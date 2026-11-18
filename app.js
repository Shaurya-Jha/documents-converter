'use strict';

/*
 * .docx / .pdf → Markdown, entirely in the browser.
 * No dependencies: .docx zips are read with DecompressionStream and parsed as
 * WordprocessingML; PDFs are parsed with a small built-in reader (xref tables
 * and streams, FlateDecode, object streams, ToUnicode CMaps, text operators).
 */

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  v: 'urn:schemas-microsoft-com:vml',
};

const MONO_FONTS = new Set([
  'courier', 'courier new', 'consolas', 'monaco', 'menlo', 'lucida console',
  'source code pro', 'fira code', 'dejavu sans mono', 'liberation mono',
]);

const IMAGE_TYPES = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  webp: 'image/webp', svg: 'image/svg+xml', bmp: 'image/bmp',
};

/* ------------------------------ zip ------------------------------ */

async function readZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);

  // Find the end-of-central-directory record.
  const stop = Math.max(0, bytes.length - 22 - 0xffff);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= stop; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('This is not a .docx file.');

  const count = view.getUint16(eocd + 10, true);
  // shortcut: zip64 is only needed for archives over 4 GB, upgrade if that changes.
  if (count === 0xffff) throw new Error('Zip64 archives are not supported.');

  let offset = view.getUint32(eocd + 16, true);
  const entries = new Map();
  const decoder = new TextDecoder();

  for (let i = 0; i < count; i++) {
    if (offset + 46 > bytes.length || view.getUint32(offset, true) !== 0x02014b50) {
      throw new Error('The .docx file is corrupt.');
    }
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    entries.set(name, { method, compressedSize, localOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  async function read(name) {
    const entry = entries.get(name);
    if (!entry) return null;
    const local = entry.localOffset;
    if (local + 30 > bytes.length || view.getUint32(local, true) !== 0x04034b50) {
      throw new Error('The .docx file is corrupt.');
    }
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const raw = bytes.subarray(start, start + entry.compressedSize);
    if (entry.method === 0) return raw;
    if (entry.method !== 8) throw new Error('Unsupported compression inside the .docx file.');
    try {
      const stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    } catch {
      throw new Error('Could not unpack the .docx file.');
    }
  }

  return { read };
}

/* ------------------------------ xml ------------------------------ */

function parseXml(bytes) {
  const doc = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml');
  if (!doc.documentElement || doc.documentElement.localName === 'parsererror' ||
      doc.getElementsByTagName('parsererror').length) {
    throw new Error('The .docx file contains invalid XML.');
  }
  return doc;
}

function child(el, name) {
  if (!el) return null;
  for (const node of el.children) if (node.localName === name) return node;
  return null;
}

function children(el, name) {
  const result = [];
  for (const node of el.children) if (node.localName === name) result.push(node);
  return result;
}

function wval(el) {
  if (!el) return null;
  return el.getAttributeNS(NS.w, 'val') ?? el.getAttribute('w:val');
}

function isOn(el) {
  if (!el) return false;
  const value = wval(el);
  return value !== '0' && value !== 'false' && value !== 'off' && value !== 'none';
}

/* --------------------------- docx parts -------------------------- */

async function loadStyles(zip) {
  const bytes = await zip.read('word/styles.xml');
  const styles = new Map();
  if (!bytes) return styles;
  const doc = parseXml(bytes);
  for (const style of doc.getElementsByTagNameNS(NS.w, 'style')) {
    const id = style.getAttributeNS(NS.w, 'styleId') || style.getAttribute('w:styleId');
    if (!id) continue;
    const name = child(style, 'name');
    styles.set(id, (wval(name) || '').toLowerCase());
  }
  return styles;
}

async function loadNumbering(zip) {
  const bytes = await zip.read('word/numbering.xml');
  const nums = new Map();
  if (!bytes) return nums;
  const doc = parseXml(bytes);

  const abstracts = new Map();
  for (const abstractNum of doc.getElementsByTagNameNS(NS.w, 'abstractNum')) {
    const id = abstractNum.getAttributeNS(NS.w, 'abstractNumId') ||
      abstractNum.getAttribute('w:abstractNumId');
    const levels = [];
    for (const level of abstractNum.getElementsByTagNameNS(NS.w, 'lvl')) {
      const ilvl = Number(level.getAttributeNS(NS.w, 'ilvl') || 0);
      levels[ilvl] = {
        fmt: wval(child(level, 'numFmt')) || 'decimal',
        start: Number(wval(child(level, 'start')) || 1),
      };
    }
    if (id != null) abstracts.set(id, levels);
  }

  for (const num of doc.getElementsByTagNameNS(NS.w, 'num')) {
    const id = num.getAttributeNS(NS.w, 'numId') || num.getAttribute('w:numId');
    const abstractId = wval(child(num, 'abstractNumId'));
    if (id != null) nums.set(id, abstracts.get(abstractId) || []);
  }
  return nums;
}

async function loadRels(zip) {
  const bytes = await zip.read('word/_rels/document.xml.rels');
  const rels = new Map();
  if (!bytes) return rels;
  const doc = parseXml(bytes);
  for (const rel of doc.getElementsByTagName('Relationship')) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (!id || !target) continue;
    rels.set(id, {
      target,
      external: rel.getAttribute('TargetMode') === 'External',
    });
  }
  return rels;
}

function base64(bytes) {
  let binary = '';
  const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
  }
  return btoa(binary);
}

function mediaPath(target) {
  return target.replace(/^\.?\//, '').replace(/^word\//, '');
}

async function loadImages(zip, targets) {
  const images = new Map();
  await Promise.all([...targets].map(async (target) => {
    const ext = target.split('.').pop().toLowerCase();
    const mime = IMAGE_TYPES[ext];
    if (!mime) return;
    try {
      const bytes = await zip.read(`word/${mediaPath(target)}`);
      if (bytes) images.set(mediaPath(target), `data:${mime};base64,${base64(bytes)}`);
    } catch {
      // A broken image should not sink the whole document.
    }
  }));
  return images;
}

/* --------------------------- markdown ---------------------------- */

const MD_SPECIAL = /([\\`*_[\]<>#])/g;

function escapeMd(text) {
  return text
    .replace(/\r\n?/g, '\n')
    .replace(MD_SPECIAL, '\\$1')
    .replace(/!(?=\[)/g, '\\!');
}

function headingLevel(styleId, styleName) {
  if (!styleId) return 0;
  const match = /heading\s*([1-6])/i.exec(`${styleId} ${styleName}`);
  if (match) return Number(match[1]);
  return /^title$/i.test(styleId) || /(^|\s)title$/.test(styleName) ? 1 : 0;
}

function isCodeStyle(styleId, styleName) {
  return /code|preformatted|verbatim/i.test(`${styleId || ''} ${styleName || ''}`);
}

function isCodeRun(runProps, ctx) {
  if (!runProps) return false;
  const styleId = wval(child(runProps, 'rStyle'));
  if (styleId && /code|verbatim|mono/i.test(`${styleId} ${ctx.styles.get(styleId) || ''}`)) {
    return true;
  }
  const fonts = child(runProps, 'rFonts');
  if (fonts) {
    for (const attr of ['ascii', 'hAnsi', 'cs', 'eastAsia']) {
      const font = (fonts.getAttributeNS(NS.w, attr) || '').toLowerCase();
      if (MONO_FONTS.has(font)) return true;
    }
  }
  return false;
}

function wrapInline(text, open, close) {
  const match = /^(\s*)([\s\S]*?)(\s*)$/.exec(text);
  if (!match || !match[2]) return text;
  return match[1] + open + match[2] + close + match[3];
}

function codeSpan(text) {
  const match = /^(\s*)([\s\S]*?)(\s*)$/.exec(text);
  if (!match || !match[2]) return text;
  const [, lead, core, trail] = match;
  const tick = '`'.repeat((core.match(/`+/g) || []).reduce((max, run) => Math.max(max, run.length), 0) + 1);
  const pad = core.startsWith('`') || core.endsWith('`') ? ' ' : '';
  return `${lead}${tick}${pad}${core}${pad}${tick}${trail}`;
}

function fenceFor(text) {
  let longest = 0;
  let run = 0;
  for (const char of text) {
    run = char === '`' ? run + 1 : 0;
    if (run > longest) longest = run;
  }
  return '`'.repeat(Math.max(3, longest + 1));
}

function renderImage(node, ctx) {
  let id = null;
  for (const blip of node.getElementsByTagNameNS(NS.a, 'blip')) {
    id = blip.getAttributeNS(NS.r, 'embed') || blip.getAttributeNS(NS.r, 'link');
    if (id) break;
  }
  if (!id) {
    for (const image of node.getElementsByTagNameNS(NS.v, 'imagedata')) {
      id = image.getAttributeNS(NS.r, 'id');
      if (id) break;
    }
  }

  let alt = '';
  const props = node.getElementsByTagNameNS(NS.wp, 'docPr');
  if (props.length) {
    const prop = props[0];
    alt = prop.getAttribute('descr') || prop.getAttribute('title') || prop.getAttribute('name') || '';
  }

  const rel = id && ctx.rels.get(id);
  if (!rel) return '';
  if (rel.external && /^https?:/i.test(rel.target)) {
    return `![${escapeMd(alt)}](${rel.target})`;
  }
  const dataUri = ctx.images.get(mediaPath(rel.target));
  return dataUri ? `![${escapeMd(alt)}](${dataUri})` : '';
}

function renderRun(run, ctx) {
  const props = child(run, 'rPr');
  const chunks = [];

  for (const node of run.children) {
    switch (node.localName) {
      case 't':
        chunks.push({ text: true, value: node.textContent });
        break;
      case 'tab':
        chunks.push({ text: true, value: ' ' });
        break;
      case 'br':
      case 'cr':
        chunks.push({ text: true, value: '  \n' });
        break;
      case 'noBreakHyphen':
        chunks.push({ text: true, value: '-' });
        break;
      case 'drawing':
      case 'pict': {
        const image = renderImage(node, ctx);
        if (image) chunks.push({ text: false, value: image });
        break;
      }
      default:
        break;
    }
  }
  if (!chunks.length) return '';

  const code = isCodeRun(props, ctx);
  let content = chunks
    .map((chunk) => (chunk.text ? (code ? chunk.value : escapeMd(chunk.value)) : chunk.value))
    .join('');
  if (!content.trim()) return content;
  if (code) return codeSpan(content);

  const sup = wval(child(props, 'vertAlign')) === 'superscript';
  const sub = wval(child(props, 'vertAlign')) === 'subscript';
  if (isOn(child(props, 'b')) || /strong/i.test(wval(child(props, 'rStyle')) || '')) content = wrapInline(content, '**', '**');
  if (isOn(child(props, 'i')) || /emphasis/i.test(wval(child(props, 'rStyle')) || '')) content = wrapInline(content, '*', '*');
  if (isOn(child(props, 'strike')) || isOn(child(props, 'dstrike'))) content = wrapInline(content, '~~', '~~');
  if (isOn(child(props, 'u'))) content = wrapInline(content, '<u>', '</u>');
  if (sup) content = wrapInline(content, '<sup>', '</sup>');
  if (sub) content = wrapInline(content, '<sub>', '</sub>');
  return content;
}

function renderInline(node, ctx, parts) {
  switch (node.localName) {
    case 'r': {
      const text = renderRun(node, ctx);
      if (text) parts.push(text);
      break;
    }
    case 'hyperlink': {
      const id = node.getAttributeNS(NS.r, 'id');
      const rel = id ? ctx.rels.get(id) : null;
      const inner = [];
      for (const nested of node.children) renderInline(nested, ctx, inner);
      // A link is already visually underlined; drop redundant <u> markers.
      const text = inner.join('').replace(/<\/?u>/g, '');
      if (!text) break;
      if (rel && rel.external && /^(https?:|mailto:)/i.test(rel.target)) {
        const url = rel.target.replace(/ /g, '%20').replace(/\(/g, '%28').replace(/\)/g, '%29');
        parts.push(`[${text}](${url})`);
      } else {
        parts.push(text);
      }
      break;
    }
    case 'ins':
    case 'smartTag':
    case 'fldSimple':
    case 'sdt':
    case 'sdtContent':
    case 'customXml':
    case 'moveTo':
      for (const nested of node.children) renderInline(nested, ctx, parts);
      break;
    default:
      break;
  }
}

function renderInlineChildren(paragraph, ctx) {
  const parts = [];
  for (const node of paragraph.children) renderInline(node, ctx, parts);
  return parts.join('');
}

function plainText(paragraph) {
  let text = '';
  const walk = (node) => {
    for (const childNode of node.children) {
      if (childNode.localName === 't') text += childNode.textContent;
      else if (childNode.localName === 'tab') text += '    ';
      else if (childNode.localName === 'br' || childNode.localName === 'cr') text += '\n';
      else if (childNode.localName !== 'del') walk(childNode);
    }
  };
  walk(paragraph);
  return text.replace(/\r\n?/g, '\n').replace(/\n+$/, '');
}

function nextNumber(ctx, numId, ilvl, level) {
  const key = `${numId}:${ilvl}`;
  let current = ctx.counters.get(key);
  if (current == null) current = (level.start || 1) - 1;
  current += 1;
  ctx.counters.set(key, current);
  for (const other of ctx.counters.keys()) {
    const [id, depth] = other.split(':');
    if (id === numId && Number(depth) > ilvl) ctx.counters.delete(other);
  }
  return current;
}

function renderParagraph(paragraph, ctx, out) {
  const props = child(paragraph, 'pPr');
  const styleId = props ? wval(child(props, 'pStyle')) : null;
  const styleName = (styleId && ctx.styles.get(styleId)) || '';
  const numPr = props ? child(props, 'numPr') : null;
  const numId = numPr ? wval(child(numPr, 'numId')) : null;
  const ilvl = numPr ? Number(wval(child(numPr, 'ilvl')) || 0) : 0;

  if (numId && numId !== '0') {
    const level = (ctx.numbering.get(numId) || [])[ilvl] || {};
    let base = ctx.baseLevels.get(numId);
    if (base == null || ilvl < base) ctx.baseLevels.set(numId, base = ilvl);
    const marker = level.fmt === 'bullet' || level.fmt === 'none'
      ? '- '
      : `${nextNumber(ctx, numId, ilvl, level)}. `;
    const indent = '    '.repeat(Math.max(0, ilvl - base));
    const text = renderInlineChildren(paragraph, ctx).trim();
    if (!text) return;
    const continuation = ' '.repeat(indent.length + marker.length);
    out.push({
      text: indent + marker + text.replace(/\n/g, `\n${continuation}`),
      list: numId,
    });
    return;
  }

  const heading = headingLevel(styleId, styleName) ||
    (props && child(props, 'outlineLvl') && Number(wval(child(props, 'outlineLvl'))) < 6
      ? Number(wval(child(props, 'outlineLvl'))) + 1
      : 0);
  if (heading) {
    const text = renderInlineChildren(paragraph, ctx).replace(/\n/g, ' ').trim();
    if (text) out.push({ text: `${'#'.repeat(heading)} ${text}` });
    return;
  }

  if (/quot/i.test(`${styleId || ''} ${styleName}`)) {
    const text = renderInlineChildren(paragraph, ctx).trim();
    if (text) out.push({ text: text.split('\n').map((line) => `> ${line}`).join('\n') });
    return;
  }

  if (isCodeStyle(styleId, styleName)) {
    const code = plainText(paragraph);
    if (code.trim()) {
      // Consecutive preformatted paragraphs are one code block.
      const last = out[out.length - 1];
      if (last && last.code) last.text += `\n${code}`;
      else out.push({ text: code, code: true });
    }
    return;
  }

  const text = renderInlineChildren(paragraph, ctx).trim();
  if (text) out.push({ text });
}

function renderTable(table, ctx) {
  const rows = [];
  for (const row of children(table, 'tr')) {
    const cells = [];
    for (const cell of children(row, 'tc')) {
      const cellProps = child(cell, 'tcPr');
      const span = cellProps ? Number(wval(child(cellProps, 'gridSpan')) || 1) : 1;
      let text = '';
      for (const paragraph of children(cell, 'p')) {
        const part = renderInlineChildren(paragraph, ctx).trim();
        if (part) text += (text ? '<br>' : '') + part;
      }
      cells.push(text.replace(/\|/g, '\\|').replace(/\n/g, '<br>'));
      for (let i = 1; i < span; i++) cells.push('');
    }
    if (cells.length) rows.push(cells);
  }
  if (!rows.length) return '';

  const width = Math.max(...rows.map((row) => row.length));
  const pad = (row) => {
    while (row.length < width) row.push('');
    return row;
  };
  const lines = [
    `| ${pad(rows[0]).map((cell) => cell || ' ').join(' | ')} |`,
    `| ${Array(width).fill('---').join(' | ')} |`,
  ];
  for (let i = 1; i < rows.length; i++) {
    lines.push(`| ${pad(rows[i]).map((cell) => cell || ' ').join(' | ')} |`);
  }
  return lines.join('\n');
}

function renderBlocks(container, ctx, out) {
  for (const node of container.children) {
    if (node.localName === 'p') {
      renderParagraph(node, ctx, out);
    } else if (node.localName === 'tbl') {
      const table = renderTable(node, ctx);
      if (table) out.push({ text: table });
    } else if (node.localName === 'sdt' || node.localName === 'customXml') {
      const content = node.localName === 'sdt' ? child(node, 'sdtContent') : node;
      if (content) renderBlocks(content, ctx, out);
    }
  }
}

/* ------------------------------ pdf ------------------------------ */

const PDF_WHITESPACE = new Set([0, 9, 10, 12, 13, 32]);
const PDF_DELIMITER = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25]);
const PDF_LATIN1 = new TextDecoder('latin1');
const PDF_ENDSTREAM = Uint8Array.from('endstream', (char) => char.charCodeAt(0));

class PdfName {
  constructor(name) { this.name = name; }
}

class PdfOperator {
  constructor(name) { this.name = name; }
}

class PdfRef {
  constructor(num) { this.num = num; }
}

class PdfStream {
  constructor(dict, data) { this.dict = dict; this.data = data; }
}

function pdfSkip(s) {
  const { b } = s;
  while (s.p < b.length) {
    const c = b[s.p];
    if (PDF_WHITESPACE.has(c)) { s.p++; continue; }
    if (c === 0x25) {
      while (s.p < b.length && b[s.p] !== 10 && b[s.p] !== 13) s.p++;
      continue;
    }
    break;
  }
}

function pdfWord(s) {
  const start = s.p;
  while (s.p < s.b.length &&
    !PDF_WHITESPACE.has(s.b[s.p]) && !PDF_DELIMITER.has(s.b[s.p])) {
    s.p++;
  }
  return PDF_LATIN1.decode(s.b.subarray(start, s.p));
}

function pdfNumber(s) {
  const { b } = s;
  const start = s.p;
  if (b[s.p] === 0x2b || b[s.p] === 0x2d) s.p++;
  while (s.p < b.length && ((b[s.p] >= 0x30 && b[s.p] <= 0x39) || b[s.p] === 0x2e)) s.p++;
  const value = Number(PDF_LATIN1.decode(b.subarray(start, s.p)));
  return Number.isFinite(value) ? value : 0;
}

function pdfNameValue(s) {
  s.p++;
  let name = '';
  while (s.p < s.b.length) {
    const c = s.b[s.p];
    if (PDF_WHITESPACE.has(c) || PDF_DELIMITER.has(c)) break;
    if (c === 0x23 && s.p + 2 < s.b.length) {
      const high = parseInt(String.fromCharCode(s.b[s.p + 1]), 16);
      const low = parseInt(String.fromCharCode(s.b[s.p + 2]), 16);
      if (Number.isInteger(high) && Number.isInteger(low)) {
        name += String.fromCharCode(high * 16 + low);
        s.p += 3;
        continue;
      }
    }
    name += String.fromCharCode(c);
    s.p++;
  }
  return new PdfName(name);
}

function pdfLiteralString(s) {
  const { b } = s;
  s.p++;
  const out = [];
  let depth = 1;
  while (s.p < b.length) {
    let c = b[s.p++];
    if (c === 0x5c) {
      c = b[s.p++];
      if (c === 0x6e) out.push(10);
      else if (c === 0x72) out.push(13);
      else if (c === 0x74) out.push(9);
      else if (c === 0x62) out.push(8);
      else if (c === 0x66) out.push(12);
      else if (c === 0x0a) { /* line continuation */ }
      else if (c === 0x0d) { if (b[s.p] === 0x0a) s.p++; }
      else if (c >= 0x30 && c <= 0x37) {
        let value = c - 0x30;
        for (let i = 0; i < 2; i++) {
          const next = b[s.p];
          if (next >= 0x30 && next <= 0x37) { value = value * 8 + next - 0x30; s.p++; } else break;
        }
        out.push(value & 0xff);
      } else if (c !== undefined) out.push(c);
      continue;
    }
    if (c === 0x28) depth++;
    else if (c === 0x29) { depth--; if (!depth) break; }
    out.push(c);
  }
  return Uint8Array.from(out);
}

function pdfHexString(s) {
  const { b } = s;
  s.p++;
  const digits = [];
  while (s.p < b.length && b[s.p] !== 0x3e) {
    const value = parseInt(String.fromCharCode(b[s.p++]), 16);
    if (Number.isInteger(value)) digits.push(value);
  }
  s.p++;
  if (digits.length % 2) digits.push(0);
  const out = new Uint8Array(digits.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = digits[i * 2] * 16 + digits[i * 2 + 1];
  return out;
}

function pdfDictValue(s) {
  const dict = new Map();
  for (;;) {
    pdfSkip(s);
    if (s.p >= s.b.length) break;
    if (s.b[s.p] === 0x3e && s.b[s.p + 1] === 0x3e) { s.p += 2; break; }
    const key = pdfObject(s);
    if (!(key instanceof PdfName)) continue;
    dict.set(key.name, pdfObject(s));
  }
  return dict;
}

function pdfObject(s) {
  pdfSkip(s);
  const { b } = s;
  if (s.p >= b.length) return null;
  const c = b[s.p];

  if (c === 0x2f) return pdfNameValue(s);
  if (c === 0x28) return pdfLiteralString(s);
  if (c === 0x3c) {
    if (b[s.p + 1] === 0x3c) { s.p += 2; return pdfDictValue(s); }
    return pdfHexString(s);
  }
  if (c === 0x5b) {
    s.p++;
    const array = [];
    for (;;) {
      pdfSkip(s);
      if (s.p >= b.length) break;
      if (b[s.p] === 0x5d) { s.p++; break; }
      const before = s.p;
      array.push(pdfObject(s));
      if (s.p === before) s.p++;
    }
    return array;
  }
  if (c === 0x5d || c === 0x3e || c === 0x29 || c === 0x7b || c === 0x7d) {
    s.p++;
    return new PdfOperator('');
  }
  if (c === 0x2b || c === 0x2d || c === 0x2e || (c >= 0x30 && c <= 0x39)) {
    const number = pdfNumber(s);
    const after = s.p;
    pdfSkip(s);
    if (Number.isInteger(number) && s.p < b.length && b[s.p] >= 0x30 && b[s.p] <= 0x39) {
      const gen = pdfNumber(s);
      pdfSkip(s);
      if (Number.isInteger(gen) && s.p < b.length && b[s.p] === 0x52) {
        const next = b[s.p + 1];
        if (next === undefined || PDF_WHITESPACE.has(next) || PDF_DELIMITER.has(next)) {
          s.p++;
          return new PdfRef(number);
        }
      }
    }
    s.p = after;
    return number;
  }

  const word = pdfWord(s);
  if (!word) { s.p++; return new PdfOperator(''); }
  if (word === 'true') return true;
  if (word === 'false') return false;
  if (word === 'null') return null;
  return new PdfOperator(word);
}

function pdfStartsWith(s, word) {
  if (s.p + word.length > s.b.length) return false;
  for (let i = 0; i < word.length; i++) {
    if (s.b[s.p + i] !== word.charCodeAt(i)) return false;
  }
  return true;
}

function pdfFindBytes(bytes, needle, from) {
  outer: for (let i = from; i <= bytes.length - needle.length; i++) {
    if (bytes[i] !== needle[0]) continue;
    for (let j = 1; j < needle.length; j++) {
      if (bytes[i + j] !== needle[j]) continue outer;
    }
    return i;
  }
  return -1;
}

function pdfConcat(parts) {
  if (parts.length === 1) return parts[0];
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

function pdfStartXref(bytes) {
  const tail = PDF_LATIN1.decode(bytes.subarray(Math.max(0, bytes.length - 2048)));
  const matches = [...tail.matchAll(/startxref\s+(\d+)/g)];
  return matches.length ? Number(matches[matches.length - 1][1]) : null;
}

async function pdfParseIndirect(state, offset) {
  const s = { b: state.bytes, p: offset };
  pdfSkip(s);
  const num = pdfNumber(s);
  pdfSkip(s);
  const gen = pdfNumber(s);
  pdfSkip(s);
  if (pdfWord(s) !== 'obj') throw new Error('Invalid PDF object.');
  const value = pdfObject(s);

  const after = s.p;
  pdfSkip(s);
  if (value instanceof Map && pdfStartsWith(s, 'stream')) {
    s.p += 6;
    if (s.b[s.p] === 13) s.p++;
    if (s.b[s.p] === 10) s.p++;
    const start = s.p;
    let end = -1;
    let exact = false;
    const length = await pdfResolve(state, value.get('Length'));
    if (Number.isInteger(length) && start + length <= s.b.length) {
      const probe = { b: s.b, p: start + length };
      pdfSkip(probe);
      if (pdfStartsWith(probe, 'endstream')) {
        end = start + length;
        exact = true;
      }
    }
    if (end < 0) {
      const found = pdfFindBytes(s.b, PDF_ENDSTREAM, start);
      end = found < 0 ? s.b.length : found;
    }
    let dataEnd = end;
    // Only a scanned end offset needs the separator whitespace trimmed; a
    // trusted /Length is exact and may itself end in a newline byte.
    while (!exact && dataEnd > start && (s.b[dataEnd - 1] === 10 || s.b[dataEnd - 1] === 13)) dataEnd--;
    return { num, value: new PdfStream(value, s.b.subarray(start, dataEnd)) };
  }
  s.p = after;
  return { num, value };
}

async function pdfInflate(bytes, raw) {
  const stream = new Blob([bytes]).stream()
    .pipeThrough(new DecompressionStream(raw ? 'deflate-raw' : 'deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function pdfFlate(bytes) {
  try {
    return await pdfInflate(bytes, false);
  } catch {
    return pdfInflate(bytes, true);
  }
}

function pdfLzw(bytes) {
  const out = [];
  let dict = [];
  const reset = () => {
    dict = [];
    for (let i = 0; i < 256; i++) dict.push([i]);
    dict.push(null, null);
  };
  reset();
  let codeSize = 9;
  let bitBuffer = 0;
  let bitCount = 0;
  let previous = null;
  let p = 0;
  while (p < bytes.length) {
    while (bitCount < codeSize && p < bytes.length) {
      bitBuffer = (bitBuffer << 8) | bytes[p++];
      bitCount += 8;
    }
    if (bitCount < codeSize) break;
    const code = (bitBuffer >> (bitCount - codeSize)) & ((1 << codeSize) - 1);
    bitCount -= codeSize;
    if (code === 256) { reset(); codeSize = 9; previous = null; continue; }
    if (code === 257) break;
    let entry;
    if (code < dict.length && dict[code]) entry = dict[code];
    else if (previous) entry = [...previous, previous[0]];
    else break;
    out.push(...entry);
    if (previous) dict.push([...previous, entry[0]]);
    previous = entry;
    if (dict.length + 1 > (1 << codeSize) && codeSize < 12) codeSize++;
  }
  return Uint8Array.from(out);
}

function pdfPredictor(data, parms) {
  const predictor = Number(parms.get('Predictor')) || 1;
  if (predictor === 1) return data;
  const colors = Number(parms.get('Colors')) || 1;
  const bits = Number(parms.get('BitsPerComponent')) || 8;
  const columns = Number(parms.get('Columns')) || 1;
  const rowLength = Math.ceil(colors * bits * columns / 8);
  if (predictor === 2) {
    if (bits !== 8) return data;
    const out = data.slice();
    for (let row = 0; row < out.length; row += rowLength) {
      for (let i = colors; i < rowLength && row + i < out.length; i++) {
        out[row + i] = (out[row + i] + out[row + i - colors]) & 0xff;
      }
    }
    return out;
  }
  const bpp = Math.ceil(colors * bits / 8);
  const rows = Math.floor(data.length / (rowLength + 1));
  const out = new Uint8Array(rows * rowLength);
  let p = 0;
  for (let row = 0; row < rows; row++) {
    const filter = data[p++];
    const base = row * rowLength;
    for (let i = 0; i < rowLength; i++) {
      const raw = data[p++];
      const left = i >= bpp ? out[base + i - bpp] : 0;
      const up = row > 0 ? out[base - rowLength + i] : 0;
      const upLeft = row > 0 && i >= bpp ? out[base - rowLength + i - bpp] : 0;
      let value;
      if (filter === 1) value = raw + left;
      else if (filter === 2) value = raw + up;
      else if (filter === 3) value = raw + ((left + up) >> 1);
      else if (filter === 4) {
        const guess = left + up - upLeft;
        const dl = Math.abs(guess - left);
        const du = Math.abs(guess - up);
        const dul = Math.abs(guess - upLeft);
        value = raw + (dl <= du && dl <= dul ? left : du <= dul ? up : upLeft);
      } else value = raw;
      out[base + i] = value & 0xff;
    }
  }
  return out;
}

async function pdfStreamData(state, stream) {
  let data = stream.data;
  let filters = await pdfResolve(state, stream.dict.get('Filter'));
  let parms = await pdfResolve(state, stream.dict.get('DecodeParms')) || [];
  if (!filters) return data;
  if (!Array.isArray(filters)) filters = [filters];
  if (!Array.isArray(parms)) parms = [parms];

  for (let i = 0; i < filters.length; i++) {
    const filter = filters[i] instanceof PdfName ? filters[i].name : '';
    if (filter === 'FlateDecode' || filter === 'Fl') data = await pdfFlate(data);
    else if (filter === 'LZWDecode' || filter === 'LZW') data = pdfLzw(data);
    else if (filter === 'ASCIIHexDecode' || filter === 'AHx') {
      const digits = [];
      for (const byte of data) {
        const value = parseInt(String.fromCharCode(byte), 16);
        if (Number.isInteger(value)) digits.push(value);
        else if (byte === 0x3e) break;
      }
      if (digits.length % 2) digits.push(0);
      const out = new Uint8Array(digits.length / 2);
      for (let j = 0; j < out.length; j++) out[j] = digits[j * 2] * 16 + digits[j * 2 + 1];
      data = out;
    } else if (filter === 'ASCII85Decode' || filter === 'A85') {
      const text = PDF_LATIN1.decode(data).replace(/\s+/g, '');
      const out = [];
      let group = [];
      for (const char of text) {
        if (char === '~') break;
        if (char === 'z' && !group.length) { out.push(0, 0, 0, 0); continue; }
        group.push(char.charCodeAt(0) - 33);
        if (group.length === 5) {
          let value = 0;
          for (const digit of group) value = value * 85 + digit;
          out.push((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff);
          group = [];
        }
      }
      if (group.length) {
        const count = group.length;
        while (group.length < 5) group.push(84);
        let value = 0;
        for (const digit of group) value = value * 85 + digit;
        const bytesOut = [(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff];
        out.push(...bytesOut.slice(0, count - 1));
      }
      data = Uint8Array.from(out);
    } else if (filter === 'RunLengthDecode' || filter === 'RL') {
      const out = [];
      let p = 0;
      while (p < data.length) {
        const length = data[p++];
        if (length === 128) break;
        if (length < 128) {
          for (let j = 0; j <= length && p < data.length; j++) out.push(data[p++]);
        } else {
          const byte = data[p++];
          for (let j = 0; j < 257 - length; j++) out.push(byte);
        }
      }
      data = Uint8Array.from(out);
    } else {
      throw new Error(`Unsupported PDF stream filter: ${filter}`);
    }
    const parm = parms[i] ? await pdfResolve(state, parms[i]) : null;
    if (parm instanceof Map) data = pdfPredictor(data, parm);
  }
  return data;
}

async function pdfApplyXrefStream(state, stream) {
  const dict = stream.dict;
  const widths = dict.get('W');
  if (!Array.isArray(widths) || widths.length < 3) return;
  let index = dict.get('Index');
  if (!Array.isArray(index)) index = [0, Number(dict.get('Size')) || 0];
  const data = await pdfStreamData(state, stream);
  const rowLength = widths.reduce((sum, width) => sum + (Number(width) || 0), 0);
  if (!rowLength) return;
  let p = 0;

  for (let i = 0; i + 1 < index.length; i += 2) {
    const start = Number(index[i]);
    const count = Number(index[i + 1]);
    for (let j = 0; j < count && p + rowLength <= data.length; j++) {
      const fields = [];
      for (const width of widths) {
        let value = 0;
        for (let n = 0; n < Number(width); n++) value = value * 256 + data[p++];
        fields.push(value);
      }
      const num = start + j;
      if (state.entries.has(num)) continue;
      const type = Number(widths[0]) === 0 ? 1 : fields[0];
      if (type === 1) state.entries.set(num, { type: 1, offset: fields[1] });
      else if (type === 2) state.entries.set(num, { type: 2, stm: fields[1], index: fields[2] });
    }
  }

  for (const key of ['Root', 'Info', 'Encrypt', 'ID', 'Prev']) {
    if (dict.has(key) && !state.trailer.has(key)) state.trailer.set(key, dict.get(key));
  }
}

async function pdfReadXref(state) {
  let offset = pdfStartXref(state.bytes);
  const seen = new Set();

  while (Number.isInteger(offset) && offset >= 0 && offset < state.bytes.length && !seen.has(offset)) {
    seen.add(offset);
    const s = { b: state.bytes, p: offset };
    pdfSkip(s);

    if (pdfStartsWith(s, 'xref')) {
      s.p += 4;
      let trailer = null;
      for (;;) {
        pdfSkip(s);
        if (s.p >= state.bytes.length) break;
        if (pdfStartsWith(s, 'trailer')) {
          s.p += 7;
          trailer = pdfObject(s);
          break;
        }
        if (!(s.b[s.p] >= 0x30 && s.b[s.p] <= 0x39)) break;
        const start = pdfNumber(s);
        const count = pdfNumber(s);
        for (let i = 0; i < count; i++) {
          pdfSkip(s);
          if (s.p >= state.bytes.length) break;
          const entryOffset = pdfNumber(s);
          pdfNumber(s);
          pdfSkip(s);
          const type = String.fromCharCode(s.b[s.p]);
          s.p++;
          const num = start + i;
          if (type === 'n' && !state.entries.has(num)) state.entries.set(num, { type: 1, offset: entryOffset });
        }
      }
      if (!(trailer instanceof Map)) break;
      for (const [key, value] of trailer) {
        if (!state.trailer.has(key)) state.trailer.set(key, value);
      }
      const hybrid = trailer.get('XRefStm');
      if (Number.isInteger(hybrid) && !seen.has(hybrid)) {
        seen.add(hybrid);
        try {
          const parsed = await pdfParseIndirect(state, hybrid);
          if (parsed.value instanceof PdfStream) await pdfApplyXrefStream(state, parsed.value);
        } catch { /* hybrid xref is optional */ }
      }
      offset = trailer.get('Prev');
    } else {
      const parsed = await pdfParseIndirect(state, offset);
      if (!(parsed.value instanceof PdfStream)) break;
      await pdfApplyXrefStream(state, parsed.value);
      const prev = parsed.value.dict.get('Prev');
      offset = Number.isInteger(prev) ? prev : null;
    }
  }
}

function pdfScanObjects(state) {
  if (state.scanned) return;
  state.scanned = true;
  const text = PDF_LATIN1.decode(state.bytes);
  for (const match of text.matchAll(/(\d+)\s+(\d+)\s+obj\b/g)) {
    state.scanOffsets.set(Number(match[1]), match.index);
  }
  if (!state.trailer.has('Root')) {
    const root = /\/Root\s+(\d+)\s+\d+\s+R/.exec(text);
    if (root) {
      state.trailer.set('Root', new PdfRef(Number(root[1])));
    } else {
      const catalog = /\/Type\s*\/Catalog/.exec(text);
      if (catalog) {
        const before = text.slice(Math.max(0, catalog.index - 4096), catalog.index);
        const objects = [...before.matchAll(/(\d+)\s+(\d+)\s+obj\b/g)];
        if (objects.length) state.trailer.set('Root', new PdfRef(Number(objects[objects.length - 1][1])));
      }
    }
  }
}

async function pdfObjectFromStream(state, stmNum, index, num) {
  const stream = await pdfGet(state, new PdfRef(stmNum));
  if (!(stream instanceof PdfStream)) return null;
  let table = state.objStms.get(stmNum);
  if (!table) {
    const data = await pdfStreamData(state, stream);
    const count = Number(stream.dict.get('N')) || 0;
    const first = Number(stream.dict.get('First')) || 0;
    const s = { b: data, p: 0 };
    table = { data, first, items: [] };
    for (let i = 0; i < count; i++) {
      const objectNum = pdfNumber(s);
      pdfSkip(s);
      const objectOffset = pdfNumber(s);
      pdfSkip(s);
      table.items.push({ num: objectNum, offset: objectOffset });
    }
    state.objStms.set(stmNum, table);
  }
  let item = table.items[index];
  if (!item || item.num !== num) item = table.items.find((entry) => entry.num === num);
  if (!item) return null;
  return pdfObject({ b: table.data, p: table.first + item.offset });
}

async function pdfLoadObject(state, num) {
  const entry = state.entries.get(num);
  if (entry && entry.type === 2) return pdfObjectFromStream(state, entry.stm, entry.index, num);

  if (entry && Number.isInteger(entry.offset)) {
    try {
      const parsed = await pdfParseIndirect(state, entry.offset);
      if (parsed.num === num) return parsed.value;
    } catch { /* fall through to the scan */ }
  }

  pdfScanObjects(state);
  const fallback = state.scanOffsets.get(num);
  if (!Number.isInteger(fallback)) return null;
  try {
    const parsed = await pdfParseIndirect(state, fallback);
    if (parsed.num === num) return parsed.value;
  } catch { /* unreadable object */ }
  return null;
}

async function pdfGet(state, ref) {
  const num = ref instanceof PdfRef ? ref.num : ref;
  if (state.cache.has(num)) return state.cache.get(num);
  const promise = pdfLoadObject(state, num);
  state.cache.set(num, promise);
  return promise;
}

async function pdfResolve(state, value) {
  return value instanceof PdfRef ? pdfGet(state, value) : value;
}

async function pdfOpen(bytes) {
  const state = {
    bytes,
    entries: new Map(),
    cache: new Map(),
    objStms: new Map(),
    trailer: new Map(),
    scanOffsets: new Map(),
    scanned: false,
    pageFonts: new Map(),
    fontCache: new Map(),
  };
  try {
    await pdfReadXref(state);
  } catch { /* corrupt xref, fall back to scanning */ }
  if (!state.entries.size) pdfScanObjects(state);
  return state;
}

async function pdfCatalog(state) {
  const root = state.trailer.get('Root');
  if (root) {
    const catalog = await pdfResolve(state, root);
    if (catalog instanceof Map) return catalog;
  }
  pdfScanObjects(state);
  for (const num of state.scanOffsets.keys()) {
    const object = await pdfGet(state, new PdfRef(num));
    if (object instanceof Map) {
      const type = object.get('Type');
      if (type instanceof PdfName && type.name === 'Catalog') return object;
    }
  }
  return null;
}

async function pdfPages(state) {
  const catalog = await pdfCatalog(state);
  if (!catalog) return [];
  const pages = [];
  const walk = async (ref, inherited) => {
    const node = await pdfResolve(state, ref);
    if (!(node instanceof Map)) return;
    const resources = await pdfResolve(state, node.get('Resources')) || inherited;
    const kids = await pdfResolve(state, node.get('Kids'));
    if (Array.isArray(kids)) {
      for (const kid of kids) await walk(kid, resources);
      return;
    }
    pages.push({ node, resources });
  };
  await walk(catalog.get('Pages'), null);
  return pages;
}

const PDF_GLYPHS = {
  space: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%',
  ampersand: '&', quotesingle: "'", parenleft: '(', parenright: ')', asterisk: '*',
  plus: '+', comma: ',', hyphen: '-', period: '.', slash: '/', colon: ':', semicolon: ';',
  less: '<', equal: '=', greater: '>', question: '?', at: '@', bracketleft: '[',
  backslash: '\\', bracketright: ']', asciicircum: '^', underscore: '_', grave: '`',
  braceleft: '{', bar: '|', braceright: '}', asciitilde: '~', zero: '0', one: '1',
  two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8',
  nine: '9', nbspace: ' ', periodcentered: '·', bullet: '•', endash: '–', emdash: '—',
  quoteleft: '‘', quoteright: '’', quotedblleft: '“', quotedblright: '”', ellipsis: '…',
  trademark: '™', copyright: '©', registered: '®', degree: '°', section: '§',
  paragraph: '¶', dagger: '†', daggerdbl: '‡', perthousand: '‰', minus: '−',
  plusminus: '±', multiply: '×', divide: '÷', onehalf: '½', onequarter: '¼',
  threequarters: '¾', guillemotleft: '«', guillemotright: '»', guilsinglleft: '‹',
  guilsinglright: '›', Euro: '€', sterling: '£', yen: '¥', cent: '¢', currency: '¤',
  brokenbar: '¦', dieresis: '¨', macron: '¯', acute: '´', cedilla: '¸', ordfeminine: 'ª',
  ordmasculine: 'º', logicalnot: '¬', mu: 'µ', onesuperior: '¹', twosuperior: '²',
  threesuperior: '³', fi: 'fi', fl: 'fl', ff: 'ff', ffi: 'ffi', ffl: 'ffl',
  ae: 'æ', AE: 'Æ', oe: 'œ', OE: 'Œ', oslash: 'ø', Oslash: 'Ø', aring: 'å', Aring: 'Å',
  germandbls: 'ß', eth: 'ð', ETH: 'Ð', thorn: 'þ', THORN: 'Þ',
};

function pdfGlyph(name) {
  if (PDF_GLYPHS[name] != null) return PDF_GLYPHS[name];
  if (name.length === 1) return name;
  const uni = /^uni([0-9a-f]{4})$/i.exec(name);
  if (uni) return String.fromCharCode(parseInt(uni[1], 16));
  const u = /^u([0-9a-f]{4,6})$/i.exec(name);
  if (u) return String.fromCodePoint(parseInt(u[1], 16));
  return '';
}

const PDF_WIN_ANSI = (() => {
  const high = {
    0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…', 0x86: '†', 0x87: '‡',
    0x88: 'ˆ', 0x89: '‰', 0x8a: 'Š', 0x8b: '‹', 0x8c: 'Œ', 0x8e: 'Ž', 0x91: '‘',
    0x92: '’', 0x93: '“', 0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—', 0x98: '˜',
    0x99: '™', 0x9a: 'š', 0x9b: '›', 0x9c: 'œ', 0x9e: 'ž', 0x9f: 'Ÿ',
  };
  const table = [];
  for (let i = 0; i < 256; i++) {
    if (i === 9) table[i] = '\t';
    else if (i === 10 || i === 13) table[i] = '\n';
    else if (i < 32 || i === 127) table[i] = '';
    else table[i] = high[i] ?? String.fromCharCode(i);
  }
  return table;
})();

function pdfDefaultFont() {
  return {
    map: null, codes: 1, simple: PDF_WIN_ANSI, bold: false, italic: false, mono: false,
    widths: null, firstChar: 0, defaultWidth: 500, missingWidth: 0,
  };
}

/* Approximate width of decoded text in glyph units (1000 = one em). */
function pdfTextWidth(font, bytes, text) {
  let units = 0;
  let count = 0;
  if (font.widths instanceof Map || font.widths) {
    if (font.map || font.codes === 1) {
      const step = font.codes || 1;
      for (let i = 0; i + step <= bytes.length; i += step) {
        let code = 0;
        for (let j = 0; j < step; j++) code = code * 256 + bytes[i + j];
        const width = font.widths instanceof Map
          ? font.widths.get(code) ?? font.defaultWidth
          : font.widths[code - font.firstChar] ?? font.defaultWidth;
        units += Number(width) || 0;
        count++;
      }
    }
  }
  if (!count) {
    const fallback = font.codes === 2 ? 1000 : 500;
    return [...text].length * fallback;
  }
  return units;
}

function pdfUtf16(bytes) {
  if (bytes.length === 1) return String.fromCharCode(bytes[0]);
  let out = '';
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    out += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
  }
  if (bytes.length % 2) out += String.fromCharCode(bytes[bytes.length - 1]);
  return out;
}

function pdfCode(bytes) {
  let code = 0;
  for (const byte of bytes) code = code * 256 + byte;
  return code;
}

function pdfParseCMap(data) {
  const map = new Map();
  let codes = 0;
  let operands = [];
  const s = { b: data, p: 0 };
  while (s.p < data.length) {
    pdfSkip(s);
    if (s.p >= data.length) break;
    const before = s.p;
    let value;
    try {
      value = pdfObject(s);
    } catch {
      s.p = before + 1;
      continue;
    }
    if (s.p === before) { s.p++; continue; }
    if (!(value instanceof PdfOperator)) { operands.push(value); continue; }

    if (value.name === 'endcodespacerange') {
      if (operands[0] instanceof Uint8Array && operands[0].length) codes = operands[0].length;
    } else if (value.name === 'endbfchar') {
      for (let i = 0; i + 1 < operands.length; i += 2) {
        if (operands[i] instanceof Uint8Array && operands[i + 1] instanceof Uint8Array) {
          map.set(pdfCode(operands[i]), pdfUtf16(operands[i + 1]));
        }
      }
    } else if (value.name === 'endbfrange') {
      for (let i = 0; i + 2 < operands.length; i += 3) {
        const lo = operands[i];
        const hi = operands[i + 1];
        const dst = operands[i + 2];
        if (!(lo instanceof Uint8Array) || !(hi instanceof Uint8Array)) continue;
        const start = pdfCode(lo);
        const end = pdfCode(hi);
        if (Array.isArray(dst)) {
          for (let code = start; code <= end && code - start < dst.length; code++) {
            if (dst[code - start] instanceof Uint8Array) map.set(code, pdfUtf16(dst[code - start]));
          }
        } else if (dst instanceof Uint8Array) {
          const base = pdfUtf16(dst);
          if (!base) continue;
          for (let code = start; code <= end; code++) {
            map.set(code, base.slice(0, -1) + String.fromCharCode(base.charCodeAt(base.length - 1) + code - start));
          }
        }
      }
    }
    operands = [];
  }
  return { map, codes };
}

async function pdfFont(state, ref) {
  const dict = await pdfResolve(state, ref);
  if (!(dict instanceof Map)) return pdfDefaultFont();
  let font = state.fontCache.get(dict);
  if (!font) {
    font = pdfLoadFont(state, dict);
    state.fontCache.set(dict, font);
  }
  return font;
}

async function pdfLoadFont(state, dict) {
  const font = pdfDefaultFont();

  const subtype = dict.get('Subtype');
  const base = dict.get('BaseFont');
  const baseName = base instanceof PdfName ? base.name : '';
  font.bold = /bold|black|heavy|semibold|demi/i.test(baseName);
  font.italic = /italic|oblique/i.test(baseName);
  font.mono = /mono|courier|consol/i.test(baseName);
  const isType0 = subtype instanceof PdfName && subtype.name === 'Type0';
  if (isType0) font.codes = 2;

  // Widths come from the simple font itself or, for Type0, its descendant CID font.
  let metrics = dict;
  if (isType0) {
    const descendants = await pdfResolve(state, dict.get('DescendantFonts'));
    if (Array.isArray(descendants)) metrics = await pdfResolve(state, descendants[0]) || dict;
  }

  if (isType0) {
    const widths = await pdfResolve(state, metrics.get('W'));
    const map = new Map();
    if (Array.isArray(widths)) {
      for (let i = 0; i + 1 < widths.length;) {
        const first = Number(widths[i++]);
        const next = await pdfResolve(state, widths[i]);
        if (Array.isArray(next)) {
          i++;
          for (let j = 0; j < next.length; j++) map.set(first + j, Number(await pdfResolve(state, next[j])) || 0);
        } else {
          const last = Number(next);
          i++;
          const width = Number(await pdfResolve(state, widths[i++])) || 0;
          for (let code = first; code <= last; code++) map.set(code, width);
        }
      }
    }
    font.widths = map;
    font.defaultWidth = Number(metrics.get('DW')) || 1000;
  } else {
    const widths = await pdfResolve(state, metrics.get('Widths'));
    if (Array.isArray(widths)) {
      font.widths = widths;
      font.firstChar = Number(metrics.get('FirstChar')) || 0;
      font.defaultWidth = 500;
    }
  }
  const descriptor = await pdfResolve(state, metrics.get('FontDescriptor'));
  if (descriptor instanceof Map) {
    font.missingWidth = Number(descriptor.get('MissingWidth')) || 0;
    if (font.widths instanceof Map) font.defaultWidth = font.missingWidth || font.defaultWidth;
  }

  const toUnicode = await pdfResolve(state, dict.get('ToUnicode'));
  if (toUnicode instanceof PdfStream) {
    try {
      const cmap = pdfParseCMap(await pdfStreamData(state, toUnicode));
      if (cmap.map.size) {
        font.map = cmap.map;
        if (cmap.codes) font.codes = cmap.codes;
      }
    } catch { /* malformed CMap, fall back to the built-in encoding */ }
  }

  if (font.codes === 1) {
    const encoding = await pdfResolve(state, dict.get('Encoding'));
    if (encoding instanceof Map) {
      const table = PDF_WIN_ANSI.slice();
      const differences = await pdfResolve(state, encoding.get('Differences'));
      if (Array.isArray(differences)) {
        let code = 0;
        for (const item of differences) {
          if (typeof item === 'number') code = item;
          else if (item instanceof PdfName) {
            table[code] = pdfGlyph(item.name);
            code++;
          }
        }
      }
      font.simple = table;
    }
  }
  return font;
}

async function pdfFontsFor(state, resources) {
  if (!(resources instanceof Map)) return new Map();
  const cached = state.pageFonts.get(resources);
  if (cached) return cached;
  const fonts = new Map();
  state.pageFonts.set(resources, fonts);
  const dict = await pdfResolve(state, resources.get('Font'));
  if (dict instanceof Map) {
    await Promise.all([...dict].map(async ([name, ref]) => {
      fonts.set(name, await pdfFont(state, ref));
    }));
  }
  return fonts;
}

function pdfDecodeString(font, bytes) {
  let text = '';
  if (font.map) {
    const step = font.codes || 1;
    for (let i = 0; i + step <= bytes.length; i += step) {
      let code = 0;
      for (let j = 0; j < step; j++) code = code * 256 + bytes[i + j];
      const value = font.map.get(code);
      if (value != null) text += value;
    }
    return text;
  }
  if (font.codes === 1) {
    for (const byte of bytes) text += font.simple[byte];
  }
  return text;
}

function pdfEstimateWidth(text, size) {
  let width = 0;
  for (const char of text) width += (char === ' ' ? 0.28 : 0.52) * size;
  return width;
}

function pdfIdentity() {
  return [1, 0, 0, 1, 0, 0];
}

function pdfMultiply(left, right) {
  return [
    left[0] * right[0] + left[1] * right[2],
    left[0] * right[1] + left[1] * right[3],
    left[2] * right[0] + left[3] * right[2],
    left[2] * right[1] + left[3] * right[3],
    left[4] * right[0] + left[5] * right[2] + right[4],
    left[4] * right[1] + left[5] * right[3] + right[5],
  ];
}

function pdfAppend(state, text, x, y, size, bold, italic, advance) {
  const lines = state.lines;
  let line = lines[lines.length - 1];
  const tolerance = Math.max(size, line ? line.size : size) * 0.6;
  const width = advance != null ? (advance / 1000) * size : pdfEstimateWidth(text, size);
  let gap = 0;

  if (state.startNew || !line || Math.abs(line.y - y) > tolerance) {
    const fresh = {
      y, x, endX: x + width, size, parts: [],
      blankBefore: Boolean(line && !state.startNew && Math.abs(line.y - y) > line.size * 1.55),
    };
    lines.push(fresh);
    line = fresh;
    state.startNew = false;
  }

  if (line.parts.length) {
    gap = x - line.endX;
    const previous = line.parts[line.parts.length - 1];
    if (gap > size * 0.12 && !/\s$/.test(previous.text) && !/^\s/.test(text)) {
      line.parts.push({ text: ' ', bold: false, italic: false, size });
    }
  }
  line.parts.push({ text, x, bold, italic, size, code: Boolean(state.font && state.font.mono) });
  line.endX = x + width;
  line.size = Math.max(line.size, size);
  // A wide internal gap means side-by-side columns, so the row stands alone.
  if (gap > size * 1.2) line.spread = true;
}

function pdfShow(state, value) {
  if (!(value instanceof Uint8Array)) return;
  const font = state.font;
  const text = pdfDecodeString(font, value);
  if (!text) return;

  // Text space -> page space, including the current transformation matrix.
  const matrix = pdfMultiply(state.tm, state.ctm);
  const scale = Math.hypot(matrix[0], matrix[1]) || 1;
  const size = state.size * scale;
  const units = font.codes === 1 || font.widths
    ? pdfTextWidth(font, value, text)
    : [...text].length * 500;
  const advance = (units / 1000) * state.size * state.hscale;

  // Rotated text (vertical watermarks, margin stamps) and near-invisible
  // figure-label text are not part of the reading flow.
  const rotated = Math.abs(matrix[1]) > Math.abs(matrix[0]) * 0.2;
  if (!rotated && size >= 2) {
    pdfAppend(state, text, matrix[4], matrix[5] + state.rise * scale, size,
      font.bold, font.italic, units);
  }

  const spaces = (text.match(/ /g) || []).length;
  state.tm[4] += advance + (state.charSpace * text.length + state.wordSpace * spaces) * state.hscale;
}

function pdfMoveText(state, tx, ty) {
  state.tlm = pdfMultiply([1, 0, 0, 1, tx, ty], state.tlm);
  state.tm = state.tlm.slice();
}

function pdfExecute(op, operands, state) {
  switch (op) {
    case 'BT':
      state.tm = pdfIdentity();
      state.tlm = pdfIdentity();
      break;
    case 'Tf': {
      const name = operands[0];
      state.font = name instanceof PdfName ? (state.fonts.get(name.name) || pdfDefaultFont()) : state.font;
      state.size = Number(operands[1]) || state.size;
      break;
    }
    case 'TL': state.leading = Number(operands[0]) || 0; break;
    case 'Tc': state.charSpace = Number(operands[0]) || 0; break;
    case 'Tw': state.wordSpace = Number(operands[0]) || 0; break;
    case 'Tz': state.hscale = (Number(operands[0]) || 100) / 100; break;
    case 'Ts': state.rise = Number(operands[0]) || 0; break;
    case 'Td': pdfMoveText(state, Number(operands[0]) || 0, Number(operands[1]) || 0); break;
    case 'TD':
      state.leading = -(Number(operands[1]) || 0);
      pdfMoveText(state, Number(operands[0]) || 0, Number(operands[1]) || 0);
      break;
    case 'Tm':
      if (operands.length >= 6) {
        state.tlm = operands.slice(0, 6).map(Number);
        state.tm = state.tlm.slice();
      }
      break;
    case 'T*': pdfMoveText(state, 0, -state.leading); break;
    case 'Tj': pdfShow(state, operands[0]); break;
    case "'":
      pdfMoveText(state, 0, -state.leading);
      pdfShow(state, operands[0]);
      break;
    case '"':
      state.wordSpace = Number(operands[0]) || 0;
      state.charSpace = Number(operands[1]) || 0;
      pdfMoveText(state, 0, -state.leading);
      pdfShow(state, operands[2]);
      break;
    case 'TJ':
      if (Array.isArray(operands[0])) {
        for (const item of operands[0]) {
          if (typeof item === 'number') {
            const shift = (-item / 1000) * state.size * state.hscale;
            if (shift > state.size * 0.18) {
              pdfAppend(state, ' ', state.tm[4], state.tm[5] + state.rise, state.size,
                state.font.bold, state.font.italic, 280);
            }            state.tm[4] += shift;
          } else {
            pdfShow(state, item);
          }
        }
      }
      break;
    default:
      break;
  }
}

function pdfSkipInlineImage(bytes, from) {
  let i = from;
  while (i < bytes.length - 1) {
    if (bytes[i] === 0x45 && bytes[i + 1] === 0x49) {
      const before = bytes[i - 1];
      const after = bytes[i + 2];
      if ((before === undefined || PDF_WHITESPACE.has(before)) &&
          (after === undefined || PDF_WHITESPACE.has(after) || PDF_DELIMITER.has(after))) {
        return i + 2;
      }
    }
    i++;
  }
  return bytes.length;
}

async function pdfDoXObject(state, resources, operands, depth) {
  const name = operands[0];
  if (!(name instanceof PdfName) || depth >= 4 || !(resources instanceof Map)) return;
  const xobjects = await pdfResolve(state, resources.get('XObject'));
  if (!(xobjects instanceof Map)) return;
  const ref = xobjects.get(name.name);
  if (!ref) return;
  const object = await pdfResolve(state, ref);
  if (!(object instanceof PdfStream)) return;
  const subtype = object.dict.get('Subtype');
  if (!(subtype instanceof PdfName) || subtype.name !== 'Form') return;
  const nested = await pdfResolve(state, object.dict.get('Resources'));
  await pdfRunContent(state, await pdfStreamData(state, object), nested instanceof Map ? nested : resources, depth + 1);
}

async function pdfRunContent(state, data, resources, depth) {
  const fonts = await pdfFontsFor(state, resources);
  const previousFonts = state.fonts;
  const previousFont = state.font;
  state.fonts = fonts;

  const s = { b: data, p: 0 };
  let operands = [];
  while (s.p < data.length) {
    pdfSkip(s);
    if (s.p >= data.length) break;
    const before = s.p;
    let value;
    try {
      value = pdfObject(s);
    } catch {
      s.p = before + 1;
      operands = [];
      continue;
    }
    if (s.p === before) { s.p++; continue; }
    if (!(value instanceof PdfOperator)) { operands.push(value); continue; }

    if (value.name === 'ID') {
      s.p = pdfSkipInlineImage(s.b, s.p);
    } else if (value.name === 'Do') {
      await pdfDoXObject(state, resources, operands, depth);
    } else if (value.name === 'q') {
      state.stack.push({ ctm: state.ctm.slice(), font: state.font, size: state.size,
        leading: state.leading, charSpace: state.charSpace, wordSpace: state.wordSpace,
        hscale: state.hscale, rise: state.rise });
    } else if (value.name === 'Q') {
      const saved = state.stack.pop();
      if (saved) Object.assign(state, saved);
    } else if (value.name === 'cm') {
      if (operands.length >= 6) state.ctm = pdfMultiply(operands.slice(0, 6).map(Number), state.ctm);
    } else {
      pdfExecute(value.name, operands, state);
    }
    operands = [];
  }

  state.fonts = previousFonts;
  state.font = previousFont;
}

async function pdfPageText(state, page, lines, pageIndex) {
  const resources = page.resources;
  const contents = await pdfResolve(state, page.node.get('Contents'));
  const chunks = [];
  const list = Array.isArray(contents) ? contents : [contents];
  for (const item of list) {
    const stream = await pdfResolve(state, item);
    if (!(stream instanceof PdfStream)) continue;
    try {
      chunks.push(await pdfStreamData(state, stream));
    } catch { /* skip undecodable content */ }
  }
  if (!chunks.length) return;

  state.lines = lines;
  state.fonts = new Map();
  state.font = pdfDefaultFont();
  state.tm = pdfIdentity();
  state.tlm = pdfIdentity();
  state.ctm = pdfIdentity();
  state.stack = [];
  state.size = 12;
  state.leading = 0;
  state.charSpace = 0;
  state.wordSpace = 0;
  state.hscale = 1;
  state.rise = 0;
  state.startNew = pageIndex > 0;

  await pdfRunContent(state, pdfConcat(chunks), resources, 0);
}

function pdfBodySize(lines) {
  const counts = new Map();
  for (const line of lines) {
    const text = line.parts.map((part) => part.text).join('').trim();
    if (!text) continue;
    const size = Math.round(line.size * 2) / 2;
    counts.set(size, (counts.get(size) || 0) + text.length);
  }
  let body = 12;
  let best = -1;
  for (const [size, count] of counts) {
    if (count > best) { best = count; body = size; }
  }
  return body;
}

function pdfMergeParts(line) {
  const parts = [];
  for (const part of line.parts) {
    const last = parts[parts.length - 1];
    if (last && last.bold === part.bold && last.italic === part.italic && last.code === part.code) {
      last.text += part.text;
    } else {
      parts.push({ ...part });
    }
  }
  return parts;
}

function pdfLineText(line, plain) {
  // Merge adjacent runs that share styling so a word split across text
  // operations does not become **Pr****o****ject**.
  let text = '';
  for (const part of pdfMergeParts(line)) {
    let chunk = escapeMd(part.text);
    if (!plain) {
      if (part.code) chunk = codeSpan(chunk);
      else {
        if (part.bold) chunk = wrapInline(chunk, '**', '**');
        if (part.italic) chunk = wrapInline(chunk, '*', '*');
      }
    }
    text += chunk;
  }
  return text.replace(/\s+/g, ' ').trim();
}

function pdfLineRaw(line) {
  return pdfMergeParts(line)
    .map((part) => part.text)
    .join('')
    .replace(/\s+/g, ' ')
    .trim();
}

function pdfLinesToMarkdown(lines) {
  if (!lines.length) return '';
  const body = pdfBodySize(lines);
  // Quietly cluster font sizes so CTM rounding (19.8 vs 19.6) does not create
  // duplicate heading levels.
  const sizes = [...new Set(
    lines.filter((line) => line.size > body * 1.15).map((line) => Math.round(line.size * 10) / 10)
  )].sort((a, b) => b - a);
  const headingSizes = [];
  for (const size of sizes) {
    if (!headingSizes.length || size < headingSizes[headingSizes.length - 1] * 0.92) {
      headingSizes.push(size);
    }
  }
  headingSizes.length = Math.min(headingSizes.length, 6);

  const levelOf = (size) => {
    for (let i = 0; i < headingSizes.length; i++) {
      if (Math.abs(size - headingSizes[i]) <= headingSizes[i] * 0.08) return i + 1;
    }
    return 0;
  };

  // Headings are single lines. A size shared by the following or preceding
  // line means that size is wrapped prose (for example a footnote), not a title.
  const isHeadingLine = lines.map((line, index) => {
    const sameSize = (other) => other && Math.abs(other.size - line.size) < 0.6;
    return !sameSize(lines[index - 1]) && !sameSize(lines[index + 1]);
  });

  // A line that reaches most of the page width was wrapped; a short line ends
  // a paragraph. This is what separates prose from stacked short lines.
  const bodyEnds = lines
    .filter((line) => Math.abs(line.size - body) < 0.6)
    .map((line) => line.endX)
    .sort((a, b) => a - b);
  // Ignore outliers (a stray wide line) so wrapped prose is not mistaken for
  // short lines that end paragraphs.
  const globalMaxEnd = bodyEnds.length ? bodyEnds[Math.floor(bodyEnds.length * 0.9)] : 0;

  // The right margin is local, and only lines starting at roughly the same x
  // (same column) count. Otherwise multi-column prose looks like short lines.
  const maxEndAt = (index) => {
    const target = lines[index];
    const ends = lines.slice(Math.max(0, index - 40), index + 40)
      .filter((line) => Math.abs(line.size - body) < 0.6 && Math.abs(line.x - target.x) < body * 2)
      .map((line) => line.endX)
      .sort((a, b) => a - b);
    return ends.length ? ends[Math.floor(ends.length * 0.9)] : globalMaxEnd;
  };

  const blocks = [];
  let paragraph = null;
  let code = null;

  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    const text = pdfLineText(line);
    if (!text) continue;
    const prevLine = index > 0 ? lines[index - 1] : null;

    const isCode = line.parts.every((part) => part.code);
    if (isCode) {
      paragraph = null;
      if (code) {
        code.lines.push(pdfLineRaw(line));
      } else {
        code = { type: 'code', lines: [pdfLineRaw(line)] };
        blocks.push(code);
      }
      continue;
    }
    code = null;

    const size = Math.round(line.size * 10) / 10;
    const level = levelOf(line.size);

    if (level > 0 && isHeadingLine[index] && text.length <= 120) {
      blocks.push({ type: 'heading', text: `${'#'.repeat(level)} ${pdfLineText(line, true)}` });
      paragraph = null;
      continue;
    }

    const bullet = /^[•·◦▪●○]\s+/.exec(text);
    const numbered = /^(\d{1,3})[.)]\s+/.exec(text);
    if (bullet || numbered) {
      const content = bullet ? text.slice(bullet[0].length) : text.slice(numbered[0].length);
      blocks.push({ type: bullet ? 'ul' : 'ol', text: `${bullet ? '-' : `${numbered[1]}.`} ${content}` });
      paragraph = null;
      continue;
    }

    const gap = prevLine ? Math.abs(prevLine.y - line.y) : 0;
    const lineHeight = (prevLine ? prevLine.size : body) * 1.3;
    // A normal wrapped line continues the paragraph; extra leading, a page or
    // column jump, or a short line starts a new one.
    const started = !paragraph || line.blankBefore || Boolean(prevLine && prevLine.spread) ||
      line.spread ||
      gap > lineHeight * 1.35 ||
      gap < lineHeight * 0.55 ||
      (paragraph && prevLine && prevLine.endX < maxEndAt(index) * 0.85);

    if (started) {
      paragraph = { type: 'paragraph', lines: [], x: line.x };
      blocks.push(paragraph);
    }
    paragraph.lines.push(text);
  }

  let markdown = '';
  let previousType = null;
  const join = (previous, next) => (/[a-z]-$/.test(previous) && /^[a-z]/.test(next)
    ? previous + next
    : `${previous} ${next}`);

  for (const block of blocks) {
    let text;
    if (block.type === 'code') {
      const fence = fenceFor(block.lines.join('\n'));
      text = `${fence}\n${block.lines.join('\n')}\n${fence}`;
    } else {
      text = block.lines ? block.lines.reduce(join) : block.text;
    }
    const tight = block.type === previousType && (block.type === 'ul' || block.type === 'ol');
    const separator = markdown ? (tight ? '\n' : '\n\n') : '';
    markdown += separator + text;
    previousType = block.type;
  }
  return markdown ? `${markdown.trim()}\n` : '';
}

async function pdfToMarkdown(input) {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('This browser is too old. Please use a recent Chrome, Edge, Firefox or Safari.');
  }
  const raw = input instanceof ArrayBuffer ? input : await input.arrayBuffer();
  if (!raw.byteLength) throw new Error('The file is empty.');
  const bytes = new Uint8Array(raw);
  if (!PDF_LATIN1.decode(bytes.subarray(0, 1024)).includes('%PDF-')) {
    throw new Error('This is not a PDF file.');
  }

  const state = await pdfOpen(bytes);
  if (state.trailer.has('Encrypt')) throw new Error('Password-protected PDFs are not supported.');

  const pages = await pdfPages(state);
  if (!pages.length) throw new Error('No pages were found in this PDF.');

  const lines = [];
  for (let i = 0; i < pages.length; i++) {
    const start = lines.length;
    await pdfPageText(state, pages[i], lines, i);
    if (lines.length > start) lines[start].blankBefore = true;
  }

  const markdown = pdfLinesToMarkdown(lines);
  if (!markdown.trim()) {
    throw new Error('No selectable text was found. Scanned PDFs need OCR first.');
  }
  return markdown;
}

/* ------------------------------ api ------------------------------ */

async function convertFile(input) {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('This browser is too old. Please use a recent Chrome, Edge, Firefox or Safari.');
  }
  const buffer = input instanceof ArrayBuffer ? input : await input.arrayBuffer();
  if (!buffer.byteLength) throw new Error('The file is empty.');
  const head = new TextDecoder('latin1').decode(new Uint8Array(buffer, 0, Math.min(1024, buffer.byteLength)));
  if (head.includes('%PDF-')) return pdfToMarkdown(buffer);
  return docxToMarkdown(buffer);
}

async function docxToMarkdown(input) {
  const buffer = input instanceof ArrayBuffer ? input : await input.arrayBuffer();
  if (!buffer.byteLength) throw new Error('The file is empty.');

  const zip = await readZip(buffer);
  const documentBytes = await zip.read('word/document.xml');
  if (!documentBytes) throw new Error('This is not a Word .docx document.');

  const doc = parseXml(documentBytes);
  const [styles, numbering, rels] = await Promise.all([
    loadStyles(zip).catch(() => new Map()),
    loadNumbering(zip).catch(() => new Map()),
    loadRels(zip).catch(() => new Map()),
  ]);

  // Only load media that the document actually references.
  const imageTargets = new Set();
  for (const element of doc.getElementsByTagName('*')) {
    for (const attr of element.attributes) {
      if (attr.namespaceURI !== NS.r) continue;
      const rel = rels.get(attr.value);
      if (rel && !rel.external) imageTargets.add(rel.target);
    }
  }
  const images = await loadImages(zip, imageTargets);

  const body = doc.getElementsByTagNameNS(NS.w, 'body')[0];
  if (!body) throw new Error('The document is empty.');

  const ctx = {
    styles,
    numbering,
    rels,
    images,
    counters: new Map(),
    baseLevels: new Map(),
  };
  const out = [];
  renderBlocks(body, ctx, out);
  if (!out.length) return '';

  let markdown = '';
  let previousList = null;
  for (const block of out) {
    if (markdown) markdown += block.list && block.list === previousList ? '\n' : '\n\n';
    if (block.code) {
      const fence = fenceFor(block.text);
      markdown += `${fence}\n${block.text}\n${fence}`;
    } else {
      markdown += block.text;
    }
    previousList = block.list || null;
  }
  return `${markdown.trim()}\n`;
}

/* ------------------------------- ui ------------------------------ */

function initUI() {
  const input = document.getElementById('file');
  const drop = document.getElementById('drop');
  if (!input || !drop) return;

  const status = document.getElementById('status');
  const result = document.getElementById('result');
  const preview = document.getElementById('preview');
  const summary = document.getElementById('summary');
  const copy = document.getElementById('copy');
  const download = document.getElementById('download');
  const PREVIEW_LIMIT = 150000;

  let currentMarkdown = '';
  let objectUrl = null;
  let busy = false;

  const setStatus = (message, isError) => {
    status.textContent = message;
    status.classList.toggle('error', Boolean(isError));
  };

  const formatBytes = (size) => {
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
    return `${(size / 1024 / 1024).toFixed(1)} MB`;
  };

  async function convert(file) {
    if (busy || !file) return;
    if (!/\.(docx|pdf)$/i.test(file.name)) {
      setStatus('Please choose a .docx or .pdf file.', true);
      return;
    }
    busy = true;
    drop.classList.add('busy');
    setStatus('Converting…');
    const started = performance.now();
    try {
      const markdown = await convertFile(file);
      if (!markdown.trim()) throw new Error('No text was found in this document.');
      currentMarkdown = markdown;
      const elapsed = Math.round(performance.now() - started);
      const outputSize = new TextEncoder().encode(markdown).length;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
      objectUrl = URL.createObjectURL(new Blob([markdown], { type: 'text/markdown;charset=utf-8' }));
      download.href = objectUrl;
      download.download = `${file.name.replace(/\.(docx|pdf)$/i, '')}.md`;
      summary.textContent = `${file.name} · ${formatBytes(file.size)} → ${formatBytes(outputSize)} · ${elapsed} ms`;
      preview.textContent = markdown.length > PREVIEW_LIMIT
        ? `${markdown.slice(0, PREVIEW_LIMIT)}\n\n… preview truncated. Download or copy for the full file.`
        : markdown;
      result.hidden = false;
      setStatus(`Converted in ${elapsed} ms.`);
    } catch (error) {
      result.hidden = true;
      setStatus(error && error.message ? error.message : 'Could not convert this file.', true);
    } finally {
      busy = false;
      drop.classList.remove('busy');
      input.value = '';
    }
  }

  input.addEventListener('change', () => convert(input.files[0]));

  drop.addEventListener('dragenter', (event) => {
    event.preventDefault();
    drop.classList.add('over');
  });
  drop.addEventListener('dragover', (event) => {
    event.preventDefault();
    drop.classList.add('over');
  });
  drop.addEventListener('dragleave', (event) => {
    if (!drop.contains(event.relatedTarget)) drop.classList.remove('over');
  });
  drop.addEventListener('drop', (event) => {
    event.preventDefault();
    drop.classList.remove('over');
    convert(event.dataTransfer && event.dataTransfer.files[0]);
  });
  drop.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      input.click();
    }
  });

  // Never let a stray drop navigate the browser away from the app.
  window.addEventListener('dragover', (event) => event.preventDefault());
  window.addEventListener('drop', (event) => event.preventDefault());

  copy.addEventListener('click', async () => {
    if (!currentMarkdown) return;
    try {
      await navigator.clipboard.writeText(currentMarkdown);
      copy.textContent = 'Copied';
      setTimeout(() => { copy.textContent = 'Copy'; }, 1200);
    } catch {
      setStatus('Copy failed — select the preview and copy manually.', true);
    }
  });
}

if (typeof document !== 'undefined') initUI();
if (typeof window !== 'undefined') {
  window.docxToMarkdown = docxToMarkdown;
  window.pdfToMarkdown = pdfToMarkdown;
  window.convertFile = convertFile;
}
