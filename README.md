# docx / pdf → markdown

A tiny, dependency-free web app that converts Word `.docx` and `.pdf` files to downloadable Markdown. Everything runs in the browser — files never leave the device, and it works offline once loaded.

The whole app is two static files (~72 KB of JS, about 20 KB gzipped, no build step): `index.html` and `app.js`. There is no PDF.js, no mammoth, no CDN.

## Run locally

```sh
python3 -m http.server 8000
# open http://localhost:8000
```

## Deploy

Upload `index.html` and `app.js` to any static host — GitHub Pages, Netlify, Cloudflare Pages, S3, Caddy, nginx. No server or API needed.

## What converts

**DOCX**
- Title, headings (1–6), paragraphs, hard line breaks
- Bold, italic, strikethrough, underline (`<u>`), superscript/subscript
- Hyperlinks, inline code, fenced code blocks
- Bullet and numbered lists, including nesting
- Tables (first row becomes the Markdown header)
- Block quotes
- Inline images, embedded as base64 data URIs

**PDF**
- Text with headings inferred from font size, paragraphs, and basic lists
- Bold, italic, and inline/fenced code from font information
- Multi-page documents, multi-column layouts, and xref/object streams
- Rotated margin text is ignored
- Images and tables are extracted as plain text (or skipped) — PDF layout does not determine Markdown structure

## Not converted

DOCX: headers/footers, footnotes/endnotes, comments, text boxes, charts/SmartArt, tracked deletions, and EMF/WMF images are skipped.

PDF: scanned/image-only PDFs (no text layer), password-protected PDFs, footnotes, and visual table structure. Use OCR first for scans.

Legacy `.doc` files are rejected; save them as `.docx` first.

## Test

Requires Chrome/Chromium and Python:

```sh
./test/run.sh
```

`test/make-fixture.py` regenerates `test/fixture.docx`. The PDF fixtures (`fixture.pdf`, `objstm.pdf`, `columns.pdf`) cover classic xref, compressed xref/object streams, and two-column layout. `test/test.html` runs the converter against any of them in the browser.
