#!/usr/bin/env bash
# Smoke test: converts the .docx and .pdf fixtures in headless Chrome and checks key output.
set -euo pipefail
cd "$(dirname "$0")/.."

python3 test/make-fixture.py >/dev/null

python3 -m http.server 8765 >/dev/null 2>&1 &
server=$!
trap 'kill "$server" 2>/dev/null || true' EXIT
sleep 1

chrome=$(command -v google-chrome || command -v chromium || command -v chromium-browser || true)
if [ -z "$chrome" ]; then
  echo "SKIP: Chrome/Chromium not found"
  exit 0
fi

out=/tmp/docx2md-test-dom.html

check() {
  if ! grep -qF -- "$1" "$out"; then
    echo "FAIL: missing output: $1"
    exit 1
  fi
}

convert() {
  "$chrome" --headless=new --disable-gpu --no-sandbox --virtual-time-budget=20000 \
    --dump-dom "http://localhost:8765/test/test.html?$1" >"$out" 2>/dev/null

  if grep -qF '<pre id="out">ERROR:' "$out"; then
    echo "FAIL: conversion threw an error ($1)"
    sed -n '/<pre id="out">/,/<\/pre>/p' "$out"
    exit 1
  fi
}

convert 'file=fixture.docx'
check '# Fixture Document'
check 'Plain **bold**, *italic* and ~~struck~~ text with a [link](https://example.com).'
check 'First line'
check '**spaced bold,** *then italic.*'
check '- Bullet one'
check '    - Nested bullet'
check '1. Step one'
check '| a\|b | 1 |'
check 'const answer = 42;'
check '![a red dot](data:image/png;base64,'

convert 'file=fixture.pdf&pdf=1'
check '# Project Notes'
check 'This is **important**, this is *subtle*, and here is a link.'
check '- First item'
check '1. Install'
check 'Alice 9 fast'
check 'Less is more.'
check 'const x = 1;'

convert 'file=objstm.pdf&pdf=1'
check '# Object stream PDF'
check 'Text stored in compressed object streams.'

convert 'file=columns.pdf&pdf=1'
check '# Two Column Document'
check 'Left column first paragraph begins here and keeps going for a while'
check 'Right column second paragraph ends the document.'

echo "PASS: docx + pdf → markdown conversion"
