#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 || $# -gt 2 ]]; then
  echo "Usage: $0 <input.docx> [output.pdf]"
  exit 1
fi

INPUT="$1"
OUTPUT="${2:-${INPUT%.docx}.native.pdf}"

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
CHROME_BIN="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

node "$ROOT_DIR/bin/docx2pdf.js" \
  --engine native \
  --native-layout structured \
  --chrome "$CHROME_BIN" \
  "$INPUT" \
  -o "$OUTPUT" \
  --overwrite

echo "Generated: $OUTPUT"
