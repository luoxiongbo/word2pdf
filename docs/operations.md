# Word-to-PDF & PDF-to-Word Operations

Minimal, high-frequency commands.

## Setup

```bash
npm install
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Word -> PDF (Web)

```bash
python3 converter_from_downloads.py
# open http://localhost:5000
```

## Word -> PDF (Node CLI)

```bash
# single file
node bin/docx2pdf.js \
  "/absolute/path/input.docx" \
  -o "/absolute/path/output.pdf" \
  --overwrite

# directory
node bin/docx2pdf.js \
  "/absolute/path/docx-dir" \
  -o "/absolute/path/output-dir" \
  --overwrite
```

## PDF -> Word (Python CLI)

```bash
# default (exact restore when possible, otherwise structured analysis)
python3 pdf_to_word.py \
  "/absolute/path/input.pdf" \
  -o "/absolute/path/output.docx" \
  --overwrite

# strict 1:1 (fail if exact restore unavailable)
python3 pdf_to_word.py \
  "/absolute/path/input.pdf" \
  -o "/absolute/path/output.docx" \
  --overwrite \
  --strict-1to1

# explicit source DOCX for exact restore
python3 pdf_to_word.py \
  "/absolute/path/input.pdf" \
  -o "/absolute/path/output.docx" \
  --overwrite \
  --source-docx "/absolute/path/original.docx"

# force structured analysis (no exact restore)
python3 pdf_to_word.py \
  "/absolute/path/input.pdf" \
  -o "/absolute/path/output.docx" \
  --overwrite \
  --no-embedded-restore \
  --no-sidecar-restore
```

## Quick Checks

```bash
npm test
python3 -m py_compile converter_from_downloads.py pdf_to_word.py
```

## Troubleshooting

- `LibreOffice not found`:
  install LibreOffice, or pass `--libreoffice` in Node CLI mode.
- PDF->Word not 1:1:
  use `--strict-1to1` to detect missing exact-restore source path.
- WPS overlap issues:
  prefer Web converter (`converter_from_downloads.py`) for Word->PDF.
