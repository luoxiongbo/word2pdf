# Word-to-PDF Operations Guide

This file keeps practical commands for day-to-day conversion and debugging.

## 1. Start Web Converter

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 converter_from_downloads.py
```

Open:

```text
http://localhost:5000
```

## 2. Node CLI Basic Conversion

```bash
node bin/docx2pdf.js \
  "/absolute/path/input.docx" \
  -o "/absolute/path/output.pdf" \
  --overwrite
```

## 3. Node CLI Batch Conversion

```bash
node bin/docx2pdf.js \
  "/absolute/path/docx-dir" \
  -o "/absolute/path/output-dir" \
  --overwrite
```

## 4. Native Engine Mode (CLI)

```bash
node bin/docx2pdf.js \
  "/absolute/path/input.docx" \
  -o "/absolute/path/output.pdf" \
  --engine native \
  --native-layout structured \
  --overwrite
```

## 5. LibreOffice Engine Mode (CLI)

```bash
node bin/docx2pdf.js \
  "/absolute/path/input.docx" \
  -o "/absolute/path/output.pdf" \
  --engine libreoffice \
  --overwrite
```

## 6. Smoke Test

```bash
npm test
```

## 7. PDF to Word (Python CLI)

Single file:

```bash
python3 pdf_to_word.py \
  "/absolute/path/input.pdf" \
  -o "/absolute/path/output.docx" \
  --overwrite
```

Batch directory:

```bash
python3 pdf_to_word.py \
  "/absolute/path/pdf-dir" \
  -o "/absolute/path/docx-dir" \
  --overwrite
```

## 8. Troubleshooting

### Symptom: textbox overlap in output

- Prefer Web converter path first (`converter_from_downloads.py`)
- Check diagnosis output (`X-Diagnosis` header)
- Compare CLI result with Web result to isolate preprocessing differences

### Symptom: `LibreOffice not found`

- Install LibreOffice and verify path
- Or pass explicit `--libreoffice <path>` in CLI mode

### Symptom: CJK text width mismatch

- Ensure proper Chinese fonts are installed
- Prefer consistent font environment across machines

### Symptom: native mode style differs from source

- This is expected for some complex layouts
- Try `--engine libreoffice` or Web mode for better fidelity
