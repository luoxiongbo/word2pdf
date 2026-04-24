# Word-to-PDF & PDF-to-Word Architecture

## Overview

This repo has three production paths:

1. Web Word->PDF (`converter_from_downloads.py`)
2. Node CLI Word->PDF (`bin/docx2pdf.js` + `lib/*`)
3. Python CLI PDF->Word (`pdf_to_word.py`)

## 1) Web Word->PDF

Pipeline:
1. Upload `.doc/.docx`
2. Preprocess WPS-specific textboxes/frames
3. Convert by LibreOffice headless
4. Embed source `.docx` into PDF for exact round-trip
5. Return PDF + diagnosis header

Strength:
- Best path for WPS-origin overlap fixes.

## 2) Node CLI Word->PDF

Core files:
- `bin/docx2pdf.js`
- `lib/converter.js`
- `lib/wpsCompat.js`
- `lib/nativeEngine.js`

Modes:
- `--engine libreoffice`
- `--engine native`

Strength:
- Batch automation and CI integration.

## 3) Python CLI PDF->Word

Entry:
- `pdf_to_word.py`

Exact restore priority:
1. `--source-docx`
2. Embedded source DOCX in PDF
3. Sidecar DOCX in same directory
4. Structured analysis fallback

Structured analysis (fallback):
- Heading/body/list classification
- Indentation and paragraph spacing reconstruction
- Basic image block insertion

## Fidelity Boundary

- Strict 1:1 is available only when exact source DOCX can be restored.
- Generic external PDFs are best-effort structural reconstruction.
