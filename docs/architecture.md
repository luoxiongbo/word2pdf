# Word-to-PDF Architecture Overview

This project has two independent conversion interfaces that share a common mission: local DOCX-to-PDF conversion.
It also includes a standalone local PDF-to-DOCX conversion script.

## 1) Web Converter Path

Entry:
- `converter_from_downloads.py`

Pipeline:
1. Receive uploaded `.doc/.docx`
2. Preprocess XML for problematic textbox/frame patterns
3. Convert using LibreOffice headless
4. Return PDF stream and diagnosis header

Key traits:
- Optimized for WPS-origin overlap scenarios
- Stronger textbox structural handling than current Node CLI compat layer

## 2) Node CLI Path

Entry:
- `bin/docx2pdf.js`

Core modules:
- `lib/converter.js`: main orchestration for LibreOffice conversion
- `lib/wpsCompat.js`: lightweight WPS preprocessing
- `lib/nativeEngine.js`: native rendering path (`--engine native`)
- `lib/fonts.js` / `lib/loFontSetup.js`: font/profile setup
- `lib/preflight.js`: runtime dependency detection

Modes:
- `--engine libreoffice`: DOCX -> LibreOffice -> PDF
- `--engine native`: DOCX XML parsing -> HTML/PDF rendering strategy

## 3) PDF -> DOCX Path

Entry:
- `pdf_to_word.py`

Pipeline:
1. Read PDF pages with PyMuPDF (`fitz`)
2. Extract text blocks in reading order
3. Write paragraphs to `.docx` via `python-docx`
4. Attempt to embed image blocks when decodable

## Design Notes

- Web and CLI are intentionally decoupled to reduce cross-impact risk.
- Web mode currently carries the more aggressive textbox overlap mitigation.
- CLI prioritizes automation and scriptability.

## Known Fidelity Boundary

No non-Word rendering path can guarantee universal pixel-identical output for all advanced DOCX constructs.
This project focuses on practical, high-fidelity output with deterministic and inspectable behavior.
