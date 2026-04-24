# Word-to-PDF Converter (word2pdf) | DOCX-to-PDF

[English](README.md) | [简体中文](README.zh-CN.md)

A local, open-source Word-to-PDF conversion toolkit focused on Chinese document workflows and WPS-origin templates.

Keywords: `word-to-pdf`, `docx-to-pdf`, `word to pdf converter`, `wps to pdf`

Repository: https://github.com/luoxiongbo/word-to-pdf

- No cloud upload
- No paid SaaS dependency
- Supports Web UI and CLI
- Includes WPS text-box overlap compatibility fixes (Web mode)
- Includes a local PDF-to-Word (`.pdf -> .docx`) utility

## Why this project

Many `.docx` files (especially from WPS/Office mixed workflows) can produce layout issues when converted with generic tools.
This project provides two practical paths:

1. `Web Converter (Python + LibreOffice)`
- Best for interactive use
- Includes stronger preprocessing for WPS text-box overlap issues

2. `Node CLI`
- Best for automation and batch jobs
- Easier integration into scripts/CI

## Feature Matrix

| Capability | Web Converter (`converter_from_downloads.py`) | Node CLI (`bin/docx2pdf.js`) |
|---|---|---|
| Local conversion | Yes | Yes |
| LibreOffice backend | Yes | Yes (`--engine libreoffice`) |
| Built-in non-LO rendering path | No | Yes (`--engine native`) |
| WPS textbox overlap preprocessing | Stronger (anchor + inline textbox handling) | Basic (`lineRule` normalization) |
| Batch directory conversion | API loop / custom scripts | Native support |
| Browser upload/download UI | Yes | No |

## Project Structure

```text
.
├── bin/                          # Node CLI entry
├── lib/                          # Node conversion core modules
├── scripts/                      # Helper scripts
├── test/                         # Node smoke tests
├── docs/
│   ├── architecture.md           # Architecture and module responsibilities
│   ├── operations.md             # Daily operations and troubleshooting
│   ├── release-checklist.md      # Open-source release checklist
│   └── images/
│       └── README.md             # Screenshot location / placeholder
├── converter_from_downloads.py   # Web converter + embedded frontend
├── pdf_to_word.py                # PDF -> Word converter (Python CLI)
├── CONTRIBUTING.md
├── CODE_OF_CONDUCT.md
├── SECURITY.md
├── requirements.txt              # Python deps for web mode
├── package.json                  # Node package metadata
├── README.zh-CN.md               # Chinese documentation
└── README.md
```

## Quick Start

### Option A: Web Converter (Recommended for layout-sensitive docs)

Prerequisites:
- macOS/Linux
- LibreOffice installed
- Python 3.10+

Install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Run server:

```bash
python3 converter_from_downloads.py
```

Open:

```text
http://localhost:5000
```

### Option B: Node CLI

Prerequisites:
- Node.js >= 16
- LibreOffice (for `--engine libreoffice`)
- Optional Chrome/Chromium (for `--engine native`)

Install:

```bash
npm install
```

Single file conversion:

```bash
node bin/docx2pdf.js \
  "/path/to/input.docx" \
  -o "/path/to/output.pdf" \
  --overwrite
```

Batch conversion:

```bash
node bin/docx2pdf.js \
  "/path/to/docx-dir" \
  -o "/path/to/output-dir" \
  --overwrite
```

### Option C: PDF to Word (Python CLI)

Single file:

```bash
python3 pdf_to_word.py "/path/to/input.pdf" -o "/path/to/output.docx" --overwrite
```

Batch directory:

```bash
python3 pdf_to_word.py "/path/to/pdf-dir" -o "/path/to/docx-dir" --overwrite
```

## Web API

### `POST /convert`

Form-data:
- `file`: `.doc` or `.docx`

Response:
- Binary PDF stream
- Header `X-Diagnosis`: preprocessing/conversion diagnosis summary

## Screenshot

Current placeholder (please replace with your real UI capture):

- Target path: `docs/images/web-ui-screenshot.png`
- README reference:

```markdown
![Web UI Screenshot](docs/images/web-ui-screenshot.png)
```

This repository currently includes a `1x1` PNG placeholder at that path.
Replace it with a real screenshot using the same file name:

![Web UI Screenshot](docs/images/web-ui-screenshot.png)

Recommended screenshot content:
- Main upload area and status panel
- Brand/header section
- One successful conversion result state

## Conversion Principles & Limits

This project aims to maximize practical fidelity for common resume/form/table templates.

Important:
- True 1:1 pixel-perfect output requires Microsoft Word's own rendering engine.
- LibreOffice/native HTML rendering are high-quality approximations, but not mathematically identical for every complex DOCX construct.
- WPS-origin `textbox` and fallback VML structures are a known source of overlap issues; Web mode includes targeted structural mitigation.

## Typical Commands

See detailed commands in:
- [docs/operations.md](docs/operations.md)

## Development

Run smoke tests:

```bash
npm test
```

Lint (if ESLint config is present):

```bash
npm run lint
```

## Open-Source Release Checklist

Before publishing, verify:
- package metadata (`author`, `repository.url`, keywords)
- license owner name
- screenshot and examples
- docs accuracy

Detailed list:
- [docs/release-checklist.md](docs/release-checklist.md)

## Contributing

Please read:
- [CONTRIBUTING.md](CONTRIBUTING.md)
- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [SECURITY.md](SECURITY.md)

## License

MIT. See [LICENSE](LICENSE).
