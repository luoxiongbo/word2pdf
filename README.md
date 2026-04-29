# Word-to-PDF & PDF-to-Word Converter (word-to-pdf)

[English](README.md) | [简体中文](README.zh-CN.md)

Local, open-source document conversion toolkit:
- `Word -> PDF` (`.doc/.docx -> .pdf`)
- `PDF -> Word` (`.pdf -> .docx`)

Keywords: `word-to-pdf`, `pdf-to-word`, `docx-to-pdf`, `pdf-to-docx`, `wps-to-pdf`, `word-to-pdf converter`

Repository: [https://github.com/luoxiongbo/word-to-pdf](https://github.com/luoxiongbo/word-to-pdf)

## Features

- Local-first and self-hostable
- Web Word-to-PDF converter with WPS textbox overlap fixes
- Node CLI Word-to-PDF converter for scripts/automation
- Python CLI PDF-to-Word converter with structure analysis
- Exact round-trip restore support for generated PDFs (when source DOCX is available)

## Tools

| Tool | Direction | Entry | Best for |
|---|---|---|---|
| Web converter | Word -> PDF | `converter_from_downloads.py` | Interactive conversion, WPS-heavy docs |
| Node CLI | Word -> PDF | `bin/docx2pdf.js` | Batch/automation/CI |
| Python CLI | PDF -> Word | `pdf_to_word.py` | PDF back-conversion |

## Quick Start

### 1) Install dependencies

```bash
# Node deps
npm install

# Python deps
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2) Word -> PDF (Web)

```bash
python3 converter_from_downloads.py
# open http://localhost:5000
```

### 3) Word -> PDF (Node CLI)

```bash
node bin/docx2pdf.js \
  "/path/to/input.docx" \
  -o "/path/to/output.pdf" \
  --overwrite
```

### 4) PDF -> Word (Python CLI)

```bash
python3 pdf_to_word.py \
  "/path/to/input.pdf" \
  -o "/path/to/output.docx" \
  --overwrite
```

## Deploy Online (Lowest Cost)

Recommended: Cloud Run (pay-as-you-go, auto scale to zero).

```bash
# 1) Install and login
gcloud auth login
gcloud auth application-default login

# 2) Enable required APIs (first time only)
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com

# 3) Deploy
PROJECT_ID="your-gcp-project-id" ./scripts/deploy_cloud_run.sh
```

After deploy, Cloud Run prints a public URL (`https://...run.app`) that users can open directly.

## Exact 1:1 Restore Rules

`pdf_to_word.py` exact restore order:
1. `--source-docx` explicit source path
2. Embedded source DOCX in PDF attachment
3. Sidecar DOCX in same directory (name matching)
4. Fallback to structured analysis

Strict mode (do not allow fallback):

```bash
python3 pdf_to_word.py \
  "/path/to/input.pdf" \
  -o "/path/to/output.docx" \
  --overwrite \
  --strict-1to1
```

Force analysis mode for external PDFs:

```bash
python3 pdf_to_word.py \
  "/path/to/input.pdf" \
  -o "/path/to/output.docx" \
  --overwrite \
  --no-embedded-restore \
  --no-sidecar-restore
```

## Screenshot

![Web UI Screenshot](docs/images/web-ui-screenshot.png)

## Project Structure

```text
.
├── converter_from_downloads.py   # Web Word->PDF
├── pdf_to_word.py                # PDF->Word CLI
├── bin/docx2pdf.js               # Node Word->PDF CLI
├── lib/                          # Node conversion internals
├── docs/                         # architecture / operations / checklist
└── README.md / README.zh-CN.md
```

## Limits

- Strict 1:1 is guaranteed only when source DOCX can be restored (embedded/sidecar/explicit).
- For generic external PDFs, output is best-effort structural reconstruction.

## Docs

- [Operations](docs/operations.md)
- [Cloud Run deployment](docs/deploy-cloud-run.md)
- [Architecture](docs/architecture.md)
- [Release checklist](docs/release-checklist.md)
- [Contributing](CONTRIBUTING.md)
- [Security](SECURITY.md)

## License

MIT, see [LICENSE](LICENSE).
