# Word-to-PDF & PDF-to-Word 转换器（word-to-pdf）

[English](README.md) | 简体中文

本地开源文档转换工具：
- `Word -> PDF`（`.doc/.docx -> .pdf`）
- `PDF -> Word`（`.pdf -> .docx`）

关键词：`word-to-pdf`、`pdf-to-word`、`docx-to-pdf`、`pdf-to-docx`、`Word 转 PDF`、`PDF 转 Word`

仓库地址：[https://github.com/luoxiongbo/word-to-pdf](https://github.com/luoxiongbo/word-to-pdf)

## 功能

- 本地优先，可自托管部署
- Web 版 Word->PDF（包含 WPS 文本框重叠修复）
- Node CLI Word->PDF（适合脚本/批处理/CI）
- Python CLI PDF->Word（结构化分析）
- 对项目生成的 PDF 支持精确回转还原

## 工具入口

| 工具 | 方向 | 入口 | 适用场景 |
|---|---|---|---|
| Web 转换器 | Word -> PDF | `converter_from_downloads.py` | 交互式使用、WPS 文档 |
| Node CLI | Word -> PDF | `bin/docx2pdf.js` | 自动化批处理 |
| Python CLI | PDF -> Word | `pdf_to_word.py` | PDF 回转 Word |

## 快速开始

### 1）安装依赖

```bash
# Node 依赖
npm install

# Python 依赖
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2）Word -> PDF（Web）

```bash
python3 converter_from_downloads.py
# 打开 http://localhost:5000
```

### 3）Word -> PDF（Node CLI）

```bash
node bin/docx2pdf.js \
  "/path/to/input.docx" \
  -o "/path/to/output.pdf" \
  --overwrite
```

### 4）PDF -> Word（Python CLI）

```bash
python3 pdf_to_word.py \
  "/path/to/input.pdf" \
  -o "/path/to/output.docx" \
  --overwrite
```

## 低成本上线（推荐）

推荐用 Cloud Run（按量计费，空闲自动缩到 0）。

```bash
# 1）安装并登录 gcloud
gcloud auth login
gcloud auth application-default login

# 2）首次启用 API
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com

# 3）部署
PROJECT_ID="你的 GCP 项目 ID" ./scripts/deploy_cloud_run.sh
```

部署成功后会输出公网地址（`https://...run.app`），用户可直接访问。

## 精确 1:1 还原规则

`pdf_to_word.py` 精确还原优先级：
1. `--source-docx` 手动指定源文件
2. PDF 内嵌源 DOCX
3. 同目录 sidecar DOCX（按文件名规则自动匹配）
4. 都没有时退化为结构分析

严格模式（不允许退化）：

```bash
python3 pdf_to_word.py \
  "/path/to/input.pdf" \
  -o "/path/to/output.docx" \
  --overwrite \
  --strict-1to1
```

对外部 PDF 强制结构分析：

```bash
python3 pdf_to_word.py \
  "/path/to/input.pdf" \
  -o "/path/to/output.docx" \
  --overwrite \
  --no-embedded-restore \
  --no-sidecar-restore
```

## 截图

![Web UI 截图](docs/images/web-ui-screenshot.png)

## 目录结构

```text
.
├── converter_from_downloads.py   # Web Word->PDF
├── pdf_to_word.py                # PDF->Word CLI
├── bin/docx2pdf.js               # Node Word->PDF CLI
├── lib/                          # Node 转换内部模块
├── docs/                         # 架构 / 操作 / 发布检查
└── README.md / README.zh-CN.md
```

## 边界说明

- 只有在能恢复源 DOCX（内嵌/sidecar/显式指定）时，才可实现严格 1:1。
- 对通用外部 PDF，仍是尽量还原的结构重建，不保证完全一致。

## 相关文档

- [操作手册](docs/operations.md)
- [Cloud Run 部署](docs/deploy-cloud-run.md)
- [架构说明](docs/architecture.md)
- [发布检查](docs/release-checklist.md)
- [贡献指南](CONTRIBUTING.md)
- [安全策略](SECURITY.md)

## License

MIT，见 [LICENSE](LICENSE)。
