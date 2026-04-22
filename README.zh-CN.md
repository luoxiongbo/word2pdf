# DocForge Convert (word2pdf)

[English](README.md) | 简体中文

一个本地化、开源的 Word 转 PDF 工具，重点面向中文文档场景与 WPS 来源文档的兼容转换。

仓库地址：[https://github.com/luoxiongbo/word2pdf](https://github.com/luoxiongbo/word2pdf)

- 不上传云端
- 不依赖付费 SaaS
- 同时支持 Web 页面和 Node CLI
- Web 模式内置 WPS 文本框重叠修复逻辑

## 为什么做这个项目

很多 `.docx`（尤其是 WPS/Office 混合流转产生的文档）在通用转换工具下容易出现排版错位、重叠、换行异常。

本项目提供两条实用路径：

1. `Web 转换器（Python + LibreOffice）`
- 适合交互式使用
- 对 WPS 文本框重叠问题有更强预处理

2. `Node CLI`
- 适合自动化和批处理
- 更容易集成进脚本/CI

## 功能对比

| 能力 | Web 转换器（`converter_from_downloads.py`） | Node CLI（`bin/docx2pdf.js`） |
|---|---|---|
| 本地转换 | 是 | 是 |
| LibreOffice 后端 | 是 | 是（`--engine libreoffice`） |
| 内置非 LO 渲染路径 | 否 | 是（`--engine native`） |
| WPS 文本框重叠兼容 | 更强（anchor + inline 文本框处理） | 基础（`lineRule` 归一化） |
| 目录批量转换 | 可通过接口/脚本实现 | 原生支持 |
| 浏览器上传下载 UI | 是 | 否 |

## 项目结构

```text
.
├── bin/                          # Node CLI 入口
├── lib/                          # Node 转换核心模块
├── scripts/                      # 辅助脚本
├── test/                         # Node 冒烟测试
├── docs/
│   ├── architecture.md           # 架构与模块职责
│   ├── operations.md             # 常用操作与排障
│   ├── release-checklist.md      # 开源发布检查清单
│   └── images/
│       └── README.md             # 截图目录说明
├── converter_from_downloads.py   # Web 转换器 + 内嵌前端
├── CONTRIBUTING.md
├── CODE_OF_CONDUCT.md
├── SECURITY.md
├── requirements.txt              # Web 模式 Python 依赖
├── package.json                  # Node 包信息
└── README.md
```

## 快速开始

### 方案 A：Web 转换器（推荐，版式敏感文档优先）

前置要求：
- macOS/Linux
- 已安装 LibreOffice
- Python 3.10+

安装依赖：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

启动服务：

```bash
python3 converter_from_downloads.py
```

打开：

```text
http://localhost:5000
```

### 方案 B：Node CLI

前置要求：
- Node.js >= 16
- LibreOffice（用于 `--engine libreoffice`）
- 可选 Chrome/Chromium（用于 `--engine native`）

安装：

```bash
npm install
```

单文件转换：

```bash
node bin/docx2pdf.js \
  "/path/to/input.docx" \
  -o "/path/to/output.pdf" \
  --overwrite
```

目录批量转换：

```bash
node bin/docx2pdf.js \
  "/path/to/docx-dir" \
  -o "/path/to/output-dir" \
  --overwrite
```

## Web API

### `POST /convert`

表单字段：
- `file`：`.doc` 或 `.docx`

返回：
- PDF 二进制流
- `X-Diagnosis` 响应头（预处理/转换诊断摘要）

## 页面截图

当前已放占位图，请替换为真实截图：

- 目标路径：`docs/images/web-ui-screenshot.png`
- README 引用：

```markdown
![Web UI Screenshot](docs/images/web-ui-screenshot.png)
```

当前仓库里该路径是 `1x1` 占位 PNG，请直接同名覆盖。

![Web UI Screenshot](docs/images/web-ui-screenshot.png)

推荐截图内容：
- 主上传区域与状态区
- 顶部品牌/标题区
- 一次成功转换后的结果状态

## 转换原则与边界

本项目目标是：在常见简历/表单/表格模板中尽可能接近原稿版式。

注意：
- 严格像素级 1:1 通常需要 Microsoft Word 原生渲染引擎。
- LibreOffice/Native HTML 渲染属于高质量近似，但对复杂 DOCX 构造不保证数学级一致。
- WPS 来源文档中的 `textbox` 与 VML fallback 是重叠高发点，Web 模式已做针对性结构修复。

## 常用命令

详见：
- [docs/operations.md](docs/operations.md)

## 开发

运行冒烟测试：

```bash
npm test
```

代码检查（若 ESLint 配置完善）：

```bash
npm run lint
```

## 开源发布检查

发布前建议核对：
- 包信息（`author`、`repository.url`、关键词）
- LICENSE 作者信息
- 截图与示例
- 文档准确性

详见：
- [docs/release-checklist.md](docs/release-checklist.md)

## 参与贡献

请先阅读：
- [CONTRIBUTING.md](CONTRIBUTING.md)
- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [SECURITY.md](SECURITY.md)

## License

MIT，见 [LICENSE](LICENSE)。
