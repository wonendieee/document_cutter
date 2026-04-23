# Document Cutter - Dify 文档切分插件

一个 Dify Tool 插件，用于在 Workflow / Agent 中按**页码范围**裁剪 **PDF / Word / Excel** 文档，支持输出文本块或**保持原格式的二进制文件**。

## 功能特性

- **多格式支持**：PDF (`.pdf`)、Word (`.docx`)、Excel (`.xlsx` / `.xls`)
- **两种输出模式**
  - `page_text` — 按页提取文本块（保留图片/表格元数据）
  - `page_file` — 按页截取后输出**同格式的新文件**（PDF→PDF，docx→docx，xlsx→xlsx），下游节点可直接作为 File 变量使用
- **灵活的页码范围**：支持 `1-10`、`5-`、`-3`、`7`、空（全部）多种写法
- **高性能**：`page_file` 模式直接操作文件对象（PDF 复制页对象、Excel 删除 Sheet、Word 删除段落范围），毫秒级完成
- **本地处理**：所有解析在插件运行时内完成，无第三方 API 依赖

## 项目结构

```
document_cutter/
├── _assets/icon.svg
├── provider/
│   ├── document_cutter.yaml
│   └── document_cutter.py
├── tools/
│   ├── split_document.yaml
│   └── split_document.py
├── core/
│   ├── parsers/
│   │   ├── pdf_parser.py        # PyMuPDF
│   │   ├── word_parser.py       # python-docx
│   │   └── excel_parser.py      # openpyxl
│   ├── splitters/
│   │   └── page_splitter.py     # page_text 模式入口
│   └── extractors/
│       └── file_extractor.py    # page_file 模式入口
├── main.py
├── manifest.yaml
├── requirements.txt
├── .env.example
├── package.ps1                  # Windows 一键打包
└── PRIVACY.md
```

## 开发调试

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置远程调试

复制 `.env.example` 为 `.env`，填入 Dify 平台的调试信息：

```
INSTALL_METHOD=remote
REMOTE_INSTALL_URL=debug.dify.ai:5003
REMOTE_INSTALL_KEY=your-key
```

### 3. 启动

```bash
python -m main
```

## 打包为 .difypkg

### Windows

```powershell
.\package.ps1
```

### macOS / Linux

```bash
# 上级目录执行
dify plugin package ./document_cutter
```

生成的 `document_cutter.difypkg` 可通过 Dify 的 **插件管理 → 安装插件 → 本地文件上传** 安装到任意 Dify 实例（含内网）。

## 使用示例

### 参数说明

| 参数 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `file` | file | 是 | 待处理的文档文件 |
| `split_mode` | select | 是 | `page_text`（文本块）或 `page_file`（同格式文件） |
| `page_range` | string | 否 | 1-based 闭区间页码。示例：`1-10`、`5-`、`-3`、`7`、空=全部 |
| `pages_per_chunk` | number | 否 | 仅 `page_text` 模式生效，每 N 页合并为一块，默认 1 |
| `delivery_mode` | select | 否 | 仅 `page_file` 模式生效。`blob`=Dify 原生文件返回；`upload_link`=上传到外部服务并返回下载链接（推荐）。默认 `upload_link` |
| `output_filename` | string | 否 | `page_file` 的自定义基础文件名（不含扩展名），留空自动命名 |

### 页码语法

| 填写 | 含义 |
|---|---|
| `1-2` | 第 1~2 页 |
| `3-5` | 第 3~5 页 |
| `5-` | 第 5 页到末尾 |
| `-3` | 前 3 页 |
| `7` | 只要第 7 页 |
| 空 | 全部 |

Word 的"页"按**手动分页符**（`<w:br w:type="page">`）计数；Excel 的"页"按 **Sheet 序号**。

---

### 模式 1：`page_text` 输出文本块

下游节点接 `json` 输出，结构：

```json
{
  "total_chunks": 2,
  "file_name": "example.docx",
  "file_type": "docx",
  "split_mode": "page_text",
  "chunks": [
    {
      "index": 0,
      "content": "第一章 项目概况\n本项目...[IMAGE:image_1]\n| 表头 | ... |",
      "metadata": {
        "page": 1,
        "images": [
          {
            "id": "image_1",
            "mime_type": "image/png",
            "base64": "iVBORw0KG...",
            "kind": "image"
          }
        ],
        "char_count": 1520
      }
    }
  ]
}
```

`pages_per_chunk > 1` 时相邻页合并，metadata 会用 `page_start` / `page_end` 替代 `page`。

---

### 模式 2：`page_file` 输出同格式文件

根据 `delivery_mode` 有两种交付方式：

#### 2.1 `delivery_mode=blob`（Dify 原生）

由 Dify 把文件写入其文件服务，下游节点通过 **Files** 变量接收。
**依赖 Dify `FILES_URL` 环境变量配置正确**：容器内部的文件服务地址与浏览器可访问地址必须一致（内网里这两个通常不一致，会导致下载链接 404）。

#### 2.2 `delivery_mode=upload_link`（推荐，绕过 Dify 文件服务）

插件把生成的文件 `POST` 到**外部 HTTP 上传服务**，拿到一个公网/内网可直接访问的下载链接返回给下游。**不依赖 Dify 的 `FILES_URL`**，内网部署更稳。

需要在插件的 **Provider 凭据** 中配置：

| 凭据 | 必填 | 说明 |
|---|---|---|
| `upload_url` | 是 | 上传端点，如 `http://host.docker.internal:8000/upload` |
| `upload_token` | 否 | 若服务需要 Bearer 鉴权 |
| `upload_file_field_name` | 否 | multipart 中文件字段名，默认 `file` |
| `response_download_url_field` | 否 | 响应 JSON 里下载链接字段名，默认 `download_url` |
| `response_file_name_field` | 否 | 响应 JSON 里文件名字段名，默认 `file_name` |
| `upload_headers_json` | 否 | 额外请求头，JSON 对象 |
| `upload_form_data_json` | 否 | 额外表单字段，JSON 对象 |

**上传服务约定**（可直接复用 `generic_excel_template_filler/service/minimal_upload_service.py`，注意把下载端点的 `media_type` 改为按扩展名返回或 `application/octet-stream`）：

- `POST {upload_url}`：multipart/form-data，文件字段名即 `upload_file_field_name`；可额外读取 `desired_name` 作为落盘文件名。
- 返回 JSON：`{"download_url": "http://...", "file_name": "xxx"}`（字段名与上面凭据对应）。

#### 输出 JSON 元信息

```json
{
  "file_name": "example_p1-3.docx",
  "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "size_bytes": 45812,
  "source_file": "example.docx",
  "page_range": "1-3",
  "delivery_mode": "upload_link",
  "download_url": "http://host.docker.internal:8000/files/example_p1-3.docx",
  "returned_file_name": "example_p1-3.docx"
}
```

`upload_link` 模式下还会通过 `create_variable_message` 把同名字段注入 Workflow 变量（包括 `download_url`），下游节点可用 `{{download_url}}` 直接引用。

#### 输出文件名规则

`{原文件名}_{tag}{原扩展名}`，`tag` 举例：
- `page_range=1-3` → `example_p1-3.docx`
- `page_range=7` → `example_p7.docx`
- `page_range=5-` → `example_p5-end.docx`
- 空 → `example_all.docx`

传 `output_filename` 时覆盖上述规则，只保留原扩展名。

---

## 典型用法

### 场景 A：配合上游章节识别节点，按章节输出独立 Word 文件

1. 上游代码节点分析文档，输出每个章节的起止页码（如 `{ "start": 5, "end": 12 }`）
2. 循环节点调用本插件，`split_mode=page_file`、`page_range={{start}}-{{end}}`
3. 每次输出一个覆盖该章节页码的独立 docx，供下游分别处理

### 场景 B：只处理封面和目录

`split_mode=page_file`、`page_range=1-2` → 输出一个只含前 2 页的 PDF/docx。

### 场景 C：快速预览前几页文本

`split_mode=page_text`、`page_range=1-3` → 得到前 3 页的文本数组，下游用 LLM 总结。

## 能力矩阵

| 内容 | PDF | Word | Excel |
|---|---|---|---|
| 按页提取文本 | ✓ | ✓ | ✓（Sheet）|
| 按页输出同格式文件 | ✓（直接拷贝页对象） | ✓（删除范围外段落/表格）| ✓（删除未选中 Sheet）|
| 文本中的图片 | ✓ (base64) | ✓ (base64) | — |
| 表格 | 部分 | ✓ (Markdown) | ✓ (Markdown) |

### Word 页码的局限

docx 只有**手动分页符**是稳定的页边界。Word 运行时根据字体/页面大小自动排版的分页**不会写进 XML**，本插件无法感知。若你的 docx 基本没有手动分页符，输出会退化成"整篇文档只有 1 页"。此时可以改用上游先转 PDF 再切的链路。

## 许可证

[MIT](LICENSE)
