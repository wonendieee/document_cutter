# Document Cutter - Dify 文档切分插件

一个 Dify Tool 插件，用于在 Workflow / Agent 中将 **Word / PDF / Excel** 文档切分为结构化文本块，支持图片、表格、图表的完整保留。

## 功能特性

- **多格式支持**：PDF (`.pdf`)、Word (`.docx`)、Excel (`.xlsx` / `.xls`)
- **三种切分模式**
  - `page` — 按页/Sheet 切分
  - `semantic` — 按标题层级/段落语义切分（支持 max_chunk_size 和 overlap）
  - `anchor` — 按上游提供的锚点 JSON 精确切分（配合自定义标题识别工作流）
- **图片/图表保留**（Word）：内嵌图片和图表 fallback 预览图以 base64 形式输出到 `metadata.images`，正文插入 `[IMAGE:image_N]` 占位符
- **表格保留**（Word / Excel）：自动转为 Markdown 表格嵌入 chunk 内容
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
│   │   ├── pdf_parser.py       # PyMuPDF
│   │   ├── word_parser.py      # python-docx
│   │   └── excel_parser.py     # openpyxl
│   └── splitters/
│       ├── page_splitter.py
│       ├── semantic_splitter.py
│       └── anchor_splitter.py
├── main.py
├── manifest.yaml
├── requirements.txt
├── .env.example
├── package.ps1                 # Windows 一键打包
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

| 参数 | 类型 | 说明 |
|---|---|---|
| `file` | file | 待切分的文档文件 |
| `split_mode` | select | `page` / `semantic` / `anchor` |
| `max_chunk_size` | number | 语义模式下单块最大字符数（默认 2000） |
| `overlap_size` | number | 语义模式下相邻块重叠字符数（默认 200） |
| `anchors` | array | anchor 模式必填。锚点列表，支持 `Array[Object]`、JSON 字符串、或含 `anchors_json` 键的对象三种输入 |

### 输出格式

```json
{
  "total_chunks": 5,
  "file_name": "example.docx",
  "file_type": "docx",
  "split_mode": "anchor",
  "chunks": [
    {
      "index": 0,
      "content": "第一章 项目概况\n本项目...[IMAGE:image_1]\n| 表头 | ... |",
      "metadata": {
        "section_standard": "作业和活动概况",
        "matched_anchor": "项目概况",
        "match_type": "primary",
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

### anchor 模式输入格式

`anchors` 参数类型为 `array`，支持以下三种输入：

**方式 1：直接传 Array[Object]（推荐，上游代码节点可直接连接）**

```json
[
  {
    "section_standard": "作业和活动概况",
    "anchor_primary": "项目概况",
    "anchor_fallbacks": ["项目背景"],
    "anchor_pos": 2,
    "confidence": 0.85,
    "found": true
  }
]
```

**方式 2：JSON 字符串**

```
"[{\"section_standard\": \"作业和活动概况\", \"anchor_primary\": \"项目概况\", ...}]"
```

**方式 3：包含 `anchors_json` 键的对象**

```json
{
  "anchors_json": [
    { "section_standard": "作业和活动概况", ... }
  ]
}
```

Dify Workflow 中，上游代码节点的 `Array[Object]` 类型输出可直接连接到本节点的 `anchors` 参数。

## 能力矩阵

| 内容 | PDF | Word | Excel |
|---|---|---|---|
| 纯文本 | ✓ | ✓ | ✓ |
| 表格 | 部分 | ✓ (Markdown) | ✓ (Markdown) |
| 图片 | ✓ (base64) | ✓ (base64) | — |
| 图表 | ✓ (已渲染为图像) | ✓ (fallback 图像) | — |
| 按页切分 | 页 | 分页符 | Sheet |
| 按语义切分 | 段落 | Heading 层级 | Sheet + 行数 |
| 按锚点切分 | — | ✓ | — |

## 许可证

[MIT](LICENSE)
