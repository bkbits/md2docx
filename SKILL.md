---
name: md2docx
description: "将 Markdown 文件或内容转换为可编辑 DOCX，支持主题色、中文字体、正文基准字号、分级标题、GFM 表格与列表、代码高亮、图片、Mermaid 和 LaTeX。用户要求 Markdown/MD 转 Word/DOCX、技术文档导出或自动化转换时使用。必需依赖：Markdown>=3.6,<4、beautifulsoup4>=4.12,<5、Pillow>=10,<12、python-docx>=1.2,<2、Pygments>=2.17,<3；可选公式依赖 matplotlib>=3.8,<4；Mermaid 可使用全局 mmdc 或通过 bunx/npx 调用 @mermaid-js/mermaid-cli。"
---

# MD2DOCX

使用 `scripts/markdown_to_word.py` 将 Markdown 文件或 Markdown 字符串转换为专业排版的 Word 文档。优先运行现有脚本，不要复制或重新实现转换逻辑。

## 默认样式

未指定样式参数时使用以下默认值：

- 主题色：`classic-blue`，对应 `#2980B9`
- 字体：宋体，DOCX 内部字体名为 `SimSun`
- 正文字号：`12pt`
- 标题字号：H1 `22pt`、H2 `18pt`、H3 `16pt`、H4 `14pt`、H5 `13pt`、H6 `12pt`
- 页面：A4，标题映射为 Word 原生 `Heading 1` 至 `Heading 6`

`--font-size` 仅表示正文基准字号。标题必须按层级相对放大，不能全部使用正文大小。

## 依赖

必需 Python 依赖：

- `Markdown>=3.6,<4`
- `beautifulsoup4>=4.12,<5`
- `Pillow>=10,<12`
- `python-docx>=1.2,<2`
- `Pygments>=2.17,<3`

在缺少依赖时，优先使用项目现有 Python 环境。需要安装时执行：

```bash
python -m pip install 'Markdown>=3.6,<4' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
```

如果系统启用了 PEP 668，不要破坏系统 Python；创建虚拟环境后安装，或使用项目已有环境。

可选能力：

- LaTeX 图片渲染：安装 `matplotlib>=3.8,<4`。未安装或公式不兼容时，以 `Cambria Math` 保留 LaTeX 源文，不丢失内容。
- Mermaid 图表：脚本依次检测全局 `mmdc`、`bunx --bun @mermaid-js/mermaid-cli`、`bunx --yes @mermaid-js/mermaid-cli`。
- `bunx`/`npx` 首次使用可能下载 Mermaid CLI，需要网络访问，并可能耗时较长。
- Mermaid CLI 需要 Chromium/Chrome。脚本会检测 `google-chrome`、`chromium` 或 `chromium-browser`。
- SVG 图片转 PNG 需要 `rsvg-convert`。

不要创建单独的 requirements 文件；依赖以本技能文档为准。

## 转换流程

1. 确认输入是 Markdown 文件还是 Markdown 内容。二者必须且只能提供一个。
2. 确认输出路径。默认文件输入会生成 `<源文件名>-转换版.docx`。
3. 未收到样式要求时使用技能默认样式。
4. 文档包含 Mermaid 且用户要求图表必须呈现时，使用 `--mermaid required`；一般转换使用默认 `auto`。
5. 执行转换并检查脚本输出的元素统计和警告。
6. 验证 DOCX 容器与 OOXML；能够渲染时再检查页面视觉效果。
7. 向用户报告输出路径、样式参数、Mermaid/公式回退情况和验证结果。

## 命令用法

在项目根目录执行，使用默认样式转换文件：

```bash
python .agents/skills/md2docx/scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx
```

要求所有 Mermaid 图必须渲染：

```bash
python .agents/skills/md2docx/scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx \
  --mermaid required \
  --mermaid-timeout 180
```

覆盖样式：

```bash
python .agents/skills/md2docx/scripts/markdown_to_word.py \
  --color professional-green \
  --font 'Microsoft YaHei' \
  --font-size 11 \
  --markdown-file document.md \
  --output document.docx
```

直接传入短 Markdown 内容：

```bash
python .agents/skills/md2docx/scripts/markdown_to_word.py \
  --markdown-content '# 标题' \
  --output output.docx
```

主题色支持网站预设名或任意 `#RRGGBB`：

- `traditional-black`
- `classic-blue`
- `professional-green`
- `modern-purple`
- `corporate-orange`
- `elegant-red`
- `minimalist-gray`
- `creative-teal`

## Mermaid 模式

- `auto`：默认。能找到渲染器时嵌入 PNG，否则保留 Mermaid 源码并发出警告。
- `required`：任何 Mermaid 图无法渲染都使转换失败。正式技术文档优先使用此模式。
- `skip`：不尝试渲染，直接保留 Mermaid 源码。

远程图片默认会下载并嵌入 DOCX。如涉及隐私、离线转换或不应访问外网，使用 `--no-download-images`。

## 支持范围

脚本支持以下 Markdown 到 Word 映射：

- H1-H6 到 Word 原生标题样式和内部书签
- 粗体、斜体、删除线、下划线、上下标和行内代码
- 有序列表、无序列表、嵌套列表和任务清单
- GFM 表格、表头主题色、交替行底色和列宽估算
- 引用块、分隔线、脚注和内部/外部链接
- 带语言标记的围栏代码块和 Pygments 语法高亮
- 本地图片、远程图片、Base64 图片和 SVG 转换
- Mermaid 图表转 PNG 并嵌入 DOCX
- `$...$` 与 `$$...$$` 公式；有 matplotlib 时尝试渲染，否则保留源文

## 文件安全

- 默认传递 `--overwrite`。
- 默认传递 `--output <源文件名>.docx`，生成文件名应与源文件名相同。
- 不修改 Markdown 源文件。
- 不删除与转换无关的用户文件或工作区变更。

对于包含 Mermaid 或本地图片的文档，还应检查 `word/media/` 中的媒体数量不是零，并确认脚本报告的 `Mermaid rendered` 数量符合源文件。若有 LibreOffice，可进一步导出 PDF 并抽查页面；未能视觉渲染时必须明确说明。
