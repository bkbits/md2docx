# md2docx

`md2docx` 是一个将 Markdown 文件或 Markdown 字符串转换为可编辑 Word 文档（`.docx`）的 UOS AI 技能。它内置转换脚本 `scripts/markdown_to_word.py`，支持常见 GitHub Flavored Markdown 结构，并提供主题色、字体和字号等基础排版参数。

## 功能特性

- 将 `.md`、`.markdown`、`.txt` 文件转换为 `.docx`
- 支持直接传入 Markdown 字符串生成 Word 文档
- 支持 H1-H6 标题，并映射为 Word 原生标题样式
- 支持粗体、斜体、删除线、下划线、上下标、行内代码
- 支持有序列表、无序列表、嵌套列表和任务清单
- 支持 GFM 表格、表头主题色、交替行底色和列宽估算
- 支持引用块、分隔线、脚注、内部链接和外部链接
- 支持围栏代码块和 Pygments 语法高亮
- 支持本地图片、远程图片、Base64 图片和 SVG 转换
- 支持 Mermaid 图表渲染为图片后嵌入 Word
- 支持 LaTeX 行内公式和块级公式；无法渲染时保留公式源文

## 目录结构

```text
md2docx/
├── SKILL.md
├── README.md
└── scripts/
    └── markdown_to_word.py
```

## 依赖安装

必需 Python 依赖：

```bash
python -m pip install 'Markdown>=3.6,<4' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
```

如果系统启用了 PEP 668，不要直接修改系统 Python 环境。可以创建虚拟环境后再安装：

```bash
python -m venv .venv
. .venv/bin/activate
python -m pip install 'Markdown>=3.6,<4' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
```

可选依赖：

- `matplotlib>=3.8,<4`：用于将 LaTeX 公式渲染为图片。
- `mmdc`、`bunx` 或 `npx`：用于调用 `@mermaid-js/mermaid-cli` 渲染 Mermaid 图表。
- `google-chrome`、`chromium` 或 `chromium-browser`：Mermaid CLI 的浏览器运行环境。
- `rsvg-convert`：用于将 SVG 图片转换为 PNG 后嵌入 DOCX。

## 命令行用法

在技能目录或项目根目录中执行脚本。

使用默认样式转换文件：

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx
```

不指定 `--output` 时，脚本会在源文件旁生成 `<源文件名>-转换版.docx`。

覆盖已存在的输出文件：

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx   --overwrite
```

直接传入 Markdown 内容：

```bash
python scripts/markdown_to_word.py \
  --markdown-content '# 标题' \
  --output output.docx
```

指定主题色、字体和正文字号：

```bash
python scripts/markdown_to_word.py \
  --color professional-green \
  --font 'Microsoft YaHei' \
  --font-size 11 \
  --markdown-file document.md \
  --output document.docx
```

要求 Mermaid 图表必须成功渲染：

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx   --mermaid required   --mermaid-timeout 180
```

离线转换或不希望下载远程图片时：

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx   --no-download-images
```

没有额外说明时使用该方式转换:

```bash
python scripts/markdown_to_word.py \
  --color professional-green \
  --markdown-file <源文件名>.md \
  --output <源文件名>.docx
```

## 样式参数

默认样式：

- 主题色：`classic-blue`，对应 `#2980B9`
- 字体：`宋体`，DOCX 内部字体名为 `SimSun`
- 正文基准字号：`12pt`
- 页面：A4
- 标题字号：H1 `22pt`、H2 `18pt`、H3 `16pt`、H4 `14pt`、H5 `13pt`、H6 `12pt`

可用主题色：

- `traditional-black`
- `classic-blue`
- `professional-green`
- `modern-purple`
- `corporate-orange`
- `elegant-red`
- `minimalist-gray`
- `creative-teal`

`--color` 也可以传入任意 `#RRGGBB` 颜色值。

## Mermaid 处理模式

- `auto`：默认模式。找到可用渲染器时嵌入 PNG，否则保留 Mermaid 源码并输出警告。
- `required`：任何 Mermaid 图无法渲染都会导致转换失败，适合正式技术文档导出。
- `skip`：不尝试渲染 Mermaid，直接保留源码。

脚本会依次检测全局 `mmdc`、`bunx --bun @mermaid-js/mermaid-cli`、`bunx --yes @mermaid-js/mermaid-cli` 和 `npx --yes @mermaid-js/mermaid-cli`。首次使用 `bunx` 或 `npx` 可能需要联网下载依赖。

## Python 调用

`markdown_to_word.py` 也可以作为 Python 模块导入：

```python
from pathlib import Path

from scripts.markdown_to_word import markdown_to_word

result = markdown_to_word(
    markdown_file=Path('document.md'),
    output_path=Path('document.docx'),
    color='classic-blue',
    font='宋体',
    font_size=12,
    mermaid_mode='auto',
    overwrite=True,
)

print(result.output_path)
print(result.warnings)
```

返回值包含输出路径、标题数量、段落数量、表格数量、图片数量、代码块数量、Mermaid 渲染数量、公式渲染数量和警告信息。

## 转换验证

转换完成后，脚本会输出类似信息：

```text
Created: /path/to/document.docx
Elements: 3 headings, 12 paragraphs, 5 list items, 1 tables, 2 images, 1 code blocks
```

建议至少检查以下内容：

- 输出文件是否存在且扩展名为 `.docx`
- 脚本是否报告 `Warning` 或 `Error`
- 包含图片或 Mermaid 的文档中，`word/media/` 是否有对应媒体文件
- 需要严格渲染 Mermaid 时，`Mermaid rendered` 数量是否符合源文件

可以用 `unzip -t` 快速检查 DOCX 容器是否完整：

```bash
unzip -t document.docx
```

## 注意事项

- 输入文件必须是 UTF-8 编码的 `.md`、`.markdown` 或 `.txt` 文件。
- 脚本不会修改 Markdown 源文件。
- 输出文件已存在时，必须传入 `--overwrite` 才会覆盖。
- 远程图片默认会被下载并嵌入 DOCX；涉及隐私或离线场景时使用 `--no-download-images`。
- Mermaid 和 LaTeX 渲染失败时，默认优先保留源码，避免内容丢失。
- `--font-size` 表示正文基准字号，标题会根据层级自动放大。

## 技能触发场景

当用户需要将 Markdown、MD 内容、技术文档、README、接口文档或项目说明导出为 Word/DOCX 时，可以使用该技能。典型请求包括：

- “把这个 md 文件转成 Word”
- “将 README.md 导出为 docx”
- “把这段 Markdown 内容生成可编辑 Word 文档”
- “技术方案里有表格、代码块和 Mermaid 图，帮我转成 docx”
