# md2docx

[English](README.md) | **简体中文**

`md2docx` 是一个将 Markdown 文件或 Markdown 字符串转换为可编辑 Word 文档（`.docx`）的 UOS AI 技能。它内置转换脚本 `scripts/markdown_to_word.py`，支持常见 GitHub Flavored Markdown 结构，并提供主题色、字体和字号等基础排版参数。

## 功能特性

- 将 `.md`、`.markdown`、`.txt` 文件转换为 `.docx`
- 支持直接传入 Markdown 字符串生成 Word 文档
- 使用 `markdown-it-py` 解析（CommonMark + GFM 表格与删除线），嵌套列表与围栏符合 CommonMark 规则
- 支持 H1-H6 标题，并映射为 Word 原生标题样式，重复标题自动去重
- 支持粗体、斜体、删除线、行内代码，以及 HTML `<u>`、`<sub>`、`<sup>` 下划线与上下标
- 支持有序列表、无序列表、嵌套列表和任务清单（复选框为 `■`/`□`，不依赖特殊符号字体），列表项内的引用、代码块、表格不会丢失
- 支持 GFM 表格、表头主题色、交替行底色和列宽估算；单元格内的嵌套表格生成真实 Word 嵌套表
- 支持引用块、分隔线、脚注（引用与回链锚点可跳转）、内部链接和外部链接
- 支持围栏代码块和 Pygments 语法高亮（`gray-background` 主题下代码为单色）；整块词法分析，多行字符串与注释高亮一致
- 代码块使用单一段落 + 段内换行：200 行代码的 `document.xml` 从约 470 KB 降到约 125 KB
- 支持本地图片、`file://` 图片、远程图片、Base64 图片和 SVG 转换
- 支持 Mermaid 图表渲染为图片后嵌入 Word；多图合并为一次 CLI 调用（共用同一浏览器），失败时自动逐图重试
- 支持 LaTeX 行内公式和块级公式；无法渲染时保留公式源文
- 公式识别在词法阶段完成，不会误伤行内代码、缩进代码块、围栏代码块或货币金额
- 自动忽略 YAML front matter 与 HTML 注释

## 目录结构

```text
md2docx/
├── SKILL.md
├── README.md
├── README.zh-CN.md
├── scripts/
│   └── markdown_to_word.py
└── tests/
    ├── conftest.py
    ├── docx_helpers.py
    ├── fake_mermaid_cli.py
    ├── test_blocks.py
    ├── test_cli_extended.py
    ├── test_cli_media.py
    ├── test_code_blocks.py
    ├── test_document.py
    ├── test_images.py
    ├── test_internals.py
    ├── test_lists.py
    ├── test_math.py
    ├── test_robustness.py
    ├── test_tables.py
    └── test_themes_fonts.py
```

## 依赖安装

必需 Python 依赖：

```bash
python -m pip install 'markdown-it-py>=3,<5' 'mdit-py-plugins>=0.4,<1' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
```

如果系统启用了 PEP 668，不要直接修改系统 Python 环境。可以创建虚拟环境后再安装：

```bash
python -m venv .venv
. .venv/bin/activate
python -m pip install 'markdown-it-py>=3,<5' 'mdit-py-plugins>=0.4,<1' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
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

没有额外说明时使用默认样式转换:

```bash
python scripts/markdown_to_word.py \
  --markdown-file <源文件名>.md \
  --output <源文件名>.docx
```

## 样式参数

默认样式：

- 主题色：`gray-background`，页面浅灰底、文字黑色、图片与 Mermaid 保持彩色
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
- `gray-background`
- `creative-teal`

`--color` 也可以传入任意 `#RRGGBB` 颜色值。

## 封面与目录

- `--cover`：使用首个 Markdown 一级标题作为封面标题，并附当前日期时间。
- `--toc`：在封面下一页插入 Word 原生 TOC 字段；使用时自动包含封面。
- Word 打开文档时会请求更新目录；也可在目录上右键选择“更新域”。

## Mermaid 处理模式

- `auto`：默认模式。找到可用渲染器时嵌入 PNG，否则保留 Mermaid 源码并输出警告。
- `required`：任何 Mermaid 图无法渲染都会导致转换失败，适合正式技术文档导出。
- `skip`：不尝试渲染 Mermaid，直接保留源码。

脚本会依次检测全局 `mmdc`、`bunx --bun @mermaid-js/mermaid-cli` 和 `npx --yes @mermaid-js/mermaid-cli`。首次使用 `bunx` 或 `npx` 可能需要联网下载依赖。

## Python 调用

`markdown_to_word.py` 也可以作为 Python 模块导入：

```python
from pathlib import Path

from scripts.markdown_to_word import markdown_to_word

result = markdown_to_word(
    markdown_file=Path('document.md'),
    output_path=Path('document.docx'),
    color='gray-background',
    font='宋体',
    font_size=12,
    mermaid_mode='auto',
    overwrite=True,
    cover=True,
    toc=True,
)

print(result.output_path)
print(result.warnings)
```

返回值包含输出路径、标题数量、段落数量、列表数量、表格数量、图片数量、代码块数量、Mermaid 渲染与回退数量、公式渲染与回退数量，以及警告信息。

## 回归测试

测试全部使用离线夹具，不依赖 Mermaid CLI、网络或 Word。需要 `pytest>=8`：

```bash
python -m pip install pytest
python -m pytest tests -q
```

覆盖范围：

- `tests/test_math.py`：货币金额、行内代码、缩进/围栏代码块中的 `$`、转义 `$`、行内/块级/列表/表格内公式、同公式缓存复用、渲染失败回退、跨行不匹配
- `tests/test_lists.py`：2/3/4 空格嵌套与层级样式上限、编号重启与 `start`、列表项内的引用/代码/图片/公式、任务清单、引用内列表、尾随换行
- `tests/test_blocks.py`：HTML 注释、front matter、H1-H6 与字号、重复标题书签、引用块、分隔线、脚注单分隔线与锚点闭合、行内格式、软换行标题、死锚点、`javascript:` 链接、外链特殊字符
- `tests/test_code_blocks.py`：未知语言、无标签/`~` 围栏、info 参数、实体与制表符保留、长行/长块、Pygments 缺失回退、段落数与体积约束
- `tests/test_tables.py`：表头识别、重复表头、交替底色、列宽与最小值分配、行单元格数不一致、多行单元格、空表、单列、两层/三层嵌套与宽度约束、单元格空白规范化、大表格
- `tests/test_images.py`：相对/上级/百分号编码/绝对/`file://` 路径、宽高限制、自然尺寸、alt 转描述、重复图去重、WebP 转 PNG、非法 base64、非法 MIME、缺图占位、远程禁用不访问网络、SVG 回退
- `tests/test_document.py`：页面与样式、主题色、标题缩放、封面/目录、书签唯一、内链与外链、参数校验、overwrite、默认输出名、容器完整性、大代码块规模
- `tests/test_themes_fonts.py`：九种主题与别名、hex 颜色、字体别名、字号边界、标题尺寸表、灰度主题黑色文本、标题 slug、书签名、图片扩展名、front matter 边界、纯函数
- `tests/test_cli_media.py`：CLI 退出码、Mermaid 三模式与批量渲染（调用次数、批量失败降级、超时）、Mermaid 配置与环境变量、远程图片 16 MB 限制、任务清单字形
- `tests/test_cli_extended.py`：`--help`、互斥参数、默认输出名、样式参数透传、cover+toc、overwrite、`--no-download-images`、错误与警告输出
- `tests/test_robustness.py`：空/纯空白/仅 front matter 文档、BOM、CRLF、扩展名、幂等性与临时目录清理（含失败路径）、大型混合文档、计数器一致性、特殊字符、rels 完整性
- `tests/test_internals.py`：Mermaid 命令优先级与 Chrome 探测、批次超时/缺失产出、嵌套表与前后文本、`section`/未知块回退、非法 `start`、列表样式缺失降级、列宽最小值、超链接空格与粗斜体、Pygments token 映射、SVG 转换成功与失败、真实子进程 CLI

`tests/docx_helpers.py` 提供 `assert_document_integrity()`：校验 zip 容器、XML 可解析、`pStyle` 均在 `styles.xml`、`numId` 均在 `numbering.xml`、`r:id` 均有关系项、书签 start/end 成对、`docPr` id 唯一。所有含图片和表格的用例都跑这项检查。

当前规模：**347 个用例，`scripts/markdown_to_word.py` 行覆盖 95%**（未覆盖主要是防御性分支与 `__main__` 子进程入口）。

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
- 远程图片默认会被下载并嵌入 DOCX，单张上限 16 MB、超时 20 秒；涉及隐私或离线场景时使用 `--no-download-images`。
- 本地图片以 Markdown 文件所在目录为基准解析，也支持绝对路径与 `file://` URI。
- 文档开头的 YAML front matter（含 `key: value` 行）会被忽略，避免泄漏到正文。
- Mermaid 和 LaTeX 渲染失败时，默认优先保留源码，避免内容丢失。
- `--font-size` 表示正文基准字号，标题会根据层级自动放大。
- 代码块不再按行拆段，而是一个段落内用换行符分隔，因此长代码块可能在行间分页。
- 嵌套列表缩进遵循 CommonMark：有序列表建议 3-4 空格，无序列表 2 空格即可。

## 技能触发场景

当用户需要将 Markdown、MD 内容、技术文档、README、接口文档或项目说明导出为 Word/DOCX 时，可以使用该技能。典型请求包括：

- “把这个 md 文件转成 Word”
- “将 README.md 导出为 docx”
- “把这段 Markdown 内容生成可编辑 Word 文档”
- “技术方案里有表格、代码块和 Mermaid 图，帮我转成 docx”
