# md2docx

**English** | [简体中文](README.zh-CN.md)

`md2docx` is a UOS AI skill that converts Markdown files or Markdown strings into editable Word documents (`.docx`). It ships the conversion script `scripts/markdown_to_word.py`, supports common GitHub Flavored Markdown constructs, and exposes basic layout options such as theme color, font, and font size.

## Features

- Converts `.md`, `.markdown`, and `.txt` files to `.docx`
- Accepts a Markdown string directly
- Parses with `markdown-it-py` (CommonMark plus GFM tables and strikethrough); nested lists and code fences follow CommonMark rules
- Supports H1-H6 headings mapped to native Word heading styles, with automatic de-duplication of repeated headings
- Supports bold, italic, strikethrough, and inline code, plus HTML `<u>`, `<sub>`, and `<sup>` underline and sub/superscript
- Supports ordered, unordered, and nested lists plus task lists (checkboxes use `■`/`□`, so no special symbol font is required); blockquotes, code blocks, and tables inside list items are preserved
- Supports GFM tables with themed header rows, zebra striping, and column-width estimation; tables nested inside cells become real Word nested tables
- Supports blockquotes, horizontal rules, footnotes (with clickable references and backlinks), and internal/external links
- Supports fenced code blocks with Pygments syntax highlighting (monochrome under the `gray-background` theme); whole-block lexing keeps highlighting consistent across multi-line strings and comments
- Code blocks use a single paragraph with in-paragraph line breaks: `document.xml` for 200 lines of code drops from about 470 KB to about 125 KB
- Supports local images, `file://` images, remote images, Base64 images, and SVG conversion
- Renders Mermaid diagrams to PNG and embeds them; multiple diagrams share one CLI invocation (and one browser), falling back to per-diagram retries on failure
- Supports inline and block LaTeX formulas; preserves the source when rendering is unavailable
- Formula detection happens during lexing, so inline code, indented code blocks, fenced code blocks, and currency amounts are not misparsed
- Ignores YAML front matter and HTML comments

## Repository layout

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

## Dependencies

Required Python dependencies:

```bash
python -m pip install 'markdown-it-py>=3,<5' 'mdit-py-plugins>=0.4,<1' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
```

If PEP 668 is enabled, do not modify the system Python environment. Create a virtual environment instead:

```bash
python -m venv .venv
. .venv/bin/activate
python -m pip install 'markdown-it-py>=3,<5' 'mdit-py-plugins>=0.4,<1' 'beautifulsoup4>=4.12,<5' 'Pillow>=10,<12' 'python-docx>=1.2,<2' 'Pygments>=2.17,<3'
```

Optional dependencies:

- `matplotlib>=3.8,<4`: renders LaTeX formulas to images.
- `mmdc`, `bunx`, or `npx`: invokes `@mermaid-js/mermaid-cli` to render Mermaid diagrams.
- `google-chrome`, `chromium`, or `chromium-browser`: browser runtime for Mermaid CLI.
- `rsvg-convert`: converts SVG images to PNG before embedding.

## Command-line usage

Run the script from the skill directory or the project root.

Convert a file with the default style:

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx
```

When `--output` is omitted, the script writes `<source-name>-转换版.docx` next to the source file.

Overwrite an existing output file:

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx   --overwrite
```

Pass Markdown content directly:

```bash
python scripts/markdown_to_word.py \
  --markdown-content '# Title' \
  --output output.docx
```

Set the theme color, font, and base font size:

```bash
python scripts/markdown_to_word.py \
  --color professional-green \
  --font 'Microsoft YaHei' \
  --font-size 11 \
  --markdown-file document.md \
  --output document.docx
```

Require every Mermaid diagram to render:

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx   --mermaid required   --mermaid-timeout 180
```

Convert offline or avoid downloading remote images:

```bash
python scripts/markdown_to_word.py \
  --markdown-file document.md \
  --output document.docx   --no-download-images
```

Default conversion when no extra options are needed:

```bash
python scripts/markdown_to_word.py \
  --markdown-file <source-name>.md \
  --output <source-name>.docx
```

## Styling options

Defaults:

- Theme: `gray-background` - light gray page background, black text, images and Mermaid diagrams keep their colors
- Font: `宋体` (`SimSun` inside the DOCX)
- Base body font size: `12pt`
- Page: A4
- Heading sizes: H1 `22pt`, H2 `18pt`, H3 `16pt`, H4 `14pt`, H5 `13pt`, H6 `12pt`

Available themes:

- `traditional-black`
- `classic-blue`
- `professional-green`
- `modern-purple`
- `corporate-orange`
- `elegant-red`
- `minimalist-gray`
- `gray-background`
- `creative-teal`

`--color` also accepts any `#RRGGBB` value.

## Cover page and table of contents

- `--cover`: uses the first level-1 Markdown heading as the cover title and adds the current date and time.
- `--toc`: inserts a native Word TOC field on the page after the cover; implies a cover.
- Word asks to update fields when the document opens; you can also right-click the TOC and choose "Update Field".

## Mermaid handling modes

- `auto`: default mode. Embeds a PNG when a renderer is available; otherwise keeps the Mermaid source and emits a warning.
- `required`: fails the conversion if any Mermaid diagram cannot be rendered; recommended for formal technical documents.
- `skip`: does not attempt to render Mermaid and keeps the source.

The script probes the global `mmdc`, then `bunx --bun @mermaid-js/mermaid-cli`, then `npx --yes @mermaid-js/mermaid-cli`. The first `bunx` or `npx` run may need network access to download packages.

## Python API

`markdown_to_word.py` can also be imported as a Python module:

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

The return value includes the output path, heading/paragraph/list/table/image/code-block counts, Mermaid and formula rendered/fallback counts, and warnings.

## Regression tests

All tests use offline fixtures; they need neither Mermaid CLI, network, nor Word. `pytest>=8` is required:

```bash
python -m pip install pytest
python -m pytest tests -q
```

Coverage:

- `tests/test_math.py`: currency amounts, inline code, `$` inside indented and fenced code, escaped `$`, inline/block/list/table formulas, cache reuse for identical formulas, render-failure fallback, unmatched delimiters across lines
- `tests/test_lists.py`: 2/3/4-space nesting and level-style caps, numbering restarts and `start`, blockquotes/code/images/formulas inside list items, task lists, lists inside blockquotes, trailing newlines
- `tests/test_blocks.py`: HTML comments, front matter, H1-H6 and font sizes, duplicate-heading bookmarks, blockquotes, horizontal rules, footnote separator and anchor closure, inline formatting, headings with soft line breaks, dead anchors, `javascript:` links, special characters in external links
- `tests/test_code_blocks.py`: unknown languages, untagged and `~` fences, info strings, entity and tab preservation, long lines and blocks, fallback without Pygments, paragraph-count and size limits
- `tests/test_tables.py`: header detection, repeated headers, zebra striping, column-width and minimum-width allocation, rows with mismatched cell counts, multi-line cells, empty tables, single columns, two- and three-level nesting with width constraints, cell whitespace normalization, large tables
- `tests/test_images.py`: relative/parent/percent-encoded/absolute/`file://` paths, width and height limits, natural size, alt text as description, duplicate-image deduplication, WebP to PNG, invalid base64, invalid MIME, missing-image placeholders, no network access when downloads are disabled, SVG fallback
- `tests/test_document.py`: page and styles, theme colors, heading scaling, cover/TOC, unique bookmarks, internal and external links, argument validation, overwrite, default output name, container integrity, large code-block size
- `tests/test_themes_fonts.py`: nine themes and aliases, hex colors, font aliases, font-size bounds, heading size table, black text under the gray theme, heading slugs, bookmark names, image extensions, front-matter boundaries, pure functions
- `tests/test_cli_media.py`: CLI exit codes, the three Mermaid modes and batch rendering (invocation counts, batch-failure fallback, timeouts), Mermaid config and environment variables, 16 MB remote-image limit, task-list glyphs
- `tests/test_cli_extended.py`: `--help`, mutually exclusive arguments, default output name, style option passthrough, cover+toc, overwrite, `--no-download-images`, error and warning output
- `tests/test_robustness.py`: empty/whitespace-only/front-matter-only documents, BOM, CRLF, extensions, idempotence and temp-directory cleanup (including failure paths), large mixed documents, counter consistency, special characters, rels integrity
- `tests/test_internals.py`: Mermaid command priority and Chrome detection, batch timeout/missing output, nested tables with surrounding text, `section`/unknown-block fallback, invalid `start`, fallback for missing list styles, minimum column widths, spaces and bold/italic in hyperlinks, Pygments token mapping, SVG conversion success and failure, real subprocess CLI

`tests/docx_helpers.py` provides `assert_document_integrity()`: it validates the zip container, XML parsing, that every `pStyle` exists in `styles.xml`, every `numId` exists in `numbering.xml`, every `r:id` has a relationship, bookmark start/end pairs, and unique `docPr` ids. Every test involving images or tables runs this check.

Current size: **347 tests and 95% line coverage of `scripts/markdown_to_word.py`** (the remainder is mostly defensive branches and the `__main__` subprocess entry point).

## Verifying the output

After conversion the script prints something like:

```text
Created: /path/to/document.docx
Elements: 3 headings, 12 paragraphs, 5 list items, 1 tables, 2 images, 1 code blocks
```

Check at least the following:

- The output file exists and ends with `.docx`
- The script reported no `Warning` or `Error`
- For documents with images or Mermaid diagrams, `word/media/` contains the expected media files
- When Mermaid rendering is required, the `Mermaid rendered` count matches the source

Use `unzip -t` for a quick container sanity check:

```bash
unzip -t document.docx
```

## Notes

- Input files must be UTF-8 encoded `.md`, `.markdown`, or `.txt` files.
- The script never modifies the Markdown source.
- If the output file already exists, `--overwrite` is required.
- Remote images are downloaded and embedded by default, with a 16 MB per-image limit and a 20-second timeout; use `--no-download-images` for privacy-sensitive or offline scenarios.
- Local images resolve relative to the Markdown file's directory; absolute paths and `file://` URIs are also supported.
- YAML front matter at the top of the document (including `key: value` lines) is ignored so it does not leak into the body.
- When Mermaid or LaTeX rendering fails, the source is preserved by default to avoid content loss.
- `--font-size` sets the base body size; headings scale automatically by level.
- Code blocks are no longer split into one paragraph per line; they use line breaks inside a single paragraph, so long blocks may paginate between lines.
- Nested list indentation follows CommonMark: 3-4 spaces for ordered lists and 2 spaces for unordered lists.

## When to use this skill

Use this skill when the user needs to export Markdown, MD content, technical documents, READMEs, API references, or project notes to Word/DOCX. Typical requests include:

- "Convert this md file to Word"
- "Export README.md as docx"
- "Turn this Markdown into an editable Word document"
- "My design doc has tables, code blocks, and Mermaid diagrams; convert it to docx"
