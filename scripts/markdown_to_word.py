#!/usr/bin/env python3
"""Convert Markdown from a file or string into a styled Word document.

The converter mirrors the core export behavior exposed by
https://markdowntoword.app/zh:

- GitHub-style Markdown structure is mapped to native Word elements.
- A theme color, document font, and base font size can be configured.
- Mermaid code fences are rendered to images when Mermaid CLI is available.
- Local, remote, and data-URI images can be embedded.
- Headings, links, tables, lists, task lists, blockquotes, fenced code,
  footnotes, strikethrough, and LaTeX source are preserved.

The public function ``markdown_to_word`` can be imported by another Python
program. The module also provides a command-line interface; run with --help.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
from datetime import datetime
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any, Iterable, Sequence, TypedDict, cast

try:
    from bs4 import BeautifulSoup, Comment, NavigableString, Tag
    from docx import Document
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.shared import Inches, Pt, RGBColor
    from docx.styles.style import ParagraphStyle
    from docx.text.run import Run
    from markdown_it import MarkdownIt
    from mdit_py_plugins.footnote import footnote_plugin
    from mdit_py_plugins.tasklists import tasklists_plugin
    from PIL import Image
except ImportError as exc:  # pragma: no cover - exercised by dependency errors
    raise SystemExit(
        "Missing dependency. See the bundled md2docx SKILL.md Dependencies "
        "section and install the required Python packages."
    ) from exc

try:
    from pygments import lex
    from pygments.lexers import TextLexer, get_lexer_by_name
    from pygments.token import Token
except ImportError:  # pragma: no cover - syntax highlighting is optional
    lex = None
    TextLexer = None
    get_lexer_by_name = None
    Token = None


THEME_COLORS = {
    "traditional-black": "000000",
    "classic-blue": "2980B9",
    "professional-green": "27AE60",
    "modern-purple": "8E44AD",
    "corporate-orange": "E67E22",
    "elegant-red": "E74C3C",
    "minimalist-gray": "34495E",
    "gray-background": "000000",
    "creative-teal": "1ABC9C",
}

THEME_ALIASES = {
    "black": "traditional-black",
    "traditional black": "traditional-black",
    "traditionalblack": "traditional-black",
    "blue": "classic-blue",
    "classic blue": "classic-blue",
    "classicblue": "classic-blue",
    "green": "professional-green",
    "professional green": "professional-green",
    "professionalgreen": "professional-green",
    "purple": "modern-purple",
    "modern purple": "modern-purple",
    "modernpurple": "modern-purple",
    "orange": "corporate-orange",
    "corporate orange": "corporate-orange",
    "corporateorange": "corporate-orange",
    "red": "elegant-red",
    "elegant red": "elegant-red",
    "elegantred": "elegant-red",
    "gray": "minimalist-gray",
    "grey": "minimalist-gray",
    "minimalist gray": "minimalist-gray",
    "minimalistgray": "minimalist-gray",
    "gray background": "gray-background",
    "graybackground": "gray-background",
    "grey background": "gray-background",
    "greybackground": "gray-background",
    "灰底黑字": "gray-background",
    "teal": "creative-teal",
    "creative teal": "creative-teal",
    "creativeteal": "creative-teal",
}

FONT_ALIASES = {
    "default": "Aptos",
    "simsun (宋体)": "SimSun",
    "宋体": "SimSun",
    "simsun": "SimSun",
    "microsoft yahei (微软雅黑)": "Microsoft YaHei",
    "微软雅黑": "Microsoft YaHei",
    "microsoft yahei": "Microsoft YaHei",
}

SUPPORTED_FILE_EXTENSIONS = {".md", ".markdown", ".txt"}
A4_WIDTH_INCHES = 8.2677
A4_HEIGHT_INCHES = 11.6929
PAGE_MARGIN_INCHES = 0.72
MAX_IMAGE_WIDTH_INCHES = A4_WIDTH_INCHES - (2 * PAGE_MARGIN_INCHES)
MAX_IMAGE_HEIGHT_INCHES = A4_HEIGHT_INCHES - (2 * PAGE_MARGIN_INCHES) - 0.35
REMOTE_IMAGE_TIMEOUT_SECONDS = 20
MAX_REMOTE_IMAGE_BYTES = 16 * 1024 * 1024

# Successor tags for ``w:settings`` children; the OOXML sequence is fixed and
# Word tolerates violations while stricter consumers may not.
SETTINGS_AFTER_DISPLAY_BACKGROUND_SHAPE = (
    "w:printPostScriptOverText",
    "w:printFractionalCharacterWidth",
    "w:printFormsData",
    "w:embedTrueTypeFonts",
    "w:embedSystemFonts",
    "w:saveSubsetFonts",
    "w:saveFormsData",
    "w:mirrorMargins",
    "w:alignBordersAndEdges",
    "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter",
    "w:gutterAtTop",
    "w:hideSpellingErrors",
    "w:hideGrammaticalErrors",
    "w:activeWritingStyle",
    "w:proofState",
    "w:formsDesign",
    "w:attachedTemplate",
    "w:linkStyles",
    "w:stylePaneFormatFilter",
    "w:stylePaneSortMethod",
    "w:documentType",
    "w:mailMerge",
    "w:revisionView",
    "w:trackRevisions",
    "w:doNotTrackMoves",
    "w:doNotTrackFormatting",
    "w:documentProtection",
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:showEnvelope",
    "w:summaryLength",
    "w:clickAndTypeStyle",
    "w:defaultTableStyle",
    "w:evenAndOddHeaders",
    "w:bookFoldRevPrinting",
    "w:bookFoldPrinting",
    "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery",
    "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin",
    "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin",
    "w:doNotShadeFormData",
    "w:noPunctuationKerning",
    "w:characterSpacingControl",
    "w:printTwoOnOne",
    "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter",
    "w:noLineBreaksBefore",
    "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema",
    "w:saveInvalidXml",
    "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly",
    "w:useXSLTWhenSaving",
    "w:saveThroughXslt",
    "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace",
    "w:updateFields",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "m:mathPr",
    "w:attachedSchema",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures",
    "w:forceUpgrade",
    "w:captions",
    "w:readModeInkLockDown",
    "w:smartTagType",
    "sl:schemaLibrary",
    "w:shapeDefaults",
    "w:doNotEmbedSmartTags",
    "w:decimalSymbol",
    "w:listSeparator",
)

SETTINGS_AFTER_UPDATE_FIELDS = (
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "m:mathPr",
    "w:attachedSchema",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures",
    "w:forceUpgrade",
    "w:captions",
    "w:readModeInkLockDown",
    "w:smartTagType",
    "sl:schemaLibrary",
    "w:shapeDefaults",
    "w:doNotEmbedSmartTags",
    "w:decimalSymbol",
    "w:listSeparator",
)


INLINE_TAGS = frozenset(
    {
        "a",
        "abbr",
        "b",
        "br",
        "cite",
        "code",
        "del",
        "em",
        "i",
        "img",
        "input",
        "ins",
        "kbd",
        "mark",
        "q",
        "s",
        "samp",
        "small",
        "span",
        "strong",
        "sub",
        "sup",
        "u",
        "var",
    }
)


def math_delimiter(state, position: int) -> tuple[bool, bool]:
    """Return ``(can_open, can_close)`` for the ``$`` at ``position``."""
    source = state.src
    previous = source[position - 1] if position > 0 else "\n"
    following = source[position + 1] if position + 1 < len(source) else "\n"
    can_open = following not in " \t\n$"
    can_close = previous not in " \t\n$" and not following.isdigit()
    return can_open, can_close


def has_unescaped_dollar(text: str) -> bool:
    """True when ``text`` holds a ``$`` that is not backslash-escaped."""
    index = text.find("$")
    while index != -1:
        backslashes = 0
        probe = index - 1
        while probe >= 0 and text[probe] == "\\":
            backslashes += 1
            probe -= 1
        if backslashes % 2 == 0:
            return True
        index = text.find("$", index + 1)
    return False


def math_inline_rule(state, silent: bool) -> bool:
    """Parse ``$...$`` with code-span and currency aware boundaries."""
    if state.src[state.pos] != "$":
        return False
    can_open, _ = math_delimiter(state, state.pos)
    if not can_open:
        if not silent:
            state.pending += "$"
        state.pos += 1
        return True

    start = state.pos + 1
    cursor = start
    closing = -1
    while True:
        cursor = state.src.find("$", cursor)
        if cursor == -1:
            break
        backslashes = 0
        probe = cursor - 1
        while probe >= 0 and state.src[probe] == "\\":
            backslashes += 1
            probe -= 1
        if backslashes % 2 == 0:
            closing = cursor
            break
        cursor += 1

    if closing == -1 or closing == start:
        if not silent:
            state.pending += "$"
        state.pos = start
        return True

    _, can_close = math_delimiter(state, closing)
    content = state.src[start:closing]
    if not can_close or has_unescaped_dollar(content):
        if not silent:
            state.pending += "$"
        state.pos = start
        return True

    if not silent:
        token = state.push("math_inline", "math", 0)
        token.markup = "$"
        token.content = content
    state.pos = closing + 1
    return True


def math_block_rule(state, start: int, end: int, silent: bool) -> bool:
    """Parse a ``$$`` block, including blocks nested in lists and quotes."""
    position = state.bMarks[start] + state.tShift[start]
    maximum = state.eMarks[start]
    if position + 2 > maximum or state.src[position : position + 2] != "$$":
        return False
    if silent:
        return True

    position += 2
    first_line = state.src[position:maximum]
    closed = False
    last_line = ""
    next_line = start
    if first_line.strip().endswith("$$"):
        first_line = first_line.strip()[:-2]
        closed = True

    while not closed:
        next_line += 1
        if next_line >= end:
            break
        line_start = state.bMarks[next_line] + state.tShift[next_line]
        line_end = state.eMarks[next_line]
        if line_start < line_end and state.tShift[next_line] < state.blkIndent:
            break
        line_text = state.src[line_start:line_end].rstrip()
        if line_text.endswith("$$"):
            marker = state.src.rfind("$$", line_start, line_end)
            last_line = state.src[line_start:marker]
            closed = True

    if not closed:
        return False

    state.line = next_line + 1
    token = state.push("math_block", "math", 0)
    token.block = True
    token.map = [start, state.line]
    token.markup = "$$"
    body = state.getLines(start + 1, next_line, state.tShift[start], True)
    token.content = (
        (first_line.strip() + "\n" if first_line.strip() else "") + body + last_line
    )
    return True


def render_math_inline(self, tokens, idx, options, env) -> str:
    latex = html.escape(tokens[idx].content, quote=True)
    return f'<span class="math-inline" data-latex="{latex}"></span>'


def render_math_block(self, tokens, idx, options, env) -> str:
    latex = html.escape(tokens[idx].content.strip(), quote=True)
    return f'<div class="math-block" data-latex="{latex}"></div>'


def heading_text(token) -> str:
    parts: list[str] = []
    pending = list(token.children or [])
    while pending:
        child = pending.pop(0)
        if child.type in {"text", "code_inline"}:
            parts.append(child.content)
        elif child.type in {"softbreak", "hardbreak"}:
            parts.append(" ")
        elif child.children:
            pending[0:0] = list(child.children)
    return "".join(parts)


def heading_id_rule(state) -> bool:
    """Attach stable, unique ids so internal links can resolve."""
    seen: dict[str, int] = {}
    tokens = state.tokens
    for index, token in enumerate(tokens):
        if token.type != "heading_open":
            continue
        inline = tokens[index + 1] if index + 1 < len(tokens) else None
        slug = slugify_heading(
            heading_text(inline) if inline is not None else "", "-"
        )
        count = seen.get(slug, 0)
        seen[slug] = count + 1
        token.attrSet("id", slug if count == 0 else f"{slug}-{count}")
    return True


def build_markdown_parser() -> MarkdownIt:
    parser = MarkdownIt("commonmark")
    parser.options["html"] = True
    parser.enable(["table", "strikethrough"])

    default_validate_link = parser.validateLink

    def validate_link(url: str) -> bool:
        if url.lower().startswith("file:"):
            return True
        return default_validate_link(url)

    parser.validateLink = validate_link
    parser.inline.ruler.after("backticks", "math_inline", math_inline_rule)
    parser.block.ruler.after("fence", "math_block", math_block_rule)
    parser.add_render_rule("math_inline", render_math_inline)
    parser.add_render_rule("math_block", render_math_block)
    parser.core.ruler.push("heading_ids", heading_id_rule)
    parser.use(footnote_plugin)
    parser.use(tasklists_plugin)
    return parser


@dataclass(frozen=True)
class InlineStyle:
    bold: bool = False
    italic: bool = False
    strike: bool = False
    underline: bool = False
    code: bool = False
    superscript: bool = False
    subscript: bool = False
    color: str | None = None
    size: float | None = None
    font: str | None = None


@dataclass
class ConversionResult:
    output_path: Path
    headings: int = 0
    paragraphs: int = 0
    lists: int = 0
    tables: int = 0
    images: int = 0
    code_blocks: int = 0
    mermaid_rendered: int = 0
    mermaid_fallback: int = 0
    formulas_rendered: int = 0
    formulas_fallback: int = 0
    warnings: list[str] = field(default_factory=list)


class MermaidRenderer:
    """Render Mermaid source with mmdc, bunx, or npx."""

    def __init__(
        self,
        work_dir: Path,
        color: str,
        font: str,
        mode: str,
        timeout_seconds: int,
    ) -> None:
        self.work_dir = work_dir
        self.color = color
        self.font = font
        self.mode = mode
        self.timeout_seconds = timeout_seconds
        self.command = self._find_command() if mode != "skip" else None
        self.config_path = work_dir / "mermaid-config.json"
        self.puppeteer_path = work_dir / "puppeteer-config.json"

        if mode == "required" and self.command is None:
            raise RuntimeError(
                "Mermaid rendering was required, but mmdc, bunx, and npx were not found."
            )
        if mode != "skip":
            self._write_configs()

    @staticmethod
    def _find_command() -> list[str] | None:
        mmdc = shutil.which("mmdc")
        if mmdc:
            return [mmdc]

        bunx = shutil.which("bunx")
        if bunx:
            return [bunx, "--bun", "@mermaid-js/mermaid-cli"]

        npx = shutil.which("npx")
        if npx:
            return [npx, "--yes", "@mermaid-js/mermaid-cli"]

        return None

    @staticmethod
    def _find_chrome() -> str | None:
        for command in ("google-chrome", "chromium", "chromium-browser"):
            executable = shutil.which(command)
            if executable:
                return executable
        if os.name == "nt":
            for candidate in (
                Path(os.environ.get("PROGRAMFILES", ""))
                / "Google/Chrome/Application/chrome.exe",
                Path(os.environ.get("PROGRAMFILES(X86)", ""))
                / "Google/Chrome/Application/chrome.exe",
                Path(os.environ.get("PROGRAMFILES(X86)", ""))
                / "Microsoft/Edge/Application/msedge.exe",
            ):
                if candidate.is_file():
                    return str(candidate)
        return None

    def _write_configs(self) -> None:
        light_color = blend_hex(self.color, "FFFFFF", 0.82)
        lighter_color = blend_hex(self.color, "FFFFFF", 0.92)
        mermaid_config = {
            "theme": "base",
            "securityLevel": "loose",
            "flowchart": {"htmlLabels": True, "useMaxWidth": True},
            "sequence": {"useMaxWidth": True, "wrap": True},
            "themeVariables": {
                "fontFamily": (
                    f'"{self.font}", "Noto Sans CJK SC", '
                    '"Source Han Sans SC", sans-serif'
                ),
                "primaryColor": f"#{light_color}",
                "primaryTextColor": "#1F2937",
                "primaryBorderColor": f"#{self.color}",
                "lineColor": f"#{self.color}",
                "secondaryColor": f"#{lighter_color}",
                "tertiaryColor": "#FFFFFF",
                "clusterBkg": f"#{lighter_color}",
                "clusterBorder": f"#{self.color}",
                "mainBkg": f"#{light_color}",
                "nodeBorder": f"#{self.color}",
                "actorBorder": f"#{self.color}",
                "actorBkg": f"#{lighter_color}",
                "actorTextColor": "#1F2937",
                "signalColor": "#374151",
                "signalTextColor": "#1F2937",
                "labelBoxBkgColor": "#FFFFFF",
                "labelBoxBorderColor": f"#{self.color}",
                "labelTextColor": "#1F2937",
                "noteBkgColor": "#FFF8D6",
                "noteBorderColor": "#D6B656",
                "noteTextColor": "#1F2937",
            },
        }
        self.config_path.write_text(
            json.dumps(mermaid_config, ensure_ascii=False), encoding="utf-8"
        )

        chrome = self._find_chrome()
        puppeteer_config: dict[str, object] = {
            "headless": True,
            "args": [
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--font-render-hinting=medium",
            ],
        }
        if chrome:
            puppeteer_config["executablePath"] = chrome
        self.puppeteer_path.write_text(
            json.dumps(puppeteer_config, ensure_ascii=False), encoding="utf-8"
        )

    def render(self, source: str, index: int) -> tuple[Path | None, str | None]:
        if self.mode == "skip":
            return None, None
        if self.command is None:
            return None, "Mermaid CLI was not found; diagram source was preserved as code."

        source_path = self.work_dir / f"mermaid-{index:03d}.mmd"
        output_path = self.work_dir / f"mermaid-{index:03d}.png"
        source_path.write_text(source, encoding="utf-8")

        command = [
            *self.command,
            "-i",
            str(source_path),
            "-o",
            str(output_path),
            "-e",
            "png",
            "-b",
            "white",
            "-s",
            "2",
            "--size",
            "1400",
            "-c",
            str(self.config_path),
            "-p",
            str(self.puppeteer_path),
            "-q",
        ]
        environment = self._environment()

        try:
            process = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                timeout=self.timeout_seconds,
                env=environment,
            )
        except subprocess.TimeoutExpired:
            message = (
                f"Mermaid diagram {index} timed out after "
                f"{self.timeout_seconds} seconds."
            )
            if self.mode == "required":
                raise RuntimeError(message)
            return None, message

        if process.returncode == 0 and output_path.exists():
            return output_path, None

        details = (process.stderr or process.stdout or "unknown error").strip()
        details = re.sub(r"\s+", " ", details)[:500]
        message = f"Mermaid diagram {index} failed: {details}"
        if self.mode == "required":
            raise RuntimeError(message)
        return None, message

    def render_all(
        self, sources: Sequence[str]
    ) -> list[tuple[Path | None, str | None]]:
        """Render every diagram, batching them into one CLI call when possible."""
        if not sources:
            return []
        if self.mode == "skip":
            return [(None, None) for _ in sources]
        if self.command is None:
            message = (
                "Mermaid CLI was not found; diagram source was preserved as code."
            )
            return [(None, message) for _ in sources]

        batched = self._render_batch(sources)
        if batched is not None:
            return batched
        return [self.render(source, index + 1) for index, source in enumerate(sources)]

    def _render_batch(
        self, sources: Sequence[str]
    ) -> list[tuple[Path | None, str | None]] | None:
        """One CLI call for all diagrams; ``None`` means fall back per diagram."""
        if any("```" in source for source in sources):
            return None

        batch_input = self.work_dir / "mermaid-batch.md"
        batch_output = self.work_dir / "mermaid-batch.png"
        blocks = "\n\n".join(
            f"```mermaid\n{source.strip()}\n```" for source in sources
        )
        batch_input.write_text(blocks + "\n", encoding="utf-8")

        command = [
            *self.command,
            "-i",
            str(batch_input),
            "-o",
            str(batch_output),
            "-e",
            "png",
            "-b",
            "white",
            "-s",
            "2",
            "--size",
            "1400",
            "-c",
            str(self.config_path),
            "-p",
            str(self.puppeteer_path),
            "-q",
        ]
        try:
            process = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                timeout=self.timeout_seconds * max(1, len(sources)),
                env=self._environment(),
            )
        except subprocess.TimeoutExpired:
            return None

        if process.returncode != 0:
            return None

        results: list[tuple[Path | None, str | None]] = []
        for index in range(1, len(sources) + 1):
            candidate = batch_output.with_name(f"{batch_output.stem}-{index}.png")
            if not candidate.exists():
                return None
            results.append((candidate, None))
        return results

    def _environment(self) -> dict[str, str]:
        environment = os.environ.copy()
        chrome = self._find_chrome()
        if chrome:
            environment["PUPPETEER_EXECUTABLE_PATH"] = chrome
            environment["PUPPETEER_SKIP_DOWNLOAD"] = "true"
        return environment


class FormulaRenderer:
    """Render LaTeX with matplotlib mathtext when available."""

    def __init__(self, work_dir: Path, color: str) -> None:
        self.work_dir = work_dir
        self.color = color
        self.cache: dict[tuple[str, bool], Path | None] = {}
        self.math_to_image: Any = None
        try:
            import importlib

            mathtext = importlib.import_module("matplotlib.mathtext")
        except ImportError:
            pass
        else:
            self.math_to_image = mathtext.math_to_image

    def render(self, latex: str, block: bool) -> Path | None:
        key = (latex, block)
        if key in self.cache:
            return self.cache[key]
        if self.math_to_image is None:
            self.cache[key] = None
            return None

        digest = hashlib.sha1(f"{block}:{latex}".encode("utf-8")).hexdigest()[:16]
        output = self.work_dir / f"formula-{digest}.png"
        expression = latex.strip()
        if not (expression.startswith("$") and expression.endswith("$")):
            expression = f"${expression}$"
        try:
            self.math_to_image(
                expression,
                str(output),
                dpi=240,
                format="png",
                color=f"#{self.color}",
            )
        except Exception:
            self.cache[key] = None
            return None
        self.cache[key] = output if output.exists() else None
        return self.cache[key]


class MarkdownDocxConverter:
    def __init__(
        self,
        markdown_text: str,
        source_dir: Path,
        output_path: Path,
        color: str,
        font: str,
        font_size: float,
        mermaid_mode: str,
        mermaid_timeout: int,
        download_images: bool,
        gray_background: bool,
        cover: bool,
        toc: bool,
    ) -> None:
        self.markdown_text = markdown_text
        self.source_dir = source_dir
        self.output_path = output_path
        self.color = color
        self.gray_background = gray_background
        self.cover = cover
        self.toc = toc
        self.text_color = "000000" if gray_background else "1F2937"
        self.page_color = "E7E6E6" if gray_background else None
        self.table_header_color = "BFBFBF" if gray_background else color
        self.font = font
        self.font_size = font_size
        self.heading_sizes = {
            1: max(font_size + 10, 22),
            2: max(font_size + 6, 17),
            3: max(font_size + 4, 15),
            4: max(font_size + 2, 13),
            5: max(font_size + 1, 12),
            6: max(font_size, 11),
        }
        self.table_font_size = max(8.0, font_size - 1.0)
        self.download_images = download_images
        self.result = ConversionResult(output_path=output_path)
        self.document = Document()
        self.parser = build_markdown_parser()
        self.bookmark_id = 1
        self.bookmarks: dict[str, str] = {}
        self.mermaid_index = 0
        self.mermaid_results: list[tuple[Path | None, str | None]] = []
        self.image_index = 0
        self._temporary_directory = tempfile.TemporaryDirectory(
            prefix="markdown-to-word-"
        )
        self.work_dir = Path(self._temporary_directory.name)
        self.mermaid = MermaidRenderer(
            self.work_dir,
            color="2980B9" if gray_background else color,
            font=font,
            mode=mermaid_mode,
            timeout_seconds=mermaid_timeout,
        )
        self.formulas = FormulaRenderer(self.work_dir, color=self.text_color)
        self._configure_document()

    def close(self) -> None:
        self._temporary_directory.cleanup()

    def _configure_document(self) -> None:
        section = self.document.sections[0]
        section.page_width = Inches(A4_WIDTH_INCHES)
        section.page_height = Inches(A4_HEIGHT_INCHES)
        section.top_margin = Inches(PAGE_MARGIN_INCHES)
        section.bottom_margin = Inches(PAGE_MARGIN_INCHES)
        section.left_margin = Inches(PAGE_MARGIN_INCHES)
        section.right_margin = Inches(PAGE_MARGIN_INCHES)
        section.header_distance = Inches(0.3)
        section.footer_distance = Inches(0.3)
        if self.page_color is not None:
            set_document_background(self.document, self.page_color)

        zoom = self.document.settings.element.find(qn("w:zoom"))
        if zoom is not None:
            zoom.set(qn("w:percent"), "100")

        normal = cast(ParagraphStyle, self.document.styles["Normal"])
        set_style_font(normal, self.font, self.font_size, self.text_color)
        normal.paragraph_format.space_after = Pt(max(4, self.font_size * 0.45))
        normal.paragraph_format.line_spacing = 1.35

        for level, size in self.heading_sizes.items():
            style = cast(
                ParagraphStyle, self.document.styles[f"Heading {level}"]
            )
            set_style_font(style, self.font, size, self.color, bold=True)
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            style.paragraph_format.keep_with_next = True
            style.paragraph_format.keep_together = True
            style.paragraph_format.space_before = Pt(14 if level <= 2 else 10)
            style.paragraph_format.space_after = Pt(6)
            style.paragraph_format.line_spacing = 1.15

        for style_name in (
            "List Bullet",
            "List Bullet 2",
            "List Bullet 3",
            "List Number",
            "List Number 2",
            "List Number 3",
        ):
            try:
                style = cast(ParagraphStyle, self.document.styles[style_name])
            except KeyError:
                continue
            set_style_font(style, self.font, self.font_size, self.text_color)
            style.paragraph_format.space_after = Pt(3)
            style.paragraph_format.line_spacing = 1.25

        self.document.core_properties.author = "Markdown to Word Python Converter"
        self.document.core_properties.subject = "Converted from Markdown"
        self.document.core_properties.keywords = (
            "Markdown, DOCX, Mermaid, Word, GFM"
        )

    def convert(self) -> ConversionResult:
        prepared = prepare_markdown(self.markdown_text)
        soup = BeautifulSoup(self.parser.render(prepared), "html.parser")
        self._render_mermaid_diagrams(soup)

        first_h1 = soup.find("h1")
        title = first_h1.get_text(" ", strip=True) if first_h1 else "未命名文档"
        self.document.core_properties.title = title

        if self.cover or self.toc:
            self._render_cover(title)
        if self.toc:
            self._render_toc()

        nodes = list(soup.contents)
        if self.cover or self.toc:
            for index, node in enumerate(nodes):
                if node is first_h1:
                    nodes.pop(index)
                    break
            for heading in soup.find_all(re.compile(r"h[2-6]")):
                heading.name = f"h{int(heading.name[1]) - 1}"
        for node in nodes:
            self._render_block(node)

        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self.document.save(str(self.output_path))
        return self.result

    def _render_mermaid_diagrams(self, soup: BeautifulSoup) -> None:
        """Render every Mermaid block up front so one browser can serve all of them."""
        sources = [
            code.get_text()
            for code in soup.find_all("code", class_="language-mermaid")
        ]
        if sources:
            self.mermaid_results = self.mermaid.render_all(sources)

    def _render_cover(self, title: str) -> None:
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(210)
        paragraph.paragraph_format.space_after = Pt(24)
        run = paragraph.add_run(title)
        apply_run_font(run, self.font, max(self.font_size + 16, 28), self.text_color)
        run.bold = True

        date_paragraph = self.document.add_paragraph()
        date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_paragraph.paragraph_format.space_after = Pt(0)
        date_run = date_paragraph.add_run(
            datetime.now().strftime("%Y年%m月%d日 %H:%M")
        )
        apply_run_font(date_run, self.font, self.font_size + 1, self.text_color)
        self._add_page_break()

    def _render_toc(self) -> None:
        heading = self.document.add_paragraph(style="Title")
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        heading.paragraph_format.space_after = Pt(18)
        run = heading.add_run("目录")
        apply_run_font(run, self.font, max(self.font_size + 6, 18), self.text_color)
        run.bold = True

        paragraph = self.document.add_paragraph()
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), 'TOC \\o "1-6" \\h \\z \\u')
        placeholder = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "更新目录以显示标题"
        placeholder.append(text)
        field.append(placeholder)
        paragraph._p.append(field)
        self._add_page_break()

        update_fields = self.document.settings.element.find(qn("w:updateFields"))
        if update_fields is None:
            update_fields = OxmlElement("w:updateFields")
            self.document.settings.element.insert_element_before(
                update_fields, *SETTINGS_AFTER_UPDATE_FIELDS
            )
        update_fields.set(qn("w:val"), "true")

    def _add_page_break(self) -> None:
        paragraph = self.document.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)

    def _render_block(self, node, *, quote: bool = False) -> None:
        if isinstance(node, Comment):
            return
        if isinstance(node, NavigableString):
            text = str(node).strip()
            if text:
                paragraph = self.document.add_paragraph()
                if quote:
                    self._style_blockquote(paragraph)
                self._add_text(paragraph, text, InlineStyle())
                self.result.paragraphs += 1
            return
        if not isinstance(node, Tag):
            return

        name = node.name.lower()
        if re.fullmatch(r"h[1-6]", name):
            self._render_heading(node, int(name[1]))
        elif name == "p":
            self._render_paragraph(node, quote=quote)
        elif name in {"ul", "ol"}:
            self._render_list(node, depth=0)
        elif name == "blockquote":
            for child in node.children:
                if isinstance(child, Comment):
                    continue
                if isinstance(child, Tag) and child.name in {"ul", "ol"}:
                    self._render_list(child, depth=0, quote=True)
                else:
                    self._render_block(child, quote=True)
        elif name == "pre":
            self._render_preformatted(node)
        elif name == "table":
            self._render_table(node)
        elif name == "hr":
            self._render_horizontal_rule()
        elif name == "div" and "math-block" in tag_classes(node):
            self._render_block_formula(
                str(node.get("data-latex") or node.get_text())
            )
        elif name == "section" and "footnotes" in tag_classes(node):
            self._render_footnotes(node)
        elif name == "div" and "footnote" in tag_classes(node):
            self._render_footnotes(node)
        elif name in {"section", "article", "main", "div"}:
            for child in node.children:
                self._render_block(child, quote=quote)
        else:
            text = node.get_text(" ", strip=True)
            if text:
                paragraph = self.document.add_paragraph()
                self._render_inline(node.children, paragraph, InlineStyle())
                self.result.paragraphs += 1

    def _render_heading(self, node: Tag, level: int) -> None:
        paragraph = self.document.add_paragraph(style=f"Heading {level}")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if level == 1 and self.result.headings == 0:
            paragraph.paragraph_format.space_after = Pt(12)
        self._render_inline(
            node.children,
            paragraph,
            InlineStyle(
                bold=True,
                color=self.color,
                size=self.heading_sizes[level],
                font=self.font,
            ),
        )
        heading_id = node.get("id") or slugify_heading(node.get_text(" ", strip=True), "-")
        bookmark_name = make_bookmark_name(heading_id)
        self._add_bookmark(paragraph, bookmark_name)
        self.bookmarks[heading_id] = bookmark_name
        self.result.headings += 1

    def _render_paragraph(self, node: Tag, *, quote: bool = False) -> None:
        paragraph = self.document.add_paragraph()
        if quote:
            self._style_blockquote(paragraph)
        self._render_inline(node.children, paragraph, InlineStyle())
        paragraph.paragraph_format.widow_control = True
        self.result.paragraphs += 1

    def _render_list(
        self,
        node: Tag,
        depth: int,
        quote: bool = False,
        bookmark_item_ids: bool = False,
    ) -> None:
        ordered = node.name.lower() == "ol"
        direct_items = node.find_all("li", recursive=False)
        style_depth = min(depth + 1, 3)
        style_name = (
            f"List Number{'' if style_depth == 1 else f' {style_depth}'}"
            if ordered
            else f"List Bullet{'' if style_depth == 1 else f' {style_depth}'}"
        )
        raw_start = node.get("start")
        try:
            start = int(str(raw_start)) if ordered and raw_start is not None else 1
        except ValueError:
            start = 1
        try:
            num_id = self._new_numbering_instance(style_name, start) if ordered else None
        except KeyError:
            num_id = None
        continuation_indent = 0.25 * (depth + 1) + 0.25

        for item_index, item in enumerate(direct_items):
            paragraph = self._open_list_paragraph(
                style_name, num_id, depth, start, item_index, ordered
            )
            if quote:
                self._style_blockquote(paragraph)

            first_paragraph_seen = False
            for child in item.children:
                if isinstance(child, Comment):
                    continue
                if isinstance(child, NavigableString):
                    text = str(child)
                    if not text.strip():
                        continue
                    if not paragraph.runs:
                        text = text.lstrip("\n")
                    self._add_text(paragraph, text, InlineStyle())
                    continue
                if not isinstance(child, Tag):
                    continue

                child_name = child.name.lower()
                if child_name in {"ul", "ol"}:
                    self._render_list(child, depth=depth + 1, quote=quote)
                elif child_name in INLINE_TAGS:
                    self._render_inline([child], paragraph, InlineStyle())
                elif child_name == "p" and not first_paragraph_seen:
                    first_paragraph_seen = True
                    self._render_inline(child.children, paragraph, InlineStyle())
                else:
                    self._render_block_indented(
                        child, continuation_indent, quote=quote
                    )

            self._trim_trailing_space(paragraph)
            paragraph.paragraph_format.widow_control = True
            item_id = str(item.get("id") or "")
            if bookmark_item_ids and item_id:
                name = make_bookmark_name(item_id)
                self.bookmarks[item_id] = name
                self._add_bookmark(paragraph, name)
            self.result.lists += 1

    @staticmethod
    def _trim_trailing_space(paragraph) -> None:
        """Drop the newline the HTML formatter leaves at the end of a list item."""
        for run in reversed(paragraph.runs):
            text = run.text
            if not text:
                continue
            stripped = text.rstrip("\n")
            if stripped != text:
                run.text = stripped
            if stripped:
                return

    def _open_list_paragraph(
        self,
        style_name: str,
        num_id: int | None,
        depth: int,
        start: int,
        item_index: int,
        ordered: bool,
    ):
        try:
            paragraph = self.document.add_paragraph(style=style_name)
            if num_id is not None:
                num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
                num_pr.get_or_add_ilvl().val = 0
                num_pr.get_or_add_numId().val = num_id
        except KeyError:
            paragraph = self.document.add_paragraph()
            paragraph.paragraph_format.left_indent = Inches(0.25 * (depth + 1))
            paragraph.paragraph_format.first_line_indent = Inches(-0.18)
            self._add_text(
                paragraph,
                f"{start + item_index}. " if ordered else "• ",
                InlineStyle(bold=ordered),
            )
        return paragraph

    def _render_block_indented(
        self, node: Tag, indent_inches: float, quote: bool = False
    ) -> None:
        before = len(self.document.paragraphs)
        self._render_block(node, quote=quote)
        for paragraph in self.document.paragraphs[before:]:
            current = paragraph.paragraph_format.left_indent
            paragraph.paragraph_format.left_indent = (
                Inches(indent_inches)
                if current is None
                else current + Inches(indent_inches)
            )

    def _new_numbering_instance(self, style_name: str, start: int) -> int:
        style = cast(ParagraphStyle, self.document.styles[style_name])
        style_num_id = int(style.element.pPr.numPr.numId.val)
        numbering = self.document.part.numbering_part.element
        base_num = numbering.num_having_numId(style_num_id)
        abstract_num_id = int(base_num.abstractNumId.val)
        num = numbering.add_num(abstract_num_id)
        num.add_lvlOverride(ilvl=0).add_startOverride(start)
        return int(num.numId)

    def _render_preformatted(self, node: Tag) -> None:
        code = node.find("code")
        source = code.get_text() if code else node.get_text()
        classes = code.get("class", []) if code else []
        language = ""
        for class_name in classes:
            if class_name.startswith("language-"):
                language = class_name[len("language-") :]
                break

        if language.lower() == "mermaid":
            self.mermaid_index += 1
            if self.mermaid_index - 1 < len(self.mermaid_results):
                rendered, warning = self.mermaid_results[self.mermaid_index - 1]
            else:
                rendered, warning = self.mermaid.render(source, self.mermaid_index)
            if rendered is not None:
                try:
                    self._add_image_from_path(
                        rendered,
                        alt=f"Mermaid diagram {self.mermaid_index}",
                        force_max_width=True,
                    )
                except Exception as exc:
                    self.result.mermaid_fallback += 1
                    self.result.warnings.append(
                        f"Mermaid diagram {self.mermaid_index} could not be "
                        f"embedded: {exc}"
                    )
                    self._render_code_block(source, "mermaid")
                    return
                self.result.mermaid_rendered += 1
                return
            self.result.mermaid_fallback += 1
            if warning:
                self.result.warnings.append(warning)
            self._render_code_block(source, "mermaid")
            return

        self._render_code_block(source, language)

    def _render_code_block(self, source: str, language: str) -> None:
        if language:
            label = self.document.add_paragraph()
            label.paragraph_format.space_before = Pt(5)
            label.paragraph_format.space_after = Pt(2)
            run = label.add_run(language.upper())
            apply_run_font(run, self.font, max(8, self.font_size - 2), self.color)
            run.bold = True

        text = source.rstrip("\n")
        base_size = max(8, self.font_size - 1)
        fallback_style = InlineStyle(code=True, size=base_size)

        paragraph = self.document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0 if language else 5)
        paragraph.paragraph_format.space_after = Pt(5)
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.left_indent = Inches(0.12)
        paragraph.paragraph_format.right_indent = Inches(0.12)
        set_paragraph_shading(paragraph, "F5F7FA")
        set_paragraph_border(
            paragraph,
            left={"val": "single", "sz": "8", "color": self.color},
        )

        lexer = None
        if lex is not None and get_lexer_by_name is not None:
            try:
                lexer = get_lexer_by_name(
                    language or "text", stripnl=False, ensurenl=False
                )
            except Exception:
                lexer = (
                    TextLexer(stripnl=False, ensurenl=False)
                    if TextLexer is not None
                    else None
                )

        if lexer is None or not text:
            self._add_text(paragraph, text or " ", fallback_style)
            self.result.code_blocks += 1
            return

        # One paragraph with in-paragraph breaks keeps the file small, and lexing
        # the whole block keeps multi-line tokens (strings, comments) correct.
        pending_text = ""
        pending_style: InlineStyle | None = None
        assert lex is not None
        for token_type, token_text in lex(text, lexer):
            cleaned = token_text.replace("\r", "")
            if not cleaned:
                continue
            token_style = pygments_style(token_type)
            style = InlineStyle(
                code=True,
                size=base_size,
                bold=token_style.get("bold", False),
                italic=token_style.get("italic", False),
                color=(
                    self.text_color
                    if self.gray_background
                    else token_style.get("color")
                ),
            )
            if pending_style is not None and style != pending_style:
                self._add_text(paragraph, pending_text, pending_style)
                pending_text = ""
            pending_style = style
            pending_text += cleaned

        if pending_text:
            self._add_text(paragraph, pending_text, pending_style or fallback_style)
        if not paragraph.runs:
            self._add_text(paragraph, " ", fallback_style)

        self.result.code_blocks += 1

    def _render_table(
        self,
        node: Tag,
        *,
        container=None,
        max_width: float | None = None,
        font_size: float | None = None,
    ) -> None:
        html_rows = self._table_rows(node)
        if not html_rows:
            return
        column_count = max(
            len(row.find_all(["th", "td"], recursive=False)) for row in html_rows
        )
        if column_count == 0:
            return
        header_rows = self._table_header_rows(node, html_rows)
        available_width = MAX_IMAGE_WIDTH_INCHES if max_width is None else max_width
        cell_font_size = self.table_font_size if font_size is None else font_size
        parent = self.document if container is None else container

        table = parent.add_table(rows=len(html_rows), cols=column_count)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = calculate_column_widths(html_rows, column_count, available_width)

        for row_index, html_row in enumerate(html_rows):
            cells = html_row.find_all(["th", "td"], recursive=False)
            word_row = table.rows[row_index]
            for column_index in range(column_count):
                cell = word_row.cells[column_index]
                cell.width = Inches(widths[column_index])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                set_cell_margins(cell, top=70, start=90, bottom=70, end=90)
                paragraph = cell.paragraphs[0]
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.15
                if column_index < len(cells):
                    html_cell = cells[column_index]
                    self._render_table_cell(
                        html_cell, cell, paragraph, cell_font_size, widths[column_index]
                    )
                    alignment = get_html_alignment(html_cell)
                    if alignment is not None:
                        paragraph.alignment = alignment
                if row_index < header_rows:
                    set_cell_shading(cell, self.table_header_color)
                    for item in cell.paragraphs:
                        for run in item.runs:
                            run.bold = True
                            run.font.color.rgb = RGBColor.from_string(
                                self.text_color if self.gray_background else "FFFFFF"
                            )
                elif (row_index - header_rows) % 2 == 1:
                    set_cell_shading(cell, "F7F9FB")

            if row_index < header_rows:
                set_repeat_table_header(word_row)
            set_keep_table_row_together(word_row)

        self.result.tables += 1
        if container is None:
            spacer = self.document.add_paragraph()
            spacer.paragraph_format.space_after = Pt(0)

    def _render_table_cell(
        self,
        html_cell: Tag,
        cell,
        paragraph,
        font_size: float,
        cell_width: float,
    ) -> None:
        style = InlineStyle(size=font_size)
        self._collapse_cell_whitespace(html_cell)
        if html_cell.find("table") is None:
            self._render_inline(html_cell.children, paragraph, style)
            self._trim_paragraph_edges(paragraph)
            return

        current = paragraph
        for child in html_cell.children:
            if isinstance(child, Comment):
                continue
            if isinstance(child, NavigableString):
                text = str(child)
                if text.strip():
                    self._add_text(current, text, style)
                continue
            if not isinstance(child, Tag):
                continue

            child_name = child.name.lower()
            if child_name == "table":
                self._render_table(
                    child,
                    container=cell,
                    max_width=max(1.0, cell_width - 0.16),
                    font_size=max(7.5, font_size - 0.5),
                )
                trailing = cell.paragraphs[-1] if cell.paragraphs else None
                if trailing is not None and not trailing.text:
                    current = trailing
                else:
                    current = self._add_cell_paragraph(cell)
                continue
            if child_name in INLINE_TAGS:
                self._render_inline([child], current, style)
                continue
            if child_name == "p":
                if current.text:
                    current = self._add_cell_paragraph(cell)
                self._render_inline(child.children, current, style)
                continue
            self._render_inline([child], current, style)

        for item in cell.paragraphs:
            self._trim_paragraph_edges(item)

    @staticmethod
    def _collapse_cell_whitespace(html_cell: Tag) -> None:
        """Treat source indentation inside HTML table cells as a single space."""
        for text_node in list(html_cell.find_all(string=True)):
            parent = text_node.parent
            if parent is None or parent.name in {"code", "pre"}:
                continue
            collapsed = re.sub(r"\s+", " ", str(text_node))
            if collapsed != str(text_node):
                text_node.replace_with(collapsed)

    @staticmethod
    def _trim_paragraph_edges(paragraph) -> None:
        runs = [child for child in paragraph._p if child.tag == qn("w:r")]
        if not runs:
            return
        for run_element, leading in ((runs[0], True), (runs[-1], False)):
            if run_element.find(qn("w:drawing")) is not None:
                continue
            text_nodes = run_element.findall(qn("w:t"))
            if not text_nodes:
                continue
            node = text_nodes[0] if leading else text_nodes[-1]
            if not node.text:
                continue
            node.text = node.text.lstrip() if leading else node.text.rstrip()

    def _add_cell_paragraph(self, cell):
        paragraph = cell.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.15
        return paragraph

    @staticmethod
    def _table_rows(node: Tag) -> list[Tag]:
        """Collect only this table's rows, ignoring nested tables."""
        rows: list[Tag] = []
        for child in node.children:
            if not isinstance(child, Tag):
                continue
            if child.name == "tr":
                rows.append(child)
            elif child.name in {"thead", "tbody", "tfoot"}:
                rows.extend(child.find_all("tr", recursive=False))
        return rows

    @staticmethod
    def _table_header_rows(node: Tag, rows: Sequence[Tag]) -> int:
        head = node.find("thead")
        if head is not None:
            return max(1, len(head.find_all("tr", recursive=False)))
        if rows and rows[0].find("th") is not None:
            return 1
        return 0

    def _render_horizontal_rule(self) -> None:
        paragraph = self.document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(7)
        set_paragraph_border(
            paragraph,
            bottom={"val": "single", "sz": "6", "color": self.color},
        )

    def _render_block_formula(self, latex: str) -> None:
        latex = html.unescape(latex).strip()
        image = self.formulas.render(latex, block=True)
        if image is not None:
            try:
                self._add_image_from_path(
                    image,
                    alt="LaTeX formula",
                    force_max_width=False,
                    max_width=MAX_IMAGE_WIDTH_INCHES * 0.85,
                )
            except Exception:
                pass
            else:
                self.result.formulas_rendered += 1
                return

        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(6)
        set_paragraph_shading(paragraph, "F8FAFC")
        self._add_text(
            paragraph,
            latex,
            InlineStyle(font="Cambria Math", size=self.font_size + 1),
        )
        self.result.formulas_fallback += 1

    def _render_footnotes(self, node: Tag) -> None:
        lists = node.find_all(["ol", "ul"], recursive=False)
        if not lists:
            for child in node.children:
                if isinstance(child, Tag):
                    self._render_block(child)
            return
        for list_node in lists:
            self._render_list(list_node, depth=0, bookmark_item_ids=True)

    def _render_inline(
        self,
        nodes: Iterable,
        paragraph,
        style: InlineStyle,
    ) -> None:
        for node in nodes:
            if isinstance(node, Comment):
                continue
            if isinstance(node, NavigableString):
                self._add_text(paragraph, str(node), style)
                continue
            if not isinstance(node, Tag):
                continue

            name = node.name.lower()
            if name in {"strong", "b"}:
                self._render_inline(node.children, paragraph, replace(style, bold=True))
            elif name in {"em", "i"}:
                self._render_inline(node.children, paragraph, replace(style, italic=True))
            elif name in {"del", "s", "strike"}:
                self._render_inline(node.children, paragraph, replace(style, strike=True))
            elif name == "u":
                self._render_inline(node.children, paragraph, replace(style, underline=True))
            elif name == "code":
                self._render_inline(node.children, paragraph, replace(style, code=True))
            elif name == "sup":
                self._render_inline(
                    node.children, paragraph, replace(style, superscript=True)
                )
            elif name == "sub":
                self._render_inline(node.children, paragraph, replace(style, subscript=True))
            elif name == "br":
                paragraph.add_run().add_break()
            elif name == "a":
                self._add_link(paragraph, node, style)
            elif name == "img":
                self._add_inline_image(paragraph, node)
            elif name == "input" and str(node.get("type") or "").lower() == "checkbox":
                checked = node.has_attr("checked")
                self._add_text(paragraph, "■" if checked else "□", style)
            elif name == "span" and "math-inline" in tag_classes(node):
                self._add_inline_formula(
                    paragraph,
                    html.unescape(
                        str(node.get("data-latex") or node.get_text())
                    ),
                    style,
                )
            elif name in {"ul", "ol"}:
                continue
            else:
                self._render_inline(node.children, paragraph, style)

    def _add_text(self, paragraph, text: str, style: InlineStyle) -> Run | None:
        if text == "":
            return None
        run = paragraph.add_run(text)
        font_name = style.font or ("Consolas" if style.code else self.font)
        font_size = style.size if style.size is not None else self.font_size
        font_color = (
            self.text_color
            if self.gray_background
            else style.color or self.text_color
        )
        apply_run_font(run, font_name, font_size, font_color)
        run.bold = style.bold
        run.italic = style.italic
        run.font.strike = style.strike
        run.underline = style.underline
        run.font.superscript = style.superscript
        run.font.subscript = style.subscript
        if style.code:
            set_run_shading(run, "EEF1F4")
        return run

    def _add_link(self, paragraph, node: Tag, style: InlineStyle) -> None:
        href = str(node.get("href") or "")
        text = str(node.get_text("", strip=False) or href)
        if not href:
            self._add_text(paragraph, text, style)
            return

        if href.startswith("#"):
            target = urllib.parse.unquote(href[1:])
            bookmark = self.bookmarks.get(target, make_bookmark_name(target))
            element = add_hyperlink(
                paragraph,
                text,
                anchor=bookmark,
                color=self.color,
                font=style.font or self.font,
                size=style.size or self.font_size,
                bold=style.bold,
                italic=style.italic,
            )
        else:
            element = add_hyperlink(
                paragraph,
                text,
                url=href,
                color=self.color,
                font=style.font or self.font,
                size=style.size or self.font_size,
                bold=style.bold,
                italic=style.italic,
            )

        anchor_id = str(node.get("id") or "")
        if anchor_id:
            name = make_bookmark_name(anchor_id)
            self.bookmarks[anchor_id] = name
            self._wrap_bookmark(paragraph, element, name)

    def _add_inline_formula(
        self, paragraph, latex: str, parent_style: InlineStyle
    ) -> None:
        latex = latex.strip()
        image = self.formulas.render(latex, block=False)
        if image is not None:
            try:
                with Image.open(image) as formula_image:
                    width_px, height_px = formula_image.size
                height_inches = min(0.34, max(0.18, self.font_size / 36))
                width_inches = min(
                    2.4,
                    height_inches * (width_px / max(height_px, 1)),
                )
                paragraph.add_run().add_picture(
                    str(image),
                    width=Inches(width_inches),
                    height=Inches(height_inches),
                )
                self.result.formulas_rendered += 1
                return
            except Exception:
                pass

        self._add_text(
            paragraph,
            latex,
            replace(parent_style, font="Cambria Math", italic=True),
        )
        self.result.formulas_fallback += 1

    def _add_inline_image(self, paragraph, node: Tag) -> None:
        source = str(node.get("src") or "")
        alt = str(node.get("alt") or "")
        try:
            image_path = self._resolve_image(source)
            if image_path is None:
                raise FileNotFoundError(source)
            stream, width, height = prepare_image_for_docx(
                image_path,
                self.work_dir,
                max_width=MAX_IMAGE_WIDTH_INCHES,
                max_height=MAX_IMAGE_HEIGHT_INCHES,
                force_max_width=False,
            )
            run = paragraph.add_run()
            inline = run.add_picture(
                stream,
                width=Inches(width),
                height=Inches(height),
            )
            if alt:
                inline._inline.docPr.set("descr", alt)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.result.images += 1
        except Exception as exc:
            message = f"Image could not be embedded ({source}): {exc}"
            self.result.warnings.append(message)
            self._add_text(
                paragraph,
                f"[{alt or 'image'}: {source}]",
                InlineStyle(italic=True, color="6B7280"),
            )

    def _add_image_from_path(
        self,
        path: Path,
        alt: str,
        force_max_width: bool,
        max_width: float = MAX_IMAGE_WIDTH_INCHES,
    ) -> None:
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(6)
        stream, width, height = prepare_image_for_docx(
            path,
            self.work_dir,
            max_width=max_width,
            max_height=MAX_IMAGE_HEIGHT_INCHES,
            force_max_width=force_max_width,
        )
        inline = paragraph.add_run().add_picture(
            stream,
            width=Inches(width),
            height=Inches(height),
        )
        inline._inline.docPr.set("descr", alt)
        self.result.images += 1

    def _resolve_image(self, source: str) -> Path | None:
        if not source:
            return None
        source = html.unescape(source)

        if source.startswith("data:"):
            match = re.match(
                r"data:image/([a-zA-Z0-9.+-]+);base64,(.+)", source, re.DOTALL
            )
            if not match:
                return None
            extension = normalize_image_extension(match.group(1))
            self.image_index += 1
            path = self.work_dir / f"data-image-{self.image_index:03d}.{extension}"
            path.write_bytes(base64.b64decode(match.group(2)))
            return path

        if re.match(r"^[A-Za-z]:[\\/]", source):
            candidate = Path(source)
            return candidate if candidate.is_file() else None

        parsed = urllib.parse.urlparse(source)
        if parsed.scheme == "file":
            candidate = Path(urllib.request.url2pathname(parsed.path))
            return candidate if candidate.is_file() else None

        if parsed.scheme in {"http", "https"}:
            if not self.download_images:
                raise RuntimeError("remote image downloading is disabled")
            self.image_index += 1
            extension = normalize_image_extension(
                Path(parsed.path).suffix.lstrip(".") or "png"
            )
            path = self.work_dir / f"remote-image-{self.image_index:03d}.{extension}"
            request = urllib.request.Request(
                source,
                headers={"User-Agent": "MarkdownToWordPython/1.0"},
            )
            with urllib.request.urlopen(
                request, timeout=REMOTE_IMAGE_TIMEOUT_SECONDS
            ) as response:
                headers = getattr(response, "headers", None)
                declared = headers.get("Content-Length") if headers else None
                if declared is not None and declared.isdigit():
                    if int(declared) > MAX_REMOTE_IMAGE_BYTES:
                        raise RuntimeError(
                            "remote image exceeds the "
                            f"{MAX_REMOTE_IMAGE_BYTES // (1024 * 1024)} MB limit"
                        )
                payload = response.read(MAX_REMOTE_IMAGE_BYTES + 1)
            if len(payload) > MAX_REMOTE_IMAGE_BYTES:
                raise RuntimeError(
                    "remote image exceeds the "
                    f"{MAX_REMOTE_IMAGE_BYTES // (1024 * 1024)} MB limit"
                )
            path.write_bytes(payload)
            return path

        local_source = urllib.parse.unquote(parsed.path or source)
        local_path = Path(local_source)
        if not local_path.is_absolute():
            local_path = self.source_dir / local_path
        local_path = local_path.resolve()
        if not local_path.is_file():
            return None
        return local_path

    def _style_blockquote(self, paragraph) -> None:
        paragraph.paragraph_format.left_indent = Inches(0.25)
        paragraph.paragraph_format.right_indent = Inches(0.08)
        paragraph.paragraph_format.space_before = Pt(2)
        paragraph.paragraph_format.space_after = Pt(3)
        set_paragraph_shading(paragraph, blend_hex(self.color, "FFFFFF", 0.93))
        set_paragraph_border(
            paragraph,
            left={"val": "single", "sz": "18", "color": self.color},
        )

    def _add_bookmark(self, paragraph, name: str) -> None:
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(self.bookmark_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(self.bookmark_id))
        self.bookmark_id += 1

        paragraph_xml = paragraph._p
        insert_at = 1 if paragraph_xml.pPr is not None else 0
        paragraph_xml.insert(insert_at, start)
        paragraph_xml.append(end)

    def _wrap_bookmark(self, paragraph, element, name: str) -> None:
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(self.bookmark_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(self.bookmark_id))
        self.bookmark_id += 1
        element.addprevious(start)
        element.addnext(end)


def markdown_to_word(
    *,
    color: str = "gray-background",
    font: str = "宋体",
    font_size: float = 12,
    markdown_file: str | Path | None = None,
    markdown_content: str | None = None,
    output_path: str | Path | None = None,
    mermaid_mode: str = "auto",
    mermaid_timeout: int = 120,
    download_images: bool = True,
    overwrite: bool = False,
    cover: bool = False,
    toc: bool = False,
) -> ConversionResult:
    """Convert one Markdown file or Markdown string into a DOCX document.

    Exactly one of ``markdown_file`` and ``markdown_content`` must be supplied.
    ``color`` accepts a website theme name or a CSS ``#RRGGBB`` value.
    """

    if (markdown_file is None) == (markdown_content is None):
        raise ValueError(
            "Provide exactly one of markdown_file or markdown_content."
        )

    theme_name = resolve_theme_name(color)
    resolved_color = resolve_color(color)
    resolved_font = resolve_font(font)
    resolved_size = validate_font_size(font_size)
    if mermaid_mode not in {"auto", "required", "skip"}:
        raise ValueError("mermaid_mode must be auto, required, or skip")

    if markdown_file is not None:
        input_path = Path(markdown_file).expanduser().resolve()
        if not input_path.is_file():
            raise FileNotFoundError(f"Markdown file not found: {input_path}")
        if input_path.suffix.lower() not in SUPPORTED_FILE_EXTENSIONS:
            raise ValueError(
                f"Unsupported input extension {input_path.suffix!r}; "
                "use .md, .markdown, or .txt."
            )
        markdown_text = input_path.read_text(encoding="utf-8-sig")
        source_dir = input_path.parent
        default_output = input_path.with_name(f"{input_path.stem}-转换版.docx")
    else:
        markdown_text = markdown_content or ""
        source_dir = Path.cwd()
        default_output = source_dir / "markdown-output.docx"

    destination = (
        Path(output_path).expanduser().resolve()
        if output_path is not None
        else default_output.resolve()
    )
    if destination.suffix.lower() != ".docx":
        destination = destination.with_suffix(".docx")
    if destination.exists() and not overwrite:
        raise FileExistsError(
            f"Output already exists: {destination}. Pass overwrite=True or --overwrite."
        )

    converter = MarkdownDocxConverter(
        markdown_text=markdown_text,
        source_dir=source_dir,
        output_path=destination,
        color=resolved_color,
        font=resolved_font,
        font_size=resolved_size,
        mermaid_mode=mermaid_mode,
        mermaid_timeout=mermaid_timeout,
        download_images=download_images,
        gray_background=theme_name == "gray-background",
        cover=cover or toc,
        toc=toc,
    )
    try:
        return converter.convert()
    finally:
        converter.close()


FRONT_MATTER_PATTERN = re.compile(
    r"\A---[ \t]*\n(?P<body>.*?)\n---[ \t]*(?:\n|\Z)", re.DOTALL
)
FRONT_MATTER_KEY_PATTERN = re.compile(r"^\s*[\w.-]+\s*:", re.MULTILINE)


def prepare_markdown(markdown_text: str) -> str:
    text = markdown_text.replace("\r\n", "\n").replace("\r", "\n")
    return strip_front_matter(text)


def strip_front_matter(text: str) -> str:
    """Drop a leading YAML front matter block so it never reaches the body."""
    match = FRONT_MATTER_PATTERN.match(text)
    if match is None:
        return text
    body = match.group("body")
    if body.strip() and FRONT_MATTER_KEY_PATTERN.search(body) is None:
        return text
    return text[match.end() :].lstrip("\n")


def resolve_theme_name(value: str) -> str | None:
    normalized = value.strip().lower().replace("_", "-")
    normalized = re.sub(r"\s+", " ", normalized)
    alias = THEME_ALIASES.get(normalized, normalized.replace(" ", "-"))
    return alias if alias in THEME_COLORS else None


def resolve_color(value: str) -> str:
    theme_name = resolve_theme_name(value)
    if theme_name is not None:
        return THEME_COLORS[theme_name]

    match = re.fullmatch(r"#?([0-9a-fA-F]{6})", value.strip())
    if match:
        return match.group(1).upper()
    short_match = re.fullmatch(r"#?([0-9a-fA-F]{3})", value.strip())
    if short_match:
        return "".join(character * 2 for character in short_match.group(1)).upper()
    choices = ", ".join(THEME_COLORS)
    raise ValueError(f"Invalid color {value!r}. Use #RRGGBB or one of: {choices}")


def resolve_font(value: str) -> str:
    stripped = value.strip()
    if not stripped:
        raise ValueError("font must not be empty")
    return FONT_ALIASES.get(stripped.lower(), stripped)


def validate_font_size(value: float) -> float:
    size = float(value)
    if not 6 <= size <= 72:
        raise ValueError("font_size must be between 6 and 72 points")
    return size


def slugify_heading(value: str, separator: str = "-") -> str:
    value = html.unescape(value).strip().lower()
    value = re.sub(r"[^\w\-\u4e00-\u9fff]+", separator, value, flags=re.UNICODE)
    value = re.sub(re.escape(separator) + r"+", separator, value)
    return value.strip(separator) or "section"


def make_bookmark_name(value: str) -> str:
    digest = hashlib.sha1(value.encode("utf-8")).hexdigest()[:12]
    return f"md_{digest}"


def blend_hex(color_a: str, color_b: str, ratio_b: float) -> str:
    ratio_b = max(0.0, min(1.0, ratio_b))
    first = tuple(int(color_a[index : index + 2], 16) for index in (0, 2, 4))
    second = tuple(int(color_b[index : index + 2], 16) for index in (0, 2, 4))
    blended = tuple(
        round(component_a * (1 - ratio_b) + component_b * ratio_b)
        for component_a, component_b in zip(first, second)
    )
    return "".join(f"{component:02X}" for component in blended)


def set_document_background(document, color: str) -> None:
    document_element = document._element
    background = document_element.find(qn("w:background"))
    if background is None:
        background = OxmlElement("w:background")
        document_element.insert(0, background)
    background.set(qn("w:color"), color)

    settings = document.settings.element
    if settings.find(qn("w:displayBackgroundShape")) is None:
        settings.insert_element_before(
            OxmlElement("w:displayBackgroundShape"),
            *SETTINGS_AFTER_DISPLAY_BACKGROUND_SHAPE,
        )


def set_style_font(style, font: str, size: float, color: str, bold: bool = False) -> None:
    style.font.name = font
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor.from_string(color)
    style.font.bold = bold
    style_element = style._element
    rpr = style_element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attribute}"), font)


def apply_run_font(run: Run, font: str, size: float, color: str) -> None:
    run.font.name = font
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attribute}"), font)


def set_run_shading(run: Run, fill: str) -> None:
    rpr = run._element.get_or_add_rPr()
    shading = rpr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        rpr.append(shading)
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)


def set_paragraph_shading(paragraph, fill: str) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    shading = ppr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        ppr.insert_element_before(
            shading,
            "w:tabs",
            "w:suppressAutoHyphens",
            "w:kinsoku",
            "w:wordWrap",
            "w:overflowPunct",
            "w:topLinePunct",
            "w:autoSpaceDE",
            "w:autoSpaceDN",
            "w:bidi",
            "w:adjustRightInd",
            "w:snapToGrid",
            "w:spacing",
            "w:ind",
            "w:contextualSpacing",
            "w:mirrorIndents",
            "w:suppressOverlap",
            "w:jc",
            "w:textDirection",
            "w:textAlignment",
            "w:textboxTightWrap",
            "w:outlineLvl",
            "w:divId",
            "w:cnfStyle",
            "w:rPr",
            "w:sectPr",
            "w:pPrChange",
        )
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)


def set_paragraph_border(paragraph, **borders: dict[str, str]) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    p_borders = ppr.find(qn("w:pBdr"))
    if p_borders is None:
        p_borders = OxmlElement("w:pBdr")
        ppr.insert_element_before(
            p_borders,
            "w:shd",
            "w:tabs",
            "w:suppressAutoHyphens",
            "w:kinsoku",
            "w:wordWrap",
            "w:overflowPunct",
            "w:topLinePunct",
            "w:autoSpaceDE",
            "w:autoSpaceDN",
            "w:bidi",
            "w:adjustRightInd",
            "w:snapToGrid",
            "w:spacing",
            "w:ind",
            "w:contextualSpacing",
            "w:mirrorIndents",
            "w:suppressOverlap",
            "w:jc",
            "w:textDirection",
            "w:textAlignment",
            "w:textboxTightWrap",
            "w:outlineLvl",
            "w:divId",
            "w:cnfStyle",
            "w:rPr",
            "w:sectPr",
            "w:pPrChange",
        )
    for edge, attributes in borders.items():
        tag = qn(f"w:{edge}")
        element = p_borders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            p_borders.append(element)
        for key, value in attributes.items():
            element.set(qn(f"w:{key}"), str(value))
        element.set(qn("w:space"), "1")


def set_cell_shading(cell, fill: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shading = tcpr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tcpr.insert_element_before(
            shading,
            "w:noWrap",
            "w:tcMar",
            "w:textDirection",
            "w:tcFitText",
            "w:vAlign",
            "w:hideMark",
            "w:headers",
            "w:cellIns",
            "w:cellDel",
            "w:cellMerge",
            "w:tcPrChange",
        )
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int, start: int, bottom: int, end: int) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    margins = tcpr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tcpr.insert_element_before(
            margins,
            "w:textDirection",
            "w:tcFitText",
            "w:vAlign",
            "w:hideMark",
            "w:headers",
            "w:cellIns",
            "w:cellDel",
            "w:cellMerge",
            "w:tcPrChange",
        )
    for name, value in (
        ("top", top),
        ("start", start),
        ("bottom", bottom),
        ("end", end),
    ):
        element = margins.find(qn(f"w:{name}"))
        if element is None:
            element = OxmlElement(f"w:{name}")
            margins.append(element)
        element.set(qn("w:w"), str(value))
        element.set(qn("w:type"), "dxa")


def set_repeat_table_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    trpr.append(tbl_header)


def set_keep_table_row_together(row) -> None:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.keep_together = True


def calculate_column_widths(
    html_rows: Sequence[Tag], count: int, max_width: float = MAX_IMAGE_WIDTH_INCHES
) -> list[float]:
    weights = [4.0] * count
    for row in html_rows:
        cells = row.find_all(["th", "td"], recursive=False)
        for index, cell in enumerate(cells[:count]):
            lines = cell.get_text("\n", strip=True).splitlines() or [""]
            visual_length = max(weighted_text_length(line) for line in lines)
            weights[index] = max(weights[index], min(visual_length, 42.0))
    total = sum(weights) or 1.0
    widths = [max_width * weight / total for weight in weights]
    minimum = min(0.72, max_width / count * 0.75)
    for index, width in enumerate(widths):
        if width < minimum:
            deficit = minimum - width
            widths[index] = minimum
            widest = max(range(count), key=lambda candidate: widths[candidate])
            if widest != index:
                widths[widest] = max(minimum, widths[widest] - deficit)
    scale = max_width / sum(widths)
    return [width * scale for width in widths]


def weighted_text_length(text: str) -> float:
    return sum(2.0 if ord(character) > 127 else 1.0 for character in text)


def tag_classes(node: Tag) -> list[str]:
    value = node.get("class")
    if value is None:
        return []
    if isinstance(value, str):
        return value.split()
    return [str(item) for item in value]


def get_html_alignment(cell: Tag):
    style = str(cell.get("style") or "").lower()
    align = str(cell.get("align") or "").lower()
    combined = f"{style};text-align:{align}"
    if "text-align:center" in combined.replace(" ", ""):
        return WD_ALIGN_PARAGRAPH.CENTER
    if "text-align:right" in combined.replace(" ", ""):
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


def add_hyperlink(
    paragraph,
    text: str,
    *,
    url: str | None = None,
    anchor: str | None = None,
    color: str,
    font: str,
    size: float,
    bold: bool,
    italic: bool,
) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    if url is not None:
        relationship_id = paragraph.part.relate_to(
            url, RT.HYPERLINK, is_external=True
        )
        hyperlink.set(qn("r:id"), relationship_id)
    elif anchor is not None:
        hyperlink.set(qn("w:anchor"), anchor)
        hyperlink.set(qn("w:history"), "1")

    run_element = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attribute}"), font)
    rpr.append(rfonts)
    color_element = OxmlElement("w:color")
    color_element.set(qn("w:val"), color)
    rpr.append(color_element)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    size_element = OxmlElement("w:sz")
    size_element.set(qn("w:val"), str(round(size * 2)))
    rpr.append(size_element)
    size_cs = OxmlElement("w:szCs")
    size_cs.set(qn("w:val"), str(round(size * 2)))
    rpr.append(size_cs)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if italic:
        rpr.append(OxmlElement("w:i"))
    run_element.append(rpr)
    text_element = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" "):
        text_element.set(qn("xml:space"), "preserve")
    text_element.text = text
    run_element.append(text_element)
    hyperlink.append(run_element)
    paragraph._p.append(hyperlink)
    return hyperlink


def normalize_image_extension(value: str) -> str:
    extension = value.lower().lstrip(".")
    mapping = {
        "jpeg": "jpg",
        "jpg": "jpg",
        "png": "png",
        "gif": "gif",
        "webp": "webp",
        "svg+xml": "svg",
        "svg": "svg",
        "bmp": "bmp",
        "tiff": "tiff",
    }
    return mapping.get(extension, "png")


def prepare_image_for_docx(
    image_path: Path,
    work_dir: Path,
    *,
    max_width: float,
    max_height: float,
    force_max_width: bool,
) -> tuple[io.BytesIO, float, float]:
    path = image_path
    if path.suffix.lower() == ".svg":
        converter = shutil.which("rsvg-convert")
        if converter is None:
            raise RuntimeError("SVG requires rsvg-convert")
        converted = work_dir / f"converted-{hashlib.sha1(str(path).encode()).hexdigest()[:12]}.png"
        process = subprocess.run(
            [converter, "-o", str(converted), str(path)],
            check=False,
            capture_output=True,
            text=True,
            timeout=30,
        )
        if process.returncode != 0 or not converted.exists():
            raise RuntimeError((process.stderr or "SVG conversion failed").strip())
        path = converted

    raw = path.read_bytes()
    with Image.open(io.BytesIO(raw)) as image:
        image.load()
        width_px, height_px = image.size
        if width_px <= 0 or height_px <= 0:
            raise ValueError("invalid image dimensions")
        if image.format not in {"PNG", "JPEG", "GIF", "BMP", "TIFF"}:
            converted_stream = io.BytesIO()
            image.convert("RGB").save(converted_stream, format="PNG")
            raw = converted_stream.getvalue()

    aspect = width_px / height_px
    natural_width = max(0.8, width_px / 160.0)
    width = max_width if force_max_width else min(max_width, natural_width)
    height = width / aspect
    if height > max_height:
        height = max_height
        width = height * aspect
    stream = io.BytesIO(raw)
    stream.seek(0)
    return stream, width, height


class PygmentsRunStyle(TypedDict, total=False):
    color: str
    bold: bool
    italic: bool


def pygments_style(token_type) -> PygmentsRunStyle:
    if Token is None:
        return {}
    if token_type in Token.Keyword:
        return {"color": "0000CC", "bold": True}
    if token_type in Token.String:
        return {"color": "A31515"}
    if token_type in Token.Comment:
        return {"color": "008000", "italic": True}
    if token_type in Token.Number:
        return {"color": "098658"}
    if token_type in Token.Name.Function or token_type in Token.Name.Class:
        return {"color": "795E26", "bold": True}
    if token_type in Token.Name.Builtin:
        return {"color": "267F99"}
    if token_type in Token.Operator:
        return {"color": "666666"}
    if token_type in Token.Error:
        return {"color": "B91C1C"}
    return {"color": "1F2937"}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Markdown to a styled Word document (.docx).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python markdown_to_word.py --markdown-file document.md\n"
            "  python markdown_to_word.py --color classic-blue "
            "--font '宋体' --font-size 12 --markdown-file document.md\n"
            "  python markdown_to_word.py --color '#2980B9' --font Arial "
            "--font-size 11 --markdown-content '# Hello' --output hello.docx"
        ),
    )
    source_group = parser.add_mutually_exclusive_group(required=True)
    source_group.add_argument(
        "--markdown-file",
        type=Path,
        help="Path to a UTF-8 .md, .markdown, or .txt file.",
    )
    source_group.add_argument(
        "--markdown-content",
        help="Markdown content supplied directly on the command line.",
    )
    parser.add_argument(
        "--color",
        default="gray-background",
        help=(
            "Theme color name or #RRGGBB (default: gray-background). Names: "
            + ", ".join(THEME_COLORS)
            + "."
        ),
    )
    parser.add_argument(
        "--font",
        default="宋体",
        help=(
            "Document font, e.g. 宋体, Microsoft YaHei, Arial, or Calibri "
            "(default: 宋体)."
        ),
    )
    parser.add_argument(
        "--font-size",
        type=float,
        default=12,
        help="Base font size in points (6-72, default: 12).",
    )
    parser.add_argument("--output", type=Path, help="Output .docx path.")
    parser.add_argument(
        "--cover",
        action="store_true",
        help="Add a cover using first level-1 heading and current date/time.",
    )
    parser.add_argument(
        "--toc",
        action="store_true",
        help="Add native Word TOC on page after cover; also adds cover.",
    )
    parser.add_argument(
        "--mermaid",
        choices=("auto", "required", "skip"),
        default="auto",
        help=(
            "Mermaid handling: auto renders when mmdc/bunx/npx is available; "
            "required fails on render errors; skip preserves source as code."
        ),
    )
    parser.add_argument(
        "--mermaid-timeout",
        type=int,
        default=120,
        help="Timeout in seconds for each Mermaid diagram (default: 120).",
    )
    parser.add_argument(
        "--no-download-images",
        action="store_true",
        help="Do not download HTTP/HTTPS images; preserve them as text references.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacing an existing output file.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    arguments = parser.parse_args(argv)
    try:
        result = markdown_to_word(
            color=arguments.color,
            font=arguments.font,
            font_size=arguments.font_size,
            markdown_file=arguments.markdown_file,
            markdown_content=arguments.markdown_content,
            output_path=arguments.output,
            mermaid_mode=arguments.mermaid,
            mermaid_timeout=arguments.mermaid_timeout,
            download_images=not arguments.no_download_images,
            overwrite=arguments.overwrite,
            cover=arguments.cover,
            toc=arguments.toc,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Created: {result.output_path}")
    print(
        "Elements: "
        f"{result.headings} headings, "
        f"{result.paragraphs} paragraphs, "
        f"{result.lists} list items, "
        f"{result.tables} tables, "
        f"{result.images} images, "
        f"{result.code_blocks} code blocks"
    )
    print(
        "Advanced: "
        f"{result.mermaid_rendered} Mermaid rendered, "
        f"{result.mermaid_fallback} Mermaid fallbacks, "
        f"{result.formulas_rendered} formulas rendered, "
        f"{result.formulas_fallback} formula fallbacks"
    )
    if result.warnings:
        print("Warnings:", file=sys.stderr)
        for warning in result.warnings:
            print(f"  - {warning}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
