"""Theme, font, size handling and the pure helper functions."""

from __future__ import annotations

import pytest

from docx_helpers import assert_document_integrity, load, style_fonts
from scripts.markdown_to_word import (
    FONT_ALIASES,
    THEME_COLORS,
    blend_hex,
    make_bookmark_name,
    normalize_image_extension,
    prepare_markdown,
    resolve_color,
    resolve_font,
    resolve_theme_name,
    slugify_heading,
    strip_front_matter,
    validate_font_size,
    weighted_text_length,
)


@pytest.mark.parametrize("name", sorted(THEME_COLORS))
def test_every_preset_theme_resolves(name):
    assert resolve_theme_name(name) == name
    assert resolve_color(name) == THEME_COLORS[name]


@pytest.mark.parametrize(
    ("alias", "expected"),
    [
        ("classic blue", "classic-blue"),
        ("ClassicBlue", "classic-blue"),
        ("professional_green", "professional-green"),
        ("grey", "minimalist-gray"),
        ("grey background", "gray-background"),
        ("灰底黑字", "gray-background"),
        ("BLUE", "classic-blue"),
    ],
)
def test_theme_aliases(alias, expected):
    assert resolve_theme_name(alias) == expected


def test_unknown_theme_name_is_not_a_theme():
    assert resolve_theme_name("chartreuse") is None
    assert resolve_theme_name("#123456") is None


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("#abcdef", "ABCDEF"),
        ("abcdef", "ABCDEF"),
        ("#ABC", "AABBCC"),
        ("#000000", "000000"),
        ("  #FfFfFf  ", "FFFFFF"),
    ],
)
def test_hex_colors(value, expected):
    assert resolve_color(value) == expected


@pytest.mark.parametrize("value", ["#12345", "#GGGGGG", "", "rgb(1,2,3)"])
def test_invalid_colors_rejected(value):
    with pytest.raises(ValueError):
        resolve_color(value)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("宋体", "SimSun"),
        ("SimSun", "SimSun"),
        ("simsun (宋体)", "SimSun"),
        ("微软雅黑", "Microsoft YaHei"),
        ("Microsoft YaHei", "Microsoft YaHei"),
        ("default", "Aptos"),
        ("Arial", "Arial"),
        ("  Calibri  ", "Calibri"),
    ],
)
def test_font_aliases(value, expected):
    assert resolve_font(value) == expected
    assert resolve_font(value) == FONT_ALIASES.get(value.strip().lower(), expected)


@pytest.mark.parametrize("size", [6, 8.5, 12, 36, 72])
def test_font_size_bounds_accepted(size):
    assert validate_font_size(size) == pytest.approx(size)


@pytest.mark.parametrize("size", [5.9, 0, -12, 72.1, 120])
def test_font_size_out_of_range(size):
    with pytest.raises(ValueError):
        validate_font_size(size)


def test_font_size_accepts_numeric_strings():
    assert validate_font_size("11") == pytest.approx(11)


def test_minimum_font_size_document_is_valid(convert):
    result = convert("# 标题\n\n正文 ⚠\n", font_size=6)

    assert result.headings == 1
    document = load(result.output_path)
    assert document.styles["Heading 1"].font.size.pt == 22
    assert_document_integrity(result.output_path)


def test_maximum_font_size_document_is_valid(convert):
    result = convert("# 标题\n\n正文\n", font_size=72)

    assert result.headings == 1
    document = load(result.output_path)
    assert document.styles["Heading 1"].font.size.pt == 82
    assert document.styles["Normal"].font.size.pt == 72


@pytest.mark.parametrize(
    ("base", "expected"),
    [
        (12, [22, 18, 16, 14, 13, 12]),
        (16, [26, 22, 20, 18, 17, 16]),
        (10, [22, 17, 15, 13, 12, 11]),
    ],
)
def test_heading_scale(convert, base, expected):
    source = "\n\n".join(f"{'#' * level} 标题{level}" for level in range(1, 7))
    result = convert(source + "\n", font_size=base)
    document = load(result.output_path)

    sizes = [document.styles[f"Heading {level}"].font.size.pt for level in range(1, 7)]
    assert sizes == [pytest.approx(value) for value in expected]


@pytest.mark.parametrize("theme", sorted(THEME_COLORS))
def test_heading_color_matches_theme(convert, theme):
    result = convert("# 标题\n\n正文\n", color=theme)
    document = load(result.output_path)
    heading = document.styles["Heading 1"]

    assert str(heading.font.color.rgb) == THEME_COLORS[theme]
    assert_document_integrity(result.output_path)


def test_gray_theme_forces_black_text(convert):
    result = convert("# 标题\n\n正文 `代码` **粗体**\n", color="gray-background")
    document = load(result.output_path)

    assert str(document.styles["Normal"].font.color.rgb) == "000000"
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            assert str(run.font.color.rgb) == "000000"


def test_colored_theme_keeps_code_highlight_colors(convert):
    result = convert("```python\nvalue = 1  # note\n```\n", color="classic-blue")
    document = load(result.output_path)
    colors = {
        str(run.font.color.rgb)
        for paragraph in document.paragraphs
        for run in paragraph.runs
    }

    assert len(colors) > 1


def test_theme_font_is_applied_to_every_style(convert):
    result = convert("# 标题\n\n- 项目\n\n正文\n", font="Microsoft YaHei")
    document = load(result.output_path)

    for style_name in ("Normal", "Heading 1", "List Bullet"):
        fonts = style_fonts(document, style_name)
        assert fonts["eastAsia"] == "Microsoft YaHei"
        assert fonts["ascii"] == "Microsoft YaHei"


def test_blend_hex_clamps_ratio():
    assert blend_hex("000000", "FFFFFF", 0.0) == "000000"
    assert blend_hex("000000", "FFFFFF", 1.0) == "FFFFFF"
    assert blend_hex("000000", "FFFFFF", 2.0) == "FFFFFF"
    assert blend_hex("000000", "FFFFFF", -1.0) == "000000"
    assert blend_hex("FF0000", "00FF00", 0.5) == "808000"


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("Hello, World!", "hello-world"),
        ("中文标题", "中文标题"),
        ("多  个   空格", "多-个-空格"),
        ("a---b", "a-b"),
        ("--前后--", "前后"),
        ("", "section"),
        ("!!!", "section"),
        ("C++ 与 C#", "c-与-c"),
        ("2024 年度报告", "2024-年度报告"),
    ],
)
def test_slugify_heading(text, expected):
    assert slugify_heading(text) == expected


def test_slugify_heading_custom_separator():
    assert slugify_heading("a b c", "_") == "a_b_c"


def test_bookmark_names_are_stable_and_safe():
    first = make_bookmark_name("中文-标题 1")
    second = make_bookmark_name("中文-标题 1")

    assert first == second
    assert first.startswith("md_")
    assert first.isascii()
    assert len(first) == len("md_") + 12
    assert make_bookmark_name("x") != make_bookmark_name("y")


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("PNG", "png"),
        (".jpeg", "jpg"),
        ("jpg", "jpg"),
        ("svg+xml", "svg"),
        ("webp", "webp"),
        ("unknown", "png"),
        ("", "png"),
    ],
)
def test_normalize_image_extension(value, expected):
    assert normalize_image_extension(value) == expected


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("abc", 3.0),
        ("中文", 4.0),
        ("a中", 3.0),
        ("", 0.0),
    ],
)
def test_weighted_text_length(text, expected):
    assert weighted_text_length(text) == expected


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("---\ntitle: t\nauthor: a\n---\n\n正文\n", "正文\n"),
        ("---\nkey: value\n---", ""),
        ("---\n\n正文\n\n---\n", "---\n\n正文\n\n---\n"),
        ("---\ntitle: t\n", "---\ntitle: t\n"),
        ("正文\n\n---\n\n尾\n", "正文\n\n---\n\n尾\n"),
        ("---\n---\n\n正文\n", "---\n---\n\n正文\n"),
        ("", ""),
    ],
)
def test_front_matter_boundaries(source, expected):
    assert strip_front_matter(source) == expected


def test_prepare_markdown_normalizes_newlines():
    assert prepare_markdown("a\r\nb\rc") == "a\nb\nc"


def test_prepare_markdown_keeps_dollar_in_fences():
    source = "```\na $b$ c\n```\n"
    assert prepare_markdown(source) == source
