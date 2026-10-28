"""Document-level layout, styles, cover/TOC, and public API validation."""

from __future__ import annotations

import re

import pytest

from docx_helpers import (
    bookmarks,
    document_xml,
    full_text,
    load,
    settings_xml,
    style_fonts,
    texts,
)
from scripts.markdown_to_word import (
    blend_hex,
    markdown_to_word,
    prepare_markdown,
    resolve_color,
    resolve_font,
    resolve_theme_name,
    slugify_heading,
    strip_front_matter,
    validate_font_size,
)


def test_page_is_a4_with_expected_margins(convert):
    result = convert("正文\n")
    section = load(result.output_path).sections[0]

    assert section.page_width.inches == pytest.approx(8.2677, abs=0.01)
    assert section.page_height.inches == pytest.approx(11.6929, abs=0.01)
    assert section.left_margin.inches == pytest.approx(0.72, abs=0.01)


def test_default_font_size_and_color(convert):
    result = convert("正文\n")
    document = load(result.output_path)
    normal = document.styles["Normal"]

    assert normal.font.size.pt == 12
    assert normal.font.name == "SimSun"
    assert style_fonts(document, "Normal")["eastAsia"] == "SimSun"
    assert str(normal.font.color.rgb) == "000000"


def test_custom_font_size_and_theme(convert):
    result = convert(
        "# 标题\n\n正文\n",
        color="classic-blue",
        font="Microsoft YaHei",
        font_size=11,
    )
    document = load(result.output_path)
    heading = document.styles["Heading 1"]

    assert document.styles["Normal"].font.size.pt == 11
    assert heading.font.color.rgb is not None
    assert str(heading.font.color.rgb) == "2980B9"
    assert style_fonts(document, "Heading 1")["eastAsia"] == "Microsoft YaHei"


def test_heading_sizes_scale_with_base_font(convert):
    result = convert("# h1\n\n## h2\n", font_size=16)
    document = load(result.output_path)

    assert document.styles["Heading 1"].font.size.pt == 26
    assert document.styles["Heading 2"].font.size.pt == 22


def test_gray_background_sets_page_color(convert):
    result = convert("正文\n")
    xml = document_xml(result.output_path)
    settings = settings_xml(result.output_path)

    assert 'w:color="E7E6E6"' in xml
    assert "displayBackgroundShape" in settings
    assert settings.index("displayBackgroundShape") < settings.index("listSeparator")


def test_theme_without_background(convert):
    result = convert("正文\n", color="professional-green")
    xml = document_xml(result.output_path)

    assert "<w:background" not in xml


def test_cover_uses_first_heading_and_removes_it_from_body(convert):
    source = "# 文档标题\n\n## 章节\n\n正文\n"
    result = convert(source, cover=True)
    document = load(result.output_path)

    assert document.paragraphs[0].text == "文档标题"
    assert re.search(
        r"\d{4}年\d{2}月\d{2}日", document.paragraphs[1].text
    ), "cover date missing"
    assert sum(1 for text in texts(result.output_path) if text == "文档标题") == 1
    assert document_xml(result.output_path).count('w:type="page"') == 1
    # h2 is promoted because the cover consumes the h1
    assert any(
        paragraph.text == "章节" and paragraph.style.name == "Heading 1"
        for paragraph in document.paragraphs
    )


def test_toc_adds_cover_field_and_updates(convert):
    result = convert("# 标题\n\n## 章节\n\n正文\n", toc=True)
    xml = document_xml(result.output_path)
    settings = settings_xml(result.output_path)

    assert "TOC" in xml
    assert 'w:fldSimple' in xml
    assert "<w:updateFields" in settings
    assert settings.index("updateFields") < settings.index("decimalSymbol")
    assert xml.count('w:type="page"') == 2


def test_toc_omits_cover_heading_count(convert):
    result = convert("# 标题\n\n## 章节\n", toc=True)

    assert result.headings == 1


def test_bookmark_names_and_ids_are_unique(convert):
    source = "# 标题一\n\n## 标题一\n\n### 标题二\n"
    result = convert(source)
    document_bookmarks = bookmarks(result.output_path)

    assert len(document_bookmarks) == len(set(document_bookmarks))
    ids = list(document_bookmarks.values())
    assert len(ids) == len(set(ids))


def test_internal_link_targets_existing_bookmark(convert):
    source = "跳到[章节](#章节)\n\n## 章节\n"
    result = convert(source)
    document_bookmarks = bookmarks(result.output_path)
    xml = document_xml(result.output_path)

    anchors = re.findall(r'w:anchor="([^"]+)"', xml)
    assert anchors, "internal link was not converted to an anchor"
    assert anchors[0] in document_bookmarks


def test_external_link_creates_relationship(convert):
    import zipfile

    result = convert("[示例](https://example.com/page)\n")
    with zipfile.ZipFile(str(result.output_path)) as archive:
        relationships = archive.read("word/_rels/document.xml.rels").decode("utf-8")
    assert "https://example.com/page" in relationships


def test_counters_reflect_content(convert):
    source = (
        "# 标题\n\n"
        "正文段落\n\n"
        "- 项目一\n- 项目二\n\n"
        "| a |\n| --- |\n| 1 |\n\n"
        "```python\nprint(1)\n```\n"
    )
    result = convert(source)

    assert result.headings == 1
    assert result.lists == 2
    assert result.tables == 1
    assert result.code_blocks == 1
    assert result.paragraphs >= 1


def test_existing_output_requires_overwrite(tmp_path):
    output = tmp_path / "exists.docx"
    output.write_bytes(b"placeholder")

    with pytest.raises(FileExistsError):
        markdown_to_word(
            markdown_content="# 标题\n",
            output_path=output,
            mermaid_mode="skip",
        )

    result = markdown_to_word(
        markdown_content="# 标题\n",
        output_path=output,
        overwrite=True,
        mermaid_mode="skip",
    )
    assert result.output_path == output.resolve()


def test_default_output_name_for_file_input(tmp_path):
    source = tmp_path / "notes.md"
    source.write_text("# 标题\n\n正文\n", encoding="utf-8")

    result = markdown_to_word(markdown_file=source, mermaid_mode="skip")

    assert result.output_path.name == "notes-转换版.docx"
    assert result.output_path.exists()
    assert source.read_text(encoding="utf-8") == "# 标题\n\n正文\n"


def test_output_suffix_is_normalized(tmp_path):
    result = markdown_to_word(
        markdown_content="# 标题\n",
        output_path=tmp_path / "report.txt",
        mermaid_mode="skip",
    )

    assert result.output_path.suffix == ".docx"


def test_requires_exactly_one_source():
    with pytest.raises(ValueError):
        markdown_to_word(markdown_content="x", markdown_file="y.md")
    with pytest.raises(ValueError):
        markdown_to_word()


def test_rejects_unsupported_extension(tmp_path):
    source = tmp_path / "notes.pdf"
    source.write_text("data", encoding="utf-8")

    with pytest.raises(ValueError):
        markdown_to_word(markdown_file=source, mermaid_mode="skip")


def test_rejects_missing_file(tmp_path):
    with pytest.raises(FileNotFoundError):
        markdown_to_word(
            markdown_file=tmp_path / "missing.md", mermaid_mode="skip"
        )


@pytest.mark.parametrize(
    ("kwargs", "error"),
    [
        ({"markdown_content": "x", "color": "not-a-color"}, ValueError),
        ({"markdown_content": "x", "font_size": 4}, ValueError),
        ({"markdown_content": "x", "font_size": 80}, ValueError),
        ({"markdown_content": "x", "mermaid_mode": "sometimes"}, ValueError),
        ({"markdown_content": "x", "font": "  "}, ValueError),
    ],
)
def test_rejects_invalid_options(tmp_path, kwargs, error):
    with pytest.raises(error):
        markdown_to_word(
            output_path=tmp_path / "out.docx",
            overwrite=True,
            **kwargs,
        )


def test_document_container_is_valid_zip(convert):
    import xml.etree.ElementTree as ElementTree
    import zipfile

    result = convert("# 标题\n\n| a |\n| --- |\n| 1 |\n")
    with zipfile.ZipFile(str(result.output_path)) as archive:
        assert archive.testzip() is None
        for name in ("word/document.xml", "word/settings.xml", "word/styles.xml"):
            ElementTree.fromstring(archive.read(name))


def test_large_code_block_keeps_content_and_size_bounded(convert):
    lines = [f"value_{index} = {index}" for index in range(120)]
    source = "```python\n" + "\n".join(lines) + "\n```\n"
    result = convert(source)

    assert result.code_blocks == 1
    text = full_text(result.output_path)
    assert "value_0 = 0" in text
    assert "value_119 = 119" in text

    document = load(result.output_path)
    assert len(document.paragraphs) == 2
    assert len(document.paragraphs[1].runs) <= len(lines) * 3
    assert len(document_xml(result.output_path)) < 150_000


def test_pure_helpers():
    assert slugify_heading("Hello, World!") == "hello-world"
    assert slugify_heading("中文标题") == "中文标题"
    assert resolve_theme_name("Classic Blue") == "classic-blue"
    assert resolve_theme_name("#123456") is None
    assert resolve_color("#abc") == "AABBCC"
    assert resolve_color("gray-background") == "000000"
    assert resolve_font("宋体") == "SimSun"
    assert resolve_font("Arial") == "Arial"
    assert validate_font_size(12) == 12
    assert blend_hex("000000", "FFFFFF", 0.5) == "808080"
    assert strip_front_matter("---\ntitle: t\n---\n\n正文\n") == "正文\n"
    assert prepare_markdown("a\r\nb") == "a\nb"
    assert prepare_markdown("---\n\n正文\n\n---\n") == "---\n\n正文\n\n---\n"
