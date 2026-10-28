"""Robustness: empty input, encodings, re-runs, temp cleanup and integrity."""

from __future__ import annotations

import tempfile
import xml.etree.ElementTree as ElementTree
import zipfile
from pathlib import Path

import pytest

import scripts.markdown_to_word as converter_module
from docx_helpers import (
    assert_document_integrity,
    drawing_count,
    full_text,
    heading_paragraphs,
    load,
    media_names,
    texts,
)
from scripts.markdown_to_word import (
    ConversionResult,
    markdown_to_word,
)


def _convert(tmp_path, source: str, name: str = "out.docx", **kwargs):
    kwargs.setdefault("mermaid_mode", "skip")
    return markdown_to_word(
        markdown_content=source,
        output_path=tmp_path / name,
        overwrite=True,
        **kwargs,
    )


def test_empty_content_produces_valid_document(tmp_path):
    result = _convert(tmp_path, "")

    assert result.headings == 0
    assert_document_integrity(result.output_path)


def test_whitespace_only_content(tmp_path):
    result = _convert(tmp_path, "   \n\n\t\n")

    assert_document_integrity(result.output_path)
    assert result.headings == 0


def test_front_matter_only_content(tmp_path):
    result = _convert(tmp_path, "---\ntitle: t\n---\n")

    assert "title" not in full_text(result.output_path)
    assert_document_integrity(result.output_path)


def test_utf8_bom_input_file(tmp_path):
    source = tmp_path / "bom.md"
    source.write_bytes("\ufeff# 标题\n\n正文\n".encode("utf-8"))

    result = markdown_to_word(
        markdown_file=source,
        output_path=tmp_path / "bom.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    text = full_text(result.output_path)
    assert "标题" in text
    assert "\ufeff" not in text


def test_crlf_input_is_normalized(tmp_path):
    result = _convert(tmp_path, "# 标题\r\n\r\n第一行\r\n第二行\r\n")

    text = full_text(result.output_path)
    assert "\r" not in text
    assert "第一行" in text


@pytest.mark.parametrize("suffix", [".md", ".markdown", ".txt", ".MD"])
def test_supported_input_extensions(tmp_path, suffix):
    source = tmp_path / f"doc{suffix}"
    source.write_text("# 标题\n", encoding="utf-8")

    result = markdown_to_word(
        markdown_file=source,
        output_path=tmp_path / "out.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.headings == 1


def test_conversion_is_idempotent(tmp_path):
    source = "# 标题\n\n正文\n\n- 项目\n\n| a |\n| --- |\n| 1 |\n"
    first = _convert(tmp_path, source, name="same.docx")
    first_text = full_text(first.output_path)
    first_body = _body_signature(first.output_path)

    second = _convert(tmp_path, source, name="same.docx")

    assert full_text(second.output_path) == first_text
    assert _body_signature(second.output_path) == first_body
    assert_document_integrity(second.output_path)


def _body_signature(path) -> str:
    """Paragraph text plus table shape, ignoring volatile ids."""
    document = load(path)
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        parts.append(f"table:{len(table.rows)}x{len(table.columns)}")
    return "|".join(parts)


def test_temporary_work_directory_is_cleaned_up(tmp_path, monkeypatch):
    created: list[Path] = []
    original = tempfile.TemporaryDirectory

    def factory(*args, **kwargs):
        directory = original(*args, **kwargs)
        created.append(Path(directory.name))
        return directory

    monkeypatch.setattr(converter_module.tempfile, "TemporaryDirectory", factory)
    _convert(tmp_path, "# 标题\n\n![缺](none.png)\n")

    assert created, "converter should create a working directory"
    assert all(not path.exists() for path in created)


def test_work_directory_is_removed_on_failure(tmp_path, monkeypatch):
    created: list[Path] = []
    original = tempfile.TemporaryDirectory

    def factory(*args, **kwargs):
        directory = original(*args, **kwargs)
        created.append(Path(directory.name))
        return directory

    monkeypatch.setattr(converter_module.tempfile, "TemporaryDirectory", factory)

    from scripts.markdown_to_word import MermaidRenderer

    monkeypatch.setattr(MermaidRenderer, "_find_command", staticmethod(lambda: None))
    with pytest.raises(RuntimeError):
        _convert(
            tmp_path,
            "```mermaid\ngraph TD\n  A --> B\n```\n",
            mermaid_mode="required",
        )

    assert all(not path.exists() for path in created)


def test_large_mixed_document(tmp_path):
    parts = ["# 总标题"]
    for index in range(20):
        parts.append(f"## 章节 {index}")
        parts.append(f"段落 {index} 内容")
        parts.append(f"| 列A | 列B |\n| --- | --- |\n| {index} | {index * 2} |")
        parts.append(f"- 项目 {index}\n- 项目 {index + 1}")
        parts.append(f"```python\nvalue = {index}\n```")
    result = _convert(tmp_path, "\n\n".join(parts) + "\n")

    assert result.headings == 21
    assert result.tables == 20
    assert result.lists == 40
    assert result.code_blocks == 20
    assert_document_integrity(result.output_path)


def test_counters_match_document_elements(tmp_path):
    source = (
        "# 标题\n\n"
        "## 二级\n\n"
        "正文\n\n"
        "- 项目\n\n"
        "| a |\n| --- |\n| 1 |\n\n"
        "```text\ncode\n```\n"
    )
    result = _convert(tmp_path, source)
    document = load(result.output_path)

    assert len(heading_paragraphs(result.output_path)) == result.headings == 2
    assert len(document.tables) == result.tables == 1
    assert len(document.paragraphs) >= result.lists


def test_heading_text_appears_once(tmp_path):
    result = _convert(tmp_path, "# 标题\n\n正文\n")

    text = full_text(result.output_path)
    assert text.count("标题") == 1
    assert text.count("正文") == 1


def test_result_dataclass_defaults(tmp_path):
    result = ConversionResult(output_path=Path("x.docx"))

    assert result.headings == 0
    assert result.warnings == []
    assert result.mermaid_rendered == 0 and result.mermaid_fallback == 0
    assert result.formulas_rendered == 0 and result.formulas_fallback == 0


def test_output_parent_directories_are_created(tmp_path):
    target = tmp_path / "deep" / "nested" / "out.docx"
    result = markdown_to_word(
        markdown_content="# 标题\n",
        output_path=target,
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.output_path.exists()


def test_only_h1_document_with_cover(tmp_path):
    result = _convert(tmp_path, "# 唯一标题\n", cover=True)

    assert result.headings == 0
    document = load(result.output_path)
    assert document.paragraphs[0].text == "唯一标题"
    assert_document_integrity(result.output_path)


def test_document_with_tables_only(tmp_path):
    source = "| a | b |\n| --- | --- |\n| 1 | 2 |\n"
    result = _convert(tmp_path, source)

    assert result.tables == 1
    assert result.headings == 0
    assert_document_integrity(result.output_path)


def test_special_characters_survive_round_trip(tmp_path):
    source = "特殊字符：< > & \" ' \\ / 与 emoji ✅\n"
    result = _convert(tmp_path, source)

    text = full_text(result.output_path)
    for fragment in ("<", ">", "&", '"', "'", "✅"):
        assert fragment in text


def test_zip_contains_required_parts(tmp_path):
    result = _convert(tmp_path, "# 标题\n\n![图](data:image/png;base64,"
                              "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGM8YWPDwMDAxAAGAA+2AUSPIZanAAAAAElFTkSuQmCC)\n")

    with zipfile.ZipFile(str(result.output_path)) as archive:
        names = set(archive.namelist())
    assert {"[Content_Types].xml", "word/document.xml", "word/styles.xml"} <= names
    assert media_names(result.output_path)


def test_no_dangling_relationships_in_complex_document(tmp_path):
    source = (
        "# 标题\n\n"
        "[外部](https://example.com/a?x=1&y=2) 与 [内部](#标题)\n\n"
        "![图](data:image/png;base64,"
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGM8YWPDwMDAxAAGAA+2AUSPIZanAAAAAElFTkSuQmCC)\n\n"
        "脚注[^1]\n\n[^1]: 内容\n"
    )
    result = _convert(tmp_path, source)

    assert_document_integrity(result.output_path)
    assert drawing_count(result.output_path) == 1
