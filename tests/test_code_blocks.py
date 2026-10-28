"""Fenced and indented code blocks, lexing behavior and fallbacks."""

from __future__ import annotations

import pytest

from docx_helpers import (
    assert_document_integrity,
    full_text,
    load,
    single_run,
)
import scripts.markdown_to_word as converter_module


def test_unknown_language_falls_back_to_text_lexer(convert):
    result = convert("```nosuchlang\nvalue = 1\n```\n")

    assert result.code_blocks == 1
    assert "value = 1" in full_text(result.output_path)
    document = load(result.output_path)
    assert document.paragraphs[0].text == "NOSUCHLANG"


def test_unlabeled_fence_has_no_label_paragraph(convert):
    result = convert("```\nplain text\n```\n")
    document = load(result.output_path)

    assert len(document.paragraphs) == 1
    assert document.paragraphs[0].text == "plain text"


def test_tilde_fence_is_supported(convert):
    result = convert("~~~python\nvalue = 2\n~~~\n")

    assert result.code_blocks == 1
    assert "value = 2" in full_text(result.output_path)


def test_fence_with_extra_info_string(convert):
    result = convert('```python title="demo" linenos\nvalue = 3\n```\n')
    document = load(result.output_path)

    assert document.paragraphs[0].text == "PYTHON"
    assert "value = 3" in document.paragraphs[1].text


def test_language_label_is_uppercased(convert):
    result = convert("```JavaScript\nconst a = 1;\n```\n")
    document = load(result.output_path)

    assert document.paragraphs[0].text == "JAVASCRIPT"


def test_html_entities_inside_code_are_literal(convert):
    source = "```html\n<div class=\"a\">&amp; &lt;</div>\n```\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert '<div class="a">' in text
    assert "&amp;" in text
    assert "&lt;" in text


def test_tabs_are_preserved_as_tab_stops(convert):
    result = convert("```python\ndef f():\n\treturn 1\n```\n")
    document = load(result.output_path)

    assert "\treturn 1" in document.paragraphs[1].text


def test_trailing_whitespace_inside_code_is_preserved(convert):
    result = convert("```text\nvalue   \nnext\n```\n")
    document = load(result.output_path)

    assert document.paragraphs[1].text == "value   \nnext"


def test_long_single_line_code(convert):
    line = "x = " + "+".join(["1"] * 200)
    result = convert(f"```python\n{line}\n```\n")

    assert line in full_text(result.output_path)


def test_very_long_code_block_keeps_structure(convert):
    lines = [f"row_{index} = {index}" for index in range(600)]
    result = convert("```python\n" + "\n".join(lines) + "\n```\n")
    document = load(result.output_path)

    assert result.code_blocks == 1
    assert len(document.paragraphs) == 2
    assert document.paragraphs[1].text.count("\n") == len(lines) - 1


def test_code_block_without_pygments_falls_back(convert, monkeypatch):
    monkeypatch.setattr(converter_module, "lex", None)
    monkeypatch.setattr(converter_module, "get_lexer_by_name", None)
    result = convert("```python\nvalue = 1\n```\n")

    assert result.code_blocks == 1
    document = load(result.output_path)
    assert document.paragraphs[1].text == "value = 1"


def test_indented_code_block_is_rendered_as_code(convert):
    result = convert("正文\n\n    indented code\n")

    assert "indented code" in full_text(result.output_path)


def test_code_block_after_table_keeps_order(convert):
    source = "| a |\n| --- |\n| 1 |\n\n```text\nafter\n```\n"
    result = convert(source)

    from docx_helpers import document_xml

    xml = document_xml(result.output_path)
    assert xml.index("<w:tbl>") < xml.index("after")


def test_code_block_document_is_structurally_valid(convert):
    source = (
        "```python\nclass A:\n    \"\"\"doc\n\n    more\n    \"\"\"\n\n    def f(self):\n"
        "        return 1\n```\n"
    )
    result = convert(source)

    assert_document_integrity(result.output_path)
    assert result.code_blocks == 1


def test_code_block_inside_blockquote(convert):
    result = convert("> ```text\n> quoted code\n> ```\n")

    assert result.code_blocks == 1
    assert "quoted code" in full_text(result.output_path)


def test_mermaid_fallback_uses_code_block_label(convert):
    result = convert("```mermaid\ngraph TD\n  A --> B\n```\n", mermaid_mode="skip")
    document = load(result.output_path)

    assert result.warnings == []
    assert document.paragraphs[0].text == "MERMAID"
    assert "graph TD" in document.paragraphs[1].text


@pytest.mark.parametrize("language", ["python", "javascript", "json", "bash", "sql"])
def test_common_languages_are_highlighted(convert, language):
    result = convert(f"```{language}\nkey = value\n```\n")

    assert result.code_blocks == 1
    assert "key = value" in full_text(result.output_path)
