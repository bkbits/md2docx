"""List structure: nesting, numbering, and block content inside items."""

from __future__ import annotations

from docx_helpers import (
    bordered_paragraphs,
    full_text,
    load,
    numbering_xml,
    paragraph_with,
    texts,
)
from scripts.markdown_to_word import markdown_to_word


def _styles_by_text(path, expected: list[str]) -> list[str]:
    document = load(path)
    styles = []
    for item in expected:
        paragraph = next(p for p in document.paragraphs if p.text.strip() == item)
        styles.append(paragraph.style.name)
    return styles


def test_ordered_list_nesting_with_four_spaces(convert):
    source = (
        "1. 第一层\n"
        "    1. 第二层 a\n"
        "    2. 第二层 b\n"
        "2. 第一层第二项\n"
    )
    result = convert(source)

    assert result.lists == 4
    assert _styles_by_text(
        result.output_path,
        ["第一层", "第二层 a", "第二层 b", "第一层第二项"],
    ) == ["List Number", "List Number 2", "List Number 2", "List Number"]


def test_ordered_list_nesting_with_three_spaces(convert):
    source = "1. 父\n   1. 子\n"
    result = convert(source)

    assert _styles_by_text(result.output_path, ["父", "子"]) == [
        "List Number",
        "List Number 2",
    ]


def test_bullet_list_nesting_with_two_spaces(convert):
    source = "- 父项\n  - 子项\n    - 孙项\n"
    result = convert(source)

    assert _styles_by_text(result.output_path, ["父项", "子项", "孙项"]) == [
        "List Bullet",
        "List Bullet 2",
        "List Bullet 3",
    ]


def test_nested_ordered_list_restarts_numbering(convert):
    source = "1. 甲\n    1. 乙\n"
    result = convert(source)

    document = load(result.output_path)
    outer = next(p for p in document.paragraphs if p.text.strip() == "甲")
    nested = next(p for p in document.paragraphs if p.text.strip() == "乙")
    assert nested._p.pPr.numPr.numId.val != outer._p.pPr.numPr.numId.val


def test_blockquote_inside_list_item_is_kept(convert):
    source = "1. 步骤一\n\n    > 注意引用\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert "注意引用" in text
    quoted = paragraph_with(result.output_path, "注意引用")
    assert quoted.paragraph_format.left_indent is not None
    assert any(
        text == "注意引用" and "left" in edges
        for text, edges in bordered_paragraphs(result.output_path)
    )


def test_fenced_code_inside_list_item_is_kept(convert):
    source = "1. 步骤\n\n    ```bash\n    echo hi\n    ```\n"
    result = convert(source)

    assert result.code_blocks == 1
    text = full_text(result.output_path)
    assert "echo hi" in text
    code_paragraph = paragraph_with(result.output_path, "echo hi")
    assert code_paragraph.paragraph_format.left_indent is not None


def test_second_paragraph_in_list_item_is_kept(convert):
    source = "- 第一段\n\n  第二段\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert "第一段" in text
    assert "第二段" in text


def test_table_inside_list_item_is_not_dropped(convert):
    source = (
        "1. 前\n\n"
        "    | a | b |\n"
        "    | --- | --- |\n"
        "    | 1 | 2 |\n"
    )
    result = convert(source)

    assert result.tables == 1
    assert "前" in full_text(result.output_path)


def test_list_item_text_has_no_trailing_newline(convert):
    result = convert("- 项目\n  - 子项\n")

    lines = [text for text in texts(result.output_path) if text.strip()]
    assert lines == ["项目", "子项"]


def test_task_list_checkboxes(convert):
    source = "- [x] 已完成\n- [ ] 未完成\n"
    result = convert(source)

    lines = [line for line in texts(result.output_path) if line.strip()]
    assert any("■" in line and "已完成" in line for line in lines)
    assert any("□" in line and "未完成" in line for line in lines)


def test_ordered_list_start_attribute(convert):
    source = "3. 三\n4. 四\n"
    result = convert(source)

    numbering = numbering_xml(result.output_path)
    assert 'w:startOverride w:val="3"' in numbering
    assert "三" in full_text(result.output_path)


def test_list_inside_blockquote(convert):
    source = "> 引用文本\n>\n> - 引用内列表\n"
    result = convert(source)

    assert "引用文本" in full_text(result.output_path)
    document = load(result.output_path)
    item = next(p for p in document.paragraphs if p.text.strip() == "引用内列表")
    assert item.style.name == "List Bullet"
    assert item.paragraph_format.left_indent is not None


def test_four_level_nesting_caps_at_third_style(convert):
    source = "- 一\n  - 二\n    - 三\n      - 四\n"
    result = convert(source)

    assert _styles_by_text(result.output_path, ["一", "二", "三", "四"]) == [
        "List Bullet",
        "List Bullet 2",
        "List Bullet 3",
        "List Bullet 3",
    ]


def test_loose_list_with_image_item(tmp_path, png_bytes):
    image = tmp_path / "icon.png"
    image.write_bytes(png_bytes)
    source_file = tmp_path / "loose.md"
    source_file.write_text("- 项目\n\n  说明 ![图标](icon.png)\n", encoding="utf-8")

    result = markdown_to_word(
        markdown_file=source_file,
        output_path=tmp_path / "loose.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1
    text = full_text(result.output_path)
    assert "项目" in text
    assert "说明" in text


def test_list_item_with_inline_formula(convert):
    result = convert("- 公式 $x^2$ 项\n")

    assert result.formulas_fallback + result.formulas_rendered == 1
    assert "项" in full_text(result.output_path)


def test_ordered_list_inside_ordered_list_keeps_numbering(convert):
    source = "1. 外层\n   1. 内层甲\n   2. 内层乙\n2. 外层二\n"
    result = convert(source)

    assert _styles_by_text(
        result.output_path, ["外层", "内层甲", "内层乙", "外层二"]
    ) == [
        "List Number",
        "List Number 2",
        "List Number 2",
        "List Number",
    ]
    numbering = numbering_xml(result.output_path)
    assert numbering.count("<w:num ") >= 2


def test_start_attribute_is_ignored_for_bullets(convert):
    result = convert("- 一\n- 二\n")

    assert result.lists == 2
    assert "一" in full_text(result.output_path)


def test_empty_list_item_does_not_crash(convert):
    result = convert("- \n- 二\n")

    assert "二" in full_text(result.output_path)
    assert result.lists >= 1


def test_task_list_with_nested_items(convert):
    source = "- [x] 主项\n  - [ ] 子项\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert "■" in text and "主项" in text
    assert "□" in text and "子项" in text
