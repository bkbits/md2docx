"""Block-level constructs: headings, quotes, rules, footnotes, formatting."""

from __future__ import annotations

from docx_helpers import (
    anchors,
    bookmarks,
    bordered_paragraphs,
    full_text,
    load,
    paragraphs,
    single_run,
    texts,
)
from docx_helpers import assert_document_integrity, link_targets


def test_block_html_comment_is_not_rendered(convert):
    source = "<!-- 这是 HTML 注释 -->\n\n可见文本。\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert "这是 HTML 注释" not in text
    assert "可见文本" in text


def test_inline_html_comment_is_not_rendered(convert):
    result = convert("行内注释<!-- 隐藏 -->后面的文字\n")

    text = full_text(result.output_path)
    assert "隐藏" not in text
    assert "行内注释" in text
    assert "后面的文字" in text


def test_yaml_front_matter_is_stripped(convert):
    source = "---\ntitle: 测试\nauthor: me\n---\n\n正文内容。\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert "title:" not in text
    assert "author:" not in text
    assert "正文内容" in text


def test_horizontal_rule_document_is_not_stripped(convert):
    result = convert("---\n\n正文内容。\n\n---\n")

    assert "正文内容" in full_text(result.output_path)


def test_heading_levels_map_to_word_styles(convert):
    source = "\n\n".join(f"{'#' * level} 标题{level}" for level in range(1, 7))
    result = convert(source + "\n")
    document = load(result.output_path)

    assert result.headings == 6
    for level in range(1, 7):
        paragraph = next(p for p in document.paragraphs if p.text == f"标题{level}")
        assert paragraph.style.name == f"Heading {level}"
    sizes = [document.styles[f"Heading {level}"].font.size.pt for level in range(1, 7)]
    assert sizes == [22, 18, 16, 14, 13, 12]


def test_duplicate_headings_get_unique_bookmarks(convert):
    source = "## 重复标题\n\n## 重复标题\n"
    result = convert(source)

    names = list(bookmarks(result.output_path))
    assert len(names) == 2
    assert len(set(names)) == 2


def test_blockquote_gets_indent_border_and_shading(convert):
    result = convert("> 引用内容\n")

    paragraph = next(
        item for item in load(result.output_path).paragraphs if "引用内容" in item.text
    )
    assert paragraph.paragraph_format.left_indent is not None
    assert any(
        "引用内容" in text and "left" in edges
        for text, edges in bordered_paragraphs(result.output_path)
    )
    shading = paragraph._p.pPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"
    )
    assert shading is not None


def test_horizontal_rule_renders_bottom_border(convert):
    result = convert("上\n\n---\n\n下\n")

    edges = [edge for _, edges in bordered_paragraphs(result.output_path) for edge in edges]
    assert "bottom" in edges


def test_footnote_reference_and_backlink_resolve(convert):
    source = "正文引用[^1]。\n\n[^1]: 脚注内容。\n"
    result = convert(source)

    text = full_text(result.output_path)
    assert "正文引用" in text
    assert "脚注内容" in text

    document_bookmarks = bookmarks(result.output_path)
    for anchor in anchors(result.output_path):
        assert anchor in document_bookmarks, f"dangling anchor {anchor}"
    assert len(document_bookmarks) >= 2


def test_footnote_separator_is_rendered_once(convert):
    source = "正文[^1]\n\n[^1]: 脚注内容\n"
    result = convert(source)

    bottom_borders = [
        text
        for text, edges in bordered_paragraphs(result.output_path)
        if "bottom" in edges
    ]
    assert len(bottom_borders) == 1


def test_inline_formatting_runs(convert):
    source = (
        "**粗体** *斜体* ~~删除线~~ `code` "
        "<u>下划线</u> <sub>下标</sub> <sup>上标</sup>\n"
    )
    result = convert(source)

    assert single_run(result.output_path, "粗体").bold is True
    assert single_run(result.output_path, "斜体").italic is True
    assert single_run(result.output_path, "删除线").font.strike is True
    assert single_run(result.output_path, "code").font.name == "Consolas"
    assert single_run(result.output_path, "下划线").underline is True
    assert single_run(result.output_path, "下标").font.subscript is True
    assert single_run(result.output_path, "上标").font.superscript is True


def test_hard_break_creates_line_break(convert):
    result = convert("第一行  \n第二行\n")

    paragraph = next(
        item
        for item in load(result.output_path).paragraphs
        if "第一行" in item.text
    )
    assert "\n" in paragraph.text


def test_strikethrough_inside_bold_is_kept(convert):
    result = convert("**粗体 ~~删除~~ 结束**\n")

    run = single_run(result.output_path, "删除")
    assert run.font.strike is True
    assert run.bold is True


def test_cover_without_h1_uses_fallback_title(convert):
    result = convert("## 只有二级标题\n\n正文\n", cover=True)
    document = load(result.output_path)

    assert document.paragraphs[0].text == "未命名文档"
    assert result.headings == 1


def test_cover_removes_h1_wherever_it_is(convert):
    source = "前言\n\n正文\n\n# 真正的标题\n\n后续\n"
    result = convert(source, cover=True)

    text = full_text(result.output_path)
    assert text.count("真正的标题") == 1
    assert load(result.output_path).paragraphs[0].text == "真正的标题"


def test_internal_link_to_missing_anchor_does_not_crash(convert):
    result = convert("[无效](#nope)\n")

    assert "无效" in full_text(result.output_path)
    assert_document_integrity(result.output_path)


def test_internal_link_with_percent_encoded_cjk(convert):
    result = convert("[跳转](#%E4%B8%AD%E6%96%87%E6%A0%87%E9%A2%98)\n\n## 中文标题\n")
    document_bookmarks = bookmarks(result.output_path)

    assert set(anchors(result.output_path)) <= set(document_bookmarks)


def test_heading_with_inline_code_is_linkable(convert):
    source = "## 使用 `code` 的标题\n\n[跳转](#使用-code-的标题)\n"
    result = convert(source)
    document_bookmarks = bookmarks(result.output_path)

    assert set(anchors(result.output_path)) <= set(document_bookmarks)


def test_external_link_with_query_string_is_preserved(convert):
    result = convert("[查询](https://example.com/a?x=1&y=2)\n")

    assert "https://example.com/a?x=1&y=2" in link_targets(result.output_path)


def test_javascript_link_is_not_converted(convert):
    result = convert("[危险](javascript:alert(1))\n")
    document_xml_text = result.output_path.read_bytes()

    assert "javascript:alert" in full_text(result.output_path)
    assert not anchors(result.output_path)
    assert b'w:anchor' not in document_xml_text


def test_multiple_footnotes_get_distinct_anchors(convert):
    source = "甲[^1] 乙[^2]\n\n[^1]: 一\n\n[^2]: 二\n"
    result = convert(source)
    document_bookmarks = bookmarks(result.output_path)
    link_anchors = anchors(result.output_path)

    assert "一" in full_text(result.output_path)
    assert "二" in full_text(result.output_path)
    assert len(set(link_anchors)) >= 4
    assert set(link_anchors) <= set(document_bookmarks)


def test_heading_with_three_duplicates(convert):
    result = convert("## 同\n\n## 同\n\n## 同\n")
    names = list(bookmarks(result.output_path))

    assert len(names) == 3
    assert len(set(names)) == 3


def test_nested_markup_combinations(convert):
    source = "**粗 *斜* 粗** 与 ~~删 `码` 除~~\n"
    result = convert(source)

    assert single_run(result.output_path, "斜").italic is True
    assert single_run(result.output_path, "码").font.name == "Consolas"


def test_paragraph_text_has_no_trailing_newline(convert):
    result = convert("普通段落\n\n另一段\n")

    assert all(not text.endswith("\n") for text in texts(result.output_path))


def test_empty_paragraph_between_blocks_is_ignored(convert):
    result = convert("第一段\n\n\n\n第二段\n")

    assert [text for text in texts(result.output_path) if text.strip()] == [
        "第一段",
        "第二段",
    ]


def test_code_block_is_one_paragraph_with_line_breaks(convert):
    source = "```\n第一行\n第二行\n第三行\n```\n"
    result = convert(source)
    document = load(result.output_path)

    assert len(document.paragraphs) == 1
    paragraph = document.paragraphs[0]
    assert paragraph.text == "第一行\n第二行\n第三行"
    assert paragraph._p.findall(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br"
    )
    assert not paragraph.paragraph_format.keep_with_next


def test_code_block_keeps_blank_lines(convert):
    result = convert("```\na\n\nb\n```\n")
    document = load(result.output_path)

    assert document.paragraphs[0].text == "a\n\nb"


def test_code_block_keeps_indentation(convert):
    result = convert("```python\ndef f():\n    return 1\n```\n")
    document = load(result.output_path)

    assert document.paragraphs[0].text == "PYTHON"
    assert "\n    return 1" in document.paragraphs[1].text


def test_multiline_string_keeps_single_highlighting(convert):
    source = '```python\ntext = """first\nsecond\nthird\"""\n```\n'
    result = convert(source, color="classic-blue")
    document = load(result.output_path)
    code_paragraph = document.paragraphs[1]

    def color_of(needle: str):
        for run in code_paragraph.runs:
            if needle in run.text:
                return str(run.font.color.rgb)
        raise AssertionError(f"no run contains {needle!r}")

    assert color_of("first") == "A31515"
    assert color_of("second") == "A31515"
    assert color_of("third") == "A31515"


def test_code_block_label_and_body_spacing(convert):
    result = convert("```bash\necho hi\n```\n")
    document = load(result.output_path)

    assert document.paragraphs[0].text == "BASH"
    body = document.paragraphs[1]
    assert body.text == "echo hi"
    assert body.paragraph_format.space_before.pt == 0


def test_empty_code_block_still_renders(convert):
    result = convert("```\n\n```\n")
    document = load(result.output_path)

    assert len(document.paragraphs) == 1
    assert document.paragraphs[0].text == " "
