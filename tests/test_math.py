"""Formula handling: currency, code contexts and LaTeX rendering."""

from __future__ import annotations

import pytest

from docx_helpers import document_xml, full_text, load
from scripts.markdown_to_word import FormulaRenderer


def _has_matplotlib() -> bool:
    try:
        import matplotlib.mathtext  # noqa: F401
    except ImportError:  # pragma: no cover - optional dependency
        return False
    return True


requires_matplotlib = pytest.mark.skipif(
    not _has_matplotlib(), reason="matplotlib is not installed"
)


def test_currency_amounts_are_not_treated_as_formulas(convert):
    source = "价格是 $5，另一个是 $10，合计 $15。"
    result = convert(source)

    assert result.formulas_rendered == 0
    assert result.formulas_fallback == 0
    assert result.images == 0
    assert full_text(result.output_path) == source


def test_inline_code_with_dollars_is_preserved(convert):
    source = "行内代码：`echo $HOME` 和 `$PATH`。"
    result = convert(source)

    assert result.formulas_rendered == 0
    assert result.formulas_fallback == 0
    text = full_text(result.output_path)
    assert "echo $HOME" in text
    assert "$PATH" in text
    assert "<span" not in text
    assert "data-latex" not in text


def test_indented_code_block_is_never_parsed_as_math(convert):
    source = "    $$x = 1$$\n    $y = 2$\n"
    result = convert(source)

    assert result.formulas_rendered == 0
    assert result.formulas_fallback == 0
    text = full_text(result.output_path)
    assert "$$x = 1$$" in text
    assert "$y = 2$" in text


def test_fenced_code_block_is_never_parsed_as_math(convert):
    source = "```python\nvalue = '$HOME'\n$$\nnot math\n$$\n```\n"
    result = convert(source)

    assert result.formulas_rendered == 0
    assert result.code_blocks == 1
    text = full_text(result.output_path)
    assert "'$HOME'" in text
    assert "not math" in text


def test_longer_closing_fence_is_supported(convert):
    source = "```text\n$$keep$$\n````\n"
    result = convert(source)

    assert result.code_blocks == 1
    assert result.formulas_rendered == 0
    assert "$$keep$$" in full_text(result.output_path)


@requires_matplotlib
def test_inline_formula_is_rendered(convert):
    result = convert("行内 $E=mc^2$ 结束。")

    assert result.formulas_rendered == 1
    assert result.formulas_fallback == 0
    paragraph = next(
        item
        for item in load(result.output_path).paragraphs
        if "行内" in item.text
    )
    assert "行内" in paragraph.text
    assert "结束" in paragraph.text
    assert paragraph._p.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")


@requires_matplotlib
def test_block_formula_is_rendered(convert):
    result = convert("$$\na^2 + b^2 = c^2\n$$\n")

    assert result.formulas_rendered == 1
    assert result.images == 1


@requires_matplotlib
def test_block_formula_inside_list_item_keeps_item_text(convert):
    source = "1. 项目一\n\n    $$\n    y = 1\n    $$\n\n2. 项目二\n"
    result = convert(source)

    assert result.formulas_rendered == 1
    text = full_text(result.output_path)
    assert "项目一" in text
    assert "项目二" in text
    assert result.lists == 2


def test_formula_falls_back_to_source_when_unavailable(convert, monkeypatch):
    monkeypatch.setattr(FormulaRenderer, "render", lambda self, latex, block: None)
    result = convert("行内 $x^2$ 结束。")

    assert result.formulas_rendered == 0
    assert result.formulas_fallback == 1
    assert "x^2" in full_text(result.output_path)


def test_unclosed_block_formula_stays_as_text(convert):
    result = convert("$$\nopen only\n")

    assert result.formulas_rendered == 0
    assert "$$" in full_text(result.output_path)


def test_escaped_dollar_is_literal_text(convert):
    source = r"价格为 \$5 与 \$10。"
    result = convert(source)

    assert result.formulas_rendered == 0
    text = full_text(result.output_path)
    assert "$5" in text
    assert "$10" in text
    assert "\\$" not in text


@requires_matplotlib
def test_inline_formula_adjacent_to_cjk(convert):
    result = convert("公式$E=mc^2$。")

    assert result.formulas_rendered == 1
    assert result.formulas_fallback == 0


@requires_matplotlib
def test_math_markup_never_leaks_as_html(convert):
    result = convert("行内 $x$ 与\n\n$$\ny\n$$\n")
    xml = document_xml(result.output_path)

    assert "data-latex" not in xml
    assert "math-inline" not in xml


@requires_matplotlib
def test_identical_formulas_reuse_one_render(convert, monkeypatch):
    import matplotlib.mathtext as mathtext

    calls: list[str] = []
    original = mathtext.math_to_image

    def counting(expression, *args, **kwargs):
        calls.append(str(expression))
        return original(expression, *args, **kwargs)

    monkeypatch.setattr(mathtext, "math_to_image", counting)
    result = convert("$x^2$ 与 $x^2$ 以及 $y^2$\n")

    assert result.formulas_rendered == 3
    assert len(calls) == 2
    assert len(set(calls)) == 2


@requires_matplotlib
def test_formula_render_failure_falls_back(convert, monkeypatch):
    import matplotlib.mathtext as mathtext

    def broken(*args, **kwargs):
        raise RuntimeError("mathtext failure")

    monkeypatch.setattr(mathtext, "math_to_image", broken)
    result = convert("行内 $\\badcommand{x}$ 结束\n")

    assert result.formulas_rendered == 0
    assert result.formulas_fallback == 1
    assert "badcommand" in full_text(result.output_path)


@requires_matplotlib
def test_single_line_block_formula(convert):
    result = convert("$$x^2$$\n")

    assert result.formulas_rendered == 1
    assert result.images == 1


@requires_matplotlib
def test_block_formula_inside_blockquote(convert):
    source = "> 引用\n>\n> $$\n> a^2\n> $$\n"
    result = convert(source)

    assert result.formulas_rendered == 1
    assert "引用" in full_text(result.output_path)


@requires_matplotlib
def test_formula_with_angle_brackets(convert):
    result = convert("$$a < b$$\n")

    assert result.formulas_rendered == 1
    assert result.images == 1


def test_inline_math_does_not_span_lines(convert):
    result = convert("价格 $5 与\n$10 两行\n")

    assert result.formulas_rendered == 0
    assert "$5" in full_text(result.output_path)


def test_dollar_in_link_url_stays_literal(convert):
    result = convert("[链接](https://example.com/?q=$x$)\n")

    assert result.formulas_rendered == 0


def test_formula_in_table_cell(convert):
    result = convert("| 公式 |\n| --- |\n| $x$ |\n")

    assert result.tables == 1
    assert result.formulas_fallback + result.formulas_rendered == 1
