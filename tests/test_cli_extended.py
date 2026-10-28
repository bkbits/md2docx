"""CLI argument handling beyond the basic happy path."""

from __future__ import annotations

import pytest

from docx_helpers import full_text, load
from scripts.markdown_to_word import main


def test_help_exits_zero(capsys):
    with pytest.raises(SystemExit) as error:
        main(["--help"])

    assert error.value.code == 0
    assert "Convert Markdown" in capsys.readouterr().out


def test_sources_are_mutually_exclusive(tmp_path):
    with pytest.raises(SystemExit) as error:
        main(
            [
                "--markdown-content",
                "# x",
                "--markdown-file",
                str(tmp_path / "a.md"),
            ]
        )

    assert error.value.code != 0


def test_missing_source_exits_nonzero():
    with pytest.raises(SystemExit) as error:
        main([])

    assert error.value.code != 0


def test_content_mode_default_output_name(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    exit_code = main(["--markdown-content", "# 标题", "--mermaid", "skip"])

    assert exit_code == 0
    assert (tmp_path / "markdown-output.docx").exists()
    assert "Created:" in capsys.readouterr().out


def test_file_mode_default_output_name(tmp_path):
    source = tmp_path / "notes.markdown"
    source.write_text("# 标题\n", encoding="utf-8")

    exit_code = main(["--markdown-file", str(source), "--mermaid", "skip"])

    assert exit_code == 0
    assert (tmp_path / "notes-转换版.docx").exists()


def test_cli_passes_style_options(tmp_path):
    output = tmp_path / "styled.docx"
    exit_code = main(
        [
            "--markdown-content",
            "# 标题",
            "--color",
            "classic-blue",
            "--font",
            "Arial",
            "--font-size",
            "10",
            "--output",
            str(output),
            "--mermaid",
            "skip",
        ]
    )

    assert exit_code == 0
    document = load(output)
    assert document.styles["Normal"].font.size.pt == 10
    assert str(document.styles["Heading 1"].font.color.rgb) == "2980B9"


def test_cli_cover_and_toc(tmp_path):
    output = tmp_path / "cover.docx"
    exit_code = main(
        [
            "--markdown-content",
            "# 标题\n\n## 章节\n\n正文\n",
            "--output",
            str(output),
            "--cover",
            "--toc",
            "--mermaid",
            "skip",
        ]
    )

    assert exit_code == 0
    text = full_text(output)
    assert text.count("标题") == 1
    assert "目录" in text


def test_cli_overwrite_flag_required(tmp_path, capsys):
    output = tmp_path / "exists.docx"
    output.write_bytes(b"placeholder")

    exit_code = main(
        ["--markdown-content", "# x", "--output", str(output), "--mermaid", "skip"]
    )
    assert exit_code == 1
    assert "Error:" in capsys.readouterr().err

    exit_code = main(
        [
            "--markdown-content",
            "# x",
            "--output",
            str(output),
            "--overwrite",
            "--mermaid",
            "skip",
        ]
    )
    assert exit_code == 0


def test_cli_no_download_images_flag(tmp_path):
    output = tmp_path / "offline.docx"
    exit_code = main(
        [
            "--markdown-content",
            "![图](https://example.com/a.png)",
            "--output",
            str(output),
            "--no-download-images",
            "--mermaid",
            "skip",
            "--overwrite",
        ]
    )

    assert exit_code == 0
    assert "example.com" in full_text(output)


def test_cli_invalid_color_reports_error(tmp_path, capsys):
    exit_code = main(
        [
            "--markdown-content",
            "# x",
            "--color",
            "not-a-color",
            "--output",
            str(tmp_path / "x.docx"),
            "--mermaid",
            "skip",
        ]
    )

    assert exit_code == 1
    assert "Invalid color" in capsys.readouterr().err


def test_cli_warnings_go_to_stderr(tmp_path, capsys):
    exit_code = main(
        [
            "--markdown-content",
            "![缺](missing.png)",
            "--output",
            str(tmp_path / "warn.docx"),
            "--mermaid",
            "skip",
            "--overwrite",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Warnings:" in captured.err
    assert "missing.png" in captured.err
