"""CLI behavior, Mermaid fallbacks, and image embedding."""

from __future__ import annotations

import base64
import io
import json
import shutil
import sys
import urllib.request
from pathlib import Path

import pytest

from docx_helpers import full_text, load, media_names
from scripts.markdown_to_word import (
    MAX_REMOTE_IMAGE_BYTES,
    MermaidRenderer,
    main,
    markdown_to_word,
)


def test_cli_reports_created_document(tmp_path, capsys):
    output = tmp_path / "cli.docx"
    exit_code = main(
        [
            "--markdown-content",
            "# 标题",
            "--output",
            str(output),
            "--mermaid",
            "skip",
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Created:" in captured.out
    assert "Elements:" in captured.out
    assert output.exists()


def test_cli_returns_error_code_on_failure(tmp_path, capsys):
    exit_code = main(
        ["--markdown-file", str(tmp_path / "missing.md"), "--output", str(tmp_path / "x.docx")]
    )
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "Error:" in captured.err


def test_mermaid_skip_preserves_source(convert):
    result = convert("```mermaid\ngraph TD\n  A --> B\n```\n", mermaid_mode="skip")

    assert result.mermaid_rendered == 0
    assert result.mermaid_fallback == 1
    assert result.warnings == []
    text = full_text(result.output_path)
    assert "graph TD" in text
    assert "A --> B" in text


def test_mermaid_auto_without_cli_warns(convert, monkeypatch):
    monkeypatch.setattr(MermaidRenderer, "_find_command", staticmethod(lambda: None))
    result = convert("```mermaid\ngraph TD\n  A --> B\n```\n", mermaid_mode="auto")

    assert result.mermaid_rendered == 0
    assert result.mermaid_fallback == 1
    assert any("Mermaid CLI was not found" in warning for warning in result.warnings)
    assert "graph TD" in full_text(result.output_path)


def test_mermaid_required_without_cli_fails(convert, monkeypatch):
    monkeypatch.setattr(MermaidRenderer, "_find_command", staticmethod(lambda: None))

    with pytest.raises(RuntimeError):
        convert("```mermaid\ngraph TD\n  A --> B\n```\n", mermaid_mode="required")


def test_mermaid_rendered_image_is_embedded(convert, monkeypatch, png_bytes, tmp_path):
    fake_png = tmp_path / "diagram.png"
    fake_png.write_bytes(png_bytes)

    def fake_render_all(self, sources):
        return [(fake_png, None) for _ in sources]

    monkeypatch.setattr(MermaidRenderer, "render_all", fake_render_all)
    result = convert("```mermaid\ngraph TD\n  A --> B\n```\n")

    assert result.mermaid_rendered == 1
    assert result.mermaid_fallback == 0
    assert result.images == 1
    assert media_names(result.output_path)


def _fake_command(monkeypatch, fail_mode: str = "") -> None:
    if fail_mode:
        monkeypatch.setenv("FAKE_MERMAID_FAIL", fail_mode)
    else:
        monkeypatch.delenv("FAKE_MERMAID_FAIL", raising=False)
    script = Path(__file__).with_name("fake_mermaid_cli.py")
    monkeypatch.setattr(
        MermaidRenderer,
        "_find_command",
        staticmethod(lambda: [sys.executable, str(script)]),
    )


def _fake_command(monkeypatch, log_path: Path, fail_mode: str = "") -> None:
    monkeypatch.setenv("FAKE_MERMAID_LOG", str(log_path))
    if fail_mode:
        monkeypatch.setenv("FAKE_MERMAID_FAIL", fail_mode)
    else:
        monkeypatch.delenv("FAKE_MERMAID_FAIL", raising=False)
    script = Path(__file__).with_name("fake_mermaid_cli.py")
    monkeypatch.setattr(
        MermaidRenderer,
        "_find_command",
        staticmethod(lambda: [sys.executable, str(script)]),
    )


def _call_count(log_path: Path) -> int:
    if not log_path.exists():
        return 0
    return len(log_path.read_text(encoding="utf-8").splitlines())


def test_mermaid_batch_uses_single_cli_call(convert, monkeypatch, tmp_path):
    log_path = tmp_path / "calls.log"
    _fake_command(monkeypatch, log_path)
    source = "".join(
        f"```mermaid\ngraph TD\n  A{index} --> B{index}\n```\n\n"
        for index in range(1, 4)
    )
    result = convert(source, mermaid_mode="auto")

    assert result.mermaid_rendered == 3
    assert result.images == 3
    assert _call_count(log_path) == 1


def test_mermaid_falls_back_to_per_diagram_when_batch_fails(
    tmp_path, monkeypatch
):
    log_path = tmp_path / "calls.log"
    _fake_command(monkeypatch, log_path, fail_mode="batch")
    sources = "".join(
        f"```mermaid\ngraph TD\n  A{index} --> B{index}\n```\n\n"
        for index in range(1, 4)
    )
    result = markdown_to_word(
        markdown_content=sources,
        output_path=tmp_path / "fallback.docx",
        overwrite=True,
    )

    assert result.mermaid_rendered == 3
    assert result.mermaid_fallback == 0
    assert _call_count(log_path) == 4  # 1 failed batch + 3 per-diagram retries


def test_mermaid_required_fails_when_cli_errors(convert, monkeypatch, tmp_path):
    _fake_command(monkeypatch, tmp_path / "calls.log", fail_mode="always")

    with pytest.raises(RuntimeError):
        convert("```mermaid\ngraph TD\n  A --> B\n```\n", mermaid_mode="required")


def test_mermaid_source_containing_fence_skips_batch(convert, monkeypatch, tmp_path):
    log_path = tmp_path / "calls.log"
    _fake_command(monkeypatch, log_path)
    source = "````mermaid\ngraph TD\n  A --> B\n```\n````\n"
    result = convert(source, mermaid_mode="auto")

    assert result.mermaid_rendered == 1
    assert _call_count(log_path) == 1


def test_mermaid_auto_falls_back_when_cli_fails(convert, monkeypatch, tmp_path):
    _fake_command(monkeypatch, tmp_path / "calls.log", fail_mode="always")
    result = convert("```mermaid\ngraph TD\n  A --> B\n```\n", mermaid_mode="auto")

    assert result.mermaid_rendered == 0
    assert result.mermaid_fallback == 1
    assert any("failed" in warning for warning in result.warnings)
    assert "graph TD" in full_text(result.output_path)



def test_remote_image_size_limit(convert, monkeypatch):
    class FakeResponse(io.BytesIO):
        headers = {"Content-Length": str(MAX_REMOTE_IMAGE_BYTES + 1)}

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

    monkeypatch.setattr(
        urllib.request,
        "urlopen",
        lambda request, timeout=None: FakeResponse(b"too large"),
    )
    result = convert("![大图](https://example.com/huge.png)\n")

    assert result.images == 0
    assert any("limit" in warning for warning in result.warnings)


def test_remote_image_stream_limit_without_content_length(convert, monkeypatch):
    class FakeResponse(io.BytesIO):
        headers = {}

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

    payload = b"0" * (MAX_REMOTE_IMAGE_BYTES + 10)
    monkeypatch.setattr(
        urllib.request,
        "urlopen",
        lambda request, timeout=None: FakeResponse(payload),
    )
    result = convert("![大图](https://example.com/huge.png)\n")

    assert result.images == 0
    assert any("limit" in warning for warning in result.warnings)


def test_local_image_is_embedded(tmp_path, png_bytes):
    image = tmp_path / "picture.png"
    image.write_bytes(png_bytes)
    source_file = tmp_path / "image.md"
    source_file.write_text("# 图片\n\n![示例](picture.png)\n", encoding="utf-8")

    result = markdown_to_word(
        markdown_file=source_file,
        output_path=tmp_path / "out.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1
    assert media_names(result.output_path)
    assert "示例" in _image_descriptions(result.output_path)


def test_absolute_image_path_is_embedded(tmp_path, png_bytes):
    image = tmp_path / "absolute.png"
    image.write_bytes(png_bytes)

    result = markdown_to_word(
        markdown_content=f"![绝对路径]({image.as_posix()})\n",
        output_path=tmp_path / "absolute.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1
    assert not result.warnings


def test_file_url_image_is_embedded(tmp_path, png_bytes):
    image = tmp_path / "uri.png"
    image.write_bytes(png_bytes)

    result = markdown_to_word(
        markdown_content=f"![file uri]({image.as_uri()})\n",
        output_path=tmp_path / "file-uri.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1


def _image_descriptions(path) -> list[str]:
    document = load(path)
    return [
        element.get("descr") or ""
        for element in document.element.body.iter(
            "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
        )
    ]


def test_missing_local_image_reports_warning(convert):
    result = convert("![缺失](missing.png)\n")

    assert result.images == 0
    assert any("missing.png" in warning for warning in result.warnings)
    assert "missing.png" in full_text(result.output_path)


def test_remote_image_is_downloaded(convert, monkeypatch, png_bytes):
    class FakeResponse(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

    monkeypatch.setattr(
        urllib.request,
        "urlopen",
        lambda request, timeout=None: FakeResponse(png_bytes),
    )
    result = convert("![远程](https://example.com/image.png)\n")

    assert result.images == 1
    assert media_names(result.output_path)


def test_remote_image_disabled_reports_warning(convert):
    result = convert(
        "![远程](https://example.com/image.png)\n", download_images=False
    )

    assert result.images == 0
    assert any("remote image" in warning.lower() for warning in result.warnings)
    assert "example.com/image.png" in full_text(result.output_path)


def test_base64_image_is_embedded(convert, png_bytes):
    encoded = base64.b64encode(png_bytes).decode("ascii")
    result = convert(f"![内嵌](data:image/png;base64,{encoded})\n")

    assert result.images == 1
    assert media_names(result.output_path)


@pytest.mark.skipif(
    shutil.which("rsvg-convert") is not None, reason="rsvg-convert is installed"
)
def test_svg_without_converter_reports_warning(convert, tmp_path):
    svg = tmp_path / "vector.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10">'
        '<rect width="10" height="10"/></svg>',
        encoding="utf-8",
    )

    result = markdown_to_word(
        markdown_content=f"![矢量]({svg})\n",
        output_path=tmp_path / "svg.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 0
    assert any("rsvg-convert" in warning for warning in result.warnings)


def test_grayscale_theme_keeps_code_tokens_black(convert):
    result = convert("```python\nvalue = 1\n```\n")
    document = load(result.output_path)
    color = document.paragraphs[-1].runs[0].font.color.rgb

    assert str(color) == "000000"


def test_mermaid_skip_mode_writes_no_config(tmp_path):
    renderer = MermaidRenderer(
        tmp_path, color="2980B9", font="SimSun", mode="skip", timeout_seconds=30
    )

    assert renderer.command is None
    assert not renderer.config_path.exists()
    assert not renderer.puppeteer_path.exists()
    assert renderer.render("graph TD\n A --> B", 1) == (None, None)
    assert renderer.render_all(["a", "b"]) == [(None, None), (None, None)]


def test_mermaid_auto_mode_writes_configs(tmp_path, monkeypatch):
    monkeypatch.setattr(
        MermaidRenderer, "_find_command", staticmethod(lambda: ["mmdc"])
    )
    renderer = MermaidRenderer(
        tmp_path, color="2980B9", font="SimSun", mode="auto", timeout_seconds=30
    )

    assert renderer.config_path.exists()
    assert renderer.puppeteer_path.exists()
    config = json.loads(renderer.config_path.read_text(encoding="utf-8"))
    assert config["theme"] == "base"
    assert config["themeVariables"]["primaryBorderColor"] == "#2980B9"
    puppeteer = json.loads(renderer.puppeteer_path.read_text(encoding="utf-8"))
    assert puppeteer["headless"] is True


def test_mermaid_environment_sets_puppeteer_paths(tmp_path, monkeypatch):
    monkeypatch.setattr(
        MermaidRenderer, "_find_chrome", staticmethod(lambda: "/usr/bin/chrome")
    )
    renderer = MermaidRenderer(
        tmp_path, color="2980B9", font="SimSun", mode="skip", timeout_seconds=30
    )
    environment = renderer._environment()

    assert environment["PUPPETEER_EXECUTABLE_PATH"] == "/usr/bin/chrome"
    assert environment["PUPPETEER_SKIP_DOWNLOAD"] == "true"


def test_mermaid_timeout_is_stored(tmp_path):
    renderer = MermaidRenderer(
        tmp_path, color="2980B9", font="SimSun", mode="skip", timeout_seconds=7
    )

    assert renderer.timeout_seconds == 7


def test_mermaid_render_returns_error_for_missing_cli(tmp_path, monkeypatch):
    monkeypatch.setattr(MermaidRenderer, "_find_command", staticmethod(lambda: None))
    renderer = MermaidRenderer(
        tmp_path, color="2980B9", font="SimSun", mode="auto", timeout_seconds=30
    )
    path, warning = renderer.render("graph TD\n A --> B", 1)

    assert path is None
    assert "not found" in warning
