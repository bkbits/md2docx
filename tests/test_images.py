"""Image embedding: sizing, formats, paths and media bookkeeping."""

from __future__ import annotations

import base64
import io
import shutil
import urllib.request

import pytest
from PIL import Image, features

from docx_helpers import (
    assert_document_integrity,
    full_text,
    image_descriptions,
    image_extents,
    load,
    media_names,
)
from scripts.markdown_to_word import (
    MAX_IMAGE_HEIGHT_INCHES,
    MAX_IMAGE_WIDTH_INCHES,
    markdown_to_word,
)

EMU_PER_INCH = 914400


def _png_bytes(size: tuple[int, int] = (48, 32), mode: str = "RGB") -> bytes:
    stream = io.BytesIO()
    Image.new(mode, size, (10, 120, 200) if mode == "RGB" else (10, 120, 200, 128)).save(
        stream, format="PNG"
    )
    return stream.getvalue()


def _convert_with_image(tmp_path, source: str, **kwargs):
    source_file = tmp_path / "source.md"
    source_file.write_text(source, encoding="utf-8")
    return markdown_to_word(
        markdown_file=source_file,
        output_path=tmp_path / "image.docx",
        overwrite=True,
        mermaid_mode="skip",
        **kwargs,
    )


def test_relative_image_path_uses_markdown_directory(tmp_path):
    (tmp_path / "assets").mkdir()
    (tmp_path / "assets" / "a.png").write_bytes(_png_bytes())
    source_file = tmp_path / "doc.md"
    source_file.write_text("![图](assets/a.png)\n", encoding="utf-8")

    result = markdown_to_word(
        markdown_file=source_file,
        output_path=tmp_path / "out.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1
    assert len(media_names(result.output_path)) == 1


def test_parent_directory_image_path(tmp_path):
    (tmp_path / "shared.png").write_bytes(_png_bytes())
    nested = tmp_path / "docs"
    nested.mkdir()
    source_file = nested / "doc.md"
    source_file.write_text("![图](../shared.png)\n", encoding="utf-8")

    result = markdown_to_word(
        markdown_file=source_file,
        output_path=tmp_path / "parent.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1


def test_percent_encoded_image_path(tmp_path):
    image = tmp_path / "my picture.png"
    image.write_bytes(_png_bytes())
    source_file = tmp_path / "doc.md"
    source_file.write_text("![图](my%20picture.png)\n", encoding="utf-8")

    result = markdown_to_word(
        markdown_file=source_file,
        output_path=tmp_path / "encoded.docx",
        overwrite=True,
        mermaid_mode="skip",
    )

    assert result.images == 1


def test_wide_image_is_clamped_to_text_width(tmp_path):
    image = tmp_path / "wide.png"
    image.write_bytes(_png_bytes((4000, 200)))
    result = _convert_with_image(tmp_path, "![宽图](wide.png)\n")

    cx, cy = image_extents(result.output_path)[0]
    assert cx / EMU_PER_INCH == pytest.approx(MAX_IMAGE_WIDTH_INCHES, abs=0.02)
    assert cy / EMU_PER_INCH <= MAX_IMAGE_HEIGHT_INCHES


def test_tall_image_is_clamped_to_page_height(tmp_path):
    image = tmp_path / "tall.png"
    image.write_bytes(_png_bytes((100, 6000)))
    result = _convert_with_image(tmp_path, "![高图](tall.png)\n")

    cx, cy = image_extents(result.output_path)[0]
    assert cy / EMU_PER_INCH <= MAX_IMAGE_HEIGHT_INCHES + 0.01
    assert cx / EMU_PER_INCH <= MAX_IMAGE_WIDTH_INCHES


def test_small_image_is_not_upscaled_beyond_natural_width(tmp_path):
    image = tmp_path / "small.png"
    image.write_bytes(_png_bytes((32, 32)))
    result = _convert_with_image(tmp_path, "![小图](small.png)\n")

    cx, _ = image_extents(result.output_path)[0]
    assert cx / EMU_PER_INCH == pytest.approx(0.8, abs=0.02)


def test_image_alt_becomes_description(tmp_path):
    image = tmp_path / "a.png"
    image.write_bytes(_png_bytes())
    result = _convert_with_image(tmp_path, '![带 "引号" 的说明](a.png)\n')

    assert '带 "引号" 的说明' in image_descriptions(result.output_path)


def test_empty_alt_leaves_description_empty(tmp_path):
    image = tmp_path / "a.png"
    image.write_bytes(_png_bytes())
    result = _convert_with_image(tmp_path, "![](a.png)\n")

    assert image_descriptions(result.output_path) == [""]


def test_repeated_image_is_embedded_once_in_media(tmp_path):
    image = tmp_path / "a.png"
    image.write_bytes(_png_bytes())
    result = _convert_with_image(tmp_path, "![一](a.png)\n\n![二](a.png)\n")

    assert result.images == 2
    assert len(media_names(result.output_path)) == 1


def test_webp_image_is_converted(tmp_path):
    if not features.check("webp"):
        pytest.skip("Pillow has no WebP support")
    image = tmp_path / "a.webp"
    stream = io.BytesIO()
    Image.new("RGB", (64, 64), (200, 10, 10)).save(stream, format="WEBP")
    image.write_bytes(stream.getvalue())
    result = _convert_with_image(tmp_path, "![webp](a.webp)\n")

    assert result.images == 1
    assert all(name.endswith(".png") for name in media_names(result.output_path))


def test_invalid_base64_data_uri_warns(tmp_path):
    result = _convert_with_image(tmp_path, "![坏图](data:image/png;base64,!!!notbase64!!!)\n")

    assert result.images == 0
    assert result.warnings
    assert "坏图" in full_text(result.output_path)


def test_non_image_data_uri_does_not_produce_image(tmp_path):
    payload = base64.b64encode(b"hello").decode("ascii")
    result = _convert_with_image(tmp_path, f"![文本](data:text/plain;base64,{payload})\n")

    assert result.images == 0
    assert media_names(result.output_path) == []


def test_missing_image_placeholder_format(tmp_path):
    result = _convert_with_image(tmp_path, "![缺失图](nope.png)\n")

    text = full_text(result.output_path)
    assert "缺失图" in text
    assert "nope.png" in text
    assert result.warnings


def test_remote_download_disabled_never_touches_network(tmp_path, monkeypatch):
    calls: list[str] = []

    def fail(*args, **kwargs):
        calls.append("called")
        raise AssertionError("network must not be used")

    monkeypatch.setattr(urllib.request, "urlopen", fail)
    result = _convert_with_image(
        tmp_path, "![远程](https://example.com/a.png)\n", download_images=False
    )

    assert calls == []
    assert result.images == 0
    assert any("remote image" in warning.lower() for warning in result.warnings)


def test_remote_image_content_type_not_trusted(tmp_path, monkeypatch):
    class FakeResponse(io.BytesIO):
        headers = {"Content-Type": "text/html"}

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

    monkeypatch.setattr(
        urllib.request,
        "urlopen",
        lambda request, timeout=None: FakeResponse(b"<html>not an image</html>"),
    )
    result = _convert_with_image(tmp_path, "![伪装](https://example.com/a.png)\n")

    assert result.images == 0
    assert result.warnings


def test_svg_with_converter_or_warning(tmp_path):
    svg = tmp_path / "vector.svg"
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="10">'
        '<rect width="20" height="10" fill="red"/></svg>',
        encoding="utf-8",
    )
    result = _convert_with_image(tmp_path, "![矢量](vector.svg)\n")

    if shutil.which("rsvg-convert"):
        assert result.images == 1
    else:
        assert result.images == 0
        assert any("rsvg-convert" in warning for warning in result.warnings)


def test_image_document_is_structurally_valid(tmp_path):
    image = tmp_path / "a.png"
    image.write_bytes(_png_bytes())
    result = _convert_with_image(
        tmp_path,
        "# 图\n\n![一](a.png)\n\n| 列 |\n| --- |\n| ![二](a.png) |\n",
    )

    assert_document_integrity(result.output_path)
    assert result.images == 2


def test_image_inside_list_item(tmp_path):
    image = tmp_path / "a.png"
    image.write_bytes(_png_bytes())
    result = _convert_with_image(tmp_path, "- 项目 ![图标](a.png)\n")

    assert result.images == 1
    document = load(result.output_path)
    assert "项目" in document.paragraphs[0].text


def test_corrupt_local_image_warns(tmp_path):
    (tmp_path / "broken.png").write_bytes(b"not a png at all")
    result = _convert_with_image(tmp_path, "![坏](broken.png)\n")

    assert result.images == 0
    assert result.warnings
