"""Shared pytest fixtures for the md2docx regression suite."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.markdown_to_word import ConversionResult, markdown_to_word  # noqa: E402


@pytest.fixture()
def convert(tmp_path):
    """Convert Markdown text to a DOCX and return the ConversionResult."""

    def _convert(markdown: str, **kwargs) -> ConversionResult:
        kwargs.setdefault("mermaid_mode", "skip")
        kwargs.setdefault("overwrite", True)
        kwargs.setdefault("output_path", tmp_path / "output.docx")
        kwargs.setdefault("markdown_content", markdown)
        return markdown_to_word(**kwargs)

    return _convert


@pytest.fixture(scope="session")
def png_bytes() -> bytes:
    """Smallest realistic PNG used by image embedding tests."""
    import io

    from PIL import Image

    stream = io.BytesIO()
    Image.new("RGB", (48, 32), (30, 120, 200)).save(stream, format="PNG")
    return stream.getvalue()
