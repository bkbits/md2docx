"""Stand-in for the mermaid CLI used by the test suite.

Reads the same ``-i``/``-o`` arguments as ``mmdc`` and writes a tiny PNG.
When the input is a Markdown file (batch mode) it emits ``-1``, ``-2`` files,
mirroring the real CLI. Every invocation appends a line to ``calls.log`` so
tests can assert how many processes were started.
"""

from __future__ import annotations

import base64
import os
import re
import sys
from pathlib import Path

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGM8YWPDwMDAxAAGAA+2AUSPIZanAAAAAElFTkSuQmCC"
)

VALUE_OPTIONS = {
    "-i",
    "--input",
    "-o",
    "--output",
    "-e",
    "--outputFormat",
    "-b",
    "--backgroundColor",
    "-s",
    "--scale",
    "--size",
    "-c",
    "--configFile",
    "-p",
    "--puppeteerConfigFile",
}


def parse_arguments(argv: list[str]) -> dict[str, str]:
    options: dict[str, str] = {}
    index = 0
    while index < len(argv):
        argument = argv[index]
        if argument in VALUE_OPTIONS and index + 1 < len(argv):
            options[argument] = argv[index + 1]
            index += 2
            continue
        index += 1
    return options


def main(argv: list[str]) -> int:
    options = parse_arguments(argv)
    input_path = options.get("-i") or options.get("--input")
    output_path = options.get("-o") or options.get("--output")
    if not input_path or not output_path:
        print("fake mmdc: missing -i or -o", file=sys.stderr)
        return 2

    source = Path(input_path)
    destination = Path(output_path)
    log_path = os.environ.get("FAKE_MERMAID_LOG")
    if log_path:
        log = Path(log_path)
    else:
        log = source.with_name("calls.log")
    with log.open("a", encoding="utf-8") as handle:
        handle.write(f"{input_path}\n")

    fail_mode = os.environ.get("FAKE_MERMAID_FAIL", "")
    is_markdown = source.suffix.lower() in {".md", ".markdown"}
    if fail_mode == "batch" and is_markdown:
        print("fake mmdc: batch failure requested", file=sys.stderr)
        return 1
    if fail_mode == "always":
        print("fake mmdc: failure requested", file=sys.stderr)
        return 1

    if is_markdown:
        text = source.read_text(encoding="utf-8")
        count = len(re.findall(r"^```mermaid\s*$", text, re.MULTILINE))
        if count == 0:
            print("fake mmdc: no diagrams found", file=sys.stderr)
            return 1
        for index in range(1, count + 1):
            candidate = destination.with_name(f"{destination.stem}-{index}{destination.suffix}")
            candidate.write_bytes(PNG)
        return 0

    destination.write_bytes(PNG)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
