"""dockit_mcp tools: convert, read a range, search, and refuse files it did not create."""
from pathlib import Path

import pytest

import dockit_mcp

TEXT_PDF = (Path(__file__).resolve().parent.parent / "packages" / "markitdown" / "tests"
            / "test_files" / "REPAIR-2022-INV-001_multipage.pdf")


def test_convert_read_search_roundtrip():
    summary = dockit_mcp.convert_document(str(TEXT_PDF), ocr="off")
    saved = next(l.split("saved: ", 1)[1] for l in summary.splitlines() if l.startswith("saved: "))
    assert "ZAVA AUTO REPAIR" not in summary  # summary only
    assert "ZAVA AUTO REPAIR" in dockit_mcp.read_converted(saved, 1, 3)
    found = dockit_mcp.search_converted(saved, "zava auto")
    assert not found.startswith("0 match") and "ZAVA AUTO REPAIR" in found


def test_refuses_files_outside_output_folder(tmp_path):
    other = tmp_path / "secret.md"
    other.write_text("x", encoding="utf-8")
    with pytest.raises(ValueError):
        dockit_mcp.read_converted(str(other))
    with pytest.raises(ValueError):
        dockit_mcp.search_converted(str(other), "x")


def test_bad_inputs():
    with pytest.raises(ValueError):
        dockit_mcp.convert_document(str(TEXT_PDF), ocr="bogus")
    with pytest.raises(FileNotFoundError):
        dockit_mcp.convert_document("no_such_file.pdf")
