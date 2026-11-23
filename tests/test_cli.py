"""dockit_cli: converts to a file and prints a short summary, never the whole text."""
from pathlib import Path

import dockit_cli

TEXT_PDF = (Path(__file__).resolve().parent.parent / "packages" / "markitdown" / "tests"
            / "test_files" / "REPAIR-2022-INV-001_multipage.pdf")


def test_converts_to_file_and_prints_summary(tmp_path, capsys):
    assert dockit_cli.main([str(TEXT_PDF), "--out", str(tmp_path), "--ocr", "off"]) == 0
    out = capsys.readouterr().out
    saved = tmp_path / f"{TEXT_PDF.name}.md"
    assert "ZAVA AUTO REPAIR" in saved.read_text(encoding="utf-8")
    assert "tokens" in out and str(saved) in out
    assert "ZAVA AUTO REPAIR" not in out  # summary only, not the content


def test_missing_file_fails_without_crashing(tmp_path, capsys):
    assert dockit_cli.main([str(tmp_path / "nope.pdf"), "--out", str(tmp_path)]) == 1
    assert "FAILED" in capsys.readouterr().out
