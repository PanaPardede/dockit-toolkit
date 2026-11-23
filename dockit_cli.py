"""DocKit from the command line: convert documents to Markdown files.

    python dockit_cli.py report.pdf slides.pptx [--out DIR] [--ocr offline|off]

Each result is saved as <out>/<filename>.md and only a short summary is printed,
so a caller (a person or an AI assistant) can read just the parts it needs.
Uses the same converter as the DocKit app, including offline OCR for scans.
"""

import argparse
import re
import sys
import tempfile
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

MAX_OUTLINE = 40


def convert_one(path: Path, out_dir: Path, ocr: str) -> str:
    from markitdown_gui import create_markitdown_instance, estimate_tokens, needs_ocr

    md, err, _ = create_markitdown_instance({"ocr_engine": ocr})
    if md is None:
        raise RuntimeError(err)
    text = (md.convert(str(path)).markdown or "").strip()
    out = out_dir / f"{path.name}.md"
    out.write_text(text, encoding="utf-8")

    used_ocr = ocr == "offline" and needs_ocr(path)
    lines = [f"{path.name}: {len(text):,} chars, ~{estimate_tokens(text):,} tokens"
             f"{', OCR used' if used_ocr else ''}", f"saved: {out}"]
    if not text:
        lines.append("warning: no text found" + ("" if ocr == "offline" else " (try --ocr offline)"))
    headings = [(i + 1, l) for i, l in enumerate(text.splitlines()) if re.match(r"#{1,3} \S", l)]
    if headings:
        lines.append(f"outline (line: heading), first {min(len(headings), MAX_OUTLINE)} of {len(headings)}:")
        lines += [f"  {n}: {l}" for n, l in headings[:MAX_OUTLINE]]
    return "\n".join(lines)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+", type=Path)
    ap.add_argument("--out", type=Path, default=Path(tempfile.gettempdir()) / "dockit-parse")
    ap.add_argument("--ocr", choices=["offline", "off"], default="offline",
                    help="offline = OCR scans and images on this computer (default)")
    args = ap.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    args.out.mkdir(parents=True, exist_ok=True)

    failed = 0
    for f in args.files:
        try:
            if not f.is_file():
                raise FileNotFoundError(f"not a file: {f}")
            print(convert_one(f, args.out, args.ocr))
        except Exception as e:
            failed += 1
            print(f"{f.name}: FAILED: {e}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
