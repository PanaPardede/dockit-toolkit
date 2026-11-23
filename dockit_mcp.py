"""DocKit as a local MCP server, for Claude Desktop and other MCP clients.

Run by the client over stdio:  python dockit_mcp.py

Tools: convert_document (file -> Markdown on disk + short summary),
read_converted (line range) and search_converted (find text), so the assistant
loads only the parts of a large document that it needs.
"""

import logging
import sys
import tempfile
from pathlib import Path

logging.disable(logging.INFO)  # stdout is the protocol channel; keep libraries quiet

import dockit_cli

try:
    from mcp.server.mcpserver import MCPServer as _Server  # mcp 2.x
except ImportError:
    from mcp.server.fastmcp import FastMCP as _Server  # mcp 1.x

OUT_DIR = Path(tempfile.gettempdir()) / "dockit-parse"
mcp = _Server("dockit")


def _converted_file(file: str) -> Path:
    p = Path(file).expanduser().resolve()
    if OUT_DIR.resolve() not in p.parents or not p.is_file():
        raise ValueError("Only files made by convert_document can be read. Convert the document first.")
    return p


@mcp.tool()
def convert_document(path: str, ocr: str = "offline") -> str:
    """Convert a PDF, Word, Excel, PowerPoint or image file on this computer to Markdown.

    Scans and images are read with offline OCR (ocr="offline", the default; use "off" to skip it).
    Returns a short summary with the saved path, size, estimated tokens and a heading outline with
    line numbers. Read the result with read_converted or search_converted instead of asking for the whole text.
    """
    if ocr not in ("offline", "off"):
        raise ValueError('ocr must be "offline" or "off"')
    src = Path(path).expanduser()
    if not src.is_file():
        raise FileNotFoundError(f"not a file: {src}")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    return dockit_cli.convert_one(src, OUT_DIR, ocr)


@mcp.tool()
def read_converted(file: str, start_line: int = 1, max_lines: int = 200) -> str:
    """Read a line range of a Markdown file made by convert_document (line numbers start at 1)."""
    lines = _converted_file(file).read_text(encoding="utf-8").splitlines()
    start = max(1, start_line)
    chunk = lines[start - 1 : start - 1 + max(1, min(max_lines, 1000))]
    end = start + len(chunk) - 1
    return f"lines {start}-{end} of {len(lines)}\n" + "\n".join(chunk)


@mcp.tool()
def search_converted(file: str, query: str, max_hits: int = 20) -> str:
    """Case-insensitive text search in a Markdown file made by convert_document. Returns 'line: text'."""
    q = query.lower()
    hits = [f"{n}: {l[:200]}" for n, l in enumerate(_converted_file(file).read_text(encoding="utf-8").splitlines(), 1)
            if q in l.lower()]
    shown = hits[: max(1, max_hits)]
    return f"{len(hits)} match(es)" + ("\n" + "\n".join(shown) if shown else "")


if __name__ == "__main__":
    sys.exit(mcp.run())
