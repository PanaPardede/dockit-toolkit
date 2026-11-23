---
name: dockit-parse
description: Use before reading any large or binary document (PDF, Word, Excel, PowerPoint, scanned PDF, photo of a document) to save tokens. Converts it to Markdown on the user's computer with DocKit (offline OCR for scans), then read only the parts you need. Trigger on "read/parse/summarize this PDF/docx/xlsx/pptx", scanned files, or long documents.
---

# dockit-parse

Convert the document to Markdown first, then read only what you need. Reading the original (scans cost the most) loads everything into the conversation.

## Preferred: the `dockit` tools (MCP)

If these tools are available, use them:

1. `convert_document(path)` converts the file and returns a short summary: size, estimated tokens, saved path, and an outline of headings with line numbers.
2. `read_converted(file, start_line, max_lines)` reads a line range. `search_converted(file, query)` finds text.

If the summary says about 2,000 tokens or fewer, read the whole file. Otherwise use the outline's line numbers or search. Do not read the original file afterwards.

## Fallback: command line

If you can run commands on the user's Windows computer and the tools above are missing:

```bash
python "<DocKit folder>/dockit_cli.py" "<file>" [--ocr offline|off]
```

It saves `%TEMP%\dockit-parse\<filename>.md` and prints the same summary. Read the saved file by line range or grep it.

## If neither works

Say that DocKit is not available in this session and ask the user to enable the `dockit` MCP server in Claude Desktop or to use Claude Code. Do not guess the document's content.

## Limits

- Offline OCR can misread handwriting, poor scans and accented letters. Check important numbers and names against the original.
- `warning: no text found` means there is no text layer and OCR found nothing. Say so.
- Never use AI Vision for this: it sends the user's files to an outside service.
