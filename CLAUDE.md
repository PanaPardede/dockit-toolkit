# DocKit — Context & Developer Guide for Claude

> **Environment**: Windows 10/11, Python 3.10+  
> **Stack**: Tkinter + ttk + sv-ttk + TkinterDnD2, PyMuPDF (`pymupdf`), `pypdf`, `win32com.client` (Office COM), `olefile`, PyInstaller

---

## 1. What is this project?

**DocKit** is an offline, privacy-first Windows desktop toolkit for document processing, conversion, merging, and optimization. It provides a clean, native 5-tab GUI interface:

1. **Tab 1: Convert to Text (`MarkItDownTab`)**
   - Extracts plain text & Markdown from various document formats (PDF, DOCX, XLSX, PPTX, HTML, Images).
   - Optional AI Vision OCR (Gemini, OpenAI, or any OpenAI-compatible Base URL) for scanned documents and diagrams.
2. **Tab 2: PowerPoint to PDF (`PptToPdfTab`)**
   - Batch converts PowerPoint files (`.pptx`, `.ppt`, `.pptm`, `.pps`, etc.) to PDF using Microsoft PowerPoint COM automation (`PowerPoint.Application`).
   - Integrated optional PDF compression post-conversion.
3. **Tab 3: Word to PDF (`WordToPdfTab`)**
   - Batch converts Microsoft Word and rich text files (`.docx`, `.doc`, `.docm`, `.rtf`, `.odt`) to PDF using Word COM automation (`Word.Application`).
   - Includes automatic detection if Word is not installed, cooperative cancellation, and optional compression.
4. **Tab 4: Merge PDF (`PdfMergerTab`)**
   - Merges multiple PDF files in a user-chosen order (`pypdf.PdfWriter`).
   - Optional post-merge compression.
5. **Tab 5: Compress PDF (`PdfOptimizerTab`)**
   - Standalone PDF compressor with 3 presets: *Light*, *Standard* (default), *Maximum*.
   - Keeps text 100% vector (never rasterizes text), downsamples images via PyMuPDF, strips duplicate streams, subsets fonts, and removes unused metadata.
   - Discards compression if savings are negligible (< 3% and < 500 KB) to protect document quality.

---

## 2. Directory & Codebase Architecture

```text
dockit/
├── markitdown_gui.py          # Application entry point & 5 GUI tabs (~3,100 lines)
├── pdf_optimizer.py           # Pure-logic PDF compression engine (zero GUI imports, ~800 lines)
├── settings.json              # User config, git-ignored (API key lives in Windows Credential Manager)
├── DocKit.spec                # PyInstaller build specification
├── requirements.txt           # Runtime deps (requirements-dev.txt adds pytest, psutil, PyInstaller)
├── packages/
│   ├── markitdown/            # Core document-to-markdown library (modified copy of microsoft/markitdown)
│   └── markitdown-ocr/        # AI Vision OCR plugin (OCRFatalError stops a conversion; transient errors leave a visible marker)
├── tests/
│   ├── conftest.py            # Isolates tests from the real settings.json and keyring
│   ├── test_pdf_optimizer.py  # Unit tests for the compression engine
│   ├── test_gui_integration.py# Integration tests for GUI tabs, workers & COM mocks
│   ├── test_ai_vision.py      # AI Vision error handling, cancel, progress
│   └── stress/                # Stress tests; fixtures are generated, not in git
└── dist/
    └── DocKit.exe             # Built executable (git-ignored)
```

---

## 3. Important Development & Execution Commands

Run these from PowerShell in the repository root:

```powershell
# 1. Run the GUI in development mode
python markitdown_gui.py

# 2. Run unit & integration tests (all must pass)
python -m pytest tests/ -q

# 3. Stress tests (generate fixtures first: python tests/stress/generate_stress_data.py)
python -m pytest tests/stress/ -v

# 4. Run OCR extension tests
python -m pytest packages/markitdown-ocr/tests/ -q

# 5. Build standalone executable (.exe)
python -m PyInstaller DocKit.spec --noconfirm
```

---

## 4. Key Engineering Contracts & Rules

When modifying this codebase, strictly follow these architectural rules:

### A. Separation of Concerns & Public API Contracts
- `pdf_optimizer.py` must remain completely decoupled from GUI libraries (`tkinter`, `ttk`).
- Never break or alter the signatures of public functions in `pdf_optimizer.py`:
  - `optimize_pdf(src_path, dst_path=..., preset_key=..., password=..., on_progress=..., cancel_event=..., ignore_signatures=...) -> OptimizationResult`
  - `inspect_pdf(src_path, password=...) -> InspectInfo`
  - `build_output_path(src_path, custom_dst=...) -> Path`

### B. COM Automation (PowerPoint & Word)
- Always call `pythoncom.CoInitialize()` at the start of worker threads and `pythoncom.CoUninitialize()` in the `finally:` block.
- Suppress popup alerts: `word_app.DisplayAlerts = False` and `powerpoint.DisplayAlerts = 1`.
- Clean up COM references: explicitly set COM objects to `None` before `CoUninitialize()`, and quit app instances if launched by the worker (`app.Quit(SaveChanges=0)`).
- Pre-check password-encrypted files with `is_office_encrypted()` to prevent COM hangs.

### C. Threading & Tkinter Safety
- Tkinter is **not** thread-safe. All UI updates from worker threads must be dispatched via `self.after(0, callback)`.
- Use `threading.Event()` for cooperative cancellation; check `cancel_event.is_set()` between batch items.
- Always guarantee `_is_running = False` and restore button states in worker `finally:` blocks.
- Beware of Python closure late-binding: when passing loop variables into inner callbacks (e.g., `on_progress`), bind them as default parameters (e.g., `def callback(p, cur=cur_file): ...`).

### D. File System Safety on Windows
- Windows holds aggressive file locks. **Never** call `unlink()` immediately followed by `Path.replace()`. Use atomic replaces or write to temporary filenames (`__opt_tmp.pdf`), and clean up `.tmp` files strictly inside `finally:` blocks.
- Save configuration atomically: write to `settings.json.tmp`, perform `fsync()`, then rename/replace to `settings.json`.

### E. Design & Copy Standards (Apple HIG & Anti-Slop)
- **Font**: Use system `Segoe UI` constants (`FONT_BODY`, `FONT_TITLE`, `FONT_BTN`).
- **Layout**: Follow a 14pt grid system (`PAD = 14`, secondary `PAD_SM = 7`).
- **Color Palette**: Use defined semantic colors (`C_BG`, `C_WHITE`, `C_PRIMARY`, `C_DANGER`, `C_TEXT`, `C_SUBTEXT`).
- **Widgets**: Prefer ttk check/radio buttons and entries so the sv-ttk theme applies.
- **Log Messages**: Standardize all log prefixes to: `[OK]`, `[FAIL]`, `[SKIP]`, `[WARN]`, `[INFO]`, and `[ERROR]`.
- **Copy**: Use direct, helpful English instructions (the app is English-only). Do not use AI fluff ("leverage", "seamlessly", "powerful") or internal developer issue tags in user-facing dialogs.
