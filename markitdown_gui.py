"""
DocKit - offline document toolkit for Windows.
Convert to text, PowerPoint to PDF, Word to PDF, merge PDF, compress PDF.
"""

import json
import logging
import os
import queue
import re
import subprocess
import sys
import threading
import time
import warnings
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk, scrolledtext
from pathlib import Path

import pdf_optimizer
from pdf_optimizer import (
    PRESETS,
    format_size,
    inspect_pdf,
    build_output_path,
    optimize_pdf,
    setup_optimizer_logging,
    PasswordRequired,
    CancelledError,
    DigitalSignatureWarning,
    CorruptPdf,
    has_digital_signatures,
)

logging.getLogger("pypdf").setLevel(logging.ERROR)
logger = logging.getLogger("MarkItDownGUI")

APP_VERSION     = "v3"
APP_NAME        = "DocKit"
KEYRING_SERVICE = "MarkItDownGUI"  # unchanged so existing users keep their saved key
KEYRING_USER    = "api_key"

FILE_WARN_MB   = 50
FILE_LIMIT_MB  = 500

# Suggestions only: the model box is editable, so users can type any newer model.
GEMINI_MODELS = [
    "gemini-3.6-flash",
    "gemini-3.8-flash",
]
OPENAI_MODELS = ["gpt-4.1-mini", "gpt-4.1", "gpt-4o-mini", "gpt-5-mini"]
GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/openai/"
# One prompt for both scanned-page OCR and embedded images: transcribe first, then describe visuals.
AI_VISION_PROMPT = (
    "Transcribe all text in this image exactly as written, in its original language, "
    "keeping the reading order. Format tables as Markdown tables. If the image contains "
    "a chart, diagram or photo, add a short description of it after the text. "
    "Return only the result, with no preamble."
)
PPT_EXTENSIONS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx"}

# ── Colour palette ────────────────────────────────────────────────────────────
C_BG        = "#FAFAFA"     # matches the sv-ttk light background
C_WHITE     = "#FFFFFF"     # card / panel
C_SURFACE   = "#FAFAFA"
C_BORDER    = "#D0D0D0"     # neutral border
C_BORDER_H  = "#1565C0"     # hover blue
C_PRIMARY   = "#1565C0"     # brand blue
C_PRIMARY_H = "#0D47A1"
C_DANGER    = "#C62828"     # brand red
C_SUCCESS   = "#2E7D32"
C_TEXT      = "#212121"     # near black
C_SUBTEXT   = "#757575"     # grey
C_BTN_SEC   = "#ECECEC"
C_BTN_SECH  = "#DCDCDC"
C_SEP       = "#E0E0E0"

FONT_TITLE_LARGE = ("Segoe UI", 12, "bold")
FONT_TITLE       = ("Segoe UI", 11, "bold")
FONT_SUBTITLE    = ("Segoe UI", 10, "bold")
FONT_BODY        = ("Segoe UI", 10)
FONT_BTN         = ("Segoe UI", 10, "bold")
FONT_SMALL       = ("Segoe UI", 9)
FONT_LINK        = ("Segoe UI", 9, "underline")
FONT_STATUS      = ("Segoe UI", 9, "italic")
FONT_CAPTION     = ("Segoe UI", 8)
FONT_MONO        = ("Consolas", 9)
PAD              = 14
PAD_SM           = 7


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _resolve_settings_file():
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        for candidate in [
            os.path.join(exe_dir, "settings.json"),
            os.path.join(os.path.dirname(exe_dir), "settings.json"),
        ]:
            if os.path.isfile(candidate):
                return candidate
        return os.path.join(exe_dir, "settings.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

SETTINGS_FILE = _resolve_settings_file()


def _resolve_icon_path():
    """Find app_icon.ico for both the frozen exe and dev mode."""
    if getattr(sys, "frozen", False):
        # PyInstaller --onefile: _MEIPASS is temp extraction dir
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            p = os.path.join(meipass, "app_icon.ico")
            if os.path.isfile(p):
                return p
        # Fallback: beside exe
        p = os.path.join(os.path.dirname(sys.executable), "app_icon.ico")
        if os.path.isfile(p):
            return p
    else:
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
        if os.path.isfile(p):
            return p
    return None


def _keyring_available():
    try:
        import keyring; return True
    except ImportError:
        return False


def load_api_key():
    if _keyring_available():
        try:
            import keyring
            key = keyring.get_password(KEYRING_SERVICE, KEYRING_USER)
            if key: return key
        except Exception: pass
    try:
        if os.path.isfile(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("api_key", "")
    except Exception: pass
    return ""


def save_api_key(api_key):
    if _keyring_available():
        try:
            import keyring
            if api_key:
                keyring.set_password(KEYRING_SERVICE, KEYRING_USER, api_key)
            else:
                try: keyring.delete_password(KEYRING_SERVICE, KEYRING_USER)
                except Exception: pass
            return True
        except Exception: pass
    return False


def load_settings():
    defaults = {
        "ai_vision_enabled": False,
        "provider": "gemini",
        "api_key": "",
        "model": "gemini-3.6-flash",
        "auto_save_md": False,
        "compress_images": True,
        "ocr_engine": "offline",   # "offline" | "ai" | "off"; see ocr_mode()
        "ocr_language": "id",
        "base_url": "",
        "pdf_compress_enabled": False,
        "pdf_compress_preset": "standard",
        "ppt_compress_enabled": False,
        "ppt_compress_preset": "standard",
        "word_compress_enabled": False,
        "word_compress_preset": "standard",
        "merger_compress_enabled": False,
        "merger_compress_preset": "standard",
        "optimizer_preset": "standard",
    }
    loaded = {}
    if os.path.isfile(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    defaults.update(loaded)
        except Exception:
            pass

    # Backward-compatible fallback: if tab-specific keys are absent, inherit from pdf_compress_enabled/preset
    if "pdf_compress_enabled" in loaded:
        base_en = bool(loaded["pdf_compress_enabled"])
        if "ppt_compress_enabled" not in loaded:
            defaults["ppt_compress_enabled"] = base_en
        if "word_compress_enabled" not in loaded:
            defaults["word_compress_enabled"] = base_en
        if "merger_compress_enabled" not in loaded:
            defaults["merger_compress_enabled"] = base_en
    if "pdf_compress_preset" in loaded:
        base_pre = str(loaded["pdf_compress_preset"])
        if "ppt_compress_preset" not in loaded:
            defaults["ppt_compress_preset"] = base_pre
        if "word_compress_preset" not in loaded:
            defaults["word_compress_preset"] = base_pre
        if "merger_compress_preset" not in loaded:
            defaults["merger_compress_preset"] = base_pre
        if "optimizer_preset" not in loaded:
            defaults["optimizer_preset"] = base_pre

    key = load_api_key()
    if key:
        defaults["api_key"] = key
    if not str(defaults.get("model") or "").strip():
        defaults["model"] = GEMINI_MODELS[0]
    return defaults


def ocr_mode(settings) -> str:
    """'ai', 'offline' or 'off'. AI needs a key; any other/legacy engine value means offline."""
    if settings.get("ai_vision_enabled") and str(settings.get("api_key", "")).strip():
        return "ai"
    return "off" if settings.get("ocr_engine") == "off" else "offline"


def save_settings(settings, show_error_parent=None):
    in_keyring = save_api_key(settings.get("api_key", ""))
    # Without a working keyring, keep the key in settings.json rather than losing it on restart.
    safe = settings if not in_keyring else {k: v for k, v in settings.items() if k != "api_key"}
    try:
        target_path = Path(SETTINGS_FILE)
        target_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_file = target_path.with_name(f"{target_path.name}.tmp")
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(safe, f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        tmp_file.replace(target_path)
        return True
    except Exception as e:
        if show_error_parent:
            messagebox.showwarning(
                "Could Not Save Settings",
                f"Settings could not be saved:\n{e}\n\nCheck that the app folder is writable.",
                parent=show_error_parent,
            )
        return False


def estimate_tokens(text: str) -> int:
    return max(1, round(len(text) / 4))

def format_tokens(n: int) -> str:
    if n >= 1_000_000: return f"~{n/1_000_000:.1f}M tokens"
    if n >= 1_000:     return f"~{n/1_000:.1f}K tokens"
    return f"~{n} tokens"

def token_context_warning(n: int):
    if n > 1_000_000: return "Larger than a 1M-token context window"
    if n > 200_000:   return "Larger than a 200K-token context window"
    if n > 128_000:   return "Larger than a 128K-token context window"
    return None


def flat_btn(parent, text, command, bg=C_PRIMARY, fg="white",
             active_bg=C_PRIMARY_H, font=FONT_BTN,
             padx=16, pady=6, width=None):
    cfg = dict(text=text, command=command, bg=bg, fg=fg,
               activebackground=active_bg, activeforeground=fg,
               font=font, relief="flat", bd=0, cursor="hand2",
               padx=padx, pady=pady)
    if width: cfg["width"] = width
    btn = tk.Button(parent, **cfg)
    btn.bind("<Enter>", lambda e: btn.config(bg=active_bg))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    return btn


def open_folder_and_select(file_path: Path):
    """Open the folder in Windows Explorer with the file selected."""
    try:
        p = Path(file_path)
        if p.exists():
            subprocess.Popen(["explorer", f"/select,{p.resolve()}"])
        elif p.parent.exists():
            os.startfile(str(p.parent))
    except Exception as e:
        logging.getLogger("MarkItDown").warning("Could not open Explorer: %s", e)


# ══════════════════════════════════════════════════════════════════════════════
# SETTINGS DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, settings, on_saved):
        super().__init__(parent)
        self.title("Text Recognition Settings")
        self.geometry("520x790")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self.transient(parent)
        self.grab_set()
        self._s = dict(settings)
        self._on_saved = on_saved
        self._build()

    def _build(self):
        p = PAD
        tk.Label(self, text="Text Recognition (OCR)",
                 font=FONT_TITLE_LARGE, fg=C_TEXT, bg=C_BG
                 ).pack(anchor="w", padx=p, pady=(p, 2))
        tk.Label(self,
                 text="How to read scanned pages and images that have no text layer.",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG, justify="left"
                 ).pack(anchor="w", padx=p, pady=(0, 8))

        self._var_mode = tk.StringVar(value=ocr_mode(self._s))
        for val, lbl, hint in [
            ("offline", "Offline OCR (recommended)",
             "Runs on this computer. Free, private, no API key. Reads printed text."),
            ("ai", "AI Vision",
             "Sends images to an AI model. Better with handwriting, tables and charts."),
            ("off", "Off", "Only extract text that is already in the file."),
        ]:
            ttk.Radiobutton(self, text=lbl, variable=self._var_mode, value=val,
                            command=self._on_mode).pack(anchor="w", padx=p)
            tk.Label(self, text=hint, font=FONT_CAPTION, fg=C_SUBTEXT, bg=C_BG
                     ).pack(anchor="w", padx=(p + 28, p), pady=(0, 4))

        # AI Vision options (enabled only when AI Vision is chosen)
        frm = ttk.LabelFrame(self, text=" AI Vision provider ", padding=(10, 6))
        frm.pack(fill="x", padx=p, pady=(6, 10))
        self._var_prov = tk.StringVar(value=self._s.get("provider", "gemini"))
        for val, lbl in [("gemini", "Google Gemini  (has a free tier)"),
                          ("openai", "OpenAI")]:
            ttk.Radiobutton(frm, text=lbl, variable=self._var_prov, value=val,
                            command=self._on_prov).pack(anchor="w", pady=2)

        # Model (editable: the list is only suggestions)
        row_m = tk.Frame(self, bg=C_BG)
        row_m.pack(fill="x", padx=p, pady=(0, 10))
        tk.Label(row_m, text="Model:", font=FONT_BODY, fg=C_TEXT, bg=C_BG
                 ).pack(side="left", padx=(0, 8))
        self._var_model = tk.StringVar(value=self._s.get("model", GEMINI_MODELS[0]))
        self._cmb = ttk.Combobox(row_m, textvariable=self._var_model,
                                  font=FONT_BODY, width=28)
        self._cmb.pack(side="left")
        tk.Label(row_m, text="pick or type", font=FONT_CAPTION,
                 fg=C_SUBTEXT, bg=C_BG).pack(side="left", padx=(6, 0))

        # API key
        tk.Label(self, text="API key:", font=FONT_BODY, fg=C_TEXT, bg=C_BG
                 ).pack(anchor="w", padx=p, pady=(0, 4))
        kf = tk.Frame(self, bg=C_BG)
        kf.pack(fill="x", padx=p, pady=(0, 4))
        self._ent = ttk.Entry(kf, show="•")
        self._ent.pack(side="left", fill="x", expand=True)
        self._ent.insert(0, self._s.get("api_key", ""))
        self._var_show = tk.BooleanVar()
        ttk.Checkbutton(kf, text="Show", variable=self._var_show,
                        command=lambda: self._ent.config(
                            show="" if self._var_show.get() else "•")
                        ).pack(side="right", padx=(6, 0))

        self._lbl_help = tk.Label(self, text="",
            font=FONT_LINK, fg=C_PRIMARY,
            bg=C_BG, cursor="hand2")
        self._lbl_help.pack(anchor="w", padx=p, pady=(0, 10))
        self._lbl_help.bind("<Button-1>", lambda e: self._open_help())

        # Base URL: any OpenAI-compatible endpoint (OpenRouter, Ollama, LM Studio...)
        tk.Label(self, text="Base URL (optional):", font=FONT_BODY, fg=C_TEXT, bg=C_BG
                 ).pack(anchor="w", padx=p, pady=(0, 4))
        self._ent_url = ttk.Entry(self)
        self._ent_url.pack(fill="x", padx=p)
        self._ent_url.insert(0, self._s.get("base_url", ""))
        tk.Label(self,
                 text="Leave empty for the provider above. For a local model, e.g.\n"
                      "Ollama: http://localhost:11434/v1  (any API key works)",
                 font=FONT_CAPTION, fg=C_SUBTEXT, bg=C_BG, justify="left"
                 ).pack(anchor="w", padx=p, pady=(2, 10))

        tk.Frame(self, bg=C_SEP, height=1).pack(fill="x", padx=p, pady=(0, 10))

        tk.Label(self, text="Options", font=FONT_SUBTITLE,
                 fg=C_TEXT, bg=C_BG).pack(anchor="w", padx=p, pady=(0, 6))

        self._var_compress = tk.BooleanVar(value=self._s.get("compress_images", True))
        ttk.Checkbutton(self, text="Shrink images before sending to AI Vision (faster, cheaper)",
                        variable=self._var_compress
                        ).pack(anchor="w", padx=p, pady=(0, 4))

        self._var_autosave = tk.BooleanVar(value=self._s.get("auto_save_md", False))
        ttk.Checkbutton(self, text="Auto-save the .md result next to the source file",
                        variable=self._var_autosave
                        ).pack(anchor="w", padx=p, pady=(0, 10))

        bf = tk.Frame(self, bg=C_BG)
        bf.pack(side="bottom", fill="x", padx=p, pady=(0, p))
        flat_btn(bf, "Save", self._save, padx=18, pady=7).pack(side="right")
        flat_btn(bf, "Cancel", self.destroy,
                 bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                 padx=14, pady=7).pack(side="right", padx=(0, 8))

        self._ai_widgets = [w for w in (frm, *frm.winfo_children(), self._cmb, self._ent,
                                        self._ent_url, *kf.winfo_children())
                            if isinstance(w, ttk.Widget) and not isinstance(w, ttk.LabelFrame)]
        self._on_prov(initial=True)
        self._on_mode()

    def _on_mode(self):
        state = ["!disabled"] if self._var_mode.get() == "ai" else ["disabled"]
        for w in self._ai_widgets:
            w.state(state)

    def _on_prov(self, initial=False):
        models = GEMINI_MODELS if self._var_prov.get() == "gemini" else OPENAI_MODELS
        self._cmb["values"] = models
        # On a provider switch, swap in that provider's default; keep a typed model on open.
        if not initial or not self._var_model.get().strip():
            self._var_model.set(models[0])
        if self._var_prov.get() == "gemini":
            self._lbl_help.config(text="Get a free Google Gemini API key")
        else:
            self._lbl_help.config(text="Open the OpenAI API keys page")

    def _open_help(self):
        url = ("https://aistudio.google.com/apikey"
               if self._var_prov.get() == "gemini"
               else "https://platform.openai.com/api-keys")
        webbrowser.open(url)

    def _save(self):
        key = self._ent.get().strip()
        model = self._var_model.get().strip()
        base_url = self._ent_url.get().strip()
        use_ai = self._var_mode.get() == "ai"
        if use_ai and not key:
            messagebox.showwarning("API Key Missing",
                "Enter an API key, or choose Offline OCR.",
                parent=self)
            return
        if use_ai and not model:
            messagebox.showwarning("Model Missing", "Pick or type a model name.", parent=self)
            return
        if base_url and not base_url.lower().startswith(("http://", "https://")):
            messagebox.showwarning("Invalid Base URL",
                "The base URL must start with http:// or https://", parent=self)
            return
        self._s.update({
            "ai_vision_enabled": use_ai,
            "ocr_engine": self._var_mode.get(),
            "provider": self._var_prov.get(),
            "api_key":  key,
            "model":    model,
            "base_url": base_url,
            "compress_images": self._var_compress.get(),
            "auto_save_md":    self._var_autosave.get(),
        })
        save_settings(self._s, show_error_parent=self)
        self._on_saved(self._s)
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 - CONVERT TO TEXT
# ══════════════════════════════════════════════════════════════════════════════

def create_markitdown_instance(settings=None, cancel_event=None, on_progress=None):
    """Return (MarkItDown | None, error | None, ai_used)."""
    from markitdown import MarkItDown
    use_ai = (settings and settings.get("ai_vision_enabled", False)
              and bool(settings.get("api_key", "").strip()))
    if use_ai:
        api_key  = settings["api_key"].strip()
        provider = settings.get("provider", "gemini")
        base_url = (settings.get("base_url") or "").strip() or (
            GEMINI_BASE_URL if provider == "gemini" else None)
        default_model = GEMINI_MODELS[0] if provider == "gemini" else OPENAI_MODELS[0]
        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key, base_url=base_url, timeout=90.0, max_retries=2)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                md = MarkItDown(
                    enable_plugins=True,
                    llm_client=client,
                    llm_model=(settings.get("model") or "").strip() or default_model,
                    llm_prompt=AI_VISION_PROMPT,
                    compress_images=settings.get("compress_images", True),
                    ocr_cancel_event=cancel_event,
                    ocr_progress_callback=on_progress,
                )
                return md, None, True
        except Exception as e:
            return None, f"Could not set up AI Vision: {e}", False
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            plain = MarkItDown()
    except Exception as e:
        return None, str(e), False
    if ocr_mode(settings or {}) == "offline":
        try:
            from markitdown_ocr import RapidOCRService, rapidocr_available
            if rapidocr_available():
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    ocr = MarkItDown(enable_plugins=True,
                                     ocr_service=RapidOCRService(cancel_event, on_progress))
                return _OfflineOcrConverter(plain, ocr), None, False
        except Exception as e:
            logger.warning("Offline OCR unavailable, converting without it: %s", e)
    return plain, None, False


OCR_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


def needs_ocr(path) -> bool:
    """True for image files and for PDFs with at least one page that has no text layer."""
    ext = os.path.splitext(str(path))[1].lower()
    if ext in OCR_IMAGE_EXTS:
        return True
    if ext != ".pdf":
        return False
    try:
        import pymupdf
        with pymupdf.open(str(path)) as doc:
            if doc.needs_pass:
                return False
            return any(len(page.get_text().strip()) < 10 for page in doc)
    except Exception:
        return False


class _OfflineOcrConverter:
    """Uses the built-in converters (better tables) unless the file needs OCR."""

    def __init__(self, plain, ocr):
        self.plain, self.ocr = plain, ocr

    def convert(self, path, **kwargs):
        return (self.ocr if needs_ocr(path) else self.plain).convert(path, **kwargs)


def _find_cause(exc, *class_names):
    """Find an exception by class name in exc, its __cause__ chain, or MarkItDown's failed attempts."""
    seen, stack = set(), [exc]
    while stack:
        e = stack.pop()
        if e is None or id(e) in seen:
            continue
        seen.add(id(e))
        if any(c.__name__ in class_names for c in type(e).__mro__):
            return e
        stack += [e.__cause__, e.__context__]
        for att in getattr(e, "attempts", None) or []:
            if getattr(att, "exc_info", None):
                stack.append(att.exc_info[1])
    return None


class MarkItDownTab(tk.Frame):
    def __init__(self, parent, settings_ref: dict, on_settings_change,
                 statusbar_update):
        super().__init__(parent, bg=C_BG)
        self._settings        = settings_ref
        self._on_settings_chg = on_settings_change
        self._update_status   = statusbar_update
        self._cancelled       = False
        self._current_job_id  = 0
        self._worker          = None
        self._cancel_event    = None
        self._file_queue: list = []
        self._current_idx     = 0
        self._dnd_available   = False
        self._container = tk.Frame(self, bg=C_BG)
        self._container.pack(fill="both", expand=True)
        self._current_screen = None
        self._show_drop_zone()

    # ── Screens ──────────────────────────────────────────────────────────────
    def _clear(self):
        if self._current_screen:
            self._current_screen.destroy()
            self._current_screen = None

    def _show_drop_zone(self):
        self._clear()
        frm = self._build_drop_zone()
        frm.pack(fill="both", expand=True)
        self._current_screen = frm

    def _build_drop_zone(self):
        outer = tk.Frame(self._container, bg=C_BG)

        # ── Slim toolbar ──
        bar = tk.Frame(outer, bg=C_BG)
        bar.pack(fill="x", padx=PAD, pady=(PAD_SM, 0))

        # Short description on the left
        tk.Label(bar, text="Turn documents into text",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG).pack(side="left")

        # AI settings: small link on the right
        mode = ocr_mode(self._settings)
        ai_indicator = {"ai": "● OCR: AI Vision", "offline": "● OCR: Offline",
                        "off": "○ OCR: Off"}[mode]
        ai_color = C_SUBTEXT if mode == "off" else C_SUCCESS
        lbl_ai = tk.Label(bar, text=ai_indicator,
                          font=FONT_SMALL, fg=ai_color, bg=C_BG, cursor="hand2")
        lbl_ai.pack(side="right")
        lbl_ai.bind("<Button-1>", lambda e: self._open_settings())

        lbl_cfg = tk.Label(bar, text="Settings  ",
                           font=FONT_LINK,
                           fg=C_PRIMARY, bg=C_BG, cursor="hand2")
        lbl_cfg.pack(side="right")
        lbl_cfg.bind("<Button-1>", lambda e: self._open_settings())

        # ── Drop zone ──
        zone_wrap = tk.Frame(outer, bg=C_BG)
        zone_wrap.pack(fill="both", expand=True, padx=PAD, pady=(PAD_SM, 0))

        # White panel
        panel = tk.Frame(zone_wrap, bg=C_WHITE, relief="flat")
        panel.pack(fill="both", expand=True)

        # Canvas drop
        self._canvas = tk.Canvas(panel, bg=C_WHITE, highlightthickness=1,
                                  highlightbackground=C_BORDER, cursor="hand2")
        self._canvas.pack(fill="both", expand=True, padx=1, pady=1)
        self._setup_dnd(outer)
        outer.after(10, self._draw_normal)
        self._canvas.bind("<Button-1>", lambda e: self._pick_single())
        self._canvas.bind("<Enter>",    lambda e: self._draw_hover())
        self._canvas.bind("<Leave>",    lambda e: self._draw_normal())
        self._canvas.bind("<Configure>",lambda e: self._redraw())
        self._canvas_hovered = False

        # ── Format label ──
        tk.Label(outer,
                 text="PDF  ·  Word  ·  Excel  ·  PowerPoint  ·  HTML  ·  Images",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG).pack(pady=(PAD_SM, 6))

        # ── Buttons ──
        br = tk.Frame(outer, bg=C_BG)
        br.pack(pady=(0, PAD_SM))
        flat_btn(br, "Choose File", self._pick_single,
                 padx=16, pady=6).pack(side="left", padx=(0, 8))
        flat_btn(br, "Choose Multiple Files", self._pick_multi,
                 bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                 font=FONT_BODY, padx=12, pady=6).pack(side="left")

        # Queue banner
        remaining = self._file_queue[self._current_idx:]
        if remaining:
            bann = tk.Frame(outer, bg="#E3F2FD", pady=4)
            bann.pack(fill="x", padx=PAD, pady=(0, 6))
            tk.Label(bann, text=f"  {len(remaining)} file(s) in queue",
                     font=FONT_SMALL, fg="#1565C0", bg="#E3F2FD").pack(side="left")
            lbl_x = tk.Label(bann, text="Clear  ",
                             font=FONT_LINK,
                             fg=C_SUBTEXT, bg="#E3F2FD", cursor="hand2")
            lbl_x.pack(side="right")
            lbl_x.bind("<Button-1>", lambda e: self._clear_queue())

        return outer

    def _redraw(self):
        if self._canvas_hovered:
            self._draw_hover()
        else:
            self._draw_normal()

    def _draw_zone(self, border, arrow_col, text_col):
        c = self._canvas
        c.delete("all")
        w = c.winfo_width()  or 400
        h = c.winfo_height() or 150
        cx, cy = w//2, h//2 - 14
        r = 24
        # Upload arrow circle
        c.create_oval(cx-r, cy-r, cx+r, cy+r,
                      outline=border, width=1, fill=C_WHITE)
        c.create_line(cx, cy+9, cx, cy-9, fill=arrow_col, width=2)
        c.create_line(cx-7, cy-2, cx, cy-11, fill=arrow_col, width=2)
        c.create_line(cx+7, cy-2, cx, cy-11, fill=arrow_col, width=2)
        # Text
        main = "Drop files here" if self._dnd_available else "Click to choose a file"
        c.create_text(cx, cy+r+16, text=main,
                      font=FONT_SUBTITLE, fill=text_col)
        c.create_text(cx, cy+r+32, text="One file or many at once",
                      font=FONT_SMALL, fill=C_SUBTEXT)
        c.configure(highlightbackground=border)

    def _draw_normal(self):
        self._canvas_hovered = False
        self._draw_zone(C_BORDER, C_SUBTEXT, C_SUBTEXT)

    def _draw_hover(self):
        self._canvas_hovered = True
        self._draw_zone(C_PRIMARY, C_PRIMARY, C_TEXT)

    def _setup_dnd(self, outer):
        try:
            outer.drop_target_register("DND_Files")
            outer.dnd_bind("<<Drop>>", self._on_drop)
            self._canvas.drop_target_register("DND_Files")
            self._canvas.dnd_bind("<<Drop>>", self._on_drop)
            self._canvas.dnd_bind("<<DragEnter>>", lambda e: self._draw_hover())
            self._canvas.dnd_bind("<<DragLeave>>", lambda e: self._draw_normal())
            self._dnd_available = True
        except Exception:
            self._dnd_available = False

    def _on_drop(self, event):
        tokens = re.findall(r'\{[^}]+\}|\S+', event.data.strip())
        paths = [t.strip("{}").strip() for t in tokens if os.path.isfile(t.strip("{}"))]
        if paths: self._enqueue(paths)

    def _pick_single(self):
        p = filedialog.askopenfilename(
            title="Choose a document",
            filetypes=[("All documents",
                        "*.pdf *.docx *.doc *.xlsx *.xls *.pptx *.ppt "
                        "*.html *.htm *.txt *.csv *.json *.xml *.zip "
                        "*.jpg *.jpeg *.png *.mp3 *.wav"),
                       ("All files", "*.*")])
        if p: self._enqueue([p])

    def _pick_multi(self):
        paths = filedialog.askopenfilenames(
            title="Choose files",
            filetypes=[("All documents",
                        "*.pdf *.docx *.doc *.xlsx *.xls *.pptx *.ppt "
                        "*.html *.htm *.txt *.csv *.json *.xml *.zip "
                        "*.jpg *.jpeg *.png *.mp3 *.wav"),
                       ("All files", "*.*")])
        if paths: self._enqueue(list(paths))

    def _clear_queue(self):
        self._current_job_id += 1
        self._cancelled = True
        if self._cancel_event:
            self._cancel_event.set()  # stops further AI Vision calls in the worker
        self._file_queue = []
        self._current_idx = 0
        self._show_drop_zone()

    def _enqueue(self, paths):
        self._file_queue = list(paths); self._current_idx = 0
        self._process_next()

    def _process_next(self):
        if self._current_idx >= len(self._file_queue):
            self._file_queue = []; self._current_idx = 0
            self._show_drop_zone(); return
        self._convert(self._file_queue[self._current_idx])

    def _open_settings(self):
        def on_saved(ns):
            self._settings.update(ns)
            self._on_settings_chg(ns)
            self._show_drop_zone()
            self._update_status()
        SettingsDialog(self.winfo_toplevel(), self._settings, on_saved)

    # ── Loading ───────────────────────────────────────────────────────────────
    def _show_loading(self, filename, ai_active=False):
        self._clear()
        total = len(self._file_queue); idx = self._current_idx + 1
        self._cancelled = False

        outer = tk.Frame(self._container, bg=C_BG)
        panel = tk.Frame(outer, bg=C_WHITE, highlightthickness=1, highlightbackground=C_BORDER)
        panel.pack(fill="both", expand=True, padx=PAD, pady=PAD)
        inner = tk.Frame(panel, bg=C_WHITE)
        inner.place(relx=0.5, rely=0.42, anchor="center")

        title = "Converting..." if total <= 1 else f"Converting file {idx} of {total}..."
        tk.Label(inner, text=title, font=FONT_TITLE, fg=C_TEXT, bg=C_WHITE).pack(pady=(0, 4))
        tk.Label(inner, text=os.path.basename(filename), font=FONT_SMALL, fg=C_SUBTEXT,
                 bg=C_WHITE, wraplength=420).pack(pady=(0, 14))

        pb = ttk.Progressbar(inner, mode="indeterminate", length=320)
        pb.pack(pady=(0, 10))
        pb.start(12)

        self._var_progress = tk.StringVar(
            value="AI Vision is on: scanned pages can take 10-60 s each." if ai_active else "")
        tk.Label(inner, textvariable=self._var_progress, font=FONT_SMALL,
                 fg=C_SUBTEXT, bg=C_WHITE).pack(pady=(0, 14))

        flat_btn(inner, "Cancel", self._clear_queue,
                 bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                 font=FONT_BODY, padx=14, pady=6).pack()

        outer.pack(fill="both", expand=True)
        self._current_screen = outer

    # ── Result ────────────────────────────────────────────────────────────────
    def _show_result(self, filename, text, ai_used=False):
        self._clear()
        total = len(self._file_queue); idx = self._current_idx + 1
        tokens = estimate_tokens(text); warn = token_context_warning(tokens)
        has_next = (self._current_idx + 1) < total

        outer = tk.Frame(self._container, bg=C_BG)

        # Slim header
        hdr = tk.Frame(outer, bg=C_BG)
        hdr.pack(fill="x", padx=PAD, pady=(PAD_SM, 4))
        tk.Label(hdr, text="Done", font=FONT_SUBTITLE,
                 fg=C_SUCCESS, bg=C_BG).pack(side="left")
        if total > 1:
            tk.Label(hdr, text=f"  {idx}/{total}",
                     font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG).pack(side="left")
        tk.Label(hdr, text=f"  {os.path.basename(filename)[:40]}",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG).pack(side="left")

        # Character + token info
        info = tk.Frame(outer, bg=C_BG)
        info.pack(fill="x", padx=PAD, pady=(0, 4))
        info_txt = f"{len(text):,} characters  ·  {format_tokens(tokens)}"
        tk.Label(info, text=info_txt, font=FONT_SMALL,
                 fg=C_SUBTEXT, bg=C_BG).pack(side="left")
        if warn:
            tk.Label(info, text=f"  ⚠ {warn}", font=FONT_SMALL,
                     fg=C_DANGER, bg=C_BG).pack(side="left")

        # Search bar (hidden until Ctrl+F)
        srch_frame = tk.Frame(outer, bg=C_SURFACE, pady=4)
        srch_var   = tk.StringVar()
        srch_info  = tk.StringVar(value="")
        si = tk.Frame(srch_frame, bg=C_SURFACE)
        si.pack(fill="x", padx=PAD)
        tk.Label(si, text="Find:", font=FONT_SMALL,
                 fg=C_TEXT, bg=C_SURFACE).pack(side="left")
        srch_ent = tk.Entry(si, textvariable=srch_var,
                            font=FONT_MONO, width=26, relief="solid", bd=1)
        srch_ent.pack(side="left", padx=(4, 8), ipady=3)
        tk.Label(si, textvariable=srch_info, font=FONT_SMALL,
                 fg=C_SUBTEXT, bg=C_SURFACE).pack(side="left")
        lbl_x = tk.Label(si, text="Close", font=FONT_LINK,
                         fg=C_SUBTEXT, bg=C_SURFACE, cursor="hand2")
        lbl_x.pack(side="right")

        # Preview
        prev_wrap = tk.Frame(outer, bg=C_BORDER, bd=1)
        prev_wrap.pack(fill="both", expand=True, padx=PAD, pady=(0, 0))
        preview = tk.Text(prev_wrap, font=FONT_MONO, wrap="word",
                          bg=C_WHITE, fg=C_TEXT, relief="flat",
                          padx=10, pady=10, state="disabled", cursor="arrow")
        sb = ttk.Scrollbar(prev_wrap, command=preview.yview)
        preview.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        preview.pack(fill="both", expand=True)
        preview.configure(state="normal")
        preview.insert("1.0", text)
        preview.configure(state="disabled")

        # Search logic
        def do_search(*_):
            term = srch_var.get()
            preview.tag_remove("hl", "1.0", "end")
            preview.tag_configure("hl", background="#FFF176")
            if not term: srch_info.set(""); return
            start, count, first = "1.0", 0, None
            while True:
                pos = preview.search(term, start, stopindex="end", nocase=True)
                if not pos: break
                end_ = f"{pos}+{len(term)}c"
                preview.tag_add("hl", pos, end_)
                if not first: first = pos
                start = end_; count += 1
            if first: preview.see(first)
            srch_info.set(f"{count} match(es)" if count else "Not found")

        srch_var.trace_add("write", do_search)

        def show_srch(e=None):
            srch_frame.pack(fill="x", before=prev_wrap)
            srch_ent.focus_set(); srch_ent.select_range(0, "end")

        def hide_srch():
            srch_frame.pack_forget()
            preview.tag_remove("hl", "1.0", "end")
            srch_info.set("")

        lbl_x.bind("<Button-1>", lambda e: hide_srch())
        preview.bind("<Control-f>", show_srch)
        outer.bind("<Control-f>", show_srch)

        # Bottom buttons
        br = tk.Frame(outer, bg=C_BG)
        br.pack(fill="x", padx=PAD, pady=PAD_SM)

        copy_btn = [None]
        def do_copy():
            outer.clipboard_clear(); outer.clipboard_append(text); outer.update()
            copy_btn[0].config(text="Copied!", bg=C_SUCCESS,
                               activebackground=C_SUCCESS)
            outer.after(2000, lambda: copy_btn[0].config(
                text="Copy to Clipboard", bg=C_PRIMARY, activebackground=C_PRIMARY_H))

        copy_btn[0] = flat_btn(br, "Copy to Clipboard", do_copy,
                                font=FONT_BTN, padx=18, pady=7)
        copy_btn[0].pack(side="left")

        if has_next:
            def next_file():
                self._current_idx += 1; self._process_next()
            flat_btn(br, "Next File", next_file,
                     bg=C_SUCCESS, fg="white", active_bg=C_SUCCESS,
                     font=FONT_BODY, padx=14, pady=7
                     ).pack(side="left", padx=(8, 0))

        sec = tk.Frame(br, bg=C_BG)
        sec.pack(side="right")
        for ext in [".md", ".txt"]:
            def save(e=ext):
                fts = [("Markdown", "*.md")] if e == ".md" else [("Text", "*.txt")]
                p = filedialog.asksaveasfilename(defaultextension=e,
                    filetypes=fts + [("All files", "*.*")])
                if p:
                    try:
                        with open(p, "w", encoding="utf-8") as f: f.write(text)
                    except Exception as ex:
                        messagebox.showerror(
                            "Could Not Save",
                            f"The file could not be saved:\n{ex}\n\n"
                            "Check that the folder is writable and the disk has free space.",
                        )
            flat_btn(sec, f"Save {ext}", save,
                     bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                     padx=10, pady=7).pack(side="left", padx=(0, 4))

        if self._settings.get("auto_save_md"):
            try:
                with open(os.path.splitext(filename)[0]+".md", "w", encoding="utf-8") as f:
                    f.write(text)
            except Exception: pass

        again = tk.Label(outer, text="← Convert another file",
                        font=FONT_SMALL, fg=C_PRIMARY, bg=C_BG, cursor="hand2")
        again.pack(pady=(2, 8))
        again.bind("<Button-1>", lambda e: self._reset_all())
        again.bind("<Enter>",    lambda e: again.config(fg=C_PRIMARY_H))
        again.bind("<Leave>",    lambda e: again.config(fg=C_PRIMARY))

        outer.pack(fill="both", expand=True)
        self._current_screen = outer

    # ── Error ─────────────────────────────────────────────────────────────────
    def _show_error(self, filename, message):
        self._clear()
        outer = tk.Frame(self._container, bg=C_BG)
        box = tk.Frame(outer, bg=C_WHITE, highlightthickness=1, highlightbackground=C_BORDER)
        box.pack(fill="both", expand=True, padx=PAD, pady=PAD)
        panel = tk.Frame(box, bg=C_WHITE)
        panel.place(relx=0.5, rely=0.42, anchor="center")
        tk.Label(panel, text="Could not convert this file",
                 font=FONT_TITLE, fg=C_DANGER, bg=C_WHITE
                 ).pack(pady=(0, 4))
        tk.Label(panel, text=os.path.basename(filename),
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_WHITE).pack(pady=(0, 10))
        frm_err = tk.Frame(panel, bg="#FDECEA", padx=12, pady=10)
        frm_err.pack(fill="x", padx=PAD)
        tk.Label(frm_err, text=message, font=FONT_BODY, fg=C_DANGER,
                 bg="#FDECEA", wraplength=480, justify="left").pack(anchor="w")
        tk.Label(panel,
                 text="Make sure the file is not open in another app.",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_WHITE
                 ).pack(pady=(10, 14))
        br = tk.Frame(panel, bg=C_WHITE)
        br.pack()
        flat_btn(br, "Try Another File", self._reset_all,
                 padx=16, pady=6).pack(side="left", padx=(0, 8))
        def copy_err():
            outer.clipboard_clear(); outer.clipboard_append(message); outer.update()
            messagebox.showinfo("Copied", "Error details copied to the clipboard.", parent=outer)
        flat_btn(br, "Copy Error", copy_err,
                 bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                 font=FONT_BODY, padx=12, pady=6).pack(side="left")
        outer.pack(fill="both", expand=True)
        self._current_screen = outer

    def _reset_all(self):
        self._file_queue = []; self._current_idx = 0
        self._show_drop_zone()

    # ── Worker ────────────────────────────────────────────────────────────────
    def _convert(self, filepath, offline=False):
        try: mb = os.path.getsize(filepath) / (1024*1024)
        except OSError: mb = 0
        if mb > FILE_LIMIT_MB:
            messagebox.showerror("File Too Large",
                f"This file is {mb:.0f} MB. The limit is {FILE_LIMIT_MB} MB."); return
        if not offline and mb > FILE_WARN_MB:
            if not messagebox.askyesno("Large File",
                f"This file is {mb:.0f} MB and may take a while. Continue?"): return
        settings = {"ai_vision_enabled": False} if offline else dict(self._settings)
        ai = bool(settings.get("ai_vision_enabled") and settings.get("api_key", "").strip())
        self._current_job_id += 1
        job_id = self._current_job_id
        cancel_event = threading.Event()
        self._cancel_event = cancel_event
        self._show_loading(filepath, ai_active=ai)

        def ui(fn):
            # Drop results of jobs that were cancelled or superseded.
            self.after(0, lambda: fn() if self._current_job_id == job_id and not self._cancelled else None)

        def worker():
            def on_image(n):
                ui(lambda: self._var_progress.set(
                    f"{'AI Vision' if ai else 'Offline OCR'}: reading image {n}..."))
            md_i, err, ai_used = create_markitdown_instance(settings, cancel_event, on_image)
            if not md_i:
                ui(lambda: self._show_error(filepath, f"Could not start the converter:\n\n{err}"))
                return
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    result = md_i.convert(filepath)
                text = (result.markdown or "").strip() or (
                    "(No text found. If this is a scan, choose Offline OCR in Settings.)"
                    if ocr_mode(settings) == "off" else "(No text found.)")
                ui(lambda: self._show_result(filepath, text, ai_used))
            except PermissionError:
                ui(lambda: self._show_error(filepath,
                    "The file is open in another app. Close it and try again."))
            except Exception as exc:
                if _find_cause(exc, "OCRCancelledError"):
                    return
                s = str(exc)
                if ai_used and _find_cause(exc, "OCRQuotaExceededError", "RateLimitError"):
                    ui(lambda: self._ask_offline_fallback(filepath,
                        "AI Vision quota or rate limit reached (HTTP 429)."))
                    return
                auth = _find_cause(exc, "OCRAuthError") if ai_used else None
                if auth:
                    ui(lambda: self._ask_offline_fallback(filepath,
                        "AI Vision rejected the request. Check the API key, model name "
                        f"and base URL in Settings.\n\n{str(auth)[:300]}"))
                    return
                if any(k in s.lower() for k in ("password", "encrypt")):
                    msg = "The file is password-protected. Remove the password and try again."
                else:
                    msg = f"Something went wrong:\n\n{s[:400]}"
                ui(lambda: self._show_error(filepath, msg))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _ask_offline_fallback(self, filepath, reason):
        if messagebox.askyesno("AI Vision Unavailable",
                               f"{reason}\n\nConvert this file without AI Vision instead?",
                               parent=self):
            self._convert(filepath, offline=True)
        else:
            self._show_error(filepath, reason)


# ══════════════════════════════════════════════════════════════════════════════
# SHARED WIDGET: COMPRESS OPTIONS (TABS 2, 3, 4)
# ══════════════════════════════════════════════════════════════════════════════

class CompressOptions(tk.Frame):
    """
    Reusable toggle + preset picker for PDF compression (tabs 2, 3, 4).
    Saves the choice to settings.json automatically.
    """
    PRESET_DISPLAY = {
        "standard": "Standard (recommended)",
        "light":    "Light (images untouched)",
        "maximum":  "Maximum (smallest file)",
    }
    DISPLAY_TO_PRESET = {v: k for k, v in PRESET_DISPLAY.items()}

    def __init__(self, parent, settings_ref=None, on_settings_change=None, bg=C_BG,
                 enable_key="pdf_compress_enabled", preset_key="pdf_compress_preset"):
        super().__init__(parent, bg=bg)
        self._s = settings_ref if settings_ref is not None else {}
        self._on_change = on_settings_change
        self._enable_key = enable_key
        self._preset_key = preset_key

        init_en = self._s.get(self._enable_key, self._s.get("pdf_compress_enabled", False))
        init_preset = self._s.get(self._preset_key, self._s.get("pdf_compress_preset", "standard"))

        self._var_en = tk.BooleanVar(value=init_en)
        self._var_preset = tk.StringVar(value=init_preset)
        self._build()

    def _build(self):
        top = tk.Frame(self, bg=self["bg"])
        top.pack(fill="x")

        self._chk = ttk.Checkbutton(
            top, text="Compress the PDF",
            variable=self._var_en, command=self._on_toggle)
        self._chk.pack(side="left")

        self._frm_preset = tk.Frame(self, bg=self["bg"])
        tk.Label(self._frm_preset, text="Preset:",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=self["bg"]).pack(side="left", padx=(18, 6))

        self._combo = ttk.Combobox(
            self._frm_preset, state="readonly", font=FONT_SMALL, width=28,
            values=list(self.PRESET_DISPLAY.values()))
        cur_disp = self.PRESET_DISPLAY.get(self._var_preset.get(), self.PRESET_DISPLAY["standard"])
        self._combo.set(cur_disp)
        self._combo.bind("<<ComboboxSelected>>", self._on_select)
        self._combo.pack(side="left")

        if self._var_en.get():
            self._frm_preset.pack(fill="x", pady=(4, 0))

    def _on_toggle(self):
        if self._var_en.get():
            self._frm_preset.pack(fill="x", pady=(4, 0))
        else:
            self._frm_preset.pack_forget()
        self._save()

    def _on_select(self, _=None):
        disp = self._combo.get()
        key = self.DISPLAY_TO_PRESET.get(disp, "standard")
        self._var_preset.set(key)
        self._save()

    def _save(self):
        self._s[self._enable_key] = self._var_en.get()
        self._s[self._preset_key] = self._var_preset.get()
        if self._enable_key == "pdf_compress_enabled":
            self._s["pdf_compress_enabled"] = self._var_en.get()
        if self._preset_key == "pdf_compress_preset":
            self._s["pdf_compress_preset"] = self._var_preset.get()
        save_settings(self._s)
        if self._on_change:
            try: self._on_change(self._s)
            except Exception: pass

    def refresh_from_settings(self, settings=None):
        if settings is not None:
            self._s = settings
        en = self._s.get(self._enable_key, self._s.get("pdf_compress_enabled", False))
        preset = self._s.get(self._preset_key, self._s.get("pdf_compress_preset", "standard"))
        self._var_en.set(en)
        self._var_preset.set(preset)
        cur_disp = self.PRESET_DISPLAY.get(preset, self.PRESET_DISPLAY["standard"])
        if hasattr(self, "_combo"):
            self._combo.set(cur_disp)
        if en:
            if hasattr(self, "_frm_preset"):
                self._frm_preset.pack(fill="x", pady=(4, 0))
        else:
            if hasattr(self, "_frm_preset"):
                self._frm_preset.pack_forget()

    def is_enabled(self) -> bool:
        return bool(self._var_en.get())

    def get_preset_key(self) -> str:
        return self._var_preset.get()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — POWERPOINT TO PDF
# ══════════════════════════════════════════════════════════════════════════════

def is_office_encrypted(filepath) -> bool:
    path = Path(filepath)
    if not path.is_file():
        return False
    try:
        with open(path, "rb") as f:
            magic = f.read(8)
        if magic == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
            try:
                import olefile
                if olefile.isOleFile(path):
                    with olefile.OleFileIO(path) as ole:
                        return ole.exists("EncryptionInfo") or ole.exists("EncryptedPackage")
            except Exception:
                pass
            with open(path, "rb") as f:
                head = f.read(4096)
                if b"EncryptionInfo" in head or b"EncryptedPackage" in head:
                    return True
        return False
    except Exception:
        return False


def _ppt_convert_single(ppt_path, output_dir, powerpoint, log_fn):
    pdf_name = ppt_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name
    pres = None
    try:
        if ppt_path.stat().st_size == 0:
            log_fn(f"[FAIL] {ppt_path.name}: File is empty (0 bytes).")
            return False
        if is_office_encrypted(ppt_path):
            log_fn(f"[SKIP] {ppt_path.name}: File is password-protected. Skipped so PowerPoint does not hang.")
            return False
        pres = powerpoint.Presentations.Open(
            str(ppt_path.resolve()), ReadOnly=True,
            Untitled=False, WithWindow=False)
        pres.SaveAs(str(pdf_path.resolve()), FileFormat=32)
        log_fn(f"[OK]   {ppt_path.name}")
        return True
    except Exception as e:
        log_fn(f"[FAIL] {ppt_path.name}: {e}")
        if pdf_path.exists():
            try:
                pdf_path.unlink(missing_ok=True)
            except Exception:
                pass
        return False
    finally:
        if pres:
            try: pres.Close()
            except Exception: pass


def _ppt_batch_worker(files, output_dir, log_fn, progress_fn, done_fn,
                      compress_enabled=False, compress_preset="standard",
                      cancel_event=None):
    import pythoncom, win32com.client
    import pymupdf
    from pdf_optimizer import has_digital_signatures

    pythoncom.CoInitialize()
    ppt_app, was_running = None, False
    try:
        ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        was_running = True
    except Exception:
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            log_fn(f"[ERROR] Could not start PowerPoint: {e}")
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            try:
                done_fn(0, len(files), output_dir, None)
            except Exception:
                pass
            return

    try:
        ppt_app.DisplayAlerts = 1  # ppAlertsNone
    except Exception:
        pass

    output_dir.mkdir(parents=True, exist_ok=True)
    total = len(files)
    log_fn(f"[INFO] Converting {total} PowerPoint file(s) to:\n{output_dir}\n{'-'*40}")

    cancelled = False
    ok = fail = 0
    comp_ok = comp_opt = comp_fail = 0
    orig_total_size = new_total_size = 0
    t0 = time.time()
    try:
        for idx, f in enumerate(files):
            if cancel_event and cancel_event.is_set():
                log_fn("[INFO] Cancelled by user.")
                cancelled = True
                break

            file_num = idx + 1
            file_base = (idx / total) * 100.0
            file_weight = (1.0 / total) * 100.0

            # Initial conversion progress
            progress_fn(file_num, total, f"File {file_num} of {total}: converting {f.name}...", int(file_base))

            conv_ok = _ppt_convert_single(f, output_dir, ppt_app, log_fn)

            if cancel_event and cancel_event.is_set():
                log_fn("[INFO] Cancelled by user.")
                cancelled = True
                break

            if conv_ok:
                ok += 1
                pdf_path = output_dir / (f.stem + ".pdf")
                if compress_enabled and pdf_path.exists():
                    tmp_target = output_dir / f"{f.stem}__opt_tmp.pdf"
                    try:
                        comp_start_pct = int(file_base + file_weight * 0.70)
                        progress_fn(file_num, total, f"File {file_num} of {total}: compressing images...", comp_start_pct)
                        log_fn(f"[INFO] Compressing {pdf_path.name}...")
                        f_orig_sz = pdf_path.stat().st_size
                        orig_total_size += f_orig_sz

                        # Check for digital signatures
                        is_signed = False
                        try:
                            chk_doc = pymupdf.open(str(pdf_path))
                            is_signed = has_digital_signatures(chk_doc)
                            chk_doc.close()
                        except Exception:
                            pass

                        if is_signed:
                            log_fn(f"[WARN] {pdf_path.name}: Document is digitally signed. Compression skipped.")
                            comp_opt += 1
                            new_total_size += f_orig_sz
                        else:
                            def on_comp_step(stage, comp_pct, _fb=file_base, _fw=file_weight, _fn=file_num, _tot=total):
                                current_pct = int(_fb + _fw * 0.70 + (_fw * 0.30) * (comp_pct / 100.0))
                                progress_fn(_fn, _tot, f"File {_fn} of {_tot}: {stage}", current_pct)

                            res = optimize_pdf(pdf_path, dst_path=tmp_target,
                                               preset_key=compress_preset,
                                               on_progress=on_comp_step,
                                               cancel_event=cancel_event,
                                               ignore_signatures=False)
                            if res.status == "success" and res.dst_path:
                                Path(res.dst_path).replace(pdf_path)
                                comp_ok += 1
                                new_total_size += res.new_size
                                log_fn(f"[OK] {pdf_path.name}: {format_size(res.orig_size)} → "
                                       f"{format_size(res.new_size)} (saved {res.saved_percent:.1f}%)")
                            elif res.status == "already_optimal":
                                comp_opt += 1
                                new_total_size += f_orig_sz
                                log_fn(f"[INFO] {pdf_path.name}: already optimal.")
                            else:
                                comp_fail += 1
                                new_total_size += f_orig_sz
                                log_fn(f"[WARN] {pdf_path.name}: Compression skipped ({res.message}).")
                    except CancelledError:
                        log_fn("[INFO] Cancelled by user.")
                        cancelled = True
                        break
                    except Exception as ce:
                        comp_fail += 1
                        if pdf_path.exists():
                            new_total_size += pdf_path.stat().st_size
                        log_fn(f"[WARN] Could not compress {pdf_path.name}: {ce}. The converted file was kept.")
                    finally:
                        if tmp_target.exists():
                            try:
                                tmp_target.unlink(missing_ok=True)
                            except Exception:
                                pass
            else:
                fail += 1

            # Per-file done
            progress_fn(file_num, total, f"File {file_num} of {total}: done", int(file_base + file_weight))
    except Exception as err:
        log_fn(f"[ERROR] PowerPoint worker stopped unexpectedly: {err}")
    finally:
        if not was_running and ppt_app:
            try: ppt_app.Quit()
            except Exception: pass
        ppt_app = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        comp_stats = None
        if not cancelled:
            log_fn(f"[INFO] Finished in {time.time()-t0:.1f}s | OK: {ok} | Failed: {fail}")
            if compress_enabled and ok > 0:
                saved_bytes = max(0, orig_total_size - new_total_size)
                saved_pct = (saved_bytes / orig_total_size * 100) if orig_total_size > 0 else 0.0
                log_fn(f"[INFO] Compression: {comp_ok} compressed, {comp_opt} already optimal, {comp_fail} failed.")
                log_fn(f"[INFO] Total size: {format_size(orig_total_size)} → {format_size(new_total_size)} "
                       f"(saved {format_size(saved_bytes)} / {saved_pct:.1f}%)")
                comp_stats = {
                    "comp_ok": comp_ok,
                    "comp_opt": comp_opt,
                    "comp_fail": comp_fail,
                    "orig_total_size": orig_total_size,
                    "new_total_size": new_total_size,
                    "saved_bytes": saved_bytes,
                    "saved_pct": saved_pct,
                }

        try:
            done_fn(ok, fail, output_dir, comp_stats, cancelled=cancelled)
        except TypeError:
            done_fn(ok, fail, output_dir, comp_stats)


class PptToPdfTab(tk.Frame):
    def __init__(self, parent, settings_ref=None, on_settings_change=None):
        super().__init__(parent, bg=C_BG)
        self._settings = settings_ref or {}
        self._on_settings_chg = on_settings_change
        self._last_out_dir = None
        self._is_running = False
        self._cancel_event = threading.Event()
        self._build()

    def _build(self):
        frm_in = ttk.LabelFrame(self, text=" File / Folder PowerPoint ", padding=(PAD, PAD_SM))
        frm_in.pack(fill="x", padx=PAD, pady=(PAD, PAD_SM))
        self._var_in = tk.StringVar()
        ttk.Entry(frm_in, textvariable=self._var_in, font=FONT_BODY
                  ).pack(fill="x", pady=(0, 6))
        br = ttk.Frame(frm_in)
        br.pack(fill="x")
        ttk.Button(br, text="Choose Folder", command=self._choose_folder
                   ).pack(side="left", padx=(0, 8))
        ttk.Button(br, text="Choose PowerPoint Files", command=self._choose_files
                   ).pack(side="left", padx=(0, 8))
        tk.Label(br, text=".pptx · .ppt · .pps and more",
                 font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG).pack(side="left")

        frm_out = ttk.LabelFrame(self, text=" Output Folder ", padding=(PAD, PAD_SM))
        frm_out.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._var_out = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self._var_out, font=FONT_BODY
                  ).pack(fill="x", pady=(0, 6))
        ttk.Button(frm_out, text="Choose Folder...",
                   command=self._choose_output).pack(anchor="w")

        # PDF compression option
        frm_opt = ttk.LabelFrame(self, text=" PDF Compression ", padding=(PAD, PAD_SM))
        frm_opt.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._compress_opt = CompressOptions(
            frm_opt, settings_ref=self._settings,
            on_settings_change=self._on_settings_chg,
            enable_key="ppt_compress_enabled",
            preset_key="ppt_compress_preset",
        )
        self._compress_opt.pack(fill="x")
        self.compress_opts = self._compress_opt

        act = tk.Frame(self, bg=C_BG)
        act.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._btn_start = flat_btn(act, "Convert to PDF", self._start, padx=16, pady=6)
        self._btn_start.pack(side="left", padx=(0, 8))
        self._btn_cancel = flat_btn(act, "Cancel", self._cancel,
                                   bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                                   font=FONT_BODY, padx=12, pady=6)
        self._btn_cancel.pack(side="left", padx=(0, 8))
        self._btn_cancel.config(state="disabled")
        self._btn_open = flat_btn(act, "Open Output Folder", self._open_out,
                                   bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                                   font=FONT_BODY, padx=12, pady=6)
        self._btn_open.pack(side="left")
        self._btn_open.config(state="disabled")
        self._var_status = tk.StringVar(value="Ready.")
        tk.Label(act, textvariable=self._var_status,
                 font=FONT_STATUS, fg=C_SUBTEXT, bg=C_BG
                 ).pack(side="left", padx=10)

        self._pbar = ttk.Progressbar(self, mode="determinate")
        self._pbar.pack(fill="x", padx=PAD, pady=(0, PAD_SM))

        frm_log = ttk.LabelFrame(self, text=" Log ", padding=(PAD, PAD_SM))
        frm_log.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self._log = scrolledtext.ScrolledText(frm_log, font=FONT_MONO,
                                               state="disabled", wrap="word",
                                               bg=C_WHITE)
        self._log.pack(fill="both", expand=True)

    def _choose_folder(self):
        d = filedialog.askdirectory(title="Choose a folder with PowerPoint files")
        if d:
            self._var_in.set(d)
            if not self._var_out.get():
                self._var_out.set(str(Path(d) / "PDF Output"))

    def _choose_files(self):
        files = filedialog.askopenfilenames(
            title="Choose PowerPoint files",
            filetypes=[("PowerPoint", "*.pptx *.ppt *.pptm *.pps *.ppsx"),
                       ("All files", "*.*")])
        if files:
            self._var_in.set("; ".join(files))
            if not self._var_out.get():
                self._var_out.set(str(Path(files[0]).parent / "PDF Output"))

    def _choose_output(self):
        d = filedialog.askdirectory(title="Choose the output folder for PDFs")
        if d: self._var_out.set(d)

    def _open_out(self):
        if self._last_out_dir and self._last_out_dir.exists():
            os.startfile(str(self._last_out_dir))

    def _append_log(self, msg):
        def _u():
            self._log.configure(state="normal")
            self._log.insert("end", msg+"\n")
            self._log.see("end")
            self._log.configure(state="disabled")
        self.after(0, _u)

    def _update_progress(self, cur, total, name, pct=None):
        def _u():
            if pct is not None:
                self._pbar["value"] = pct
            else:
                self._pbar["value"] = int(cur / total * 100)
            self._var_status.set(name)
        self.after(0, _u)

    def _on_done(self, ok, fail, out_dir, comp_stats=None, cancelled=False):
        def _u():
            self._is_running = False
            self._btn_start.config(state="normal")
            self._btn_cancel.config(state="disabled")
            if cancelled:
                self._var_status.set("Cancelled by user.")
                return
            self._last_out_dir = out_dir
            if out_dir.exists(): self._btn_open.config(state="normal")
            self._pbar["value"] = 100
            self._var_status.set(f"Done: {ok} converted, {fail} failed.")
            if ok + fail > 0:
                msg = f"Conversion finished.\n\nConverted: {ok}\nFailed: {fail}\n\nFolder: {out_dir}"
                if comp_stats:
                    msg += (
                        f"\n\nCompression:\n"
                        f"• {comp_stats['comp_ok']} file(s) compressed\n"
                        f"• {comp_stats['comp_opt']} file(s) already optimal / skipped\n"
                        f"• {comp_stats['comp_fail']} file(s) failed to compress (uncompressed PDF kept)\n"
                        f"• Total size: {format_size(comp_stats['orig_total_size'])} → {format_size(comp_stats['new_total_size'])} "
                        f"(saved {format_size(comp_stats['saved_bytes'])} / {comp_stats['saved_pct']:.1f}%)"
                    )
                messagebox.showinfo("Done", msg)
        self.after(0, _u)

    def _cancel(self):
        if self._is_running:
            self._cancel_event.set()
            self._var_status.set("Cancelling...")
            self._btn_cancel.config(state="disabled")

    def _start(self):
        raw = self._var_in.get().strip()
        if not raw:
            messagebox.showwarning("Nothing Selected", "Choose a folder or PowerPoint files first.")
            return
        paths = [Path(p.strip('\'" ')) for p in raw.split(";") if p.strip('\'" ')]
        existing = [p for p in paths if p.exists()]
        if not existing:
            messagebox.showerror(
                "Not Found",
                "The selected PowerPoint file or folder no longer exists.\n\n"
                "Choose it again with 'Choose PowerPoint Files' or 'Choose Folder'."
            )
            return
        files: list[Path] = []
        for p in existing:
            if p.is_file() and p.suffix.lower() in PPT_EXTENSIONS:
                files.append(p)
            elif p.is_dir():
                for ext in PPT_EXTENSIONS:
                    files.extend(p.glob(f"*{ext}"))
        seen, unique = set(), []
        for f in files:
            r = f.resolve()
            if r not in seen: seen.add(r); unique.append(f)
        files = sorted(unique, key=lambda x: x.name.lower())
        if not files:
            messagebox.showwarning(
                "No Files Found",
                "No PowerPoint files (.ppt, .pptx) were found there.\n\n"
                "Pick a folder that contains PowerPoint presentations."
            )
            return
        out_raw = self._var_out.get().strip()
        out_dir = (Path(out_raw.strip('\'" ')) if out_raw
                   else (existing[0] if existing[0].is_dir() else existing[0].parent) / "PDF Output")
        self._cancel_event.clear()
        self._is_running = True
        self._btn_start.config(state="disabled")
        self._btn_cancel.config(state="normal")
        self._btn_open.config(state="disabled")
        self._pbar["value"] = 0
        self._var_status.set("Preparing...")
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        threading.Thread(
            target=_ppt_batch_worker,
            args=(files, out_dir, self._append_log, self._update_progress, self._on_done,
                  self._compress_opt.is_enabled(), self._compress_opt.get_preset_key(),
                  self._cancel_event),
            daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — WORD TO PDF
# ══════════════════════════════════════════════════════════════════════════════

WORD_EXTENSIONS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".odt"}

# wdFormatPDF constant for Word COM SaveAs
_WD_FORMAT_PDF = 17


def _word_convert_single(doc_path: Path, output_dir: Path, word_app, log_fn) -> bool:
    """Convert one Word document to PDF via COM automation.

    Returns True on success, False on failure (already logs the reason).
    """
    pdf_name = doc_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name
    doc = None
    try:
        if doc_path.stat().st_size == 0:
            log_fn(f"[FAIL] {doc_path.name}: File is empty (0 bytes).")
            return False
        if is_office_encrypted(doc_path):
            log_fn(f"[SKIP] {doc_path.name}: File is password-protected. Skipped.")
            return False
        doc = word_app.Documents.Open(
            str(doc_path.resolve()),
            ReadOnly=True,
            Visible=False,
            ConfirmConversions=False,
        )
        doc.SaveAs2(str(pdf_path.resolve()), FileFormat=_WD_FORMAT_PDF)
        log_fn(f"[OK]   {doc_path.name}")
        return True
    except Exception as e:
        log_fn(f"[FAIL] {doc_path.name}: {e}")
        if pdf_path.exists():
            try:
                pdf_path.unlink(missing_ok=True)
            except Exception:
                pass
        return False
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def _word_batch_worker(
    files, output_dir, log_fn, progress_fn, done_fn,
    compress_enabled=False, compress_preset="standard",
    cancel_event=None,
):
    """Worker thread: batch convert Word documents to PDF then optionally compress."""
    import pythoncom
    import win32com.client
    import pymupdf
    from pdf_optimizer import has_digital_signatures

    pythoncom.CoInitialize()
    word_app, was_running = None, False
    try:
        word_app = win32com.client.GetActiveObject("Word.Application")
        was_running = True
    except Exception:
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
        except Exception as e:
            log_fn(f"[ERROR] Microsoft Word not found: {e}")
            log_fn("[ERROR] Make sure Microsoft Word is installed and not blocked.")
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            try:
                done_fn(0, len(files), output_dir, None, word_missing=True)
            except TypeError:
                done_fn(0, len(files), output_dir, None)
            return

    output_dir.mkdir(parents=True, exist_ok=True)
    total = len(files)
    log_fn(f"[INFO] Converting {total} Word file(s) to:\n{output_dir}\n{'-'*40}")

    cancelled = False
    ok = fail = 0
    comp_ok = comp_opt = comp_fail = 0
    orig_total_size = new_total_size = 0
    t0 = time.time()

    try:
        for idx, f in enumerate(files):
            if cancel_event and cancel_event.is_set():
                log_fn("[INFO] Cancelled by user.")
                cancelled = True
                break

            file_num = idx + 1
            file_base = (idx / total) * 100.0
            file_weight = (1.0 / total) * 100.0

            progress_fn(file_num, total, f"File {file_num} of {total}: converting {f.name}...", int(file_base))

            conv_ok = _word_convert_single(f, output_dir, word_app, log_fn)

            if cancel_event and cancel_event.is_set():
                log_fn("[INFO] Cancelled by user.")
                cancelled = True
                break

            if conv_ok:
                ok += 1
                pdf_path = output_dir / (f.stem + ".pdf")
                if compress_enabled and pdf_path.exists():
                    tmp_target = output_dir / f"{f.stem}__opt_tmp.pdf"
                    try:
                        comp_start_pct = int(file_base + file_weight * 0.70)
                        progress_fn(file_num, total, f"File {file_num} of {total}: compressing PDF...", comp_start_pct)
                        log_fn(f"[INFO] Compressing {pdf_path.name}...")
                        f_orig_sz = pdf_path.stat().st_size
                        orig_total_size += f_orig_sz

                        is_signed = False
                        try:
                            chk_doc = pymupdf.open(str(pdf_path))
                            is_signed = has_digital_signatures(chk_doc)
                            chk_doc.close()
                        except Exception:
                            pass

                        if is_signed:
                            log_fn(f"[WARN] {pdf_path.name}: Document is digitally signed. Compression skipped.")
                            comp_opt += 1
                            new_total_size += f_orig_sz
                        else:
                            def on_comp_step(stage, comp_pct, _fb=file_base, _fw=file_weight, _fn=file_num, _tot=total):
                                current_pct = int(_fb + _fw * 0.70 + (_fw * 0.30) * (comp_pct / 100.0))
                                progress_fn(_fn, _tot, f"File {_fn} of {_tot}: {stage}", current_pct)

                            res = optimize_pdf(
                                pdf_path, dst_path=tmp_target,
                                preset_key=compress_preset,
                                on_progress=on_comp_step,
                                cancel_event=cancel_event,
                                ignore_signatures=False,
                            )
                            if res.status == "success" and res.dst_path:
                                Path(res.dst_path).replace(pdf_path)
                                comp_ok += 1
                                new_total_size += res.new_size
                                log_fn(
                                    f"[OK] {pdf_path.name}: "
                                    f"{format_size(res.orig_size)} → {format_size(res.new_size)} "
                                    f"(saved {res.saved_percent:.1f}%)"
                                )
                            elif res.status == "already_optimal":
                                comp_opt += 1
                                new_total_size += f_orig_sz
                                log_fn(f"[INFO] {pdf_path.name}: PDF is already optimal.")
                            else:
                                comp_fail += 1
                                new_total_size += f_orig_sz
                                log_fn(f"[WARN] {pdf_path.name}: Compression skipped ({res.message}).")
                    except CancelledError:
                        log_fn("[INFO] Cancelled by user.")
                        cancelled = True
                        break
                    except Exception as ce:
                        comp_fail += 1
                        if pdf_path.exists():
                            new_total_size += pdf_path.stat().st_size
                        log_fn(f"[WARN] Could not compress {pdf_path.name}: {ce}. The PDF was kept.")
                    finally:
                        if tmp_target.exists():
                            try:
                                tmp_target.unlink(missing_ok=True)
                            except Exception:
                                pass
            else:
                fail += 1

            progress_fn(file_num, total, f"File {file_num} of {total}: done", int(file_base + file_weight))

    except Exception as err:
        log_fn(f"[ERROR] Word worker stopped unexpectedly: {err}")
    finally:
        if not was_running and word_app:
            try:
                word_app.Quit(SaveChanges=0)
            except Exception:
                pass
        word_app = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        comp_stats = None
        if not cancelled:
            elapsed = time.time() - t0
            log_fn(f"[INFO] Finished in {elapsed:.1f}s | OK: {ok} | Failed: {fail}")
            if compress_enabled and ok > 0:
                saved_bytes = max(0, orig_total_size - new_total_size)
                saved_pct = (saved_bytes / orig_total_size * 100) if orig_total_size > 0 else 0.0
                log_fn(f"[INFO] Compression: {comp_ok} compressed, {comp_opt} already optimal, {comp_fail} failed.")
                log_fn(
                    f"[INFO] Total size: {format_size(orig_total_size)} → {format_size(new_total_size)} "
                    f"(saved {format_size(saved_bytes)} / {saved_pct:.1f}%)"
                )
                comp_stats = {
                    "comp_ok": comp_ok,
                    "comp_opt": comp_opt,
                    "comp_fail": comp_fail,
                    "orig_total_size": orig_total_size,
                    "new_total_size": new_total_size,
                    "saved_bytes": saved_bytes,
                    "saved_pct": saved_pct,
                }

        try:
            done_fn(ok, fail, output_dir, comp_stats, cancelled=cancelled)
        except TypeError:
            done_fn(ok, fail, output_dir, comp_stats)


class WordToPdfTab(tk.Frame):
    """Batch-convert Word files (.docx, .doc, ...) to PDF."""

    def __init__(self, parent, settings_ref=None, on_settings_change=None):
        super().__init__(parent, bg=C_BG)
        self._settings = settings_ref or {}
        self._on_settings_chg = on_settings_change
        self._last_out_dir = None
        self._cancel_event = threading.Event()
        self._is_running = False
        self._build()

    def _build(self):
        frm_in = ttk.LabelFrame(self, text=" File / Folder Word ", padding=(PAD, PAD_SM))
        frm_in.pack(fill="x", padx=PAD, pady=(PAD, PAD_SM))
        self._var_in = tk.StringVar()
        ttk.Entry(frm_in, textvariable=self._var_in, font=FONT_BODY).pack(fill="x", pady=(0, 6))
        br = ttk.Frame(frm_in)
        br.pack(fill="x")
        ttk.Button(br, text="Choose Folder", command=self._choose_folder).pack(side="left", padx=(0, 8))
        ttk.Button(br, text="Choose Word Files", command=self._choose_files).pack(side="left", padx=(0, 8))
        tk.Label(
            br,
            text=".docx · .doc · .docm · .rtf · .odt",
            font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG,
        ).pack(side="left")

        frm_out = ttk.LabelFrame(self, text=" Output Folder ", padding=(PAD, PAD_SM))
        frm_out.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._var_out = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self._var_out, font=FONT_BODY).pack(fill="x", pady=(0, 6))
        ttk.Button(frm_out, text="Choose Folder...", command=self._choose_output).pack(anchor="w")

        # Compression option — same widget used by PPT and Merge tabs
        frm_opt = ttk.LabelFrame(self, text=" PDF Compression ", padding=(PAD, PAD_SM))
        frm_opt.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._compress_opt = CompressOptions(
            frm_opt,
            settings_ref=self._settings,
            on_settings_change=self._on_settings_chg,
            enable_key="word_compress_enabled",
            preset_key="word_compress_preset",
        )
        self._compress_opt.pack(fill="x")
        self.compress_opts = self._compress_opt

        act = tk.Frame(self, bg=C_BG)
        act.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._btn_start = flat_btn(act, "Convert to PDF", self._start, padx=16, pady=6)
        self._btn_start.pack(side="left", padx=(0, 8))
        self._btn_cancel = flat_btn(
            act, "Cancel", self._cancel,
            bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
            font=FONT_BODY, padx=12, pady=6,
        )
        self._btn_cancel.pack(side="left", padx=(0, 8))
        self._btn_cancel.config(state="disabled")
        self._btn_open = flat_btn(
            act, "Open Output Folder", self._open_out,
            bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
            font=FONT_BODY, padx=12, pady=6,
        )
        self._btn_open.pack(side="left")
        self._btn_open.config(state="disabled")
        self._var_status = tk.StringVar(value="Ready.")
        tk.Label(
            act, textvariable=self._var_status,
            font=FONT_STATUS, fg=C_SUBTEXT, bg=C_BG,
        ).pack(side="left", padx=10)

        self._pbar = ttk.Progressbar(self, mode="determinate")
        self._pbar.pack(fill="x", padx=PAD, pady=(0, PAD_SM))

        frm_log = ttk.LabelFrame(self, text=" Log ", padding=(PAD, PAD_SM))
        frm_log.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self._log = scrolledtext.ScrolledText(
            frm_log, font=FONT_MONO, state="disabled", wrap="word", bg=C_WHITE,
        )
        self._log.pack(fill="both", expand=True)

    # ── Input ───────────────────────────────────────────────────────────

    def _choose_folder(self):
        d = filedialog.askdirectory(title="Choose a folder with Word files")
        if d:
            self._var_in.set(d)
            if not self._var_out.get():
                self._var_out.set(str(Path(d) / "PDF Output"))

    def _choose_files(self):
        exts = " ".join(f"*{e}" for e in sorted(WORD_EXTENSIONS))
        files = filedialog.askopenfilenames(
            title="Choose Word files",
            filetypes=[("Word / RTF / ODT", exts), ("All files", "*.*")],
        )
        if files:
            self._var_in.set("; ".join(files))
            if not self._var_out.get():
                self._var_out.set(str(Path(files[0]).parent / "PDF Output"))

    def _choose_output(self):
        d = filedialog.askdirectory(title="Choose the output folder for PDFs")
        if d:
            self._var_out.set(d)

    def _open_out(self):
        if self._last_out_dir and self._last_out_dir.exists():
            os.startfile(str(self._last_out_dir))

    # ── Log & Progress ────────────────────────────────────────────────────────

    def _append_log(self, msg):
        def _u():
            self._log.configure(state="normal")
            self._log.insert("end", msg + "\n")
            self._log.see("end")
            self._log.configure(state="disabled")
        self.after(0, _u)

    def _update_progress(self, cur, total, label, pct=None):
        def _u():
            self._pbar["value"] = pct if pct is not None else int(cur / total * 100)
            self._var_status.set(label)
        self.after(0, _u)

    # ── Done callback ─────────────────────────────────────────────────────────

    def _on_done(self, ok, fail, out_dir, comp_stats=None, word_missing=False, cancelled=False):
        def _u():
            self._is_running = False
            self._btn_start.config(state="normal")
            self._btn_cancel.config(state="disabled")
            if word_missing:
                self._var_status.set("Microsoft Word not found.")
                messagebox.showwarning(
                    "Microsoft Word Not Found",
                    "Microsoft Word could not be found or started.\n\n"
                    "Word to PDF needs Microsoft Word installed on this computer.",
                )
                return
            if cancelled:
                self._var_status.set("Cancelled by user.")
                return
            self._last_out_dir = out_dir
            if out_dir.exists():
                self._btn_open.config(state="normal")
            self._pbar["value"] = 100
            self._var_status.set(f"Done: {ok} converted, {fail} failed.")
            if ok + fail > 0:
                msg = f"Conversion finished.\n\nConverted: {ok}\nFailed: {fail}\n\nFolder: {out_dir}"
                if comp_stats:
                    msg += (
                        f"\n\nCompression:\n"
                        f"• {comp_stats['comp_ok']} file(s) compressed\n"
                        f"• {comp_stats['comp_opt']} file(s) already optimal / skipped\n"
                        f"• {comp_stats['comp_fail']} file(s) failed to compress (uncompressed PDF kept)\n"
                        f"• Total: {format_size(comp_stats['orig_total_size'])} → "
                        f"{format_size(comp_stats['new_total_size'])} "
                        f"(saved {format_size(comp_stats['saved_bytes'])} / {comp_stats['saved_pct']:.1f}%)"
                    )
                messagebox.showinfo("Done", msg)
        self.after(0, _u)

    # ── Cancel ────────────────────────────────────────────────────────────────

    def _cancel(self):
        if self._is_running:
            self._cancel_event.set()
            self._var_status.set("Cancelling...")
            self._btn_cancel.config(state="disabled")

    # ── Start ─────────────────────────────────────────────────────────────────

    def _start(self):
        raw = self._var_in.get().strip()
        if not raw:
            messagebox.showwarning("Nothing Selected", "Choose a folder or Word files first.")
            return

        paths = [Path(p.strip("'\" ")) for p in raw.split(";") if p.strip("'\" ")]
        existing = [p for p in paths if p.exists()]
        if not existing:
            messagebox.showerror(
                "Not Found",
                "The selected Word file or folder no longer exists.\n\n"
                "Choose it again with 'Choose Word Files' or 'Choose Folder'."
            )
            return

        files: list[Path] = []
        for p in existing:
            if p.is_file() and p.suffix.lower() in WORD_EXTENSIONS:
                files.append(p)
            elif p.is_dir():
                for ext in WORD_EXTENSIONS:
                    files.extend(p.glob(f"*{ext}"))

        seen, unique = set(), []
        for f in files:
            r = f.resolve()
            if r not in seen:
                seen.add(r)
                unique.append(f)
        files = sorted(unique, key=lambda x: x.name.lower())

        if not files:
            messagebox.showwarning(
                "No Files Found",
                "No Word files (.docx, .doc, .docm, .rtf, .odt) were found there.\n\n"
                "Pick a folder that contains Word documents.",
            )
            return

        out_raw = self._var_out.get().strip()
        out_dir = (
            Path(out_raw.strip("'\" ")) if out_raw
            else (existing[0] if existing[0].is_dir() else existing[0].parent) / "PDF Output"
        )

        self._cancel_event.clear()
        self._is_running = True
        self._btn_start.config(state="disabled")
        self._btn_cancel.config(state="normal")
        self._btn_open.config(state="disabled")
        self._pbar["value"] = 0
        self._var_status.set("Preparing...")
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

        threading.Thread(
            target=_word_batch_worker,
            args=(
                files, out_dir,
                self._append_log, self._update_progress, self._on_done,
                self._compress_opt.is_enabled(), self._compress_opt.get_preset_key(),
                self._cancel_event,
            ),
            daemon=True,
        ).start()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — MERGE PDF
# ══════════════════════════════════════════════════════════════════════════════

def _merge_pdfs(pdf_paths, output_path, cancel_event=None):
    from pypdf import PdfWriter, PdfReader
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    readers = []
    total_pages = 0
    try:
        for p in pdf_paths:
            if cancel_event and cancel_event.is_set():
                raise InterruptedError("Merge cancelled")
            p_res = str(p.resolve())
            r = PdfReader(p_res)
            readers.append(r)
            if r.is_encrypted:
                decrypted = 0
                try:
                    decrypted = r.decrypt("")
                except Exception:
                    decrypted = 0
                if not decrypted:
                    raise ValueError(
                        f"File '{p.name}' is encrypted and needs a password. Remove the password first."
                    )
            try:
                total_pages += len(r.pages)
            except Exception as pe:
                if "decrypt" in str(pe).lower() or type(pe).__name__ == "FileNotDecryptedError":
                    raise ValueError(
                        f"File '{p.name}' is encrypted and needs a password. Remove the password first."
                    ) from pe
                raise
            writer.append(r)

        if cancel_event and cancel_event.is_set():
            raise InterruptedError("Merge cancelled")

        with open(output_path, "wb") as f:
            writer.write(f)

        return total_pages
    except BaseException:
        if output_path.exists():
            try:
                output_path.unlink(missing_ok=True)
            except Exception:
                pass
        raise
    finally:
        try:
            writer.close()
        except Exception:
            pass
        for r in readers:
            try:
                r.close()
            except Exception:
                pass


class PdfMergerTab(tk.Frame):
    def __init__(self, parent, settings_ref=None, on_settings_change=None):
        super().__init__(parent, bg=C_BG)
        self._settings = settings_ref or {}
        self._on_settings_chg = on_settings_change
        self._items: list[Path] = []
        self._last_merged = None
        self._cancel_event = threading.Event()
        self._is_running = False
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=PAD, pady=(PAD, PAD_SM))
        for lbl, cmd in [("Add Files...", self._add_files),
                          ("Add Folder...", self._add_folder),
                          ("Move Up", self._move_up), ("Move Down", self._move_down),
                          ("Remove", self._remove), ("Clear", self._clear)]:
            ttk.Button(top, text=lbl, command=cmd).pack(side="left", padx=(0, 6))

        frm_lst = ttk.LabelFrame(self, text=" PDF Order ", padding=(PAD, PAD_SM))
        frm_lst.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD_SM))
        sb = ttk.Scrollbar(frm_lst, orient="vertical")
        self._lb = tk.Listbox(frm_lst, yscrollcommand=sb.set, font=FONT_SMALL,
                               selectmode="extended", activestyle="none",
                               bg=C_WHITE)
        sb.config(command=self._lb.yview)
        sb.pack(side="right", fill="y")
        self._lb.pack(fill="both", expand=True)

        frm_out = ttk.LabelFrame(self, text=" Save Merged PDF As ", padding=(PAD, PAD_SM))
        frm_out.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._var_out = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self._var_out, font=FONT_BODY
                  ).pack(fill="x", pady=(0, 6))
        ttk.Button(frm_out, text="Choose Name & Location...",
                   command=self._browse_out).pack(anchor="w")

        # PDF compression option
        frm_opt = ttk.LabelFrame(self, text=" PDF Compression ", padding=(PAD, PAD_SM))
        frm_opt.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._compress_opt = CompressOptions(
            frm_opt, settings_ref=self._settings,
            on_settings_change=self._on_settings_chg,
            enable_key="merger_compress_enabled",
            preset_key="merger_compress_preset",
        )
        self._compress_opt.pack(fill="x")
        self.compress_opts = self._compress_opt

        act = tk.Frame(self, bg=C_BG)
        act.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._btn_merge = flat_btn(act, "Merge PDF", self._start_merge, padx=16, pady=6)
        self._btn_merge.pack(side="left", padx=(0, 8))
        self._btn_cancel = flat_btn(act, "Cancel", self.do_cancel,
                                    bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                                    font=FONT_BODY, padx=12, pady=6)
        self._btn_cancel.pack(side="left", padx=(0, 8))
        self._btn_cancel.config(state="disabled")
        self._btn_pdf = flat_btn(act, "Open Result", self._open_merged,
                                  bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                                  font=FONT_BODY, padx=12, pady=6)
        self._btn_pdf.pack(side="left")
        self._btn_pdf.config(state="disabled")
        self._var_status = tk.StringVar(value="No PDFs added yet.")
        tk.Label(act, textvariable=self._var_status,
                 font=FONT_STATUS, fg=C_SUBTEXT, bg=C_BG
                 ).pack(side="left", padx=10)

        frm_log = ttk.LabelFrame(self, text=" Log ", padding=(PAD, PAD_SM))
        frm_log.pack(fill="x", padx=PAD, pady=(0, PAD))
        self._log = scrolledtext.ScrolledText(frm_log, height=4, font=FONT_MONO,
                                               state="disabled", wrap="word",
                                               bg=C_WHITE)
        self._log.pack(fill="both", expand=True)

    def _refresh(self):
        self._lb.delete(0, "end")
        for i, p in enumerate(self._items, 1):
            self._lb.insert("end", f"{i}.  {p.name}   [{p.parent}]")
        self._var_status.set(f"{len(self._items)} PDF file(s) selected.")
        if self._items and not self._var_out.get():
            self._var_out.set(str(self._items[0].parent / "Merged.pdf"))

    def _add_files(self):
        fs = filedialog.askopenfilenames(title="Choose PDF files",
            filetypes=[("PDF", "*.pdf"), ("All files", "*.*")])
        for f in fs:
            p = Path(f)
            if p not in self._items: self._items.append(p)
        self._refresh()

    def _add_folder(self):
        d = filedialog.askdirectory(title="Choose a folder with PDFs")
        if d:
            for f in sorted(Path(d).glob("*.pdf"), key=lambda x: x.name.lower()):
                if f not in self._items: self._items.append(f)
            self._refresh()

    def _move_up(self):
        sel = list(self._lb.curselection())
        if not sel or sel[0] == 0: return
        i = sel[0]
        self._items[i-1], self._items[i] = self._items[i], self._items[i-1]
        self._refresh(); self._lb.selection_set(i-1)

    def _move_down(self):
        sel = list(self._lb.curselection())
        if not sel or sel[0] >= len(self._items)-1: return
        i = sel[0]
        self._items[i+1], self._items[i] = self._items[i], self._items[i+1]
        self._refresh(); self._lb.selection_set(i+1)

    def _remove(self):
        for i in reversed(sorted(self._lb.curselection())): del self._items[i]
        self._refresh()

    def _clear(self):
        self._items.clear(); self._var_out.set(""); self._refresh()

    def _browse_out(self):
        p = filedialog.asksaveasfilename(
            title="Save merged PDF as", defaultextension=".pdf",
            initialfile=(Path(self._var_out.get()).name if self._var_out.get()
                         else "Merged.pdf"),
            filetypes=[("PDF", "*.pdf")])
        if p: self._var_out.set(p)

    def _open_merged(self):
        if self._last_merged and self._last_merged.exists():
            os.startfile(str(self._last_merged))

    def _log_msg(self, msg):
        def _u():
            self._log.configure(state="normal")
            self._log.insert("end", msg+"\n"); self._log.see("end")
            self._log.configure(state="disabled")
        self.after(0, _u)

    def do_cancel(self):
        if self._is_running:
            self._cancel_event.set()
            self.after(0, lambda: self._var_status.set("Cancelling..."))
            self._btn_cancel.config(state="disabled")
            self._log_msg("[INFO] Cancelling merge...")

    def _start_merge(self):
        if len(self._items) < 2:
            messagebox.showwarning("Not Enough Files", "Add at least 2 PDF files to merge.")
            return
        out_str = self._var_out.get().strip()
        if not out_str:
            messagebox.showwarning("Choose Where to Save",
                "Choose a file name and folder for the merged PDF.")
            return
        out_path = Path(out_str.strip('\'" '))
        items = list(self._items)

        self._is_running = True
        self._cancel_event.clear()
        self._btn_merge.config(state="disabled")
        self._btn_cancel.config(state="normal", bg=C_DANGER, fg="white")
        self._btn_pdf.config(state="disabled")
        self._var_status.set("Merging...")
        self._log.configure(state="normal"); self._log.delete("1.0","end")
        self._log.configure(state="disabled")

        def worker():
            t0 = time.time()
            self._log_msg(f"[INFO] Merging {len(items)} file...")
            tmp_opt = None
            merge_success = False
            try:
                pages = _merge_pdfs(items, out_path, cancel_event=self._cancel_event)
                elapsed = time.time()-t0

                if self._cancel_event.is_set():
                    raise InterruptedError("Merge cancelled")

                # Compress the merged file if enabled
                comp_result = None
                comp_status = "none"
                if self._compress_opt.is_enabled() and out_path.exists():
                    self._log_msg(f"[INFO] Compressing merged file {out_path.name}...")
                    self.after(0, lambda: self._var_status.set("Compressing merged PDF..."))

                    # Check for digital signatures
                    is_signed = False
                    try:
                        import pymupdf
                        chk_doc = pymupdf.open(str(out_path))
                        is_signed = has_digital_signatures(chk_doc)
                        chk_doc.close()
                    except Exception:
                        pass

                    if is_signed:
                        comp_status = "signed"
                        self._log_msg("[WARN] Merged document is digitally signed. Compression skipped.")
                    else:
                        try:
                            tmp_opt = out_path.parent / f"{out_path.stem}__opt_tmp_merged.pdf"
                            res = optimize_pdf(out_path, dst_path=tmp_opt,
                                               preset_key=self._compress_opt.get_preset_key(),
                                               cancel_event=self._cancel_event,
                                               ignore_signatures=False)
                            if res.status == "success" and res.dst_path:
                                Path(res.dst_path).replace(out_path)
                                comp_result = res
                                comp_status = "success"
                                self._log_msg(f"[OK] {format_size(res.orig_size)} → "
                                              f"{format_size(res.new_size)} (saved {res.saved_percent:.1f}%)")
                            elif res.status == "already_optimal":
                                comp_status = "already_optimal"
                                self._log_msg("[INFO] Merged PDF is already optimal.")
                            else:
                                comp_status = "skipped"
                                self._log_msg(f"[WARN] Compression skipped: {res.message}")
                        except CancelledError:
                            raise InterruptedError("Merge cancelled")
                        except Exception as ce:
                            comp_status = "failed"
                            self._log_msg(f"[WARN] Could not compress the merged file: {ce}. The merged PDF was kept.")

                merge_success = True

                def ok():
                    self._is_running = False
                    self._btn_merge.config(state="normal")
                    self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
                    self._last_merged = out_path
                    self._btn_pdf.config(state="normal")
                    self._var_status.set(f"Done: {pages} pages in total.")
                    self._log_msg(f"[OK] {len(items)} file → {out_path.name}  "
                                  f"({pages} pages, {elapsed:.2f}s)")
                    msg_text = f"PDFs merged.\nTotal: {pages} pages\n"
                    if comp_result:
                        msg_text += f"Size: {format_size(comp_result.orig_size)} → {format_size(comp_result.new_size)} " \
                                    f"(saved {comp_result.saved_percent:.1f}%)\n\n"
                    elif comp_status == "already_optimal":
                        try:
                            msg_text += f"Size: {format_size(out_path.stat().st_size)} (already optimal)\n\n"
                        except Exception:
                            pass
                    elif comp_status == "signed":
                        try:
                            msg_text += f"Size: {format_size(out_path.stat().st_size)} (compression skipped: signed document)\n\n"
                        except Exception:
                            pass
                    else:
                        try:
                            msg_text += f"Size: {format_size(out_path.stat().st_size)}\n\n"
                        except Exception:
                            pass
                    msg_text += f"Saved to: {out_path}"
                    messagebox.showinfo("Done", msg_text)
                self.after(0, ok)

            except (InterruptedError, CancelledError):
                def on_cancel():
                    self._is_running = False
                    self._btn_merge.config(state="normal")
                    self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
                    self._btn_pdf.config(state="normal" if (self._last_merged and self._last_merged.exists()) else "disabled")
                    self._var_status.set("Cancelled.")
                    self._log_msg("[INFO] Merge cancelled by user.")
                    messagebox.showinfo("Cancelled", "The merge was cancelled.")
                self.after(0, on_cancel)

            except ValueError as ve:
                def on_val_error():
                    self._is_running = False
                    self._btn_merge.config(state="normal")
                    self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
                    self._btn_pdf.config(state="normal" if (self._last_merged and self._last_merged.exists()) else "disabled")
                    self._var_status.set("Failed: encrypted file.")
                    self._log_msg(f"[FAIL] {ve}")
                    messagebox.showerror(
                        "Could Not Merge PDFs",
                        f"A file is password-protected or encrypted:\n{ve}\n\n"
                        "Remove the password from that file, then merge again."
                    )
                self.after(0, on_val_error)

            except Exception as e:
                def fail():
                    self._is_running = False
                    self._btn_merge.config(state="normal")
                    self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
                    self._btn_pdf.config(state="normal" if (self._last_merged and self._last_merged.exists()) else "disabled")
                    self._var_status.set("Failed.")
                    self._log_msg(f"[FAIL] {e}")
                    messagebox.showerror(
                        "Could Not Merge PDFs",
                        f"Something went wrong while merging:\n{e}\n\n"
                        "Make sure no PDF is open in another app and none of them is damaged."
                    )
                self.after(0, fail)
            finally:
                if not merge_success:
                    if out_path.exists():
                        try:
                            out_path.unlink(missing_ok=True)
                        except Exception:
                            pass
                if tmp_opt and tmp_opt.exists():
                    try:
                        tmp_opt.unlink(missing_ok=True)
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — COMPRESS PDF (STANDALONE)
# ══════════════════════════════════════════════════════════════════════════════

class PdfOptimizerTab(tk.Frame):
    """
    Tab 5: standalone PDF compressor.
    Drag-and-drop, quick pre-check, 3 presets, file info,
    cooperative cancel, progress bar, result panel and log.
    """
    def __init__(self, parent, settings_ref=None, on_settings_change=None):
        super().__init__(parent, bg=C_BG)
        self._settings = settings_ref or {}
        self._on_settings_chg = on_settings_change

        self._input_path: Optional[Path] = None
        self._inspect_info: Optional[pdf_optimizer.InspectInfo] = None
        self._is_running = False
        self._cancel_event = threading.Event()
        self._queue = queue.Queue()
        self._dnd_available = False
        self._last_result: Optional[pdf_optimizer.OptimizationResult] = None
        self._password_used: Optional[str] = None

        self._build()
        self._poll_queue()

    def _build(self):
        # 1. Drop zone & file picker
        self._frm_drop = tk.Frame(self, bg=C_WHITE, highlightthickness=1,
                                  highlightbackground=C_BORDER, padx=PAD, pady=PAD)
        self._frm_drop.pack(fill="x", padx=PAD, pady=(PAD, PAD_SM))

        self._lbl_drop = tk.Label(
            self._frm_drop,
            text="Drop a PDF here, or use the button below",
            font=FONT_BODY, fg=C_TEXT, bg=C_WHITE)
        self._lbl_drop.pack(pady=(4, 6))

        br_pick = tk.Frame(self._frm_drop, bg=C_WHITE)
        br_pick.pack()
        flat_btn(br_pick, "Choose PDF...", self._choose_file, padx=14, pady=5).pack(side="left")

        self._setup_dnd()

        # 2. File info
        self._frm_info = ttk.LabelFrame(self, text=" File Info ", padding=(PAD, PAD_SM))
        self._frm_info.pack(fill="x", padx=PAD, pady=(0, PAD_SM))

        self._lbl_fname = tk.Label(self._frm_info, text="No file selected",
                                   font=FONT_TITLE, fg=C_TEXT, bg=C_BG, anchor="w")
        self._lbl_fname.pack(fill="x")

        self._lbl_fmeta = tk.Label(self._frm_info, text="-",
                                   font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG, anchor="w")
        self._lbl_fmeta.pack(fill="x", pady=(2, 0))

        self._lbl_fstatus = tk.Label(self._frm_info, text="-",
                                     font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG, anchor="w")
        self._lbl_fstatus.pack(fill="x", pady=(2, 0))

        # 3. Compression level
        frm_pre = ttk.LabelFrame(self, text=" Compression Level ", padding=(PAD, PAD_SM))
        frm_pre.pack(fill="x", padx=PAD, pady=(0, PAD_SM))

        self._var_preset = tk.StringVar(value=self._settings.get("optimizer_preset", self._settings.get("pdf_compress_preset", "standard")))

        for key, title, desc in [
            ("standard", "Standard (recommended)", "150 DPI, JPEG quality 80. Balanced; text stays sharp"),
            ("light",    "Light",                      "Images untouched. Removes duplicates and metadata (lossless)"),
            ("maximum",  "Maximum",                    "110 DPI, JPEG quality 65. Smallest file, for email or upload limits"),
        ]:
            rf = tk.Frame(frm_pre, bg=C_BG)
            rf.pack(fill="x", pady=2)
            rb = ttk.Radiobutton(rf, text=title, variable=self._var_preset, value=key,
                                 command=self._on_preset_change)
            rb.pack(anchor="w")
            tk.Label(rf, text=desc, font=FONT_SMALL, fg=C_SUBTEXT, bg=C_BG).pack(anchor="w", padx=(28, 0))

        # 4. Output location
        frm_out = ttk.LabelFrame(self, text=" Output File ", padding=(PAD, PAD_SM))
        frm_out.pack(fill="x", padx=PAD, pady=(0, PAD_SM))
        self._var_out = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self._var_out, font=FONT_BODY).pack(fill="x", pady=(0, 4))
        ttk.Button(frm_out, text="Change Location...", command=self._change_output).pack(anchor="w")

        # 5. Actions & status
        act = tk.Frame(self, bg=C_BG)
        act.pack(fill="x", padx=PAD, pady=(0, PAD_SM))

        self._btn_start = flat_btn(act, "Compress", self._start_compress, padx=16, pady=6)
        self._btn_start.pack(side="left", padx=(0, 8))
        self._btn_start.config(state="disabled")

        self._btn_cancel = flat_btn(act, "Cancel", self._cancel_compress,
                                    bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                                    padx=14, pady=6)
        self._btn_cancel.pack(side="left", padx=(0, 10))
        self._btn_cancel.config(state="disabled")

        self._var_status = tk.StringVar(value="Choose a PDF to start.")
        tk.Label(act, textvariable=self._var_status,
                 font=FONT_STATUS, fg=C_SUBTEXT, bg=C_BG).pack(side="left", fill="x")

        # 6. Progress bar
        self._pbar = ttk.Progressbar(self, mode="determinate")
        self._pbar.pack(fill="x", padx=PAD, pady=(0, PAD_SM))

        # 7. Result panel (hidden until done)
        self._frm_res = tk.Frame(self, bg=C_WHITE, relief="solid", bd=1, padx=10, pady=8)
        # packed only when done

        # 8. Log Panel
        frm_log = ttk.LabelFrame(self, text=" Log ", padding=(PAD, PAD_SM))
        frm_log.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self._log = scrolledtext.ScrolledText(frm_log, font=FONT_MONO, height=5,
                                               state="disabled", wrap="word", bg=C_WHITE)
        self._log.pack(fill="both", expand=True)

    def _setup_dnd(self):
        try:
            self._frm_drop.drop_target_register("DND_Files")
            self._frm_drop.dnd_bind("<<Drop>>", self._on_drop)
            self._dnd_available = True
        except Exception:
            self._dnd_available = False

    def _on_drop(self, event):
        tokens = re.findall(r'\{[^}]+\}|\S+', event.data.strip())
        paths = [t.strip("{}").strip() for t in tokens if os.path.isfile(t.strip("{}"))]
        if not paths:
            return
        if len(paths) > 1:
            self._log_msg(f"[WARN] Several files dropped. Using the first one: {Path(paths[0]).name}")
            messagebox.showinfo("Multiple Files Dropped", "This tab compresses one file at a time.\nOnly the first file will be used.")
        self._load_file(Path(paths[0]))

    def _choose_file(self):
        f = filedialog.askopenfilename(
            title="Choose a PDF to compress",
            filetypes=[("PDF Document", "*.pdf"), ("All files", "*.*")])
        if f:
            self._load_file(Path(f))

    def _load_file(self, path: Path):
        if path.suffix.lower() != ".pdf":
            messagebox.showerror("Wrong File Type", "Choose a .pdf file.")
            return

        self._input_path = path
        self._password_used = None
        self._frm_res.pack_forget()

        # Pre-check & inspect
        pw = None
        info = None
        try:
            info = inspect_pdf(path, password=None)
        except CorruptPdf as ce:
            messagebox.showerror("Damaged File", f"This PDF cannot be read:\n{ce}")
            self._log_msg(f"[FAIL] {path.name} is damaged: {ce}")
            return

        if info.needs_password:
            pw = simpledialog.askstring("Password Required",
                f"'{path.name}' is password-protected.\nEnter the password:",
                show="*")
            if not pw:
                self._log_msg(f"[WARN] Password entry cancelled for {path.name}")
                return

            max_retries = 3
            for attempt in range(1, max_retries + 1):
                try:
                    info = inspect_pdf(path, password=pw)
                except CorruptPdf as ce:
                    messagebox.showerror("Damaged File", f"This PDF cannot be read:\n{ce}")
                    self._log_msg(f"[FAIL] {path.name} is damaged: {ce}")
                    return

                if info.is_valid:
                    break
                if attempt < max_retries:
                    pw = simpledialog.askstring("Password Salah",
                        f"Wrong password. Try again ({attempt}/{max_retries}):",
                        show="*")
                    if not pw:
                        self._log_msg(f"[WARN] Password entry cancelled for {path.name}")
                        return
                else:
                    messagebox.showerror("Could Not Open", "Wrong password 3 times. Stopped.")
                    self._log_msg(f"[WARN] Wrong password 3 times for {path.name}")
                    return

        if not info or not info.is_valid:
            messagebox.showerror("Could Not Open", "The PDF could not be opened, or the password is wrong.")
            return

        self._inspect_info = info
        self._password_used = pw

        # Update labels
        self._lbl_fname.config(text=path.name)
        self._lbl_fmeta.config(
            text=f"Size: {format_size(info.file_size)}   ·   "
                 f"Pages: {info.page_count}   ·   "
                 f"Images: {info.images_count}")

        status_text = "Ready to compress."
        if info.has_signatures:
            status_text = "Digitally signed: compressing will invalidate the signature"
        elif info.needs_password:
            status_text = "Encrypted (password OK)"
        elif info.is_encrypted:
            status_text = "Encrypted / restricted (encryption is kept)"
        self._lbl_fstatus.config(text=status_text)

        out_path = build_output_path(path)
        self._var_out.set(str(out_path))

        self._btn_start.config(state="normal")
        self._var_status.set("Ready. Pick a level and click 'Compress'.")
        self._log_msg(f"[INFO] Loaded: {path.name} ({format_size(info.file_size)}, {info.page_count} pages)")

    def _change_output(self):
        cur = self._var_out.get().strip()
        initial = Path(cur).name if cur else "compressed.pdf"
        p = filedialog.asksaveasfilename(
            title="Choose where to save",
            defaultextension=".pdf",
            initialfile=initial,
            filetypes=[("PDF Document", "*.pdf")])
        if p:
            self._var_out.set(p)

    def _on_preset_change(self):
        k = self._var_preset.get()
        self._settings["optimizer_preset"] = k
        self._settings["pdf_compress_preset"] = k
        save_settings(self._settings)
        if self._on_settings_chg:
            try: self._on_settings_chg(self._settings)
            except Exception: pass

    def _log_msg(self, msg):
        def _u():
            self._log.configure(state="normal")
            timestamp = time.strftime("[%H:%M:%S]")
            self._log.insert("end", f"{timestamp} {msg}\n")
            self._log.see("end")
            self._log.configure(state="disabled")
        self.after(0, _u)

    def _start_compress(self):
        if not self._input_path or not self._input_path.is_file():
            messagebox.showwarning("No File", "Choose a PDF first.")
            return

        # Confirm very large files
        if self._inspect_info and self._inspect_info.file_size > 500 * 1024 * 1024:
            if not messagebox.askyesno(
                "File Sangat Besar",
                f"This file is very large ({format_size(self._inspect_info.file_size)}).\n"
                "Compression may take a long time.\n\nContinue?"):
                return

        # Digital signature warning
        ignore_sigs = False
        if self._inspect_info and self._inspect_info.has_signatures:
            if not messagebox.askyesno(
                "Digitally Signed Document",
                "This document is digitally signed.\n"
                "Compressing it will invalidate the signature.\n\n"
                "Continue anyway?"):
                return
            ignore_sigs = True

        out_str = self._var_out.get().strip()
        if not out_str:
            messagebox.showwarning("No Output Location", "Choose where to save the result.")
            return
        dst_path = Path(out_str)

        self._is_running = True
        self._cancel_event.clear()
        self._btn_start.config(state="disabled")
        self._btn_cancel.config(state="normal", bg=C_DANGER, fg="white")
        self._pbar["value"] = 0
        self._frm_res.pack_forget()

        src = self._input_path
        pw = self._password_used
        preset_k = self._var_preset.get()

        self._log_msg(f"[INFO] Compressing: {src.name} (Preset: {preset_k})...")

        def worker():
            try:
                def on_p(stage, pct):
                    self._queue.put(("progress", stage, pct))

                res = optimize_pdf(
                    src,
                    dst_path=dst_path,
                    preset_key=preset_k,
                    password=pw,
                    on_progress=on_p,
                    cancel_event=self._cancel_event,
                    ignore_signatures=ignore_sigs
                )
                self._queue.put(("done", res))
            except CancelledError:
                self._queue.put(("cancelled", None))
            except PasswordRequired:
                self._queue.put(("error", "A password is needed to open this document."))
            except MemoryError:
                self._queue.put(("error", "Not enough memory for this file. Close other apps or use the Light preset."))
            except (PermissionError, OSError) as oe:
                self._queue.put(("error_permission", str(oe)))
            except Exception as e:
                self._queue.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_compress(self):
        if self._is_running:
            self._cancel_event.set()
            self._var_status.set("Cancelling...")
            self._btn_cancel.config(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                item = self._queue.get_nowait()
                if not item:
                    continue
                if isinstance(item, tuple) and len(item) == 3:
                    msg_type = item[0]
                    data = (item[1], item[2])
                elif isinstance(item, tuple) and len(item) == 2:
                    msg_type, data = item
                elif isinstance(item, tuple) and len(item) == 1:
                    msg_type = item[0]
                    data = None
                else:
                    msg_type = item[0]
                    data = item[1:] if len(item) > 1 else None

                if msg_type == "progress":
                    if isinstance(data, (tuple, list)):
                        stage = data[0] if len(data) > 0 else ""
                        pct = data[1] if len(data) > 1 else 0
                    else:
                        stage, pct = data, 0
                    self._pbar["value"] = pct
                    self._var_status.set(f"{stage} ({pct}%)")
                elif msg_type == "done":
                    self._on_done(data)
                elif msg_type == "cancelled":
                    self._on_cancelled()
                elif msg_type == "error":
                    self._on_error(data, is_permission=False)
                elif msg_type == "error_permission":
                    self._on_error(data, is_permission=True)
        except queue.Empty:
            pass
        except Exception as e:
            logger.error("Error processing queue message: %s", e)
        finally:
            self.after(100, self._poll_queue)

    def _on_done(self, res: pdf_optimizer.OptimizationResult):
        self._is_running = False
        self._btn_start.config(state="normal")
        self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
        self._last_result = res

        if res.status == "already_optimal":
            self._pbar["value"] = 100
            self._var_status.set("Done: the file is already optimal.")
            self._log_msg(f"[INFO] {res.message}")
            messagebox.showinfo("Already Optimal",
                "This file is already well optimised.\n"
                "Compressing would not save meaningful space, so no new file was written.")
            return

        self._pbar["value"] = 100
        self._var_status.set(f"Done. Saved {res.saved_percent:.1f}%.")
        self._log_msg(f"[OK] {res.src_path} → {res.dst_path} "
                      f"({format_size(res.orig_size)} → {format_size(res.new_size)}, "
                      f"saved {res.saved_percent:.1f}%)")

        # Show result panel
        for w in self._frm_res.winfo_children():
            w.destroy()

        self._frm_res.pack(fill="x", padx=PAD, pady=(0, PAD_SM), before=self._log.master)

        tk.Label(self._frm_res, text=" Compression Result ",
                 font=FONT_TITLE, fg=C_SUCCESS, bg=C_WHITE).pack(anchor="w")

        info_box = tk.Frame(self._frm_res, bg=C_WHITE)
        info_box.pack(fill="x", pady=(4, 6))

        tk.Label(info_box,
                 text=f"Before: {format_size(res.orig_size)}   ➔   "
                      f"After: {format_size(res.new_size)}   "
                      f"(saved {format_size(res.saved_bytes)} / {res.saved_percent:.1f}%)",
                 font=FONT_SUBTITLE, fg=C_TEXT, bg=C_WHITE).pack(anchor="w")

        # Warning notes
        notes = []
        if res.images_skipped > 0:
            notes.append(f"{res.images_skipped} image(s) skipped (transparency / 1-bit / special format)")
        if res.cmyk_converted > 0:
            notes.append(f"{res.cmyk_converted} CMYK image(s) converted to RGB (colours may shift slightly)")

        if notes:
            lbl_note = tk.Label(info_box, text="Note: " + "; ".join(notes),
                                font=FONT_SMALL, fg=C_SUBTEXT, bg=C_WHITE)
            lbl_note.pack(anchor="w", pady=(2, 0))

        br_btn = tk.Frame(self._frm_res, bg=C_WHITE)
        br_btn.pack(fill="x", pady=(4, 0))

        flat_btn(br_btn, "Open Result",
                 lambda: os.startfile(str(res.dst_path)) if res.dst_path and Path(res.dst_path).exists() else None,
                 padx=14, pady=5).pack(side="left", padx=(0, 8))

        flat_btn(br_btn, "Open Folder",
                 lambda: open_folder_and_select(Path(res.dst_path)) if res.dst_path else None,
                 bg=C_BTN_SEC, fg=C_TEXT, active_bg=C_BTN_SECH,
                 font=FONT_BODY, padx=12, pady=5).pack(side="left")

    def _on_cancelled(self):
        self._is_running = False
        self._btn_start.config(state="normal")
        self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
        self._pbar["value"] = 0
        self._var_status.set("Compression cancelled.")
        self._log_msg("[INFO] Compression cancelled by user.")

    def _on_error(self, err_msg: str, is_permission: bool = False):
        self._is_running = False
        self._btn_start.config(state="normal")
        self._btn_cancel.config(state="disabled", bg=C_BTN_SEC, fg=C_TEXT)
        self._pbar["value"] = 0
        self._var_status.set("Failed.")
        self._log_msg(f"[ERROR] {err_msg}")
        if is_permission or any(k in err_msg.lower() for k in ("permission", "akses ditolak", "denied", "read-only")):
            if messagebox.askyesno(
                "Cannot Write Output",
                f"The output location is not writable (access denied or disk full):\n\n{err_msg}\n\n"
                "Choose another location?",
                parent=self
            ):
                self._change_output()
        else:
            messagebox.showerror(
                "Could Not Compress",
                f"Something went wrong while compressing:\n\n{err_msg}\n\n"
                "Try this:\n"
                "• Use the 'Light' preset for very long or image-heavy documents.\n"
                "• Close the file in other apps (Acrobat, your browser).\n"
                "• Check that the output folder is writable and the disk has space."
            )


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════════

class MToolsApp:
    def __init__(self, initial_files=None):
        try:
            import tkinterdnd2
            self.root = tkinterdnd2.Tk()
        except ImportError:
            self.root = tk.Tk()

        self.root.title(f"{APP_NAME} — Document Toolkit")
        self.root.geometry("700x800")
        self.root.minsize(600, 680)
        self.root.configure(bg=C_BG)
        self.root.resizable(True, True)

        # Set window icon
        icon_path = _resolve_icon_path()
        if icon_path and os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass
            try:
                from PIL import ImageTk, Image
                _img = ImageTk.PhotoImage(Image.open(icon_path))
                self.root.iconphoto(True, _img)
                self._icon_ref = _img
            except Exception:
                pass

        self._settings = load_settings()
        setup_optimizer_logging()

        # Style: Windows 11 look via sv-ttk; fall back to the plain theme if it is missing.
        style = ttk.Style()
        try:
            import sv_ttk
            sv_ttk.set_theme("light")
            style.configure("TLabelframe.Label", foreground=C_SUBTEXT)
        except Exception:
            style.theme_use("default")
            style.configure("TProgressbar", troughcolor=C_SEP,
                            background=C_PRIMARY, thickness=5)
            style.configure("TNotebook", background=C_BG, borderwidth=0)
            style.configure("TNotebook.Tab", font=FONT_SMALL, padding=[14, 6],
                            background=C_BTN_SEC, foreground=C_TEXT)
            style.map("TNotebook.Tab",
                      background=[("selected", C_WHITE)],
                      foreground=[("selected", C_PRIMARY)])
            style.configure("TLabelframe", background=C_BG)
            style.configure("TLabelframe.Label", background=C_BG,
                            font=FONT_SMALL, foreground=C_SUBTEXT)

        # Status bar is packed before the notebook so the notebook cannot squeeze it out.
        self._statusbar = tk.Frame(self.root, bg=C_SURFACE, pady=3,
                                   relief="flat", bd=0)
        self._statusbar.pack(fill="x", side="bottom")
        tk.Frame(self._statusbar, bg=C_SEP, height=1).pack(fill="x")
        self._lbl_status = tk.Label(self._statusbar, text="",
                                     font=FONT_CAPTION, fg=C_SUBTEXT,
                                     bg=C_SURFACE, anchor="w")
        self._lbl_status.pack(fill="x", padx=10, pady=2)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=PAD_SM, pady=(PAD_SM, 0))

        self._tab_md = MarkItDownTab(
            notebook,
            settings_ref=self._settings,
            on_settings_change=self._on_settings_changed,
            statusbar_update=self._refresh_statusbar,
        )
        notebook.add(self._tab_md, text="Convert to Text")

        self._tab_ppt = PptToPdfTab(
            notebook,
            settings_ref=self._settings,
            on_settings_change=self._on_settings_changed,
        )
        notebook.add(self._tab_ppt, text="PowerPoint to PDF")

        self._tab_word = WordToPdfTab(
            notebook,
            settings_ref=self._settings,
            on_settings_change=self._on_settings_changed,
        )
        notebook.add(self._tab_word, text="Word to PDF")

        self._tab_merger = PdfMergerTab(
            notebook,
            settings_ref=self._settings,
            on_settings_change=self._on_settings_changed,
        )
        notebook.add(self._tab_merger, text="Merge PDF")

        self._tab_opt = PdfOptimizerTab(
            notebook,
            settings_ref=self._settings,
            on_settings_change=self._on_settings_changed,
        )
        notebook.add(self._tab_opt, text="Compress PDF")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._refresh_statusbar()
        # Warm up MarkItDown in background thread to eliminate initial cold-start lag (~1.3s)
        def _warmup():
            try:
                from markitdown import MarkItDown
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    _md = MarkItDown(enable_plugins=True)
                    del _md
            except Exception:
                pass

        threading.Thread(target=_warmup, daemon=True).start()

        if initial_files:
            valid = [f for f in initial_files if os.path.isfile(f)]
            if valid:
                self.root.after(300, lambda: self._tab_md._enqueue(valid))

        self.root.mainloop()

    def _refresh_statusbar(self):
        mode = ocr_mode(self._settings)
        if mode == "ai":
            prov  = self._settings.get("provider", "gemini")
            model = self._settings.get("model", "")
            prov_name = "Custom URL" if (self._settings.get("base_url") or "").strip() else (
                "Gemini" if prov == "gemini" else "OpenAI")
            txt = f"OCR: AI Vision ({prov_name} / {model})"
        else:
            txt = "OCR: Offline (on this computer)" if mode == "offline" else "OCR: Off"
        self._lbl_status.config(text=f"  {APP_NAME} {APP_VERSION}   |   {txt}")

    def _on_settings_changed(self, new_settings):
        self._settings.update(new_settings)
        # Notify tabs with CompressOptions to refresh UI
        for tab in (getattr(self, "_tab_ppt", None),
                    getattr(self, "_tab_word", None),
                    getattr(self, "_tab_merger", None)):
            comp = getattr(tab, "_compress_opt", getattr(tab, "compress_opts", None)) if tab else None
            if comp:
                try:
                    comp.refresh_from_settings(self._settings)
                except Exception:
                    pass
        if getattr(self, "_tab_opt", None) and hasattr(self._tab_opt, "_var_preset"):
            opt_p = self._settings.get("optimizer_preset", "standard")
            self._tab_opt._var_preset.set(opt_p)
        self._refresh_statusbar()

    def _on_close(self):
        # Save settings on normal close
        try:
            save_settings(self._settings)
        except Exception:
            pass

        # Check PDF optimizer running
        if getattr(self, "_tab_opt", None) and self._tab_opt._is_running:
            if not messagebox.askyesno(
                "Quit?",
                "PDF compression is still running.\nCancel it and quit?"):
                return
            self._tab_opt._cancel_compress()
        # Check PowerPoint conversion running
        if getattr(self, "_tab_ppt", None) and self._tab_ppt._is_running:
            if not messagebox.askyesno(
                "Quit?",
                "PowerPoint to PDF is still running.\nCancel it and quit?"):
                return
            self._tab_ppt._cancel()
        # Check Word-to-PDF conversion running
        if getattr(self, "_tab_word", None) and self._tab_word._is_running:
            if not messagebox.askyesno(
                "Quit?",
                "Word to PDF is still running.\nCancel it and quit?"):
                return
            self._tab_word._cancel()
        self.root.destroy()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    initial_files = []
    for arg in sys.argv[1:]:
        p = arg.strip().strip('"')
        if p and os.path.isfile(p):
            initial_files.append(p)
    MToolsApp(initial_files=initial_files if initial_files else None)