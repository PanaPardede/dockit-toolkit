"""
test_empirical_challenger_m2.py — Empirical challenge suite for Milestone M2.

Tests:
1. Settings persistence across all 5 tab keys:
   - Default loading schema integrity
   - Independent mutation & reload round-trip
   - Legacy backward-compatibility fallback
   - CompressOptions component state & tab binding
   - Tab synchronization across tabs via MToolsApp._on_settings_changed
2. Atomic write verification:
   - Verification that settings.json.tmp is cleanly replaced and no .tmp persists in normal sequential saves
   - Repeated sequential save consistency
3. Forbidden token audit:
   - Strict audit for '(FR-2.5)', '[KOMPRES]', '[BATAL]', 'Kurang File' in markitdown_gui.py
"""

import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path

import markitdown_gui
from markitdown_gui import (
    load_settings,
    save_settings,
    CompressOptions,
    PptToPdfTab,
    WordToPdfTab,
    PdfMergerTab,
)


class TestSettingsPersistence5Tabs(unittest.TestCase):
    """Verifies settings persistence across all 5 tabs."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.settings_file = Path(self.temp_dir) / "settings.json"
        self._orig_settings_file = markitdown_gui.SETTINGS_FILE
        markitdown_gui.SETTINGS_FILE = str(self.settings_file)

    def tearDown(self):
        markitdown_gui.SETTINGS_FILE = self._orig_settings_file
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_default_settings_contains_all_5_tab_keys(self):
        """Verify defaults contain dedicated keys for all 5 tabs."""
        s = load_settings()
        # Tab 1: OCR / Text
        self.assertIn("ocr_engine", s)
        self.assertIn("ocr_language", s)
        self.assertIn("ai_vision_enabled", s)
        # Tab 2: PPT to PDF
        self.assertIn("ppt_compress_enabled", s)
        self.assertIn("ppt_compress_preset", s)
        # Tab 3: Word to PDF
        self.assertIn("word_compress_enabled", s)
        self.assertIn("word_compress_preset", s)
        # Tab 4: PDF Merger
        self.assertIn("merger_compress_enabled", s)
        self.assertIn("merger_compress_preset", s)
        # Tab 5: PDF Optimizer
        self.assertIn("optimizer_preset", s)

    def test_independent_mutation_and_persistence_roundtrip(self):
        """Mutate keys across all 5 tabs independently, save, reload, and verify no cross-contamination."""
        s = load_settings()

        # Tab 1
        s["ocr_engine"] = "tesseract"
        s["ocr_language"] = "eng"
        s["ai_vision_enabled"] = True
        s["model"] = "gemini-3.6-flash"

        # Tab 2
        s["ppt_compress_enabled"] = True
        s["ppt_compress_preset"] = "maximum"

        # Tab 3
        s["word_compress_enabled"] = False
        s["word_compress_preset"] = "light"

        # Tab 4
        s["merger_compress_enabled"] = True
        s["merger_compress_preset"] = "standard"

        # Tab 5
        s["optimizer_preset"] = "maximum"

        ok = save_settings(s)
        self.assertTrue(ok)

        # Reload from disk
        reloaded = load_settings()

        # Check Tab 1
        self.assertEqual(reloaded["ocr_engine"], "tesseract")
        self.assertEqual(reloaded["ocr_language"], "eng")
        self.assertTrue(reloaded["ai_vision_enabled"])

        # Check Tab 2
        self.assertTrue(reloaded["ppt_compress_enabled"])
        self.assertEqual(reloaded["ppt_compress_preset"], "maximum")

        # Check Tab 3
        self.assertFalse(reloaded["word_compress_enabled"])
        self.assertEqual(reloaded["word_compress_preset"], "light")

        # Check Tab 4
        self.assertTrue(reloaded["merger_compress_enabled"])
        self.assertEqual(reloaded["merger_compress_preset"], "standard")

        # Check Tab 5
        self.assertEqual(reloaded["optimizer_preset"], "maximum")

    def test_legacy_settings_fallback(self):
        """Legacy settings file with only pdf_compress_* should propagate to all tabs gracefully."""
        legacy = {
            "pdf_compress_enabled": True,
            "pdf_compress_preset": "maximum",
        }
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(legacy, f)

        s = load_settings()
        self.assertTrue(s["ppt_compress_enabled"])
        self.assertEqual(s["ppt_compress_preset"], "maximum")
        self.assertTrue(s["word_compress_enabled"])
        self.assertEqual(s["word_compress_preset"], "maximum")
        self.assertTrue(s["merger_compress_enabled"])
        self.assertEqual(s["merger_compress_preset"], "maximum")
        self.assertEqual(s["optimizer_preset"], "maximum")

    def test_compress_options_component_save(self):
        """CompressOptions properly writes to tab-specific enable and preset keys."""
        import tkinter as tk
        root = tk.Tk()
        try:
            settings_store = {
                "word_compress_enabled": False,
                "word_compress_preset": "standard",
            }
            co = CompressOptions(
                root,
                settings_ref=settings_store,
                enable_key="word_compress_enabled",
                preset_key="word_compress_preset",
            )
            # Toggle checkbutton
            co._var_en.set(True)
            co._var_preset.set("maximum")
            co._save()

            self.assertTrue(settings_store["word_compress_enabled"])
            self.assertEqual(settings_store["word_compress_preset"], "maximum")

            # Reload settings from disk
            s_disk = load_settings()
            self.assertTrue(s_disk["word_compress_enabled"])
            self.assertEqual(s_disk["word_compress_preset"], "maximum")
        finally:
            root.destroy()

    def test_mtools_app_settings_sync_updates_tab_ui(self):
        """
        Tests whether MToolsApp._on_settings_changed successfully updates tab compression controls.
        Identifies attribute mismatch between MToolsApp (compress_opts) and tabs (_compress_opt).
        """
        import tkinter as tk
        root = tk.Tk()
        try:
            app = markitdown_gui.MToolsApp.__new__(markitdown_gui.MToolsApp)
            app._settings = markitdown_gui.load_settings()
            app._tab_ppt = markitdown_gui.PptToPdfTab(root, settings_ref=app._settings)
            app._tab_word = markitdown_gui.WordToPdfTab(root, settings_ref=app._settings)
            app._tab_merger = markitdown_gui.PdfMergerTab(root, settings_ref=app._settings)
            app._lbl_status = tk.Label(root)

            # Change setting for ppt_compress_enabled to the opposite of current
            current_en = app._tab_ppt._compress_opt._var_en.get()
            target_en = not current_en
            new_settings = dict(app._settings)
            new_settings["ppt_compress_enabled"] = target_en

            app._on_settings_changed(new_settings)

            # The tab UI variable MUST reflect the change
            self.assertEqual(
                app._tab_ppt._compress_opt._var_en.get(),
                target_en,
                "MToolsApp._on_settings_changed failed to update PptToPdfTab._compress_opt (line 3041 checks compress_opts instead of _compress_opt)",
            )
        finally:
            root.destroy()


class TestAtomicWriteAndTmpSemantics(unittest.TestCase):
    """Verifies atomic write semantics and absence of lingering .tmp files."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.settings_file = Path(self.temp_dir) / "settings.json"
        self._orig_settings_file = markitdown_gui.SETTINGS_FILE
        markitdown_gui.SETTINGS_FILE = str(self.settings_file)

    def tearDown(self):
        markitdown_gui.SETTINGS_FILE = self._orig_settings_file
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_single_thread_save_creates_target_and_removes_tmp(self):
        """Normal sequential save creates settings.json and leaves NO .tmp file."""
        tmp_file = self.settings_file.with_name(f"{self.settings_file.name}.tmp")
        s = {"test_key": "test_value"}
        ok = save_settings(s)
        self.assertTrue(ok)
        self.assertTrue(self.settings_file.is_file())
        self.assertFalse(tmp_file.exists(), f"Leftover temporary file found: {tmp_file}")

        # Verify content
        with open(self.settings_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.assertEqual(data.get("test_key"), "test_value")

    def test_repeated_sequential_saves_leave_no_tmp(self):
        """Repeated sequential saves update target and never leak .tmp."""
        tmp_file = self.settings_file.with_name(f"{self.settings_file.name}.tmp")
        for i in range(10):
            save_settings({"seq": i})
            self.assertFalse(tmp_file.exists(), f"Iteration {i} left .tmp file")
            with open(self.settings_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.assertEqual(data["seq"], i)


class TestForbiddenTokens(unittest.TestCase):
    """Audits codebase for forbidden tokens."""

    def test_absence_of_forbidden_tokens_in_gui(self):
        gui_path = Path(markitdown_gui.__file__)
        content = gui_path.read_text(encoding="utf-8")

        tokens = [
            "(FR-2.5)",
            "[KOMPRES]",
            "[BATAL]",
            "Kurang File",
        ]

        found_violations = {}
        lines = content.splitlines()
        for token in tokens:
            matching_lines = [
                (i + 1, line.strip())
                for i, line in enumerate(lines)
                if token in line
            ]
            if matching_lines:
                found_violations[token] = matching_lines

        self.assertEqual(
            found_violations,
            {},
            f"Forbidden tokens found in markitdown_gui.py: {found_violations}",
        )


if __name__ == "__main__":
    unittest.main()
