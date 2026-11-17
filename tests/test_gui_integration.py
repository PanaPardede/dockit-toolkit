"""
test_gui_integration.py — Tes integrasi komponen GUI (CompressOptions, PdfOptimizerTab, PptToPdfTab, PdfMergerTab).
"""

import os
import shutil
import tempfile
import unittest
import tkinter as tk
from pathlib import Path

import pymupdf
from PIL import Image

import markitdown_gui
from markitdown_gui import (
    CompressOptions,
    PdfOptimizerTab,
    PptToPdfTab,
    WordToPdfTab,
    PdfMergerTab,
    _merge_pdfs,
    _word_batch_worker,
)
import pdf_optimizer


class TestGuiIntegration(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = tk.Tk()
        cls.root.withdraw()  # Sembunyikan window saat testing

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_gui_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_sample_pdf(self, name="sample.pdf", pages=2) -> Path:
        p = self.temp_dir / name
        doc = pymupdf.open()
        for i in range(pages):
            page = doc.new_page()
            page.insert_text((50, 50), f"Halaman {i+1}")
        doc.save(str(p))
        doc.close()
        return p

    def test_compress_options_widget(self):
        """AC-03: Toggle checkbox & preset combobox tersimpan dan terpanggil."""
        settings = {"pdf_compress_enabled": False, "pdf_compress_preset": "standard"}
        changed = []

        def on_change(s):
            changed.append(dict(s))

        opt = CompressOptions(self.root, settings_ref=settings, on_settings_change=on_change)
        self.assertFalse(opt.is_enabled())
        self.assertEqual(opt.get_preset_key(), "standard")

        # Toggle aktif
        opt._var_en.set(True)
        opt._on_toggle()
        self.assertTrue(opt.is_enabled())
        self.assertTrue(settings["pdf_compress_enabled"])
        self.assertGreater(len(changed), 0)

        # Switch preset to Maximum
        opt._combo.set("Maximum (smallest file)")
        opt._on_select()
        self.assertEqual(opt.get_preset_key(), "maximum")
        self.assertEqual(settings["pdf_compress_preset"], "maximum")

        opt.destroy()

    def test_pdf_merger_tab_has_compress_options(self):
        """AC-03, AC-04: PdfMergerTab memiliki opsi kompresi dan default nonaktif."""
        tab = PdfMergerTab(self.root)
        self.assertTrue(hasattr(tab, "_compress_opt"))
        self.assertIsInstance(tab._compress_opt, CompressOptions)
        self.assertFalse(tab._compress_opt.is_enabled())
        tab.destroy()

    def test_ppt_to_pdf_tab_has_compress_options(self):
        """AC-03, AC-04: PptToPdfTab memiliki opsi kompresi dan default nonaktif."""
        tab = PptToPdfTab(self.root)
        self.assertTrue(hasattr(tab, "_compress_opt"))
        self.assertIsInstance(tab._compress_opt, CompressOptions)
        self.assertFalse(tab._compress_opt.is_enabled())
        tab.destroy()

    def test_pdf_optimizer_tab_ui_and_load(self):
        """AC-01, AC-02, AC-05: Tab 4 memuat file, info diekstrak, output path terbentuk."""
        pdf_file = self._create_sample_pdf("dokumen_penting.pdf", pages=3)
        tab = PdfOptimizerTab(self.root)

        # Standard is the default
        self.assertEqual(tab._var_preset.get(), "standard")

        # Muat file
        tab._load_file(pdf_file)
        self.assertEqual(tab._input_path, pdf_file)
        self.assertIsNotNone(tab._inspect_info)
        self.assertEqual(tab._inspect_info.page_count, 3)
        self.assertIn("dokumen_penting.pdf", tab._lbl_fname.cget("text"))

        # Output path default: <nama>_compressed.pdf
        out = Path(tab._var_out.get())
        self.assertEqual(out.name, "dokumen_penting_compressed.pdf")

        tab.destroy()

    def test_merger_logic_with_and_without_compression(self):
        """AC-04: Merge berhasil baik saat kompresi aktif maupun mati."""
        p1 = self._create_sample_pdf("part1.pdf", pages=2)
        p2 = self._create_sample_pdf("part2.pdf", pages=3)
        merged = self.temp_dir / "merged.pdf"

        # 1. Tanpa kompresi (standar)
        pages = _merge_pdfs([p1, p2], merged)
        self.assertEqual(pages, 5)
        self.assertTrue(merged.exists())

        # 2. Dengan kompresi
        res = pdf_optimizer.optimize_pdf(merged, preset_key="standard")
        self.assertIn(res.status, ("success", "already_optimal"))

    def test_ppt_batch_progress_and_summary_logic(self):
        """FR-2.5: Batch PPT progress 70/30 weighting dan ringkasan kompresi."""
        from unittest.mock import patch

        f1 = self.temp_dir / "deck1.pptx"
        f2 = self.temp_dir / "deck2.pptx"
        f1.touch()
        f2.touch()

        out_dir = self.temp_dir / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        pdf1 = out_dir / "deck1.pdf"
        pdf2 = out_dir / "deck2.pdf"
        p_dummy = self._create_sample_pdf("dummy.pdf", pages=2)
        shutil.copy(p_dummy, pdf1)
        shutil.copy(p_dummy, pdf2)

        progress_calls = []
        def mock_progress(cur, total, msg, pct=None):
            progress_calls.append((cur, total, msg, pct))

        done_calls = []
        def mock_done(ok, fail, out_d, comp_stats):
            done_calls.append((ok, fail, out_d, comp_stats))

        with patch("markitdown_gui._ppt_convert_single", return_value=True), \
             patch("win32com.client.GetActiveObject"), \
             patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"):

            markitdown_gui._ppt_batch_worker(
                [f1, f2], out_dir,
                log_fn=lambda m: None,
                progress_fn=mock_progress,
                done_fn=mock_done,
                compress_enabled=True,
                compress_preset="standard"
            )

        self.assertEqual(len(done_calls), 1)
        ok, fail, out_d, comp_stats = done_calls[0]
        self.assertEqual(ok, 2)
        self.assertEqual(fail, 0)
        self.assertIsNotNone(comp_stats)
        self.assertGreaterEqual(comp_stats["comp_ok"] + comp_stats["comp_opt"], 2)

        # Cek bobot progress: untuk file 1 dari 2, konversi 70% dari 50% = 35%
        percentages = [p[3] for p in progress_calls if p[3] is not None]
        self.assertIn(35, percentages)

    def test_word_to_pdf_tab_has_compress_options(self):
        """Verify WordToPdfTab instantiates CompressOptions, sets enable_key='word_compress_enabled' and preset_key='word_compress_preset', and exposes _compress_opt and compress_opts."""
        tab = WordToPdfTab(self.root)
        self.assertTrue(hasattr(tab, "_compress_opt"))
        self.assertTrue(hasattr(tab, "compress_opts"))
        self.assertIs(tab._compress_opt, tab.compress_opts)
        self.assertIsInstance(tab._compress_opt, CompressOptions)
        self.assertEqual(tab._compress_opt._enable_key, "word_compress_enabled")
        self.assertEqual(tab._compress_opt._preset_key, "word_compress_preset")
        self.assertFalse(tab._compress_opt.is_enabled())
        tab.destroy()

        # Test with custom settings reference
        settings = {"word_compress_enabled": True, "word_compress_preset": "maximum"}
        tab2 = WordToPdfTab(self.root, settings_ref=settings)
        self.assertTrue(tab2._compress_opt.is_enabled())
        self.assertEqual(tab2._compress_opt.get_preset_key(), "maximum")
        tab2.destroy()

    def test_word_batch_progress_and_missing_word_logic(self):
        """Test _word_batch_worker behavior with mocked win32com.client.Dispatch raising an exception (missing Word).
        Verify that word_missing=True is passed to done_fn, no directory is created prematurely,
        buttons reset to idle state, and warning is shown."""
        from unittest.mock import patch

        f1 = self.temp_dir / "doc1.docx"
        f1.touch()
        out_dir = self.temp_dir / "uncreated_out_dir"

        done_calls = []
        def mock_done(ok, fail, out_d, comp_stats, **kwargs):
            done_calls.append((ok, fail, out_d, comp_stats, kwargs))

        log_calls = []
        def mock_log(msg):
            log_calls.append(msg)

        # Mock Word COM failure (both GetActiveObject and Dispatch raise exception)
        with patch("win32com.client.GetActiveObject", side_effect=Exception("Word not running")), \
             patch("win32com.client.Dispatch", side_effect=Exception("Word not installed")), \
             patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"):

            _word_batch_worker(
                [f1], out_dir,
                log_fn=mock_log,
                progress_fn=lambda *a, **k: None,
                done_fn=mock_done,
                compress_enabled=False,
            )

        # 1. Verify done_fn received word_missing=True
        self.assertEqual(len(done_calls), 1)
        ok, fail, out_d, comp_stats, kwargs = done_calls[0]
        self.assertEqual(ok, 0)
        self.assertEqual(fail, 1)
        self.assertEqual(out_d, out_dir)
        self.assertIsNone(comp_stats)
        self.assertTrue(kwargs.get("word_missing"))

        # 2. Verify no directory created prematurely
        self.assertFalse(out_dir.exists())

        # 3. Verify WordToPdfTab UI handles word_missing=True properly
        tab = WordToPdfTab(self.root)
        tab._is_running = True
        tab._btn_start.config(state="disabled")
        tab._btn_cancel.config(state="normal")

        with patch("tkinter.messagebox.showwarning") as mock_warn:
            tab._on_done(0, 1, out_dir, None, word_missing=True)
            self.root.update()

            self.assertFalse(tab._is_running)
            self.assertEqual(tab._btn_start.cget("state"), "normal")
            self.assertEqual(tab._btn_cancel.cget("state"), "disabled")
            self.assertEqual(tab._var_status.get(), "Microsoft Word not found.")
            mock_warn.assert_called_once()
            args, _ = mock_warn.call_args
            self.assertIn("Microsoft Word Not Found", args[0])

        tab.destroy()

    def test_word_batch_conversion_and_compression_logic(self):
        """Test batch progression with mocked Word conversion and compression,
        verifying 70/30 weighted progress updates and closure variable stability."""
        from unittest.mock import patch, MagicMock

        f1 = self.temp_dir / "report1.docx"
        f2 = self.temp_dir / "report2.docx"
        f1.touch()
        f2.touch()

        out_dir = self.temp_dir / "word_out"
        out_dir.mkdir(parents=True, exist_ok=True)
        pdf1 = out_dir / "report1.pdf"
        pdf2 = out_dir / "report2.pdf"

        # Create dummy source PDFs that conversion would generate
        p_dummy = self._create_sample_pdf("dummy_doc.pdf", pages=2)
        shutil.copy(p_dummy, pdf1)
        shutil.copy(p_dummy, pdf2)

        progress_calls = []
        def mock_progress(cur, total, msg, pct=None):
            progress_calls.append((cur, total, msg, pct))

        done_calls = []
        def mock_done(ok, fail, out_d, comp_stats, **kwargs):
            done_calls.append((ok, fail, out_d, comp_stats, kwargs))

        recorded_closure_args = []
        def mock_optimize_pdf(src_path, dst_path=None, preset_key=None, on_progress=None, cancel_event=None, **kwargs):
            if on_progress:
                on_progress("Testing downsampling...", 50)
                defaults = on_progress.__defaults__ if hasattr(on_progress, "__defaults__") else ()
                recorded_closure_args.append(defaults)
            if dst_path:
                shutil.copy(src_path, dst_path)
            res = MagicMock()
            res.status = "success"
            res.dst_path = dst_path
            res.orig_size = 10000
            res.new_size = 6000
            res.saved_percent = 40.0
            return res

        with patch("markitdown_gui._word_convert_single", return_value=True), \
             patch("win32com.client.GetActiveObject"), \
             patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("markitdown_gui.optimize_pdf", side_effect=mock_optimize_pdf):

            _word_batch_worker(
                [f1, f2], out_dir,
                log_fn=lambda m: None,
                progress_fn=mock_progress,
                done_fn=mock_done,
                compress_enabled=True,
                compress_preset="standard",
            )

        self.assertEqual(len(done_calls), 1)
        ok, fail, out_d, comp_stats, _ = done_calls[0]
        self.assertEqual(ok, 2)
        self.assertEqual(fail, 0)
        self.assertIsNotNone(comp_stats)
        self.assertEqual(comp_stats["comp_ok"], 2)

        # 70/30 weighting check
        pcts = [p[3] for p in progress_calls if p[3] is not None]
        self.assertIn(0, pcts)
        self.assertIn(35, pcts)
        self.assertIn(42, pcts)
        self.assertIn(50, pcts)
        self.assertIn(85, pcts)
        self.assertIn(92, pcts)
        self.assertIn(100, pcts)

        # Verify closure variable stability: default args captured file_num=1 then file_num=2
        self.assertEqual(len(recorded_closure_args), 2)
        self.assertEqual(recorded_closure_args[0], (0.0, 50.0, 1, 2))
        self.assertEqual(recorded_closure_args[1], (50.0, 50.0, 2, 2))

    def test_word_cancellation_logic(self):
        """Test cooperative cancellation during batch processing,
        verifying that cancel_event stops the loop and cleans up intermediate scratch files."""
        import threading
        from unittest.mock import patch

        f1 = self.temp_dir / "canc1.docx"
        f2 = self.temp_dir / "canc2.docx"
        f3 = self.temp_dir / "canc3.docx"
        f1.touch()
        f2.touch()
        f3.touch()

        out_dir = self.temp_dir / "canc_out"
        out_dir.mkdir(parents=True, exist_ok=True)
        pdf1 = out_dir / "canc1.pdf"
        p_dummy = self._create_sample_pdf("canc_dummy.pdf", pages=1)
        shutil.copy(p_dummy, pdf1)

        cancel_ev = threading.Event()
        tmp_target = out_dir / "canc1__opt_tmp.pdf"

        done_calls = []
        def mock_done(ok, fail, out_d, comp_stats, **kwargs):
            done_calls.append((ok, fail, out_d, comp_stats, kwargs))

        def mock_optimize_with_cancel(src_path, dst_path=None, **kwargs):
            if dst_path:
                Path(dst_path).touch()
            cancel_ev.set()
            raise pdf_optimizer.CancelledError("User pressed cancel")

        with patch("markitdown_gui._word_convert_single", return_value=True), \
             patch("win32com.client.GetActiveObject"), \
             patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("markitdown_gui.optimize_pdf", side_effect=mock_optimize_with_cancel):

            _word_batch_worker(
                [f1, f2, f3], out_dir,
                log_fn=lambda m: None,
                progress_fn=lambda *a, **k: None,
                done_fn=mock_done,
                compress_enabled=True,
                compress_preset="standard",
                cancel_event=cancel_ev,
            )

        # 1. Verify done_fn received cancelled=True
        self.assertEqual(len(done_calls), 1)
        ok, fail, out_d, comp_stats, kwargs = done_calls[0]
        self.assertTrue(kwargs.get("cancelled"))

        # 2. Verify loop stopped early (did not convert file 2 or 3)
        self.assertFalse((out_dir / "canc2.pdf").exists())
        self.assertFalse((out_dir / "canc3.pdf").exists())

        # 3. Verify intermediate scratch file is cleaned up
        self.assertFalse(tmp_target.exists())


if __name__ == "__main__":
    unittest.main()
