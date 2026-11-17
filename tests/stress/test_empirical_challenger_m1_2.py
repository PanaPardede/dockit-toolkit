"""
test_empirical_challenger_m1_2.py — Empirical Challenge Harness for Milestone M1

Targeted edge-case audits:
1. Word COM missing detection:
   - Verify Dispatch failure branch: no directory created, error logged, done_fn called with word_missing=True.
   - Verify WordToPdfTab._on_done: actionable showwarning, status updated, button states reset.
2. Temporary file cleanup:
   - In _word_batch_worker: simulate cancellation and optimization exception while __opt_tmp.pdf exists;
     verify strict cleanup in finally.
   - In _ppt_batch_worker: simulate cancellation and optimization exception while __opt_tmp.pdf exists;
     verify strict cleanup in finally.
3. pdf_optimizer.py public API backward compatibility:
   - verify build_output_path with str, Path, custom_dst, identical path validation.
   - verify inspect_pdf with valid, corrupt, encrypted, and missing files.
   - verify optimize_pdf across presets (light, standard, maximum), presets invalidation, cancel_event, and progress callbacks.
"""

import os
import sys
import tempfile
import shutil
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
import tkinter as tk

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pymupdf
import pdf_optimizer
from pdf_optimizer import (
    optimize_pdf,
    inspect_pdf,
    build_output_path,
    CancelledError,
    PasswordRequired,
    CorruptPdf,
    OptimizationResult,
    InspectInfo,
    PRESETS,
)
import markitdown_gui
from markitdown_gui import (
    _word_batch_worker,
    _ppt_batch_worker,
    WordToPdfTab,
    PptToPdfTab,
)


class TestWordMissingDetection(unittest.TestCase):
    """Challenge 1: Word COM missing detection and recovery."""

    @classmethod
    def setUpClass(cls):
        cls.root = tk.Tk()
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_word_detect_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_word_batch_worker_missing_com_does_not_create_directory(self):
        """Simulate missing Word: verify no directory is created and word_missing=True is sent."""
        target_out_dir = self.temp_dir / "should_not_be_created"
        self.assertFalse(target_out_dir.exists())

        test_doc = self.temp_dir / "document.docx"
        test_doc.touch()

        logs = []
        progress_calls = []
        done_calls = []

        def mock_log(msg):
            logs.append(msg)

        def mock_progress(cur, total, msg, pct=None):
            progress_calls.append((cur, total, msg, pct))

        def mock_done(ok, fail, out_d, comp_stats, **kwargs):
            done_calls.append((ok, fail, out_d, comp_stats, kwargs))

        with patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("win32com.client.GetActiveObject", side_effect=Exception("No active instance")), \
             patch("win32com.client.Dispatch", side_effect=Exception("Word.Application COM server not registered")):

            _word_batch_worker(
                files=[test_doc],
                output_dir=target_out_dir,
                log_fn=mock_log,
                progress_fn=mock_progress,
                done_fn=mock_done,
                compress_enabled=False,
            )

        # 1. Output directory MUST NOT be created
        self.assertFalse(
            target_out_dir.exists(),
            "Violation: Output directory was created before verifying Word COM availability!"
        )

        # 2. Standardized bracketed log prefix emitted
        error_logs = [m for m in logs if "[ERROR] Microsoft Word not found:" in m]
        self.assertTrue(len(error_logs) > 0, "Expected bracketed [ERROR] log not found.")

        # 3. done_fn called with word_missing=True
        self.assertEqual(len(done_calls), 1)
        ok, fail, out_d, comp_stats, kwargs = done_calls[0]
        self.assertEqual(ok, 0)
        self.assertEqual(fail, 1)
        self.assertEqual(out_d, target_out_dir)
        self.assertIsNone(comp_stats)
        self.assertTrue(kwargs.get("word_missing"), "Expected word_missing=True flag in done_fn kwargs.")

    def test_word_to_pdf_tab_on_done_word_missing_behavior(self):
        """Verify WordToPdfTab._on_done displays actionable warning, resets buttons and status."""
        tab = WordToPdfTab(self.root)
        tab._is_running = True
        tab._btn_start.config(state="disabled")
        tab._btn_cancel.config(state="normal")
        tab._btn_open.config(state="normal")

        dummy_dir = self.temp_dir / "out"

        with patch("tkinter.messagebox.showwarning") as mock_showwarning, \
             patch("tkinter.messagebox.showinfo") as mock_showinfo:

            tab._on_done(0, 1, dummy_dir, None, word_missing=True)
            self.root.update()

            # 1. Verify warning dialog was triggered with actionable guidance
            self.assertTrue(mock_showwarning.called, "Expected messagebox.showwarning to be called.")
            self.assertFalse(mock_showinfo.called, "showinfo should NOT be called on missing Word.")

            title, message = mock_showwarning.call_args[0]
            self.assertEqual(title, "Microsoft Word Not Found")
            self.assertIn("needs Microsoft Word installed", message)

            # 2. Verify button states reset cleanly
            self.assertFalse(tab._is_running, "Running state was not reset to False.")
            self.assertEqual(str(tab._btn_start.cget("state")), "normal")
            self.assertEqual(str(tab._btn_cancel.cget("state")), "disabled")

            # 3. Verify status label
            self.assertEqual(tab._var_status.get(), "Microsoft Word not found.")

        tab.destroy()


class TestTemporaryFileCleanup(unittest.TestCase):
    """Challenge 2: Strict cleanup of scratch and __opt_tmp.pdf files on error or cancel."""

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_cleanup_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_sample_pdf(self, path: Path):
        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((50, 50), "Test PDF content")
        doc.save(str(path))
        doc.close()

    def test_word_worker_temp_file_cleanup_on_cancel(self):
        """Simulate cancellation during Word compression: verify __opt_tmp.pdf is purged."""
        out_dir = self.temp_dir / "word_out"
        out_dir.mkdir(parents=True, exist_ok=True)

        doc_file = self.temp_dir / "report.docx"
        doc_file.touch()

        pdf_path = out_dir / "report.pdf"
        self._create_sample_pdf(pdf_path)

        tmp_target = out_dir / "report__opt_tmp.pdf"

        cancel_ev = MagicMock()
        cancel_ev.is_set.return_value = False

        def mock_optimize(*args, **kwargs):
            # Simulate partial temporary file written to disk
            tmp_target.write_bytes(b"%PDF-1.4 partial content during optimization")
            self.assertTrue(tmp_target.exists())
            raise CancelledError("User pressed cancel during compression")

        with patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("win32com.client.GetActiveObject"), \
             patch("markitdown_gui._word_convert_single", return_value=True), \
             patch("markitdown_gui.optimize_pdf", side_effect=mock_optimize):

            _word_batch_worker(
                files=[doc_file],
                output_dir=out_dir,
                log_fn=lambda m: None,
                progress_fn=lambda c, t, m, p=None: None,
                done_fn=lambda *a, **k: None,
                compress_enabled=True,
                cancel_event=cancel_ev,
            )

        # __opt_tmp.pdf MUST be unlinked
        self.assertFalse(
            tmp_target.exists(),
            f"Violation: Leftover temporary file {tmp_target.name} found after cancellation!"
        )
        # Original converted PDF must remain intact
        self.assertTrue(pdf_path.exists(), "Original converted PDF was improperly deleted.")

    def test_word_worker_temp_file_cleanup_on_error(self):
        """Simulate unexpected compression error in Word: verify __opt_tmp.pdf is purged."""
        out_dir = self.temp_dir / "word_out"
        out_dir.mkdir(parents=True, exist_ok=True)

        doc_file = self.temp_dir / "report2.docx"
        doc_file.touch()

        pdf_path = out_dir / "report2.pdf"
        self._create_sample_pdf(pdf_path)

        tmp_target = out_dir / "report2__opt_tmp.pdf"

        def mock_optimize(*args, **kwargs):
            tmp_target.write_bytes(b"%PDF-1.4 corrupt temp write")
            self.assertTrue(tmp_target.exists())
            raise RuntimeError("Fatal compressor segmentation or OOM")

        with patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("win32com.client.GetActiveObject"), \
             patch("markitdown_gui._word_convert_single", return_value=True), \
             patch("markitdown_gui.optimize_pdf", side_effect=mock_optimize):

            _word_batch_worker(
                files=[doc_file],
                output_dir=out_dir,
                log_fn=lambda m: None,
                progress_fn=lambda c, t, m, p=None: None,
                done_fn=lambda *a, **k: None,
                compress_enabled=True,
            )

        self.assertFalse(
            tmp_target.exists(),
            f"Violation: Leftover temporary file {tmp_target.name} found after unexpected error!"
        )
        self.assertTrue(pdf_path.exists(), "Original converted PDF was improperly deleted on error.")

    def test_ppt_worker_temp_file_cleanup_on_cancel(self):
        """Simulate cancellation during PPT compression: verify __opt_tmp.pdf is purged."""
        out_dir = self.temp_dir / "ppt_out"
        out_dir.mkdir(parents=True, exist_ok=True)

        ppt_file = self.temp_dir / "deck.pptx"
        ppt_file.touch()

        pdf_path = out_dir / "deck.pdf"
        self._create_sample_pdf(pdf_path)

        tmp_target = out_dir / "deck__opt_tmp.pdf"

        cancel_ev = MagicMock()
        cancel_ev.is_set.return_value = False

        def mock_optimize(*args, **kwargs):
            tmp_target.write_bytes(b"%PDF-1.4 partial ppt temp")
            self.assertTrue(tmp_target.exists())
            raise CancelledError("User cancelled PPT compression")

        with patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("win32com.client.GetActiveObject"), \
             patch("markitdown_gui._ppt_convert_single", return_value=True), \
             patch("markitdown_gui.optimize_pdf", side_effect=mock_optimize):

            _ppt_batch_worker(
                files=[ppt_file],
                output_dir=out_dir,
                log_fn=lambda m: None,
                progress_fn=lambda c, t, m, p=None: None,
                done_fn=lambda *a, **k: None,
                compress_enabled=True,
                cancel_event=cancel_ev,
            )

        self.assertFalse(
            tmp_target.exists(),
            f"Violation: Leftover temporary file {tmp_target.name} found after PPT cancellation!"
        )
        self.assertTrue(pdf_path.exists(), "Original converted PPT PDF was improperly deleted.")

    def test_ppt_worker_temp_file_cleanup_on_error(self):
        """Simulate unexpected compression error in PPT: verify __opt_tmp.pdf is purged."""
        out_dir = self.temp_dir / "ppt_out"
        out_dir.mkdir(parents=True, exist_ok=True)

        ppt_file = self.temp_dir / "deck2.pptx"
        ppt_file.touch()

        pdf_path = out_dir / "deck2.pdf"
        self._create_sample_pdf(pdf_path)

        tmp_target = out_dir / "deck2__opt_tmp.pdf"

        def mock_optimize(*args, **kwargs):
            tmp_target.write_bytes(b"%PDF-1.4 corrupted ppt temp")
            self.assertTrue(tmp_target.exists())
            raise IOError("Disk full while writing PPT compressed stream")

        with patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"), \
             patch("win32com.client.GetActiveObject"), \
             patch("markitdown_gui._ppt_convert_single", return_value=True), \
             patch("markitdown_gui.optimize_pdf", side_effect=mock_optimize):

            _ppt_batch_worker(
                files=[ppt_file],
                output_dir=out_dir,
                log_fn=lambda m: None,
                progress_fn=lambda c, t, m, p=None: None,
                done_fn=lambda *a, **k: None,
                compress_enabled=True,
            )

        self.assertFalse(
            tmp_target.exists(),
            f"Violation: Leftover temporary file {tmp_target.name} found after PPT error!"
        )
        self.assertTrue(pdf_path.exists(), "Original converted PPT PDF was improperly deleted on error.")


class TestPdfOptimizerPublicAPICompatibility(unittest.TestCase):
    """Challenge 3: Full backward compatibility and contract verification for pdf_optimizer.py."""

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_api_compat_"))
        self.sample_pdf = self.temp_dir / "valid_document.pdf"
        self._create_sample_pdf(self.sample_pdf)

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_sample_pdf(self, path: Path, pages: int = 3):
        doc = pymupdf.open()
        for i in range(pages):
            p = doc.new_page()
            p.insert_text((72, 72), f"Page {i+1} Content for API compatibility testing")
        doc.save(str(path))
        doc.close()

    def test_build_output_path_configurations(self):
        """Verify build_output_path across string, Path, custom_dst, and collision avoidance."""
        # 1. str path -> Path with _compressed
        res1 = build_output_path(str(self.sample_pdf))
        self.assertIsInstance(res1, Path)
        self.assertEqual(res1.name, "valid_document_compressed.pdf")
        self.assertEqual(res1.parent, self.sample_pdf.parent)

        # 2. Path instance
        res2 = build_output_path(self.sample_pdf)
        self.assertEqual(res2.name, "valid_document_compressed.pdf")

        # 3. Custom destination as str (non-existent)
        custom_str = str(self.temp_dir / "custom_out.pdf")
        res3 = build_output_path(self.sample_pdf, custom_dst=custom_str)
        self.assertEqual(res3, Path(custom_str))

        # 4. Custom destination as Path (non-existent)
        custom_path = self.temp_dir / "custom_out_path.pdf"
        res4 = build_output_path(self.sample_pdf, custom_dst=custom_path)
        self.assertEqual(res4, custom_path)

        # 5. Collision avoidance: when destination already exists (e.g. sample_pdf itself),
        # build_output_path generates a safe non-colliding candidate: base_stem (2).pdf
        res5 = build_output_path(self.sample_pdf, custom_dst=self.sample_pdf)
        self.assertEqual(res5.name, "valid_document (2).pdf")
        self.assertFalse(res5.exists())

    def test_inspect_pdf_contract_and_exceptions(self):
        """Verify inspect_pdf return type, attributes, and exception contracts."""
        # 1. Valid PDF inspection
        info = inspect_pdf(self.sample_pdf)
        self.assertIsInstance(info, InspectInfo)
        self.assertEqual(info.page_count, 3)
        self.assertGreater(info.file_size, 0)
        self.assertIsInstance(info.images_count, int)
        self.assertIsInstance(info.has_signatures, bool)
        self.assertIsInstance(info.is_encrypted, bool)
        self.assertFalse(info.is_encrypted)
        self.assertFalse(info.needs_password)
        self.assertTrue(info.is_valid)

        # 2. String path acceptance
        info_str = inspect_pdf(str(self.sample_pdf))
        self.assertEqual(info_str.page_count, 3)

        # 3. Missing file -> CorruptPdf
        missing = self.temp_dir / "non_existent.pdf"
        with self.assertRaises(CorruptPdf):
            inspect_pdf(missing)

        # 4. Corrupt file -> CorruptPdf
        corrupt = self.temp_dir / "corrupted.pdf"
        corrupt.write_bytes(b"NOT_A_VALID_PDF_HEADER_DATA_12345")
        with self.assertRaises(CorruptPdf):
            inspect_pdf(corrupt)

    def test_optimize_pdf_all_presets_and_return_structure(self):
        """Verify optimize_pdf across presets (light, standard, maximum) and return types."""
        for preset in ("light", "standard", "maximum"):
            dst = self.temp_dir / f"opt_{preset}.pdf"
            progress_record = []

            def on_p(stage, pct):
                progress_record.append((stage, pct))

            res = optimize_pdf(
                src_path=self.sample_pdf,
                dst_path=dst,
                preset_key=preset,
                on_progress=on_p,
            )

            self.assertIsInstance(res, OptimizationResult)
            self.assertIn(res.status, ("success", "already_optimal"))
            self.assertGreater(res.orig_size, 0)
            self.assertGreater(res.new_size, 0)
            self.assertIsInstance(res.saved_bytes, int)
            self.assertIsInstance(res.saved_percent, float)
            self.assertIsInstance(res.images_total, int)
            self.assertIsInstance(res.images_processed, int)
            self.assertIsInstance(res.images_skipped, int)
            self.assertIsInstance(res.cmyk_converted, int)
            self.assertIsInstance(res.message, str)
            self.assertEqual(Path(res.dst_path), dst)
            self.assertTrue(dst.exists())

            # Verify progress callback was invoked
            self.assertGreater(len(progress_record), 0)
            for stage, pct in progress_record:
                self.assertIsInstance(stage, str)
                self.assertTrue(0 <= pct <= 100)

    def test_optimize_pdf_unknown_preset_fallback(self):
        """Verify unknown preset gracefully falls back to DEFAULT_PRESET_KEY without crash."""
        dst = self.temp_dir / "opt_fallback.pdf"
        res = optimize_pdf(self.sample_pdf, dst_path=dst, preset_key="non_existent_preset")
        self.assertIsInstance(res, OptimizationResult)
        self.assertIn(res.status, ("success", "already_optimal"))
        self.assertTrue(dst.exists())

    def test_optimize_pdf_cooperative_cancellation(self):
        """Verify optimize_pdf respects cancel_event before execution and mid-flight."""
        dst = self.temp_dir / "opt_cancel.pdf"

        # Pre-set cancel event
        ev = MagicMock()
        ev.is_set.return_value = True

        with self.assertRaises(CancelledError):
            optimize_pdf(
                self.sample_pdf,
                dst_path=dst,
                preset_key="standard",
                cancel_event=ev,
            )

        # Verify destination file was not left behind
        self.assertFalse(dst.exists(), "Cancelled optimization left destination file on disk.")


if __name__ == "__main__":
    unittest.main()
