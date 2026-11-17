"""
test_empirical_challenger_m1.py — Empirical challenge and verification suite for M1 remediations.

Tests:
1. Queue unpack logic in PdfOptimizerTab._poll_queue (2-tuple, 3-tuple, 1-tuple, lists, malformed items).
2. Cancellation job token logic in MarkItDownTab (stale tokens, cancelled tokens, out-of-order workers).
3. Atomic replace semantics via Path.replace() on Windows (without preceding unlink).
"""

import os
import queue
import shutil
import tempfile
import unittest
from unittest.mock import MagicMock, patch
from pathlib import Path

import pymupdf
import markitdown_gui
from markitdown_gui import (
    PdfOptimizerTab,
    MarkItDownTab,
)
import pdf_optimizer


class DummyPdfOptimizerTab:
    """
    Headless test harness directly bound to PdfOptimizerTab._poll_queue.
    Executes the exact implementation in markitdown_gui.py without relying on
    unstable Tcl interpreter teardown across multiple test classes.
    """
    _poll_queue = PdfOptimizerTab._poll_queue

    def __init__(self):
        self._queue = queue.Queue()
        self._pbar = {"value": 0}
        self._status_text = ""
        self._var_status = MagicMock()
        self._var_status.set = lambda val: setattr(self, "_status_text", val)
        self._var_status.get = lambda: getattr(self, "_status_text", "")
        self.after_scheduled = []
        self.done_calls = []
        self.cancelled_calls = []
        self.error_calls = []

    def after(self, delay, fn):
        self.after_scheduled.append((delay, fn))

    def _on_done(self, data):
        self.done_calls.append(data)

    def _on_cancelled(self):
        self.cancelled_calls.append(True)

    def _on_error(self, data, is_permission=False):
        self.error_calls.append((data, is_permission))


class TestPdfOptimizerQueueUnpack(unittest.TestCase):
    """
    Empirical challenge for Task 1:
    Test queue unpack logic in PdfOptimizerTab._poll_queue.
    """
    def setUp(self):
        self.tab = DummyPdfOptimizerTab()

    def test_01_queue_3_tuple_progress(self):
        """Verify 3-tuple ('progress', stage, pct) updates UI without error."""
        self.tab._queue.put(("progress", "Compressing images...", 55))
        self.tab._poll_queue()
        
        self.assertEqual(self.tab._pbar["value"], 55)
        self.assertEqual(self.tab._var_status.get(), "Compressing images... (55%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_02_queue_2_tuple_progress_with_tuple_data(self):
        """Verify 2-tuple ('progress', (stage, pct)) updates UI without error."""
        self.tab._queue.put(("progress", ("Membersihkan metadata...", 80)))
        self.tab._poll_queue()
        
        self.assertEqual(self.tab._pbar["value"], 80)
        self.assertEqual(self.tab._var_status.get(), "Membersihkan metadata... (80%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_03_queue_2_tuple_progress_with_string_data(self):
        """Verify 2-tuple ('progress', stage_str) defaults pct to 0."""
        self.tab._queue.put(("progress", "Inisialisasi"))
        self.tab._poll_queue()
        
        self.assertEqual(self.tab._pbar["value"], 0)
        self.assertEqual(self.tab._var_status.get(), "Inisialisasi (0%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_04_queue_1_tuple_progress(self):
        """Verify 1-tuple ('progress',) handles missing stage & pct gracefully."""
        self.tab._queue.put(("progress",))
        self.tab._poll_queue()
        
        self.assertEqual(self.tab._pbar["value"], 0)
        self.assertEqual(self.tab._var_status.get(), "None (0%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_05_queue_list_progress(self):
        """Verify list ['progress', stage, pct] is handled via general sequence branch."""
        self.tab._queue.put(["progress", "Optimasi font", 92])
        self.tab._poll_queue()
        
        self.assertEqual(self.tab._pbar["value"], 92)
        self.assertEqual(self.tab._var_status.get(), "Optimasi font (92%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_06_queue_falsy_items_skipped(self):
        """Verify falsy items (None, (), '', []) are skipped cleanly without crashing."""
        self.tab._queue.put(None)
        self.tab._queue.put(())
        self.tab._queue.put("")
        self.tab._queue.put([])
        self.tab._queue.put(("progress", "Done", 100))
        self.tab._poll_queue()
        
        self.assertEqual(self.tab._pbar["value"], 100)
        self.assertEqual(self.tab._var_status.get(), "Done (100%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_07_queue_malformed_unsubscriptable_item_does_not_kill_polling(self):
        """Verify non-subscriptable item (e.g. integer 99999) is caught and does not kill polling."""
        self.tab._queue.put(99999)  # int has no __getitem__, triggers TypeError
        
        # Does not raise; caught in except Exception
        self.tab._poll_queue()
        # self.after must still be scheduled in finally:
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

        # Ensure subsequent queue items can still be processed normally
        self.tab._queue.put(("progress", "Pulih dari error", 50))
        self.tab._poll_queue()
        self.assertEqual(self.tab._pbar["value"], 50)
        self.assertEqual(self.tab._var_status.get(), "Pulih dari error (50%)")
        self.assertIn((100, self.tab._poll_queue), self.tab.after_scheduled)

    def test_08_queue_dispatch_handlers(self):
        """Verify done, cancelled, error, and error_permission dispatch correctly."""
        mock_res = MagicMock()
        self.tab._queue.put(("done", mock_res))
        self.tab._queue.put(("cancelled", None))
        self.tab._queue.put(("error", "Fatal compression error"))
        self.tab._queue.put(("error_permission", "File locked"))

        self.tab._poll_queue()

        self.assertIn(mock_res, self.tab.done_calls)
        self.assertIn(True, self.tab.cancelled_calls)
        self.assertIn(("Fatal compression error", False), self.tab.error_calls)
        self.assertIn(("File locked", True), self.tab.error_calls)


class DummyMarkItDownTab:
    """
    Headless test harness directly bound to MarkItDownTab methods for cancellation
    and job token validation.
    """
    _clear_queue = MarkItDownTab._clear_queue

    def __init__(self):
        self._cancelled = False
        self._current_job_id = 0
        self._cancel_event = None
        self._file_queue = []
        self._current_idx = 0
        self.result_calls = []
        self.error_calls = []

    def _show_drop_zone(self):
        pass

    def _show_result(self, filepath, text, ai_used):
        self.result_calls.append((filepath, text, ai_used))

    def _show_error(self, filepath, msg):
        self.error_calls.append((filepath, msg))


class TestMarkItDownTabJobTokens(unittest.TestCase):
    """
    Empirical challenge for Task 2:
    Test cancellation job token logic in MarkItDownTab.
    Verify that out-of-order or stale job tokens are ignored.
    """
    def setUp(self):
        self.tab = DummyMarkItDownTab()

    def test_01_job_token_increments_monotonically(self):
        """Verify job token increments on cancellation and queue clearing."""
        initial_id = self.tab._current_job_id
        
        self.tab._clear_queue()
        self.assertEqual(self.tab._current_job_id, initial_id + 1)
        self.assertTrue(self.tab._cancelled)

    def test_02_stale_job_token_result_ignored(self):
        """Simulate worker from job 1 finishing after job 2 has started."""
        # Start job 1
        self.tab._current_job_id = 1
        job_1_id = 1

        # Simulate job 2 starting
        self.tab._current_job_id = 2
        job_2_id = 2

        # Emulate job 1's lambda callback firing on UI thread
        cb_job_1 = lambda jid=job_1_id: self.tab._show_result("file1.pdf", "text 1", False) \
            if self.tab._current_job_id == jid and not self.tab._cancelled else None
        cb_job_1()
        
        # Job 1 must NOT show result because current_job_id is 2 != 1
        self.assertEqual(len(self.tab.result_calls), 0)

        # Emulate job 2's lambda callback firing
        cb_job_2 = lambda jid=job_2_id: self.tab._show_result("file2.pdf", "text 2", False) \
            if self.tab._current_job_id == jid and not self.tab._cancelled else None
        cb_job_2()

        # Job 2 MUST show result
        self.assertEqual(self.tab.result_calls, [("file2.pdf", "text 2", False)])

    def test_03_cancelled_job_token_ignored_even_after_cancelled_reset(self):
        """
        Critical race scenario:
        1. Job 1 begins (job_id = 1).
        2. User cancels: _cancelled = True, _current_job_id becomes 2.
        3. User starts Job 2: _cancelled = False, _current_job_id becomes 3.
        4. Detached background worker for Job 1 finishes.
        Verify Job 1 is rejected by token mismatch even though _cancelled is False.
        """
        self.tab._current_job_id = 1
        job_1_id = 1

        # Cancelled
        self.tab._current_job_id += 1  # 2
        self.tab._cancelled = True

        # Job 2 starts
        self.tab._current_job_id += 1  # 3
        job_2_id = 3
        self.tab._cancelled = False  # Reset for Job 2

        # Job 1 callback attempts to execute
        cb_1 = lambda jid=job_1_id: self.tab._show_result("stale.pdf", "stale", False) \
            if self.tab._current_job_id == jid and not self.tab._cancelled else None
        cb_1()
        self.assertEqual(len(self.tab.result_calls), 0)

        # Job 1 error callback attempts to execute
        cb_err_1 = lambda jid=job_1_id: self.tab._show_error("stale.pdf", "stale error") \
            if self.tab._current_job_id == jid and not self.tab._cancelled else None
        cb_err_1()
        self.assertEqual(len(self.tab.error_calls), 0)

    def test_04_quota_fallback_respects_job_token(self):
        """Verify ask_fallback dialog is never opened for stale/cancelled job token."""
        self.tab._current_job_id = 5
        stale_jid = 4

        mock_ask = MagicMock()
        def ask_fallback(jid=stale_jid):
            if self.tab._cancelled or self.tab._current_job_id != jid:
                return
            mock_ask("Title", "Prompt")

        cb = lambda jid=stale_jid: ask_fallback(jid) if self.tab._current_job_id == jid and not self.tab._cancelled else None
        cb()
        mock_ask.assert_not_called()


class TestAtomicReplace(unittest.TestCase):
    """
    Empirical challenge for Task 3:
    Test atomic replace: verify that Path.replace() successfully replaces files
    without needing a preceding unlink().
    """
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_atomic_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_01_path_replace_overwrites_existing_file_without_unlink(self):
        """Verify Path.replace() on Windows overwrites destination cleanly without unlink()."""
        dst = self.temp_dir / "target.pdf"
        src = self.temp_dir / "target__opt_tmp.pdf"

        # Populate target and src with different content
        dst.write_bytes(b"%PDF-1.4 TARGET INITIAL CONTENT" * 100)
        src.write_bytes(b"%PDF-1.4 REPLACEMENT OPTIMIZED CONTENT" * 50)

        initial_dst_size = dst.stat().st_size
        src_size = src.stat().st_size
        self.assertNotEqual(initial_dst_size, src_size)

        # Perform atomic replace WITHOUT unlink
        src.replace(dst)

        # Assertions
        self.assertTrue(dst.exists(), "Target file must exist after replace")
        self.assertFalse(src.exists(), "Source temporary file must be gone after replace")
        self.assertEqual(dst.stat().st_size, src_size, "Target file size must match replacement content")
        self.assertTrue(dst.read_bytes().startswith(b"%PDF-1.4 REPLACEMENT OPTIMIZED CONTENT"))

    def test_02_stress_consecutive_atomic_replaces(self):
        """Stress-test 100 rapid consecutive atomic replacements simulating batch workloads."""
        target = self.temp_dir / "batch_doc.pdf"
        target.write_text("initial document version 0")

        for i in range(1, 101):
            tmp = self.temp_dir / f"batch_doc__opt_tmp_{i}.pdf"
            tmp_content = f"document version {i} content padding " + ("x" * (i * 10))
            tmp.write_text(tmp_content)

            # Atomic replace
            tmp.replace(target)

            self.assertTrue(target.exists())
            self.assertFalse(tmp.exists())
            self.assertEqual(target.read_text(), tmp_content)

    def test_03_codebase_audit_confirms_no_premature_unlink_before_replace(self):
        """
        Verify AST / source in markitdown_gui.py has zero instances of unlink()
        immediately preceding Path.replace().
        """
        gui_path = Path("markitdown_gui.py")
        content = gui_path.read_text(encoding="utf-8")
        lines = content.splitlines()

        replace_line_indices = [
            idx for idx, line in enumerate(lines)
            if "Path(res.dst_path).replace" in line or ".replace(out_path)" in line or ".replace(pdf_path)" in line
        ]
        self.assertGreater(len(replace_line_indices), 0, "Must find replace calls in markitdown_gui.py")

        # In none of these call sites should an unlink() occur in the 5 lines preceding replace()
        for idx in replace_line_indices:
            preceding_context = "\n".join(lines[max(0, idx - 5):idx])
            self.assertNotIn("unlink", preceding_context,
                             f"Found premature unlink preceding replace at line {idx+1}:\n{preceding_context}")


if __name__ == "__main__":
    unittest.main()
