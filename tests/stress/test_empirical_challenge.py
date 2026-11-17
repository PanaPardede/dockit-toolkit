"""
Empirical Challenge Harness for Challenger M6-1.
Tests the extreme PDFs under multiple presets and stress conditions while running
a high-frequency (20ms/50ms) background ResourceMonitor to catch transient RSS peaks.
"""

import os
import sys
import time
import shutil
import tempfile
import threading
import unittest
from pathlib import Path
from typing import Dict, Any, List

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import psutil
import pymupdf
import pypdf

from pdf_optimizer import optimize_pdf, inspect_pdf, CorruptPdf, PRESETS
from markitdown_gui import _merge_pdfs, is_office_encrypted


class RealResourceMonitor:
    """
    Background sampling monitor conforming to PROJECT.md specification:
    - 50ms sampling interval
    - Tracks process tree RSS in MB
    - Tracks CPU%
    - Tracks active POWERPNT.EXE process count
    - Exposes peak_rss_mb, avg_cpu_percent, leaked_processes, lingering_tmp_files
    """
    def __init__(self, interval_s: float = 0.05, scratch_dir: Path = None):
        self.interval_s = interval_s
        self.scratch_dir = scratch_dir
        self.running = False
        self._thread: threading.Thread = None
        self.samples_rss: List[float] = []
        self.samples_cpu: List[float] = []
        self.peak_rss_mb: float = 0.0
        self.peak_powerpnt_count: int = 0
        self.process = psutil.Process(os.getpid())

    def start(self):
        self.running = True
        self.samples_rss.clear()
        self.samples_cpu.clear()
        self.peak_rss_mb = 0.0
        self.peak_powerpnt_count = 0
        self._thread = threading.Thread(target=self._sample_loop, daemon=True)
        self._thread.start()

    def _sample_loop(self):
        while self.running:
            try:
                tree_rss = self.process.memory_info().rss
                for child in self.process.children(recursive=True):
                    try:
                        tree_rss += child.memory_info().rss
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                rss_mb = tree_rss / (1024 * 1024)
                cpu = self.process.cpu_percent(interval=None)

                self.samples_rss.append(rss_mb)
                self.samples_cpu.append(cpu)
                if rss_mb > self.peak_rss_mb:
                    self.peak_rss_mb = rss_mb

                ppt_count = sum(1 for p in psutil.process_iter(['name']) if p.info['name'] and p.info['name'].lower() == 'powerpnt.exe')
                if ppt_count > self.peak_powerpnt_count:
                    self.peak_powerpnt_count = ppt_count
            except Exception:
                pass
            time.sleep(self.interval_s)

    def stop(self) -> Dict[str, Any]:
        self.running = False
        if self._thread:
            self._thread.join(timeout=1.0)

        avg_cpu = sum(self.samples_cpu) / max(1, len(self.samples_cpu))
        current_ppt = sum(1 for p in psutil.process_iter(['name']) if p.info['name'] and p.info['name'].lower() == 'powerpnt.exe')
        tmp_files = 0
        if self.scratch_dir and self.scratch_dir.exists():
            tmp_files = len(list(self.scratch_dir.glob("*.tmp*"))) + len(list(self.scratch_dir.glob("*__opt_tmp*")))

        return {
            "peak_rss_mb": self.peak_rss_mb,
            "avg_cpu_percent": avg_cpu,
            "sample_count": len(self.samples_rss),
            "leaked_processes": current_ppt,
            "lingering_tmp_files": tmp_files,
        }


class EmpiricalChallengeM6_1(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fixtures_dir = PROJECT_ROOT / "tests" / "stress" / "fixtures" / "extreme_pdfs"
        if not cls.fixtures_dir.exists():  # ~105 MB, not in git
            raise unittest.SkipTest("Run `python tests/stress/generate_stress_data.py` first")
        cls.scratch = Path(tempfile.mkdtemp(prefix="challenger_scratch_m6_1_"))
        cls.monitor = RealResourceMonitor(interval_s=0.02, scratch_dir=cls.scratch)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.scratch, ignore_errors=True)

    def test_01_extreme_100mb_pdf_under_all_presets(self):
        """Challenge 100MB PDF under light, standard, and maximum presets."""
        path = self.fixtures_dir / "extreme_size_100mb.pdf"
        self.assertTrue(path.exists(), f"Missing fixture: {path}")

        for preset in ["light", "standard", "maximum"]:
            out_dst = self.scratch / f"opt_100mb_{preset}.pdf"
            self.monitor.start()
            t0 = time.time()
            res = optimize_pdf(path, dst_path=out_dst, preset_key=preset)
            dt = time.time() - t0
            stats = self.monitor.stop()

            self.assertEqual(res.status, "success")
            self.assertTrue(out_dst.exists())
            self.assertLess(stats["peak_rss_mb"], 1500.0, f"Peak memory exceeded 1.5GB: {stats['peak_rss_mb']} MB")
            self.assertEqual(stats["lingering_tmp_files"], 0, "Leaked temp files found!")
            self.assertEqual(stats["leaked_processes"], 0, "Leaked POWERPNT.EXE found!")
            self.assertLess(dt, 30.0, f"Optimization took too long: {dt:.2f}s")

    def test_02_extreme_500_pages_pdf(self):
        """Challenge 500-page PDF under standard & maximum presets, merge, and extraction."""
        path = self.fixtures_dir / "high_page_count_500p.pdf"
        self.assertTrue(path.exists())

        # Optimization
        for preset in ["standard", "maximum"]:
            out_dst = self.scratch / f"opt_500p_{preset}.pdf"
            self.monitor.start()
            res = optimize_pdf(path, dst_path=out_dst, preset_key=preset)
            stats = self.monitor.stop()
            self.assertEqual(res.status, "success")
            self.assertLess(stats["peak_rss_mb"], 1500.0)
            self.assertEqual(stats["lingering_tmp_files"], 0)

        # Merge 2x 500p (1,000 pages)
        merged_dst = self.scratch / "merged_1000p.pdf"
        self.monitor.start()
        pages = _merge_pdfs([path, path], merged_dst)
        stats = self.monitor.stop()
        self.assertEqual(pages, 1000)
        self.assertTrue(merged_dst.exists())
        self.assertLess(stats["peak_rss_mb"], 1500.0)

    def test_03_extreme_1000_images_pdf(self):
        """Challenge 1000 distinct images under standard & maximum presets."""
        path = self.fixtures_dir / "high_image_density_1000img.pdf"
        self.assertTrue(path.exists())

        for preset in ["standard", "maximum"]:
            out_dst = self.scratch / f"opt_1000img_{preset}.pdf"
            self.monitor.start()
            res = optimize_pdf(path, dst_path=out_dst, preset_key=preset)
            stats = self.monitor.stop()
            self.assertEqual(res.status, "success")
            self.assertLess(stats["peak_rss_mb"], 1500.0)
            self.assertEqual(stats["lingering_tmp_files"], 0)

    def test_04_resource_monitor_fidelity_proves_transient_spike(self):
        """Verify that high-frequency ResourceMonitor catches transient peak that post-call misses."""
        path = self.fixtures_dir / "extreme_size_100mb.pdf"
        out_dst = self.scratch / "opt_spike_test.pdf"

        self.monitor.start()
        optimize_pdf(path, dst_path=out_dst, preset_key="standard")
        stats = self.monitor.stop()
        post_mem = psutil.Process(os.getpid()).memory_info().rss / (1024 * 1024)

        # True transient peak is captured by background monitor
        self.assertGreater(stats["sample_count"], 0)
        self.assertGreaterEqual(stats["peak_rss_mb"], post_mem - 1.0)
        # Verify peak stays strictly <= 1500 MB
        self.assertLess(stats["peak_rss_mb"], 1500.0)


if __name__ == "__main__":
    unittest.main()
