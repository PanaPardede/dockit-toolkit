"""Offline OCR (RapidOCR): mode selection, when OCR runs, row grouping, real recognition."""
import io
import tempfile
import threading
import unittest
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from markitdown_gui import (
    _OfflineOcrConverter,
    create_markitdown_instance,
    needs_ocr,
    ocr_mode,
)
from markitdown_ocr import OCRCancelledError, RapidOCRService, rapidocr_available
from markitdown_ocr._rapidocr_service import group_into_lines

ROOT = Path(__file__).resolve().parent.parent
TEST_FILES = ROOT / "packages" / "markitdown" / "tests" / "test_files"
SCANNED_PDF = TEST_FILES / "MEDRPT-2024-PAT-3847_medical_report_scan.pdf"
TEXT_PDF = TEST_FILES / "REPAIR-2022-INV-001_multipage.pdf"


def _box(x0, y0, x1, y1):
    return [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]


class TestGrouping(unittest.TestCase):
    def test_same_row_fragments_join_left_to_right(self):
        boxes = [_box(200, 10, 300, 30), _box(10, 12, 100, 28), _box(10, 50, 120, 70)]
        txts = ["MEDRPT-1", "Report ID:", "Next line"]
        self.assertEqual(group_into_lines(boxes, txts), "Report ID:  MEDRPT-1\nNext line")

    def test_empty(self):
        self.assertEqual(group_into_lines([], []), "")
        self.assertEqual(group_into_lines([_box(0, 0, 5, 5)], ["  "]), "")


class TestModeAndRouting(unittest.TestCase):
    def test_ocr_mode(self):
        self.assertEqual(ocr_mode({}), "offline")
        self.assertEqual(ocr_mode({"ocr_engine": "default"}), "offline")  # legacy value
        self.assertEqual(ocr_mode({"ocr_engine": "off"}), "off")
        self.assertEqual(ocr_mode({"ai_vision_enabled": True, "api_key": "k"}), "ai")
        self.assertEqual(ocr_mode({"ai_vision_enabled": True, "api_key": " "}), "offline")

    def test_needs_ocr(self):
        self.assertTrue(needs_ocr("photo.JPG"))
        self.assertFalse(needs_ocr("report.docx"))
        self.assertTrue(needs_ocr(SCANNED_PDF))
        self.assertFalse(needs_ocr(TEXT_PDF))

    @unittest.skipUnless(rapidocr_available(), "rapidocr not installed")
    def test_instance_per_mode(self):
        md, err, ai = create_markitdown_instance({"ocr_engine": "offline"})
        self.assertIsNone(err)
        self.assertFalse(ai)
        self.assertIsInstance(md, _OfflineOcrConverter)
        md, _, _ = create_markitdown_instance({"ocr_engine": "off"})
        self.assertNotIsInstance(md, _OfflineOcrConverter)

    def test_cancel_before_recognition(self):
        ev = threading.Event()
        ev.set()
        with self.assertRaises(OCRCancelledError):
            RapidOCRService(cancel_event=ev).extract_text(io.BytesIO(b"x"))


@unittest.skipUnless(rapidocr_available(), "rapidocr not installed")
class TestRealOcr(unittest.TestCase):
    def test_image_file(self):
        img = Image.new("RGB", (900, 200), "white")
        font = ImageFont.truetype("arial.ttf", 48)
        ImageDraw.Draw(img).text((30, 60), "Invoice number 48213", fill="black", font=font)
        tmp = Path(tempfile.mkdtemp()) / "invoice.png"
        img.save(tmp)
        seen = []
        md, _, _ = create_markitdown_instance({"ocr_engine": "offline"}, on_progress=seen.append)
        text = md.convert(str(tmp)).markdown
        self.assertIn("48213", text)
        self.assertEqual(seen, [1])

    def test_scanned_pdf(self):
        md, _, _ = create_markitdown_instance({"ocr_engine": "offline"})
        text = md.convert(str(SCANNED_PDF)).markdown
        self.assertIn("SEATTLE GENERAL HOSPITAL", text)
        self.assertIn("Report ID:  MEDRPT-2024-PAT-3847", text)  # same-row grouping

    def test_text_pdf_keeps_builtin_tables(self):
        md, _, _ = create_markitdown_instance({"ocr_engine": "offline"})
        text = md.convert(str(TEXT_PDF)).markdown
        self.assertIn("ZAVA AUTO REPAIR", text)
        self.assertIn("|", text)  # built-in converter's Markdown tables, not OCR text


if __name__ == "__main__":
    unittest.main()
