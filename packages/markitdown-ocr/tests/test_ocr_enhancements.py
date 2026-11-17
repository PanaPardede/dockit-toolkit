"""
Unit tests for OCR optimizations and error handling enhancements:
- Image dimension filtering (skipping < 80x80 or > 15:1 aspect ratio)
- JPEG compression & resizing to <= 1024x1024
- Quota / HTTP 429 error propagation across service and converters
"""

import io
from unittest.mock import MagicMock
from PIL import Image
import pytest

from markitdown_ocr._ocr_service import (
    LLMVisionOCRService,
    OCRQuotaExceededError,
    OCRResult,
    is_quota_or_rate_limit_error,
    should_skip_image,
)
from markitdown_ocr._docx_converter_with_ocr import DocxConverterWithOCR
from markitdown_ocr._pdf_converter_with_ocr import PdfConverterWithOCR
from markitdown_ocr._pptx_converter_with_ocr import PptxConverterWithOCR
from markitdown_ocr._xlsx_converter_with_ocr import XlsxConverterWithOCR
from markitdown import StreamInfo


def _make_test_image(width: int, height: int, mode: str = "RGB") -> io.BytesIO:
    """Helper to create an in-memory PNG image."""
    img = Image.new(mode, (width, height), color=(255, 0, 0) if mode == "RGB" else (255, 0, 0, 128))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


class TestImageDimensionFiltering:
    """Test should_skip_image filtering rule."""

    def test_skip_small_icons_and_bullets(self):
        # Smaller than 80x80 in width or height
        assert should_skip_image(Image.new("RGB", (16, 16))) is True
        assert should_skip_image(Image.new("RGB", (32, 32))) is True
        assert should_skip_image(Image.new("RGB", (64, 64))) is True
        assert should_skip_image(Image.new("RGB", (79, 79))) is True
        assert should_skip_image(Image.new("RGB", (79, 100))) is True
        assert should_skip_image(Image.new("RGB", (100, 79))) is True

    def test_keep_normal_images(self):
        # 80x80 or larger with reasonable aspect ratio
        assert should_skip_image(Image.new("RGB", (80, 80))) is False
        assert should_skip_image(Image.new("RGB", (100, 100))) is False
        assert should_skip_image(Image.new("RGB", (500, 100))) is False
        assert should_skip_image(Image.new("RGB", (1024, 768))) is False

    def test_skip_extreme_aspect_ratio(self):
        # Aspect ratio > 15:1
        assert should_skip_image(Image.new("RGB", (1600, 100))) is True  # 16:1
        assert should_skip_image(Image.new("RGB", (100, 1600))) is True  # 16:1
        assert should_skip_image(Image.new("RGB", (2000, 100))) is True  # 20:1
        # Aspect ratio <= 15:1
        assert should_skip_image(Image.new("RGB", (1500, 100))) is False  # 15:1
        assert should_skip_image(Image.new("RGB", (1200, 100))) is False  # 12:1

    def test_handles_bytes_and_stream(self):
        stream = _make_test_image(50, 50)
        assert should_skip_image(stream) is True
        # Verify stream position is preserved
        assert stream.tell() == 0

        raw_bytes = stream.getvalue()
        assert should_skip_image(raw_bytes) is True

        valid_stream = _make_test_image(200, 200)
        assert should_skip_image(valid_stream) is False
        assert valid_stream.tell() == 0


class TestQuotaErrorDetection:
    """Test is_quota_or_rate_limit_error helper."""

    def test_quota_error_status_code(self):
        exc = Exception("Rate limit hit")
        exc.status_code = 429
        assert is_quota_or_rate_limit_error(exc) is True

    def test_quota_error_response_status_code(self):
        exc = Exception("API error")
        mock_resp = MagicMock()
        mock_resp.status_code = 429
        exc.response = mock_resp
        assert is_quota_or_rate_limit_error(exc) is True

    def test_quota_error_message(self):
        assert is_quota_or_rate_limit_error(Exception("HTTP 429 Too Many Requests")) is True
        assert is_quota_or_rate_limit_error(Exception("You exceeded your current quota")) is True
        assert is_quota_or_rate_limit_error(Exception("RESOURCE_EXHAUSTED: rate limit exceeded")) is True

    def test_non_quota_error(self):
        assert is_quota_or_rate_limit_error(Exception("Invalid file format")) is False
        assert is_quota_or_rate_limit_error(ValueError("Bad parameter")) is False


class TestLLMVisionOCRService:
    """Test OCR service image compression and error propagation."""

    def test_resizing_and_jpeg_compression(self):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_choice = MagicMock()
        mock_choice.message.content = "Detected Text"
        mock_response.choices = [mock_choice]
        mock_client.chat.completions.create.return_value = mock_response

        service = LLMVisionOCRService(client=mock_client, model="gemini-3.6-flash", compress_images=True)

        # Create a large 2000x1500 image
        large_img_stream = _make_test_image(2000, 1500)
        result = service.extract_text(large_img_stream)

        assert result.text == "Detected Text"
        assert mock_client.chat.completions.create.called

        # Verify the image sent was compressed to JPEG
        call_kwargs = mock_client.chat.completions.create.call_args[1]
        msg_content = call_kwargs["messages"][0]["content"]
        image_url = msg_content[1]["image_url"]["url"]
        assert image_url.startswith("data:image/jpeg;base64,")

    def test_quota_error_propagates(self):
        mock_client = MagicMock()
        quota_err = Exception("Error 429: Quota exceeded for model gemini-3.6-flash")
        quota_err.status_code = 429
        mock_client.chat.completions.create.side_effect = quota_err

        service = LLMVisionOCRService(client=mock_client, model="gemini-3.6-flash")
        img_stream = _make_test_image(200, 200)

        with pytest.raises(OCRQuotaExceededError) as exc_info:
            service.extract_text(img_stream)
        assert "Quota" in str(exc_info.value) or "429" in str(exc_info.value)

    def test_normal_error_returns_ocr_result(self):
        mock_client = MagicMock()
        mock_client.chat.completions.create.side_effect = Exception("Malformed image content")

        service = LLMVisionOCRService(client=mock_client, model="gemini-3.6-flash")
        img_stream = _make_test_image(200, 200)

        result = service.extract_text(img_stream)
        assert result.text.startswith("[AI Vision failed on this image")  # visible, not silently dropped
        assert "Malformed image content" in result.error


class QuotaThrowingMockOCRService:
    """Mock service that raises OCRQuotaExceededError."""
    def extract_text(self, image_stream, **kwargs):
        raise OCRQuotaExceededError("Rate limit / quota exceeded (429)")


class TestConverterQuotaPropagation:
    """Ensure all converters propagate OCRQuotaExceededError instead of swallowing it."""

    def test_docx_propagates_quota_error(self, tmp_path):
        from pathlib import Path
        test_file = Path(__file__).parent / "ocr_test_data" / "docx_image_start.docx"
        if not test_file.exists():
            pytest.skip("Test file not found")

        converter = DocxConverterWithOCR()
        svc = QuotaThrowingMockOCRService()

        with open(test_file, "rb") as f:
            with pytest.raises(OCRQuotaExceededError):
                converter.convert(f, StreamInfo(extension=".docx"), ocr_service=svc)

    def test_pdf_propagates_quota_error(self):
        from pathlib import Path
        test_file = Path(__file__).parent / "ocr_test_data" / "pdf_image_start.pdf"
        if not test_file.exists():
            pytest.skip("Test file not found")

        converter = PdfConverterWithOCR()
        svc = QuotaThrowingMockOCRService()

        with open(test_file, "rb") as f:
            with pytest.raises(OCRQuotaExceededError):
                converter.convert(f, StreamInfo(extension=".pdf"), ocr_service=svc)

    def test_pptx_propagates_quota_error(self):
        from pathlib import Path
        test_file = Path(__file__).parent / "ocr_test_data" / "pptx_image_start.pptx"
        if not test_file.exists():
            pytest.skip("Test file not found")

        converter = PptxConverterWithOCR()
        svc = QuotaThrowingMockOCRService()

        with open(test_file, "rb") as f:
            with pytest.raises(OCRQuotaExceededError):
                converter.convert(f, StreamInfo(extension=".pptx"), ocr_service=svc)

    def test_xlsx_propagates_quota_error(self):
        from pathlib import Path
        test_file = Path(__file__).parent / "ocr_test_data" / "xlsx_image_start.xlsx"
        if not test_file.exists():
            pytest.skip("Test file not found")

        converter = XlsxConverterWithOCR()
        svc = QuotaThrowingMockOCRService()

        with open(test_file, "rb") as f:
            with pytest.raises(OCRQuotaExceededError):
                converter.convert(f, StreamInfo(extension=".xlsx"), ocr_service=svc)

    def test_markitdown_core_propagates_quota_error(self):
        """Ensure MarkItDown.convert does not silently fall back to built-ins when quota error occurs."""
        from pathlib import Path
        from markitdown import MarkItDown

        test_file = Path(__file__).parent / "ocr_test_data" / "docx_image_start.docx"
        if not test_file.exists():
            pytest.skip("Test file not found")

        mock_client = MagicMock()
        err = Exception("Rate limit exceeded")
        err.status_code = 429
        mock_client.chat.completions.create.side_effect = err

        md = MarkItDown(
            enable_plugins=True,
            llm_client=mock_client,
            llm_model="gemini-3.6-flash",
        )

        with pytest.raises((OCRQuotaExceededError, Exception)) as exc_info:
            md.convert(str(test_file))
        assert "429" in str(exc_info.value) or "Quota" in str(exc_info.value) or exc_info.type.__name__ == "OCRQuotaExceededError"

    def test_docx_skipped_image_does_not_misalign_subsequent_ocr(self, tmp_path):
        """Ensure small skipped images do not displace or misalign OCR blocks of subsequent real images."""
        from docx import Document
        from PIL import Image

        doc_path = tmp_path / "mixed_images.docx"
        doc = Document()
        doc.add_paragraph("Paragraph 1: Before icon")

        # Small 16x16 icon (should be skipped)
        icon_buf = _make_test_image(16, 16)
        doc.add_picture(icon_buf)

        doc.add_paragraph("Paragraph 2: Between icon and chart")

        # Large 300x200 image (should be OCR-ed)
        chart_buf = _make_test_image(300, 200)
        doc.add_picture(chart_buf)

        doc.add_paragraph("Paragraph 3: After chart")
        doc.save(str(doc_path))

        mock_svc = MagicMock()
        mock_svc.extract_text.return_value = OCRResult(text="CHART_TEXT_CONTENT", backend_used="mock")

        converter = DocxConverterWithOCR(ocr_service=mock_svc)
        with open(doc_path, "rb") as f:
            res = converter.convert(f, StreamInfo(extension=".docx"))

        md = res.text_content
        # 1. OCR text should appear
        assert "CHART_TEXT_CONTENT" in md
        # 2. OCR text must appear AFTER Paragraph 2, NOT before it!
        pos_p2 = md.index("Paragraph 2: Between icon and chart")
        pos_ocr = md.index("CHART_TEXT_CONTENT")
        pos_p3 = md.index("Paragraph 3: After chart")
        assert pos_p2 < pos_ocr < pos_p3
        # 3. Only the large image was extracted, not the 16x16 icon
        assert mock_svc.extract_text.call_count == 1
