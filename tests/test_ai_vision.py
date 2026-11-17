"""AI Vision error handling: fatal errors stop the conversion, transient ones stay visible."""
import io
import threading
import unittest
from unittest.mock import MagicMock

from PIL import Image

from markitdown_ocr import (
    LLMVisionOCRService,
    OCRAuthError,
    OCRCancelledError,
    OCRQuotaExceededError,
)
from markitdown_gui import _find_cause, create_markitdown_instance


def _png():
    buf = io.BytesIO()
    Image.new("RGB", (200, 200), "white").save(buf, format="PNG")
    buf.seek(0)
    return buf


class _HttpError(Exception):
    def __init__(self, status):
        super().__init__(f"HTTP {status}")
        self.status_code = status


def _service(side_effect=None, reply="hello"):
    client = MagicMock()
    if side_effect:
        client.chat.completions.create.side_effect = side_effect
    else:
        client.chat.completions.create.return_value.choices = [
            MagicMock(message=MagicMock(content=reply))]
    return LLMVisionOCRService(client, "m"), client


class TestOcrService(unittest.TestCase):
    def test_ok(self):
        svc, _ = _service(reply=" text ")
        self.assertEqual(svc.extract_text(_png()).text, "text")

    def test_auth_error_is_fatal(self):
        for status in (400, 401, 403, 404):
            svc, _ = _service(_HttpError(status))
            with self.assertRaises(OCRAuthError):
                svc.extract_text(_png())

    def test_quota_is_fatal(self):
        svc, _ = _service(_HttpError(429))
        with self.assertRaises(OCRQuotaExceededError):
            svc.extract_text(_png())

    def test_transient_error_leaves_visible_marker(self):
        svc, _ = _service(TimeoutError("slow"))
        res = svc.extract_text(_png())
        self.assertIn("AI Vision failed", res.text)
        self.assertIn("TimeoutError", res.text)

    def test_cancel_skips_api_call(self):
        svc, client = _service()
        svc.cancel_event = threading.Event()
        svc.cancel_event.set()
        with self.assertRaises(OCRCancelledError):
            svc.extract_text(_png())
        client.chat.completions.create.assert_not_called()

    def test_progress_callback_counts_images(self):
        seen = []
        svc, _ = _service()
        svc.on_progress = seen.append
        svc.extract_text(_png()); svc.extract_text(_png())
        self.assertEqual(seen, [1, 2])


class TestGuiHelpers(unittest.TestCase):
    def test_find_cause_walks_chain_and_attempts(self):
        inner = OCRAuthError("bad key")
        att = MagicMock(exc_info=(type(inner), inner, None))
        outer = Exception("conversion failed")
        outer.attempts = [att]
        self.assertIs(_find_cause(outer, "OCRAuthError"), inner)
        self.assertIs(_find_cause(outer, "OCRFatalError"), inner)  # matches base class
        self.assertIsNone(_find_cause(outer, "OCRCancelledError"))

    def test_instance_passes_cancel_event_and_custom_url(self):
        ev = threading.Event()
        md, err, ai = create_markitdown_instance({
            "ai_vision_enabled": True, "api_key": "x", "provider": "openai",
            "model": "my-model", "base_url": "http://localhost:11434/v1"}, ev)
        self.assertIsNone(err)
        self.assertTrue(ai)
        services = [c.converter.ocr_service for c in md._converters
                    if getattr(c.converter, "ocr_service", None)]
        self.assertTrue(services)
        self.assertIs(services[0].cancel_event, ev)
        self.assertEqual(services[0].model, "my-model")
        self.assertEqual(str(services[0].client.base_url), "http://localhost:11434/v1/")


if __name__ == "__main__":
    unittest.main()
