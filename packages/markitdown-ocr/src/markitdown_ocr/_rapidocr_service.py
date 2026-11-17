"""
Offline OCR service using RapidOCR (ONNX Runtime, CPU). Same interface as
LLMVisionOCRService, so every OCR converter accepts either one.
"""

import logging
import threading
from typing import Any, BinaryIO

from ._ocr_service import OCRCancelledError, OCRResult

_engine = None
_engine_lock = threading.Lock()


def rapidocr_available() -> bool:
    try:
        import rapidocr  # noqa: F401
        return True
    except ImportError:
        return False


def _get_engine():
    """Load the ONNX models once per process (about 1 s); later calls reuse them."""
    global _engine
    with _engine_lock:
        if _engine is None:
            logging.getLogger("RapidOCR").setLevel(logging.WARNING)
            from rapidocr import RapidOCR
            _engine = RapidOCR()
        return _engine


def group_into_lines(boxes, txts) -> str:
    """Join text fragments that sit on the same visual row, top to bottom, left to right.

    boxes: per fragment, 4 (x, y) corner points. RapidOCR returns label/value pairs
    such as "Report ID:" and "MEDRPT-1" as separate fragments; this puts them back
    on one line.
    """
    items = []
    for box, txt in zip(boxes, txts):
        txt = (txt or "").strip()
        if not txt:
            continue
        ys = [p[1] for p in box]
        xs = [p[0] for p in box]
        items.append((min(ys), max(ys), min(xs), txt))
    if not items:
        return ""
    items.sort(key=lambda it: ((it[0] + it[1]) / 2, it[2]))

    lines: list[list[tuple]] = []
    for it in items:
        mid = (it[0] + it[1]) / 2
        if lines:
            top = min(i[0] for i in lines[-1])
            bottom = max(i[1] for i in lines[-1])
            if top <= mid <= bottom:  # vertical centre falls inside the current row
                lines[-1].append(it)
                continue
        lines.append([it])
    return "\n".join("  ".join(i[3] for i in sorted(line, key=lambda i: i[2])) for line in lines)


class RapidOCRService:
    """OCR on this computer. No API key, no network."""

    def __init__(self, cancel_event: Any = None, on_progress: Any = None) -> None:
        self.cancel_event = cancel_event
        self.on_progress = on_progress
        self.images_sent = 0

    def extract_text(self, image_stream: BinaryIO, prompt: str | None = None, **kwargs: Any) -> OCRResult:
        if self.cancel_event is not None and self.cancel_event.is_set():
            raise OCRCancelledError("Cancelled by user")
        self.images_sent += 1
        if self.on_progress is not None:
            self.on_progress(self.images_sent)
        try:
            image_stream.seek(0)
            result = _get_engine()(image_stream.read())
            boxes = getattr(result, "boxes", None)
            txts = getattr(result, "txts", None)
            if boxes is None or not txts:
                return OCRResult(text="", backend_used="rapidocr")
            return OCRResult(text=group_into_lines(boxes, txts), backend_used="rapidocr")
        except OCRCancelledError:
            raise
        except Exception as e:
            reason = f"{e.__class__.__name__}: {str(e)[:120]}"
            return OCRResult(text=f"[Offline OCR failed on this image - {reason}]",
                             backend_used="rapidocr", error=str(e))
        finally:
            image_stream.seek(0)
