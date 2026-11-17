"""Image files (photos or scans of documents) to text via an OCR service."""

from typing import Any, BinaryIO

from markitdown import DocumentConverter, DocumentConverterResult, StreamInfo

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


class ImageConverterWithOCR(DocumentConverter):
    def __init__(self, ocr_service: Any):
        super().__init__()
        self.ocr_service = ocr_service

    def accepts(self, file_stream: BinaryIO, stream_info: StreamInfo, **kwargs: Any) -> bool:
        ext = (stream_info.extension or "").lower()
        mime = (stream_info.mimetype or "").lower()
        return ext in IMAGE_EXTENSIONS or (mime.startswith("image/") and "svg" not in mime)

    def convert(self, file_stream: BinaryIO, stream_info: StreamInfo, **kwargs: Any) -> DocumentConverterResult:
        result = self.ocr_service.extract_text(file_stream)
        return DocumentConverterResult(markdown=result.text.strip())
