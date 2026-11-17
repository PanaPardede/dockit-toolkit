"""
Enhanced DOCX Converter with OCR support for embedded images.
Extracts images from Word documents and performs OCR while maintaining context.
"""

import io
import re
import sys
from typing import Any, BinaryIO, Optional

from markitdown.converters import HtmlConverter
from markitdown.converter_utils.docx.pre_process import pre_process_docx
from markitdown import DocumentConverterResult, StreamInfo
from markitdown._exceptions import (
    MissingDependencyException,
    MISSING_DEPENDENCY_MESSAGE,
)
from ._ocr_service import (
    LLMVisionOCRService,
    OCRFatalError,
    should_skip_image,
)

# Try loading dependencies
_dependency_exc_info = None
try:
    import mammoth
    from docx import Document
except ImportError:
    _dependency_exc_info = sys.exc_info()

# Placeholder injected into HTML so that mammoth never sees the OCR markers.
# Must be a single token with no special markdown characters.
_PLACEHOLDER = "MARKITDOWNOCRBLOCK{}"

_UNDERLINE_STYLE_MAP = "u => u"


def _read_embedded_style_map(file_stream: BinaryIO) -> Optional[str]:
    """Read the style map embedded in a .docx, if it has one."""
    position = file_stream.tell()
    file_stream.seek(0)
    try:
        return mammoth.read_embedded_style_map(file_stream)
    finally:
        file_stream.seek(position)


class DocxConverterWithOCR(HtmlConverter):
    """
    Enhanced DOCX Converter with OCR support for embedded images.
    Maintains document flow while extracting text from images inline.
    """

    def __init__(self, ocr_service: Optional[LLMVisionOCRService] = None):
        super().__init__()
        self._html_converter = HtmlConverter()
        self.ocr_service = ocr_service

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension == ".docx":
            return True

        if mimetype.startswith(
            "application/vnd.openxmlformats-officedocument.wordprocessingml"
        ):
            return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> DocumentConverterResult:
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".docx",
                    feature="docx",
                )
            ) from _dependency_exc_info[1].with_traceback(
                _dependency_exc_info[2]
            )  # type: ignore[union-attr]

        # Get OCR service if available (from kwargs or instance)
        ocr_service: Optional[LLMVisionOCRService] = (
            kwargs.get("ocr_service") or self.ocr_service
        )

        # Pre-process once, up front, so that every subsequent read sees the
        # repaired archive. pre_process_docx() also fixes .docx files whose ZIP
        # local file headers disagree with the central directory; reading the
        # original stream first would make those raise, or silently yield no
        # images and drop the OCR output.
        pre_process_stream = pre_process_docx(file_stream)

        # Read the embedded style map and combine with any provided style map
        caller_style_map = kwargs.get("style_map")
        embedded_style_map = _read_embedded_style_map(pre_process_stream)

        style_map = "\n".join(
            part
            for part in (
                caller_style_map,
                embedded_style_map,
                _UNDERLINE_STYLE_MAP,
            )
            if part
        )

        if ocr_service:
            ocr_texts: list[str] = []

            def convert_img(image: Any) -> dict[str, str]:
                try:
                    with image.open() as stream:
                        img_bytes = stream.read()
                    if should_skip_image(img_bytes):
                        return {"src": "", "data-skip": "true"}
                    img_stream = io.BytesIO(img_bytes)
                    ocr_result = ocr_service.extract_text(img_stream)
                    if ocr_result.text.strip():
                        idx = len(ocr_texts)
                        ocr_texts.append(ocr_result.text.strip())
                        return {"src": "", "data-ocr-idx": str(idx)}
                    return {"src": "", "data-skip": "true"}
                except OCRFatalError:
                    raise
                except Exception:
                    return {"src": "", "data-skip": "true"}

            # 1. Convert DOCX → HTML via mammoth with document-order image OCR
            pre_process_stream.seek(0)
            html_result = mammoth.convert_to_html(
                pre_process_stream,
                style_map=style_map,
                include_embedded_style_map=False,
                convert_image=mammoth.images.img_element(convert_img),
            ).value

            # 2. Cleanly strip skipped images
            html_clean = re.sub(r'<img[^>]*data-skip="true"[^>]*>', "", html_result)

            # 3. Replace OCR images with placeholders
            html_with_placeholders = re.sub(
                r'<img[^>]*data-ocr-idx="(\d+)"[^>]*>',
                lambda m: f"<p>{_PLACEHOLDER.format(m.group(1))}</p>",
                html_clean,
            )

            # 4. Convert HTML → markdown
            md_result = self._html_converter.convert_string(
                html_with_placeholders, **kwargs
            )
            md = md_result.markdown

            # 5. Swap placeholders for the actual OCR blocks (post-conversion
            #    so * and _ are never escaped by the markdown converter).
            for i, raw_text in enumerate(ocr_texts):
                placeholder = _PLACEHOLDER.format(i)
                ocr_block = f"*[Image OCR]\n{raw_text}\n[End OCR]*"
                md = md.replace(placeholder, ocr_block)

            return DocumentConverterResult(markdown=md)
        else:
            # Standard conversion without OCR
            pre_process_stream.seek(0)
            return self._html_converter.convert_string(
                mammoth.convert_to_html(
                    pre_process_stream,
                    style_map=style_map,
                    include_embedded_style_map=False,
                ).value,
                **kwargs,
            )

    def _extract_and_ocr_images(
        self, file_stream: BinaryIO, ocr_service: LLMVisionOCRService
    ) -> dict[str, str]:
        """
        Extract images from DOCX and OCR them.

        Returns:
            Dict mapping image relationship IDs to raw OCR text (no markers).
        """
        ocr_map = {}

        try:
            file_stream.seek(0)
            doc = Document(file_stream)

            for rel in doc.part.rels.values():
                if "image" in rel.target_ref.lower():
                    try:
                        image_bytes = rel.target_part.blob
                        if should_skip_image(image_bytes):
                            continue
                        image_stream = io.BytesIO(image_bytes)
                        ocr_result = ocr_service.extract_text(image_stream)

                        if ocr_result.text.strip():
                            # Store raw text only — markers added later
                            ocr_map[rel.rId] = ocr_result.text.strip()

                    except OCRFatalError:
                        raise
                    except Exception:
                        continue

        except OCRFatalError:
            raise
        except Exception:
            pass

        return ocr_map

    def _inject_placeholders(
        self, html: str, ocr_map: dict[str, str]
    ) -> tuple[str, list[str]]:
        """
        Replace <img> tags with numbered placeholder tokens.

        Returns:
            (html_with_placeholders, ordered list of raw OCR texts)
        """
        if not ocr_map:
            return html, []

        ocr_texts = list(ocr_map.values())
        used: list[int] = []

        def replace_img(match: re.Match) -> str:  # type: ignore[type-arg]
            for i in range(len(ocr_texts)):
                if i not in used:
                    used.append(i)
                    return f"<p>{_PLACEHOLDER.format(i)}</p>"
            return ""  # remove image if all OCR texts already used

        result = re.sub(r"<img[^>]*>", replace_img, html)

        # Any OCR texts that had no matching <img> tag go at the end
        for i in range(len(ocr_texts)):
            if i not in used:
                result += f"<p>{_PLACEHOLDER.format(i)}</p>"

        return result, ocr_texts
