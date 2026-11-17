"""
OCR Service Layer for MarkItDown
Provides LLM Vision-based image text extraction.
"""

import base64
import io
from typing import Any, BinaryIO
from dataclasses import dataclass

from markitdown import StreamInfo


class OCRFatalError(Exception):
    """An OCR error that must stop the whole conversion (retrying other images is pointless)."""


class OCRQuotaExceededError(OCRFatalError):
    """Raised when LLM Vision OCR encounters a quota limit or rate limit (HTTP 429)."""


class OCRAuthError(OCRFatalError):
    """Raised when the API key, model name or endpoint is rejected (HTTP 400/401/403/404)."""


class OCRCancelledError(OCRFatalError):
    """Raised when the caller's cancel_event is set."""


def is_auth_or_config_error(exc: Exception) -> bool:
    """True for errors that will fail identically on every image (bad key, model or URL)."""
    status_code = getattr(exc, "status_code", None)
    if status_code is None:
        status_code = getattr(getattr(exc, "response", None), "status_code", None)
    if status_code in (400, 401, 403, 404):
        return True
    name = exc.__class__.__name__
    return name in ("AuthenticationError", "PermissionDeniedError", "NotFoundError", "APIConnectionError")


def is_quota_or_rate_limit_error(exc: Exception) -> bool:
    """Check if an exception is a rate limit or quota exceeded error (HTTP 429)."""
    if isinstance(exc, OCRQuotaExceededError):
        return True

    status_code = getattr(exc, "status_code", None)
    if status_code == 429:
        return True

    response = getattr(exc, "response", None)
    if response is not None and getattr(response, "status_code", None) == 429:
        return True

    cls_name = exc.__class__.__name__.lower()
    if "ratelimit" in cls_name or "quota" in cls_name:
        return True

    msg = str(exc).lower()
    keywords = ("429", "quota", "rate_limit", "rate limit", "resource_exhausted", "resourceexhausted")
    return any(kw in msg for kw in keywords)


def should_skip_image(
    image: Any,
    min_width: int = 80,
    min_height: int = 80,
    max_aspect_ratio: float = 15.0,
) -> bool:
    """
    Check if an image should be skipped for OCR based on dimensions.

    Skips images smaller than 80x80 pixels (icons, bullets, avatars)
    and images with extreme aspect ratio (> 15:1, lines/borders).

    Args:
        image: PIL Image, bytes, or BinaryIO stream
        min_width: Minimum width in pixels (default: 80)
        min_height: Minimum height in pixels (default: 80)
        max_aspect_ratio: Maximum aspect ratio allowed (default: 15.0)

    Returns:
        True if the image should be skipped, False otherwise.
    """
    try:
        from PIL import Image

        w, h = 0, 0
        if isinstance(image, Image.Image):
            w, h = image.size
        elif isinstance(image, (bytes, bytearray)):
            with Image.open(io.BytesIO(image)) as img:
                w, h = img.size
        elif hasattr(image, "read") and hasattr(image, "seek"):
            pos = image.tell()
            try:
                image.seek(0)
                with Image.open(image) as img:
                    w, h = img.size
            finally:
                image.seek(pos)
        else:
            return False

        if w <= 0 or h <= 0:
            return True

        if w < min_width or h < min_height:
            return True

        aspect_ratio = max(w / h, h / w)
        if aspect_ratio > max_aspect_ratio:
            return True

        return False
    except Exception:
        return False


@dataclass
class OCRResult:
    """Result from OCR extraction."""

    text: str
    confidence: float | None = None
    backend_used: str | None = None
    error: str | None = None


class LLMVisionOCRService:
    """OCR service using LLM vision models (OpenAI-compatible)."""

    def __init__(
        self,
        client: Any,
        model: str,
        default_prompt: str | None = None,
        compress_images: bool = True,
        cancel_event: Any = None,
        on_progress: Any = None,
    ) -> None:
        """
        Initialize LLM Vision OCR service.

        Args:
            client: OpenAI-compatible client
            model: Model name (e.g., 'gpt-4o', 'gemini-3.6-flash')
            default_prompt: Default prompt for OCR extraction
            compress_images: Whether to resize and compress images before API calls
        """
        self.client = client
        self.model = model
        self.default_prompt = default_prompt or (
            "Extract all text from this image. "
            "Return ONLY the extracted text, maintaining the original "
            "layout and order. Do not add any commentary or description."
        )
        self.compress_images = compress_images
        self.cancel_event = cancel_event
        self.on_progress = on_progress  # called with the 1-based image count before each request
        self.images_sent = 0

    def extract_text(
        self,
        image_stream: BinaryIO,
        prompt: str | None = None,
        stream_info: StreamInfo | None = None,
        compress_images: bool | None = None,
        **kwargs: Any,
    ) -> OCRResult:
        """Extract text using LLM vision."""
        if self.client is None:
            return OCRResult(
                text="",
                backend_used="llm_vision",
                error="LLM client not configured",
            )

        if self.cancel_event is not None and self.cancel_event.is_set():
            raise OCRCancelledError("Cancelled by user")
        self.images_sent += 1
        if self.on_progress is not None:
            self.on_progress(self.images_sent)

        should_compress = self.compress_images if compress_images is None else compress_images

        try:
            image_stream.seek(0)
            data_uri: str | None = None

            if should_compress:
                try:
                    from PIL import Image

                    img = Image.open(image_stream)
                    # Convert to RGB (composite alpha on white background if transparent)
                    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                        bg = Image.new("RGB", img.size, (255, 255, 255))
                        if img.mode != "RGBA":
                            img = img.convert("RGBA")
                        bg.paste(img, mask=img.split()[3])
                        img = bg
                    elif img.mode != "RGB":
                        img = img.convert("RGB")

                    # Max 1600px: 1024px made small print on full A4 scans unreadable
                    resample_filter = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)
                    img.thumbnail((1600, 1600), resample_filter)

                    # Compress to JPEG (quality 80, optimize=True)
                    out_stream = io.BytesIO()
                    img.save(out_stream, format="JPEG", quality=80, optimize=True)
                    out_stream.seek(0)
                    b64_jpeg = base64.b64encode(out_stream.read()).decode("utf-8")
                    data_uri = f"data:image/jpeg;base64,{b64_jpeg}"
                except Exception:
                    data_uri = None

            if not data_uri:
                image_stream.seek(0)
                content_type: str | None = None
                if stream_info:
                    content_type = stream_info.mimetype

                if not content_type:
                    try:
                        from PIL import Image

                        image_stream.seek(0)
                        img = Image.open(image_stream)
                        fmt = img.format.lower() if img.format else "png"
                        content_type = f"image/{fmt}"
                    except Exception:
                        content_type = "image/png"

                image_stream.seek(0)
                base64_image = base64.b64encode(image_stream.read()).decode("utf-8")
                data_uri = f"data:{content_type};base64,{base64_image}"

            actual_prompt = prompt or self.default_prompt
            messages = [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": actual_prompt},
                        {
                            "type": "image_url",
                            "image_url": {"url": data_uri},
                        },
                    ],
                }
            ]

            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    timeout=90.0,  # dense full-page scans regularly take 20-60 s
                )
            except TypeError:
                # Fallback for clients/mocks that do not accept timeout kwarg
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                )

            text = response.choices[0].message.content
            return OCRResult(
                text=text.strip() if text else "",
                backend_used="llm_vision",
            )
        except OCRFatalError:
            raise
        except Exception as e:
            if is_quota_or_rate_limit_error(e):
                raise OCRQuotaExceededError(f"AI Vision Quota/Rate Limit Exceeded: {e}") from e
            if is_auth_or_config_error(e):
                raise OCRAuthError(f"AI Vision request rejected: {e}") from e
            # Transient failure (timeout, 5xx): keep going, but leave a visible marker
            # instead of silently dropping this image's content.
            reason = f"{e.__class__.__name__}: {str(e)[:120]}"
            return OCRResult(
                text=f"[AI Vision failed on this image - {reason}]",
                backend_used="llm_vision",
                error=str(e),
            )
        finally:
            image_stream.seek(0)
