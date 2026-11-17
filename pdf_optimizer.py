"""
pdf_optimizer.py - PDF compression engine for DocKit.
No Tkinter imports, so it can be tested and used on its own.
"""

import os
import io
import gc
import time
import logging
import logging.handlers
from pathlib import Path
from dataclasses import dataclass
from typing import Callable, Optional, Dict, List, Tuple

import pymupdf

logger = logging.getLogger("MarkItDown.PdfOptimizer")


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING SETUP
# ══════════════════════════════════════════════════════════════════════════════

def get_log_dir() -> Path:
    """Return the log folder: %LOCALAPPDATA%\\MarkItDown\\logs\\."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base_dir = Path(local_app_data)
    else:
        base_dir = Path.home() / "AppData" / "Local"
    log_dir = base_dir / "MarkItDown" / "logs"
    return log_dir


def setup_optimizer_logging() -> Optional[Path]:
    """Set up a rotating log handler (1 MB, 3 backups) in AppData."""
    try:
        log_dir = get_log_dir()
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / "markitdown.log"

        app_logger = logging.getLogger("MarkItDown")
        # Avoid adding the handler twice
        if not any(isinstance(h, logging.handlers.RotatingFileHandler) for h in app_logger.handlers):
            handler = logging.handlers.RotatingFileHandler(
                str(log_file),
                maxBytes=1024 * 1024,  # 1 MB
                backupCount=3,
                encoding="utf-8"
            )
            formatter = logging.Formatter(
                "[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            )
            handler.setFormatter(formatter)
            app_logger.addHandler(handler)
            app_logger.setLevel(logging.INFO)
            logger.info("Logging initialized at %s", log_file)
        return log_file
    except Exception as e:
        logger.warning("Could not set up file logging: %s", e)
        return None


# ══════════════════════════════════════════════════════════════════════════════
# DATA CLASSES & EXCEPTIONS
# ══════════════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class Preset:
    key: str
    name: str
    description: str
    target_dpi: Optional[int]       # None: images untouched
    dpi_threshold: Optional[int]    # Downsample only above this DPI
    jpeg_quality: Optional[int]     # JPEG quality
    subset_fonts: bool              # Font subsetting


PRESETS: Dict[str, Preset] = {
    "light": Preset(
        key="light",
        name="Light",
        description="Removes duplicate objects and streams; images untouched (lossless)",
        target_dpi=None,
        dpi_threshold=None,
        jpeg_quality=None,
        subset_fonts=False,
    ),
    "standard": Preset(
        key="standard",
        name="Standard (recommended)",
        description="Balanced (150 DPI, quality 80, text stays vector)",
        target_dpi=150,
        dpi_threshold=180,
        jpeg_quality=80,
        subset_fonts=True,
    ),
    "maximum": Preset(
        key="maximum",
        name="Maximum",
        description="Smallest file for email / upload limits (110 DPI, quality 65)",
        target_dpi=110,
        dpi_threshold=130,
        jpeg_quality=65,
        subset_fonts=True,
    ),
}

DEFAULT_PRESET_KEY = "standard"


class OptimizerError(Exception):
    """Base exception for the PDF optimizer."""
    pass


class PasswordRequired(OptimizerError):
    """Raised when the PDF needs a password to open."""
    pass


class InvalidPassword(OptimizerError):
    """Raised when the given password is wrong."""
    pass


class CorruptPdf(OptimizerError):
    """Raised when the PDF structure is damaged and cannot be read."""
    pass


class CancelledError(OptimizerError):
    """Raised when the user cancels."""
    pass


class DigitalSignatureWarning(OptimizerError):
    """Raised when the PDF has a digital signature that compression would break."""
    pass


@dataclass
class InspectInfo:
    src_path: str
    file_size: int
    page_count: int
    is_encrypted: bool
    needs_password: bool
    has_signatures: bool
    images_count: int
    is_valid: bool
    error_message: Optional[str] = None


@dataclass
class OptimizationResult:
    status: str             # "success", "already_optimal", "cancelled", "skipped", "error"
    src_path: str
    dst_path: Optional[str]
    orig_size: int
    new_size: int
    saved_bytes: int
    saved_percent: float
    images_total: int
    images_processed: int
    images_skipped: int
    cmyk_converted: int
    message: str


# ══════════════════════════════════════════════════════════════════════════════
# INSPECTION & HELPER UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def format_size(bytes_val: int) -> str:
    """Format bytes as a human-readable string (KB, MB, GB)."""
    if bytes_val < 0:
        return "0 B"
    if bytes_val < 1024:
        return f"{bytes_val} B"
    if bytes_val < 1024 * 1024:
        return f"{bytes_val / 1024:.1f} KB"
    if bytes_val < 1024 * 1024 * 1024:
        return f"{bytes_val / (1024 * 1024):.2f} MB"
    return f"{bytes_val / (1024 * 1024 * 1024):.2f} GB"


def has_digital_signatures(doc: pymupdf.Document) -> bool:
    """Detect whether the PDF contains a digital signature."""
    try:
        flags = doc.get_sigflags()
        if flags > 0:
            return True
    except Exception:
        pass

    try:
        sig_type = getattr(pymupdf, "PDF_WIDGET_TYPE_SIGNATURE", 6)
        for page in doc:
            for widget in page.widgets():
                if widget.field_type == sig_type:
                    return True
    except Exception:
        pass
    return False


def to_extended_path(p: Path | str) -> str:
    """Handle long paths (>260 chars on Windows) with the prefix \\\\?\\ (E13)."""
    path_str = str(Path(p).resolve())
    if os.name == "nt" and len(path_str) >= 260 and not path_str.startswith("\\\\?\\"):
        return f"\\\\?\\{path_str}"
    return path_str


def inspect_pdf(src_path: str | Path, password: Optional[str] = None) -> InspectInfo:
    """Open the PDF and collect metadata without optimising it."""
    path = Path(src_path)
    if not path.is_file():
        raise CorruptPdf(f"File not found: {path}")

    file_size = path.stat().st_size
    doc = None
    try:
        doc = pymupdf.open(to_extended_path(path))
    except Exception as e:
        logger.error("Could not open file %s: %s", path.name, e)
        raise CorruptPdf(f"File cannot be read or is damaged: {e}")

    try:
        is_encrypted = doc.is_encrypted
        needs_password = bool(doc.needs_pass)

        if needs_password:
            if password:
                if not doc.authenticate(password):
                    return InspectInfo(
                        src_path=str(path),
                        file_size=file_size,
                        page_count=0,
                        is_encrypted=True,
                        needs_password=True,
                        has_signatures=False,
                        images_count=0,
                        is_valid=False,
                        error_message="Password salah."
                    )
            else:
                return InspectInfo(
                    src_path=str(path),
                    file_size=file_size,
                    page_count=0,
                    is_encrypted=True,
                    needs_password=True,
                    has_signatures=False,
                    images_count=0,
                    is_valid=False,
                    error_message="The document is password-protected."
                )

        page_count = len(doc)
        has_sigs = has_digital_signatures(doc)

        # Count unique images
        unique_xrefs = set()
        for page in doc:
            for img in page.get_images(full=True):
                unique_xrefs.add(img[0])

        return InspectInfo(
            src_path=str(path),
            file_size=file_size,
            page_count=page_count,
            is_encrypted=is_encrypted,
            needs_password=False,
            has_signatures=has_sigs,
            images_count=len(unique_xrefs),
            is_valid=True,
            error_message=None
        )
    finally:
        if doc:
            doc.close()


def build_output_path(src_path: str | Path, custom_dst: Optional[str | Path] = None) -> Path:
    """
    Build an output path that never overwrites an existing file.
    Pattern: <name>_compressed.pdf, then _compressed (2).pdf, and so on.
    """
    src = Path(src_path)
    if custom_dst:
        cd = Path(custom_dst)
        if cd.is_dir():
            target_dir = cd
            base_stem = f"{src.stem}_compressed"
        else:
            # User picked a specific output file
            if not cd.exists():
                return cd
            target_dir = cd.parent
            base_stem = cd.stem
    else:
        target_dir = src.parent
        base_stem = f"{src.stem}_compressed"

    candidate = target_dir / f"{base_stem}.pdf"
    if not candidate.exists():
        return candidate

    idx = 2
    while True:
        candidate = target_dir / f"{base_stem} ({idx}).pdf"
        if not candidate.exists():
            return candidate
        idx += 1


def validate_output(orig_page_count: int, dst_path: str | Path, password: Optional[str] = None) -> bool:
    """Check the output opens and has the same page count."""
    p = Path(dst_path)
    if not p.is_file() or p.stat().st_size == 0:
        return False
    doc = None
    try:
        doc = pymupdf.open(to_extended_path(p))
        if doc.needs_pass:
            if not password or not doc.authenticate(password):
                return False
        if len(doc) != orig_page_count:
            return False
        return True
    except Exception:
        return False
    finally:
        if doc:
            doc.close()


# ══════════════════════════════════════════════════════════════════════════════
# IMAGE COLLECTION & RECOMPRESSION ENGINE
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class ImageCandidate:
    xref: int
    first_page_idx: int
    width: int
    height: int
    bpc: int
    colorspace_name: str
    filter_name: str
    has_smask: bool
    max_display_width_pt: float
    effective_dpi: float
    is_bilevel: bool


def collect_images(doc: pymupdf.Document) -> Dict[int, ImageCandidate]:
    """
    Collect every unique image in the PDF with its effective DPI.
    Uses the largest placement (lowest DPI) so the image never looks blurry
    wherever it is shown.
    """
    found: Dict[int, ImageCandidate] = {}

    for pno in range(len(doc)):
        page = doc[pno]
        for img in page.get_images(full=True):
            xref = img[0]
            smask = img[1]
            w = img[2]
            h = img[3]
            bpc = img[4]
            cs = img[5]
            filt = img[8] if len(img) > 8 else ""

            # Check mask / bilevel
            is_smask = bool(smask)
            is_bilevel = (bpc == 1) or (filt in ("JBIG2Decode", "CCITTFaxDecode", "/JBIG2Decode", "/CCITTFaxDecode"))

            # Find the displayed size on this page
            rects = page.get_image_rects(xref)
            current_max_w = 0.0
            for r in rects:
                if r.width > current_max_w:
                    current_max_w = float(r.width)

            if xref not in found:
                found[xref] = ImageCandidate(
                    xref=xref,
                    first_page_idx=pno,
                    width=w,
                    height=h,
                    bpc=bpc,
                    colorspace_name=cs,
                    filter_name=filt,
                    has_smask=is_smask,
                    max_display_width_pt=current_max_w,
                    effective_dpi=0.0,
                    is_bilevel=is_bilevel,
                )
            else:
                cand = found[xref]
                if is_smask:
                    cand.has_smask = True
                if current_max_w > cand.max_display_width_pt:
                    cand.max_display_width_pt = current_max_w
        del page

    # Effective DPI from the largest displayed size
    for cand in found.values():
        if cand.max_display_width_pt > 0:
            cand.effective_dpi = cand.width / (cand.max_display_width_pt / 72.0)
        else:
            cand.effective_dpi = 0.0

    return found


def _recompress_single_image(
    doc: pymupdf.Document,
    cand: ImageCandidate,
    preset: Preset
) -> Tuple[bool, int, str]:
    """
    Downsample and re-encode one image.
    Returns (success, bytes_saved, action_message).
    action_message is one of: 'ok', 'skipped_smask', 'skipped_bilevel',
    'skipped_larger', 'cmyk_converted', ...
    """
    # Transparency (SMask) is skipped
    if cand.has_smask:
        return False, 0, "skipped_smask"

    # 1-bit, JBIG2, CCITT are skipped
    if cand.is_bilevel:
        return False, 0, "skipped_bilevel"

    # Already JPEG and below the threshold (or size unknown): skip
    is_already_jpeg = cand.filter_name in ("DCTDecode", "/DCTDecode")
    if is_already_jpeg:
        if cand.effective_dpi == 0 or (preset.dpi_threshold is not None and cand.effective_dpi <= preset.dpi_threshold):
            return False, 0, "skipped_low_dpi"

    orig_stream = None
    pix = None
    pix_scaled = None
    data = None
    try:
        try:
            orig_stream = doc.xref_stream_raw(cand.xref)
            orig_len = len(orig_stream)
        except Exception:
            orig_len = 0

        try:
            pix = pymupdf.Pixmap(doc, cand.xref)
        except Exception as e:
            logger.debug("Could not open pixmap for xref %d: %s", cand.xref, e)
            return False, 0, "skipped_error"

        # Alpha mask check
        if pix.alpha:
            return False, 0, "skipped_smask"

        if pix.colorspace is None:
            return False, 0, "skipped_bilevel"

        # Convert CMYK to RGB
        cmyk_converted = False
        if pix.colorspace.n >= 4 or pix.n >= 4:
            try:
                pix_rgb = pymupdf.Pixmap(pymupdf.csRGB, pix)
                del pix
                pix = pix_rgb
                cmyk_converted = True
            except Exception as e:
                logger.debug("CMYK to RGB failed for xref %d: %s", cand.xref, e)
                return False, 0, "skipped_error"

        # Downsample if above the threshold
        if (
            preset.dpi_threshold is not None
            and preset.target_dpi is not None
            and cand.effective_dpi > preset.dpi_threshold
        ):
            scale = preset.target_dpi / cand.effective_dpi
            target_w = max(1, int(pix.width * scale))
            target_h = max(1, int(pix.height * scale))
            if target_w < pix.width and target_h < pix.height:
                try:
                    pix_scaled = pymupdf.Pixmap(pix, target_w, target_h)
                    del pix
                    pix = pix_scaled
                    pix_scaled = None
                except Exception as e:
                    # Fallback resize via Pillow if needed
                    try:
                        from PIL import Image
                        mode = "RGB" if pix.n == 3 else ("L" if pix.n == 1 else "RGB")
                        pil_img = Image.frombytes(mode, (pix.width, pix.height), bytes(pix.samples), "raw", mode, pix.stride)
                        pil_resized = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
                        del pil_img
                        buf = io.BytesIO()
                        pil_resized.save(buf, format="JPEG", quality=preset.jpeg_quality or 80, optimize=True)
                        del pil_resized
                        data = buf.getvalue()
                        buf.close()
                        del buf
                        if orig_len > 0 and len(data) >= orig_len:
                            return False, 0, "skipped_larger"
                        doc[cand.first_page_idx].replace_image(cand.xref, stream=data)
                        saved = orig_len - len(data) if orig_len > 0 else 0
                        action = "cmyk_converted" if cmyk_converted else "ok"
                        return True, saved, action
                    except Exception as pe:
                        logger.debug("Pillow resize fallback failed: %s", pe)

        # Encode to JPEG
        try:
            quality = preset.jpeg_quality if preset.jpeg_quality is not None else 80
            data = pix.tobytes("jpeg", jpg_quality=quality)
        except Exception as e:
            logger.debug("JPEG encode failed for xref %d: %s", cand.xref, e)
            return False, 0, "skipped_error"

        # Keep the original if re-encoding is not smaller
        if orig_len > 0 and len(data) >= orig_len:
            return False, 0, "skipped_larger"

        try:
            doc[cand.first_page_idx].replace_image(cand.xref, stream=data)
        except Exception as e:
            logger.debug("replace_image failed for xref %d: %s", cand.xref, e)
            return False, 0, "skipped_error"

        saved = orig_len - len(data) if orig_len > 0 else 0
        action = "cmyk_converted" if cmyk_converted else "ok"
        return True, saved, action
    finally:
        del pix
        del pix_scaled
        del orig_stream
        del data


# ══════════════════════════════════════════════════════════════════════════════
# MAIN OPTIMIZATION PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def optimize_pdf(
    src_path: str | Path,
    dst_path: Optional[str | Path] = None,
    preset_key: str = DEFAULT_PRESET_KEY,
    password: Optional[str] = None,
    on_progress: Optional[Callable[[str, int], None]] = None,
    cancel_event: Optional[any] = None,
    ignore_signatures: bool = False,
) -> OptimizationResult:
    """
    Full PDF optimisation pipeline.
    Writes a new PDF; the original file is never changed or overwritten.
    """
    src = Path(src_path)
    if not src.is_file():
        raise CorruptPdf(f"File not found: {src}")

    orig_size = src.stat().st_size
    preset = PRESETS.get(preset_key, PRESETS[DEFAULT_PRESET_KEY])

    def report(stage: str, percent: int):
        if on_progress:
            try:
                on_progress(stage, percent)
            except Exception:
                pass

    report("Opening file...", 5)
    if cancel_event and cancel_event.is_set():
        raise CancelledError("Cancelled by user.")

    doc = None
    dst_tmp = None
    try:
        # 1. Open the document
        try:
            doc = pymupdf.open(to_extended_path(src))
        except Exception as e:
            logger.error("Could not open PDF %s: %s", src.name, e)
            raise CorruptPdf(f"File cannot be read or is damaged: {e}")

        # Encryption & password (record state before authenticate, which clears it)
        was_encrypted = bool(doc.is_encrypted)
        if doc.needs_pass:
            if not password or not doc.authenticate(password):
                raise PasswordRequired("A password is needed to open this PDF.")

        # Digital signature check
        if has_digital_signatures(doc) and not ignore_signatures:
            raise DigitalSignatureWarning(
                "This PDF is digitally signed. "
                "Compressing it will invalidate the signature."
            )

        orig_page_count = len(doc)
        final_dst = build_output_path(src, dst_path)
        final_dst.parent.mkdir(parents=True, exist_ok=True)

        # Temp file in the same folder (same-volume atomic rename)
        dst_tmp = final_dst.parent / f"{final_dst.name}.tmp_{os.getpid()}_{time.time_ns()}"

        report("Analysing images...", 15)
        if cancel_event and cancel_event.is_set():
            raise CancelledError("Cancelled by user.")

        # 2. Image processing (skipped for the Light preset)
        images_total = 0
        images_processed = 0
        images_skipped = 0
        cmyk_converted_count = 0

        if preset.target_dpi is not None:
            candidates = collect_images(doc)
            images_total = len(candidates)
            cand_list = list(candidates.values())

            for i, cand in enumerate(cand_list, 1):
                if cancel_event and cancel_event.is_set():
                    raise CancelledError("Cancelled by user.")

                pct = 15 + int(65 * (i / max(1, images_total)))
                report(f"Compressing images ({i}/{images_total})...", pct)

                ok, saved_bytes, action = _recompress_single_image(doc, cand, preset)
                if ok:
                    images_processed += 1
                    if action == "cmyk_converted":
                        cmyk_converted_count += 1
                else:
                    images_skipped += 1

                if i % 25 == 0:
                    gc.collect()
        else:
            report("Skipping images (Light preset)...", 75)

        if cancel_event and cancel_event.is_set():
            raise CancelledError("Cancelled by user.")

        # 3. Subset fonts (Standard & Maximum presets)
        if preset.subset_fonts and hasattr(doc, "subset_fonts"):
            report("Subsetting fonts...", 82)
            try:
                doc.subset_fonts()
            except Exception as e:
                logger.warning("Font subsetting failed for %s: %s", src.name, e)

        # 4. Clean metadata
        report("Cleaning metadata...", 88)
        try:
            doc.set_metadata({})
            if hasattr(doc, "del_xml_metadata"):
                doc.del_xml_metadata()
        except Exception as e:
            logger.debug("Could not clean metadata: %s", e)

        # 5. Save to the temp file
        report("Saving...", 92)
        save_kwargs = dict(
            garbage=4,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            clean=True,
            use_objstms=True,
        )
        if was_encrypted:
            save_kwargs["encryption"] = pymupdf.PDF_ENCRYPT_KEEP

        try:
            doc.save(str(dst_tmp), **save_kwargs)
        except Exception as e:
            logger.error("Could not save compressed document %s: %s", dst_tmp.name, e)
            raise OptimizerError(f"Could not save the compressed PDF: {e}")
        finally:
            try:
                doc.close()
                doc = None
            except Exception:
                pass

        # 6. Validate and check the savings threshold
        if not dst_tmp.exists():
            raise OptimizerError("The temp file was not created.")

        new_size = dst_tmp.stat().st_size
        saved_bytes = orig_size - new_size
        saved_percent = (saved_bytes / orig_size * 100.0) if orig_size > 0 else 0.0

        # Threshold: < 3% AND < 500 KB (or the new file is not smaller)
        is_savings_negligible = (saved_bytes < 500 * 1024) and (saved_percent < 3.0)
        if new_size >= orig_size or is_savings_negligible:
            # Discard the temp file
            try:
                dst_tmp.unlink(missing_ok=True)
            except Exception:
                pass
            dst_tmp = None

            msg = "Already optimal: no meaningful savings."
            logger.info("%s: %s (original: %s, result: %s, saved: %s / %.1f%%)",
                        src.name, msg, format_size(orig_size), format_size(new_size),
                        format_size(max(0, saved_bytes)), max(0.0, saved_percent))

            return OptimizationResult(
                status="already_optimal",
                src_path=str(src),
                dst_path=None,
                orig_size=orig_size,
                new_size=orig_size,
                saved_bytes=0,
                saved_percent=0.0,
                images_total=images_total,
                images_processed=images_processed,
                images_skipped=images_skipped,
                cmyk_converted=cmyk_converted_count,
                message=msg,
            )

        # 7. Validate the output file
        report("Checking the result...", 96)
        if not validate_output(orig_page_count, dst_tmp, password=password):
            try:
                dst_tmp.unlink(missing_ok=True)
            except Exception:
                pass
            dst_tmp = None
            raise OptimizerError("Output check failed: damaged document or page count mismatch.")

        # 8. Atomic rename
        report("Menyelesaikan...", 99)
        dst_tmp.replace(final_dst)
        dst_tmp = None
        report("Done.", 100)

        msg = f"Compressed ({format_size(orig_size)} → {format_size(new_size)}, saved {saved_percent:.1f}%)"
        logger.info("%s: %s → %s (%s)", src.name, format_size(orig_size),
                    format_size(new_size), final_dst.name)

        return OptimizationResult(
            status="success",
            src_path=str(src),
            dst_path=str(final_dst),
            orig_size=orig_size,
            new_size=new_size,
            saved_bytes=saved_bytes,
            saved_percent=saved_percent,
            images_total=images_total,
            images_processed=images_processed,
            images_skipped=images_skipped,
            cmyk_converted=cmyk_converted_count,
            message=msg,
        )

    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
        if dst_tmp is not None:
            try:
                dst_tmp.unlink(missing_ok=True)
            except Exception:
                pass
