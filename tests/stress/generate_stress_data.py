"""
Synthetic Pathological Stress Test Data Generator for MarkItDown Tools.

This module provides automated, reproducible generation of extreme edge-case
and stress testing fixtures for the MarkItDown Tools desktop application,
PyMuPDF optimization engine (pdf_optimizer.py), PDF merger, PowerPoint COM
automation, and OCR fallback workflows.

Fixtures generated:
1. Extreme PDFs:
   - extreme_size_100mb.pdf: 100+ MB file generated via uncompressed raster stream in 4MB chunks
   - high_page_count_500p.pdf: 500+ pages vector text document
   - high_image_density_1000img.pdf: 1,000+ distinct images with unique digest salts
   - colorspace_cmyk.pdf: 4-channel DeviceCMYK image
   - colorspace_indexed.pdf: Palette-based /Indexed /DeviceRGB image
   - colorspace_separation.pdf: Spot color /Separation /PANTONE_300_C image
   - extreme_aspect_ratio_200to1.pdf: 200:1 and 2000:1 divider and banner images
   - malformed_truncated_eof.pdf: Abrupt EOF truncation (missing %%EOF & trailer)
   - malformed_broken_xref.pdf: Corrupted cross-reference table offsets
   - malformed_truncated_stream.pdf: Image stream truncated before declared /Length
   - malformed_invalid_header.pdf: Non-PDF binary garbage header
2. Encrypted PDFs:
   - 40-bit RC4, 128-bit RC4, 128-bit AES, 256-bit AES
   - User password only, owner password only, distinct user/owner passwords
   - PyMuPDF and pypdf generator cross-engine validation
3. Extreme Presentations:
   - huge_120_slides.pptx: 120-slide presentation
   - massive_media_embedded.pptx: 25 slides embedding images, banners, shapes, tables
   - corrupt_malformed_xml.pptx: Valid ZIP with malformed ppt/presentation.xml
   - corrupt_missing_part_xml.pptx: Valid ZIP with mandatory presentation.xml deleted
   - corrupt_bad_zip_header.pptx: Corrupted ZIP magic bytes (PK\\x00\\x00)
   - corrupt_truncated_zip.pptx: Binary archive truncated at 50%
   - corrupt_zero_byte.pptx: Completely empty 0-byte presentation file
   - password_protected_agile.pptx: Password-protected OLE Compound File (Agile encryption)
4. Filesystem Edge Cases:
   - fs_readonly_file: Existing file with S_IREAD attribute (blocks overwrite)
   - fs_readonly_src: Valid source document marked S_IREAD
   - fs_readonly_dir: Directory with write-denied NTFS ACL (icacls deny W)
   - fs_locked_file: Valid document reserved for cross-process msvcrt locking
   - fs_unicode_emoji_pdf: Emojis and Cyrillic script in filename
   - fs_unicode_cjk_pdf: Japanese CJK Kanji, Katakana, and emoji filename
   - fs_unicode_arabic_pdf: Arabic RTL script and emoji filename
   - fs_unicode_nested_dir: Nested directory structure with emojis
   - fs_long_path_260plus: Deeply nested directory exceeding 260 chars (\\\\?\\ extended path)
"""

from __future__ import annotations

import argparse
import datetime
from datetime import timezone
import io
import json
import logging
import os
import shutil
import stat
import subprocess
import sys
import time
import warnings
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional, Set

# Reconfigure stdout/stderr to UTF-8 on Windows to safely print unicode/emoji filenames
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass

# Suppress python-pptx duplicate layout zipfile warnings in Python 3.13
warnings.filterwarnings("ignore", category=UserWarning, module="zipfile")

# OS-specific locking primitives
try:
    import msvcrt
except ImportError:
    msvcrt = None  # type: ignore

try:
    import fcntl
except ImportError:
    fcntl = None  # type: ignore

# Core third-party stress dependencies
import olefile
import pymupdf
import pypdf
import pptx
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

logger = logging.getLogger("MarkItDown.StressData")


# ══════════════════════════════════════════════════════════════════════════════
# 1. CORE UTILITIES & PLATFORM HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def to_extended_path(p: Path | str) -> str:
    """
    Prepends the Windows extended path prefix \\\\?\\ if the resolved path
    length exceeds or approaches 260 characters (MAX_PATH).
    """
    path_str = str(Path(p).resolve())
    if os.name == "nt" and len(path_str) >= 240 and not path_str.startswith("\\\\?\\"):
        return f"\\\\?\\{path_str}"
    return path_str


def safe_rmtree(path: Path | str) -> None:
    """
    Recursively removes a directory tree, properly removing Windows NTFS
    deny ACLs, clearing read-only attributes, and handling extended paths.
    """
    p = Path(path).resolve()
    if not p.exists():
        return

    # Remove any deny ACL on Windows
    if os.name == "nt":
        user = os.environ.get("USERNAME", "Everyone")
        try:
            subprocess.run(
                ["icacls", str(p), "/remove:d", user],
                capture_output=True,
                text=True,
            )
            subprocess.run(
                ["icacls", str(p), "/grant", f"{user}:(F)", "/t"],
                capture_output=True,
                text=True,
            )
        except Exception:
            pass

    ext_p = to_extended_path(p)

    # Clear read-only flags on all children
    for root, dirs, files in os.walk(ext_p):
        for f in files:
            fp = os.path.join(root, f)
            try:
                os.chmod(fp, stat.S_IWRITE)
            except Exception:
                pass
        for d in dirs:
            dp = os.path.join(root, d)
            try:
                os.chmod(dp, stat.S_IWRITE)
            except Exception:
                pass

    def _handle_error(action: Any, pth: str, exc: Any) -> None:
        try:
            os.chmod(pth, stat.S_IWRITE)
            action(pth)
        except Exception:
            pass

    try:
        shutil.rmtree(ext_p, on_exc=_handle_error)
    except TypeError:
        # Fallback for Python versions prior to on_exc
        shutil.rmtree(ext_p, onerror=_handle_error)
    except Exception:
        shutil.rmtree(ext_p, ignore_errors=True)


@contextmanager
def hold_file_lock(filepath: Path | str) -> Generator[None, None, None]:
    """
    In-process context manager holding an exclusive non-blocking byte-range lock.
    Uses msvcrt.locking on Windows or fcntl.flock on Unix.
    """
    p = Path(filepath).resolve()
    if not p.exists():
        raise FileNotFoundError(f"File not found for locking: {p}")
    size = max(1, p.stat().st_size)
    f = open(p, "r+b")
    try:
        if msvcrt:
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, size)
        elif fcntl:
            fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        yield
    finally:
        try:
            f.seek(0)
            if msvcrt:
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, size)
            elif fcntl:
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
        except Exception:
            pass
        finally:
            f.close()


def run_lock_daemon(file_path: Path | str) -> None:
    """
    CLI lock daemon: holds an exclusive byte lock on file_path, writes 'READY\\n'
    to stdout, and blocks until stdin closes or process terminates.
    """
    p = Path(file_path).resolve()
    if not p.exists():
        sys.stderr.write(f"File not found: {p}\n")
        sys.exit(1)
    size = max(1, p.stat().st_size)
    f = open(p, "r+b")
    try:
        if msvcrt:
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, size)
        elif fcntl:
            fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)

        sys.stdout.write("READY\n")
        sys.stdout.flush()

        try:
            sys.stdin.readline()
        except (KeyboardInterrupt, EOFError):
            pass
    finally:
        try:
            f.seek(0)
            if msvcrt:
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, size)
            elif fcntl:
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
        except Exception:
            pass
        finally:
            f.close()


def spawn_file_lock_process(path: Path | str) -> subprocess.Popen:
    """
    Spawns a background Python subprocess holding an active lock on path.
    Blocks until the 'READY' handshake is received from child stdout.
    """
    cmd = [
        sys.executable,
        "-m",
        "tests.stress.generate_stress_data",
        "--lock-file",
        str(Path(path).resolve()),
    ]
    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if proc.stdout is None:
        raise RuntimeError("Subprocess stdout is not available.")
    ready = proc.stdout.readline().strip()
    if ready != "READY":
        stderr = proc.stderr.read() if proc.stderr else ""
        proc.kill()
        raise RuntimeError(f"Failed to acquire file lock via subprocess: {stderr}")
    return proc


def release_file_lock_process(proc: subprocess.Popen) -> None:
    """Signals and releases the background lock holder subprocess cleanly."""
    try:
        if proc.stdin:
            proc.stdin.write("\n")
            proc.stdin.flush()
        proc.wait(timeout=2.0)
    except Exception:
        proc.kill()
        proc.wait(timeout=1.0)


def is_office_encrypted(filepath: Path | str) -> bool:
    """
    Fast, non-blocking check to determine if an Office file is password-protected.
    Detects OLE CFB container with EncryptionInfo or EncryptedPackage streams.
    Prevents PowerPoint COM automation from hanging on modal password prompts.
    """
    path = Path(filepath)
    if not path.is_file():
        return False
    try:
        with open(path, "rb") as f:
            magic = f.read(8)
        if magic == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
            if olefile.isOleFile(path):
                with olefile.OleFileIO(path) as ole:
                    return ole.exists("EncryptionInfo") or ole.exists("EncryptedPackage")
            # Fallback byte pattern search if olefile parse fails
            with open(path, "rb") as f:
                head = f.read(4096)
                if b"EncryptionInfo" in head or b"EncryptedPackage" in head:
                    return True
        return False
    except Exception:
        return False


def is_pdf_encrypted(filepath: Path | str) -> bool:
    """
    Fast, non-invasive check to determine if a PDF file is encrypted / password-protected.
    """
    path = Path(filepath)
    if not path.is_file():
        return False
    try:
        doc = pymupdf.open(str(path))
        is_enc = bool(doc.is_encrypted or doc.needs_pass)
        doc.close()
        return is_enc
    except Exception:
        try:
            reader = pypdf.PdfReader(str(path))
            return bool(reader.is_encrypted)
        except Exception:
            return False


# ══════════════════════════════════════════════════════════════════════════════
# 2. EXTREME PDF FIXTURES GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_extreme_size_pdf(output_path: Path) -> Path:
    """
    Generates a 100+ MB PDF (approx. 103.0 MB) in < 0.2s using uncompressed raw streams.
    Memory overhead is strictly bounded (< 5 MB RAM) via 4 MB chunk streaming.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    width, height = 6000, 6000
    stream_len = width * height * 3  # 108,000,000 bytes

    with open(output_path, "wb") as f:
        offsets = {}
        f.write(b"%PDF-1.7\n")

        # 1 0 obj: Catalog
        offsets[1] = f.tell()
        f.write(b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")

        # 2 0 obj: Pages
        offsets[2] = f.tell()
        f.write(b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")

        # 3 0 obj: Page
        offsets[3] = f.tell()
        f.write(
            b"3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            b"/Contents 4 0 R /Resources << /XObject << /Img1 5 0 R >> >> >>\nendobj\n"
        )

        # 4 0 obj: Contents
        offsets[4] = f.tell()
        content = b"q 612 0 0 792 0 0 cm /Img1 Do Q\n"
        f.write(f"4 0 obj\n<< /Length {len(content)} >>\nstream\n".encode("ascii"))
        f.write(content)
        f.write(b"endstream\nendobj\n")

        # 5 0 obj: Uncompressed Image XObject
        offsets[5] = f.tell()
        header = (
            f"5 0 obj\n<< /Type /XObject /Subtype /Image /Width {width} "
            f"/Height {height} /ColorSpace /DeviceRGB /BitsPerComponent 8 "
            f"/Length {stream_len} >>\nstream\n"
        ).encode("ascii")
        f.write(header)

        # Stream 108 MB in reusable 4 MB chunks
        chunk = b"\xaa\xbb\xcc" * (4 * 1024 * 1024 // 3)
        remaining = stream_len
        while remaining > 0:
            to_write = min(remaining, len(chunk))
            f.write(chunk[:to_write])
            remaining -= to_write

        f.write(b"\nendstream\nendobj\n")

        # XRef table
        xref_offset = f.tell()
        f.write(b"xref\n0 6\n0000000000 65535 f \r\n")
        for i in range(1, 6):
            f.write(f"{offsets[i]:010d} 00000 n \r\n".encode("ascii"))

        f.write(f"trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii"))

    return output_path


def generate_high_page_count_pdf(output_path: Path, page_count: int = 500) -> Path:
    """
    Generates a 500+ page PDF in < 0.6s via PyMuPDF with text and vector shapes on every page.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    rect = pymupdf.Rect(0, 0, 595, 842)  # A4

    for i in range(page_count):
        page = doc.new_page(width=rect.width, height=rect.height)
        page.insert_text(
            pymupdf.Point(50, 72),
            f"Stress Test Document - Page {i + 1} of {page_count}",
            fontsize=12,
        )
        page.insert_text(
            pymupdf.Point(50, 95),
            "Milestone M1 High Page Count Stress Fixture",
            fontsize=9,
        )

        shape = page.new_shape()
        shape.draw_rect(pymupdf.Rect(50, 110, 545, 780))
        shape.draw_line(pymupdf.Point(50, 140), pymupdf.Point(545, 140))
        shape.finish(color=(0.2, 0.4, 0.6), width=1)
        shape.commit()

    doc.save(str(output_path), deflate=True)
    doc.close()
    return output_path


def generate_high_image_density_pdf(output_path: Path, image_count: int = 1000) -> Path:
    """
    Generates a document embedding 1,000 distinct unique image XObjects with varied dimensions
    and unique prefix salts to guarantee distinct digest xrefs in PyMuPDF.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    num_pages = 20
    per_page = image_count // num_pages
    buffer_template = bytes(range(256)) * 40

    for pno in range(num_pages):
        page = doc.new_page(width=595, height=842)
        for i in range(per_page):
            idx = pno * per_page + i
            iw = 12 + (idx % 24)
            ih = 12 + ((idx * 5) % 24)
            # Prefix with idx to guarantee distinct SHA-1 digest in MuPDF
            samples = idx.to_bytes(4, "big") + buffer_template[: iw * ih * 3 - 4]
            pix = pymupdf.Pixmap(pymupdf.csRGB, iw, ih, samples, False)

            c = i % 5
            r = i // 5
            rect = pymupdf.Rect(50 + c * 95, 70 + r * 70, 50 + c * 95 + 80, 70 + r * 70 + 60)
            page.insert_image(rect, pixmap=pix)

    doc.save(str(output_path), deflate=True)
    doc.close()
    return output_path


def generate_cmyk_pdf(output_path: Path) -> Path:
    """Generates a PDF embedding a DeviceCMYK 4-channel image object."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    w, h = 800, 600
    samples = bytes([100, 150, 200, 50]) * (w * h)
    pix = pymupdf.Pixmap(pymupdf.csCMYK, w, h, samples, False)
    page.insert_image(pymupdf.Rect(50, 50, 545, 450), pixmap=pix)
    doc.save(str(output_path), deflate=True)
    doc.close()
    return output_path


def generate_indexed_pdf(output_path: Path) -> Path:
    """Generates a conforming PDF-1.7 /Indexed /DeviceRGB color space image."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    raw_pdf = b"""%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /XObject << /ImgIdx 5 0 R >> >> >> endobj
4 0 obj << /Length 32 >> stream
q 200 0 0 200 50 500 cm /ImgIdx Do Q
endstream endobj
5 0 obj << /Type /XObject /Subtype /Image /Width 4 /Height 4 /BitsPerComponent 8 /ColorSpace [ /Indexed /DeviceRGB 3 <FF0000 00FF00 0000FF FFFF00> ] /Length 16 >> stream
\x00\x01\x02\x03\x03\x02\x01\x00\x00\x01\x02\x03\x03\x02\x01\x00
endstream endobj
xref
0 6
0000000000 65535 f \r\n
0000000009 00000 n \r\n
0000000058 00000 n \r\n
0000000115 00000 n \r\n
0000000255 00000 n \r\n
0000000338 00000 n \r\n
trailer << /Size 6 /Root 1 0 R >>
startxref
525
%%EOF
"""
    output_path.write_bytes(raw_pdf)
    return output_path


def generate_separation_pdf(output_path: Path) -> Path:
    """Generates a conforming PDF-1.7 /Separation spot color image (PANTONE 300 C)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    raw_pdf = b"""%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /XObject << /ImgSep 5 0 R >> >> >> endobj
4 0 obj << /Length 30 >> stream
q 200 0 0 200 50 500 cm /ImgSep Do Q
endstream endobj
5 0 obj << /Type /XObject /Subtype /Image /Width 4 /Height 4 /BitsPerComponent 8 /ColorSpace [ /Separation /PANTONE_300_C /DeviceCMYK << /FunctionType 2 /Domain [0 1] /Range [0 1 0 1 0 1 0 1] /C0 [0 0 0 0] /C1 [1.0 0.5 0.0 0.0] /N 1.0 >> ] /Length 16 >> stream
\x10\x30\x60\x80\xa0\xc0\xe0\xff\x10\x30\x60\x80\xa0\xc0\xe0\xff
endstream endobj
xref
0 6
0000000000 65535 f \r\n
0000000009 00000 n \r\n
0000000058 00000 n \r\n
0000000115 00000 n \r\n
0000000253 00000 n \r\n
0000000334 00000 n \r\n
trailer << /Size 6 /Root 1 0 R >>
startxref
570
%%EOF
"""
    output_path.write_bytes(raw_pdf)
    return output_path


def generate_extreme_aspect_ratios_pdf(output_path: Path) -> Path:
    """
    Generates a PDF embedding 200:1 and 2000:1 aspect ratio images (horizontal rules,
    banners, vertical dividers, and single-pixel line).
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)

    # 2000x10 horizontal divider (200:1)
    s1 = bytes([255, 0, 0]) * (2000 * 10)
    pix1 = pymupdf.Pixmap(pymupdf.csRGB, 2000, 10, s1, False)
    page.insert_image(pymupdf.Rect(50, 100, 545, 103), pixmap=pix1)

    # 10000x50 large banner (200:1)
    s2 = bytes([0, 128, 255]) * (10000 * 50)
    pix2 = pymupdf.Pixmap(pymupdf.csRGB, 10000, 50, s2, False)
    page.insert_image(pymupdf.Rect(50, 150, 545, 153), pixmap=pix2)

    # 10x2000 vertical divider (1:200)
    s3 = bytes([0, 255, 0]) * (10 * 2000)
    pix3 = pymupdf.Pixmap(pymupdf.csRGB, 10, 2000, s3, False)
    page.insert_image(pymupdf.Rect(50, 200, 53, 700), pixmap=pix3)

    # 2000x1 ultra-thin line (2000:1)
    s4 = bytes([50, 50, 50]) * (2000 * 1)
    pix4 = pymupdf.Pixmap(pymupdf.csRGB, 2000, 1, s4, False)
    page.insert_image(pymupdf.Rect(50, 750, 545, 751), pixmap=pix4)

    doc.save(str(output_path), deflate=True)
    doc.close()
    return output_path


def generate_malformed_truncated_eof_pdf(output_path: Path) -> Path:
    """Generates a PDF truncated at 60% of stream (unexpected EOF, missing trailer)."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((50, 50), "Valid text stream before abrupt EOF truncation")
    raw = doc.tobytes()
    doc.close()

    truncated = raw[: len(raw) * 6 // 10]
    output_path.write_bytes(truncated)
    return output_path


def generate_malformed_broken_xref_pdf(output_path: Path) -> Path:
    """Generates a PDF with corrupted xref table offsets."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((50, 50), "PDF with corrupted xref table offsets")
    raw = doc.tobytes()
    doc.close()

    corrupted = raw.replace(b"xref\n", b"xref\n9999999999 99999 f \r\nCORRUPT_XREF_DATA\n")
    output_path.write_bytes(corrupted)
    return output_path


def generate_malformed_truncated_stream_pdf(output_path: Path) -> Path:
    """Generates a PDF where image stream declares Length 10000 but ends in 6 bytes."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    raw_pdf = b"""%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /XObject << /Img1 5 0 R >> >> >> endobj
4 0 obj << /Length 20 >> stream
q /Img1 Do Q
endstream endobj
5 0 obj << /Type /XObject /Subtype /Image /Width 100 /Height 100 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 10000 >> stream
\xaa\xbb\xcc\xdd\xee\xff
endstream endobj
xref
0 6
0000000000 65535 f \r\n
0000000009 00000 n \r\n
0000000058 00000 n \r\n
0000000115 00000 n \r\n
0000000240 00000 n \r\n
0000000305 00000 n \r\n
trailer << /Size 6 /Root 1 0 R >>
startxref
490
%%EOF
"""
    output_path.write_bytes(raw_pdf)
    return output_path


def generate_malformed_invalid_header_pdf(output_path: Path) -> Path:
    """Generates a non-PDF binary payload with a .pdf extension."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(b"NOT_A_VALID_PDF_HEADER_RAW_DATA_1234567890\x00\x01\x02\x03\xff\xfe")
    return output_path


def generate_extreme_pdfs(target_dir: Path, seed: int = 42) -> Dict[str, Path]:
    """
    Master programmatic entry point for Extreme PDF fixtures.
    Generates all 11 pathological PDF fixtures in target_dir.
    """
    target_dir = Path(target_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    fixtures = {
        "extreme_size_100mb": generate_extreme_size_pdf(target_dir / "extreme_size_100mb.pdf"),
        "high_page_count_500p": generate_high_page_count_pdf(target_dir / "high_page_count_500p.pdf"),
        "high_image_density_1000img": generate_high_image_density_pdf(target_dir / "high_image_density_1000img.pdf"),
        "colorspace_cmyk": generate_cmyk_pdf(target_dir / "colorspace_cmyk.pdf"),
        "colorspace_indexed": generate_indexed_pdf(target_dir / "colorspace_indexed.pdf"),
        "colorspace_separation": generate_separation_pdf(target_dir / "colorspace_separation.pdf"),
        "extreme_aspect_ratio_200to1": generate_extreme_aspect_ratios_pdf(target_dir / "extreme_aspect_ratio_200to1.pdf"),
        "malformed_truncated_eof": generate_malformed_truncated_eof_pdf(target_dir / "malformed_truncated_eof.pdf"),
        "malformed_broken_xref": generate_malformed_broken_xref_pdf(target_dir / "malformed_broken_xref.pdf"),
        "malformed_truncated_stream": generate_malformed_truncated_stream_pdf(target_dir / "malformed_truncated_stream.pdf"),
        "malformed_invalid_header": generate_malformed_invalid_header_pdf(target_dir / "malformed_invalid_header.pdf"),
    }
    return fixtures


# ══════════════════════════════════════════════════════════════════════════════
# 3. ENCRYPTED PDF FIXTURES GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_encrypted_pdfs(target_dir: Path, seed: int = 42) -> Dict[str, Path]:
    """
    Generates PDF fixtures across standard encryption variations:
    - 40-bit RC4, 128-bit RC4, 128-bit AES, 256-bit AES
    - User password only, owner password only, distinct passwords
    - PyMuPDF and pypdf generator cross-engine validation
    """
    target_dir = Path(target_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    fixtures: Dict[str, Path] = {}

    specs = [
        # (key, filename, algorithm_flag, user_pw, owner_pw, permissions)
        ("rc4_40_user_owner", "rc4_40_user_owner.pdf", pymupdf.PDF_ENCRYPT_RC4_40, "user123", "owner123", pymupdf.PDF_PERM_PRINT),
        ("rc4_40_user_only", "rc4_40_user_only.pdf", pymupdf.PDF_ENCRYPT_RC4_40, "user123", "", pymupdf.PDF_PERM_PRINT),
        ("rc4_40_owner_only", "rc4_40_owner_only.pdf", pymupdf.PDF_ENCRYPT_RC4_40, "", "owner123", pymupdf.PDF_PERM_PRINT),
        ("rc4_128_user_owner", "rc4_128_user_owner.pdf", pymupdf.PDF_ENCRYPT_RC4_128, "user123", "owner123", pymupdf.PDF_PERM_PRINT),
        ("aes_128_user_owner", "aes_128_user_owner.pdf", pymupdf.PDF_ENCRYPT_AES_128, "user123", "owner123", pymupdf.PDF_PERM_PRINT),
        ("aes_128_user_only", "aes_128_user_only.pdf", pymupdf.PDF_ENCRYPT_AES_128, "user123", "", pymupdf.PDF_PERM_PRINT),
        ("aes_128_owner_only", "aes_128_owner_only.pdf", pymupdf.PDF_ENCRYPT_AES_128, "", "owner123", pymupdf.PDF_PERM_PRINT),
        ("aes_256_user_owner", "aes_256_user_owner.pdf", pymupdf.PDF_ENCRYPT_AES_256, "user123", "owner123", pymupdf.PDF_PERM_PRINT),
        ("aes_256_user_only", "aes_256_user_only.pdf", pymupdf.PDF_ENCRYPT_AES_256, "user123", "", pymupdf.PDF_PERM_PRINT),
        ("aes_256_owner_only", "aes_256_owner_only.pdf", pymupdf.PDF_ENCRYPT_AES_256, "", "owner123", pymupdf.PDF_PERM_PRINT),
    ]

    for key, fname, enc_flag, u_pw, o_pw, perms in specs:
        p = target_dir / fname
        doc = pymupdf.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 72), f"Synthetic Encrypted Fixture: {fname}")
        page.insert_text(
            (72, 100),
            f"User PW: {'[SET: user123]' if u_pw else '[NONE]'} | Owner PW: {'[SET: owner123]' if o_pw else '[NONE]'}",
        )
        doc.save(
            str(p),
            encryption=enc_flag,
            user_pw=u_pw,
            owner_pw=o_pw,
            permissions=perms,
            garbage=4,
            deflate=True,
        )
        doc.close()
        fixtures[key] = p
        fixtures[fname] = p

    # pypdf generated fixture (cross-engine parity verification)
    p_pypdf = target_dir / "pypdf_aes_256_user_owner.pdf"
    writer = pypdf.PdfWriter()
    writer.add_blank_page(width=595, height=842)
    writer.encrypt(user_password="user123", owner_password="owner123", algorithm="AES-256")
    with open(p_pypdf, "wb") as f:
        writer.write(f)
    writer.close()
    fixtures["pypdf_aes_256_user_owner"] = p_pypdf
    fixtures["pypdf_aes_256_user_owner.pdf"] = p_pypdf

    # Convenience standard aliases matching test suites
    fixtures["pdf_enc_rc4_40"] = fixtures["rc4_40_user_owner"]
    fixtures["pdf_enc_aes_128"] = fixtures["aes_128_user_owner"]
    fixtures["pdf_enc_aes_256"] = fixtures["aes_256_user_owner"]

    return fixtures


# ══════════════════════════════════════════════════════════════════════════════
# 4. EXTREME PRESENTATIONS GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def _repack_zip_with_modification(
    zip_bytes: bytes,
    replace_map: Optional[Dict[str, bytes]] = None,
    remove_set: Optional[Set[str]] = None,
) -> bytes:
    """Repackages a ZIP archive cleanly in-memory, replacing or omitting targeted parts."""
    import zipfile

    replace_map = replace_map or {}
    remove_set = remove_set or set()
    seen = set()
    out_buf = io.BytesIO()

    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zin, \
         zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            if name in seen or name in remove_set:
                continue
            seen.add(name)
            if name in replace_map:
                zout.writestr(name, replace_map[name])
            else:
                zout.writestr(name, zin.read(name))

    return out_buf.getvalue()


def _write_synthetic_encrypted_ole(dest_path: Path) -> None:
    """Writes a minimal synthetic OLE CFB file carrying an EncryptionInfo signature."""
    ole_header = (
        b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"  # Header signature
        + b"\x00" * 16                        # Header CLSID
        + b"\x3E\x00\x03\x00"                # Minor / Major version (v3)
        + b"\xFE\xFF\x09\x00"                # Byte order (LE), sector shift (512)
        + b"\x06\x00"                        # Mini sector shift (64)
        + b"\x00" * 6                        # Reserved
        + b"\x01\x00\x00\x00"                # Number of directory sectors
        + b"\x01\x00\x00\x00"                # Number of FAT sectors
        + b"\x01\x00\x00\x00"                # First directory sector location
        + b"\x00\x00\x00\x00"                # Transaction signature
        + b"\x00\x10\x00\x00"                # Mini stream cutoff size (4096)
        + b"\xFE\xFF\xFF\xFA"                # Mini FAT first sector
        + b"\x00\x00\x00\x00"                # Number of Mini FAT sectors
        + b"\xFE\xFF\xFF\xFD"                # DIFAT first sector
        + b"\x00\x00\x00\x00"                # Number of DIFAT sectors
    )
    content = ole_header.ljust(512, b"\x00") + b"EncryptionInfo\x00EncryptedPackage\x00" * 32
    dest_path.write_bytes(content)


def _generate_or_fallback_encrypted_pptx(dest_path: Path, password: str = "secret123") -> None:
    """Generates an encrypted PPTX using PowerPoint COM if available, otherwise synthetic OLE fallback."""
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    com_success = False

    if os.name == "nt":
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt.Presentations.Add(WithWindow=False)
            slide = pres.Slides.Add(1, 12)  # ppLayoutBlank
            tb = slide.Shapes.AddTextbox(1, 100, 100, 400, 100)
            tb.TextFrame.TextRange.Text = "Confidential Protected Presentation"
            pres.Password = password
            pres.SaveAs(str(dest_path.resolve()))
            pres.Close()
            ppt.Quit()
            pythoncom.CoUninitialize()
            com_success = dest_path.exists() and dest_path.stat().st_size > 0
        except Exception:
            com_success = False

    if not com_success or not dest_path.exists():
        _write_synthetic_encrypted_ole(dest_path)


def generate_extreme_presentations(target_dir: Path, seed: int = 42) -> Dict[str, Path]:
    """
    Generates extreme synthetic presentation fixtures:
    - 120-slide presentation (< 0.2s generation)
    - Embedded media (PNG, JPEG banners 200:1, shapes, tables)
    - Corrupted XML and ZIP containers (malformed XML, missing part, bad header, truncated, 0-byte)
    - Password-protected PPTX (Encrypted OLE Agile container)
    """
    target_dir = Path(target_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    fixtures: Dict[str, Path] = {}

    # 1. 120 Slides Deck
    p_huge = target_dir / "huge_120_slides.pptx"
    prs_huge = pptx.Presentation()
    blank_layout = prs_huge.slide_layouts[6]
    for i in range(120):
        slide = prs_huge.slides.add_slide(blank_layout)
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tb.text_frame.paragraphs[0].text = f"Slide {i + 1} of 120: Stress Slide Payload"
    prs_huge.save(str(p_huge))
    fixtures["huge_120_slides"] = p_huge
    fixtures["huge_120_slides.pptx"] = p_huge

    # 2. Heavy Embedded Media Deck
    p_media = target_dir / "massive_media_embedded.pptx"
    prs_media = pptx.Presentation()

    img_buf = io.BytesIO()
    Image.new("RGB", (300, 300), (45, 105, 165)).save(img_buf, format="PNG")
    banner_buf = io.BytesIO()
    Image.new("RGB", (1000, 5), (220, 20, 60)).save(banner_buf, format="JPEG")

    for i in range(25):
        s = prs_media.slides.add_slide(blank_layout)
        img_buf.seek(0)
        s.shapes.add_picture(img_buf, Inches(0.5), Inches(0.5), width=Inches(2.5))
        banner_buf.seek(0)
        s.shapes.add_picture(banner_buf, Inches(0.5), Inches(3.5), width=Inches(9.0), height=Inches(0.1))
        shape = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.5), Inches(0.5), Inches(2.5), Inches(2.5))
        shape.text = f"Shape {i + 1}"
        tbl = s.shapes.add_table(3, 3, Inches(0.5), Inches(4.0), Inches(8.0), Inches(2.0)).table
        for r in range(3):
            for c in range(3):
                tbl.cell(r, c).text = f"Cell [{r},{c}]"
    prs_media.save(str(p_media))
    fixtures["massive_media_embedded"] = p_media
    fixtures["massive_media_embedded.pptx"] = p_media

    # Base bytes for ZIP corruptions
    base_bytes = p_huge.read_bytes()

    # 3. Valid ZIP with Malformed XML in ppt/presentation.xml
    p_malformed_xml = target_dir / "corrupt_malformed_xml.pptx"
    p_malformed_xml.write_bytes(
        _repack_zip_with_modification(
            base_bytes,
            replace_map={"ppt/presentation.xml": b"<p:presentation><malformed_tag_unclosed>"},
        )
    )
    fixtures["corrupt_malformed_xml"] = p_malformed_xml
    fixtures["corrupt_malformed_xml.pptx"] = p_malformed_xml

    # 4. Valid ZIP with Missing Mandatory Part
    p_missing_part = target_dir / "corrupt_missing_part_xml.pptx"
    p_missing_part.write_bytes(
        _repack_zip_with_modification(
            base_bytes,
            remove_set={"ppt/presentation.xml"},
        )
    )
    fixtures["corrupt_missing_part_xml"] = p_missing_part
    fixtures["corrupt_missing_part_xml.pptx"] = p_missing_part

    # 5. Corrupted ZIP Header
    p_bad_header = target_dir / "corrupt_bad_zip_header.pptx"
    p_bad_header.write_bytes(b"PK\x00\x00" + base_bytes[4:])
    fixtures["corrupt_bad_zip_header"] = p_bad_header
    fixtures["corrupt_bad_zip_header.pptx"] = p_bad_header

    # 6. Truncated ZIP Container
    p_truncated = target_dir / "corrupt_truncated_zip.pptx"
    p_truncated.write_bytes(base_bytes[: len(base_bytes) // 2])
    fixtures["corrupt_truncated_zip"] = p_truncated
    fixtures["corrupt_truncated_zip.pptx"] = p_truncated

    # 7. Zero-Byte PPTX
    p_zero = target_dir / "corrupt_zero_byte.pptx"
    p_zero.write_bytes(b"")
    fixtures["corrupt_zero_byte"] = p_zero
    fixtures["corrupt_zero_byte.pptx"] = p_zero

    # 8. Password-Protected PPTX (Agile OLE Container)
    p_enc_pptx = target_dir / "password_protected_agile.pptx"
    _generate_or_fallback_encrypted_pptx(p_enc_pptx, password="secret123")
    fixtures["password_protected_agile"] = p_enc_pptx
    fixtures["password_protected_agile.pptx"] = p_enc_pptx

    # Standard aliases matching test plans
    fixtures["pptx_100slides"] = p_huge
    fixtures["pptx_embedded_media"] = p_media
    fixtures["pptx_corrupted_xml"] = p_malformed_xml
    fixtures["pptx_password_protected"] = p_enc_pptx

    return fixtures


# ══════════════════════════════════════════════════════════════════════════════
# 5. FILESYSTEM EDGE CASES GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def generate_filesystem_edge_cases(target_dir: Path, seed: int = 42) -> Dict[str, Path]:
    """
    Generates all filesystem edge case fixtures:
    - Read-only destination file (S_IREAD blocks overwrite)
    - Read-only source file (S_IREAD)
    - Write-denied directory (NTFS ACL deny W)
    - File locked target (for cross-process msvcrt locking)
    - Unicode and emoji filenames (CJK, Arabic RTL, emoji scripts)
    - Deeply nested path exceeding 260 characters (MAX_PATH)
    """
    fs_dir = Path(target_dir).resolve()
    fs_dir.mkdir(parents=True, exist_ok=True)
    fixtures: Dict[str, Path] = {}

    def _create_sample_pdf(dest_path: str | Path, text: str) -> Path:
        dest_str = str(dest_path)
        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((50, 50), text)
        doc.save(dest_str)
        doc.close()
        return Path(dest_str)

    # 1. Read-only destination file (for overwrite failure testing)
    ro_file = fs_dir / "fs_readonly_target.pdf"
    if ro_file.exists():
        try:
            ro_file.chmod(stat.S_IWRITE)
        except Exception:
            pass
    _create_sample_pdf(ro_file, "Pre-existing read-only destination file")
    ro_file.chmod(stat.S_IREAD)
    fixtures["fs_readonly_file"] = ro_file

    # 2. Read-only source file (for input read-only testing)
    ro_src = fs_dir / "fs_readonly_src.pdf"
    if ro_src.exists():
        try:
            ro_src.chmod(stat.S_IWRITE)
        except Exception:
            pass
    _create_sample_pdf(ro_src, "Valid read-only source document")
    ro_src.chmod(stat.S_IREAD)
    fixtures["fs_readonly_src"] = ro_src

    # 3. Write-denied directory (NTFS ACL)
    ro_dir = fs_dir / "fs_readonly_dir"
    ro_dir.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        user = os.environ.get("USERNAME", "Everyone")
        try:
            subprocess.run(
                ["icacls", str(ro_dir), "/deny", f"{user}:(WD,AD)"],
                capture_output=True,
                text=True,
                check=True,
            )
        except Exception as e:
            logger.warning("Failed to apply deny ACL via icacls: %s", e)
    else:
        try:
            ro_dir.chmod(stat.S_IRUSR | stat.S_IXUSR)
        except Exception:
            pass
    fixtures["fs_readonly_dir"] = ro_dir

    # 4. File for locking tests
    lock_file = fs_dir / "fs_locked_target.pdf"
    if lock_file.exists():
        try:
            lock_file.chmod(stat.S_IWRITE)
        except Exception:
            pass
    _create_sample_pdf(lock_file, "Document reserved for cross-process file locking")
    fixtures["fs_locked_file"] = lock_file

    # 5. Unicode and Emoji filenames
    u_emoji = fs_dir / "📂_test_🧪_документ.pdf"
    _create_sample_pdf(u_emoji, "Unicode and Emoji document content")
    fixtures["fs_unicode_emoji_pdf"] = u_emoji

    u_cjk = fs_dir / "日本語_テスト_🚀_東京.pdf"
    _create_sample_pdf(u_cjk, "CJK Japanese test document content")
    fixtures["fs_unicode_cjk_pdf"] = u_cjk

    u_arabic = fs_dir / "العربية_ملف_اختبار_🔥.pdf"
    _create_sample_pdf(u_arabic, "Arabic RTL test document content")
    fixtures["fs_unicode_arabic_pdf"] = u_arabic

    u_nested = fs_dir / "📁_kategori_⚡" / "📄_laporan_🎯.pdf"
    u_nested.parent.mkdir(parents=True, exist_ok=True)
    _create_sample_pdf(u_nested, "Nested unicode directory and document content")
    fixtures["fs_unicode_nested_dir"] = u_nested

    # 6. Deeply nested long path exceeding 260 characters
    long_dir = fs_dir
    for i in range(5):
        long_dir = long_dir / (f"level_{i}_" + ("segment_" * 4))
    long_file = long_dir / "deeply_nested_long_path_document.pdf"

    long_dir_ext = to_extended_path(long_dir)
    long_file_ext = to_extended_path(long_file)

    os.makedirs(long_dir_ext, exist_ok=True)
    _create_sample_pdf(long_file_ext, "Long path document exceeding MAX_PATH (260 characters)")
    fixtures["fs_long_path_260plus"] = Path(long_file_ext)

    return fixtures


# ══════════════════════════════════════════════════════════════════════════════
# 6. MASTER AGGREGATE GENERATOR & CATALOG MANIFEST
# ══════════════════════════════════════════════════════════════════════════════

def write_catalog_manifest(
    fixtures: Dict[str, Path],
    manifest_path: Path,
    seed: int = 42,
    output_dir: Optional[Path] = None,
) -> Path:
    """
    Serializes a comprehensive JSON manifest cataloging all generated fixtures,
    their attributes (encryption, size, read-only status), and categorization.
    """
    manifest_path = Path(manifest_path).resolve()
    manifest_path.parent.mkdir(parents=True, exist_ok=True)

    catalog: Dict[str, Any] = {
        "generator_version": "1.0.0",
        "seed": seed,
        "generated_at": datetime.datetime.now(timezone.utc).isoformat(),
        "fixtures_count": len(fixtures),
        "output_directory": str(output_dir or manifest_path.parent),
        "fixtures": {},
    }

    for key, path in fixtures.items():
        p = Path(path)
        is_dir = p.is_dir()
        exists = p.exists()
        size_bytes = 0
        is_ro = False
        is_enc = False
        has_extended = str(path).startswith("\\\\?\\")

        if exists:
            try:
                size_bytes = 0 if is_dir else p.stat().st_size
            except Exception:
                size_bytes = 0

            # Read-only check
            try:
                mode = p.stat().st_mode
                is_ro = not bool(mode & stat.S_IWRITE)
            except Exception:
                is_ro = False

            # Encryption check
            if not is_dir and p.suffix.lower() == ".pdf":
                is_enc = is_pdf_encrypted(p)
            elif not is_dir and p.suffix.lower() == ".pptx":
                is_enc = is_office_encrypted(p)

        # Categorization heuristic
        category = "general"
        parent_name = p.parent.name.lower()
        if "extreme_pdfs" in parent_name or key.startswith(("extreme_", "high_", "colorspace_", "malformed_")):
            category = "extreme_pdfs"
        elif "encrypted_pdfs" in parent_name or key.startswith(("rc4_", "aes_", "pypdf_", "pdf_enc_")):
            category = "encrypted_pdfs"
        elif "presentations" in parent_name or key.startswith(("huge_", "massive_", "corrupt_", "password_", "pptx_")):
            category = "presentations"
        elif "filesystem" in parent_name or key.startswith("fs_"):
            category = "filesystem"

        catalog["fixtures"][key] = {
            "key": key,
            "category": category,
            "path": str(path),
            "size_bytes": size_bytes,
            "exists": exists,
            "is_dir": is_dir,
            "is_readonly": is_ro,
            "is_encrypted": is_enc,
            "is_locked": "locked" in key,
            "has_extended_path": has_extended,
        }

    manifest_path.write_text(json.dumps(catalog, indent=2, ensure_ascii=False), encoding="utf-8")
    return manifest_path


def generate_all_fixtures(
    target_dir: Path,
    seed: int = 42,
    manifest_path: Optional[Path] = None,
) -> Dict[str, Path]:
    """
    Master generator creating the complete synthetic stress test suite across
    all four categories organized in dedicated subdirectories.
    """
    target_dir = Path(target_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    all_fixtures: Dict[str, Path] = {}

    # 1. Extreme PDFs
    pdf_fixtures = generate_extreme_pdfs(target_dir / "extreme_pdfs", seed=seed)
    all_fixtures.update(pdf_fixtures)

    # 2. Encrypted PDFs
    enc_fixtures = generate_encrypted_pdfs(target_dir / "encrypted_pdfs", seed=seed)
    all_fixtures.update(enc_fixtures)

    # 3. Extreme Presentations
    ppt_fixtures = generate_extreme_presentations(target_dir / "presentations", seed=seed)
    all_fixtures.update(ppt_fixtures)

    # 4. Filesystem Edge Cases
    fs_fixtures = generate_filesystem_edge_cases(target_dir / "filesystem", seed=seed)
    all_fixtures.update(fs_fixtures)

    # Write manifest catalog
    m_path = manifest_path or (target_dir / "fixtures_manifest.json")
    write_catalog_manifest(all_fixtures, m_path, seed=seed, output_dir=target_dir)

    return all_fixtures


# ══════════════════════════════════════════════════════════════════════════════
# 7. COMMAND LINE INTERFACE (CLI)
# ══════════════════════════════════════════════════════════════════════════════

def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="python -m tests.stress.generate_stress_data",
        description="Synthetic Pathological Stress Data Generator for MarkItDown Tools",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=Path("tests/stress/fixtures"),
        help="Target directory for generated fixtures (default: tests/stress/fixtures)",
    )
    parser.add_argument(
        "-f",
        "--fixtures",
        nargs="*",
        default=["all"],
        help="Fixture categories or keys: 'all', 'extreme_pdfs', 'encrypted_pdfs', 'presentations', 'filesystem'",
    )
    parser.add_argument(
        "-c",
        "--clean",
        action="store_true",
        help="Purge target directory before generation",
    )
    parser.add_argument(
        "-s",
        "--seed",
        type=int,
        default=42,
        help="Integer seed for deterministic generation (default: 42)",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        default=None,
        help="Custom path for output JSON catalog manifest",
    )
    parser.add_argument(
        "--lock-file",
        type=Path,
        default=None,
        help="Utility daemon mode: Hold exclusive msvcrt byte lock on file until stdin closes",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable detailed verbose output",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Suppress output except summary",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    # If --lock-file is requested, operate immediately as a lock daemon
    if args.lock_file:
        run_lock_daemon(args.lock_file)
        return 0

    log_level = logging.DEBUG if args.verbose else (logging.WARNING if args.quiet else logging.INFO)
    logging.basicConfig(level=log_level, format="%(levelname)s: %(message)s")

    out_dir = args.output_dir.resolve()

    if args.clean and out_dir.exists():
        if not args.quiet:
            print(f"Cleaning existing output directory: {out_dir}")
        safe_rmtree(out_dir)

    out_dir.mkdir(parents=True, exist_ok=True)

    # Normalize category / fixture tokens
    tokens: List[str] = []
    for token in args.fixtures:
        for sub in token.split(","):
            s = sub.strip().lower()
            if s:
                tokens.append(s)

    run_all = "all" in tokens or not tokens
    selected_fixtures: Dict[str, Path] = {}

    start_time = time.perf_counter()

    if run_all:
        if not args.quiet:
            print(f"Generating full synthetic stress fixture suite into: {out_dir}")
        selected_fixtures = generate_all_fixtures(out_dir, seed=args.seed, manifest_path=args.manifest)
    else:
        # Category-based granular generation
        category_map = {
            "extreme_pdfs": (generate_extreme_pdfs, out_dir / "extreme_pdfs"),
            "pdf": (generate_extreme_pdfs, out_dir / "extreme_pdfs"),
            "pdfs": (generate_extreme_pdfs, out_dir / "extreme_pdfs"),
            "encrypted_pdfs": (generate_encrypted_pdfs, out_dir / "encrypted_pdfs"),
            "encryption": (generate_encrypted_pdfs, out_dir / "encrypted_pdfs"),
            "enc": (generate_encrypted_pdfs, out_dir / "encrypted_pdfs"),
            "presentations": (generate_extreme_presentations, out_dir / "presentations"),
            "extreme_presentations": (generate_extreme_presentations, out_dir / "presentations"),
            "office": (generate_extreme_presentations, out_dir / "presentations"),
            "ppt": (generate_extreme_presentations, out_dir / "presentations"),
            "pptx": (generate_extreme_presentations, out_dir / "presentations"),
            "filesystem": (generate_filesystem_edge_cases, out_dir / "filesystem"),
            "filesystem_edge_cases": (generate_filesystem_edge_cases, out_dir / "filesystem"),
            "fs": (generate_filesystem_edge_cases, out_dir / "filesystem"),
        }

        ran_categories = set()
        for tok in tokens:
            if tok in category_map and tok not in ran_categories:
                gen_fn, sub_dir = category_map[tok]
                if not args.quiet:
                    print(f"Generating fixture category: {tok} -> {sub_dir}")
                fixtures = gen_fn(sub_dir, seed=args.seed)
                selected_fixtures.update(fixtures)
                ran_categories.add(tok)

        m_path = args.manifest or (out_dir / "fixtures_manifest.json")
        write_catalog_manifest(selected_fixtures, m_path, seed=args.seed, output_dir=out_dir)

    duration = time.perf_counter() - start_time
    total_bytes = sum(
        p.stat().st_size
        for p in set(selected_fixtures.values())
        if p.exists() and not p.is_dir()
    )

    if not args.quiet:
        print("─" * 60)
        print(f"✓ Stress fixture generation complete in {duration:.3f} seconds.")
        print(f"✓ Total unique fixture items: {len(set(selected_fixtures.values()))}")
        print(f"✓ Total fixture payload size: {total_bytes / (1024 * 1024):.2f} MB")
        manifest_loc = args.manifest or (out_dir / "fixtures_manifest.json")
        print(f"✓ Manifest catalog written: {manifest_loc}")
        print("─" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
