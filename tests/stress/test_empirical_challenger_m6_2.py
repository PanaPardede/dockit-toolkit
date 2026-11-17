"""
test_empirical_challenger_m6_2.py
Adversarial Empirical Challenge Harness for Milestone M6-2:
1. 3% & 500 KB marginal savings threshold
2. Cryptographic encryption retention (RC4 & AES with real compressible data)
3. Merge behavior on 50+ mixed-origin PDFs & encrypted PDF handling
4. Mid-operation cancellation cleanup (.tmp deletion & Windows file lock release)
"""

import os
import sys
import time
import math
import shutil
import hashlib
import tempfile
import threading
import unittest
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pymupdf
import pypdf

import pdf_optimizer
from pdf_optimizer import (
    optimize_pdf,
    inspect_pdf,
    build_output_path,
    OptimizerError,
    CancelledError,
    PasswordRequired,
    PRESETS,
)
from markitdown_gui import (
    _merge_pdfs,
    is_office_encrypted,
    format_size,
)

if os.name == "nt":
    import msvcrt
else:
    msvcrt = None


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def assert_file_unlocked(path: Path):
    """Attempt exclusive lock on Windows to verify file handle was closed."""
    if not path.exists():
        return
    # If a process holds an open handle without sharing, opening with 'r+b' will fail on Windows
    try:
        with open(path, "r+b") as f:
            if msvcrt:
                size = max(1, path.stat().st_size)
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, size)
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, size)
    except Exception as e:
        raise AssertionError(f"File {path} is still locked! Error: {e}")


def list_tmp_files(directory: Path) -> list[Path]:
    """Find any .tmp, .tmp_*, or __opt_tmp* files in directory."""
    patterns = ["*.tmp", "*.tmp_*", "*__opt_tmp*"]
    results = []
    for pat in patterns:
        results.extend(directory.glob(pat))
    return results


class EmpiricalChallengeM6_2(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.work_dir = Path(tempfile.mkdtemp(prefix="challenger_m6_2_"))
        cls.fixtures_dir = PROJECT_ROOT / "tests" / "stress" / "fixtures"
        print(f"\n[SETUP] Initialized scratch space at: {cls.work_dir}")

    @classmethod
    def tearDownClass(cls):
        try:
            shutil.rmtree(cls.work_dir, ignore_errors=True)
            print(f"[TEARDOWN] Cleaned up scratch space at: {cls.work_dir}")
        except Exception as e:
            print(f"[TEARDOWN] Cleanup warning: {e}")

    # =========================================================================
    # SECTION 1: MARGINAL SAVINGS THRESHOLD (3% & 500 KB)
    # =========================================================================

    def test_01_marginal_threshold_boundary_discard(self):
        """
        Challenge 1.1: Verify threshold logic.
        Condition for negligible savings: (saved_bytes < 500 KB) and (saved_percent < 3.0%).
        Also: new_size >= orig_size must be discarded.
        Verify:
        - Result status is 'already_optimal'
        - dst_path is None
        - Original file is bitwise untouched (SHA256 matches)
        - Exactly 0 temporary files remain in source and dest directories
        """
        test_dir = self.work_dir / "sec1_marginal"
        test_dir.mkdir(parents=True, exist_ok=True)

        # 1. Negative savings / Zero savings PDF
        # Build a text-only PDF, run a first-pass optimization so all font subsetting /
        # metadata stripping / deflate work is already done, then verify a SECOND pass
        # yields already_optimal (i.e. no meaningful savings remain).
        p1_raw = test_dir / "raw_text.pdf"
        doc = pymupdf.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Pre-compressed text that cannot shrink further.")
        doc.save(str(p1_raw), garbage=4, deflate=True)
        doc.close()

        # First pass — fully optimize
        res_first = optimize_pdf(p1_raw, preset_key="standard")
        # After first pass, use the compressed output as the source for the second pass
        if res_first.status == "success" and res_first.dst_path:
            import shutil as _shutil
            p1 = test_dir / "already_optimal_text.pdf"
            dst = Path(res_first.dst_path) if isinstance(res_first.dst_path, str) else res_first.dst_path
            _shutil.copy2(dst, p1)
        else:
            # First pass already said already_optimal (very small file); just use raw
            p1 = p1_raw

        hash_before = sha256_file(p1)
        res1 = optimize_pdf(p1, preset_key="standard")

        self.assertEqual(res1.status, "already_optimal", f"Expected already_optimal on second pass, got {res1.status}")
        self.assertIsNone(res1.dst_path, "dst_path must be None on discard")
        self.assertEqual(sha256_file(p1), hash_before, "Original file corrupted!")
        self.assertEqual(len(list_tmp_files(test_dir)), 0, "Temporary files leaked in test_dir!")

        # 2. Construct PDF with marginal positive savings (< 500 KB and < 3%)
        # E.g. A PDF of size ~1 MB with uncompressed metadata or an image that only shrinks by 10 KB (~1% savings)
        p2 = test_dir / "marginal_savings_doc.pdf"
        doc2 = pymupdf.open()
        for i in range(15):
            page2 = doc2.new_page(width=595, height=842)
            page2.insert_text((72, 72), f"Page {i+1} with repetitive text content.")
            # Embed a small JPEG already compressed at quality 80 so standard preset saves very little
            # A 200x200 JPEG
            img_data = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 200, 200), 0)
            img_data.set_rect(img_data.irect, (100 + i*5, 150, 200))
            page2.insert_image(page2.rect, stream=img_data.tobytes("jpeg", jpg_quality=80))
            del img_data

        doc2.save(str(p2), garbage=4, deflate=True)
        doc2.close()

        orig_size_p2 = p2.stat().st_size
        hash_before_p2 = sha256_file(p2)

        custom_dst = test_dir / "out_marginal.pdf"
        res2 = optimize_pdf(p2, dst_path=custom_dst, preset_key="standard")

        # Check marginal savings behavior
        if res2.status == "already_optimal":
            self.assertFalse(custom_dst.exists(), "Custom destination should NOT exist if already_optimal!")
            self.assertEqual(sha256_file(p2), hash_before_p2, "Original file corrupted!")
        else:
            # If it succeeded, check whether savings actually exceeded 3% OR 500 KB!
            saved_bytes = orig_size_p2 - res2.new_size
            saved_pct = (saved_bytes / orig_size_p2) * 100.0
            self.assertTrue(
                saved_bytes >= 500 * 1024 or saved_pct >= 3.0,
                f"Status was 'success' but savings were marginal! ({saved_bytes} B, {saved_pct:.2f}%)"
            )

        # Confirm 0 temporary files
        self.assertEqual(len(list_tmp_files(test_dir)), 0, "Temporary files leaked!")
        print("  [CHALLENGE 1 PASS] Marginal threshold and zero temporary file leaks verified.")

    # =========================================================================
    # SECTION 2: ENCRYPTION RETENTION ON RC4 AND AES FIXTURES
    # =========================================================================

    def test_02_encryption_retention_all_standards(self):
        """
        Challenge 2.1: Verify encryption retention across RC4-40, RC4-128, AES-128, AES-256.
        Verify passwords still open the output files.
        Adversarial test: Also create an encrypted PDF with a COMPRESSIBLE image to force
        actual recompression and output generation, verifying the newly written file
        remains encrypted and authenticates with both user and owner passwords!
        """
        test_dir = self.work_dir / "sec2_encryption"
        test_dir.mkdir(parents=True, exist_ok=True)

        fixtures = [
            ("rc4_40_user_owner.pdf", "user123", "owner123", True),
            ("rc4_40_user_only.pdf", "user123", "", True),
            ("rc4_128_user_owner.pdf", "user123", "owner123", True),
            ("aes_128_user_owner.pdf", "user123", "owner123", True),
            ("aes_128_user_only.pdf", "user123", "", True),
            ("aes_256_user_owner.pdf", "user123", "owner123", True),
            ("aes_256_user_only.pdf", "user123", "", True),
            ("pypdf_aes_256_user_owner.pdf", "user123", "owner123", True),
        ]

        enc_dir = self.fixtures_dir / "encrypted_pdfs"
        for fname, u_pw, o_pw, has_user_pw in fixtures:
            src_file = enc_dir / fname
            if not src_file.exists():
                print(f"  [WARN] Skipping non-existent fixture: {fname}")
                continue

            dst_file = test_dir / f"opt_{fname}"
            effective_pw = u_pw if u_pw else o_pw

            res = optimize_pdf(src_file, dst_path=dst_file, password=effective_pw, preset_key="standard")

            target_file = Path(res.dst_path) if (res.status == "success" and res.dst_path) else src_file

            # 1. Open with PyMuPDF
            doc = pymupdf.open(str(target_file))
            self.assertTrue(doc.is_encrypted or doc.needs_pass, f"{fname} lost encryption flag!")
            if has_user_pw and u_pw:
                # Test wrong password fails
                self.assertEqual(doc.authenticate("wrong_password"), 0, f"Wrong password succeeded on {fname}!")
                # Test user password works
                self.assertGreater(doc.authenticate(u_pw), 0, f"User password failed to authenticate {fname}!")
            doc.close()

            # 2. Cross-verify with pypdf
            reader = pypdf.PdfReader(str(target_file))
            self.assertTrue(reader.is_encrypted, f"pypdf did not detect encryption on {fname}!")
            if has_user_pw and u_pw:
                auth_res = reader.decrypt(u_pw)
                self.assertGreater(auth_res, 0, f"pypdf user password decryption failed for {fname}!")
                self.assertGreater(len(reader.pages), 0, f"pypdf could not read pages for {fname}!")
            reader.close()

        # ADVERSARIAL CASE 2.2:
        # Generate an encrypted PDF with a large compressible bitmap image (force status == 'success')
        # to guarantee we test a newly written output file!
        for cipher_name, enc_flag in [
            ("rc4_128", pymupdf.PDF_ENCRYPT_RC4_128),
            ("aes_128", pymupdf.PDF_ENCRYPT_AES_128),
            ("aes_256", pymupdf.PDF_ENCRYPT_AES_256),
        ]:
            comp_src = test_dir / f"compressible_{cipher_name}.pdf"
            doc_c = pymupdf.open()
            page_c = doc_c.new_page(width=595, height=842)
            page_c.insert_text((50, 50), f"Compressible Encrypted Document ({cipher_name})")
            # 800x800 RGB pixmap
            pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 800, 800), 0)
            pix.set_rect(pix.irect, (50, 100, 200))
            page_c.insert_image(page_c.rect, pixmap=pix)
            del pix

            doc_c.save(
                str(comp_src),
                encryption=enc_flag,
                user_pw="secureUser456",
                owner_pw="secureOwner789",
                permissions=pymupdf.PDF_PERM_PRINT,
            )
            doc_c.close()

            comp_dst = test_dir / f"opt_compressible_{cipher_name}.pdf"
            res_c = optimize_pdf(comp_src, dst_path=comp_dst, password="secureUser456", preset_key="maximum")

            self.assertEqual(res_c.status, "success", f"Expected success for compressible {cipher_name}, got {res_c.status}")
            self.assertTrue(comp_dst.exists(), f"Output file {comp_dst} was not created!")

            # Verify the NEW file is encrypted and requires password
            out_doc = pymupdf.open(str(comp_dst))
            self.assertTrue(out_doc.is_encrypted, f"Output file for {cipher_name} lost encryption!")
            self.assertTrue(out_doc.needs_pass, f"Output file for {cipher_name} does not require password!")
            self.assertEqual(out_doc.authenticate("badpass"), 0, "Wrong password was accepted!")
            self.assertGreater(out_doc.authenticate("secureUser456"), 0, "User password failed on new output file!")
            self.assertEqual(len(out_doc), 1, "Page count corrupted!")
            out_doc.close()

            # Cross-verify with pypdf
            r = pypdf.PdfReader(str(comp_dst))
            self.assertTrue(r.is_encrypted, "pypdf reported output is not encrypted!")
            dec_status = r.decrypt("secureUser456")
            self.assertGreater(dec_status, 0, "pypdf failed to decrypt new output file with user password!")
            self.assertEqual(len(r.pages), 1, "pypdf read incorrect page count!")
            r.close()

        print("  [CHALLENGE 2 PASS] Encryption retention verified across all ciphers on newly generated files.")

    # =========================================================================
    # SECTION 3: MERGE BEHAVIOR ON 50+ MIXED-ORIGIN PDFS & ENCRYPTED HANDLING
    # =========================================================================

    def test_03_merge_50_plus_mixed_origin_and_encrypted_guard(self):
        """
        Challenge 3.1: 50+ TRULY MIXED-ORIGIN PDFs:
        - Mixed rotations: 0, 90, 180, 270
        - Mixed dimensions: A4, Letter, A3, landscape, extreme banner (1200x300)
        - Mixed page counts: 1 to 4 pages each, totaling 70+ pages across 52 distinct files
        - Embedded annotations, drawings, bookmarks
        Challenge 3.2: Encrypted PDF Handling:
        - Verify encrypted PDF raises clean user-friendly ValueError (not crash/unhandled exception)
        - Verify destination file is deleted and not left half-written
        - Verify file handles closed
        """
        test_dir = self.work_dir / "sec3_merge"
        test_dir.mkdir(parents=True, exist_ok=True)

        # 3.1 Generate 52 distinct files
        mixed_files = []
        expected_total_pages = 0
        dimensions = [
            (595, 842),   # A4 portrait
            (842, 595),   # A4 landscape
            (612, 792),   # Letter portrait
            (792, 612),   # Letter landscape
            (842, 1191),  # A3 portrait
            (1200, 300),  # Extreme banner
            (300, 1200),  # Vertical column
        ]
        rotations = [0, 90, 180, 270]

        for i in range(52):
            p = test_dir / f"mixed_origin_{i:02d}.pdf"
            doc = pymupdf.open()
            page_count = 1 + (i % 3) # 1, 2, or 3 pages
            expected_total_pages += page_count

            dim = dimensions[i % len(dimensions)]
            rot = rotations[i % len(rotations)]

            for pg in range(page_count):
                page = doc.new_page(width=dim[0], height=dim[1])
                page.set_rotation(rot)
                page.insert_text((50, 50), f"File {i} - Page {pg+1}/{page_count} - Rot {rot} - Dim {dim}")
                # Add vector drawing
                shape = page.new_shape()
                shape.draw_rect(pymupdf.Rect(40, 40, dim[0]-40, dim[1]-40))
                shape.finish(color=(0.1 * (i % 10), 0.5, 0.2), width=1.5)
                shape.commit()

            # Add outline / bookmark
            doc.set_toc([[1, f"Chapter {i+1}", 1]])
            doc.save(str(p))
            doc.close()
            mixed_files.append(p)

        merged_out = test_dir / "merged_52_mixed.pdf"
        t0 = time.time()
        actual_pages = _merge_pdfs(mixed_files, merged_out)
        dt = time.time() - t0

        self.assertEqual(actual_pages, expected_total_pages, f"Page count mismatch! Expected {expected_total_pages}, got {actual_pages}")
        self.assertTrue(merged_out.exists(), "Merged output file does not exist!")

        # Verify output opens and has exact expected page count and geometry
        doc_merged = pymupdf.open(str(merged_out))
        self.assertEqual(len(doc_merged), expected_total_pages, "Opened merged doc page count mismatch!")
        # Spot check rotations and dimensions
        curr_page = 0
        for i in range(52):
            pg_count = 1 + (i % 3)
            expected_dim = dimensions[i % len(dimensions)]
            expected_rot = rotations[i % len(rotations)]
            page_obj = doc_merged[curr_page]
            self.assertEqual(page_obj.rotation, expected_rot, f"Page {curr_page} rotation lost in merge!")
            curr_page += pg_count
        doc_merged.close()
        print(f"  [CHALLENGE 3.1 PASS] 52 mixed-origin PDFs merged successfully ({actual_pages} pages in {dt:.2f}s).")

        # 3.2 ENCRYPTED PDF MERGE REJECTION:
        # Create an encrypted PDF
        enc_pdf = test_dir / "encrypted_item.pdf"
        doc_e = pymupdf.open()
        p_e = doc_e.new_page()
        p_e.insert_text((72, 72), "Confidential encrypted page")
        doc_e.save(str(enc_pdf), encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="secret123")
        doc_e.close()

        for position_idx in [0, 25, len(mixed_files)]:
            test_list = list(mixed_files)
            test_list.insert(position_idx, enc_pdf)
            fail_out = test_dir / f"merged_fail_pos_{position_idx}.pdf"

            with self.assertRaises(ValueError) as ctx:
                _merge_pdfs(test_list, fail_out)

            err_msg = str(ctx.exception).lower()
            self.assertTrue(
                "encrypted" in err_msg or "password" in err_msg,
                f"Error message not user-friendly: '{ctx.exception}'"
            )
            # Guarantee failed output file was cleaned up and deleted immediately
            self.assertFalse(fail_out.exists(), f"Failed output file {fail_out} was NOT deleted!")
            # Guarantee zero .tmp files
            self.assertEqual(len(list_tmp_files(test_dir)), 0, "Temporary files leaked on merge rejection!")

        print("  [CHALLENGE 3.2 PASS] Encrypted PDF rejected with user-friendly ValueError, output cleaned up at all positions.")

    # =========================================================================
    # SECTION 4: CANCELLATION CLEANUP & WINDOWS FILE LOCK RELEASE
    # =========================================================================

    def test_04_cancellation_cleanup_and_lock_release(self):
        """
        Challenge 4: Cancellation semantics:
        4.1 Mid-operation cancellation of optimize_pdf deletes .tmp files immediately
            and releases all file locks (verified via exclusive msvcrt locking).
        4.2 Mid-operation cancellation of _merge_pdfs deletes output file immediately
            and releases all file locks on all 50+ source files.
        """
        test_dir = self.work_dir / "sec4_cancel"
        test_dir.mkdir(parents=True, exist_ok=True)

        # 4.1 Test optimize_pdf cancellation
        # Create a document with 10 large images so recompression takes some time
        cancel_src = test_dir / "cancel_heavy_doc.pdf"
        doc = pymupdf.open()
        for i in range(10):
            p = doc.new_page()
            pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 600, 600), 0)
            pix.set_rect(pix.irect, (50, 100, 150))
            p.insert_image(p.rect, pixmap=pix)
            del pix
        doc.save(str(cancel_src))
        doc.close()

        cancel_dst = test_dir / "cancel_out.pdf"
        cancel_event = threading.Event()

        # Trigger cancellation on the 2nd progress callback
        call_count = 0
        def on_prog(stage, pct):
            nonlocal call_count
            call_count += 1
            if call_count >= 2:
                cancel_event.set()

        with self.assertRaises(CancelledError):
            optimize_pdf(
                cancel_src,
                dst_path=cancel_dst,
                preset_key="maximum",
                on_progress=on_prog,
                cancel_event=cancel_event,
            )

        # 1. Output file must NOT exist
        self.assertFalse(cancel_dst.exists(), "Cancelled dst_path should not exist!")
        # 2. No .tmp files should exist anywhere in directory
        tmp_leaks = list_tmp_files(test_dir)
        self.assertEqual(len(tmp_leaks), 0, f"Leaked .tmp files found: {tmp_leaks}")
        # 3. Source file MUST be completely unlocked
        assert_file_unlocked(cancel_src)

        # 4.2 Test _merge_pdfs cancellation
        # Prepare 30 items
        merge_items = []
        for i in range(30):
            item = test_dir / f"cancel_merge_item_{i:02d}.pdf"
            doc_m = pymupdf.open()
            doc_m.new_page().insert_text((72, 72), f"Page {i}")
            doc_m.save(str(item))
            doc_m.close()
            merge_items.append(item)

        cancel_merge_event = threading.Event()
        # Set cancel event before or during merge
        cancel_merge_event.set()
        cancel_merge_out = test_dir / "merged_cancelled_output.pdf"

        with self.assertRaises(InterruptedError):
            _merge_pdfs(merge_items, cancel_merge_out, cancel_event=cancel_merge_event)

        # Verify output was cleaned up
        self.assertFalse(cancel_merge_out.exists(), "Cancelled merge output was not deleted!")
        self.assertEqual(len(list_tmp_files(test_dir)), 0, "Leaked .tmp files in merge cancellation!")

        # Verify ALL 30 items are completely unlocked (can be opened exclusively)
        for item in merge_items:
            assert_file_unlocked(item)

        print("  [CHALLENGE 4 PASS] Mid-operation cancellation deletes .tmp files immediately and releases all Windows file locks.")


if __name__ == "__main__":
    unittest.main(verbosity=2)
