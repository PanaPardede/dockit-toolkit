"""
test_pdf_optimizer.py — Unit test komprehensif untuk core engine pdf_optimizer.py.
"""

import os
import shutil
import tempfile
import hashlib
import threading
import unittest
from pathlib import Path

import pymupdf
from PIL import Image

import pdf_optimizer
from pdf_optimizer import (
    PRESETS,
    inspect_pdf,
    build_output_path,
    optimize_pdf,
    validate_output,
    PasswordRequired,
    CorruptPdf,
    CancelledError,
    DigitalSignatureWarning,
    has_digital_signatures,
)


def get_file_hash(path: Path) -> str:
    """Menghitung SHA-256 file."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


class TestPdfOptimizer(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="test_pdf_opt_"))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_text_pdf(self, filename="sample_text.pdf", pages=3) -> Path:
        p = self.temp_dir / filename
        doc = pymupdf.open()
        for i in range(pages):
            page = doc.new_page(width=595, height=842)
            page.insert_text((72, 72), f"Halaman {i+1} — Dokumen Uji Teks Vektor.\n"
                                       "Teks ini harus tetap dapat dicari dan diseleksi 100%.",
                             fontsize=12)
        doc.save(str(p))
        doc.close()
        return p

    def _create_image_heavy_pdf(self, filename="sample_images.pdf", img_size=(1600, 1200)) -> Path:
        """Membuat PDF dengan gambar beresolusi tinggi (> 300 DPI saat ditampilkan)."""
        p = self.temp_dir / filename
        doc = pymupdf.open()
        
        # Buat dummy foto RGB dengan variasi warna
        img = Image.new("RGB", img_size, color=(100, 150, 200))
        # Isi dengan pattern agar tidak terkompresi terlalu datar oleh flate
        for x in range(0, img_size[0], 20):
            for y in range(0, img_size[1], 20):
                img.putpixel((x, y), ((x*3)%256, (y*5)%256, ((x+y)*7)%256))
        
        img_bytes = io = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        io.close()
        img.save(img_bytes.name, format="PNG")

        try:
            for i in range(2):
                page = doc.new_page(width=595, height=842)
                page.insert_text((50, 50), f"Halaman Gambar {i+1}", fontsize=14)
                # Tampilkan selebar 300 pt (~4.1 inci).
                # 1600 px / 4.16 inci = ~384 DPI efektif!
                page.insert_image(pymupdf.Rect(50, 100, 350, 500), filename=img_bytes.name)
            doc.save(str(p))
        finally:
            doc.close()
            os.remove(img_bytes.name)
        return p

    def test_preset_definitions(self):
        self.assertIn("light", PRESETS)
        self.assertIn("standard", PRESETS)
        self.assertIn("maximum", PRESETS)

        std = PRESETS["standard"]
        self.assertEqual(std.target_dpi, 150)
        self.assertEqual(std.dpi_threshold, 180)
        self.assertEqual(std.jpeg_quality, 80)
        self.assertTrue(std.subset_fonts)

        max_p = PRESETS["maximum"]
        self.assertEqual(max_p.target_dpi, 110)
        self.assertEqual(max_p.dpi_threshold, 130)
        self.assertEqual(max_p.jpeg_quality, 65)

        light = PRESETS["light"]
        self.assertIsNone(light.target_dpi)
        self.assertFalse(light.subset_fonts)

    def test_build_output_path(self):
        src = self.temp_dir / "document.pdf"
        src.touch()

        # Kasus 1: Belum ada file compressed
        out1 = build_output_path(src)
        self.assertEqual(out1.name, "document_compressed.pdf")
        out1.touch()

        # Kasus 2: Sudah ada document_compressed.pdf
        out2 = build_output_path(src)
        self.assertEqual(out2.name, "document_compressed (2).pdf")
        out2.touch()

        # Kasus 3: Sudah ada (2), buat (3)
        out3 = build_output_path(src)
        self.assertEqual(out3.name, "document_compressed (3).pdf")

    def test_inspect_valid_pdf(self):
        src = self._create_text_pdf("inspect_test.pdf", pages=4)
        info = inspect_pdf(src)
        self.assertTrue(info.is_valid)
        self.assertEqual(info.page_count, 4)
        self.assertGreater(info.file_size, 0)
        self.assertFalse(info.needs_password)
        self.assertFalse(info.has_signatures)

    def test_inspect_corrupt_file(self):
        bad_file = self.temp_dir / "corrupt.pdf"
        bad_file.write_text("bukan pdf sungguhan sama sekali")
        with self.assertRaises(CorruptPdf):
            inspect_pdf(bad_file)

    def test_optimization_preserves_text_and_pages(self):
        """AC-06, AC-11, AC-13: Hash asli utuh, teks identik, jumlah halaman sama."""
        src = self._create_image_heavy_pdf("heavy_doc.pdf")
        hash_before = get_file_hash(src)
        doc_before = pymupdf.open(str(src))
        text_before = [doc_before[i].get_text() for i in range(len(doc_before))]
        page_count_before = len(doc_before)
        doc_before.close()

        res = optimize_pdf(src, preset_key="standard")
        self.assertEqual(res.status, "success")
        self.assertIsNotNone(res.dst_path)
        out_path = Path(res.dst_path)
        self.assertTrue(out_path.exists())

        # AC-06: File asli tidak berubah sama sekali
        hash_after = get_file_hash(src)
        self.assertEqual(hash_before, hash_after)

        # AC-13: Jumlah halaman sama
        doc_after = pymupdf.open(str(out_path))
        self.assertEqual(len(doc_after), page_count_before)

        # AC-11: Teks tidak berubah
        text_after = [doc_after[i].get_text() for i in range(len(doc_after))]
        self.assertEqual(text_before, text_after)
        doc_after.close()

        # AC-17: Penghematan minimal 40% pada gambar berat resolusi tinggi
        self.assertGreater(res.saved_percent, 40.0)
        self.assertLess(res.new_size, res.orig_size)

    def test_already_optimal_threshold(self):
        """FR-06 / E1 / AC-19: Jika penghematan < 3% dan < 500 KB, file hasil dibuang."""
        # PDF yang sudah dikompresi sebelumnya (corpus uji e / AC-19)
        src = self.temp_dir / "already_optimal.pdf"
        doc = pymupdf.open()
        p = doc.new_page()
        p.insert_text((50, 50), "Dokumen yang sudah sangat optimal.")
        doc.save(str(src), garbage=4, deflate=True, clean=True, use_objstms=True)
        doc.close()

        res = optimize_pdf(src, preset_key="standard")

        # Harus mengembalikan already_optimal
        self.assertEqual(res.status, "already_optimal")
        self.assertIsNone(res.dst_path)
        # Pastikan tidak ada file _compressed yang dibuat di folder
        compressed_files = list(self.temp_dir.glob("*_compressed*.pdf"))
        self.assertEqual(len(compressed_files), 0)

    def test_cancellation(self):
        """AC-10 / E14: Pembatalan menghentikan proses dan membersihkan file sementara."""
        src = self._create_image_heavy_pdf("cancel_test.pdf")
        cancel_evt = threading.Event()
        cancel_evt.set()  # batalkan segera

        with self.assertRaises(CancelledError):
            optimize_pdf(src, cancel_event=cancel_evt)

        # Pastikan tidak ada file sementara .tmp_* tertinggal
        tmp_files = list(self.temp_dir.glob("*.tmp_*"))
        self.assertEqual(len(tmp_files), 0)

    def test_password_protected_pdf(self):
        """E2, E3: PDF berpassword dan verifikasi enkripsi dipertahankan pada output."""
        src = self.temp_dir / "protected.pdf"
        doc = pymupdf.open()
        p = doc.new_page()
        p.insert_text((50, 50), "Rahasia negara tingkat tinggi")
        doc.save(str(src), encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="kunci123")
        doc.close()

        # Tanpa password harus raise PasswordRequired
        with self.assertRaises(PasswordRequired):
            optimize_pdf(src, password=None)

        # Dengan password salah harus raise PasswordRequired
        with self.assertRaises(PasswordRequired):
            optimize_pdf(src, password="salah")

        # Dengan password benar harus berhasil diinspeksi
        info = inspect_pdf(src, password="kunci123")
        self.assertTrue(info.is_valid)
        self.assertEqual(info.page_count, 1)

        # Optimasi dengan password yang benar
        res = optimize_pdf(src, password="kunci123", preset_key="light")
        self.assertIn(res.status, ("success", "already_optimal"))
        if res.status == "success" and res.dst_path:
            # Verifikasi output tetap terenkripsi dan butuh password yang sama
            out_doc = pymupdf.open(res.dst_path)
            self.assertTrue(out_doc.needs_pass, "File output harus tetap membutuhkan password!")
            # Autentikasi dengan password yang benar harus berhasil
            auth_ok = out_doc.authenticate("kunci123")
            self.assertTrue(auth_ok > 0, "Password asli harus bisa membuka file output!")
            self.assertEqual(len(out_doc), 1)
            out_doc.close()

    def test_cmyk_image_conversion(self):
        """E8: Gambar CMYK dikonversi ke RGB."""
        src = self.temp_dir / "cmyk_test.pdf"
        doc = pymupdf.open()
        page = doc.new_page(width=500, height=500)
        cmyk_pix = pymupdf.Pixmap(pymupdf.csCMYK, pymupdf.IRect(0, 0, 400, 400), 0)
        cmyk_pix.clear_with(200)
        page.insert_image(pymupdf.Rect(50, 50, 200, 200), pixmap=cmyk_pix)
        doc.save(str(src))
        doc.close()

        res = optimize_pdf(src, preset_key="standard")
        self.assertIn(res.status, ("success", "already_optimal"))
        # Pastikan gambar CMYK terdeteksi dan diproses
        if res.status == "success":
            self.assertGreaterEqual(res.cmyk_converted, 1)

    def test_transparent_image_smask_skipped(self):
        """E7 / AC-15: Gambar bertransparansi (SMask) dilewati."""
        src = self.temp_dir / "smask_test.pdf"
        doc = pymupdf.open()
        page = doc.new_page(width=400, height=400)
        
        # Buat image dengan alpha mask
        rgba_pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 100, 100), 1)
        rgba_pix.clear_with(150)
        page.insert_image(pymupdf.Rect(50, 50, 250, 250), pixmap=rgba_pix)
        doc.save(str(src))
        doc.close()

        res = optimize_pdf(src, preset_key="standard")
        # Karena gambar memiliki transparansi, gambar dilewati
        self.assertGreaterEqual(res.images_skipped, 1)

    def test_light_preset_does_not_touch_images(self):
        """Preset Ringan tidak menyentuh gambar sama sekali."""
        src = self._create_image_heavy_pdf("light_test.pdf")
        res = optimize_pdf(src, preset_key="light")
        self.assertEqual(res.images_processed, 0)

    def test_digital_signature_warning(self):
        """E10: Dokumen bertanda tangan memicu DigitalSignatureWarning kecuali di-ignore."""
        src = self.temp_dir / "sig_test.pdf"
        doc = pymupdf.open()
        p = doc.new_page()
        # Tambahkan form field widget tipe signature
        w = pymupdf.Widget()
        w.rect = pymupdf.Rect(50, 50, 200, 100)
        w.field_type = getattr(pymupdf, "PDF_WIDGET_TYPE_SIGNATURE", 6)
        w.field_name = "Signature1"
        p.add_widget(w)
        doc.save(str(src))
        doc.close()

        # Harus raise DigitalSignatureWarning jika ignore_signatures=False
        with self.assertRaises(DigitalSignatureWarning):
            optimize_pdf(src, ignore_signatures=False)

        # Harus berhasil diproses jika ignore_signatures=True
        res = optimize_pdf(src, ignore_signatures=True)
        self.assertIn(res.status, ("success", "already_optimal"))

    def test_validate_output_function(self):
        src = self._create_text_pdf("valid_check.pdf", pages=3)
        self.assertTrue(validate_output(3, src))
        self.assertFalse(validate_output(4, src))  # jumlah halaman salah
        self.assertFalse(validate_output(3, self.temp_dir / "tidak_ada.pdf"))

    def test_corpus_maximum_preset_heavy_image(self):
        """AC-17: File berat gambar (a) pada preset Maksimal menyusut >= 60%."""
        src = self._create_image_heavy_pdf("heavy_max.pdf")
        res = optimize_pdf(src, preset_key="maximum")
        self.assertEqual(res.status, "success")
        self.assertGreaterEqual(res.saved_percent, 60.0)

    def test_corpus_scanned_pdf(self):
        """AC-17: File scan (b) 5 halaman dengan resolusi scan menyusut >= 40% pada Standar."""
        src = self.temp_dir / "scan_corpus.pdf"
        doc = pymupdf.open()
        # Buat dummy scan 800x1000 pada A4 (595x842) -> ~100 DPI
        # Atau 1600x2000 pada A4 -> ~200 DPI
        scan_img = Image.new("RGB", (1200, 1600), color=(240, 240, 240))
        for y in range(0, 1600, 30):
            for x in range(50, 1150, 50):
                scan_img.putpixel((x, y), (20, 20, 20))
        tmp_img = self.temp_dir / "scan_tmp.png"
        scan_img.save(str(tmp_img), format="PNG")
        try:
            for i in range(5):
                page = doc.new_page(width=595, height=842)
                page.insert_image(pymupdf.Rect(0, 0, 595, 842), filename=str(tmp_img))
            doc.save(str(src))
        finally:
            doc.close()
            tmp_img.unlink(missing_ok=True)

        res = optimize_pdf(src, preset_key="standard")
        self.assertEqual(res.status, "success")
        self.assertGreaterEqual(res.saved_percent, 40.0)

    def test_corpus_pure_text_light_preset(self):
        """AC-18: File teks murni (c) pada Ringan tidak membesar dan sudah optimal pada run ke-2."""
        src = self._create_text_pdf("pure_text_corpus.pdf", pages=5)
        orig_size = src.stat().st_size
        res1 = optimize_pdf(src, preset_key="light")
        if res1.status == "success":
            self.assertLessEqual(res1.new_size, orig_size, "File hasil Ringan tidak boleh lebih besar!")
            # Run ke-2 harus already_optimal
            res2 = optimize_pdf(res1.dst_path, preset_key="light")
            self.assertEqual(res2.status, "already_optimal")
        else:
            self.assertEqual(res1.status, "already_optimal")

    def test_encryption_retention_all_standards(self):
        """M2 Deliverable: Memverifikasi preservasi PDF_ENCRYPT_KEEP untuk RC4-40, RC4-128, AES-128, AES-256."""
        standards = [
            (pymupdf.PDF_ENCRYPT_RC4_40, "RC4_40"),
            (pymupdf.PDF_ENCRYPT_RC4_128, "RC4_128"),
            (pymupdf.PDF_ENCRYPT_AES_128, "AES_128"),
            (pymupdf.PDF_ENCRYPT_AES_256, "AES_256"),
        ]

        for method, name in standards:
            src = self.temp_dir / f"enc_{name}.pdf"
            doc = pymupdf.open()
            page = doc.new_page(width=400, height=400)
            page.insert_text((50, 50), f"Testing encryption retention: {name}")
            # Masukkan gambar uncompressed agar kompresi aktif
            pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 500, 500), 0)
            pix.clear_with(150)
            page.insert_image(pymupdf.Rect(50, 80, 250, 280), pixmap=pix)
            doc.save(str(src), encryption=method, user_pw="user123", owner_pw="owner456")
            doc.close()

            res = optimize_pdf(src, password="user123", preset_key="standard")
            self.assertEqual(res.status, "success", f"Standard {name} harus berhasil dioptimasi!")
            self.assertIsNotNone(res.dst_path)

            out_doc = pymupdf.open(res.dst_path)
            self.assertTrue(out_doc.is_encrypted, f"Output {name} harus tetap terenkripsi!")
            self.assertEqual(out_doc.needs_pass, 1, f"Output {name} harus memerlukan password!")
            auth_res = out_doc.authenticate("user123")
            self.assertGreater(auth_res, 0, f"Password asli harus valid untuk output {name}!")
            self.assertEqual(len(out_doc), 1)
            out_doc.close()

    def test_marginal_savings_threshold_boundaries(self):
        """M2 Deliverable: Verifikasi ambang batas 3% & 500 KB di bawah skenario penghematan marjinal."""
        from unittest.mock import patch

        src = self._create_text_pdf("marginal_test.pdf", pages=2)
        orig_hash = get_file_hash(src)
        real_stat = Path.stat

        # Kasus 1: Hemat 100 KB pada 20 MB (< 500 KB DAN < 3.0%) -> Dibuang (already_optimal)
        def mock_stat_1(path_obj, *args, **kwargs):
            st = real_stat(path_obj, *args, **kwargs)
            m = unittest.mock.MagicMock(wraps=st)
            m.st_mode = st.st_mode
            if ".tmp_" in path_obj.name:
                m.st_size = 20 * 1024 * 1024 - 100 * 1024
                return m
            elif path_obj.name == src.name:
                m.st_size = 20 * 1024 * 1024
                return m
            return st

        with patch.object(Path, "stat", mock_stat_1):
            res1 = optimize_pdf(src, preset_key="light")
            self.assertEqual(res1.status, "already_optimal")
            self.assertIsNone(res1.dst_path)
            self.assertEqual(get_file_hash(src), orig_hash)
            self.assertEqual(len(list(self.temp_dir.glob("*.tmp_*"))), 0)

        # Kasus 2: Hemat 600 KB pada 30 MB (>= 500 KB meskipun < 3.0%) -> Disimpan (success)
        def mock_stat_2(path_obj, *args, **kwargs):
            st = real_stat(path_obj, *args, **kwargs)
            m = unittest.mock.MagicMock(wraps=st)
            m.st_mode = st.st_mode
            if ".tmp_" in path_obj.name:
                m.st_size = 30 * 1024 * 1024 - 600 * 1024
                return m
            elif path_obj.name == src.name:
                m.st_size = 30 * 1024 * 1024
                return m
            return st

        with patch.object(Path, "stat", mock_stat_2):
            res2 = optimize_pdf(src, preset_key="light")
            self.assertEqual(res2.status, "success")
            self.assertIsNotNone(res2.dst_path)
            self.assertEqual(res2.saved_bytes, 600 * 1024)
            self.assertEqual(len(list(self.temp_dir.glob("*.tmp_*"))), 0)

        # Kasus 3: Hemat 100 KB pada 1 MB (>= 3.0% meskipun < 500 KB) -> Disimpan (success)
        def mock_stat_3(path_obj, *args, **kwargs):
            st = real_stat(path_obj, *args, **kwargs)
            m = unittest.mock.MagicMock(wraps=st)
            m.st_mode = st.st_mode
            if ".tmp_" in path_obj.name:
                m.st_size = 1024 * 1024 - 100 * 1024
                return m
            elif path_obj.name == src.name:
                m.st_size = 1024 * 1024
                return m
            return st

        with patch.object(Path, "stat", mock_stat_3):
            res3 = optimize_pdf(src, preset_key="light")
            self.assertEqual(res3.status, "success")
            self.assertIsNotNone(res3.dst_path)
            self.assertEqual(res3.saved_bytes, 100 * 1024)
            self.assertEqual(len(list(self.temp_dir.glob("*.tmp_*"))), 0)

    def test_zero_tmp_leak_on_save_exception(self):
        """M2 Deliverable: Garansi 0 file .tmp bocor jika doc.save gagal/melempar error."""
        from unittest.mock import patch
        from pdf_optimizer import OptimizerError

        src = self._create_text_pdf("save_fail.pdf", pages=2)
        with patch.object(pymupdf.Document, "save", side_effect=OSError("Simulated I/O disk error")):
            with self.assertRaises(OptimizerError):
                optimize_pdf(src, preset_key="standard")

        # Pastikan tidak ada file sementara .tmp_* tertinggal di folder
        tmp_files = list(self.temp_dir.glob("*.tmp_*"))
        self.assertEqual(len(tmp_files), 0)

    def test_zero_tmp_leak_on_validation_failure(self):
        """M2 Deliverable: Garansi 0 file .tmp bocor jika validate_output gagal."""
        from unittest.mock import patch
        from pdf_optimizer import OptimizerError

        src = self._create_text_pdf("val_fail.pdf", pages=2)
        with patch("pdf_optimizer.validate_output", return_value=False):
            with self.assertRaises(OptimizerError):
                optimize_pdf(src, preset_key="standard")

        tmp_files = list(self.temp_dir.glob("*.tmp_*"))
        self.assertEqual(len(tmp_files), 0)

    def test_zero_tmp_leak_on_midway_cancellation(self):
        """M2 Deliverable: Garansi 0 file .tmp bocor saat pembatalan kooperatif di tengah proses."""
        src = self._create_image_heavy_pdf("mid_cancel.pdf")
        cancel_evt = threading.Event()

        def progress_cb(stage, pct):
            if pct >= 20:
                cancel_evt.set()

        with self.assertRaises(CancelledError):
            optimize_pdf(src, on_progress=progress_cb, cancel_event=cancel_evt)

        tmp_files = list(self.temp_dir.glob("*.tmp_*"))
        self.assertEqual(len(tmp_files), 0)


if __name__ == "__main__":
    unittest.main()

