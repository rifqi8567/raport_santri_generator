from flask import Flask, render_template, request, send_file, url_for, jsonify, Response
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import io
import traceback
from num2words import num2words
from datetime import datetime
import tempfile
import zipfile
import shutil
import glob
from datetime import datetime
import math

# 
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# 

app = Flask(__name__)
UPLOAD_FOLDER = "/tmp"
STATIC_FOLDER = "static"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)
logo_x = "logo.png"  # Placeholder for logo path
logo_y = "logo2.png"  # Placeholder for logo path
pdfmetrics.registerFont(TTFont("Calligrapher", os.path.join("static", "font.TTF")))
pdfmetrics.registerFont(TTFont("NotoNaskhArabic", os.path.join("static", "NotoNaskhArabic.ttf")))

# Global storage for uploaded data
biodata_storage = None
nilai_storage = {}
kompetensi_storage = {}
tahsin_tahfidz_storage = {}
karakter_storage = {}

# Mapping wali kelas
wali_kelas_map = {
    "7A": "Muhammad Naufal, S.Pd.",
    "7B": "Thalib Muammar, S.Ag.",
    "8": "Muhammad Radi, S.Pd.",
    "9": "Muhammad Akbar, S.Pd.",
    "10A": "Khalid Zaid Hamzah. P, S.H.",
    "10B": "Muhammad Rifqi Thufail",
    "11": "Kukuh Ahyar Pattani",
    "12": "Wafiq M. A. A. A.P",
}

nip_map = {
    "kepsek": "09025234",
    "7A": "09025248",
    "7B": "09025249",
    "8": "09025250",
    "9": "08023170",
    "10A": "09025251",
    "10B": "09025255",
    "11": "09025252",
    "12": "09025253",
}


def normalize_kelas(kelas):
    """Normalisasi nama kelas supaya konsisten (7A, 10B, 11, 12, dll)."""
    if not kelas:
        return ""
    kelas = str(kelas).strip().upper()
    if kelas.isdigit():
        return str(int(kelas))  # "07" -> "7", "11" tetap "11"
    return kelas


# School level mapping
def get_school_level(kelas):
    """Determine school level based on class"""
    if kelas in ["7A", "7B", "8", "9"]:
        return "SMP"
    elif kelas in ["10A", "10B", "11", "12"]:
        return "SMA"
    return "SMA"  # default


def number_to_words(n):
    """Convert number to Indonesian words"""
    try:
        return num2words(n, lang="id").replace("koma nol", "").title()
    except:
        return str(n)


def get_predikat(nilai):
    if nilai >= 90:
        return "A"
    elif nilai >= 80:
        return "B"
    elif nilai >= 70:
        return "C"
    else:
        return "D"


def draw_wrapped_text(
    c, text, x, y, max_width, line_height=11, font="Times-Roman", size=10
):
    """
    Gambar teks panjang dengan word-wrap supaya tidak keluar dari cell tabel.
    """
    c.setFont(font, size)
    words = text.split()
    line = ""
    current_y = y
    for word in words:
        test_line = (line + " " + word).strip()
        if c.stringWidth(test_line, font, size) <= max_width:
            line = test_line
        else:
            c.drawString(x, current_y, line)
            current_y -= line_height
            line = word
    if line:
        c.drawString(x, current_y, line)
    return current_y


def calc_text_height(c, text, max_width, font="Times-Roman", size=10, line_height=11):
    words = text.split()
    lines = []
    line = ""
    for word in words:
        test_line = (line + " " + word).strip()
        if c.stringWidth(test_line, font, size) <= max_width:
            line = test_line
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    return len(lines), lines


def draw_centered_text(c, text, x, y, w, h, font="Times-Roman", size=10, bold=False):
    """Tulis teks di tengah cell (horizontal & vertical center)"""
    if bold:
        c.setFont("Times-Bold", size)
    else:
        c.setFont(font, size)
    text_width = c.stringWidth(text, font if not bold else "Times-Bold", size)
    c.drawString(x + (w - text_width) / 2, y + (h - size) / 2 + 3, text)


# def draw_watermark_logo(c, width, height):
#     """Draw watermark logo in center with low opacity"""
#     bg_logo_path = os.path.join(STATIC_FOLDER, "logo_bg.png")
#     if os.path.exists(bg_logo_path):
#         try:
#             # Save current graphics state
#             c.saveState()
#             # Set low opacity for watermark
#             c.setFillAlpha(0.12)
#             bg_logo = ImageReader(bg_logo_path)
#             bg_size = 280
#             c.drawImage(
#                 bg_logo,
#                 width / 2 - bg_size / 2,
#                 height / 2 - bg_size / 2,
#                 width=bg_size,
#                 height=bg_size,
#                 mask="auto",
#             )
#             # Restore graphics state
#             c.restoreState()
#         except:
#             pass


def draw_signatures(c, width, height, left_margin, row, wali_kelas_map, nip_map):
    """
    Signature section: Kepala Sekolah (kiri) & Wali Kelas (kanan) + TTD otomatis
    """
    signature_y = 130
    # Ambil kelas dari row biodata
    kelas = normalize_kelas(row.get("Kelas", ""))

    # Cari wali kelas & nip
    wali_kelas = wali_kelas_map.get(kelas, "MUHAMAD AKBAR, S.Pd")
    nip_wali = nip_map.get(kelas, "-")

    # Posisi kolom
    col_width = (width - 2 * left_margin) / 2
    col_x = [
        left_margin + col_width / 2,  # Kepala Sekolah (kiri)
        left_margin + col_width * 1.5,  # Wali Kelas (kanan)
    ]

    # === Kepala Sekolah ===
    c.setFont("Times-Roman", 11)
    tanggal_hari_ini = datetime.now().strftime("%d %B %Y")
    c.drawCentredString(col_x[0], signature_y + 20, f"Bekasi, {tanggal_hari_ini}")
    c.drawCentredString(col_x[0], signature_y, "Kepala Sekolah")

    # TTD Kepala Sekolah
    ttd_kepsek_path = os.path.join("static", "ttd_kepsek.png")
    if os.path.exists(ttd_kepsek_path):
        try:
            ttd_width, ttd_height = 120, 60
            c.drawImage(
                ttd_kepsek_path,
                col_x[0] - (ttd_width / 2),
                signature_y - 50,
                width=ttd_width,
                height=ttd_height,
                mask="auto",
                preserveAspectRatio=True,
            )
        except Exception as e:
            print(f"Error loading kepsek signature: {e}")

    # Garis + nama + NIP Kepala Sekolah
    c.line(col_x[0] - 70, signature_y - 45, col_x[0] + 70, signature_y - 45)
    c.setFont("Times-Bold", 11)
    c.drawCentredString(col_x[0], signature_y - 60, "Wildan Arif Priambodo, Lc.")
    c.setFont("Times-Roman", 10)
    c.drawCentredString(
        col_x[0], signature_y - 75, f"NIP. {nip_map.get('kepsek', '-')}"
    )

    # === Wali Kelas ===
    c.setFont("Times-Roman", 11)
    c.drawCentredString(col_x[1], signature_y + 20, "Wali Kelas")

    # TTD Wali Kelas
    ttd_wali_path = os.path.join("static", f"ttd_wali_{kelas}.png")
    if os.path.exists(ttd_wali_path):
        try:
            ttd_width, ttd_height = 120, 60
            c.drawImage(
                ttd_wali_path,
                col_x[1] - (ttd_width / 2),
                signature_y - 45,
                width=ttd_width,
                height=ttd_height,
                mask="auto",
                preserveAspectRatio=True,
            )
        except Exception as e:
            print(f"Error loading wali signature: {e}")

    # Garis + nama + NIP Wali Kelas
    c.line(col_x[1] - 70, signature_y - 45, col_x[1] + 70, signature_y - 45)
    c.setFont("Times-Bold", 11)
    c.drawCentredString(col_x[1], signature_y - 60, str(wali_kelas))
    c.setFont("Times-Roman", 10)
    c.drawCentredString(col_x[1], signature_y - 75, f"NIP. {nip_wali}")


def create_cover_page(c, width, height, nis, nisn, nama, logo_path, school_level):
    """Page 1: Cover"""
    try:
        center_x = width / 2
        top_y = height - 80

        # Draw watermark logo
        # draw_watermark_logo(c, width, height)

        # Header title - Enhanced formatting with formal fonts
        c.setFont("Times-Bold", 18)
        c.setFillColor(colors.black)
        c.drawCentredString(
            center_x, top_y, "RAPORT PENILAIAN SUMATIF AKHIR SEMESTER I (PSAS I) T.A. 2025/2026"
        )

        # Bagian sekolah, pisah font
        text_y = top_y - 24

        # Hitung posisi teks manual
        prefix = f"{school_level} PESANTREN "
        middle = "RABBAANII "
        suffix = "ISLAMIC SCHOOL"

        # Hitung lebar tiap bagian untuk menentukan posisi X
        prefix_width = c.stringWidth(prefix, "Times-Bold", 16)
        middle_width = c.stringWidth(
            middle, "Calligrapher", 20
        )  # ukuran sedikit lebih besar
        suffix_width = c.stringWidth(suffix, "Times-Bold", 16)

        total_width = prefix_width + middle_width + suffix_width
        start_x = center_x - (total_width / 2)

        # Gambar teks dengan font berbeda
        c.setFont("Times-Bold", 16)
        c.drawString(start_x, text_y, prefix)

        c.setFont("Calligrapher", 20)  # font khusus untuk "RABBAANII"
        c.drawString(start_x + prefix_width, text_y, middle)

        c.setFont("Times-Bold", 16)
        c.drawString(start_x + prefix_width + middle_width, text_y, suffix)

        # Baris berikutnya tetap normal
        c.setFont("Times-Bold", 15)
        c.drawCentredString(center_x, top_y - 46, "KECAMATAN CIKARANG TIMUR")

        # School logo - positioned higher for better balance
        if os.path.exists(logo_path):
            try:
                logo = ImageReader(logo_path)
                logo_size = 130
                logo_y = (height / 2) - (logo_size / 2) + 20
                c.drawImage(
                    logo,
                    center_x - logo_size / 2,
                    logo_y,
                    width=logo_size,
                    height=logo_size,
                    mask="auto",
                )
            except:
                pass

        # NIS / NISN box - Enhanced styling
        box_width = 380
        box_height = 40
        nisn_y = height / 2 - 160

        c.setStrokeColor(colors.black)
        c.setLineWidth(1.5)
        c.setFillColor(colors.lightgrey)
        c.rect(
            center_x - box_width / 2, nisn_y, box_width, box_height, fill=1, stroke=1
        )

        c.setFont("Times-Bold", 15)
        c.setFillColor(colors.black)
        nis_text = f"{nis} / {nisn}" if nis and nisn else f"{nis or nisn or ''}"
        c.drawCentredString(center_x, nisn_y + 15, nis_text)

        # Name box - Enhanced styling
        name_y = nisn_y - 70
        c.setFillColor(colors.white)
        c.rect(
            center_x - box_width / 2, name_y, box_width, box_height, fill=1, stroke=1
        )

        c.setFont("Times-Bold", 18)
        c.setFillColor(colors.black)
        c.drawCentredString(center_x, name_y + 15, str(nama).upper())

    except Exception as e:
        print(f"Error creating cover page: {e}")


def draw_wrapped_field(
    c,
    label,
    value,
    x,
    y,
    colon_x,
    max_width,
    line_height=14,
    font="Times-Roman",
    size=12,
):
    """Menggambar field dengan label + value yang panjang (wrap sejajar setelah titik dua)"""
    c.setFont(font, size)
    c.drawString(x, y, label)

    if value:
        c.drawString(colon_x, y, ":")
        text_x = colon_x + 15

        words = str(value).split()
        line = ""
        current_y = y
        for word in words:
            test_line = (line + " " + word).strip()
            if c.stringWidth(test_line, font, size) <= max_width:
                line = test_line
            else:
                c.drawString(text_x, current_y, line)
                current_y -= line_height
                line = word
        if line:
            c.drawString(text_x, current_y, line)

        return current_y - line_height  # kembalikan posisi Y setelah teks
    else:
        return y - line_height


def wrap_text_lines(c, text, max_width, font="Times-Roman", size=10):
    """Return list of wrapped lines for given text sesuai lebar max_width."""
    words = text.split()
    lines, line = [], ""
    for word in words:
        test_line = (line + " " + word).strip()
        if c.stringWidth(test_line, font, size) <= max_width:
            line = test_line
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


def format_score(v):
    try:
        if v is None:
            return ""
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            return ""
        f = float(s)
        # Jika integer (88.0) -> "88"
        if f.is_integer():
            return str(int(f))
        # Jika desimal, rapikan (88.50 -> 88.5, 88.00 -> 88)
        s = f"{f:.2f}".rstrip("0").rstrip(".")
        return s
    except Exception:
        # fallback: kembalikan apa adanya
        return str(v) if v is not None else ""

# page 2 data diri
def create_data_page(c, width, height, row, school_level):
    """Page 2: Personal Data"""
    try:
        # draw_watermark_logo(c, width, height)

        # Header
        c.setFillColor(colors.black)

        prefix = f"DATA DIRI SISWA {school_level} "
        middle = "RABBAANII "
        suffix = "ISLAMIC SCHOOL"

        # Hitung lebar total agar tetap rata tengah
        prefix_width = c.stringWidth(prefix, "Times-Bold", 18)
        middle_width = c.stringWidth(
            middle, "Calligrapher", 22
        )  # dibuat sedikit lebih besar
        suffix_width = c.stringWidth(suffix, "Times-Bold", 18)

        total_width = prefix_width + middle_width + suffix_width
        start_x = (width - total_width) / 2
        y_pos = height - 50

        # Gambar teks
        c.setFont("Times-Bold", 18)
        c.drawString(start_x, y_pos, prefix)

        c.setFont("Calligrapher", 22)
        c.drawString(start_x + prefix_width, y_pos, middle)

        c.setFont("Times-Bold", 18)
        c.drawString(start_x + prefix_width + middle_width, y_pos, suffix)

        y_start = height - 90
        left_margin = 80
        line_height = 24
        current_y = y_start

        c.setFont("Times-Roman", 12)

        def safe_get(key, default=""):
            value = row.get(key, default)
            return str(value) if pd.notna(value) and value != "" else default

        # form_fields dibuat tuple (label, value, key_name)
        form_fields = [
            ("1.  Nama Siswa", safe_get("Nama"), "Nama"),
            (
                "2.  Nomor Induk Siswa/NISN",
                (
                    f"{safe_get('NIS', '')} / {safe_get('NISN', '')}".replace(
                        " / ", "/"
                    )
                    if safe_get("NIS") or safe_get("NISN")
                    else ""
                ),
                "NISN",
            ),
            (
                "3.  Tempat, Tanggal Lahir",
                safe_get("TempatTanggalLahir"),
                "TempatTanggalLahir",
            ),
            ("4.  Jenis Kelamin", safe_get("JenisKelamin"), "JenisKelamin"),
            ("5.  Agama", safe_get("Agama"), "Agama"),
            (
                "6.  Pendidikan Sebelumnya",
                safe_get("PendidikanSebelumnya"),
                "PendidikanSebelumnya",
            ),
            ("7.  Alamat Siswa", safe_get("Alamat"), "Alamat"),
            ("", "", ""),
            ("8.  Nama Orang Tua", "", ""),
            ("     a. Ayah", safe_get("NamaAyah"), "NamaAyah"),
            ("     b. Ibu", safe_get("NamaIbu"), "NamaIbu"),
            ("", "", ""),
            ("9.  Pekerjaan Orang Tua", "", ""),
            ("     a. Ayah", safe_get("PekerjaanAyah"), "PekerjaanAyah"),
            ("     b. Ibu", safe_get("PekerjaanIbu"), "PekerjaanIbu"),
            ("", "", ""),
            ("10. Alamat Orang Tua", "", ""),
            ("     a. Ayah", safe_get("AlamatAyah"), "AlamatAyah"),
            ("     b. Ibu", safe_get("AlamatIbu"), "AlamatIbu"),
            ("", "", ""),
            ("11. Wali Siswa", "", ""),
            ("     a. Nama", safe_get("NamaWali"), "NamaWali"),
            ("     b. Pekerjaan", safe_get("PekerjaanWali"), "PekerjaanWali"),
            ("     c. Alamat", safe_get("AlamatWali"), "AlamatWali"),
        ]

        alamat_keys = ["Alamat", "AlamatAyah", "AlamatIbu", "AlamatWali"]

        for label, value, key_name in form_fields:
            if label == "":
                current_y -= line_height * 0.6
                continue

            colon_x = left_margin + 200
            max_width = width - colon_x - 60

            # cek: kalau label ada kata 'Alamat' atau key_name ada di daftar alamat
            if "Alamat" in label or key_name in alamat_keys:
                current_y = draw_wrapped_field(
                    c,
                    label,
                    value,
                    left_margin,
                    current_y,
                    colon_x,
                    max_width,
                    line_height=14,
                )
            else:
                # field normal
                c.setFont("Times-Roman", 12)
                c.drawString(left_margin, current_y, label)
                if value:
                    c.drawString(colon_x, current_y, ":")
                    c.drawString(colon_x + 15, current_y, str(value))
                current_y -= line_height

        signature_y = 140
        signature_x = width - 240

        c.setFont("Times-Roman", 11)
        # Tanggal statis
        c.drawString(signature_x, signature_y + 40, "Bekasi, 21 Juli 2025")

        # Baris 1: Kepala SMA (atau sesuai school_level)
        jabatan_text = f"Kepala {school_level}"
        c.setFont("Times-Roman", 11)
        c.drawString(signature_x, signature_y + 20, jabatan_text)

        # Baris 2: Pesantren RABBAANII Islamic School
        prefix = "Pesantren "
        middle = "RABBAANII"
        suffix = " Islamic School"

        prefix_width = c.stringWidth(prefix, "Times-Roman", 11)
        middle_width = c.stringWidth(middle, "Calligrapher", 13)

        c.setFont("Times-Roman", 11)
        c.drawString(signature_x, signature_y + 5, prefix)

        c.setFont("Calligrapher", 13)
        c.drawString(signature_x + prefix_width, signature_y + 5, middle)

        c.setFont("Times-Roman", 11)
        c.drawString(signature_x + prefix_width + middle_width, signature_y + 5, suffix)

        # === Spasi untuk tanda tangan ===
        ttd_path = os.path.join(STATIC_FOLDER, "ttd_kepsek.png")
        if os.path.exists(ttd_path):
            try:
                ttd_width = 120
                ttd_height = 60
                c.drawImage(
                    ttd_path,
                    signature_x + 30,
                    signature_y - 60,  # posisinya di atas garis
                    width=ttd_width,
                    height=ttd_height,
                    mask="auto",
                    preserveAspectRatio=True,
                )
            except Exception as e:
                print(f"Error loading signature image: {e}")

        # Garis tanda tangan
        c.line(signature_x, signature_y - 50, signature_x + 180, signature_y - 50)

        # Nama Kepala Sekolah
        c.setFont("Times-Bold", 12)
        c.drawString(signature_x + 30, signature_y - 65, "Wildan Arif Priambodo, Lc.")

        # NIP
        c.setFont("Times-Roman", 10)
        c.drawString(signature_x + 30, signature_y - 80, "NIP. 09025234")

    except Exception as e:
        print(f"Error creating data page: {e}")

# page 3 nilai
def create_score_page(c, width, height, row, nilai_df, wali_kelas_map, school_level):
    """Page 3: Scores with header logos and transparent background logo"""
    try:
        center_x = width / 2
        left_margin = 50

        # === Background watermark logo ===
        # bg_logo_path = os.path.join(STATIC_FOLDER, "logo_bg.png")
        # if os.path.exists(bg_logo_path):
        #     bg_logo = ImageReader(bg_logo_path)
        #     bg_size = 300
        #     c.drawImage(
        #         bg_logo,
        #         center_x - bg_size / 2,
        #         height / 2 - bg_size / 2,
        #         width=bg_size,
        #         height=bg_size,
        #         mask="auto",
        #     )

        # === Header logos kiri & kanan ===
        logo_left_path = os.path.join(STATIC_FOLDER, "logo.png")
        logo_right_path = os.path.join(STATIC_FOLDER, "logo2.png")

        logo_size = 70
        logo_y = height - 90

        if os.path.exists(logo_left_path):
            logo_left = ImageReader(logo_left_path)
            c.drawImage(
                logo_left, 40, logo_y, width=logo_size, height=logo_size, mask="auto"
            )

        if os.path.exists(logo_right_path):
            logo_right = ImageReader(logo_right_path)
            c.drawImage(
                logo_right,
                width - 40 - logo_size,
                logo_y,
                width=logo_size,
                height=logo_size,
                mask="auto",
            )

        # === Header text ===
        header_y = height - 40
        c.setFont("Times-Bold", 13)
        c.setFillColor(colors.black)
        c.drawCentredString(
            center_x, header_y, "LAPORAN PENILAIAN SUMATIF AKHIR SEMESTER I (PSAS I)"
        )

        prefix = f"{school_level} PESANTREN "
        middle = "RABBAANII"
        prefix_width = c.stringWidth(prefix, "Times-Bold", 13)
        middle_width = c.stringWidth(middle, "Calligrapher", 15)
        total_width = prefix_width + middle_width
        start_x = center_x - (total_width / 2)
        y_pos = header_y - 18

        c.setFont("Times-Bold", 13)
        c.drawString(start_x, y_pos, prefix)
        c.setFont("Calligrapher", 15)
        c.drawString(start_x + prefix_width, y_pos, middle)

        c.setFont("Times-Bold", 12)
        c.drawCentredString(center_x, header_y - 36, "TAHUN PELAJARAN 2025/2026")
        c.drawCentredString(center_x, header_y - 54, "KECAMATAN CIKARANG TIMUR")

        # === Student information ===
        info_y = header_y - 100
        c.setFont("Times-Roman", 11)
        c.drawString(70, info_y, "Nama")
        c.drawString(120, info_y, f": {row.get('Nama', '')}")
        c.drawString(70, info_y - 15, "NIS/NISN")
        c.drawString(120, info_y - 15, f": {row.get('NIS', '')}/{row.get('NISN', '')}")
        c.drawString(350, info_y, "Kelas")
        c.drawString(400, info_y, f": {row.get('Kelas', '')}")
        c.drawString(350, info_y - 15, "Semester")
        c.drawString(400, info_y - 15, ": I (SATU)")

        # === Table setup ===
        table_start_y = info_y - 50
        table_x = 70
        table_width = width - 140

        col_no_width = 30
        col_mapel_width = 200
        col_kkm_width = 40
        col_nilai_width = 80
        col_huruf_width = (
            table_width
            - col_no_width
            - col_mapel_width
            - col_kkm_width
            - col_nilai_width
        )

        # === Header ===
        header_height = 40
        c.setStrokeColor(colors.black)
        c.setLineWidth(1)
        c.setFillColor(colors.lightgrey)
        c.rect(
            table_x, table_start_y - header_height, table_width, header_height, fill=1
        )

        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 13)
        c.rect(table_x, table_start_y - header_height, col_no_width, header_height)
        draw_centered_text(
            c,
            "NO",
            table_x,
            table_start_y - header_height,
            col_no_width,
            header_height,
            bold=True,
        )

        x_pos = table_x + col_no_width
        c.rect(x_pos, table_start_y - header_height, col_mapel_width, header_height)
        draw_centered_text(
            c,
            "MATA PELAJARAN",
            x_pos,
            table_start_y - header_height,
            col_mapel_width,
            header_height,
            bold=True,
        )

        x_pos += col_mapel_width
        c.rect(x_pos, table_start_y - header_height, col_kkm_width, header_height)
        draw_centered_text(
            c,
            "KKM",
            x_pos,
            table_start_y - header_height,
            col_kkm_width,
            header_height,
            bold=True,
        )

        x_pos += col_kkm_width
        c.rect(
            x_pos,
            table_start_y - header_height,
            col_nilai_width + col_huruf_width,
            header_height,
        )
        draw_centered_text(
            c,
            "NILAI",
            x_pos,
            table_start_y - 20,
            col_nilai_width + col_huruf_width,
            20,
            bold=True,
        )

        c.rect(x_pos, table_start_y - header_height, col_nilai_width, 20)
        draw_centered_text(
            c,
            "ANGKA",
            x_pos,
            table_start_y - header_height,
            col_nilai_width,
            20,
            bold=True,
        )
        c.rect(
            x_pos + col_nilai_width, table_start_y - header_height, col_huruf_width, 20
        )
        draw_centered_text(
            c,
            "HURUF",
            x_pos + col_nilai_width,
            table_start_y - header_height,
            col_huruf_width,
            20,
            bold=True,
        )

        # === Table content ===
        current_y = table_start_y - header_height
        total, count = 0, 0
        padding = 5
        base_row_height = 28
        print("JUMLAH BARIS nilai_df:", len(nilai_df))
        print("KOLOM:", list(nilai_df.columns))
        print(nilai_df.head())
        if not nilai_df.empty:
            for i, (_, nilai_row) in enumerate(nilai_df.iterrows(), start=1):
                mapel = str(nilai_row.get("MataPelajaran", ""))
                arab = str(nilai_row.get("Arab", ""))
                kkm = str(nilai_row.get("KKM", ""))
                nilai = nilai_row.get("Nilai", 0)
                try:
                    nilai_int = int(float(nilai))
                except (ValueError, TypeError):
                    nilai_int = 0
                huruf = number_to_words(nilai_int)

                # Process Arabic text
                reshaped_arab = arabic_reshaper.reshape(arab)
                bidi_arab = get_display(reshaped_arab)

                # Calculate widths
                c.setFont("NotoNaskhArabic", 12)
                arab_width = c.stringWidth(bidi_arab, "NotoNaskhArabic", 12)
                
                # Max width for Indo text (total - arab - padding)
                mapel_max_width = col_mapel_width - arab_width - 15

                # Calculate lines for wrapping
                lines_mapel = wrap_text_lines(c, mapel, mapel_max_width, font="Times-Roman", size=11)
                lines_huruf = wrap_text_lines(c, huruf, col_huruf_width - 10, font="Times-Roman", size=12)

                # Determine row height
                height_mapel = len(lines_mapel) * 14
                height_huruf = len(lines_huruf) * 14
                row_height = max(base_row_height, height_mapel + 10, height_huruf + 10)

                current_y -= row_height
                
                # Draw Rectangles
                c.rect(table_x, current_y, col_no_width, row_height)
                c.rect(table_x + col_no_width, current_y, col_mapel_width, row_height)
                c.rect(
                    table_x + col_no_width + col_mapel_width,
                    current_y,
                    col_kkm_width,
                    row_height,
                )
                c.rect(
                    table_x + col_no_width + col_mapel_width + col_kkm_width,
                    current_y,
                    col_nilai_width,
                    row_height,
                )
                c.rect(
                    table_x
                    + col_no_width
                    + col_mapel_width
                    + col_kkm_width
                    + col_nilai_width,
                    current_y,
                    col_huruf_width,
                    row_height,
                )

                # Draw NO
                draw_centered_text(c, str(i), table_x, current_y, col_no_width, row_height)

                # Draw MATA PELAJARAN (Indo kiri – Arab kanan)
                text_top_y = current_y + row_height - 15
                
                # --- Garis pemisah vertikal antara Indo & Arab ---
                separator_x = table_x + col_no_width + (col_mapel_width / 2)
                c.setLineWidth(0.5)
                c.line(separator_x, current_y, separator_x, current_y + row_height)
                
                # Indo (Wrapped)
                draw_wrapped_text(c, mapel, table_x + col_no_width + 5, text_top_y, mapel_max_width, line_height=14, font="Times-Roman", size=11)

                # Arab (Right aligned)
                c.setFont("NotoNaskhArabic", 12)
                c.drawRightString(table_x + col_no_width + col_mapel_width - 5, text_top_y, bidi_arab)

                # Draw KKM & Nilai
                draw_centered_text(
                    c,
                    kkm,
                    table_x + col_no_width + col_mapel_width,
                    current_y,
                    col_kkm_width,
                    row_height,
                )
                draw_centered_text(
                    c,
                    str(nilai_int),
                    table_x + col_no_width + col_mapel_width + col_kkm_width,
                    current_y,
                    col_nilai_width,
                    row_height,
                )

                # Draw Huruf (Wrapped)
                draw_wrapped_text(
                    c,
                    huruf,
                    table_x + col_no_width + col_mapel_width + col_kkm_width + col_nilai_width + 5,
                    text_top_y,
                    col_huruf_width - 10,
                    line_height=14,
                    font="Times-Roman",
                    size=12
                )

                total += nilai_int
                count += 1

        # === Jumlah Nilai ===
        huruf_total = number_to_words(total)
        lines_total = wrap_text_lines(
            c, huruf_total, col_huruf_width - 10, font="Times-Bold", size=12
        )
        row_height_total = max(base_row_height, 12 * len(lines_total) + 10)
        current_y -= row_height_total

        c.setFont("Times-Bold", 12)
        c.rect(table_x, current_y, col_no_width + col_mapel_width, row_height_total)
        c.rect(
            table_x + col_no_width + col_mapel_width,
            current_y,
            col_kkm_width,
            row_height_total,
        )
        c.rect(
            table_x + col_no_width + col_mapel_width + col_kkm_width,
            current_y,
            col_nilai_width,
            row_height_total,
        )
        c.rect(
            table_x + col_no_width + col_mapel_width + col_kkm_width + col_nilai_width,
            current_y,
            col_huruf_width,
            row_height_total,
        )

        draw_centered_text(
            c,
            "Jumlah Nilai",
            table_x,
            current_y,
            col_no_width + col_mapel_width,
            row_height_total,
            bold=True,
        )

        draw_centered_text(
            c,
            str(total),
            table_x + col_no_width + col_mapel_width + col_kkm_width,
            current_y,
            col_nilai_width,
            row_height_total,
            bold=True,
        )

        text_y = current_y + row_height_total - 15
        for line in lines_total:
            c.drawString(
                table_x
                + col_no_width
                + col_mapel_width
                + col_kkm_width
                + col_nilai_width
                + padding,
                text_y,
                line,
            )
            text_y -= 12

        # === Nilai Rata-rata ===
        rata2 = total // count if count else 0
        huruf_rata2 = number_to_words(rata2)
        lines_rata = wrap_text_lines(
            c, huruf_rata2, col_huruf_width - 10, font="Times-Bold", size=12
        )
        row_height_rata = max(base_row_height, 12 * len(lines_rata) + 10)
        current_y -= row_height_rata

        c.rect(table_x, current_y, col_no_width + col_mapel_width, row_height_rata)
        c.rect(
            table_x + col_no_width + col_mapel_width,
            current_y,
            col_kkm_width,
            row_height_rata,
        )
        c.rect(
            table_x + col_no_width + col_mapel_width + col_kkm_width,
            current_y,
            col_nilai_width,
            row_height_rata,
        )
        c.rect(
            table_x + col_no_width + col_mapel_width + col_kkm_width + col_nilai_width,
            current_y,
            col_huruf_width,
            row_height_rata,
        )

        draw_centered_text(
            c,
            "Nilai Rata-rata",
            table_x,
            current_y,
            col_no_width + col_mapel_width,
            row_height_rata,
            bold=True,
        )

        draw_centered_text(
            c,
            str(rata2),
            table_x + col_no_width + col_mapel_width + col_kkm_width,
            current_y,
            col_nilai_width,
            row_height_rata,
            bold=True,
        )

        text_y = current_y + row_height_rata - 15
        for line in lines_rata:
            c.drawString(
                table_x
                + col_no_width
                + col_mapel_width
                + col_kkm_width
                + col_nilai_width
                + padding,
                text_y,
                line,
            )
            text_y -= 12

        # === Signatures ===
        draw_signatures(c, width, height, left_margin, row, wali_kelas_map, nip_map)

    except Exception as e:
        print(f"Error creating score page: {e}")


def safe_int(value, default=0):
    """Convert nilai ke int dengan aman (support format '92,5')."""
    try:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.strip().replace(",", ".")
        return int(float(value))
    except Exception:
        return default


def create_kompetensi_page(
    c, width, height, row, kompetensi_df, school_level, wali_kelas_map
):
    """Page: Kompetensi Keahlian (untuk SMA)"""
    try:
        center_x = width / 2
        left_margin = 50

        # === Background watermark logo ===
        # bg_logo_path = os.path.join(STATIC_FOLDER, "logo_bg.png")
        # if os.path.exists(bg_logo_path):
        #     bg_logo = ImageReader(bg_logo_path)
        #     bg_size = 300
        #     c.drawImage(
        #         bg_logo,
        #         center_x - bg_size / 2,
        #         height / 2 - bg_size / 2,
        #         width=bg_size,
        #         height=bg_size,
        #         mask="auto",
        #     )

        # === Header logos kiri & kanan ===
        logo_left_path = os.path.join(STATIC_FOLDER, "logo.png")
        logo_right_path = os.path.join(STATIC_FOLDER, "logo2.png")

        logo_size = 70
        logo_y = height - 90

        if os.path.exists(logo_left_path):
            logo_left = ImageReader(logo_left_path)
            c.drawImage(
                logo_left, 40, logo_y, width=logo_size, height=logo_size, mask="auto"
            )

        if os.path.exists(logo_right_path):
            logo_right = ImageReader(logo_right_path)
            c.drawImage(
                logo_right,
                width - 40 - logo_size,
                logo_y,
                width=logo_size,
                height=logo_size,
                mask="auto",
            )

        # === Header text ===
        header_y = height - 40
        c.setFont("Times-Bold", 13)
        c.setFillColor(colors.black)
        c.drawCentredString(
            center_x, header_y, "LAPORAN PENILAIAN SUMATIF AKHIR SEMESTER I (PSAS I)"
        )

        prefix = f"{school_level} PESANTREN "
        middle = "RABBAANII"

        prefix_width = c.stringWidth(prefix, "Times-Bold", 13)
        middle_width = c.stringWidth(middle, "Calligrapher", 15)
        start_x = center_x - (prefix_width + middle_width) / 2
        y_pos = header_y - 18

        c.setFont("Times-Bold", 13)
        c.drawString(start_x, y_pos, prefix)
        c.setFont("Calligrapher", 15)
        c.drawString(start_x + prefix_width, y_pos, middle)

        c.setFont("Times-Bold", 13)
        c.drawCentredString(center_x, header_y - 36, "TAHUN PELAJARAN 2025/2026")
        c.drawCentredString(center_x, header_y - 54, "KECAMATAN CIKARANG TIMUR")

        # === Student information ===
        info_y = header_y - 100
        c.setFont("Times-Roman", 11)
        c.drawString(70, info_y, f"Nama")
        c.drawString(120, info_y, f": {row.get('Nama', '')}")
        c.drawString(70, info_y - 15, f"NIS/NISN")
        c.drawString(120, info_y - 15, f": {row.get('NIS', '')}/{row.get('NISN', '')}")
        c.drawString(350, info_y, f"Kelas")
        c.drawString(400, info_y, f": {row.get('Kelas', '')}")
        c.drawString(350, info_y - 15, f"Semester")
        c.drawString(400, info_y - 15, ": I (SATU)")

        # === Setup tabel ===
        table_start_y = info_y - 70
        table_x = left_margin
        col_no_width = 45
        col_mapel_width = 250
        col_pengetahuan_width = 80
        col_keterampilan_width = 80
        col_akhir_width = 70
        col_predikat_width = 75
        available_width = width - (2 * left_margin)

        # scale jika tidak muat
        table_width = (
            col_no_width
            + col_mapel_width
            + col_pengetahuan_width
            + col_keterampilan_width
            + col_akhir_width
            + col_predikat_width
        )
        if table_width > available_width:
            scale = available_width / table_width
            col_no_width = round(col_no_width * scale)
            col_mapel_width = round(col_mapel_width * scale)
            col_pengetahuan_width = round(col_pengetahuan_width * scale)
            col_keterampilan_width = round(col_keterampilan_width * scale)
            col_akhir_width = round(col_akhir_width * scale)
            col_predikat_width = available_width - (
                col_no_width
                + col_mapel_width
                + col_pengetahuan_width
                + col_keterampilan_width
                + col_akhir_width
            )

        # Header
        header_height = 50
        c.setStrokeColor(colors.black)
        c.setFillColor(colors.lightgrey)
        c.rect(
            table_x,
            table_start_y - header_height,
            available_width,
            header_height,
            fill=1,
        )
        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 13)

        # NO
        c.rect(
            table_x, table_start_y - header_height, col_no_width, header_height, fill=0
        )
        draw_centered_text(
            c,
            "NO",
            table_x,
            table_start_y - header_height,
            col_no_width,
            header_height,
            bold=True,
        )

        # Mata Pelajaran
        x_pos = table_x + col_no_width
        c.rect(
            x_pos, table_start_y - header_height, col_mapel_width, header_height, fill=0
        )
        draw_centered_text(
            c,
            "Mata Pelajaran",
            x_pos,
            table_start_y - header_height,
            col_mapel_width,
            header_height,
            bold=True,
        )

        # Nilai Akademik
        x_pos += col_mapel_width
        nilai_akademik_width = (
            col_pengetahuan_width + col_keterampilan_width + col_akhir_width
        )
        c.rect(x_pos, table_start_y - 25, nilai_akademik_width, 25, fill=0)
        draw_centered_text(
            c,
            "NILAI AKADEMIK",
            x_pos,
            table_start_y - 25,
            nilai_akademik_width,
            25,
            bold=True,
        )

        # Sub header
        c.rect(x_pos, table_start_y - header_height, col_pengetahuan_width, 25, fill=0)
        draw_centered_text(
            c,
            "Pengetahuan",
            x_pos,
            table_start_y - header_height,
            col_pengetahuan_width,
            25,
            bold=True,
        )
        c.rect(
            x_pos + col_pengetahuan_width,
            table_start_y - header_height,
            col_keterampilan_width,
            25,
            fill=0,
        )
        draw_centered_text(
            c,
            "Keterampilan",
            x_pos + col_pengetahuan_width,
            table_start_y - header_height,
            col_keterampilan_width,
            25,
            bold=True,
        )
        c.rect(
            x_pos + col_pengetahuan_width + col_keterampilan_width,
            table_start_y - header_height,
            col_akhir_width,
            25,
            fill=0,
        )
        draw_centered_text(
            c,
            "Nilai Akhir",
            x_pos + col_pengetahuan_width + col_keterampilan_width,
            table_start_y - header_height,
            col_akhir_width,
            25,
            bold=True,
        )

        # Predikat
        x_pos += nilai_akademik_width
        c.rect(
            x_pos,
            table_start_y - header_height,
            col_predikat_width,
            header_height,
            fill=0,
        )
        draw_centered_text(
            c,
            "Predikat",
            x_pos,
            table_start_y - header_height,
            col_predikat_width,
            header_height,
            bold=True,
        )

        # === Isi Tabel ===
        current_y = table_start_y - header_height
        row_height = 25
        c.setFont("Times-Roman", 12)

        data_to_use = (
            kompetensi_df.to_dict("records")
            if (kompetensi_df is not None and not kompetensi_df.empty)
            else []
        )

        total_pengetahuan = total_keterampilan = total_akhir = 0
        count = 0

        for i, data in enumerate(data_to_use, start=1):
            current_y -= row_height
            pengetahuan = safe_int(data.get("Pengetahuan", 0))
            keterampilan = safe_int(data.get("Keterampilan", 0))
            nilai_akhir = safe_int(data.get("NilaiAkhir", 0))
            mata_pelajaran = str(data.get("MataPelajaran", ""))
            predikat = get_predikat(nilai_akhir)

            draw_table_row(
                c,
                table_x,
                current_y,
                [
                    (col_no_width, str(i)),
                    (col_mapel_width, mata_pelajaran),
                    (col_pengetahuan_width, str(pengetahuan)),
                    (col_keterampilan_width, str(keterampilan)),
                    (col_akhir_width, str(nilai_akhir)),
                    (col_predikat_width, predikat),
                ],
                row_height,
                align_second_left=True,
            )

            total_pengetahuan += pengetahuan
            total_keterampilan += keterampilan
            total_akhir += nilai_akhir
            count += 1

        # Rata-rata
        if count > 0:
            avg_pengetahuan = round(total_pengetahuan / count)
            avg_keterampilan = round(total_keterampilan / count)
            avg_akhir = round(total_akhir / count)
        else:
            avg_pengetahuan = avg_keterampilan = avg_akhir = 0

        avg_predikat = get_predikat(avg_akhir)
        current_y -= row_height
        draw_table_row(
            c,
            table_x,
            current_y,
            [
                (col_no_width + col_mapel_width, "Nilai Rata-rata"),
                (col_pengetahuan_width, str(avg_pengetahuan)),
                (col_keterampilan_width, str(avg_keterampilan)),
                (col_akhir_width, str(avg_akhir)),
                (col_predikat_width, avg_predikat),
            ],
            row_height,
            align_first_left=True,
            bold=True,
        )

        draw_signatures(c, width, height, left_margin, row, wali_kelas_map, nip_map)

    except Exception as e:
        print(f"Error creating kompetensi page: {e}")
        raise


def normalize_nis(nis_value):
    """Normalisasi NIS biar konsisten"""
    if pd.isna(nis_value):
        return ""
    return str(nis_value).strip().replace(" ", "").replace("-", "").lower()


def create_tahsin_tahfidz_page(
    c, width, height, row, tahsin_tahfidz_df=None, wali_kelas_map=None
):
    """
    Page: Laporan Tahsin & Tahfidz
    """
    try:
        center_x = width / 2
        left_margin = 50

        # --- Ambil data biodata
        nis_biodata = normalize_nis(row.get("NIS", ""))
        nama_biodata = str(row.get("Nama", "")).strip().lower()
        print(f"[DEBUG] NIS biodata: {nis_biodata}, Nama biodata: {nama_biodata}")

        # --- Tentukan jenjang
        kelas = row.get("Kelas", "")
        school_level = get_school_level(kelas)
        if school_level == "SMP":
            school_header = "SMP PESANTREN RABBAANII"
        elif school_level == "SMA":
            school_header = "SMA PESANTREN RABBAANII"
        else:
            school_header = "SD PESANTREN RABBAANII"

        # === Background watermark ===
        # bg_logo_path = os.path.join(STATIC_FOLDER, "logo_bg.png")
        # if os.path.exists(bg_logo_path):
        #     try:
        #         c.drawImage(
        #             bg_logo_path,
        #             center_x - 150,
        #             height / 2 - 150,
        #             width=300,
        #             height=300,
        #             mask="auto",
        #         )
            # except Exception as e:
            #     print(f"[DEBUG] Gagal load background: {e}")

        # === Header logos ===
        logo_left_path = os.path.join(STATIC_FOLDER, "logo.png")
        logo_right_path = os.path.join(STATIC_FOLDER, "logo2.png")
        logo_size, logo_y = 70, height - 90
        if os.path.exists(logo_left_path):
            c.drawImage(logo_left_path, 40, logo_y, logo_size, logo_size, mask="auto")
        if os.path.exists(logo_right_path):
            c.drawImage(
                logo_right_path,
                width - 40 - logo_size,
                logo_y,
                logo_size,
                logo_size,
                mask="auto",
            )

        # === Header text ===
        header_y = height - 40
        c.setFont("Times-Bold", 13)
        c.drawCentredString(
            center_x, header_y, "LAPORAN PENILAIAN SUMATIF AKHIR SEMESTER I (PSAS I)"
            
        )
        prefix = school_header.replace("RABBAANII", "").strip() + " "
        middle = "RABBAANII"
        prefix_width = c.stringWidth(prefix, "Times-Bold", 13)
        middle_width = c.stringWidth(middle, "Calligrapher", 15)
        start_x = center_x - (prefix_width + middle_width) / 2
        y_pos = header_y - 18
        c.setFont("Times-Bold", 13)
        c.drawString(start_x, y_pos, prefix)
        c.setFont("Calligrapher", 15)
        c.drawString(start_x + prefix_width, y_pos, middle)
        c.setFont("Times-Bold", 13)
        c.drawCentredString(center_x, header_y - 36, "TAHSIN DAN TAHFIDZ")
        c.drawCentredString(center_x, header_y - 54, "TAHUN PELAJARAN 2025/2026")

        # === Student info ===
        info_y = header_y - 80
        c.setFont("Times-Roman", 11)

        pembimbing, penguji = "-", "-"
        student_data = pd.DataFrame()

        # --- Cari di DataFrame Excel ---
        if tahsin_tahfidz_df is not None and not tahsin_tahfidz_df.empty:
            df = tahsin_tahfidz_df.copy()
            df["NIS_norm"] = df["NIS"].apply(normalize_nis)

            # Debug sample
            print("[DEBUG] Sample NIS Excel:", df["NIS"].head(3).tolist())
            print("[DEBUG] Sample NIS_norm Excel:", df["NIS_norm"].head(3).tolist())
            print("[DEBUG] Sample Nama Excel:", df["Nama"].head(3).tolist())

            # Cari berdasarkan NIS
            student_data = df[df["NIS_norm"] == nis_biodata]

            # Jika gagal, coba pakai nama
            if student_data.empty:
                print(f"[DEBUG] Tidak ketemu NIS: {nis_biodata}, coba pakai nama...")
                student_data = df[
                    df["Nama"].astype(str).str.strip().str.lower() == nama_biodata
                ]

            if student_data.empty:
                print(f"[DEBUG] Masih gagal match: {nis_biodata} / {nama_biodata}")
            else:
                print(f"[DEBUG] Data ditemukan untuk {row.get('Nama','')}:")
                print(student_data.to_dict("records")[0])
                student_row = student_data.iloc[0]
                pembimbing = student_row.get("Pembimbing", "-")
                penguji = student_row.get("Penguji", "-")

        # --- Cetak info siswa ---
        c.drawString(50, info_y, "Nama")
        c.drawString(120, info_y, f": {row.get('Nama', '')}")
        c.drawString(350, info_y, "Pembimbing")
        c.drawString(420, info_y, f": {pembimbing}")
        c.drawString(50, info_y - 15, "Kelas")
        c.drawString(120, info_y - 15, f": {row.get('Kelas', '')}")
        c.drawString(350, info_y - 15, "Penguji")
        c.drawString(420, info_y - 15, f": {penguji}")

        # === Setup tabel ===
        table_start_y = info_y - 60
        table_x = left_margin
        col_no, col_mapel = 40, 180
        col_kkm, col_kel, col_taj, col_makh, col_nilai = 50, 70, 60, 60, 70
        col_widths = [col_kkm, col_kel, col_taj, col_makh, col_nilai]
        sub_headers = ["KKM", "Kelancaran", "Tajwid", "Makhroj", "Nilai"]
        row_height = 25
        current_y = table_start_y

        # Header utama
        c.setFillColor(colors.lightgrey)
        c.rect(table_x, current_y - row_height, col_no, row_height, fill=1)
        c.rect(table_x + col_no, current_y - row_height, col_mapel, row_height, fill=1)
        c.rect(
            table_x + col_no + col_mapel,
            current_y - row_height,
            sum(col_widths),
            row_height,
            fill=1,
        )
        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 13)
        draw_centered_text(
            c, "NO", table_x, current_y - row_height, col_no, row_height, bold=True
        )
        draw_centered_text(
            c,
            "MATA PELAJARAN",
            table_x + col_no,
            current_y - row_height,
            col_mapel,
            row_height,
            bold=True,
        )
        draw_centered_text(
            c,
            "NILAI TAHSIN DAN TAHFIDZ",
            table_x + col_no + col_mapel,
            current_y - row_height,
            sum(col_widths),
            row_height,
            bold=True,
        )
        current_y -= row_height


       # TAHSIN header
        c.setFillColor(colors.lightgrey)
        c.rect(table_x, current_y - row_height, col_no, row_height, fill=1)
        c.rect(table_x + col_no, current_y - row_height, col_mapel, row_height, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 13)
        draw_centered_text(c, "TAHSIN", table_x + col_no, current_y - row_height, col_mapel, row_height, bold=True)

        x_pos = table_x + col_no + col_mapel
        for header, w in zip(sub_headers, col_widths):
            c.setFillColor(colors.lightgrey)
            c.rect(x_pos, current_y - row_height, w, row_height, fill=1)
            c.setFillColor(colors.black)
            draw_centered_text(c, header, x_pos, current_y - row_height, w, row_height, bold=True)
            x_pos += w
        current_y -= row_height

        # Data TAHSIN
        if not student_data.empty:
            tahsin_row = student_data.iloc[0]
            
            # Nomor 1
            c.rect(table_x, current_y - row_height, col_no, row_height)
            draw_centered_text(c, "1", table_x, current_y - row_height, col_no, row_height)
            
            # Deskripsi dari Excel dengan Paragraph untuk auto-wrap
            c.rect(table_x + col_no, current_y - row_height, col_mapel, row_height)
            tahsin_desc = str(tahsin_row.get("Tahsin_Surah", ""))
            
            # Import Paragraph dari reportlab
            from reportlab.platypus import Paragraph
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_LEFT
            
            # Buat style untuk paragraph
            styles = getSampleStyleSheet()
            custom_style = ParagraphStyle(
                'CustomStyle',
                parent=styles['Normal'],
                fontName='Times-Roman',
                fontSize=10,
                leading=12,
                alignment=TA_LEFT,
                spaceAfter=0,
                spaceBefore=0,
            )
            
            # Buat paragraph dengan auto-wrap
            para = Paragraph(tahsin_desc, custom_style)
            para_width = col_mapel - 10
            para_height = row_height - 4
            
            # Wrap and draw paragraph
            w, h = para.wrap(para_width, para_height)
            para.drawOn(
                    c,
                    table_x + col_no + 5,
                    current_y - row_height + (row_height - h) / 2
                )

            # Nilai-nilai TAHSIN dari Excel
            tahsin_vals = [
                format_score(tahsin_row.get("Tahsin_KKM", "")),
                format_score(tahsin_row.get("Tahsin_Kelancaran", "")),
                format_score(tahsin_row.get("Tahsin_Tajwid", "")),
                format_score(tahsin_row.get("Tahsin_Makhroj", "")),
                format_score(tahsin_row.get("Tahsin_Nilai", "")),
            ]
            
            x_pos = table_x + col_no + col_mapel
            c.setFont("Times-Roman", 10)
            for v, w in zip(tahsin_vals, col_widths):
                c.rect(x_pos, current_y - row_height, w, row_height)
                draw_centered_text(c, str(v), x_pos, current_y - row_height, w, row_height)
                x_pos += w
            current_y -= row_height
        # TAHFIDZ header
        c.setFillColor(colors.lightgrey)
        c.rect(table_x, current_y - row_height, col_no, row_height, fill=1)
        c.rect(table_x + col_no, current_y - row_height, col_mapel, row_height, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 13)
        draw_centered_text(c, "TAHFIDZ", table_x + col_no, current_y - row_height, col_mapel, row_height, bold=True)


        x_pos = table_x + col_no + col_mapel
        for header, w in zip(sub_headers, col_widths):
            c.setFillColor(colors.lightgrey)
            c.rect(x_pos, current_y - row_height, w, row_height, fill=1)
            c.setFillColor(colors.black)
            draw_centered_text(
                c, header, x_pos, current_y - row_height, w, row_height, bold=True
            )
            x_pos += w
        current_y -= row_height

        # Data TAHFIDZ
        if not student_data.empty:
            tahfidz_row = student_data.iloc[0]
            c.rect(table_x, current_y - row_height, col_no, row_height)
            draw_centered_text(
                c, "2", table_x, current_y - row_height, col_no, row_height
            )
            c.rect(table_x + col_no, current_y - row_height, col_mapel, row_height)
            c.drawString(
                table_x + col_no + 5,
                current_y - row_height / 2 - 3,
                str(tahfidz_row.get("Tahfidz_Surah", "")),
            )
            tahfidz_vals = [
                format_score(tahfidz_row.get("Tahfidz_KKM", "")),
                format_score(tahfidz_row.get("Tahfidz_Kelancaran", "")),
                format_score(tahfidz_row.get("Tahfidz_Tajwid", "")),
                format_score(tahfidz_row.get("Tahfidz_Makhroj", "")),
                format_score(tahfidz_row.get("Tahfidz_Nilai", "")),
            ]

            x_pos = table_x + col_no + col_mapel
            for v, w in zip(tahfidz_vals, col_widths):
                c.rect(x_pos, current_y - row_height, w, row_height)
                draw_centered_text(
                    c, str(v), x_pos, current_y - row_height, w, row_height
                )
                x_pos += w
            current_y -= row_height

        # Signature
        draw_signatures(c, width, height, left_margin, row, wali_kelas_map, nip_map)

    except Exception as e:
        print(f"Error create_tahsin_tahfidz_page: {e}")
        print(traceback.format_exc())


# === Rubrik Penilaian Karakter 1-10 ===
KARAKTER_RUBRIK = {
    "Integritas": {
        1: "Ananda sangat sering melanggar aturan, tidak jujur, dan sering mengabaikan tanggung jawabnya. Belum memahami arti komitmen dan konsekuensi.",
        2: "Ananda jarang menunjukkan kejujuran, sering menyalahi aturan, serta mengabaikan tugas sederhana yang diberikan.",
        3: "Ananda kadang mematuhi aturan tetapi masih sering lalai. Tanggung jawab belum dijalankan dengan baik dan masih bergantung pada arahan guru.",
        4: "Ananda mulai belajar bersikap jujur dan disiplin, meski masih tidak konsisten dan kadang mencari alasan untuk menghindari kewajiban.",
        5: "Ananda cukup mampu menjaga integritas. Mulai menunjukkan sikap jujur meski masih sesekali melakukan pelanggaran kecil.",
        6: "Ananda cukup disiplin dan bertanggung jawab pada sebagian besar tugas, meskipun konsistensinya masih belum stabil.",
        7: "Ananda jujur, disiplin, dan bertanggung jawab dalam banyak kesempatan. Mulai menjadi contoh di lingkungannya.",
        8: "Ananda memiliki integritas yang baik, konsisten menjalankan kewajiban, dan menunjukkan sikap tanggung jawab dalam kegiatan kelompok.",
        9: "Ananda menjunjung tinggi kejujuran, disiplin, dan tanggung jawab. Sikapnya menjadi teladan yang baik bagi teman-temannya.",
        10: "Ananda konsisten penuh dalam kejujuran, disiplin, dan tanggung jawab. Mampu menjadi panutan dalam menjaga integritas di lingkungan sekolah.",
    },
    "Religius": {
        1: "Ananda jarang melaksanakan ibadah, tidak terbiasa berdoa, dan kurang memahami nilai keagamaan.",
        2: "Ananda melaksanakan ibadah hanya bila diingatkan. Doa harian jarang dilakukan.",
        3: "Ananda mulai melaksanakan ibadah tetapi belum rutin. Masih sering lalai atau lupa berdoa.",
        4: "Ananda sesekali berdoa sebelum dan sesudah kegiatan, tetapi belum konsisten melakukannya.",
        5: "Ananda cukup konsisten beribadah, meski kadang masih kurang khusyuk atau terlewat.",
        6: "Ananda beribadah secara teratur dan mulai membiasakan doa dalam aktivitas sehari-hari.",
        7: "Ananda rajin beribadah, berdoa sebelum dan sesudah kegiatan, serta menghargai nilai religius.",
        8: "Ananda religius, menjaga akhlak sehari-hari, melaksanakan ibadah tepat waktu, dan menjadi contoh bagi teman-temannya.",
        9: "Ananda sangat religius, istiqamah beribadah, menunjukkan akhlak mulia, serta memberi pengaruh positif pada lingkungannya.",
        10: "Ananda sangat religius, konsisten menjaga ibadah, penuh akhlak mulia, dan menjadi teladan utama dalam sikap Islami.",
    },
    "Nasionalis": {
        1: "Ananda tidak menunjukkan rasa cinta tanah air, sering mengabaikan simbol bangsa, dan tidak peduli terhadap lingkungan.",
        2: "Ananda kurang menghargai simbol negara dan jarang terlibat dalam kegiatan kebangsaan.",
        3: "Ananda mulai mengenal simbol kebangsaan meskipun belum menunjukkan kebanggaan nyata.",
        4: "Ananda kadang menunjukkan kepedulian terhadap kegiatan kebangsaan di sekolah.",
        5: "Ananda cukup menghargai simbol bangsa dan mulai menumbuhkan rasa nasionalisme.",
        6: "Ananda memiliki kesadaran nasional yang baik, tetapi penerapannya masih terbatas pada kegiatan tertentu.",
        7: "Ananda bangga sebagai warga negara Indonesia dan menunjukkan nasionalisme dalam kegiatan sekolah.",
        8: "Ananda aktif dalam kegiatan kebangsaan, menghargai budaya, dan mendukung persatuan.",
        9: "Ananda memiliki nasionalisme tinggi, konsisten menghargai keberagaman, dan bangga pada bangsa Indonesia.",
        10: "Ananda menjadi teladan nasionalis, cinta tanah air, aktif menjaga persatuan, dan membela nilai bangsa dalam berbagai kesempatan.",
    },
    "Mandiri": {
        1: "Ananda sangat bergantung pada orang lain dan tidak mampu menyelesaikan tugas tanpa bantuan.",
        2: "Ananda sering meminta bantuan bahkan untuk tugas sederhana.",
        3: "Ananda mulai mencoba mandiri, tetapi sering masih membutuhkan arahan intensif.",
        4: "Ananda kadang bisa menyelesaikan tugas sendiri, namun masih belum percaya diri.",
        5: "Ananda cukup mandiri, berusaha menyelesaikan tugas meski sesekali butuh bantuan.",
        6: "Ananda mampu mengatasi kesulitan sederhana secara mandiri.",
        7: "Ananda mandiri, menyelesaikan tugas dengan baik tanpa banyak bergantung pada orang lain.",
        8: "Ananda percaya diri, tangguh, dan dapat mengambil keputusan secara mandiri.",
        9: "Ananda mandiri dalam hampir semua aspek belajar dan menjadi contoh bagi teman-temannya.",
        10: "Ananda sangat mandiri, percaya diri tinggi, mampu mengatasi tantangan kompleks secara efektif.",
    },
    "Gotong Royong": {
        1: "Ananda enggan bekerja sama, cenderung individualis, dan kurang peduli terhadap orang lain.",
        2: "Ananda jarang terlibat dalam kerja kelompok meski diminta.",
        3: "Ananda mulai berpartisipasi dalam kerja kelompok, meski masih pasif.",
        4: "Ananda kadang membantu teman, namun belum konsisten dalam kebersamaan.",
        5: "Ananda cukup kooperatif, mau bekerja sama bila diarahkan.",
        6: "Ananda aktif bekerja sama dalam kelompok, meski masih terbatas pada kegiatan tertentu.",
        7: "Ananda terbiasa bekerja sama dan membantu teman, serta kooperatif dalam kegiatan.",
        8: "Ananda aktif mendukung kerja tim, menunjukkan kepedulian, dan berkontribusi positif.",
        9: "Ananda menjunjung tinggi gotong royong, menjadi teladan dalam kerja kelompok.",
        10: "Ananda selalu menunjukkan semangat kebersamaan, peduli terhadap teman, dan menjadi panutan dalam gotong royong.",
    },
    "Disiplin": {
        1: "Ananda sangat sering melanggar aturan dan tidak mengikuti tata tertib sekolah.",
        2: "Ananda jarang disiplin, sering terlambat, dan mengabaikan tanggung jawab.",
        3: "Ananda mulai berusaha disiplin, meski masih sering lalai.",
        4: "Ananda kadang disiplin dalam kegiatan, tetapi tidak konsisten.",
        5: "Ananda cukup disiplin, mematuhi sebagian besar aturan sekolah.",
        6: "Ananda disiplin dalam banyak hal, meski masih ada kekurangan.",
        7: "Ananda taat pada tata tertib dan konsisten menjaga kedisiplinan.",
        8: "Ananda sangat disiplin, mematuhi aturan, dan menjadi contoh bagi teman-temannya.",
        9: "Ananda memiliki disiplin tinggi, mampu mengingatkan teman-temannya.",
        10: "Ananda menjadi teladan utama dalam kedisiplinan, konsisten dan tegas terhadap aturan.",
    },
    "Sopan Santun": {
        1: "Ananda sering bersikap tidak sopan, menggunakan bahasa yang kurang baik, dan tidak menghormati orang lain.",
        2: "Ananda jarang menunjukkan sikap santun, masih sering berbicara kasar atau tidak menghargai guru.",
        3: "Ananda mulai berusaha bersikap sopan, meski masih belum konsisten.",
        4: "Ananda kadang menunjukkan kesantunan, tetapi hanya dalam situasi tertentu.",
        5: "Ananda cukup sopan, berusaha menghormati guru dan teman-teman.",
        6: "Ananda santun dalam banyak hal, tetapi sesekali masih kurang peka terhadap etika.",
        7: "Ananda terbiasa bersikap sopan, menghargai guru, dan ramah kepada teman.",
        8: "Ananda menunjukkan sikap santun dan ramah, menjaga tutur kata dan perilaku sehari-hari.",
        9: "Ananda santun dalam setiap kesempatan, mampu memberi teladan dalam sikap dan tutur kata.",
        10: "Ananda sangat santun, konsisten menjaga sikap hormat, tutur kata halus, dan menjadi panutan akhlak mulia.",
    },
}


def get_karakter_deskripsi(karakter, nilai):
    try:
        nilai = int(float(nilai))
    except:
        return ""
    rubrik = KARAKTER_RUBRIK.get(karakter, {})
    return rubrik.get(nilai, "")


# === Fungsi Utama ===
def create_karakter_page(c, width, height, row, karakter_df=None, wali_kelas_map=None):
    """Page: Laporan Perkembangan Karakter (langsung pakai deskripsi dari Excel)"""
    try:
        center_x = width / 2
        left_margin = 50

        # # === Background watermark logo ===
        # bg_logo_path = os.path.join(STATIC_FOLDER, "logo_bg.png")
        # if os.path.exists(bg_logo_path):
        #     bg_logo = ImageReader(bg_logo_path)
        #     bg_size = 300
        #     c.drawImage(
        #         bg_logo,
        #         center_x - bg_size / 2,
        #         height / 2 - bg_size / 2,
        #         width=bg_size,
        #         height=bg_size,
        #         mask="auto",
        #     )

        # === Header logos kiri & kanan ===
        logo_left_path = os.path.join(STATIC_FOLDER, "logo.png")
        logo_right_path = os.path.join(STATIC_FOLDER, "logo2.png")

        logo_size = 70
        logo_y = height - 90

        if os.path.exists(logo_left_path):
            logo_left = ImageReader(logo_left_path)
            c.drawImage(
                logo_left, 40, logo_y, width=logo_size, height=logo_size, mask="auto"
            )

        if os.path.exists(logo_right_path):
            logo_right = ImageReader(logo_right_path)
            c.drawImage(
                logo_right,
                width - 40 - logo_size,
                logo_y,
                width=logo_size,
                height=logo_size,
                mask="auto",
            )

        # === Header text ===
        header_y = height - 40
        line_gap = 18
        c.setFont("Times-Bold", 13)
        c.setFillColor(colors.black)
        c.drawCentredString(
            center_x, header_y, "LAPORAN PENILAIAN SUMATIF AKHIR SEMESTER I (PSAS I)"
        )

        # Baris 2 → split agar RABBAANII pakai calligrapher
        prefix = f"{row.get('Kelas', '')} PESANTREN "
        middle = "RABBAANII"

        prefix_width = c.stringWidth(prefix, "Times-Bold", 13)
        middle_width = c.stringWidth(middle, "Calligrapher", 15)
        total_width = prefix_width + middle_width
        start_x = center_x - (total_width / 2)
        y_pos = header_y - line_gap

        c.setFont("Times-Bold", 13)
        c.drawString(start_x, y_pos, prefix)

        c.setFont("Calligrapher", 15)
        c.drawString(start_x + prefix_width, y_pos, middle)

        # Baris 3 & 4
        c.setFont("Times-Bold", 13)
        c.drawCentredString(
            center_x, header_y - 2 * line_gap, "TAHUN PELAJARAN 2025/2026"
        )
        c.drawCentredString(
            center_x, header_y - 3 * line_gap, "KECAMATAN CIKARANG TIMUR"
        )

        # === Informasi siswa ===
        info_y = header_y - 100
        c.setFont("Times-Roman", 11)
        c.drawString(50, info_y, "Nama")
        c.drawString(120, info_y, f": {row.get('Nama', '')}")
        c.drawString(350, info_y, "Kelas")
        c.drawString(420, info_y, f": {row.get('Kelas', '')}")

        c.drawString(50, info_y - 15, "NIS / NISN")
        nis_nisn = f"{row.get('NIS', '')} / {row.get('NISN', '')}"
        c.drawString(120, info_y - 15, f": {nis_nisn}")
        c.drawString(350, info_y - 15, "Semester")
        c.drawString(420, info_y - 15, ": I (SATU)")

        # === Ambil data karakter berdasarkan NIS ===
        karakter_data = {}
        if isinstance(karakter_df, pd.DataFrame) and not karakter_df.empty:
            nis = row.get("NIS", "")
            student_row = karakter_df[karakter_df["NIS"] == nis]
            if not student_row.empty:
                karakter_data = student_row.iloc[0].to_dict()

        # === Setup Tabel ===
        table_y = info_y - 60
        table_x = left_margin
        table_width = width - 2 * left_margin

        col_no_width = 30
        col_karakter_width = 120
        col_deskripsi_width = table_width - (col_no_width + col_karakter_width)

        row_height = 50
        header_height = 30

        # Header tabel
        c.setFillColor(colors.lightgrey)
        c.rect(table_x, table_y - header_height, col_no_width, header_height, fill=1)
        c.rect(
            table_x + col_no_width,
            table_y - header_height,
            col_karakter_width,
            header_height,
            fill=1,
        )
        c.rect(
            table_x + col_no_width + col_karakter_width,
            table_y - header_height,
            col_deskripsi_width,
            header_height,
            fill=1,
        )

        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 11)
        draw_centered_text(
            c, "No", table_x, table_y - header_height, col_no_width, header_height
        )
        draw_centered_text(
            c,
            "Karakter",
            table_x + col_no_width,
            table_y - header_height,
            col_karakter_width,
            header_height,
        )
        draw_centered_text(
            c,
            "Deskripsi",
            table_x + col_no_width + col_karakter_width,
            table_y - header_height,
            col_deskripsi_width,
            header_height,
        )

        # === Isi Tabel ===
        current_y = table_y - header_height
        karakter_map = {
            "Integritas": "Integritas_Deskripsi",
            "Religius": "Religius_Deskripsi",
            "Nasionalis": "Nasionalis_Deskripsi",
            "Mandiri": "Mandiri_Deskripsi",
            "Gotong Royong": "GotongRoyong_Deskripsi",
            "Disiplin": "Disiplin_Deskripsi",
            "Sopan Santun": "SopanSantun_Deskripsi",
        }

        for idx, (karakter, col_name) in enumerate(karakter_map.items(), start=1):
            nilai = karakter_data.get(col_name, "")
            deskripsi = get_karakter_deskripsi(karakter, nilai)

            if not deskripsi:
                continue

            # Kotak row
            c.rect(table_x, current_y - row_height, col_no_width, row_height)
            c.rect(
                table_x + col_no_width,
                current_y - row_height,
                col_karakter_width,
                row_height,
            )
            c.rect(
                table_x + col_no_width + col_karakter_width,
                current_y - row_height,
                col_deskripsi_width,
                row_height,
            )

            # Isi teks
            c.setFont("Times-Roman", 11)
            draw_centered_text(
                c, str(idx), table_x, current_y - row_height, col_no_width, row_height
            )
            draw_centered_text(
                c,
                karakter,
                table_x + col_no_width,
                current_y - row_height,
                col_karakter_width,
                row_height,
            )

            # Wrap teks deskripsi
            words = deskripsi.split()
            line = ""
            text_y = current_y - 15
            for word in words:
                if (
                    c.stringWidth(line + " " + word, "Times-Roman", 10)
                    < col_deskripsi_width - 10
                ):
                    line += " " + word
                else:
                    c.drawString(
                        table_x + col_no_width + col_karakter_width + 5,
                        text_y,
                        line.strip(),
                    )
                    text_y -= 12
                    line = word
            if line:
                c.drawString(
                    table_x + col_no_width + col_karakter_width + 5,
                    text_y,
                    line.strip(),
                )

            current_y -= row_height

        # === Tanda tangan ===
        draw_signatures(c, width, height, left_margin, row, wali_kelas_map, nip_map)

    except Exception as e:
        print(f"Error creating karakter page: {e}")
        raise


def draw_table_row(
    c,
    start_x,
    y,
    cells,
    height,
    align_first_left=False,
    align_second_left=False,
    bold=False,
):
    """Helper function to draw a table row with formal fonts"""
    current_x = start_x

    for i, (width, text) in enumerate(cells):
        # Draw cell border
        c.rect(current_x, y, width, height, fill=0)

        # Determine alignment and font
        if (i == 0 and align_first_left) or (i == 1 and align_second_left):
            # Left align with padding
            if bold:
                c.setFont("Times-Bold", 10)
            else:
                c.setFont("Times-Roman", 10)
            c.drawString(current_x + 8, y + 9, text)
        else:
            # Center align
            draw_centered_text(
                c, text, current_x, y, width, height, font="Times-Roman", bold=bold
            )

        current_x += width


@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/upload_biodata", methods=["POST"])
def upload_biodata():
    global biodata_storage
    try:
        file = request.files.get("file")
        if not file or file.filename == "":
            return jsonify({"status": "error", "message": "No file uploaded"}), 400

        # Save file temporarily
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Read Excel file
        biodata_storage = pd.read_excel(filepath, sheet_name=0)

        # Clean up temporary file
        os.remove(filepath)

        return jsonify(
            {
                "status": "success",
                "message": f"Biodata uploaded successfully! {len(biodata_storage)} students found.",
            }
        )

    except Exception as e:
        print(f"Error uploading biodata: {e}")
        return (
            jsonify({"status": "error", "message": f"Error processing file: {str(e)}"}),
            500,
        )


@app.route("/upload_nilai", methods=["POST"])
def upload_nilai():
    global nilai_storage
    try:
        file = request.files.get("file")
        kelas = normalize_kelas(request.form.get("kelas"))

        if not file or file.filename == "":
            return jsonify({"status": "error", "message": "No file uploaded"}), 400

        if not kelas:
            return jsonify({"status": "error", "message": "Please select a class"}), 400

        # Save file temporarily
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Read Excel file
        nilai_df = pd.read_excel(filepath, sheet_name=0)
        nilai_storage[kelas] = nilai_df

        # Clean up temporary file
        os.remove(filepath)

        return jsonify(
            {
                "status": "success",
                "message": f"Nilai for class {kelas} uploaded successfully! {len(nilai_df)} records found.",
            }
        )

    except Exception as e:
        print(f"Error uploading nilai: {e}")
        return (
            jsonify({"status": "error", "message": f"Error processing file: {str(e)}"}),
            500,
        )


@app.route("/upload_kompetensi", methods=["POST"])
def upload_kompetensi():
    global kompetensi_storage
    try:
        file = request.files.get("file")
        kelas = normalize_kelas(request.form.get("kelas"))

        if not file or file.filename == "":
            return jsonify({"status": "error", "message": "No file uploaded"}), 400

        if not kelas:
            return jsonify({"status": "error", "message": "Please select a class"}), 400

        # Save file temporarily
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Read Excel file
        kompetensi_df = pd.read_excel(filepath, sheet_name=0)
        kompetensi_storage[kelas] = kompetensi_df

        # Clean up temporary file
        os.remove(filepath)

        return jsonify(
            {
                "status": "success",
                "message": f"Kompetensi for class {kelas} uploaded successfully! {len(kompetensi_df)} records found.",
            }
        )

    except Exception as e:
        print(f"Error uploading kompetensi: {e}")
        return (
            jsonify({"status": "error", "message": f"Error processing file: {str(e)}"}),
            500,
        )


@app.route("/upload_tahsin_tahfidz", methods=["POST"])
def upload_tahsin_tahfidz():
    global tahsin_tahfidz_storage
    try:
        file = request.files.get("file")
        kelas = normalize_kelas(request.form.get("kelas"))

        if not file or file.filename == "":
            return jsonify({"status": "error", "message": "No file uploaded"}), 400

        if not kelas:
            return jsonify({"status": "error", "message": "Please select a class"}), 400

        # Validate file extension
        if not file.filename.lower().endswith((".xlsx", ".xls")):
            return (
                jsonify(
                    {
                        "status": "error",
                        "message": "Please upload an Excel file (.xlsx or .xls)",
                    }
                ),
                400,
            )

        # Save file temporarily
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        try:
            # Read Excel file
            tahsin_tahfidz_df = pd.read_excel(filepath, sheet_name=0)

            # Validate required columns
            required_columns = ["NIS", "Nama", "Pembimbing", "Penguji"]
            missing_columns = [
                col for col in required_columns if col not in tahsin_tahfidz_df.columns
            ]

            if missing_columns:
                os.remove(filepath)
                return (
                    jsonify(
                        {
                            "status": "error",
                            "message": f"Missing required columns: {', '.join(missing_columns)}",
                        }
                    ),
                    400,
                )

            # Optional columns for Tahsin data
            tahsin_columns = [
                "Tahsin_Surah",
                "Tahsin_KKM",
                "Tahsin_Kelancaran",
                "Tahsin_Tajwid",
                "Tahsin_Makhroj",
                "Tahsin_Nilai",
            ]

            # Optional columns for Tahfidz data
            tahfidz_columns = [
                "Tahfidz_Surah",
                "Tahfidz_KKM",
                "Tahfidz_Kelancaran",
                "Tahfidz_Tajwid",
                "Tahfidz_Makhroj",
                "Tahfidz_Nilai",
            ]

            # Check if at least one set of data exists
            has_tahsin = all(col in tahsin_tahfidz_df.columns for col in tahsin_columns)
            has_tahfidz = all(
                col in tahsin_tahfidz_df.columns for col in tahfidz_columns
            )

            if not has_tahsin and not has_tahfidz:
                os.remove(filepath)
                return (
                    jsonify(
                        {
                            "status": "error",
                            "message": "Excel file must contain either Tahsin or Tahfidz data columns",
                        }
                    ),
                    400,
                )

            # Clean data - remove empty rows
            tahsin_tahfidz_df = tahsin_tahfidz_df.dropna(subset=["NIS", "Nama"])

            # Store the data
            if tahsin_tahfidz_storage is None:
                tahsin_tahfidz_storage = {}

            tahsin_tahfidz_storage[kelas] = tahsin_tahfidz_df

            # Clean up temporary file
            os.remove(filepath)

            data_type = []
            if has_tahsin:
                data_type.append("Tahsin")
            if has_tahfidz:
                data_type.append("Tahfidz")

            return jsonify(
                {
                    "status": "success",
                    "message": f"Tahsin Tahfidz data for class {kelas} uploaded successfully! {len(tahsin_tahfidz_df)} records found with {' and '.join(data_type)} data.",
                    "data_info": {
                        "class": kelas,
                        "records": len(tahsin_tahfidz_df),
                        "has_tahsin": has_tahsin,
                        "has_tahfidz": has_tahfidz,
                        "columns": list(tahsin_tahfidz_df.columns),
                    },
                }
            )

        except Exception as e:
            # Clean up temporary file if error occurs
            if os.path.exists(filepath):
                os.remove(filepath)
            raise e

    except Exception as e:
        print(f"Error uploading tahsin tahfidz: {e}")
        return (
            jsonify({"status": "error", "message": f"Error processing file: {str(e)}"}),
            500,
        )


@app.route("/upload_karakter", methods=["POST"])
def upload_karakter():
    global karakter_storage
    try:
        file = request.files.get("file")
        kelas = normalize_kelas(request.form.get("kelas"))

        if not file or file.filename == "":
            return jsonify({"status": "error", "message": "No file uploaded"}), 400

        if not kelas:
            return jsonify({"status": "error", "message": "Please select a class"}), 400

        # Save file temporarily
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Read Excel file
        karakter_df = pd.read_excel(filepath, sheet_name=0)
        karakter_storage[kelas] = karakter_df

        # Clean up temporary file
        os.remove(filepath)

        return jsonify(
            {
                "status": "success",
                "message": f"Karakter for class {kelas} uploaded successfully! {len(karakter_df)} records found.",
            }
        )

    except Exception as e:
        print(f"Error uploading karakter: {e}")
        return (
            jsonify({"status": "error", "message": f"Error processing file: {str(e)}"}),
            500,
        )


@app.route("/generate_pdf", methods=["POST"])
def generate_pdf():
    global biodata_storage, nilai_storage, kompetensi_storage, tahsin_tahfidz_storage, karakter_storage, wali_kelas_map

    try:
        print("=== Starting PDF generation ===")

        if biodata_storage is None:
            return (
                jsonify({"status": "error", "message": "Please upload biodata first"}),
                400,
            )

        if not nilai_storage:
            return (
                jsonify(
                    {"status": "error", "message": "Please upload nilai data first"}
                ),
                400,
            )

        # Direktori sementara
        temp_dir = tempfile.mkdtemp()
        print(f"Created temp directory: {temp_dir}")

        pdf_files = []
        processed_students = 0

        try:
            for idx, row in biodata_storage.iterrows():
                try:
                    kelas = normalize_kelas(row.get("Kelas", ""))
                    if not kelas or kelas not in nilai_storage:
                        print(
                            f"Skipping student {idx+1} {row.get('Nama')}: kelas={kelas}, available={list(nilai_storage.keys())}"
                        )
                        continue

                    nilai_df = nilai_storage[kelas]
                    nis = row.get("NIS", "")
                    student_nilai = (
                        nilai_df[nilai_df["NIS"] == nis] if nis else pd.DataFrame()
                    )

                    school_level = get_school_level(kelas)
                    logo_path = os.path.join(STATIC_FOLDER, "logo.png")

                    # Handle kompetensi data for SMA
                    kompetensi_df = None
                    if (
                        school_level == "SMA"
                        and kompetensi_storage
                        and kelas in kompetensi_storage
                    ):
                        kompetensi_df = kompetensi_storage[kelas]
                        if nis and "NIS" in kompetensi_df.columns:
                            kompetensi_df = kompetensi_df[kompetensi_df["NIS"] == nis]

                    # Handle tahsin tahfidz data for all levels
                    tahsin_tahfidz_df = None
                    if tahsin_tahfidz_storage and kelas in tahsin_tahfidz_storage:
                        tahsin_tahfidz_df = tahsin_tahfidz_storage[kelas]
                        student_name = row.get("Nama", "")
                        if student_name and "Nama" in tahsin_tahfidz_df.columns:
                            tahsin_tahfidz_df = tahsin_tahfidz_df[
                                tahsin_tahfidz_df["Nama"] == student_name
                            ]
                        elif nis and "NIS" in tahsin_tahfidz_df.columns:
                            tahsin_tahfidz_df = tahsin_tahfidz_df[
                                tahsin_tahfidz_df["NIS"] == nis
                            ]

                    # Handle karakter data for all levels (SMP and SMA)
                    karakter_df = None
                    if karakter_storage and kelas in karakter_storage:
                        karakter_df = karakter_storage[kelas]
                        student_name = row.get("Nama", "")
                        if student_name and "Nama" in karakter_df.columns:
                            karakter_df = karakter_df[
                                karakter_df["Nama"].str.strip().str.lower()
                                == student_name.strip().lower()
                            ]
                        elif nis and "NIS" in karakter_df.columns:
                            karakter_df = karakter_df[karakter_df["NIS"] == nis]

                    student_name = row.get("Nama", f"Student_{idx+1}")
                    clean_name = "".join(
                        [c if c.isalnum() else "_" for c in student_name]
                    )
                    pdf_path = os.path.join(temp_dir, f"{clean_name}.pdf")

                    # === Generate PDF ===
                    buffer = io.BytesIO()
                    c = canvas.Canvas(buffer, pagesize=A4)
                    width, height = A4

                    # 1. Cover Page (for all levels)
                    create_cover_page(
                        c,
                        width,
                        height,
                        row.get("NIS", ""),
                        row.get("NISN", ""),
                        row.get("Nama", f"Student {idx+1}"),
                        logo_path,
                        school_level,
                    )
                    c.showPage()

                    # 2. Data Page (for all levels)
                    create_data_page(c, width, height, row, school_level)
                    c.showPage()

                    # 3. Score Page (for all levels) - FIXED
                    create_score_page(
                        c,
                        width,
                        height,
                        row,
                        student_nilai,
                        wali_kelas_map,
                        school_level,
                    )
                    c.showPage()

                    # 4. Level-specific pages
                    if school_level == "SMA":
                        # Kompetensi Page for SMA - FIXED
                        create_kompetensi_page(
                            c,
                            width,
                            height,
                            row,
                            kompetensi_df,
                            school_level,
                            wali_kelas_map,
                        )
                        c.showPage()

                    # 5. Tahsin Tahfidz Page for all levels (SMP and SMA) - FIXED
                    if tahsin_tahfidz_df is not None or kelas in tahsin_tahfidz_storage:
                        create_tahsin_tahfidz_page(
                            c, width, height, row, tahsin_tahfidz_df, wali_kelas_map
                        )
                        c.showPage()

                    # 6. Karakter Page for all levels (SMP and SMA) - FIXED
                    if school_level in ["SMP", "SMA"]:  # Exclude SD for now
                        create_karakter_page(
                            c, width, height, row, karakter_df, wali_kelas_map
                        )
                        c.showPage()

                    # Note: SD only gets cover, data, and score pages (3 pages total)

                    c.save()
                    buffer.seek(0)

                    with open(pdf_path, "wb") as f:
                        f.write(buffer.getvalue())

                    pdf_files.append(pdf_path)
                    processed_students += 1

                    # Log page count based on school level
                    page_count = 3  # Base pages: Cover, Data, Score
                    if school_level == "SMA":
                        page_count += 1  # Add Kompetensi
                    if tahsin_tahfidz_df is not None or kelas in tahsin_tahfidz_storage:
                        page_count += 1  # Add Tahsin Tahfidz
                    if school_level in ["SMP", "SMA"]:
                        page_count += 1  # Add Karakter

                    print(
                        f"Generated {page_count} pages for {student_name} ({school_level})"
                    )

                except Exception as e:
                    print(f"Error processing student {idx+1} ({student_name}): {e}")
                    print(traceback.format_exc())
                    continue

            if processed_students == 0:
                shutil.rmtree(temp_dir)
                return (
                    jsonify(
                        {"status": "error", "message": "No students could be processed"}
                    ),
                    400,
                )

             # === Buat ZIP ===
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for pdf_path in pdf_files:
                    zip_file.write(pdf_path, os.path.basename(pdf_path))
            zip_buffer.seek(0)

            # Hapus folder sementara
            shutil.rmtree(temp_dir)
            print(f"Successfully processed {processed_students} students")
            print("Cleaned up temp directory")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"raport_siswa_{timestamp}.zip"

            return Response(
                zip_buffer.getvalue(),
                mimetype="application/zip",
                headers={
                    "Content-Disposition": f"attachment; filename={zip_filename}",
                    "Cache-Control": "no-cache, no-store, must-revalidate",
                    "Pragma": "no-cache",
                    "Expires": "0",
                },
            )

        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            raise e

    except Exception as e:
        print(f"General error: {e}")
        print(traceback.format_exc())
        return (
            jsonify({"status": "error", "message": f"Internal server error: {str(e)}"}),
            500,
        )


# Update route status
@app.route("/status")
def status():
    global biodata_storage, nilai_storage, kompetensi_storage, tahsin_tahfidz_storage

    biodata_count = len(biodata_storage) if biodata_storage is not None else 0
    nilai_classes = list(nilai_storage.keys()) if nilai_storage else []
    kompetensi_classes = list(kompetensi_storage.keys()) if kompetensi_storage else []
    tahsin_tahfidz_classes = (
        list(tahsin_tahfidz_storage.keys()) if tahsin_tahfidz_storage else []
    )

    return jsonify(
        {
            "biodata_uploaded": biodata_count > 0,
            "biodata_count": biodata_count,
            "nilai_classes": nilai_classes,
            "kompetensi_classes": kompetensi_classes,
            "tahsin_tahfidz_classes": tahsin_tahfidz_classes,
            "ready_to_generate": biodata_count > 0 and len(nilai_classes) > 0,
        }
    )


if __name__ == "__main__":
    app.run(port=5001, debug=True)