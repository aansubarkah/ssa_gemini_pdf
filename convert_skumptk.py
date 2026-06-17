# convert_skumptk.py
# Mengubah Laporan_SKUMPTK_dummy.pdf menjadi Laporan_SKUMPTK.xlsx
#
# CARA PAKAI (pilih salah satu):
#
# A) Google Colab (paling mudah, tanpa instalasi):
#    1. Buka https://colab.research.google.com
#    2. Klik "Upload" dan unggah file Laporan_SKUMPTK_dummy.pdf
#    3. Buat cell baru, salin seluruh kode ini, lalu jalankan (tombol Play).
#    4. File Laporan_SKUMPTK.xlsx akan muncul di panel kiri (ikon folder) -> klik kanan -> Download.
#
# B) Komputer lokal (jika Python sudah terinstall):
#    1. Simpan file ini dan file PDF di folder yang sama.
#    2. Jalankan di terminal:  pip install pdfplumber pandas openpyxl
#    3. Jalankan:  python convert_skumptk.py
#    4. File Laporan_SKUMPTK.xlsx akan dibuat di folder yang sama.

import re
from datetime import datetime
import pandas as pdfp  # dipakai untuk DataFrame

# pdfplumber dipakai untuk membaca PDF
try:
    import pdfplumber
except ImportError:
    print("Library pdfplumber belum terpasang. Jalankan: pip install pdfplumber pandas openpyxl")
    raise

INPUT_PDF = "Laporan_SKUMPTK_dummy.pdf"
OUTPUT_XLSX = "Laporan_SKUMPTK.xlsx"

COLUMNS = [
    "nama", "nip", "tempat", "tanggal_lahir", "jenis_kelamin", "agama",
    "status_kepegawaian", "pangkat", "golongan", "instansi", "masa_kerja",
    "alamat", "nama", "pekerjaan", "tanggal_lahir", "tanggal_perkawinan",
    "keterangan", "status",
]


def get_item(text, n):
    """Ambil nilai setelah tanda ':' pada baris bernomor 'n' (mis. '1. Nama / Nip : ...')."""
    for line in text.splitlines():
        m = re.match(rf"^{n}\.\s*(.+?)\s*:\s*(.*)", line.strip())
        if m:
            return m.group(2).strip()  # kembalikan teks setelah tanda ':'
    return None


def norm_date(s):
    """Ubah teks tanggal menjadi objek tanggal Excel (YYYY-MM-DD)."""
    if not s or s.strip() == "-":
        return None
    s = s.strip()
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def parse_pdf(path):
    rows = []
    num_pages = 0
    with pdfplumber.open(path) as pdf:
        num_pages = len(pdf.pages)
        for page in pdf.pages:
            txt = page.extract_text()

            # --- Data Pegawai ---
            v1 = get_item(txt, 1) or ""
            nama = v1.split("/")[0].strip()
            nip = v1.split("/")[-1].strip()

            v2 = get_item(txt, 2) or ""
            tempat = v2.split("/")[0].strip()
            tgl_pegawai = norm_date(v2.split("/")[-1])

            jenis_kelamin = get_item(txt, 3)
            agama = get_item(txt, 4)
            status_kepegawaian = get_item(txt, 5)

            v7 = get_item(txt, 7) or ""
            gm = re.match(r"\(([^)]+)\)\s*(.*)", v7)
            golongan, pangkat = (gm.group(1), gm.group(2)) if gm else ("", v7)

            instansi = get_item(txt, 8)
            masa_kerja = get_item(txt, 9)
            alamat = get_item(txt, 11)

            # --- Tabel Susunan Keluarga ---
            for table in page.extract_tables():
                for r in table[1:]:  # lewati baris judul
                    if not any(r):
                        continue
                    # kolom tabel: No, Nama, Pekerjaan, Tanggal Lahir, Tanggal Perkawinan, Keterangan, Status
                    rows.append([
                        nama, nip, tempat, tgl_pegawai, jenis_kelamin, agama,
                        status_kepegawaian, pangkat, golongan, instansi, masa_kerja,
                        alamat,
                        r[1],                         # nama (keluarga)
                        r[2],                         # pekerjaan
                        norm_date(r[3]) if r[3] != "-" else None,   # tanggal_lahir
                        norm_date(r[4]) if r[4] != "-" else None,   # tanggal_perkawinan
                        r[5],                         # keterangan
                        r[6],                         # status
                    ])
    return rows, num_pages


def main():
    rows, num_pages = parse_pdf(INPUT_PDF)
    df = pdfp.DataFrame(rows, columns=COLUMNS)
    df.to_excel(OUTPUT_XLSX, index=False, engine="openpyxl")

    print(f"Jumlah halaman dibaca : {num_pages}")
    print(f"Total baris Excel     : {len(df)}")
    print("\n5 baris pertama:")
    print(df.head().to_string())
    print(f"\nFile '{OUTPUT_XLSX}' sudah dibuat.")


if __name__ == "__main__":
    main()
