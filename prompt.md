# Prompt: Mengubah Store.pdf Menjadi Excel dengan Python (Gemini Sandbox)

> **Tujuan:** Meminta Gemini untuk menulis dan menjalankan kode Python yang mengekstrak tabel dari `Store.pdf` dan menyimpannya sebagai `Store.xlsx`.
> **Untuk siapa:** Siswa / karyawan yang tidak punya pengalaman pemrograman. Hanya ubah bagian yang ditandai **(BISA DIEDIT)**.
> **Catatan penting:** Versi ini sengaja tidak berisi kode Python maupun nama pustaka Python. Pengguna cukup menjelaskan apa yang diinginkan dengan bahasa sehari-hari.

---

## 1. TUJUAN (Apa yang kita inginkan)

Ambil data tabel toko dari lampiran `Store.pdf` dan simpan sebagai file Excel bernama `Store.xlsx`.

- **File masukan:** `Store.pdf`
- **File keluaran:** `Store.xlsx`
- **Isi yang diharapkan:** satu lembar kerja yang berisi semua baris data dari tabel PDF.
- **Hal penting:**
  - Tabel tersebar di **15 halaman**.
  - Setiap halaman memiliki **6 kolom yang sama**:
    1. `BusinessEntityID`
    2. `Name`
    3. `SalesPersonID`
    4. `Demographics`
    5. `rowguid`
    6. `ModifiedDate`

---

## 2. GAMBARAN DATA (Tabel seperti apa)

```
BusinessEntityID | Name                    | SalesPersonID | Demographics | rowguid | ModifiedDate
292              | Next-Door Bike Store    | 279           | ...          | ...     | ...
294              | Professional Sales and  | 276           | ...          | ...     | ...
```

- Baris pertama di setiap halaman adalah **judul kolom (header)**. Jangan sampai muncul berulang kali di file Excel akhir.
- Kolom `Demographics` berisi teks XML yang panjang.
- Kolom `rowguid` berupa UUID tetapi bisa terpecah menjadi beberapa sel; harus disatukan kembali menjadi satu nilai per baris.

---

## 3. INSTRUKSI UNTUK GEMINI (Langkah-langkah dalam bahasa sehari-hari)

### 3.1 Persiapan

Gemini akan menggunakan lingkungan Python di dalam kotak pasirnya. Mintalah Gemini untuk menginstal alat apa pun yang diperlukan, lalu menulis dan menjalankan skripnya sendiri.

### 3.2 Baca setiap halaman

- Baca seluruh PDF dari halaman 1 sampai halaman 15.
- Temukan tabel di setiap halaman.

### 3.3 Bersihkan data

- Gunakan baris pertama di halaman 1 sebagai judul kolom.
- Lewati baris pertama di setiap halaman berikutnya, sehingga judul kolom hanya muncul sekali.
- Gabungkan semua baris lainnya menjadi satu daftar berurutan.
- Jika satu informasi (terutama `rowguid` atau `Demographics`) terpecah menjadi beberapa sel dalam satu baris, satukan kembali menjadi satu nilai.
- Hapus baris yang benar-benar kosong.

### 3.4 Buat file Excel

- Buat workbook Excel.
- Letakkan tabel gabungan tersebut di lembar pertama.
- Simpan file dengan nama `Store.xlsx`.

### 3.5 Verifikasi dan laporkan

Setelah menyimpan, mintalah Gemini menampilkan:

- Berapa banyak baris data yang berhasil diambil (tidak termasuk judul kolom).
- Pratinjau 5 baris pertama.
- Konfirmasi bahwa `Store.xlsx` telah dibuat.
- Peringatan jika ada nilai yang hilang atau baris yang terlihat kacau.

---

## 4. FILE KELUARAN YANG DIHARAPKAN

| File          | Keterangan                                              |
|---------------|---------------------------------------------------------|
| `Store.xlsx`  | File Excel bersih dengan satu lembar: semua baris toko  |

---

## 5. PANDUAN MENGATASI MASALAH (Edit jika muncul masalah baru)

| Gejala                                              | Penyebab yang mungkin                                           | Perintah tambahan untuk ditanyakan ke Gemini                          |
|-----------------------------------------------------|-------------------------------------------------------------------|------------------------------------------------------------------------|
| Judul kolom muncul di tengah data                   | Baris judul di halaman berikutnya tidak dilewati                  | "Lewati baris pertama di setiap halaman setelah halaman 1"             |
| `rowguid` terlihat pecah menjadi beberapa kolom     | Satu nilai logis terbagi ke beberapa sel                          | "Satukan potongan-potongan rowguid di setiap baris menjadi satu kolom" |
| Ada baris dengan jumlah kolom yang salah            | Ada baris teks tambahan yang tertangkap sebagai tabel             | "Hanya simpan baris yang memiliki 6 nilai tidak kosong"                |
| Baris yang sama muncul lebih dari sekali            | Tabel sama muncul di setiap halaman dan judul kolom tersimpan dua kali | "Hapus baris duplikat, dan pertahankan hanya satu judul kolom di paling atas" |
| Kolom di Excel terlalu sempit                       | Lebar kolom bawaan                                              | "Sesuaikan lebar kolom di Excel setelah menyimpan"                     |

---

## 6. PENINGKATAN OPSIONAL (Hanya jika diperlukan)

- **Uraikan XML di kolom `Demographics`** menjadi kolom-kolom tersendiri seperti `AnnualRevenue`, `BankName`, `BusinessType`, `YearOpened`, `Specialty`, `SquareFeet`, `Brands`, `Internet`, `NumberEmployees`.
- **Format** kolom `ModifiedDate` sebagai kolom tanggal yang benar.
- **Ubah nama lembar kerja** dari `Sheet1` menjadi `Stores`.

---

## 7. PROMPT YANG DIKIRIM KE GEMINI (Salin dari sini ke chat)

```text
Anda memiliki lingkungan Python dalam kotak pasir dan file bernama Store.pdf dilampirkan di chat ini.

PDF ini terdiri dari 15 halaman. Setiap halaman berisi tabel toko yang sama dengan 6 kolom berikut:
BusinessEntityID, Name, SalesPersonID, Demographics, rowguid, ModifiedDate.

Tolong gunakan Python untuk:
1. Menginstal alat apa pun yang belum tersedia.
2. Membaca setiap halaman Store.pdf.
3. Mengambil tabel dari setiap halaman.
4. Menggunakan baris pertama halaman 1 sebagai judul kolom.
5. Melewati baris judul di semua halaman berikutnya.
6. Menyatukan sel yang terpecah di setiap baris, terutama pada kolom rowguid dan Demographics, sehingga setiap baris memiliki tepat 6 nilai yang sesuai dengan judul kolom.
7. Menyimpan tabel gabungan sebagai Store.xlsx.
8. Memberitahu berapa baris data yang berhasil diambil, menunjukkan 5 baris pertama, mengonfirmasi bahwa Store.xlsx telah dibuat, dan memberi peringatan jika ada baris yang kacau.

Terakhir, berikan file Store.xlsx agar saya bisa mengunduhnya.
```

---

## 8. CATATAN UNTUK SISWA

- **Jangan ubah bagian 7 kecuali hasil dari Gemini salah.**
- Jika hasilnya salah, periksa dulu bagian 5 (Panduan Mengatasi Masalah).
- Jika diperlukan perbaikan, tambahkan satu kalimat di akhir prompt pada bagian 7 yang menjelaskan masalahnya, contohnya:
  - `"Nilai rowguid terpecah menjadi dua sel; satukan kembali."`
  - `"Beberapa baris memiliki 7 kolom, seharusnya 6; hapus sel kosong yang berlebih."`
