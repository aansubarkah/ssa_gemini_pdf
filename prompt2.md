# Prompt: Mengubah Laporan SKUMPTK PDF Menjadi Excel dengan Python (Gemini Sandbox)

> **Tujuan:** Meminta Gemini untuk menulis dan menjalankan kode Python yang mengekstrak data dari `Laporan_SKUMPTK_dummy.pdf` dan menyimpannya sebagai file Excel.
> **Untuk siapa:** Siswa / karyawan yang tidak punya pengalaman pemrograman. Hanya ubah bagian yang ditandai **(BISA DIEDIT)**.
> **Catatan penting:** Versi ini sengaja tidak berisi kode Python maupun nama pustaka Python. Pengguna cukup menjelaskan apa yang diinginkan dengan bahasa sehari-hari.

---

## 1. TUJUAN (Apa yang kita inginkan)

Ambil data pegawai dan data keluarga dari lampiran `Laporan_SKUMPTK_dummy.pdf`, lalu simpan sebagai file Excel bernama `Laporan_SKUMPTK.xlsx`.

- **File masukan:** `Laporan_SKUMPTK_dummy.pdf`
- **File keluaran:** `Laporan_SKUMPTK.xlsx`
- **Isi yang diharapkan:** satu lembar kerja di mana setiap baris berisi data satu **anggota keluarga** yang tercatat dalam tabel.
- **Hal penting:**
  - Setiap halaman PDF adalah satu surat keterangan untuk **satu pegawai**.
  - **Masing-masing halaman memuat kolom dari pegawai** (seperti nama, nip, pangkat, dan sebagainya) dan **tabel susunan keluarga**.
  - Data pegawai harus ditulis ulang pada setiap baris anggota keluarga.

---

## 2. KOLOM HASIL EXCEL

Semua kolom di Excel harus menggunakan nama di bawah ini:

| No | Nama Kolom         | Sumber Data                                                 |
|----|---------------------|-------------------------------------------------------------|
| 1  | `nama`              | Nama pegawai dari poin 1 halaman PDF                        |
| 2  | `nip`               | NIP pegawai dari poin 1 halaman PDF                         |
| 3  | `tempat`            | Tempat lahir pegawai dari poin 2 halaman PDF                |
| 4  | `tanggal_lahir`     | Tanggal lahir pegawai dari poin 2 halaman PDF               |
| 5  | `jenis_kelamin`     | Jenis kelamin pegawai dari poin 3 halaman PDF               |
| 6  | `agama`             | Agama pegawai dari poin 4 halaman PDF                       |
| 7  | `status_kepegawaian`| Status kepegawaian dari poin 5 halaman PDF                  |
| 8  | `pangkat`           | Pangkat pegawai dari poin 7 halaman PDF                     |
| 9  | `golongan`          | Golongan pegawai dari poin 7 halaman PDF                    |
| 10 | `instansi`          | Instansi / satuan kerja dari poin 8 halaman PDF             |
| 11 | `masa_kerja`        | Masa kerja golongan dari poin 9 halaman PDF                 |
| 12 | `alamat`            | Alamat pegawai dari poin 11 halaman PDF                     |
| 13 | `nama` (tabel)      | Nama anggota keluarga dari tabel susunan keluarga           |
| 14 | `pekerjaan`         | Pekerjaan anggota keluarga dari tabel                       |
| 15 | `tanggal_lahir`     | Tanggal lahir anggota keluarga dari tabel, format YYYY-MM-DD|
| 16 | `tanggal_perkawinan`| Tanggal perkawinan anggota keluarga dari tabel              |
| 17 | `keterangan`        | Keterangan hubungan dari tabel (ISTRI, ANAK, SUAMI, dll.)   |
| 18 | `status`            | Status dari tabel (D, T, atau C)                            |

Ada dua kolom dengan nama `nama` dan dua kolom dengan nama `tanggal_lahir`. Ini sengaja. Yang pertama adalah data pegawai, yang kedua adalah data dari tabel keluarga.

---

## 3. GAMBARAN DATA DALAM PDF

Contoh isi satu halaman PDF:

```
1. Nama / Nip : BAMBANG SUKARNO, SE., M.Si / 197505122000037004
2. Tempat/ Tanggal Lahir : TIMBUKTU / 12-05-1975
3. Jenis Kelamin : Laki-laki
4. Agama : ISLAM
5. Status Kepegawaian : PEGAWAI TETAP (PNS)
6. Jabatan Struk/Fungs : KA.BIDANG PERENCANAAN,...
7. Pangkat/Golongan : (IV/a) Pembina
8. Instansi/satker : DINAS PERINDUSTRIAN DAN PERDAGANGAN
9. Masa Kerja Gol : 18 Tahun -5 Bulan
10. Digaji menurut : Tabel Gapok 01-01-2024 Rp. 1184000
11. Alamat : JL. LEMBEH TANAH MERAH NO. 45 KEL. TANAH MERAH

dengan susunan keluarga sebagai berikut :

No.  Nama                    Pekerjaan  Tanggal Lahir  Tanggal Perkawinan  Keterangan  Status
1    SRI HAYATI, S.Pd        GURU       1977-08-15     1999-06-20          ISTRI       D
2    RIZKI AMANDA SUKARNO    -          2001-03-10     -                   ANAK        D
```

### Cara membaca poin 7 (`Pangkat/Golongan`):

Nilai poin 7 memiliki pola: `(golongan) pangkat`.

Contoh:

| Teks asli                    | Kolom `pangkat`      | Kolom `golongan` |
|------------------------------|----------------------|------------------|
| `(IV/a) Pembina`             | `Pembina`            | `IV/a`           |
| `(III/e) Penata Muda Tk.I`   | `Penata Muda Tk.I`   | `III/e`          |
| `(2) Ahli Pertama`           | `Ahli Pertama`       | `2`              |
| `(III/d) Penata Tk.I`        | `Penata Tk.I`        | `III/d`          |
| `(II/t) Penolong Tk.II`      | `Penolong Tk.II`     | `II/t`           |
| `(1) Teknisi Muda`           | `Teknisi Muda`       | `1`              |

Jadi, isi di dalam kurung menjadi kolom `golongan`, sisanya menjadi kolom `pangkat`.

### Cara membaca poin 1 (`Nama / Nip`):

Nilai poin 1 memiliki pola: `NAMA LENGKAP / NIP`.

Contoh:

| Teks asli                                                       | Kolom `nama`                    | Kolom `nip`             |
|-----------------------------------------------------------------|---------------------------------|-------------------------|
| `BAMBANG SUKARNO, SE., M.Si / 197505122000037004`               | `BAMBANG SUKARNO, SE., M.Si`    | `197505122000037004`    |
| `SRI RAHAYU, SE., M.M / 198001252005018003`                     | `SRI RAHAYU, SE., M.M`          | `198001252005018003`    |

### Cara membaca poin 2 (`Tempat/ Tanggal Lahir`):

Nilai poin 2 memiliki pola: `TEMPAT / TANGGAL`.

Contoh:

| Teks asli                 | Kolom `tempat`     | Kolom `tanggal_lahir` |
|---------------------------|--------------------|-----------------------|
| `TIMBUKTU / 12-05-1975`   | `TIMBUKTU`         | `1975-05-12`          |
| `NEGERI DI AWAN / 25-01-1980` | `NEGERI DI AWAN` | `1980-01-25`          |

Tanggal lahir pegawai di PDF menggunakan format DD-MM-YYYY, tetapi di Excel harus disimpan sebagai format tanggal yang dikenali Excel, yaitu YYYY-MM-DD.

### Catatan tentang tanggal di tabel keluarga:

Tanggal lahir dan tanggal perkawinan di tabel sudah dalam bentuk `YYYY-MM-DD`. Jika kosong, biarkan kosong.

---

## 4. INSTRUKSI UNTUK GEMINI (Langkah-langkah dalam bahasa sehari-hari)

### 4.1 Persiapan

Gemini akan menggunakan lingkungan Python di dalam kotak pasirnya. Mintalah Gemini untuk menginstal alat apa pun yang diperlukan, lalu menulis dan menjalankan skripnya sendiri.

### 4.2 Baca setiap halaman

- Baca seluruh PDF dari halaman pertama sampai halaman terakhir.
- Untuk setiap halaman, ambil:
  - Data pegawai dari poin 1 sampai poin 11.
  - Tabel susunan keluarga yang berada di bagian bawah halaman.

### 4.3 Ambil data pegawai

Dari setiap halaman, ambil nilai berikut:

- `Nama`: teks sebelum tanda `/` di poin 1.
- `Nip`: teks setelah tanda `/` di poin 1.
- `Tempat`: teks sebelum tanda `/` di poin 2.
- `Tanggal Lahir`: teks setelah tanda `/` di poin 2, lalu ubah ke format YYYY-MM-DD.
- `Jenis Kelamin` dari poin 3.
- `Agama` dari poin 4.
- `Status Kepegawaian` dari poin 5.
- `Pangkat`: teks di poin 7 setelah kurung tutup.
- `Golongan`: teks di dalam kurung pada poin 7.
- `Instansi` dari poin 8.
- `Masa Kerja` dari poin 9.
- `Alamat` dari poin 11.

### 4.4 Ambil data tabel keluarga

- Temukan tabel yang dimulai dengan kolom `No.`, `Nama`, `Pekerjaan`, `Tanggal Lahir`, `Tanggal Perkawinan`, `Keterangan`, dan `Status`.
- Lewati baris judul tabel.
- Baca setiap baris anggota keluarga.
- Jika nilai `Tanggal Lahir` atau `Tanggal Perkawinan` kosong, biarkan kosong.

### 4.5 Gabungkan menjadi satu baris per anggota keluarga

Untuk setiap baris di tabel keluarga, buat satu baris di Excel yang berisi:

- Semua data pegawai dari halaman yang sama.
- Lalu data anggota keluarga dari tabel tersebut.

Jadi jika seorang pegawai memiliki 4 anggota keluarga, pegawai tersebut akan memiliki 4 baris di Excel.

### 4.6 Buat file Excel

- Buat workbook Excel.
- Letakkan semua baris hasil di lembar pertama.
- Gunakan urutan kolom dan nama kolom dari tabel di bagian 2.
- Pastikan kolom `tanggal_lahir` (pegawai) dan kolom `tanggal_lahir` (tabel) serta kolom `tanggal_perkawinan` disimpan sebagai **tanggal yang bisa dikenali Excel**, bukan teks biasa.
- Simpan file dengan nama `Laporan_SKUMPTK.xlsx`.

### 4.7 Verifikasi dan laporkan

Setelah menyimpan, mintalah Gemini menampilkan:

- Berapa banyak halaman PDF yang berhasil dibaca.
- Berapa banyak baris total di Excel.
- 5 baris pertama dari Excel sebagai pratinjau.
- Konfirmasi bahwa `Laporan_SKUMPTK.xlsx` telah dibuat.
- Peringatan jika ada halaman atau baris yang tidak bisa dibaca dengan bersih.

---

## 5. FILE KELUARAN YANG DIHARAPKAN

| File                    | Keterangan                                                       |
|-------------------------|------------------------------------------------------------------|
| `Laporan_SKUMPTK.xlsx`  | File Excel bersih dengan satu lembar: satu baris per anggota keluarga |

---

## 6. PANDUAN MENGATASI MASALAH (Edit jika muncul masalah baru)

| Gejala                                                         | Penyebab yang mungkin                                                  | Perintah tambahan untuk ditanyakan ke Gemini                                |
|----------------------------------------------------------------|------------------------------------------------------------------------|------------------------------------------------------------------------------|
| Tanggal lahir muncul sebagai teks, bukan tanggal               | Excel tidak mengenali format tanggal                                   | "Simpan kolom tanggal lahir sebagai tanggal Excel yang sesuai"               |
| Data pegawai muncul hanya pada baris pertama keluarganya       | Data pegawai tidak ditulis ulang untuk setiap anggota keluarga         | "Ulangi data pegawai pada setiap baris anggota keluarga"                     |
| Kolom `pangkat` dan `golongan` tercampur                       | Poin 7 tidak dipisahkan dengan benar                                   | "Pisahkan nilai poin 7: teks dalam kurung ke kolom golongan, sisanya ke pangkat" |
| Ada baris yang kosong di tengah Excel                          | Baris kosong dari tabel ikut tersimpan                                 | "Hapus baris yang seluruhnya kosong"                                         |
| Anggota keluarga tidak muncul                                  | Tabel tidak terdeteksi                                                 | "Cari tabel yang dimulai dengan kolom No., Nama, Pekerjaan, Tanggal Lahir"   |
| Kolom `nama` dan `tanggal_lahir` bertabrakan                   | Excel menolak nama kolom yang sama                                     | "Tetap pertahankan dua kolom bernama nama dan dua kolom bernama tanggal_lahir sesuai urutan yang diminta" |
| Gemini hanya menulis penjelasan / tidak mengeluarkan file | Tombol Code Execution belum aktif, atau Gemini membaca PDF secara visual | "Nyalakan Code Execution di pengaturan obrolan, lalu jalankan kode Python dan berikan file .xlsx. Jangan ketik ulang isi PDF." |

---

## 7. PROMPT YANG DIKIRIM KE GEMINI (Salin dari sini ke chat)

**PRASYARAT — wajib dicek sebelum kirim prompt:**
1. Pastikan tombol **Code Execution** (atau "Advanced" → code execution) dalam obrolan Gemini **aktif / nyala**. Jika mati, Gemini hanya akan menulis penjelasan, bukan menjalankan kode.
2. Jangan menulis instruksi manual "ketik ulang tabel ini" — itu membuat Gemini sekadar menyalin isi PDF sebagai teks.
3. Lampirkan file PDF, lalu salin teks di bawah ini.

```text
Gunakan CODE EXECUTION (jalankan Python). JANGAN membaca isi PDF dengan mata sendiri dan JANGAN mengetik ulang datanya sebagai teks. Anda WAJIB menulis kode Python, MENJALANKANNYA di sandbox, dan mengembalikan file .xlsx. Jika tidak menjalankan kode, jawaban Anda salah.

Berikut file PDF bernama Laporan_SKUMPTK_dummy.pdf. Ekstrak datanya dengan Python dan kirimkan kembali sebagai file Excel bernama Laporan_SKUMPTK.xlsx.

Setiap halaman PDF berisi satu surat keterangan untuk satu pegawai dengan tabel susunan keluarga di bawahnya.

Buat satu baris per anggota keluarga. Setiap baris harus berisi data pegawai dari halaman yang sama, lalu diikuti data anggota keluarga dari tabel.

Kolom hasil Excel (urut dari kiri ke kanan):
nama, nip, tempat, tanggal_lahir, jenis_kelamin, agama, status_kepegawaian, pangkat, golongan, instansi, masa_kerja, alamat, nama, pekerjaan, tanggal_lahir, tanggal_perkawinan, keterangan, status.

Catatan: ada dua kolom bernama "nama" dan dua kolom bernama "tanggal_lahir". Yang pertama berasal dari data pegawai, yang kedua dari tabel keluarga.

Cara membaca data pegawai dari setiap halaman:
- Poin 1 "Nama / Nip": pisahkan dengan tanda "/". Kiri = nama, kanan = nip.
- Poin 2 "Tempat/ Tanggal Lahir": pisahkan dengan tanda "/". Kiri = tempat, kanan = tanggal lahir. Tanggal lahir pegawai di PDF memakai format DD-MM-YYYY; ubah menjadi YYYY-MM-DD dan simpan sebagai tanggal yang dikenali Excel.
- Poin 3 = jenis kelamin.
- Poin 4 = agama.
- Poin 5 = status kepegawaian.
- Poin 7 "Pangkat/Golongan": teks di dalam kurung = golongan; teks setelah kurung tutup = pangkat.
- Poin 8 = instansi.
- Poin 9 = masa kerja.
- Poin 11 = alamat.

Cara membaca tabel keluarga dari setiap halaman:
- Cari tabel dengan kolom: No., Nama, Pekerjaan, Tanggal Lahir, Tanggal Perkawinan, Keterangan, Status.
- Lewati baris judulnya.
- Baca setiap baris anggota keluarga.
- Jika Tanggal Lahir atau Tanggal Perkawinan kosong, biarkan kosong.

Pastikan semua kolom tanggal dihasilkan sebagai tanggal Excel yang benar, bukan teks biasa.

Setelah selesai, tampilkan jumlah halaman yang dibaca, total baris Excel, 5 baris pertama, konfirmasi file Laporan_SKUMPTK.xlsx sudah dibuat, dan beri peringatan jika ada masalah.

Jalankan kodenya di sandbox dengan CODE EXECUTION, lalu lampirkan file Laporan_SKUMPTK.xlsx agar bisa saya unduh. Jangan hanya menuliskan langkah-langkah, potongan kode, atau penjelasan — hasil akhirnya HARUS berupa file .xlsx yang bisa diunduh.
```

---

## 8. CATATAN UNTUK SISWA

- **Jangan ubah bagian 7 kecuali hasil dari Gemini salah.**
- Jika hasilnya salah, periksa dulu bagian 6 (Panduan Mengatasi Masalah).
- Jika diperlukan perbaikan, tambahkan satu kalimat di akhir prompt pada bagian 7, contohnya:
  - `"Tanggal lahir pegawai masih berupa teks, ubah menjadi tanggal Excel."`
  - `"Data pegawai hanya ada di baris pertama keluarganya, ulangi pada semua baris."`
