# Panduan Pengguna — Tanpa Rumus

## 1. Fungsi Tanpa Rumus

Menu **Tanpa Rumus** membuat salinan final dari workbook delivery yang sebelumnya sudah diperbarui melalui Update Schedule. Pada salinan tersebut, formula di worksheet bulanan dihitung lalu dibekukan menjadi nilai (*paste values*), dan worksheet sumber FSS `Date`/`DATE`/`DATA` dihapus.

File sumber tidak diubah. Hasil dari setiap pilihan dibuat sebagai file baru dan dikemas dalam satu ZIP.

## 2. Jenis output

Di halaman Tanpa Rumus tersedia tiga pilihan:

1. **MMKI keseluruhan** — memilih workbook MMKI, tanpa filter Receiving Area.
2. **CKD keseluruhan** — memilih workbook CKD, tanpa filter Receiving Area.
3. **CKD per kelompok receiving** — memilih workbook CKD dan satu atau beberapa kelompok. Semua kode di dalam satu kelompok digabung menjadi satu file output. Aplikasi tidak membuat satu file untuk setiap kode.

Contoh nama hasil dari workbook `20261018 CKD DELIVERY SCH.xlsx`:

- `20261018 CKD DELIVERY SCH (7R01 VIETNAM).xlsx`
- `20261018 CKD DELIVERY SCH (7R02 MALAYSIA & 7R03 L300).xlsx`
- `20261018 CKD DELIVERY SCH (7R06 THAILAND).xlsx`
- `20261018 CKD DELIVERY SCH (7R07 BANGLADESH & 7R08 7R10 RV MALAYSIA).xlsx`
- `20261018 CKD DELIVERY SCH (7R09 7R11 CMC TAIWAN).xlsx`

Label output dapat diubah pada konfigurasi kelompok.

## 3. Kelompok receiving bawaan

Konfigurasi awal yang dibuat otomatis:

| Nama kelompok | Kode | Label output |
|---|---|---|
| Vietnam | `7R01` | `7R01 VIETNAM` |
| Malaysia & L300 | `7R02`, `7R03` | `7R02 MALAYSIA & 7R03 L300` |
| Thailand | `7R06` | `7R06 THAILAND` |
| Bangladesh & RV Malaysia | `7R07`, `7R08`, `7R10` | `7R07 BANGLADESH & 7R08 7R10 RV MALAYSIA` |
| CMC Taiwan | `7R09`, `7R11` | `7R09 7R11 CMC TAIWAN` |

Satu kode receiving hanya boleh terdaftar pada satu kelompok aktif. Jika kode yang sama dimasukkan ke dua kelompok, backend akan menolak penyimpanan.

### Menambah kode atau negara baru

1. Klik **Tambah kelompok** atau edit kelompok yang ada.
2. Isi nama kelompok.
3. Isi label yang ingin muncul pada nama file.
4. Isi satu atau beberapa kode, pisahkan dengan koma. Contoh: `7R12, 7R13`.
5. Klik **Simpan kelompok**.
6. Klik **Periksa workbook** lagi agar kode baru diperiksa terhadap isi sheet Date.

Jika sheet Date berisi kode negara `7Rxx` yang belum dipetakan, proses CKD per kelompok diblokir sampai kode itu dimasukkan ke kelompok. Aplikasi tidak menebak negara dari nomor kode. Kode non-`7R` ditampilkan pada preview, tetapi tidak memerlukan mapping kelompok negara.

Jika ada baris data FSS dengan Receiving Area kosong, proses per kelompok juga diblokir agar order tidak terlewat atau salah dikelompokkan.

Konfigurasi disimpan di SQLite lokal, bukan hanya di state browser. Menghapus kelompok menonaktifkannya sehingga konfigurasi lama tetap dapat ditemukan pada snapshot riwayat.

## 4. Cara memakai Tanpa Rumus

### Langkah 1 — Siapkan workbook sumber

Gunakan workbook yang sudah melalui proses Update Schedule:

- Workbook MMKI untuk output MMKI keseluruhan.
- Workbook CKD untuk CKD keseluruhan dan/atau CKD per kelompok receiving.

Format yang didukung: `.xlsx` dan `.xlsm`. Batas ukuran upload saat ini 50 MB per file.

### Langkah 2 — Pilih file

Pada kartu **Input MMKI**, pilih workbook MMKI jika ingin membuat output MMKI.

Pada kartu **Input CKD**, pilih workbook CKD jika ingin membuat output CKD keseluruhan atau CKD per kelompok.

Kedua sumber tidak wajib dipilih bersamaan. Jika hanya workbook MMKI yang dipilih, opsi CKD otomatis tidak aktif dan hanya output MMKI yang dapat dipilih. Jika hanya workbook CKD yang dipilih, opsi MMKI tidak aktif dan hanya pilihan CKD yang tersedia. Jika kedua sumber dipilih, semua pilihan output dapat digunakan.

### Langkah 3 — Pilih jenis output

Centang satu atau beberapa opsi:

- MMKI keseluruhan.
- CKD keseluruhan.
- CKD per kelompok receiving.

Jika memilih CKD per kelompok, pilih kelompok yang ingin dibuat. Secara awal semua kelompok aktif akan dipilih.

### Langkah 4 — Periksa workbook

Klik **Periksa workbook**.

Selama pemeriksaan, tombol menampilkan indikator berputar dan penghitung waktu. Pemeriksaan MMKI memverifikasi sheet sumber dan nama sheet bulanan; pemeriksaan CKD juga membaca header FSS dan menghitung Receiving Area. Waktu pemeriksaan ditampilkan pada ringkasan hasil.

Aplikasi membaca workbook CKD secara langsung dari struktur file untuk menemukan:

- Worksheet sumber (`Date`, `DATE`, atau `DATA`).
- Baris header dan baris awal data.
- Kolom Receiving Area.
- Kode Receiving Area dan jumlah order per kode.
- Worksheet bulanan yang namanya diawali `YYYYMM`, misalnya `202610`, `202611`, `202612`, atau `202701`.
- Kode `7Rxx` yang belum mempunyai kelompok.
- Baris data yang Receiving Area-nya kosong.

Jika ada kode `7Rxx` yang belum dipetakan, edit atau tambahkan kelompok, lalu jalankan **Periksa workbook** kembali.

### Langkah 5 — Buat file

Setelah pemeriksaan lolos, klik **Buat file Tanpa Rumus**.

Backend membuat setiap output dari salinan workbook sumber yang sama. Output Vietnam tidak digunakan sebagai sumber output Malaysia; output Malaysia juga tidak digunakan sebagai sumber output Thailand.

Setelah selesai, browser mengunduh `FSS_Control_Center_Tanpa_Rumus.zip`.

Indikator proses di bagian atas aplikasi menampilkan fitur yang berjalan dan waktu yang sudah berlalu. Ringkasan hasil menampilkan waktu total proses dan durasi masing-masing output. Riwayat juga menyimpan durasi proses tersebut.

Pengguna boleh berpindah ke menu lain melalui sidebar ketika Build Data FSS, Update Schedule, atau Tanpa Rumus berjalan. Halaman yang sudah dibuka tetap dipertahankan di aplikasi sehingga state, file terpilih, dan request aktif tidak di-reset. Jangan menutup tab browser atau menghentikan terminal backend sampai proses selesai; pemulihan setelah browser ditutup atau halaman di-refresh belum dijamin.

Isi ZIP:

- Workbook hasil untuk setiap pilihan yang berhasil.
- `manifest.json` berisi ringkasan teknis proses dan snapshot kelompok.
- `laporan_tanpa_rumus.csv` berisi ringkasan setiap output.

Jika sebagian output gagal, ZIP berisi output yang berhasil dan laporan menunjukkan output mana yang gagal. Jika tidak ada output yang berhasil, backend mengembalikan pesan error dan tidak mengunduh ZIP final.

## 5. Apa yang dilakukan aplikasi di dalam workbook?

### Output keseluruhan MMKI dan CKD

1. Membuat salinan workbook sumber.
2. Membuka salinan melalui Microsoft Excel desktop.
3. Menghitung workbook.
4. Pada worksheet bulanan yang namanya diawali `YYYYMM`, formula diganti dengan nilai hasil hitung terakhir.
5. Menghapus worksheet Date dari salinan.
6. Menyimpan dan memeriksa file output.

Tidak ada filter Receiving Area pada output keseluruhan.

### Output CKD per kelompok

Untuk setiap kelompok, aplikasi menjalankan proses terpisah dari workbook CKD sumber yang sama:

1. Membuat salinan baru dari CKD asli.
2. Membaca header sheet Date dan menemukan kolom FSS berdasarkan nama header.
3. Jika formula workbook cocok dengan jalur cepat, aplikasi menghitung hanya agregat milik kode Receiving Area kelompok tersebut. Sheet Date tidak perlu ditulis ulang karena akan dihapus dari output.
4. Jika formula workbook tidak cocok dengan jalur cepat, aplikasi memakai jalur fallback: mempertahankan baris kelompok, mengosongkan sel data FSS non-target, tidak menghapus baris worksheet secara fisik, dan tidak mengosongkan helper formula di luar kolom FSS.
5. Memastikan worksheet non-bulanan tidak memiliki formula yang masih merujuk ke Date. Jika masih ada, penghapusan Date diblokir agar formula tidak rusak.
6. Membekukan formula menjadi nilai hanya pada worksheet bulanan yang namanya diawali `YYYYMM` dan memvalidasi bahwa formula bulan sudah tidak tersisa.
7. Menghapus sheet Date, menyimpan, memvalidasi, dan memasukkan output ke ZIP.

Hasil jalur cepat secara logika setara dengan filter per Receiving Area untuk formula `SUMPRODUCT` yang dikenali. Jalur fallback tetap tersedia untuk template yang polanya berbeda, tetapi dapat lebih lambat karena Excel harus menghitung ulang workbook.

### Jalur cepat untuk template CKD yang dikenali

Pada template CKD Delivery yang memakai pola helper `Date!B = S&H&M&T` dan formula bulanan `SUMPRODUCT` standar, aplikasi menggunakan jalur cepat:

1. Membaca baris data FSS dan membuat indeks jumlah `Order Quantity` berdasarkan kombinasi tanggal, delivery period, parts number, serta akhiran Division yang benar-benar terdaftar di formula helper workbook.
2. Membuat agregat sekali dari sumber CKD. Agregat ini dipakai ulang untuk output CKD keseluruhan dan semua kelompok receiving yang dipilih.
3. Mengganti hasil formula `SUMPRODUCT` yang merujuk ke Date dengan angka dari indeks Python, menggunakan penulisan matriks per area data, bukan satu panggilan Excel untuk setiap sel.
4. Menghitung formula lokal yang tersisa pada sheet bulanan, membekukan formula bulanan menjadi nilai, lalu menghapus Date.

Dengan jalur cepat ini, Excel tidak perlu menghitung ulang lebih dari satu juta formula helper di Date atau menjalankan puluhan ribu `SUMPRODUCT` yang masing-masing mencari sampai 90.000 baris. Filter grup dilakukan pada agregat per Receiving Area; file sumber tetap tidak diubah. Jika pola formula tidak cocok atau tidak bisa diverifikasi, aplikasi menggunakan jalur kompatibilitas Excel. Jalur fallback dapat lebih lambat, tetapi tidak memaksakan hasil dari pola yang tidak dikenali.

Untuk contoh workbook CKD yang dipakai saat pengujian, 198.152 formula `SUMPRODUCT` pada empat sheet bulanan menghasilkan nilai agregat yang cocok dengan seluruh 198.152 cached result formula sumber. Preview pada contoh workbook 21 MB selesai sekitar 6,4 detik di lingkungan pengujian. Angka ini bukan jaminan waktu total pada setiap laptop; waktu membuka dan menyimpan workbook tetap bergantung pada CPU, RAM, versi Excel, dan ukuran file.

## 6. Syarat workbook

- Workbook harus `.xlsx` atau `.xlsm`.
- Harus mempunyai satu worksheet sumber bernama `Date`, `DATE`, atau `DATA`.
- Header FSS dikenali dari nama seperti Vendor code, Receiving Area, Delivery date, dan parts number.
- Harus mempunyai setidaknya satu worksheet bulanan yang namanya diawali enam digit `YYYYMM`.
- Worksheet sumber tidak boleh diproteksi untuk proses filter grup.
- Struktur workbook tidak boleh diproteksi jika sheet Date harus dihapus.
- Jika kolom FSS mengandung formula, proses filter grup dihentikan agar formula tersebut tidak ikut dikosongkan.
- Jika formula pada sheet non-bulanan masih merujuk ke Date, penghapusan Date diblokir.

## 7. Persyaratan komputer

Operasi Excel Tanpa Rumus menggunakan xlwings dan Microsoft Excel desktop pada Windows. Backend harus berjalan di komputer yang memiliki Excel desktop terpasang.

Jika Excel tidak tersedia, endpoint akan mengembalikan pesan bahwa Excel desktop diperlukan. Browser sendiri tidak menghitung formula workbook.

## 8. Catatan keamanan

- File asli tidak dibuka untuk ditulis.
- Setiap output dibuat dari sumber asli secara independen.
- Macro `.xlsm` tidak dijalankan selama pemrosesan.
- External links tidak diperbarui otomatis.
- Koneksi data eksternal tidak di-refresh.
- Jika struktur workbook tidak bisa dipertahankan atau divalidasi, proses output tersebut akan gagal dengan laporan.

## 9. Troubleshooting

**Kode 7R baru muncul sebagai belum dipetakan**  
Tambahkan kode itu ke kelompok yang sesuai atau buat kelompok baru, lalu periksa workbook lagi.

**Receiving Area kosong**  
Periksa dan lengkapi Receiving Area pada sheet Date sebelum membuat output per kelompok.

**Sheet Date tidak ditemukan**  
Pastikan workbook schedule yang dipilih adalah workbook hasil Update Schedule dan mempunyai worksheet Date/DATE/DATA.

**Worksheet bulanan tidak ditemukan**  
Pastikan nama tab bulanan diawali `YYYYMM`, misalnya `202610` atau `202610(CKD)`.

**Masih ada formula yang merujuk ke Date**  
Backend sengaja tidak menghapus Date karena formula sheet lain masih membutuhkan sheet tersebut. Periksa workbook dan tentukan apakah formula tersebut memang perlu dipertahankan atau perlu penyesuaian manual.

**Excel atau xlwings tidak tersedia**  
Jalankan backend pada Windows yang sudah memiliki Microsoft Excel desktop dan pasang dependency dari `be/requirements.txt`.

**Sebagian output gagal**  
Buka `manifest.json` dan `laporan_tanpa_rumus.csv` di ZIP untuk melihat output dan alasan kegagalannya. File yang berhasil tetap disediakan, tetapi output yang gagal tidak dimasukkan.
