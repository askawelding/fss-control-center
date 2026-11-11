# FSS Control Center — Web Lokal

Aplikasi web lokal untuk mengolah file FSS dan menyiapkan workbook delivery tanpa rumus. Frontend menggunakan React + TypeScript, backend menggunakan Python + FastAPI, dan operasi workbook delivery dijalankan melalui Microsoft Excel desktop pada Windows dengan xlwings.

## Yang tersedia

- **Build Data FSS**: menggabungkan file Firm dan Forecast menjadi satu master FSS.
- **Update Schedule DATA**: fitur schedule yang sudah tersedia pada backend/front-end proyek.
- **Tanpa Rumus**: membuat output MMKI keseluruhan, CKD keseluruhan, atau CKD per kelompok Receiving Area. Template CKD dengan pola formula yang dikenali menggunakan agregasi Python untuk menghindari kalkulasi ulang `SUMPRODUCT` yang sangat berat di Excel.
- **Konfigurasi receiving**: menambah, mengubah, dan menonaktifkan kelompok kode receiving. Konfigurasi disimpan permanen di SQLite.
- **Riwayat Tanpa Rumus**: mencatat input, kelompok yang dipakai, status, dan laporan setiap proses.
- **Status pekerjaan lintas menu**: halaman fitur tetap tersimpan ketika berpindah melalui sidebar. Proses yang sedang berjalan tidak di-reset hanya karena membuka menu lain; banner global menampilkan waktu berjalan dan status akhir.
- **Pengukuran waktu**: Tanpa Rumus menampilkan durasi pemeriksaan, durasi setiap output, dan waktu total. Build Data FSS dan Update Schedule menampilkan durasi berjalan pada banner global.

## Persyaratan komputer

Untuk menjalankan seluruh fitur Tanpa Rumus yang memproses Excel:

- Windows 10/11.
- Python yang didukung dependency proyek. Contoh perintah di bawah menggunakan Python 3.14 yang sudah tersedia pada komputer pengujian pengguna.
- Node.js LTS dan npm.
- Microsoft Excel desktop terpasang.
- Koneksi internet hanya diperlukan saat pertama kali mengunduh dependency. Pemrosesan workbook berlangsung lokal.

> Catatan: backend dapat dijalankan pada sistem lain untuk pengembangan API, tetapi operasi Tanpa Rumus yang mengendalikan Excel desktop menggunakan xlwings dan ditujukan untuk Windows + Microsoft Excel.

## Menjalankan aplikasi untuk pertama kali (PowerShell)

### 1. Backend Python

Buka PowerShell di folder `be`, lalu jalankan:

```powershell
py -3.14 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

`Set-ExecutionPolicy -Scope Process` hanya berlaku untuk jendela PowerShell saat ini. Jika tidak ingin mengaktifkan environment, jalankan pip dan uvicorn melalui `\.venv\Scripts\python.exe`.

Biarkan jendela backend tetap terbuka selama aplikasi dipakai. Backend berjalan di `http://127.0.0.1:8000`.

### 2. Frontend React

Buka PowerShell kedua di folder `fe`, lalu jalankan:

```powershell
npm.cmd install
npm.cmd run dev
```

Gunakan `npm.cmd` di PowerShell jika eksekusi `npm.ps1` diblokir oleh Execution Policy. Buka alamat yang ditampilkan Vite, biasanya `http://127.0.0.1:5173`.

Untuk menghentikan server, tekan `Ctrl+C` pada masing-masing terminal.

## Menjalankan pengujian

Backend:

```powershell
cd be
.\.venv\Scripts\python.exe -m pytest -q
```

Frontend:

```powershell
cd fe
npm.cmd run lint
npm.cmd run build
```

Untuk detail alur Tanpa Rumus, lihat [`docs/TANPA_RUMUS_USER_GUIDE_ID.md`](docs/TANPA_RUMUS_USER_GUIDE_ID.md).

Catatan: perpindahan antar-menu dalam aplikasi didukung saat proses berjalan. Jangan menutup tab browser, me-refresh halaman, atau menghentikan backend sebelum hasil diterima; penyimpanan dan pemulihan pekerjaan setelah tindakan tersebut belum dijamin.

## Lokasi data lokal

Secara default database SQLite dibuat di `be/data/fss_control_center.db`. Folder kerja, output sementara, ZIP, dan log proses Tanpa Rumus dibuat pada direktori sementara sistem operasi; ZIP diunduh melalui browser. Backend menghapus direktori sementara setelah respons selesai dikirim.

Untuk memindahkan lokasi database, atur environment variable `FSS_DATA_DIR` ke folder lokal yang dapat ditulis sebelum menjalankan backend.

## Privasi

Backend diikat ke `127.0.0.1` secara default. File yang dipilih pengguna dikirim hanya ke backend lokal pada komputer yang sama. Tidak ada proses Tanpa Rumus yang mengirim workbook ke layanan AI atau cloud.
