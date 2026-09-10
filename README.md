# Formatra — Conversion Server

Backend opsional yang memberikan hasil konversi Word/PDF/PPT dengan fidelity
tinggi (mendekati iLovePDF/Smallpdf) dengan cara memakai rendering/parsing
engine sungguhan alih-alih HTML/CSS di browser:

- **Word → PDF** dan **PPT → PDF**: **LibreOffice headless** (tidak berubah).
- **PDF → Word**: **`pdf2docx`** (Python) sebagai engine utama, dengan
  fallback otomatis ke LibreOffice per-request bila `pdf2docx` tidak
  terpasang atau gagal pada file tertentu. `pdf2docx` merekonstruksi halaman
  sebagai pohon layout (blok teks dengan posisi/bounding box asli, grid
  tabel, gambar) lalu membangun ulang `.docx` dari struktur itu — jauh lebih
  dekat ke posisi/spacing/alignment/kolom/tabel PDF asli dibanding filter
  `writer_pdf_import` LibreOffice, yang cenderung me-reflow PDF menjadi
  paragraf Writer biasa.

Jika server ini **tidak** dijalankan/di-deploy, aplikasi frontend tetap
berfungsi seperti biasa dan otomatis memakai konverter bawaan browser
(fidelity lebih rendah, lihat batasan di README utama).

## Kebutuhan

- Node.js 18+ (hanya modul bawaan Node dipakai — **tidak perlu `npm install`**
  untuk server ini).
- LibreOffice terpasang di server dan bisa dipanggil lewat perintah `soffice`.
  - Ubuntu/Debian: `sudo apt-get install -y libreoffice`
  - Disarankan juga install font tambahan agar dokumen dengan font non-default
    (Calibri, Cambria, dst.) tidak "melompat" karena font-substitution:
    `sudo apt-get install -y fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation`
    (Carlito ≈ Calibri, Caladea ≈ Cambria secara metrik, jadi layout tidak
    bergeser meski nama font tetap berbeda).
- Python 3 + paket `pdf2docx`, **hanya** dipakai oleh route PDF → Word:
  ```bash
  pip3 install pdf2docx
  # Debian/Ubuntu terbaru (PEP 668) mungkin perlu:
  pip3 install --break-system-packages pdf2docx
  ```
  Jika langkah ini dilewati, route PDF → Word tetap berfungsi — otomatis
  jatuh ke LibreOffice (perilaku lama) untuk setiap request, hanya dengan
  fidelity yang lebih rendah.

## Menjalankan

```bash
cd server
node index.mjs
```

Server berjalan di `http://localhost:8787` secara default. Endpoint:

- `GET  /api/health` — cek server & LibreOffice siap.
- `POST /api/convert/word-to-pdf` — body: bytes file `.doc/.docx/.rtf/.odt` mentah, header `X-Filename` = nama file asli.
- `POST /api/convert/ppt-to-pdf` — body: bytes file `.ppt/.pptx/.odp`.
- `POST /api/convert/pdf-to-word` — body: bytes file `.pdf`.

Semua endpoint mengembalikan file hasil konversi langsung sebagai response
body (bukan JSON+base64), atau `{"error": "..."}` dengan status 4xx/5xx bila
gagal.

## Variabel lingkungan

| Variabel | Default | Keterangan |
|---|---|---|
| `PORT` | `8787` | Port HTTP server |
| `ALLOWED_ORIGIN` | `*` | Origin yang diizinkan CORS — **set ke domain frontend Anda di production**, jangan biarkan `*` |
| `MAX_FILE_SIZE_MB` | `150` | Batas ukuran file upload (upload di-stream langsung ke disk, tidak dibuffer penuh di memori) |
| `SOFFICE_TIMEOUT_MS` | `90000` | Timeout **dasar** per konversi LibreOffice; otomatis bertambah untuk file besar (lihat `SOFFICE_TIMEOUT_PER_MB_MS`) sebelum proses soffice di-kill paksa |
| `SOFFICE_TIMEOUT_PER_MB_MS` | `2500` | Tambahan waktu timeout (ms) per MB ukuran file, di atas `SOFFICE_TIMEOUT_MS` |
| `SOFFICE_TIMEOUT_MAX_MS` | `600000` | Batas atas timeout LibreOffice walau file sangat besar |
| `MAX_CONCURRENT_JOBS` | `2` | Berapa proses soffice boleh jalan bersamaan (naikkan sesuai jumlah CPU/RAM server) |
| `SOFFICE_BIN` | `soffice` | Path/nama binary LibreOffice bila tidak ada di `PATH` |
| `PDF_TO_WORD_ENGINE` | `pdf2docx` | Set ke `libreoffice` untuk mematikan `pdf2docx` sepenuhnya dan kembali ke perilaku lama (LibreOffice saja) untuk route PDF → Word |
| `PYTHON_BIN` | `python3` | Path/nama binary Python bila tidak ada di `PATH` |
| `PDF2DOCX_TIMEOUT_MS` | `120000` | Timeout **dasar** konversi `pdf2docx`; otomatis bertambah untuk file besar (lihat `PDF2DOCX_TIMEOUT_PER_MB_MS`) sebelum proses di-kill paksa dan request jatuh ke fallback LibreOffice |
| `PDF2DOCX_TIMEOUT_PER_MB_MS` | `4000` | Tambahan waktu timeout (ms) per MB ukuran file, di atas `PDF2DOCX_TIMEOUT_MS` |
| `PDF2DOCX_TIMEOUT_MAX_MS` | `600000` | Batas atas timeout `pdf2docx` walau file sangat besar |

`GET /api/health` sekarang juga melaporkan `sofficeBin`, `pdf2docxEnabled`, dan `maxConcurrentJobs` — berguna untuk memverifikasi cepat bahwa deployment benar-benar memakai LibreOffice/pdf2docx server-side, bukan diam-diam gagal lalu klien jatuh ke fallback browser.

## Menghubungkan ke frontend

Di root proyek frontend, buat file `.env` (lihat `.env.example`):

```
VITE_CONVERT_API_URL=http://localhost:8787
```

Lalu jalankan `npm run dev` / `npm run build` seperti biasa. Jika variabel ini
tidak diisi, frontend tidak akan pernah mencoba menghubungi server — otomatis
memakai mode browser sepenuhnya.

## Deploy dengan Docker

```bash
cd server
docker build -t formatra-server .
docker run -p 8787:8787 -e ALLOWED_ORIGIN=https://app-anda.com formatra-server
```

## Arsitektur & keamanan singkat

- Setiap request mendapat direktori kerja sementara sendiri (`os.tmpdir()`)
  **dan** profil LibreOffice terisolasi sendiri (`-env:UserInstallation=...`).
  Ini bukan cuma soal kerapian: LibreOffice headless memperlakukan direktori
  profil sebagai lock satu-penulis, jadi berbagi profil antar request bisa
  membuat proses saling mengunci/hang atau korup.
- File & folder sementara **selalu** dihapus di blok `finally`, termasuk saat
  konversi gagal/timeout.
- Nama file dari client hanya dipakai untuk membaca ekstensi (dicocokkan ke
  whitelist), tidak pernah dipakai langsung sebagai path filesystem.
- Ukuran body dibatasi saat streaming diterima (tidak menunggu file penuh
  masuk memori dulu baru dicek).
- Jumlah proses `soffice` bersamaan dibatasi lewat semaphore sederhana di
  memori, supaya lonjakan upload tidak membebani host secara berlebihan.

## Keterbatasan yang perlu diketahui

- **Bukan 100% identik** dengan hasil iLovePDF/Smallpdf (mereka memakai engine
  proprietary/berbayar). LibreOffice adalah pendekatan open-source dengan
  fidelity terbaik yang realistis untuk self-hosted.
- **Font**: jika font yang dipakai dokumen sumber tidak terpasang di server,
  LibreOffice mengganti dengan font pengganti (font-substitution) yang
  metriknya mendekati tapi tidak selalu identik — bisa menggeser sedikit
  posisi teks pada dokumen yang sangat padat/presisi.
- **PDF → Word** tetap yang paling sulit: PDF pada dasarnya adalah kumpulan
  instruksi menggambar (bukan model dokumen terstruktur). Engine `pdf2docx`
  yang dipakai sekarang jauh lebih baik dalam mempertahankan posisi, spacing,
  alignment, kolom, tabel, dan gambar dibanding LibreOffice, tapi PDF yang
  sangat kompleks (layout non-standar, PDF hasil scan/gambar tanpa teks asli,
  font yang sangat tidak umum) tetap bisa menghasilkan penyesuaian kecil.
- Endpoint ini melakukan **satu kali proses tanpa progress asli** dari
  LibreOffice (CLI-nya tidak melaporkan progress bertahap) — progress yang
  ditampilkan di UI selama tahap "mengonversi" adalah simulasi, bukan
  progress sebenarnya dari server.
- Tidak ada antrian persisten (in-memory saja) — jika server di-restart saat
  ada job berjalan, job tersebut hilang dan client akan menerima error/timeout
  lalu (untuk Word→PDF/PPT→PDF/PDF→Word) otomatis fallback ke mode browser.
