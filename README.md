# Pivot Studio

Aplikasi Flask lokal: CSV/XLSX Converter, Pivot Table interaktif, dan AI Insights (Groq).

## Menjalankan
```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env        # isi GROQ_API_KEY
python run.py               # buka http://127.0.0.1:5000
pytest                      # menjalankan tes
```

## Definisi agregasi
| Fungsi | Perilaku |
|---|---|
| SUM, AVERAGE, MEDIAN | Field numerik, nilai kosong diabaikan |
| VAR, STD | Population (VAR.P, STD.P) |
| COUNT | Semua baris pada grup |
| COUNT NON-NULL | Hanya nilai tidak kosong |
| COUNT DISTINCT | Nilai unik, kosong diabaikan |
| MODE | Nilai terbanyak, seri dipilih yang terkecil, kosong jika semua unik |

## Catatan
- Tanpa batas ukuran upload (isi `MAX_CONTENT_MB` di `.env` bila ingin membatasi). File besar memakai banyak RAM.
- File upload disimpan sementara di `instance/uploads/<sesi>` dan dibersihkan otomatis setelah `FILE_TTL_HOURS`.
- Tanggal berbentuk teks dibaca dengan urutan hari-bulan-tahun bila tipe kolom diubah ke Tanggal.
- Drag-and-drop memakai API HTML5 bawaan browser (tanpa library). Tiap area juga punya menu Tambah dan tombol geser untuk layar sentuh dan keyboard.
- API key hanya dibaca dari `.env`. Jangan commit `.env`.
