# Generate Compile PKS -TNI AD

Aplikasi web sederhana (Streamlit) untuk menggantikan proses input manual
Compile PKS di kantor pusat. Kanwil upload file Excel PKS mereka masing-masing,
aplikasi otomatis:

1. Membaca 3 sheet (`LAPORAN SOSIALISASI`, `LAPORAN PUBLIKASI`, `LAPORAN SIARAN PERS`)
   dari tiap file kanwil — deteksi baris header otomatis, jadi tahan kalau
   format judul antar kanwil sedikit beda.
2. Menggabungkan semua kanwil ke 1 file Compile, urut sesuai daftar kanwil resmi,
   dengan penomoran ("No") per-blok kanwil seperti format asli.
3. Menghitung ulang sheet `COMPILE DATA JUNI` (rekap jumlah kegiatan per satker/kanwil,
   NIHIL otomatis di-exclude dari hitungan).
4. Update sheet `CHECKLIST` — menandai kanwil mana yang sudah upload.
5. Menjaga 3 PivotTable asli tetap ada, dan menyesuaikan range datanya otomatis
   + set `refreshOnLoad`, jadi begitu file dibuka di Excel, pivot akan
   otomatis coba refresh (kalau belum ter-update sendiri, tinggal klik
   **Data > Refresh All** sekali).

## Isi folder

- `app.py` — kode aplikasi Streamlit
- `template_compile_pks.xlsx` — **template dasar** (struktur & PivotTable asli,
  data dikosongkan). Ini "master file" yang dipakai ulang tiap generate.
- `requirements.txt` — daftar library Python yang dibutuhkan

## Cara jalankan (lokal, buat testing)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Lalu buka `http://localhost:8501` di browser.

## Cara deploy supaya bisa diakses tim/kanwil lain

Beberapa opsi termudah, urut dari paling gampang:

1. **Streamlit Community Cloud** (gratis, paling cepat) — push folder ini ke
   repo GitHub (private), lalu deploy lewat streamlit.io/cloud. Dapat link
   `https://namaapp.streamlit.app` yang bisa dibagikan.
2. **Server internal kantor** — jalankan `streamlit run app.py --server.port 8501`
   di VM/server internal, lalu akses via jaringan intranet kantor.
3. **Docker** — bisa dibungkus jadi container kalau tim IT kamu maunya
   begitu, tinggal minta saya buatkan Dockerfile-nya kalau perlu.

## Hal yang WAJIB disesuaikan ke kondisi kantor kamu

- **`CANONICAL_KANWIL`** di `app.py` — saat ini saya isi berdasarkan sheet
  `CHECKLIST` di file contoh kamu (23 kanwil). Cek ulang & lengkapi kalau ada
  kanwil yang belum masuk daftar.
- **Template `template_compile_pks.xlsx`** — ini saya buat dari file Compile
  Juni 2026 yang kamu kasih (header + PivotTable dipertahankan). Kalau bulan
  depan ada perubahan kolom/format di file pusat, template ini perlu di-update
  mengikuti.
- Aplikasi ini **mengasumsikan** semua kanwil pakai struktur kolom yang sama
  persis dengan contoh file Banten kamu. Kalau ternyata beberapa kanwil punya
  variasi kolom (urutan beda, kolom tambahan, dll), perlu penyesuaian di
  bagian `SHEETS` config.

## Catatan soal PivotTable

PivotTable Excel itu terikat pada "cache" yang tersimpan di dalam file — bukan
formula yang otomatis hitung ulang seperti biasa. Karena keterbatasan Python
(tidak ada library yang bisa membuat PivotTable native dari nol dengan andal),
pendekatan yang dipakai di sini adalah: **pertahankan PivotTable yang sudah ada
di template, lalu setiap generate baru, sesuaikan range sumber datanya**. Ini
membuat pivot tetap akurat tanpa perlu dibuat ulang manual tiap bulan.
