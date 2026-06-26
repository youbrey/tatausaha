# SIPS – Struktur Modular (hasil pemisahan dari app_v6.py)

File `app_v6.py` (≈2900 baris, 1 file) sudah dipecah menjadi paket
Python terstruktur agar mudah diperbaiki/dikembangkan. **Surat
Perjalanan Dinas dan Surat Undangan kini benar-benar terpisah**, baik
logikanya maupun form UI-nya.

## Cara menjalankan
```
pip install -r requirements.txt
python main.py
```
Letakkan semua file template `.docx` (surat tugas, SPD, pemberitahuan,
daftar hadir, undangan paripurna) di folder yang sama dengan
`main.py`, sesuai nama yang didefinisikan di `config.py`.

## Struktur folder
```
config.py            # SEMUA konstanta: path template, daftar kota, dll
optional_deps.py      # import library opsional (docx2pdf, PyMuPDF, tkcalendar)
database.py            # data personel & riwayat surat (TANPA kode UI)
main.py                 # entry point

utils/                  # fungsi murni, tidak ada UI
  nomor.py              # increment nomor surat / SPD
  geo.py                # deteksi nama & wilayah kota
  tanggal.py             # tanggal, hari, terbilang (Bahasa Indonesia)

docx_helpers/            # manipulasi level-rendah file Word
  formatting.py          # border, lebar kolom, shading sel
  table_ops.py            # isi tabel docx dari data dinamis
  combine.py               # gabung banyak docx jadi satu

letters/                  # >>> LOGIKA SURAT, per jenis surat <<<
  surat_tugas.py           # Surat Tugas DPRD & ASN      -- PERJALANAN DINAS
  pemberitahuan.py         # Surat Pemberitahuan          -- PERJALANAN DINAS
  sppd.py                  # SPD/SPPD DPRD & ASN          -- PERJALANAN DINAS
  daftar_hadir.py          # Daftar Hadir                 -- PERJALANAN DINAS
  undangan.py              # Undangan Paripurna/Biasa     -- TERPISAH TOTAL

ui/                        # tampilan (CustomTkinter), dipecah per bagian
  app.py                   # kelas utama, merakit semua mixin
  sidebar.py                # panel kiri: menu, import DB, navigasi
  personnel_panel.py        # panel kanan: filter & checklist personel
  tujuan_panel.py            # input multi "Kota Tujuan Bertugas"
  perjalanan_form.py         # form Perjalanan Dinas (panel tengah)
  undangan_form.py            # form Surat Undangan (panel tengah, mode lain)
  preview_panel.py             # panel live preview (thread terpisah)
  context_builder.py           # jembatan form -> letters/* (perjalanan dinas)
```

## Kalau ingin memodifikasi sesuatu
- **Ubah path template / nama file / daftar kota** → edit `config.py` saja.
- **Ubah logika Surat Tugas/SPD/Pemberitahuan/Daftar Hadir** → edit file
  terkait di `letters/` (tidak akan memengaruhi Surat Undangan).
- **Ubah logika Surat Undangan** → edit `letters/undangan.py` dan
  `ui/undangan_form.py` saja (tidak memengaruhi Surat Perjalanan Dinas).
- **Ubah tampilan/posisi widget** → edit file mixin di `ui/` yang sesuai
  bagian (sidebar, form, panel personel, preview).
- **Tambah jenis surat baru** → buat file baru di `letters/`, lalu
  panggil dari mixin UI yang sesuai (atau buat mixin form baru).

## Catatan penting
File asli `app_v6.py` sangat besar (≈2900 baris dalam satu class
Tkinter raksasa). Refactor ini menyusun ulang strukturnya menjadi
modul-modul di atas menggunakan pola **mixin** untuk kelas UI (karena
widget Tkinter saling terikat lewat `self`), sambil menarik semua
logika pembuatan dokumen (python-docx/docxtpl) ke `letters/` sebagai
fungsi murni yang tidak bergantung pada widget apa pun -- sehingga
bisa dites atau dipakai ulang tanpa membuka jendela aplikasi.

Karena ukuran file asli yang sangat besar, sebagian detail kecil (misal
pesan placeholder, urutan field tertentu) mungkin perlu disesuaikan
kembali dengan kebutuhan asli Anda -- silakan cek tiap file di
`letters/` dan `ui/` lalu sesuaikan, strukturnya sudah memudahkan untuk
itu.
