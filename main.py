"""
main.py
========
Titik masuk aplikasi SIPS. Jalankan dengan:

    python main.py

Struktur folder:
    config.py            -> semua konstanta (path template, daftar kota, dst.)
    optional_deps.py      -> import library opsional (docx2pdf, PyMuPDF, tkcalendar)
    database.py            -> manajemen data personel & riwayat surat (tanpa UI)

    utils/
        nomor.py            -> penomoran surat
        geo.py              -> deteksi nama & wilayah kota
        tanggal.py          -> tanggal, hari, terbilang (Bahasa Indonesia)

    docx_helpers/
        formatting.py       -> styling tabel/sel docx (border, lebar kolom, dll)
        table_ops.py        -> isi tabel docx dari data dinamis
        combine.py          -> gabung banyak file docx jadi satu

    letters/                -- LOGIKA SURAT (terpisah per jenis surat) --
        surat_tugas.py      -> Surat Tugas DPRD & ASN (perjalanan dinas)
        pemberitahuan.py    -> Surat Pemberitahuan (perjalanan dinas)
        sppd.py              -> SPD/SPPD DPRD & ASN (perjalanan dinas)
        daftar_hadir.py      -> Daftar Hadir (perjalanan dinas)
        undangan.py          -> Surat Undangan Paripurna/Biasa (TERPISAH dari
                                 perjalanan dinas)

    ui/                      -- TAMPILAN (CustomTkinter), per bagian --
        app.py               -> kelas utama SIPSApp, merakit semua mixin
        sidebar.py           -> panel kiri (menu & navigasi)
        personnel_panel.py   -> panel kanan (filter & checklist personel)
        tujuan_panel.py       -> input multi kota tujuan
        perjalanan_form.py    -> form Perjalanan Dinas (panel tengah)
        undangan_form.py      -> form Surat Undangan (panel tengah, mode lain)
        preview_panel.py       -> panel live preview
        context_builder.py     -> jembatan form -> letters/* (perjalanan dinas)
"""

from ui.app import SIPSApp

if __name__ == "__main__":
    app = SIPSApp()
    app.mainloop()
