"""
config.py
=========
Semua konstanta & konfigurasi aplikasi SIPS dikumpulkan di sini.
Jika suatu saat perlu mengubah nama file template, path data, atau
daftar kota, cukup edit file ini -- tidak perlu menyentuh kode logika
atau UI.
"""

# ---------------------------------------------------------------------------
# PATH DATA APLIKASI
# ---------------------------------------------------------------------------
DATA_FILE = "sips_data.json"
HISTORY_FILE = "sips_history.json"
DATABASE_XLSX = "database_dprd_asn.xlsx"

# ---------------------------------------------------------------------------
# TEMPLATE: SURAT TUGAS
# Dipisah antara format "biasa" (<=3 pelaksana) dan format "tabel"
# (>3 pelaksana), masing-masing untuk DPRD dan ASN.
# ---------------------------------------------------------------------------
TEMPLATE_ST_DPRD_BIASA = "surat_tugas_dprd_biasa.docx"
TEMPLATE_ST_DPRD_TABEL = "surat_tugas_dprd_tabel.docx"
TEMPLATE_ST_ASN_BIASA = "surat_tugas_asn_biasa.docx"
TEMPLATE_ST_ASN_TABEL = "surat_tugas_asn_tabel.docx"

# Template surat lain
TEMPLATE_PEMBERITAHUAN = "pemberitahuan_dprd.docx"
TEMPLATE_SPD_DEPAN = "SPD_DPRD.docx"
TEMPLATE_SPD_BELAKANG = "SPD_BELAKANG.docx"
TEMPLATE_DAFTAR_HADIR = "DAFTAR_HADIR_DPRD.docx"

# Template Surat Undangan Paripurna (dicoba berurutan sampai ketemu)
TEMPLATE_UNDANGAN_PARIPURNA_CANDIDATES = [
    "templates/surat_undangan/paripurna/CONTOH_UNDANGAN_PARIPURNA.docx",
    "rapat_paripurna.docx",
]

# ---------------------------------------------------------------------------
# DAFTAR JENIS SURAT UNTUK PANEL LIVE PREVIEW
# Format tuple: (label_tampilan, kode_template, mode_context)
# ---------------------------------------------------------------------------
PREVIEW_TEMPLATES = [
    ("Surat Tugas (DPRD)",     "__surat_tugas_dprd__",    "ctx"),
    ("Surat Tugas (ASN)",      "__surat_tugas_asn__",     "ctx"),
    ("Surat Pemberitahuan",    TEMPLATE_PEMBERITAHUAN,    "ctx"),
    ("SPD DPRD - Halaman Depan",    TEMPLATE_SPD_DEPAN,        "person_dprd"),
    ("SPD DPRD - Halaman Belakang", TEMPLATE_SPD_BELAKANG,     "person_dprd"),
    ("SPD ASN - Halaman Depan",     TEMPLATE_SPD_DEPAN,        "person_asn"),
    ("SPD ASN - Halaman Belakang",  TEMPLATE_SPD_BELAKANG,     "person_asn"),
    ("Daftar Hadir",                "__daftar_hadir__",        "ctx"),
    ("Undangan Paripurna",          "__undangan_paripurna__",  "ctx"),
]

# ---------------------------------------------------------------------------
# KATEGORI PELAKSANA DPRD
# Urutan kategori pelaksana DPRD sesuai slot pada surat pemberitahuan
# (pelaksana_tugas_1..4 / jlh_pelaksana_dprd1..4).
# ---------------------------------------------------------------------------
KATEGORI_DPRD_ORDER = ["Pimpinan DPRD", "Komisi I", "Komisi II", "Komisi III"]

# ---------------------------------------------------------------------------
# DAFTAR WILAYAH
# ---------------------------------------------------------------------------
# Kota/kabupaten di Sulawesi Utara (untuk deteksi transportasi & tanggal)
SULAWESI_UTARA_CITIES = [
    "Bitung", "Manado", "Tomohon", "Kotamobagu",
    "Minahasa", "Minahasa Utara", "Minahasa Selatan", "Minahasa Tenggara",
    "Bolaang Mongondow", "Bolaang Mongondow Utara", "Bolaang Mongondow Selatan",
    "Bolaang Mongondow Timur", "Kepulauan Sangihe", "Kepulauan Talaud", "Kepulauan Sitaro"
]

# Kota di Jabodetabek (untuk penentuan tujuan awal SPD halaman belakang)
JABODETABEK_CITIES = ["Jakarta", "Bekasi", "Tangerang", "Depok", "Bogor"]

# Database kota tujuan untuk autocomplete input "Kota Tujuan Bertugas"
DATABASE_TUJUAN = [
    "Kota Manado", "Kota Bitung", "Kota Tomohon", "Kota Kotamobagu",
    "Kabupaten Minahasa", "Kabupaten Minahasa Utara", "Kabupaten Minahasa Selatan",
    "Kabupaten Minahasa Tenggara", "Kabupaten Bolaang Mongondow",
    "Kabupaten Bolaang Mongondow Utara", "Kabupaten Bolaang Mongondow Selatan",
    "Kabupaten Bolaang Mongondow Timur", "Kabupaten Kepulauan Sangihe",
    "Kabupaten Kepulauan Talaud", "Kabupaten Kepulauan Sitaro",
    "DKI Jakarta", "Kota Surabaya", "Kota Bandung", "Kota Medan",
    "Kota Semarang", "Kota Makassar", "Kota Palembang", "Kota Tangerang",
    "Kota Tangerang Selatan", "Kota Bekasi", "Kota Depok", "Kota Yogyakarta",
    "Kota Surakarta (Solo)", "Kota Balikpapan", "Kota Samarinda",
    "Kota Banjarmasin", "Kota Pontianak", "Kota Denpasar", "Kota Mataram",
    "Kota Kupang", "Kota Ambon", "Kota Jayapura", "Kota Sorong",
    "Kota Palu", "Kota Kendari", "Kota Gorontalo", "Kota Palangkaraya",
    "Kota Tarakan", "Kota Banda Aceh", "Kota Padang", "Kota Pekanbaru",
    "Kota Jambi", "Kota Bengkulu", "Kota Bandar Lampung", "Kota Pangkalpinang",
    "Kota Tanjungpinang"
]

# Daftar penerima surat undangan paripurna (urutan = urutan halaman)
PENERIMA_UNDANGAN_PARIPURNA = [
    "PIMPINAN DAN ANGGOTA DPRD KOTA BITUNG",
    "WALI KOTA BITUNG",
    "WAKIL WALI KOTA BITUNG",
    "SEKRETARIS DAERAH KOTA BITUNG",
    "FORKOPIMDA KOTA BITUNG",
    "FORUM KOORDINASI LINTAS SEKTORAL KEAMANAN, KETENTERAMAN DAN KETERTIBAN MASYARAKAT KOTA BITUNG",
    "1. Para Staf Ahli Wali Kota Bitung\n2. Para Asisten Sekda Kota Bitung\n3. Para Kepala Perangkat Daerah di Lingkungan Pemerintah Kota Bitung\n4. Para Kepala Bagian di Lingkungan Setda Kota Bitung\n5. Dirut Perumda Air Minum Duasudara Bitung\n6. Dirut Perumda Bangun Bitung\n7. Dirut Perumda Pasar Kota Bitung\n8. Dirut Rumah Sakit Pratama Bitung\n9. Para Camat se -- Kota Bitung",
    "TENAGA AHLI FRAKSI DPRD KOTA BITUNG"
]

# ---------------------------------------------------------------------------
# TEMA TAMPILAN (CustomTkinter)
# ---------------------------------------------------------------------------
APPEARANCE_MODE = "light"
COLOR_THEME = "blue"

APP_TITLE = "SIPS - Aplikasi Pembuat Surat Perjalanan Dinas DPRD Bitung"
APP_GEOMETRY = "1900x900"
APP_MIN_SIZE = (1500, 800)
APP_VERSION_LABEL = "v7.0 © DPRD Kota Bitung"
