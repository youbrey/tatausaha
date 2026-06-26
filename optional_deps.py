"""
optional_deps.py
=================
Mengelola import library opsional yang dipakai aplikasi.
Dipisah agar bagian lain dari kode tidak perlu peduli try/except import,
cukup `from optional_deps import HAS_FITZ, fitz, ...`.
"""

# Library untuk Live Preview PDF (Word -> PDF)
try:
    from docx2pdf import convert as convert_to_pdf_word
    HAS_DOCX2PDF = True
except ImportError:
    convert_to_pdf_word = None
    HAS_DOCX2PDF = False

# Library untuk render PDF -> gambar (preview)
try:
    import fitz  # PyMuPDF
    from PIL import Image
    HAS_FITZ = True
except ImportError:
    fitz = None
    Image = None
    HAS_FITZ = False

# Library kalender untuk input tanggal
try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    DateEntry = None
    HAS_TKCALENDAR = False
