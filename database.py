"""
database.py
============
Mengelola data personel (DPRD & ASN) dan riwayat nomor surat.
Tidak ada kode UI di sini sama sekali -- kelas ini bisa dites atau
dipakai ulang tanpa perlu membuka jendela aplikasi.
"""

import os
import json

import pandas as pd

from config import DATA_FILE, HISTORY_FILE, DATABASE_XLSX


def normalize_keys(data_list):
    normalized = []
    for d in data_list:
        if isinstance(d, dict):
            normalized.append({str(k).lower().strip(): v for k, v in d.items()})
    return normalized


def read_dprd_asn_from_excel_file(path):
    """Membaca file Excel database (sheet DPRD & sheet ASN) menjadi
    dua list of dict mentah (raw_dprd, raw_asn)."""
    xls = pd.ExcelFile(path)
    sheet_dprd = next((s for s in xls.sheet_names if "dprd" in s.lower() or "anggota" in s.lower()), None)
    sheet_asn = next((s for s in xls.sheet_names if "asn" in s.lower() or "pendamping" in s.lower()), None)

    raw_dprd, raw_asn = [], []
    if sheet_dprd:
        df = pd.read_excel(xls, sheet_name=sheet_dprd)
        df.columns = [str(c).strip() for c in df.columns]
        for _, row in df.iterrows():
            raw_dprd.append({
                "nama": str(row.get("Nama", row.get("NAMA", ""))).strip(),
                "jabatan": str(row.get("Jabatan", row.get("JABATAN", ""))).strip(),
                "kategori": str(row.get("Kategori", row.get("KATEGORI", "Custom"))).strip()
            })
    if sheet_asn:
        df = pd.read_excel(xls, sheet_name=sheet_asn)
        df.columns = [str(c).strip() for c in df.columns]
        for _, row in df.iterrows():
            raw_asn.append({
                "nama": str(row.get("NAMA", row.get("Nama", ""))).strip(),
                "nip": str(row.get("NIP", "-")).strip(),
                "pangkat": str(row.get("PANGKAT/GOLONGAN", row.get("Pangkat", "-"))).strip(),
                "jabatan": str(row.get("JABATAN", row.get("Jabatan", "-"))).strip()
            })
    return raw_dprd, raw_asn


def read_dprd_asn_from_csv_file(path):
    """Membaca file CSV database (heuristik: ada kolom NIP -> ASN,
    selain itu -> DPRD)."""
    df = pd.read_csv(path)
    df.columns = [str(c).strip() for c in df.columns]
    cols_lower = [c.lower() for c in df.columns]
    new_dprd, new_asn = [], []
    if "nip" in cols_lower:
        for _, row in df.iterrows():
            new_asn.append({
                "nama": str(row.get("NAMA", row.get("Nama", ""))).strip(),
                "nip": str(row.get("NIP", "-")).strip(),
                "pangkat": str(row.get("PANGKAT/GOLONGAN", row.get("Pangkat", "-"))).strip(),
                "jabatan": str(row.get("JABATAN", row.get("Jabatan", "-"))).strip()
            })
    else:
        for _, row in df.iterrows():
            new_dprd.append({
                "nama": str(row.get("Nama", row.get("NAMA", ""))).strip(),
                "jabatan": str(row.get("Jabatan", row.get("JABATAN", ""))).strip(),
                "kategori": str(row.get("Kategori", row.get("KATEGORI", "Custom"))).strip()
            })
    return new_dprd, new_asn


DEFAULT_DPRD = [{"nama": "VIVY JEANET GANAP, S.E.", "jabatan": "KETUA", "kategori": "Pimpinan DPRD"}]
DEFAULT_ASN = [{"nama": "Drs. ALBERT M. SARESE, M.Si.", "nip": "19681011 199010 1 002",
                "pangkat": "PEMBINA UTAMA MUDA IV/c", "jabatan": "Sekretaris DPRD"}]


class PersonnelDatabase:
    """Menyimpan & memuat data personel DPRD dan ASN dari Excel/JSON."""

    def __init__(self):
        self.db_dprd = []
        self.db_asn = []

    def load(self):
        if os.path.exists(DATABASE_XLSX):
            try:
                raw_dprd, raw_asn = read_dprd_asn_from_excel_file(DATABASE_XLSX)
                if raw_dprd or raw_asn:
                    if raw_dprd:
                        self.db_dprd = normalize_keys(raw_dprd)
                    if raw_asn:
                        self.db_asn = normalize_keys(raw_asn)
                    self.save()
                    return
            except Exception as e:
                print(f"Gagal membaca {DATABASE_XLSX}: {e}")

        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    data = json.load(f)
                    self.db_dprd = normalize_keys(data.get("dprd", []))
                    self.db_asn = data.get("asn", []).copy()
                    return
            except Exception:
                pass

        self.db_dprd = list(DEFAULT_DPRD)
        self.db_asn = list(DEFAULT_ASN)
        self.save()

    def save(self):
        try:
            with open(DATA_FILE, "w") as f:
                json.dump({"dprd": self.db_dprd, "asn": self.db_asn}, f, indent=4)
        except Exception:
            pass

    def import_from_file(self, file_path):
        """Import database dari file .xlsx/.xls/.csv yang dipilih pengguna.
        Mengembalikan (jumlah_dprd, jumlah_asn). Melempar Exception jika gagal."""
        ext = os.path.splitext(file_path)[1].lower()
        new_dprd, new_asn = [], []
        if ext in (".xlsx", ".xls"):
            new_dprd, new_asn = read_dprd_asn_from_excel_file(file_path)
            if not new_dprd and not new_asn:
                raise ValueError("Tidak ditemukan sheet yang sesuai.")
        else:
            new_dprd, new_asn = read_dprd_asn_from_csv_file(file_path)

        if new_dprd:
            self.db_dprd = normalize_keys(new_dprd)
        if new_asn:
            self.db_asn = normalize_keys(new_asn)
        self.save()
        return len(self.db_dprd), len(self.db_asn)


class HistoryDatabase:
    """Menyimpan & memuat riwayat data surat yang sudah pernah dibuat,
    berdasarkan nomor surat."""

    def __init__(self):
        self.history_data = {}

    def load(self):
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r") as f:
                    self.history_data = json.load(f)
            except Exception:
                self.history_data = {}
        else:
            self.history_data = {}

    def save(self):
        try:
            with open(HISTORY_FILE, "w") as f:
                json.dump(self.history_data, f, indent=4)
        except Exception as e:
            print("Gagal menyimpan riwayat:", e)

    def record(self, nomor_surat, data):
        self.history_data[nomor_surat] = data
        self.save()

    def get(self, nomor_surat):
        return self.history_data.get(nomor_surat)

    def keys(self):
        return list(self.history_data.keys())
