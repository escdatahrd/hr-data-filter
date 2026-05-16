import os
import re
import io
import json
import html
import math
import shutil
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# HR PORTAL V40 - FULL-WIDTH SEARCH + DATE DISPLAY FIX
# =========================================================
# Fokus versi ini:
# 1. Menyesuaikan keterbatasan akses developer di server klien.
# 2. Tidak mengasumsikan akses live ke komputer HRD, LAN, atau shared folder lokal.
# 3. Sumber update resmi adalah upload manual Excel Master melalui Streamlit.
# 4. master_database.csv tetap menjadi cache database matang agar dashboard ringan.
# 5. Menjaga kolom ganda seperti NPWP agar tidak hilang.
# 6. Memisahkan NPWP checklist dokumen dari NO NPWP nomor pajak.
# 7. Menarik NIK/NPWP yang salah kamar dari seluruh baris.
# 8. Membersihkan NIK/NPWP yang nyasar di kolom pengalaman kerja.

DATA_FILE = os.getenv("DATA_FILE", "master_database.csv")
DOC_FOLDER = os.getenv("DOC_FOLDER", "dokumen_pelamar")
UPLOAD_ARCHIVE_FOLDER = os.getenv("UPLOAD_ARCHIVE_FOLDER", "uploaded_excels")
DATABASE_BACKUP_FOLDER = os.getenv("DATABASE_BACKUP_FOLDER", "database_backups")
CLEANING_REPORT_FOLDER = os.getenv("CLEANING_REPORT_FOLDER", "cleaning_reports")
EXPECTED_CLIENT_EXCEL_NAME = os.getenv("EXPECTED_CLIENT_EXCEL_NAME", "01 ESC DBTA (1).xlsx")
HR_PORTAL_ADMIN_PASSWORD = os.getenv("HR_PORTAL_ADMIN_PASSWORD", "").strip()
HR_PORTAL_VIEWER_PASSWORD = os.getenv("HR_PORTAL_VIEWER_PASSWORD", "").strip()
HR_PORTAL_PASSWORD = os.getenv("HR_PORTAL_PASSWORD", "").strip()
# Batas baris default untuk rendering tabel dashboard.
# Data penuh tetap ada dan bisa dicari/diexport, tetapi UI tidak perlu menggambar ribuan baris setiap rerun.
DASHBOARD_TABLE_LIMIT = int(os.getenv("DASHBOARD_TABLE_LIMIT", "300"))

SKIP_SHEET_KEYWORDS = ["tat", "kode", "data pendukung", "catatan"]
EMPTY_TOKENS = {"", "nan", "none", "null", "-", "--", "belum ada", "belum ada di db", "tidak ada", "n/a", "na"}
CHECKMARK_TOKENS = {"v", "V", "√", "✓", "ok", "OK", "ada", "ADA", "ya", "YA", "yes", "YES"}

DISPLAY_ORDER = [
    "NO", "NAMA", "STRATA", "KEAHLIAN", "JENIS IJAZAH", "TAHUN LULUS IJAZAH",
    "PROPINSI/KOTA", "PENERBIT SKA/SKK", "BERLAKU SKA", "TGL EXPIRED SKA", "SKA BY",
    "SERT BGH", "SERT BIM", "SIMPAN", "IPTB", "CV", "REF", "IJASAH", "SKA",
    "ASOSIASI", "NPWP", "PAJAK", "KTP", "PENILAIAN", "SUMBER", "PROYEK TERAKHIR",
    "KETERANGAN", "NO. TELP", "EMAIL", "KRPENGALAMAN KERJA (TAHUN)",
    "NO NIK", "NO NPWP", "KATEGORI_ASAL", "CLEANING_NOTES",
]

FORM_COLUMNS = [
    "NAMA", "STRATA", "KEAHLIAN", "JENIS IJAZAH", "TAHUN LULUS IJAZAH", "PROPINSI/KOTA",
    "PENERBIT SKA/SKK", "BERLAKU SKA", "TGL EXPIRED SKA", "SKA BY", "CV", "REF",
    "IJASAH", "NPWP", "PAJAK", "KTP", "SUMBER", "PROYEK TERAKHIR",
    "KETERANGAN", "NO. TELP", "EMAIL", "KRPENGALAMAN KERJA (TAHUN)", "NO NIK", "NO NPWP",
]

# Kolom checklist mengikuti bentuk Excel asli: terisi "v" berarti dokumen/flag tersedia.
DOCUMENT_CHECKLIST_COLUMNS = [
    "CV", "REF", "IJASAH", "SKA", "ASOSIASI", "NPWP", "PAJAK", "KTP",
    "SERT BGH", "SERT BIM", "IPTB",
]

# Label UI dibuat lebih manusiawi, sedangkan nama kolom CSV tetap mengikuti Excel lama.
FIELD_LABELS = {
    "NO": "No.",
    "NAMA": "Nama Personil",
    "STRATA": "Strata",
    "KEAHLIAN": "Keahlian / SKA",
    "JENIS IJAZAH": "Jurusan / Ijazah",
    "TAHUN LULUS IJAZAH": "Tahun Lulus",
    "PROPINSI/KOTA": "Domisili",
    "PENERBIT SKA/SKK": "Penerbit SKA/SKK",
    "BERLAKU SKA": "Masa Berlaku SKA",
    "TGL EXPIRED SKA": "Tanggal Expired SKA",
    "SKA BY": "SKA Oleh",
    "CV": "CV",
    "REF": "Referensi",
    "IJASAH": "Ijazah",
    "SKA": "Dokumen SKA/SKK",
    "ASOSIASI": "Dokumen Asosiasi",
    "NPWP": "Dokumen NPWP",
    "PAJAK": "Dokumen Pajak",
    "KTP": "KTP",
    "PENILAIAN": "Form Penilaian",
    "SERT BGH": "Sertifikat BGH",
    "SERT BIM": "Sertifikat BIM",
    "SIMPAN": "Sudah Diarsipkan",
    "IPTB": "IPTB",
    "SUMBER": "Sumber Data",
    "PROYEK TERAKHIR": "Proyek Terakhir",
    "KETERANGAN": "Keterangan",
    "NO. TELP": "Nomor Telepon",
    "EMAIL": "Email",
    "KRPENGALAMAN KERJA (TAHUN)": "Pengalaman Kerja (Tahun)",
    "NO NIK": "Nomor NIK",
    "NO NPWP": "Nomor NPWP",
    "KATEGORI_ASAL": "Kategori Asal",
    "CLEANING_NOTES": "Catatan Auto-Cleaning",
}

CHECKLIST_GROUPS = [
    ("Berkas Utama", ["CV", "REF", "IJASAH", "SKA", "KTP", "NPWP", "PAJAK"]),
    ("Sertifikasi Tambahan", ["ASOSIASI", "SERT BGH", "SERT BIM", "IPTB"]),
]

CHECKLIST_HELP = {
    "CV": "Centang jika CV tersedia.",
    "REF": "Centang jika dokumen referensi tersedia.",
    "IJASAH": "Centang jika dokumen ijazah tersedia.",
    "SKA": "Centang jika file sertifikat SKA/SKK tersedia.",
    "ASOSIASI": "Centang jika dokumen asosiasi profesi tersedia.",
    "NPWP": "Centang jika dokumen NPWP tersedia. Nomor NPWP diisi di bagian Nomor & Kontak.",
    "PAJAK": "Centang jika dokumen pajak tersedia.",
    "KTP": "Centang jika KTP tersedia.",
    "SERT BGH": "Centang jika sertifikat BGH tersedia.",
    "SERT BIM": "Centang jika sertifikat BIM tersedia.",
    "IPTB": "Centang jika dokumen IPTB tersedia.",
}

def ui_label(column_name):
    return FIELD_LABELS.get(str(column_name), str(column_name))

PRIMARY_FORM_COLUMNS = [
    "NAMA", "STRATA", "JENIS IJAZAH", "TAHUN LULUS IJAZAH", "PROPINSI/KOTA",
    "KEAHLIAN", "PENERBIT SKA/SKK", "BERLAKU SKA", "TGL EXPIRED SKA", "SKA BY",
    "KRPENGALAMAN KERJA (TAHUN)",
]

CONTACT_FORM_COLUMNS = ["NO NIK", "NO NPWP", "NO. TELP", "EMAIL"]
NOTES_FORM_COLUMNS = ["SUMBER", "PROYEK TERAKHIR", "KETERANGAN"]

MONTH_MAP = {
    "januari": "january", "jan": "jan",
    "februari": "february", "feb": "feb",
    "maret": "march", "mar": "mar",
    "april": "april", "apr": "apr",
    "mei": "may",
    "juni": "june", "jun": "jun",
    "juli": "july", "jul": "jul",
    "agustus": "august", "agusutus": "august", "agust": "aug", "agus": "aug", "agu": "aug", "ags": "aug",
    "september": "september", "sep": "sep",
    "oktober": "october", "okto": "oct", "okt": "oct",
    "november": "november", "nov": "nov",
    "desember": "december", "desmber": "december", "des": "dec",
}


def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)


def resolve_app_path(path):
    """Mengubah path relatif menjadi path di working directory aplikasi."""
    if os.path.isabs(path):
        return path
    return os.path.abspath(path)


def file_mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return None


def format_datetime_from_timestamp(timestamp):
    if timestamp is None:
        return "-"
    return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")

def format_datetime_for_filename():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def format_datetime_human():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_checked_value(value):
    """Mengubah isi Excel checklist menjadi boolean untuk UI checkbox."""
    text = str(value).strip()
    if text.lower() in EMPTY_TOKENS:
        return False
    if text.lower() in {"0", "false", "tidak", "no", "x", "none"}:
        return False
    # Di Excel asli, checklist bisa berupa v, V, √, ✓, OK, ADA, atau sel terisi.
    return bool(text)


def checklist_to_cell(is_checked):
    """Format simpan balik ke CSV agar tetap kompatibel dengan checklist Excel asli."""
    return "v" if is_checked else ""


def get_auth_mode():
    """Mode login sederhana. Jika password kosong, aplikasi berjalan tanpa login."""
    if HR_PORTAL_ADMIN_PASSWORD or HR_PORTAL_VIEWER_PASSWORD or HR_PORTAL_PASSWORD:
        return "enabled"
    return "disabled"


def require_login():
    """Login opsional untuk persiapan akses publik.

    - Jika tidak ada env password, user dianggap Admin agar deployment lokal tetap mudah.
    - Jika HR_PORTAL_ADMIN_PASSWORD diisi, password itu mendapat role Admin.
    - Jika HR_PORTAL_VIEWER_PASSWORD diisi, password itu mendapat role Viewer.
    - HR_PORTAL_PASSWORD adalah fallback password sederhana dengan role Admin.
    """
    if get_auth_mode() == "disabled":
        return "Admin"

    if st.session_state.get("authenticated"):
        return st.session_state.get("role", "Viewer")

    st.title("HR Portal")
    st.caption("Masukkan password untuk membuka aplikasi.")
    with st.form("login_form"):
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Masuk", use_container_width=True)

    if submitted:
        if HR_PORTAL_ADMIN_PASSWORD and password == HR_PORTAL_ADMIN_PASSWORD:
            st.session_state["authenticated"] = True
            st.session_state["role"] = "Admin"
            st.rerun()
        elif HR_PORTAL_VIEWER_PASSWORD and password == HR_PORTAL_VIEWER_PASSWORD:
            st.session_state["authenticated"] = True
            st.session_state["role"] = "Viewer"
            st.rerun()
        elif HR_PORTAL_PASSWORD and password == HR_PORTAL_PASSWORD:
            st.session_state["authenticated"] = True
            st.session_state["role"] = "Admin"
            st.rerun()
        else:
            st.error("Password salah.")

    st.stop()


def logout_button():
    if get_auth_mode() == "enabled":
        with st.sidebar:
            st.caption(f"Login sebagai: {st.session_state.get('role', 'Viewer')}")
            if st.button("Logout"):
                st.session_state.pop("authenticated", None)
                st.session_state.pop("role", None)
                st.rerun()


def create_database_backup(label="backup"):
    """Membuat backup database aktif. Tidak error bila database belum ada."""
    data_path = resolve_app_path(DATA_FILE)
    if not os.path.exists(data_path):
        return None
    ensure_folder(resolve_app_path(DATABASE_BACKUP_FOLDER))
    safe_label = re.sub(r"[^A-Za-z0-9_-]", "_", str(label or "backup"))
    backup_name = f"master_database_{format_datetime_for_filename()}_{safe_label}.csv"
    backup_path = os.path.join(resolve_app_path(DATABASE_BACKUP_FOLDER), backup_name)
    shutil.copy2(data_path, backup_path)
    return backup_path


def save_latest_good_database_copy():
    data_path = resolve_app_path(DATA_FILE)
    if not os.path.exists(data_path):
        return None
    ensure_folder(resolve_app_path(DATABASE_BACKUP_FOLDER))
    latest_good = os.path.join(resolve_app_path(DATABASE_BACKUP_FOLDER), "master_database_latest_good.csv")
    shutil.copy2(data_path, latest_good)
    return latest_good


def list_database_backups():
    folder = resolve_app_path(DATABASE_BACKUP_FOLDER)
    if not os.path.exists(folder):
        return []
    backups = []
    for filename in os.listdir(folder):
        if not filename.lower().endswith(".csv"):
            continue
        path = os.path.join(folder, filename)
        if os.path.isfile(path):
            backups.append({
                "filename": filename,
                "path": path,
                "modified_ts": file_mtime(path),
                "modified_at": format_datetime_from_timestamp(file_mtime(path)),
                "size_mb": round(os.path.getsize(path) / (1024 * 1024), 2),
            })
    return sorted(backups, key=lambda x: x["modified_ts"] or 0, reverse=True)


def restore_database_from_backup(backup_path):
    if not backup_path or not os.path.exists(backup_path):
        raise FileNotFoundError("File backup tidak ditemukan.")
    create_database_backup("before_restore")
    shutil.copy2(backup_path, resolve_app_path(DATA_FILE))
    st.cache_data.clear()


def reset_database_file():
    """Reset database matang dari UI, menggantikan SOP manual hapus CSV."""
    data_path = resolve_app_path(DATA_FILE)
    if os.path.exists(data_path):
        create_database_backup("before_reset")
        os.remove(data_path)
    st.cache_data.clear()


def database_status():
    """Status cache CSV yang dipakai dashboard."""
    path = resolve_app_path(DATA_FILE)
    exists = os.path.exists(path)
    size_mb = round(os.path.getsize(path) / (1024 * 1024), 2) if exists else 0
    modified_ts = file_mtime(path)
    return {
        "path": path,
        "exists": exists,
        "size_mb": size_mb,
        "modified_ts": modified_ts,
        "modified_at": format_datetime_from_timestamp(modified_ts),
    }


def upload_archive_status():
    """Status file Excel terakhir yang diupload lewat Streamlit."""
    folder = resolve_app_path(UPLOAD_ARCHIVE_FOLDER)
    latest_path = os.path.join(folder, "latest_uploaded.xlsx")
    exists = os.path.exists(latest_path)
    size_mb = round(os.path.getsize(latest_path) / (1024 * 1024), 2) if exists else 0
    modified_ts = file_mtime(latest_path)
    return {
        "folder": folder,
        "latest_path": latest_path,
        "exists": exists,
        "size_mb": size_mb,
        "modified_ts": modified_ts,
        "modified_at": format_datetime_from_timestamp(modified_ts),
    }


def safe_filename(filename):
    name = os.path.basename(str(filename or "uploaded.xlsx"))
    name = re.sub(r"[^A-Za-z0-9_.() -]", "_", name).strip()
    return name or "uploaded.xlsx"


def save_uploaded_excel_snapshot(uploaded_file):
    """Menyimpan salinan upload untuk audit/reprocess, lalu mengembalikan path latest."""
    ensure_folder(resolve_app_path(UPLOAD_ARCHIVE_FOLDER))
    original_name = safe_filename(getattr(uploaded_file, "name", "uploaded.xlsx"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = resolve_app_path(UPLOAD_ARCHIVE_FOLDER)
    latest_path = os.path.join(folder, "latest_uploaded.xlsx")
    archive_path = os.path.join(folder, f"{timestamp}_{original_name}")

    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    content = uploaded_file.read()

    with open(latest_path, "wb") as latest_file:
        latest_file.write(content)
    with open(archive_path, "wb") as archive_file:
        archive_file.write(content)

    return latest_path, archive_path, original_name

def dataframe_to_excel_bytes(df, sheet_name="Data"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


def generate_cleaning_report(df, stats):
    """Membuat laporan cleaning yang bisa diunduh HRD tanpa membaca log teknis."""
    ensure_folder(resolve_app_path(CLEANING_REPORT_FOLDER))
    timestamp = format_datetime_for_filename()
    report_path = os.path.join(resolve_app_path(CLEANING_REPORT_FOLDER), f"cleaning_report_{timestamp}.xlsx")

    df_report = normalize_final_dataframe(df).copy() if not df.empty else pd.DataFrame(columns=DISPLAY_ORDER)
    if "EXPIRED_SKA" not in df_report.columns:
        df_report["EXPIRED_SKA"] = df_report.apply(derive_expired_ska, axis=1) if not df_report.empty else []
    if not df_report.empty:
        df_report["EXPIRED_DATE_OBJ"] = df_report["EXPIRED_SKA"].apply(parse_expired_date)

    auto_notes = pd.DataFrame()
    if "CLEANING_NOTES" in df_report.columns:
        auto_notes = df_report[df_report["CLEANING_NOTES"].astype(str).str.strip().ne("")]

    tanggal_gagal = pd.DataFrame()
    if "EXPIRED_SKA" in df_report.columns and "EXPIRED_DATE_OBJ" in df_report.columns:
        tanggal_gagal = df_report[df_report["EXPIRED_SKA"].astype(str).str.strip().ne("") & df_report["EXPIRED_DATE_OBJ"].isna()]

    summary_rows = [
        ["Waktu proses", stats.get("processed_at", format_datetime_human())],
        ["Nama file upload", stats.get("uploaded_filename", "-")],
        ["Total record", stats.get("total_records", len(df_report))],
        ["Personil unik", stats.get("unique_persons", df_report["NAMA"].nunique() if "NAMA" in df_report.columns else 0)],
        ["NIK terisi", stats.get("rows_with_nik", 0)],
        ["NPWP terisi", stats.get("rows_with_npwp", 0)],
        ["Baris auto-fix", stats.get("rows_with_notes", 0)],
        ["Tanggal gagal dibaca", int(len(tanggal_gagal))],
        ["Sheet diproses", ", ".join(stats.get("processed_sheets", []))],
        ["Sheet dilewati", ", ".join(stats.get("skipped_sheets", []))],
        ["Backup sebelum upload", stats.get("backup_before_upload", "-") or "-"],
    ]
    summary_df = pd.DataFrame(summary_rows, columns=["Item", "Nilai"])

    cols_notes = [c for c in ["NAMA", "KEAHLIAN", "JENIS IJAZAH", "NO NIK", "NO NPWP", "KRPENGALAMAN KERJA (TAHUN)", "KATEGORI_ASAL", "CLEANING_NOTES"] if c in auto_notes.columns]
    cols_dates = [c for c in ["NAMA", "KEAHLIAN", "BERLAKU SKA", "TGL EXPIRED SKA", "EXPIRED_SKA", "KATEGORI_ASAL"] if c in tanggal_gagal.columns]

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Ringkasan")
        (auto_notes[cols_notes] if cols_notes else auto_notes).to_excel(writer, index=False, sheet_name="Auto Fix")
        (tanggal_gagal[cols_dates] if cols_dates else tanggal_gagal).to_excel(writer, index=False, sheet_name="Tanggal Gagal")
        pd.DataFrame({"Sheet Diproses": stats.get("processed_sheets", [])}).to_excel(writer, index=False, sheet_name="Sheet Diproses")
        pd.DataFrame({"Sheet Dilewati": stats.get("skipped_sheets", [])}).to_excel(writer, index=False, sheet_name="Sheet Dilewati")

    return report_path


def read_report_bytes(report_path):
    if report_path and os.path.exists(report_path):
        with open(report_path, "rb") as report_file:
            return report_file.read()
    return None


def process_uploaded_excel_workflow(uploaded_file):
    """Workflow handover-safe: simpan upload, backup DB lama, proses, buat report, simpan latest good."""
    latest_path, archive_path, original_name = save_uploaded_excel_snapshot(uploaded_file)
    backup_path = create_database_backup("before_upload")

    success, total, stats = proses_excel_baru(latest_path)
    stats = dict(stats or {})
    stats.update({
        "uploaded_filename": original_name,
        "latest_upload_path": latest_path,
        "archive_upload_path": archive_path,
        "backup_before_upload": backup_path,
        "processed_at": format_datetime_human(),
    })

    if success:
        df_after = load_data_uncached()
        stats["total_records"] = int(len(df_after))
        stats["unique_persons"] = int(df_after["NAMA"].nunique()) if "NAMA" in df_after.columns else 0
        report_path = generate_cleaning_report(df_after, stats)
        latest_good_path = save_latest_good_database_copy()
        stats["report_path"] = report_path
        stats["latest_good_database"] = latest_good_path
        st.cache_data.clear()

    return success, total, stats


def latest_cleaning_report_status():
    folder = resolve_app_path(CLEANING_REPORT_FOLDER)
    if not os.path.exists(folder):
        return None
    reports = []
    for filename in os.listdir(folder):
        if filename.lower().endswith(".xlsx"):
            path = os.path.join(folder, filename)
            reports.append((file_mtime(path) or 0, path, filename))
    if not reports:
        return None
    reports.sort(reverse=True)
    return {"path": reports[0][1], "filename": reports[0][2], "modified_at": format_datetime_from_timestamp(reports[0][0])}

def normalize_text(value):
    """Membersihkan nilai sel tanpa merusak format penting seperti NPWP bertitik/strip."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\u00a0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    if text.lower() in EMPTY_TOKENS:
        return ""
    return text


def normalize_header(value):
    text = normalize_text(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = text.upper().replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if text in {"JENIS IJASAH", "JENIS IJASA"}:
        text = "JENIS IJAZAH"
    if text in {"IJASA DAN KELULUSAN", "IJASAH DAN KELULUSAN"}:
        text = "IJAZAH DAN KELULUSAN"
    text = text.replace("SERTIFIKAT", "SERT")
    text = text.replace("PROPINSI/ KOTA", "PROPINSI/KOTA")
    text = text.replace("PROPINSI / KOTA", "PROPINSI/KOTA")
    text = text.replace("DOMISILI (KOTA/PROVINSI)", "PROPINSI/KOTA")
    text = text.replace("BERLAKU SKA/SKK", "BERLAKU SKA")
    text = text.replace("PENERBIT SKA", "PENERBIT SKA/SKK") if text == "PENERBIT SKA" else text
    text = text.replace("NO.", "NO") if text == "NO." else text
    return text


def make_unique_columns(headers):
    """Membuat nama kolom unik tanpa membuang kolom ganda."""
    seen = {}
    result = []
    for i, header in enumerate(headers):
        base = normalize_header(header) or f"UNNAMED_{i}"
        count = seen.get(base, 0) + 1
        seen[base] = count
        result.append(base if count == 1 else f"{base}__{count}")
    return result


def base_header(column_name):
    return re.sub(r"__\d+$", "", str(column_name))


def digits_only(value):
    return re.sub(r"\D", "", normalize_text(value))


def is_empty(value):
    return normalize_text(value).lower() in EMPTY_TOKENS


def is_checkmark(value):
    return normalize_text(value) in CHECKMARK_TOKENS


def valid_nik_digits(digits):
    """Validasi ringan NIK: 16 digit, kode provinsi masuk akal, tanggal lahir masuk akal."""
    if not re.fullmatch(r"\d{16}", digits):
        return False
    if digits.startswith("00"):
        return False
    try:
        province = int(digits[0:2])
        day = int(digits[6:8])
        month = int(digits[8:10])
    except ValueError:
        return False
    if not (11 <= province <= 99):
        return False
    if day > 40:
        day -= 40
    if not (1 <= day <= 31):
        return False
    if not (1 <= month <= 12):
        return False
    return True


def is_probable_nik(value):
    return valid_nik_digits(digits_only(value))


def is_probable_npwp(value, header_context=""):
    text = normalize_text(value)
    digits = digits_only(text)
    header_context = header_context.upper()
    if not digits:
        return False
    if is_probable_nik(text):
        # 16 digit yang valid sebagai NIK jangan dipaksa menjadi NPWP.
        return False
    if len(digits) == 15 and ("." in text or "-" in text or "NPWP" in header_context):
        return True
    if len(digits) == 16 and "NPWP" in header_context:
        return True
    return False


def is_date_like(value):
    text = normalize_text(value).lower()
    if not text:
        return False
    if any(month in text for month in MONTH_MAP):
        return True
    if re.search(r"\b\d{4}-\d{1,2}-\d{1,2}\b", text):
        return True
    if re.search(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", text):
        return True
    if "00:00:00" in text:
        return True
    return False



def is_probable_email(value):
    text = normalize_text(value)
    if not text or is_checkmark(text):
        return False
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", text))


def is_probable_phone(value):
    text = normalize_text(value)
    if not text or is_checkmark(text):
        return False
    digits = digits_only(text)
    if len(digits) < 9 or len(digits) > 15:
        return False
    # Hindari NIK/NPWP yang juga 15-16 digit.
    if is_probable_nik(text) or is_probable_npwp(text, "NO NPWP"):
        return False
    # Nomor telepon lokal/HP Indonesia biasanya diawali 0, 62, +62, atau memiliki separator.
    if text.startswith("+62") or digits.startswith("62") or digits.startswith("08") or digits.startswith("8"):
        return True
    return bool(re.search(r"[+()\-\s]", text) and len(digits) >= 9)


def is_source_like(value):
    """Nilai sumber biasanya nama kanal/vendor, bukan email, bukan nomor, bukan checklist."""
    text = normalize_text(value)
    if not text or is_checkmark(text):
        return False
    if is_probable_email(text) or is_probable_phone(text):
        return False
    if is_probable_nik(text) or is_probable_npwp(text, "NO NPWP"):
        return False
    if any(ch.isalpha() for ch in text) and len(text) >= 3:
        return True
    return False


def repair_text_contact_fields(out, raw_row=None, column_meta=None):
    """Memperbaiki kasus salah-kamar umum di file HRD.

    Contoh nyata sheet NON TKNK:
    - Sumber berisi tanda centang, sedangkan sumber asli 'Maganghub' ditaruh di kolom Email.
    - Nomor telepon ditaruh di kolom kosong/UNNAMED setelah Email.
    - Keterangan kadang hanya berisi tanda centang sehingga tidak bermakna sebagai keterangan.
    """
    notes = []

    email = normalize_text(out.get("EMAIL", ""))
    sumber = normalize_text(out.get("SUMBER", ""))
    ket = normalize_text(out.get("KETERANGAN", ""))
    telp = normalize_text(out.get("NO. TELP", ""))

    # SUMBER dan KETERANGAN bukan checklist. Jika hanya tanda centang, kosongkan.
    if is_checkmark(sumber):
        out["SUMBER"] = ""
        notes.append("Tanda centang di kolom sumber dikosongkan")
        sumber = ""
    if is_checkmark(ket):
        out["KETERANGAN"] = ""
        notes.append("Tanda centang di kolom keterangan dikosongkan")

    # Email yang tidak mengandung @ dan terlihat seperti nama sumber dipindahkan ke SUMBER.
    if email and not is_probable_email(email):
        if is_probable_phone(email):
            if not telp:
                out["NO. TELP"] = email
                notes.append("Nomor telepon dipindahkan dari kolom email")
            out["EMAIL"] = ""
        elif is_source_like(email):
            if not out.get("SUMBER"):
                out["SUMBER"] = email
                notes.append("Sumber data dipindahkan dari kolom email")
            out["EMAIL"] = ""
        elif is_checkmark(email):
            out["EMAIL"] = ""
            notes.append("Tanda centang di kolom email dikosongkan")

    # Jika NO. TELP kosong, cari nomor telepon yang nyasar di kolom kosong/kolom lain.
    if raw_row is not None and column_meta is not None and not normalize_text(out.get("NO. TELP", "")):
        phone_candidates = []
        for meta in column_meta:
            value = normalize_text(raw_row.get(meta["col"], ""))
            if not value:
                continue
            if is_probable_phone(value):
                base = meta.get("base", "")
                canonical = meta.get("canonical", "")
                priority = 0 if "TELP" in canonical.upper() or "TELP" in base.upper() else 1
                phone_candidates.append((priority, meta.get("idx", 9999), value))
        if phone_candidates:
            phone_candidates.sort(key=lambda x: (x[0], x[1]))
            out["NO. TELP"] = phone_candidates[0][2]
            notes.append("Nomor telepon ditemukan dari kolom lain")

    old_notes = normalize_text(out.get("CLEANING_NOTES", ""))
    if notes:
        out["CLEANING_NOTES"] = "; ".join([x for x in [old_notes] + notes if x])
    return out

def canonical_header(column_name, index, all_base_headers):
    """Menentukan nama final kolom berdasarkan header dan posisi.

    Poin penting: header 'NPWP' ada dua makna.
    - NPWP sebelum PAJAK/KTP = checklist dokumen.
    - NPWP di area akhir setelah EMAIL/KR/NIK = nomor NPWP.
    """
    h = base_header(column_name)
    compact = h.replace(" ", "")
    email_idx = next((i for i, x in enumerate(all_base_headers) if x == "EMAIL"), 10**6)
    pajak_idx = next((i for i, x in enumerate(all_base_headers) if x == "PAJAK"), 10**6)

    direct_map = {
        "NO": "NO",
        "NOMOR": "NO",
        "INDEX": "NO",
        "NO URUT": "NO",
        "NAMA": "NAMA",
        "STRATA": "STRATA",
        "KEAHLIAN": "KEAHLIAN",
        "SKA/SKK AKTIF YANG DIMILIKI": "KEAHLIAN",
        "SKA/SKK YANG DIMILIKI": "KEAHLIAN",
        "JENIS IJAZAH": "JENIS IJAZAH",
        "IJAZAH DAN KELULUSAN": "JENIS IJAZAH",
        "IJASA DAN KELULUSAN": "JENIS IJAZAH",
        "PENDIDIKAN": "JENIS IJAZAH",
        "TAHUN LULUS IJAZAH": "TAHUN LULUS IJAZAH",
        "PROPINSI/KOTA": "PROPINSI/KOTA",
        "PROPINSI": "PROPINSI/KOTA",
        "DOMISILI": "PROPINSI/KOTA",
        "KOTA/PROVINSI": "PROPINSI/KOTA",
        "PENERBIT SKA/SKK": "PENERBIT SKA/SKK",
        "BERLAKU SKA": "BERLAKU SKA",
        "TGL EXPIRED SKA": "TGL EXPIRED SKA",
        "TGL EXPIRED SKA/SKK": "TGL EXPIRED SKA",
        "SKA BY": "SKA BY",
        "SERT BGH": "SERT BGH",
        "SERT BIM": "SERT BIM",
        "SIMPAN": "SIMPAN",
        "IPTB": "IPTB",
        "CV": "CV",
        "REF": "REF",
        "IJAZAH FILE": "IJASAH",
        "SKA": "SKA",
        "ASOSIASI": "ASOSIASI",
        "PAJAK": "PAJAK",
        "KTP": "KTP",
        "PENILAIAN": "PENILAIAN",
        "SUMBER": "SUMBER",
        "PROYEK TERAKHIR": "PROYEK TERAKHIR",
        "KETERANGAN": "KETERANGAN",
        "CATATAN": "KETERANGAN",
        "KET": "KETERANGAN",
        "NO TELP": "NO. TELP",
        "NO. TELP": "NO. TELP",
        "TELP": "NO. TELP",
        "TELEPON": "NO. TELP",
        "EMAIL": "EMAIL",
        "STATUS KERJA": "STATUS_KERJA",
    }

    if h in direct_map:
        return direct_map[h]

    if h in {"IJAZAH", "IJASAH"}:
        # Header tunggal IJAZAH/IJASAH pada file klien adalah checklist dokumen.
        # Pendidikan asli biasanya bernama JENIS IJAZAH atau PENDIDIKAN.
        return "IJASAH"

    if "PENGALAMAN" in h:
        return "KRPENGALAMAN KERJA (TAHUN)__RAW"
    if h == "KR":
        return "KR"
    if compact in {"NONIK", "NIK"}:
        return "NO NIK__RAW"
    if compact in {"NONPWP"}:
        return "NO NPWP__RAW"
    if h == "NPWP":
        if index > email_idx or index > pajak_idx + 4:
            return "NO NPWP__RAW"
        return "NPWP"
    return h


def column_priority_for_nik(canonical, base, index):
    b = base.upper()
    c = canonical.upper()
    if "NIK" in c or "NIK" in b:
        return 0
    if "NPWP" in c or "NPWP" in b:
        return 1
    if "PENGALAMAN" in c or "PENGALAMAN" in b:
        return 2
    return 3


def column_priority_for_npwp(canonical, base, index):
    b = base.upper()
    c = canonical.upper()
    if "NO NPWP" in c or "NO NPWP" in b:
        return 0
    if "NPWP" in c or "NPWP" in b:
        return 1
    return 3


def clean_experience(value):
    text = normalize_text(value)
    if not text:
        return ""
    if is_probable_nik(text) or is_probable_npwp(text, "NO NPWP") or is_date_like(text):
        return ""
    if re.search(r"@", text):
        return ""
    digits = digits_only(text)
    if len(digits) >= 8:
        return ""
    low = text.lower().replace(",", ".")
    if re.fullmatch(r"\d{1,2}(\.\d{1,2})?", low):
        try:
            if 0 <= float(low) <= 60:
                return text
        except ValueError:
            return ""
    if re.fullmatch(r"\d{1,2}\s*[-/]\s*\d{1,2}\s*(tahun|thn)?", low):
        return text
    if re.fullmatch(r"\d{1,2}\s*(tahun|thn)", low):
        return text
    return ""


def pick_first_nonempty(values):
    for value in values:
        text = normalize_text(value)
        if text:
            return text
    return ""


def append_note(notes, message):
    if message and message not in notes:
        notes.append(message)


def auto_align_row(row, column_meta):
    """Membaca satu baris mentah, lalu mengembalikan nilai final + catatan cleaning."""
    notes = []
    nik_candidates = []
    npwp_candidates = []
    domisili_candidates = []
    pengalaman_candidates = []

    for meta in column_meta:
        col = meta["col"]
        base = meta["base"]
        canonical = meta["canonical"]
        idx = meta["idx"]
        value = normalize_text(row.get(col, ""))
        if not value:
            continue

        if is_probable_nik(value):
            nik_candidates.append((column_priority_for_nik(canonical, base, idx), idx, digits_only(value), canonical, base))
            if "PENGALAMAN" in canonical.upper() or "PENGALAMAN" in base.upper():
                append_note(notes, "NIK dipindahkan dari kolom pengalaman")
            elif "NPWP" in canonical.upper() or "NPWP" in base.upper():
                append_note(notes, "NIK ditemukan di area NPWP lalu dipindahkan")
            continue

        if is_probable_npwp(value, f"{canonical} {base}"):
            npwp_candidates.append((column_priority_for_npwp(canonical, base, idx), idx, value, canonical, base))
            if "NPWP" not in canonical.upper() and "NPWP" not in base.upper():
                append_note(notes, "NPWP dipindahkan dari kolom lain")
            continue

        if "NIK" in canonical.upper() or "NIK" in base.upper() or "NO NPWP" in canonical.upper():
            if len(value) > 3 and not any(ch.isdigit() for ch in value) and not is_checkmark(value):
                domisili_candidates.append(value)
                append_note(notes, "Teks lokasi dipindahkan dari area NIK/NPWP")

        if "PENGALAMAN" in canonical.upper():
            clean_exp = clean_experience(value)
            if clean_exp:
                pengalaman_candidates.append(clean_exp)
            elif value:
                append_note(notes, "Nilai pengalaman tidak valid dibersihkan")

    nik_candidates.sort(key=lambda x: (x[0], x[1]))
    npwp_candidates.sort(key=lambda x: (x[0], x[1]))

    selected_nik = ""
    if nik_candidates:
        selected_nik = nik_candidates[0][2]
        unique_nik = sorted(set(x[2] for x in nik_candidates))
        if len(unique_nik) > 1:
            append_note(notes, "Ada lebih dari satu kandidat NIK")

    selected_npwp = ""
    for _, _, candidate, _, _ in npwp_candidates:
        if digits_only(candidate) != selected_nik:
            selected_npwp = candidate
            break

    unique_npwp = sorted(set(digits_only(x[2]) for x in npwp_candidates if digits_only(x[2]) != selected_nik))
    if len(unique_npwp) > 1:
        append_note(notes, "Ada lebih dari satu kandidat NPWP")

    selected_pengalaman = pick_first_nonempty(pengalaman_candidates)
    selected_domisili_from_wrong_col = pick_first_nonempty(domisili_candidates)

    return {
        "NO NIK": selected_nik,
        "NO NPWP": selected_npwp,
        "KRPENGALAMAN KERJA (TAHUN)": selected_pengalaman,
        "DOMISILI_FROM_WRONG_COL": selected_domisili_from_wrong_col,
        "CLEANING_NOTES": "; ".join(notes),
    }


def derive_expired_ska(row_dict):
    expired = normalize_text(row_dict.get("TGL EXPIRED SKA", ""))
    berlaku = normalize_text(row_dict.get("BERLAKU SKA", ""))
    return expired or berlaku


def clean_dataframe_from_sheet(df_raw, sheet_name):
    header_row_index = -1
    for i, row in df_raw.iterrows():
        normalized_values = [normalize_header(v) for v in row.values]
        if "NAMA" in normalized_values:
            header_row_index = i
            break

    if header_row_index == -1:
        return pd.DataFrame()

    headers_raw = df_raw.iloc[header_row_index].tolist()
    unique_columns = make_unique_columns(headers_raw)
    df = df_raw.iloc[header_row_index + 1:].reset_index(drop=True).copy()
    df.columns = unique_columns

    base_headers = [base_header(c) for c in df.columns]
    column_meta = []
    for idx, col in enumerate(df.columns):
        b = base_header(col)
        column_meta.append({
            "idx": idx,
            "col": col,
            "base": b,
            "canonical": canonical_header(col, idx, base_headers),
        })

    rows = []
    for _, row in df.iterrows():
        aligned = auto_align_row(row, column_meta)
        out = {col: "" for col in DISPLAY_ORDER}
        out["KATEGORI_ASAL"] = sheet_name

        for meta in column_meta:
            canonical = meta["canonical"]
            if canonical.endswith("__RAW"):
                continue
            if canonical not in out:
                continue
            value = normalize_text(row.get(meta["col"], ""))
            if not value:
                continue
            if canonical in {"NO NIK", "NO NPWP", "KRPENGALAMAN KERJA (TAHUN)"}:
                continue
            if canonical == "NPWP" and is_probable_nik(value):
                continue
            if out.get(canonical):
                # Gabungkan kolom duplicate yang maknanya sama, misalnya ASOSIASI ganda.
                if value not in out[canonical].split(" | "):
                    out[canonical] = f"{out[canonical]} | {value}"
            else:
                out[canonical] = value

        if aligned["DOMISILI_FROM_WRONG_COL"] and not out.get("PROPINSI/KOTA"):
            out["PROPINSI/KOTA"] = aligned["DOMISILI_FROM_WRONG_COL"]

        out["NO NIK"] = aligned["NO NIK"]
        out["NO NPWP"] = aligned["NO NPWP"]
        out["KRPENGALAMAN KERJA (TAHUN)"] = aligned["KRPENGALAMAN KERJA (TAHUN)"] or clean_experience(out.get("KRPENGALAMAN KERJA (TAHUN)", ""))
        out["CLEANING_NOTES"] = aligned["CLEANING_NOTES"]
        out = repair_text_contact_fields(out, row, column_meta)

        if not normalize_text(out.get("NAMA", "")):
            continue
        if normalize_header(out.get("NAMA", "")) == "NAMA":
            continue

        # Kolom helper kompatibilitas untuk metrik dan filter lama.
        out["PENDIDIKAN"] = out.get("JENIS IJAZAH", "")
        out["KEAHLIAN_SKA"] = out.get("KEAHLIAN", "")
        out["DOMISILI"] = out.get("PROPINSI/KOTA", "")
        out["EXPIRED_SKA"] = derive_expired_ska(out)
        out["TAHUN_LULUS"] = out.get("TAHUN LULUS IJAZAH", "")
        rows.append(out)

    return pd.DataFrame(rows)


def normalize_final_dataframe(df):
    if df.empty:
        return pd.DataFrame(columns=DISPLAY_ORDER)

    df = df.copy().fillna("")
    df.columns = [normalize_header(c) for c in df.columns]

    legacy_mapping = {
        "PENDIDIKAN": "JENIS IJAZAH",
        "KEAHLIAN_SKA": "KEAHLIAN",
        "DOMISILI": "PROPINSI/KOTA",
        "PENGALAMAN": "KRPENGALAMAN KERJA (TAHUN)",
        "NO NIK FINAL": "NO NIK",
        "NO NPWP FINAL": "NO NPWP",
    }
    for old, new in legacy_mapping.items():
        if old in df.columns and new not in df.columns:
            df[new] = df[old]

    for col in DISPLAY_ORDER:
        if col not in df.columns:
            df[col] = ""

    for col in df.columns:
        df[col] = df[col].apply(normalize_text)

    # Perbaikan ringan untuk database lama yang masih memuat NIK di PENGALAMAN.
    for idx, row in df.iterrows():
        notes = []
        exp = normalize_text(row.get("KRPENGALAMAN KERJA (TAHUN)", ""))
        nik = normalize_text(row.get("NO NIK", ""))
        npwp = normalize_text(row.get("NO NPWP", ""))

        if not nik and is_probable_nik(exp):
            df.at[idx, "NO NIK"] = digits_only(exp)
            df.at[idx, "KRPENGALAMAN KERJA (TAHUN)"] = ""
            append_note(notes, "NIK lama dipindahkan dari pengalaman")
        elif exp and not clean_experience(exp):
            df.at[idx, "KRPENGALAMAN KERJA (TAHUN)"] = ""

        if is_probable_nik(npwp) and digits_only(npwp) == digits_only(df.at[idx, "NO NIK"]):
            df.at[idx, "NO NPWP"] = ""
            append_note(notes, "NO NPWP berisi duplikat NIK lalu dikosongkan")

        # Perbaiki database lama tanpa perlu edit manual: sumber/email/keterangan yang salah kamar.
        row_dict = {col: df.at[idx, col] if col in df.columns else "" for col in df.columns}
        row_dict = repair_text_contact_fields(row_dict)
        for col in ["SUMBER", "EMAIL", "NO. TELP", "KETERANGAN", "CLEANING_NOTES"]:
            if col in df.columns and col in row_dict:
                df.at[idx, col] = row_dict[col]

        if notes:
            old_note = normalize_text(df.at[idx, "CLEANING_NOTES"] if "CLEANING_NOTES" in df.columns else row.get("CLEANING_NOTES", ""))
            df.at[idx, "CLEANING_NOTES"] = "; ".join([x for x in [old_note] + notes if x])

    df["PENDIDIKAN"] = df.get("JENIS IJAZAH", "")
    df["KEAHLIAN_SKA"] = df.get("KEAHLIAN", "")
    df["DOMISILI"] = df.get("PROPINSI/KOTA", "")
    df["EXPIRED_SKA"] = df.apply(derive_expired_ska, axis=1)
    df["TAHUN_LULUS"] = df.get("TAHUN LULUS IJAZAH", "")

    ordered = [c for c in DISPLAY_ORDER if c in df.columns]
    helpers = ["PENDIDIKAN", "KEAHLIAN_SKA", "DOMISILI", "EXPIRED_SKA", "TAHUN_LULUS"]
    rest = [c for c in df.columns if c not in ordered + helpers]
    return df[ordered + helpers + rest]


def proses_excel_baru(uploaded_file):
    excel_file = pd.read_excel(uploaded_file, sheet_name=None, header=None, dtype=str)
    all_data = []
    skipped_sheets = []
    processed_sheets = []
    empty_or_unread_sheets = []

    for sheet_name, df_raw in excel_file.items():
        sheet_clean = str(sheet_name).strip().lower()
        if any(keyword in sheet_clean for keyword in SKIP_SHEET_KEYWORDS):
            skipped_sheets.append(sheet_name)
            continue

        cleaned = clean_dataframe_from_sheet(df_raw, sheet_name)
        if not cleaned.empty:
            all_data.append(cleaned)
            processed_sheets.append(sheet_name)
        else:
            empty_or_unread_sheets.append(sheet_name)

    if not all_data:
        return False, 0, {
            "skipped_sheets": skipped_sheets,
            "processed_sheets": processed_sheets,
            "empty_or_unread_sheets": empty_or_unread_sheets,
            "notes": "Tidak ada sheet yang berhasil dibaca.",
        }

    master_df = pd.concat(all_data, ignore_index=True)
    master_df = normalize_final_dataframe(master_df)
    master_df = master_df[master_df["NAMA"].astype(str).str.strip() != ""]

    # Pertahankan input manual dari web lama bila ada.
    if os.path.exists(DATA_FILE):
        try:
            old_df = pd.read_csv(DATA_FILE, dtype=str).fillna("")
            old_df = normalize_final_dataframe(old_df)
            if "KATEGORI_ASAL" in old_df.columns:
                manual_df = old_df[old_df["KATEGORI_ASAL"].isin(["Input Web", "Input Web (Diedit)"])]
                if not manual_df.empty:
                    master_df = pd.concat([master_df, manual_df], ignore_index=True)
        except Exception:
            pass

    # Jangan dedupe hanya berdasarkan nama; satu orang bisa punya banyak SKA.
    dedupe_cols = ["NAMA", "KEAHLIAN", "JENIS IJAZAH", "TGL EXPIRED SKA", "KATEGORI_ASAL"]
    master_df = master_df.drop_duplicates(subset=[c for c in dedupe_cols if c in master_df.columns], keep="last")
    master_df = master_df.reset_index(drop=True)
    master_df["NO"] = range(1, len(master_df) + 1)

    if "EXPIRED_SKA" not in master_df.columns:
        master_df["EXPIRED_SKA"] = master_df.apply(derive_expired_ska, axis=1)
    date_series = master_df["EXPIRED_SKA"].apply(parse_expired_date) if "EXPIRED_SKA" in master_df.columns else pd.Series([], dtype="datetime64[ns]")

    master_df.to_csv(DATA_FILE, index=False)
    stats = {
        "processed_sheets": processed_sheets,
        "skipped_sheets": skipped_sheets,
        "empty_or_unread_sheets": empty_or_unread_sheets,
        "rows_with_nik": int(master_df["NO NIK"].astype(str).str.strip().ne("").sum()) if "NO NIK" in master_df.columns else 0,
        "rows_with_npwp": int(master_df["NO NPWP"].astype(str).str.strip().ne("").sum()) if "NO NPWP" in master_df.columns else 0,
        "rows_with_notes": int(master_df["CLEANING_NOTES"].astype(str).str.strip().ne("").sum()) if "CLEANING_NOTES" in master_df.columns else 0,
        "rows_with_unreadable_dates": int(master_df["EXPIRED_SKA"].astype(str).str.strip().ne("").sum() - date_series.notna().sum()) if "EXPIRED_SKA" in master_df.columns else 0,
    }
    return True, len(master_df), stats


def normalize_date_string(value):
    text = normalize_text(value).lower()
    if not text:
        return ""
    text = text.replace("s/d", "s.d").replace("sd", "s.d")
    text = re.sub(r"\s+", " ", text)

    if "s.d" in text:
        text = text.split("s.d")[-1].strip(" :-")
    elif " sampai " in text:
        text = text.split(" sampai ")[-1].strip(" :-")
    elif re.search(r"\d\s*-\s*\d", text) and not re.match(r"^\d{4}-\d{1,2}-\d{1,2}", text):
        text = re.split(r"\s+-\s+", text)[-1].strip()

    text = text.replace("00:00:00", "").strip()
    for indo, eng in sorted(MONTH_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        text = re.sub(rf"\b{indo}\b", eng, text)
    return text


def parse_expired_date(value):
    normalized = normalize_date_string(value)
    if not normalized:
        return pd.NaT

    # Excel sering menyimpan tanggal sebagai serial number, misalnya 44729.
    if re.fullmatch(r"\d{5}", normalized):
        try:
            serial = int(normalized)
            if 20000 <= serial <= 60000:
                return pd.to_datetime(serial, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            return pd.NaT

    # ISO yyyy-mm-dd harus diparse month-first False/day-first False agar 2029-01-11 tidak menjadi 2029-11-01.
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", normalized):
        return pd.to_datetime(normalized, errors="coerce", format="%Y-%m-%d")

    return pd.to_datetime(normalized, errors="coerce", dayfirst=True)


INDONESIAN_MONTH_NAMES = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

DATE_DISPLAY_COLUMNS = {"TGL EXPIRED SKA", "EXPIRED_SKA"}


def normalize_column_name(value):
    return normalize_text(value).upper().replace("\n", " ").strip()


def format_indonesian_date(value):
    """Format tanggal untuk tampilan saja. Data CSV asli tidak diubah."""
    text = normalize_text(value)
    if not text:
        return ""

    text = text.replace(" 00:00:00", "").replace("T00:00:00", "").strip()

    # Format ISO dari pandas/Excel, contoh 2027-05-24 atau 2027-05-24 00:00:00.
    iso_match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{2}:\d{2})?", normalize_text(value))
    if iso_match:
        try:
            dt = pd.to_datetime(f"{iso_match.group(1)}-{iso_match.group(2)}-{iso_match.group(3)}", format="%Y-%m-%d", errors="coerce")
            if pd.notna(dt):
                return f"{int(dt.day)} {INDONESIAN_MONTH_NAMES.get(int(dt.month), dt.strftime('%B'))} {int(dt.year)}"
        except Exception:
            return text

    # Timestamp pandas lain yang masih mungkin muncul.
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}.*", normalize_text(value)):
        try:
            dt = pd.to_datetime(value, errors="coerce")
            if pd.notna(dt):
                return f"{int(dt.day)} {INDONESIAN_MONTH_NAMES.get(int(dt.month), dt.strftime('%B'))} {int(dt.year)}"
        except Exception:
            pass

    # Selain ISO, pertahankan teks asli yang sudah terbaca, cukup hilangkan jam kosong.
    return text


def clean_display_value(column, value):
    if normalize_column_name(column) in DATE_DISPLAY_COLUMNS:
        return format_indonesian_date(value)
    text = normalize_text(value)
    if text.lower() in EMPTY_TOKENS:
        return ""
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def prepare_display_dataframe(table_df):
    display_df = table_df.copy()
    for col in display_df.columns:
        display_df[col] = display_df[col].apply(lambda value, c=col: clean_display_value(c, value))
    return display_df



def load_data_uncached():
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=DISPLAY_ORDER)
    df = pd.read_csv(DATA_FILE, dtype=str).fillna("")
    df = normalize_final_dataframe(df)
    if "EXPIRED_SKA" not in df.columns:
        df["EXPIRED_SKA"] = df.apply(derive_expired_ska, axis=1)
    df["EXPIRED_DATE_OBJ"] = df["EXPIRED_SKA"].apply(parse_expired_date)
    return df


@st.cache_data(show_spinner=False)
def load_data():
    return load_data_uncached()


def save_dataframe(df, backup_label="before_manual_save"):
    create_database_backup(backup_label)
    df = normalize_final_dataframe(df)
    df = df.drop(columns=["EXPIRED_DATE_OBJ"], errors="ignore")
    df.to_csv(DATA_FILE, index=False)
    save_latest_good_database_copy()
    st.cache_data.clear()


def safe_folder_name(name):
    text = normalize_text(name)
    text = re.sub(r"[\\/:*?\"<>|]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "Tanpa Nama"


def cari_folder_klien(nama_personil):
    if not os.path.exists(DOC_FOLDER):
        return None
    nama_asli = normalize_text(nama_personil).lower().replace(" ", "").replace(",", "").replace(".", "")
    try:
        for folder in os.listdir(DOC_FOLDER):
            path = os.path.join(DOC_FOLDER, folder)
            if os.path.isdir(path):
                folder_clean = folder.lower().replace(" ", "").replace("_", "").replace(",", "").replace(".", "")
                if folder_clean == nama_asli or folder_clean == f"pelamar{nama_asli}":
                    return path
    except Exception:
        return None
    return None


def folder_for_person(nama_personil):
    ensure_folder(DOC_FOLDER)
    existing = cari_folder_klien(nama_personil)
    if existing:
        return existing
    folder = os.path.join(DOC_FOLDER, safe_folder_name(nama_personil))
    ensure_folder(folder)
    return folder


def apply_global_search(df, query):
    query = normalize_text(query)
    if not query:
        return df
    mask = df.apply(lambda row: row.astype(str).str.contains(query, case=False, regex=False, na=False).any(), axis=1)
    return df[mask]

# =========================================================
# PROFESSIONAL FLOW UI (V27.1)
# =========================================================

def format_number(value):
    try:
        return f"{int(value):,}".replace(",", ".")
    except Exception:
        return str(value)


def get_kpi_values(df):
    today = pd.Timestamp.now().normalize()
    total_personil_unik = int(df["NAMA"].nunique()) if not df.empty and "NAMA" in df.columns else 0
    total_record = int(len(df))
    punya_ska = int(df["KEAHLIAN"].replace("", np.nan).notna().sum()) if not df.empty and "KEAHLIAN" in df.columns else 0
    expired = int((df["EXPIRED_DATE_OBJ"] < today).sum()) if not df.empty and "EXPIRED_DATE_OBJ" in df.columns else 0
    aktif = max(punya_ska - expired, 0)
    return {
        "personil": total_personil_unik,
        "record": total_record,
        "aktif": aktif,
        "expired": expired,
        "nik": int(df["NO NIK"].astype(str).str.strip().ne("").sum()) if not df.empty and "NO NIK" in df.columns else 0,
        "npwp": int(df["NO NPWP"].astype(str).str.strip().ne("").sum()) if not df.empty and "NO NPWP" in df.columns else 0,
    }


def render_metric_cards(df):
    kpi = get_kpi_values(df)
    cards = [
        ("Personil", f"{format_number(kpi['personil'])}", "orang"),
        ("Record", f"{format_number(kpi['record'])}", "baris"),
        ("SKA Aktif", f"{format_number(kpi['aktif'])}", "aman"),
        ("SKA Expired", f"{format_number(kpi['expired'])}", "perlu cek"),
    ]
    cols = st.columns(4)
    for col, (label, value, sub) in zip(cols, cards):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">{label}</div>
                    <div class="kpi-value">{value}</div>
                    <div class="kpi-sub">{sub}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def get_display_columns(df):
    hidden = {"EXPIRED_DATE_OBJ", "PENDIDIKAN", "KEAHLIAN_SKA", "DOMISILI", "STATUS_KERJA", "EXPIRED_SKA", "TAHUN_LULUS"}
    visible = [c for c in DISPLAY_ORDER if c in df.columns and c not in hidden]
    visible += [c for c in df.columns if c not in visible and c not in hidden and not c.startswith("UNNAMED")]
    return visible


def page_header(title, eyebrow=None, right_text=None):
    badge_html = f'<span class="page-badge">{eyebrow}</span>' if eyebrow else ""
    right_html = f'<span class="page-right">{right_text}</span>' if right_text else ""
    st.markdown(
        f"""
        <div class="page-header">
            <div>
                <div class="page-title">{title} {badge_html}</div>
            </div>
            <div>{right_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_empty_database():
    st.markdown(
        """
        <div class="empty-state">
            <div class="empty-title">Database belum tersedia</div>
            <div class="empty-sub">Upload Excel Master untuk mulai menggunakan dashboard.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Update Database", type="primary", use_container_width=True):
        st.session_state["active_page"] = "Update Database"
        st.rerun()


FILTER_STOPWORDS = {
    "ska", "skk", "skt", "sert", "sertifikat", "dokumen", "keahlian",
    "ahli", "teknik", "jenjang", "bidang", "sub", "klasifikasi",
}


def normalize_filter_text(value):
    text = normalize_text(value).lower()
    text = re.sub(r"[^0-9a-zA-ZÀ-ÿ]+", " ", text)
    return " ".join(text.split())


def smart_filter_tokens(query, keep_stopwords=False):
    text = normalize_filter_text(query)
    if not text:
        return []
    tokens = text.split()
    if not keep_stopwords:
        tokens = [token for token in tokens if token not in FILTER_STOPWORDS]
    return tokens or text.split()


def smart_contains_mask(series, query, keep_stopwords=False):
    tokens = smart_filter_tokens(query, keep_stopwords=keep_stopwords)
    if not tokens:
        return pd.Series(True, index=series.index)
    normalized_series = series.astype(str).map(normalize_filter_text)
    mask = pd.Series(True, index=series.index)
    for token in tokens:
        mask &= normalized_series.str.contains(token, case=False, regex=False, na=False)
    return mask


def filter_dataframe(df, query, ijazah_filter, keahlian_filter, penerbit_filter, status_filter):
    filtered_df = apply_global_search(df.copy(), query)
    if ijazah_filter.strip() and "JENIS IJAZAH" in filtered_df.columns:
        filtered_df = filtered_df[smart_contains_mask(filtered_df["JENIS IJAZAH"], ijazah_filter, keep_stopwords=True)]
    if keahlian_filter.strip() and "KEAHLIAN" in filtered_df.columns:
        # Smart filter: input seperti "SKK Jalan" tetap cocok dengan data Excel
        # yang tertulis "Ahli Teknik Jalan - Madya/Utama".
        filtered_df = filtered_df[smart_contains_mask(filtered_df["KEAHLIAN"], keahlian_filter, keep_stopwords=False)]
    if penerbit_filter.strip() and "PENERBIT SKA/SKK" in filtered_df.columns:
        filtered_df = filtered_df[smart_contains_mask(filtered_df["PENERBIT SKA/SKK"], penerbit_filter, keep_stopwords=True)]

    today = pd.Timestamp.now().normalize()
    if status_filter == "Aktif" and "EXPIRED_DATE_OBJ" in filtered_df.columns:
        filtered_df = filtered_df[(filtered_df["EXPIRED_DATE_OBJ"].isna()) | (filtered_df["EXPIRED_DATE_OBJ"] >= today)]
    elif status_filter == "Expired" and "EXPIRED_DATE_OBJ" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["EXPIRED_DATE_OBJ"] < today]
    elif status_filter == "Tanggal Bermasalah" and "EXPIRED_DATE_OBJ" in filtered_df.columns:
        source_col = "EXPIRED_SKA" if "EXPIRED_SKA" in filtered_df.columns else "TGL EXPIRED SKA"
        if source_col in filtered_df.columns:
            filtered_df = filtered_df[filtered_df[source_col].astype(str).str.strip().ne("") & filtered_df["EXPIRED_DATE_OBJ"].isna()]
    return filtered_df


def render_profile_section(title, row, fields):
    values = []
    for field in fields:
        value = clean_display_value(field, row.get(field, ""))
        if value:
            values.append((field, value))
    if not values:
        return
    st.markdown(f'<div class="profile-section-title">{title}</div>', unsafe_allow_html=True)
    for field, value in values:
        st.markdown(
            f"""
            <div class="profile-field">
                <div class="profile-label">{ui_label(field)}</div>
                <div class="profile-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_profile_checklist_section(title, row, fields):
    available = []
    missing = []
    for field in fields:
        if is_checked_value(row.get(field, "")):
            available.append(field)
        else:
            missing.append(field)
    if not available and not missing:
        return
    st.markdown(f'<div class="profile-section-title">{title}</div>', unsafe_allow_html=True)
    chips = []
    for field in available:
        chips.append(f'<span class="doc-chip doc-yes">✓ {html.escape(ui_label(field))}</span>')
    for field in missing:
        chips.append(f'<span class="doc-chip doc-no">– {html.escape(ui_label(field))}</span>')
    st.markdown(f'<div class="doc-chip-wrap">{"".join(chips)}</div>', unsafe_allow_html=True)


def render_profile_card(row):
    if row is None:
        st.markdown(
            """
            <div class="profile-placeholder">
                <div class="empty-title">Pilih personil</div>
                <div class="empty-sub">Profil akan muncul di sini.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    nama = normalize_text(row.get("NAMA", "-"))
    keahlian = normalize_text(row.get("KEAHLIAN", ""))
    domisili = normalize_text(row.get("PROPINSI/KOTA", ""))
    st.markdown(
        f"""
        <div class="profile-card-head">
            <div class="profile-name">{nama}</div>
            <div class="profile-meta">{keahlian or domisili or 'Data personil'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    render_profile_section("Identitas", row, ["STRATA", "JENIS IJAZAH", "TAHUN LULUS IJAZAH", "PROPINSI/KOTA", "KATEGORI_ASAL"])
    render_profile_section("Sertifikasi", row, ["KEAHLIAN", "PENERBIT SKA/SKK", "BERLAKU SKA", "TGL EXPIRED SKA", "SKA BY", "KRPENGALAMAN KERJA (TAHUN)"])
    render_profile_section("Nomor & Kontak", row, ["NO NIK", "NO NPWP", "NO. TELP", "EMAIL"])
    render_profile_checklist_section("Dokumen", row, DOCUMENT_CHECKLIST_COLUMNS)
    render_profile_section("Catatan", row, ["SUMBER", "PROYEK TERAKHIR", "KETERANGAN", "CLEANING_NOTES"])



def render_light_table(table_df, height=440):
    """Render table as controlled light HTML to avoid dark-mode dataframe contrast issues."""
    if table_df is None or table_df.empty:
        st.markdown('<div class="empty-state"><div class="empty-title">Tidak ada data</div><div class="empty-sub">Ubah kata kunci atau filter pencarian.</div></div>', unsafe_allow_html=True)
        return

    safe_df = prepare_display_dataframe(table_df).fillna("").astype(str)
    headers = "".join(f"<th>{html.escape(ui_label(col))}</th>" for col in safe_df.columns)
    rows_html = []
    for _, row in safe_df.iterrows():
        cells = "".join(f"<td>{html.escape(str(row[col]))}</td>" for col in safe_df.columns)
        rows_html.append(f"<tr>{cells}</tr>")

    table_html = f"""
    <div class="table-wrap" style="max-height:{int(height)}px;">
        <table class="light-table">
            <thead><tr>{headers}</tr></thead>
            <tbody>{''.join(rows_html)}</tbody>
        </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)

def render_search_page(df):
    page_header("Cari Data", right_text=f"{format_number(len(df))} record")
    if df.empty:
        render_empty_database()
        return

    render_metric_cards(df)
    st.markdown('<div class="section-spacer"></div>', unsafe_allow_html=True)

    toolbar_left, toolbar_right = st.columns([5, 1])
    with toolbar_left:
        global_search = st.text_input(
            "Cari",
            placeholder="Cari nama, NIK, NPWP, domisili, keahlian...",
            label_visibility="collapsed",
            key="search_global",
        )
    with toolbar_right:
        if st.button("Refresh", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    with st.expander("Filter", expanded=False):
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            ijazah_filter = st.text_input("Ijazah", placeholder="Contoh: Sipil", key="filter_ijazah")
        with f2:
            keahlian_filter = st.text_input("Keahlian", placeholder="Contoh: Jalan / Gedung / Arsitek", key="filter_keahlian")
        with f3:
            penerbit_filter = st.text_input("Penerbit", placeholder="Nama penerbit", key="filter_penerbit")
        with f4:
            status_filter = st.selectbox("Status", ["Semua", "Aktif", "Expired", "Tanggal Bermasalah"], key="filter_status")

    filtered_df = filter_dataframe(df, global_search, ijazah_filter, keahlian_filter, penerbit_filter, status_filter)
    visible_cols = get_display_columns(filtered_df)
    default_cols = [c for c in ["NO", "NAMA", "JENIS IJAZAH", "KEAHLIAN", "TGL EXPIRED SKA", "NO NIK", "NO NPWP", "PROPINSI/KOTA"] if c in visible_cols]

    # Profil tidak lagi ditampilkan sebagai daftar panjang di panel kanan.
    # Profil hanya muncul jika pencarian cukup spesifik, sehingga area baca tetap lebar.
    selected_row = None
    search_is_active = bool(normalize_text(global_search)) or any([
        normalize_text(ijazah_filter),
        normalize_text(keahlian_filter),
        normalize_text(penerbit_filter),
        status_filter != "Semua",
    ])

    if search_is_active and not filtered_df.empty:
        if len(filtered_df) == 1:
            selected_row = filtered_df.iloc[0]
        elif len(filtered_df) <= 12:
            st.markdown('<div class="section-title">Profil Personil</div>', unsafe_allow_html=True)
            options = []
            lookup = {}
            for idx, row in filtered_df.iterrows():
                nama = normalize_text(row.get("NAMA", ""))
                keahlian = normalize_text(row.get("KEAHLIAN", ""))
                kategori = normalize_text(row.get("KATEGORI_ASAL", ""))
                label_parts = [nama]
                if keahlian:
                    label_parts.append(keahlian)
                if kategori:
                    label_parts.append(kategori)
                label = " · ".join(label_parts)
                # Pastikan label unik saat satu orang punya beberapa record SKA.
                if label in lookup:
                    label = f"{label} · baris {idx}"
                options.append(label)
                lookup[label] = idx
            selected_label = st.selectbox(
                "Pilih record profil",
                options,
                index=0,
                format_func=lambda x: x,
                label_visibility="collapsed",
                key="profile_result_select",
            )
            selected_row = filtered_df.loc[lookup[selected_label]]
        else:
            st.markdown(
                f"""
                <div class="profile-placeholder profile-wide-message">
                    <div class="empty-title">{format_number(len(filtered_df))} hasil ditemukan</div>
                    <div class="empty-sub">Persempit kata kunci untuk membuka profil personil secara penuh.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    if selected_row is not None:
        st.markdown('<div class="section-title">Profil Personil</div>', unsafe_allow_html=True)
        render_profile_card(selected_row)

    st.markdown(f'<div class="section-title">Hasil Pencarian <span>{format_number(len(filtered_df))} data</span></div>', unsafe_allow_html=True)
    with st.expander("Kolom", expanded=False):
        selected_cols = st.multiselect("Kolom", options=visible_cols, default=default_cols, format_func=ui_label, label_visibility="collapsed")
    selected_cols = selected_cols or default_cols or visible_cols[:8]
    table_df = filtered_df[selected_cols].copy() if selected_cols else filtered_df.copy()
    table_show = table_df.head(DASHBOARD_TABLE_LIMIT)
    render_light_table(table_show, height=470)
    if len(table_df) > DASHBOARD_TABLE_LIMIT:
        st.caption(f"Menampilkan {DASHBOARD_TABLE_LIMIT} baris pertama.")
    export_df = prepare_display_dataframe(table_df)
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download Hasil", data=csv_bytes, file_name="hasil_pencarian.csv", mime="text/csv", use_container_width=True)

def render_kelola_personil(df):
    page_header("Kelola Data", eyebrow="Admin", right_text="Tambah / Edit Personil")
    mode = st.radio(
        "Mode",
        ["Tambah Personil", "Edit Personil"],
        horizontal=True,
        label_visibility="collapsed",
    )
    working_df = df.drop(columns=["EXPIRED_DATE_OBJ"], errors="ignore").copy()
    working_df = normalize_final_dataframe(working_df)

    selected_index = None
    defaults = {col: "" for col in DISPLAY_ORDER}

    if mode == "Edit Personil":
        if working_df.empty:
            st.warning("Belum ada data.")
            return

        search_edit = st.text_input(
            "Cari record",
            placeholder="Cari nama, NIK, atau keahlian yang akan diedit...",
            label_visibility="collapsed",
            key="edit_record_search",
        )
        candidate_df = working_df.copy()
        if search_edit.strip():
            mask = candidate_df.apply(
                lambda row: row.astype(str).str.contains(search_edit.strip(), case=False, regex=False, na=False).any(),
                axis=1,
            )
            candidate_df = candidate_df[mask]

        labels = []
        lookup = {}
        for idx, row in candidate_df.head(50).iterrows():
            nama = normalize_text(row.get("NAMA", ""))
            keahlian = normalize_text(row.get("KEAHLIAN", ""))
            domisili = normalize_text(row.get("PROPINSI/KOTA", ""))
            label_detail = " · ".join([x for x in [keahlian, domisili] if x])
            label = f"{nama}" + (f" — {label_detail}" if label_detail else "")
            labels.append(label)
            lookup[label] = idx

        if not labels:
            st.warning("Record tidak ditemukan.")
            return

        selected_label = st.radio("Pilih record", labels, label_visibility="collapsed", key="edit_record_radio")
        selected_index = lookup[selected_label]
        defaults = {col: normalize_text(working_df.at[selected_index, col]) if col in working_df.columns else "" for col in DISPLAY_ORDER}

    with st.form("form_kelola_personil"):
        st.markdown('<div class="section-title">Data Personil & Sertifikasi</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        form_values = {}
        for i, col in enumerate(PRIMARY_FORM_COLUMNS):
            target_col = [c1, c2, c3][i % 3]
            with target_col:
                form_values[col] = st.text_input(ui_label(col), value=defaults.get(col, ""), key=f"form_text_{col}")

        st.markdown('<div class="section-title">Kelengkapan Berkas</div>', unsafe_allow_html=True)
        doc_values = {}
        for group_title, group_cols in CHECKLIST_GROUPS:
            st.markdown(f'<div class="checklist-group-title">{html.escape(group_title)}</div>', unsafe_allow_html=True)
            group_layout = st.columns(3, gap="medium")
            for i, col in enumerate(group_cols):
                if col not in DOCUMENT_CHECKLIST_COLUMNS:
                    continue
                with group_layout[i % 3]:
                    doc_values[col] = st.checkbox(
                        ui_label(col),
                        value=is_checked_value(defaults.get(col, "")),
                        help=CHECKLIST_HELP.get(col),
                        key=f"doc_check_{col}",
                    )

        st.markdown('<div class="section-title">Nomor & Kontak</div>', unsafe_allow_html=True)
        n1, n2, n3, n4 = st.columns(4)
        for i, col in enumerate(CONTACT_FORM_COLUMNS):
            with [n1, n2, n3, n4][i]:
                form_values[col] = st.text_input(ui_label(col), value=defaults.get(col, ""), key=f"form_contact_{col}")

        st.markdown('<div class="section-title">Catatan</div>', unsafe_allow_html=True)
        m1, m2, m3 = st.columns(3)
        for i, col in enumerate(NOTES_FORM_COLUMNS):
            with [m1, m2, m3][i]:
                form_values[col] = st.text_input(ui_label(col), value=defaults.get(col, ""), key=f"form_notes_{col}")

        submitted = st.form_submit_button("Simpan Data", type="primary", use_container_width=True)

    if not submitted:
        return

    if not normalize_text(form_values.get("NAMA", "")):
        st.error("Nama personil wajib diisi.")
        return

    output_df = working_df.copy()
    for col in DISPLAY_ORDER:
        if col not in output_df.columns:
            output_df[col] = ""

    # Saat edit, kolom yang tidak lagi ditampilkan di form tetap dipertahankan agar data lama tidak hilang.
    # Saat tambah baru, kolom tersembunyi dibiarkan kosong.
    row_data = {col: (normalize_text(defaults.get(col, "")) if mode == "Edit Personil" else "") for col in DISPLAY_ORDER}
    for col, val in form_values.items():
        row_data[col] = normalize_text(val)
    for col, checked in doc_values.items():
        row_data[col] = checklist_to_cell(checked)

    row_data["KATEGORI_ASAL"] = "Input Web" if mode == "Tambah Personil" else "Input Web (Diedit)"
    row_data["CLEANING_NOTES"] = normalize_text(defaults.get("CLEANING_NOTES", ""))

    if mode == "Tambah Personil":
        output_df = pd.concat([output_df, pd.DataFrame([row_data])], ignore_index=True)
        folder_for_person(row_data["NAMA"])
        success_message = "Data baru tersimpan."
    else:
        for col, val in row_data.items():
            output_df.at[selected_index, col] = val
        folder_for_person(row_data["NAMA"])
        success_message = "Data berhasil diperbarui."

    output_df = normalize_final_dataframe(output_df.drop(columns=["EXPIRED_DATE_OBJ"], errors="ignore"))
    if "NO" in output_df.columns:
        output_df["NO"] = range(1, len(output_df) + 1)
    create_database_backup("before_manual_edit")
    output_df.to_csv(resolve_app_path(DATA_FILE), index=False)
    save_latest_good_database_copy()
    st.cache_data.clear()
    st.success(success_message)
    st.rerun()

def render_update_database_page(df):
    page_header("Update Database", eyebrow="Admin", right_text="Excel Master")

    status_col, action_col = st.columns([1, 1])
    db_status = database_status()
    upload_status = upload_archive_status()
    with status_col:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="status-label">Database Aktif</div>
                <div class="status-value">{'Tersedia' if db_status['exists'] else 'Kosong'}</div>
                <div class="status-sub">{db_status['modified_at']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with action_col:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="status-label">Upload Terakhir</div>
                <div class="status-value">{'Ada' if upload_status['exists'] else 'Belum ada'}</div>
                <div class="status-sub">{upload_status['modified_at']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown('<div class="section-spacer"></div>', unsafe_allow_html=True)
    file_excel_baru = st.file_uploader("Pilih file Excel", type=["xlsx"], key="excel_upload_v27")

    if file_excel_baru:
        uploaded_name = safe_filename(file_excel_baru.name)
        file_ok = uploaded_name == EXPECTED_CLIENT_EXCEL_NAME
        st.markdown(
            f"""
            <div class="upload-card">
                <div class="upload-name">{uploaded_name}</div>
                <div class="upload-sub">{round(getattr(file_excel_baru, 'size', 0) / (1024 * 1024), 2)} MB · {'Nama sesuai' if file_ok else 'Nama berbeda'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if not file_ok:
            st.warning(f"File standar: {EXPECTED_CLIENT_EXCEL_NAME}")

        if st.button("Proses Excel", type="primary", use_container_width=True):
            with st.spinner("Memproses Excel..."):
                try:
                    sukses, total, stats = process_uploaded_excel_workflow(file_excel_baru)
                except Exception as exc:
                    sukses, total, stats = False, 0, {"error": str(exc)}
            if not sukses:
                st.error(stats.get("error") or "Excel tidak dapat diproses.")
                return
            st.session_state["last_upload_result"] = stats
            st.cache_data.clear()
            st.success(f"Database diperbarui. {format_number(total)} record dimuat.")

    stats = st.session_state.get("last_upload_result")
    if stats:
        c1, c2, c3, c4 = st.columns(4)
        items = [
            ("Record", stats.get("total_records", 0)),
            ("Personil", stats.get("unique_persons", 0)),
            ("NIK", stats.get("rows_with_nik", 0)),
            ("Auto-fix", stats.get("rows_with_notes", 0)),
        ]
        for col, (label, value) in zip([c1, c2, c3, c4], items):
            with col:
                st.markdown(f'<div class="mini-card"><b>{format_number(value)}</b><span>{label}</span></div>', unsafe_allow_html=True)
        report_path = stats.get("report_path")
        report_bytes = read_report_bytes(report_path)
        b1, b2 = st.columns(2)
        with b1:
            if report_bytes:
                st.download_button("Download Laporan", data=report_bytes, file_name=os.path.basename(report_path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with b2:
            if st.button("Buka Dashboard", use_container_width=True):
                st.session_state["active_page"] = "Cari Data"
                st.rerun()


def render_documents_page(df):
    page_header("Dokumen")
    ensure_folder(DOC_FOLDER)
    if df.empty or "NAMA" not in df.columns:
        render_empty_database()
        return

    left, right = st.columns([1.2, 1])
    with left:
        names = sorted(list(df["NAMA"].dropna().unique()))
        selected_name = st.selectbox("Personil", ["Pilih personil"] + names, label_visibility="collapsed")
    with right:
        st.markdown(f'<div class="status-card"><div class="status-label">Folder Dokumen</div><div class="status-value">{len(os.listdir(DOC_FOLDER)) if os.path.exists(DOC_FOLDER) else 0}</div><div class="status-sub">folder</div></div>', unsafe_allow_html=True)

    if selected_name == "Pilih personil":
        return

    folder = cari_folder_klien(selected_name)
    if not folder:
        st.warning("Folder tidak ditemukan.")
        if st.button("Buat Folder", type="primary"):
            new_folder = folder_for_person(selected_name)
            st.success(f"Folder dibuat: {new_folder}")
        return

    st.markdown(f'<div class="folder-path">{folder}</div>', unsafe_allow_html=True)
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    if not files:
        st.info("Folder kosong.")
        return

    for filename in files:
        file_path = os.path.join(folder, filename)
        col_name, col_btn = st.columns([3, 1])
        with col_name:
            st.markdown(f'<div class="file-row">{filename}</div>', unsafe_allow_html=True)
        with col_btn:
            with open(file_path, "rb") as file_obj:
                st.download_button("Download", data=file_obj.read(), file_name=filename, use_container_width=True)


def render_admin_page(df):
    page_header("Admin & Backup", eyebrow="Admin")
    db_status = database_status()
    upload_status = upload_archive_status()
    report_status = latest_cleaning_report_status()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="status-card"><div class="status-label">Database</div><div class="status-value">{"OK" if db_status["exists"] else "Kosong"}</div><div class="status-sub">{db_status["size_mb"]} MB</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="status-card"><div class="status-label">Upload</div><div class="status-value">{"OK" if upload_status["exists"] else "-"}</div><div class="status-sub">{upload_status["modified_at"]}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="status-card"><div class="status-label">Report</div><div class="status-value">{"OK" if report_status else "-"}</div><div class="status-sub">{report_status["modified_at"] if report_status else "-"}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-spacer"></div>', unsafe_allow_html=True)
    a1, a2 = st.columns(2)
    with a1:
        if st.button("Buat Backup", use_container_width=True):
            backup_path = create_database_backup("manual")
            if backup_path:
                st.success(os.path.basename(backup_path))
            else:
                st.warning("Database kosong.")
    with a2:
        if st.button("Refresh", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    backups = list_database_backups()
    if backups:
        labels = [f"{b['filename']} · {b['modified_at']} · {b['size_mb']} MB" for b in backups]
        lookup = {label: backups[i] for i, label in enumerate(labels)}
        selected = st.selectbox("Backup", labels, label_visibility="collapsed")
        selected_backup = lookup[selected]
        d1, d2 = st.columns(2)
        with d1:
            with open(selected_backup["path"], "rb") as backup_file:
                st.download_button("Download Backup", data=backup_file.read(), file_name=selected_backup["filename"], mime="text/csv", use_container_width=True)
        with d2:
            confirm_restore = st.checkbox("Konfirmasi restore")
            if st.button("Restore", use_container_width=True, disabled=not confirm_restore):
                restore_database_from_backup(selected_backup["path"])
                st.success("Restore berhasil.")
                st.rerun()
    else:
        st.info("Belum ada backup.")

    if report_status:
        report_bytes = read_report_bytes(report_status["path"])
        if report_bytes:
            st.download_button("Download Laporan Cleaning", data=report_bytes, file_name=report_status["filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown('<div class="danger-zone">Reset Database</div>', unsafe_allow_html=True)
    confirm_reset = st.checkbox("Konfirmasi reset database")
    if st.button("Reset Database", disabled=not confirm_reset, use_container_width=True):
        reset_database_file()
        st.success("Database direset.")
        st.rerun()



def get_allowed_pages(role):
    pages = ["Cari Data", "Update Database", "Dokumen"] if role == "Admin" else ["Cari Data", "Dokumen"]
    if role == "Admin":
        pages += ["Kelola Data", "Admin & Backup"]
    return pages


def render_app_header(df, role):
    kpi = get_kpi_values(df) if not df.empty else {"personil": 0, "record": 0, "aktif": 0, "expired": 0}
    role_badge = "Admin" if role == "Admin" else "Viewer"
    db_label = "Database siap" if not df.empty else "Database kosong"
    st.markdown(
        f"""
        <div class="topbar">
            <div class="brand-lockup">
                <div class="brand-mark">HR</div>
                <div>
                    <div class="brand-title">Portal Database HRD</div>
                    <div class="brand-subtitle">Sistem Manajemen Tenaga Ahli</div>
                </div>
            </div>
            <div class="header-actions">
                <span class="status-pill">{db_label}</span>
                <span class="role-pill">{role_badge}</span>
            </div>
        </div>
        <div class="summary-strip">
            <div><span>Personil</span><b>{format_number(kpi['personil'])}</b></div>
            <div><span>Record</span><b>{format_number(kpi['record'])}</b></div>
            <div><span>SKA Aktif</span><b>{format_number(kpi['aktif'])}</b></div>
            <div><span>SKA Expired</span><b>{format_number(kpi['expired'])}</b></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if get_auth_mode() == "enabled":
        logout_col, spacer = st.columns([1, 5])
        with logout_col:
            if st.button("Logout", use_container_width=True):
                st.session_state.pop("authenticated", None)
                st.session_state.pop("role", None)
                st.rerun()



def render_sidebar_navigation(role):
    pages = get_allowed_pages(role)
    active = st.session_state.get("active_page", pages[0])
    if active not in pages:
        active = pages[0]
        st.session_state["active_page"] = active

    with st.sidebar:
        st.markdown(
            """
            <div class="side-brand">
                <div class="side-logo">HR</div>
                <div>
                    <div class="side-title">HR Portal</div>
                    <div class="side-subtitle">Database Personil</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        page = st.radio("Navigasi", pages, index=pages.index(active), label_visibility="collapsed", key="left_navigation")
        st.session_state["active_page"] = page
        st.markdown('<div class="side-footer">V41 Smart Filter UI</div>', unsafe_allow_html=True)
    return page

def inject_professional_css():
    """Tema V35: light professional content, readable left navigation, and native checklist controls."""
    st.markdown(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

            :root {
                --bg: #f5f7fb;
                --panel: #ffffff;
                --panel-soft: #f8fafc;
                --border: #d8e1ec;
                --border-strong: #b7c4d5;
                --text: #111827;
                --muted: #475569;
                --soft: #64748b;
                --primary: #1f4e79;
                --primary-hover: #163a5c;
                --primary-soft: #e8f1f8;
                --accent: #0f766e;
                --accent-hover: #0b5f59;
                --danger: #b42318;
                --success: #087443;
                --shadow: 0 10px 24px rgba(17, 24, 39, .08);
            }

            html, body, .stApp, [data-testid="stAppViewContainer"] {
                font-family: 'Inter', sans-serif !important;
                background: var(--bg) !important;
                color: var(--text) !important;
            }

            [data-testid="stHeader"] {
                background: #ffffff !important;
                border-bottom: 1px solid var(--border) !important;
            }

            .block-container {
                max-width: 1360px !important;
                padding: 3.25rem 1.75rem 3rem 1.75rem !important;
            }

            section[data-testid="stSidebar"] {
                display: block !important;
                background: #ffffff !important;
                border-right: 1px solid var(--border) !important;
                box-shadow: 8px 0 20px rgba(17, 24, 39, .04) !important;
            }
            section[data-testid="stSidebar"] * {
                color: var(--text) !important;
            }
            .side-brand {
                display: flex;
                align-items: center;
                gap: 12px;
                padding: 10px 4px 22px 4px;
                margin-bottom: 12px;
                border-bottom: 1px solid var(--border);
            }
            .side-logo {
                width: 44px;
                height: 44px;
                border-radius: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                background: var(--primary);
                color: #ffffff !important;
                font-weight: 800;
                box-shadow: 0 8px 16px rgba(31, 78, 121, .18);
            }
            .side-title { color: var(--text) !important; font-weight: 800; font-size: 1rem; line-height: 1.15; }
            .side-subtitle { color: var(--muted) !important; font-size: .8rem; font-weight: 600; margin-top: 3px; }
            .side-footer { margin-top: 28px; padding-top: 14px; border-top: 1px solid var(--border); color: var(--soft) !important; font-size: .75rem; font-weight: 700; }
            section[data-testid="stSidebar"] div[role="radiogroup"] label {
                background: #ffffff !important;
                border: 1px solid var(--border) !important;
                border-radius: 12px !important;
                padding: .76rem .86rem !important;
                margin-bottom: .5rem !important;
                box-shadow: none !important;
            }
            section[data-testid="stSidebar"] div[role="radiogroup"] label:hover { background: var(--primary-soft) !important; border-color: #c6d7e8 !important; }
            section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked),
            section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] { background: var(--primary) !important; border-color: var(--primary) !important; }
            section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) *,
            section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] * { color: #ffffff !important; font-weight: 800 !important; }

            .topbar { background: var(--panel) !important; border: 1px solid var(--border) !important; border-radius: 18px; padding: 24px 24px; box-shadow: var(--shadow); display: flex; align-items: center; justify-content: space-between; gap: 18px; margin-top: 10px; margin-bottom: 16px; }
            .brand-lockup { display: flex; align-items: center; gap: 14px; }
            .brand-mark { width: 50px; height: 50px; border-radius: 14px; background: var(--primary); color: #ffffff !important; display: flex; align-items: center; justify-content: center; font-weight: 800; box-shadow: 0 10px 18px rgba(31, 78, 121, .18); }
            .brand-title { color: var(--text) !important; font-size: 1.55rem; line-height: 1.05; font-weight: 800; letter-spacing: -.025em; }
            .brand-subtitle { color: var(--muted) !important; font-size: .94rem; margin-top: 6px; font-weight: 600; }
            .header-actions { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; justify-content: flex-end; }
            .status-pill, .role-pill { border-radius: 999px; padding: 8px 12px; font-size: .82rem; font-weight: 800; border: 1px solid var(--border); color: var(--text) !important; background: #ffffff !important; }
            .status-pill { background: #ecfdf3 !important; border-color: #bbf7d0 !important; color: #166534 !important; }
            .role-pill { background: var(--primary-soft) !important; border-color: #cfe0ef !important; color: var(--primary) !important; }
            .summary-strip { background: var(--panel) !important; border: 1px solid var(--border) !important; border-radius: 16px; padding: 12px 16px; display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px; }
            .summary-strip div { border-right: 1px solid var(--border); padding: 4px 14px; }
            .summary-strip div:last-child { border-right: none; }
            .summary-strip span { color: var(--muted) !important; display:block; font-size:.76rem; font-weight:800; text-transform:uppercase; letter-spacing:.05em; }
            .summary-strip b { color: var(--text) !important; display:block; font-size:1.24rem; font-weight:800; margin-top:2px; }

            .page-header { background: var(--panel) !important; border: 1px solid var(--border) !important; border-radius: 18px; padding: 18px 20px; margin: 18px 0 16px 0; display: flex; align-items: center; justify-content: space-between; gap: 14px; box-shadow: 0 4px 14px rgba(17, 24, 39, .04); }
            .page-title { color: var(--text) !important; font-size:1.38rem; font-weight:800; letter-spacing:-.02em; }
            .page-badge { display:inline-flex; border-radius:999px; padding:5px 9px; margin-left:8px; background:var(--primary-soft) !important; color:var(--primary) !important; font-size:.74rem; font-weight:800; }
            .page-right { color: var(--muted) !important; font-weight: 800; font-size: .9rem; }
            .section-spacer { height: 16px; }
            .section-title { color: var(--text) !important; font-size:1.05rem; font-weight:800; margin:18px 0 10px 0; }
            .section-title span { color: var(--muted) !important; font-size:.86rem; margin-left:8px; font-weight:700; }

            .kpi-card, .status-card, .profile-card-head, .profile-placeholder, .upload-card, .mini-card, .folder-path, .file-row, .empty-state { background: var(--panel) !important; border: 1px solid var(--border) !important; box-shadow: 0 4px 14px rgba(17, 24, 39, .04); }
            .kpi-card { border-radius: 18px; padding: 18px 20px; min-height: 120px; }
            .kpi-label, .status-label { color: var(--muted) !important; font-size:.76rem; font-weight:800; text-transform:uppercase; letter-spacing:.06em; }
            .kpi-value { color: var(--primary) !important; font-size:1.9rem; font-weight:800; line-height:1.08; margin-top:14px; }
            .kpi-sub, .status-sub { color: var(--muted) !important; font-size:.88rem; font-weight:600; margin-top:6px; }
            .status-card { border-radius:18px; padding:18px 20px; min-height:112px; }
            .status-value { color: var(--text) !important; font-size:1.35rem; font-weight:800; margin-top:10px; }
            .profile-card-head, .profile-placeholder { border-radius:18px; padding:20px; margin-bottom:12px; }
            .profile-name { color: var(--text) !important; font-size:1.08rem; font-weight:800; line-height:1.25; }
            .profile-meta, .empty-sub { color: var(--muted) !important; font-size:.92rem; margin-top:6px; }
            .profile-section-title { color: var(--primary) !important; font-weight:800; margin:18px 0 8px 0; font-size:.92rem; }
            .profile-field { background:#ffffff !important; border:1px solid var(--border) !important; border-radius:12px; padding:10px 12px; margin-bottom:8px; }
            .profile-wide-message { margin: 10px 0 18px 0; }
            .profile-card-head + .profile-section-title { margin-top: 14px; }

            .profile-label { color:var(--muted) !important; font-size:.72rem; font-weight:800; text-transform:uppercase; letter-spacing:.05em; }
            .profile-value { color:var(--text) !important; font-size:.94rem; font-weight:700; margin-top:4px; word-break:break-word; }
            .doc-chip-wrap { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
            .doc-chip { display:inline-flex; align-items:center; border-radius:999px; padding:7px 10px; font-size:.78rem; font-weight:800; border:1px solid var(--border); }
            .doc-yes { background:#ecfdf3 !important; color:#166534 !important; border-color:#bbf7d0 !important; }
            .doc-no { background:#f8fafc !important; color:#64748b !important; border-color:#e2e8f0 !important; }
            .checklist-panel-readable {
                background:#ffffff !important;
                border:1px solid var(--border) !important;
                border-radius:16px;
                padding:14px 16px 12px 16px;
                margin-bottom:10px;
            }
            .checklist-group-title {
                color:var(--primary) !important;
                font-size:.82rem;
                font-weight:900;
                text-transform:uppercase;
                letter-spacing:.055em;
                margin:14px 0 10px 0;
            }
            .form-note {
                color: var(--muted) !important;
                font-size: .88rem;
                font-weight: 600;
                margin: -2px 0 10px 0;
            }
            /* Checklist memakai checkbox native Streamlit agar tidak muncul double-box/ghost box. */
            [data-testid="stCheckbox"] {
                margin: .12rem 0 .52rem 0 !important;
                min-height: 34px !important;
            }
            [data-testid="stCheckbox"] p {
                color: var(--text) !important;
                font-weight: 800 !important;
                font-size: .96rem !important;
                white-space: normal !important;
                overflow: visible !important;
                text-overflow: clip !important;
                margin: 0 !important;
                line-height: 1.25 !important;
            }
            .empty-state { border-radius:18px; padding:36px; text-align:center; border-style:dashed !important; }
            .empty-title { color:var(--text) !important; font-size:1.12rem; font-weight:800; }
            .upload-card { border-radius:18px; padding:16px 18px; margin:12px 0; }
            .upload-name { color:var(--text) !important; font-size:1.02rem; font-weight:800; }
            .upload-sub { color:var(--muted) !important; font-size:.9rem; margin-top:4px; font-weight:600; }
            .mini-card { border-radius:16px; padding:16px; }
            .mini-card b { color:var(--primary) !important; font-size:1.35rem; font-weight:800; }
            .mini-card span { color:var(--muted) !important; font-size:.82rem; font-weight:800; text-transform:uppercase; letter-spacing:.04em; }
            .folder-path, .file-row { border-radius:14px; padding:12px 14px; color:var(--text) !important; font-weight:700; margin-bottom:10px; }
            .danger-zone { margin-top:28px; padding:14px 16px; border-radius:16px; background:#fff5f5 !important; border:1px solid #fecaca !important; color:var(--danger) !important; font-weight:800; }

            .stApp, .stApp p, .stApp li, .stApp label, div[data-testid="stMarkdownContainer"], div[data-testid="stMarkdownContainer"] p, div[data-testid="stMarkdownContainer"] li, div[data-testid="stCaptionContainer"], div[data-testid="stCaptionContainer"] p { color: var(--text) !important; }
            small, .caption, .stCaptionContainer { color: var(--muted) !important; }

            [data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea, [data-baseweb="select"] > div, [data-testid="stFileUploader"] section, [data-testid="stExpander"], [data-testid="stForm"] { background: #ffffff !important; color: var(--text) !important; border-color: var(--border) !important; border-radius: 14px !important; }
            [data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea { border: 1px solid var(--border-strong) !important; min-height: 44px !important; padding-left: 12px !important; }
            input::placeholder, textarea::placeholder { color:#6b7280 !important; opacity:1 !important; }

            [data-baseweb="select"], [data-baseweb="select"] *, [data-baseweb="popover"], [data-baseweb="popover"] *, [data-baseweb="menu"], [data-baseweb="menu"] *, ul[role="listbox"], ul[role="listbox"] *, div[role="option"], div[role="option"] * { color: var(--text) !important; text-shadow: none !important; }
            [data-baseweb="popover"] > div, [data-baseweb="menu"], ul[role="listbox"], div[role="listbox"] { background: #ffffff !important; border: 1px solid var(--border) !important; box-shadow: 0 16px 34px rgba(17,24,39,.18) !important; }
            div[role="option"], li[role="option"] { background: #ffffff !important; color: var(--text) !important; }
            div[role="option"]:hover, li[role="option"]:hover { background: var(--primary-soft) !important; color: var(--primary) !important; }

            div[role="radiogroup"] label { background: #ffffff !important; color: var(--text) !important; border: 1px solid var(--border) !important; border-radius: 12px !important; padding: .65rem .75rem !important; margin-bottom: .42rem !important; }
            div[role="radiogroup"] label:hover { background: var(--primary-soft) !important; border-color: #c6d7e8 !important; }
            div[role="radiogroup"] label:has(input:checked), div[role="radiogroup"] label[data-checked="true"] { background: var(--primary-soft) !important; border-color: var(--primary) !important; }
            div[role="radiogroup"] label *, div[role="radiogroup"] label:has(input:checked) * { color: var(--text) !important; font-weight: 700 !important; }

            .stButton button, .stDownloadButton button, button[data-testid="baseButton-secondary"], button[kind="secondary"] { background: #ffffff !important; border: 1px solid var(--border-strong) !important; color: var(--text) !important; border-radius: 12px !important; min-height: 44px !important; font-weight: 800 !important; box-shadow: 0 2px 8px rgba(17,24,39,.04) !important; }
            .stButton button *, .stDownloadButton button *, button[kind="secondary"] * { color: var(--text) !important; font-weight:800 !important; }
            .stButton button:hover, .stDownloadButton button:hover, button[kind="secondary"]:hover { border-color: var(--primary) !important; background: var(--primary-soft) !important; color: var(--primary) !important; }
            .stButton button:hover *, .stDownloadButton button:hover *, button[kind="secondary"]:hover * { color: var(--primary) !important; }
            .stButton button[kind="primary"], button[data-testid="baseButton-primary"], button[kind="primary"] { background: var(--accent) !important; border-color: var(--accent) !important; color:#ffffff !important; box-shadow:0 8px 18px rgba(15,118,110,.18) !important; }
            .stButton button[kind="primary"] *, button[data-testid="baseButton-primary"] *, button[kind="primary"] * { color:#ffffff !important; }
            .stButton button[kind="primary"]:hover, button[data-testid="baseButton-primary"]:hover, button[kind="primary"]:hover { background:var(--accent-hover) !important; border-color:var(--accent-hover) !important; }
            .stButton button:disabled, .stDownloadButton button:disabled, button:disabled { background:#eef2f6 !important; border-color:#d8dee8 !important; color:#8a97a8 !important; opacity:1 !important; box-shadow:none !important; }
            .stButton button:disabled *, .stDownloadButton button:disabled *, button:disabled * { color:#8a97a8 !important; }

            [data-testid="stExpander"] { border:1px solid var(--border) !important; box-shadow:0 2px 8px rgba(17,24,39,.03); }
            [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * { color:var(--text) !important; font-weight:800 !important; }

            .table-wrap { overflow: auto; border: 1px solid var(--border); border-radius: 14px; background: #ffffff; box-shadow: 0 3px 10px rgba(17,24,39,.04); }
            .light-table { width: 100%; border-collapse: collapse; min-width: 860px; background: #ffffff; color: var(--text); font-size: .88rem; }
            .light-table thead th { position: sticky; top: 0; z-index: 2; background: #edf3f8; color: var(--primary); text-align: left; padding: 12px 14px; border-bottom: 1px solid var(--border); font-size: .76rem; text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }
            .light-table tbody td { padding: 11px 14px; border-bottom: 1px solid #edf2f7; color: var(--text); background: #ffffff; vertical-align: top; white-space: nowrap; }
            .light-table tbody tr:nth-child(even) td { background: #f9fbfd; }
            .light-table tbody tr:hover td { background: var(--primary-soft); }

            [data-testid="stAlert"] { border-radius:14px !important; border:1px solid var(--border) !important; }
            [data-testid="stAlert"] * { color:var(--text) !important; }

            @media (max-width: 900px) {
                .topbar { flex-direction: column; align-items: flex-start; }
                .summary-strip { grid-template-columns: repeat(2, 1fr); }
                .summary-strip div:nth-child(2) { border-right: none; }
                .block-container { padding: 1rem !important; }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

def main():
    st.set_page_config(page_title="HR Portal", page_icon="HR", layout="wide", initial_sidebar_state="expanded")
    inject_professional_css()

    role = require_login()

    ensure_folder(DOC_FOLDER)
    ensure_folder(resolve_app_path(UPLOAD_ARCHIVE_FOLDER))
    ensure_folder(resolve_app_path(DATABASE_BACKUP_FOLDER))
    ensure_folder(resolve_app_path(CLEANING_REPORT_FOLDER))

    df = load_data()
    render_app_header(df, role)
    page = render_sidebar_navigation(role)

    if page == "Cari Data":
        render_search_page(df)
    elif page == "Update Database":
        if role != "Admin":
            st.warning("Akses admin diperlukan.")
        else:
            render_update_database_page(df)
    elif page == "Dokumen":
        render_documents_page(df)
    elif page == "Kelola Data":
        if role != "Admin":
            st.warning("Akses admin diperlukan.")
        else:
            render_kelola_personil(df)
    elif page == "Admin & Backup":
        if role != "Admin":
            st.warning("Akses admin diperlukan.")
        else:
            render_admin_page(df)


if __name__ == "__main__":
    main()
