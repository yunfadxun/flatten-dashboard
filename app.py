# ============================================================
# STREAMLIT APP
# DASHBOARD SURVEILANS CAMPAK-RUBELLA
#
# TAHAP 1:
# - Upload Excel
# - Flatten header baris Excel 9-10
#
# TAHAP 2:
# - Upload ZIP GeoJSON
# - Identifikasi boundary ADM1 dan ADM2
# - Matching provinsi dan kabupaten/kota
# ============================================================

import re
import zipfile
import unicodedata
from io import BytesIO
from pathlib import Path, PurePosixPath

import pandas as pd
import streamlit as st

from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# 1. KONFIGURASI APLIKASI
# ============================================================

st.set_page_config(
    page_title="Dashboard Surveilans MR",
    page_icon="📊",
    layout="wide"
)

SHEET_ANALISIS = "Tabel Analisis"
SHEET_KABUPATEN = "Tabel Analisis Pusat by Kab"

# Index pandas berbasis nol:
# 8 = baris Excel 9
# 9 = baris Excel 10
# 10 = baris Excel 11
HEADER_START = 8
HEADER_END = 9
DATA_START = 10

MAX_FILE_SIZE_MB = 50
MAX_BOUNDARY_SIZE_MB = 200


# ============================================================
# 2. SESSION STATE
# ============================================================

DEFAULT_SESSION_STATE = {
    "hasil_flatten": [],
    "hasil_flatten_gagal": [],
    "hasil_matching": None,
    "boundary_zip_name": None,
}

for key, default_value in DEFAULT_SESSION_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = default_value


# ============================================================
# 3. FUNGSI NORMALISASI DASAR
# ============================================================

def normalisasi_dasar(teks):
    """
    Normalisasi umum untuk nama wilayah:
    - Mengubah menjadi huruf kapital
    - Menghapus aksen
    - Mengganti tanda baca dengan spasi
    - Merapikan spasi
    """
    if teks is None or pd.isna(teks):
        return ""

    hasil = str(teks).strip().upper()

    if not hasil or hasil in {"NAN", "NONE", "NULL"}:
        return ""

    hasil = unicodedata.normalize("NFKD", hasil)
    hasil = "".join(
        karakter
        for karakter in hasil
        if not unicodedata.combining(karakter)
    )

    hasil = hasil.replace("&", " DAN ")
    hasil = re.sub(r"[/_,.;:()\-]+", " ", hasil)
    hasil = re.sub(r"[^A-Z0-9 ]+", " ", hasil)
    hasil = re.sub(r"\s+", " ", hasil).strip()

    return hasil


def normalisasi_nama_sheet(nama):
    """Normalisasi nama sheet untuk pencarian."""
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(nama).lower()
    )


def normalisasi_provinsi(teks):
    """Normalisasi dan alias nama provinsi."""
    hasil = normalisasi_dasar(teks)

    if not hasil:
        return ""

    # Hapus prefiks administratif bila ada.
    hasil = re.sub(
        r"^(PROVINSI|PROPINSI|PROV)\s+",
        "",
        hasil
    ).strip()

    alias = {
        "JAKARTA": "DKI JAKARTA",
        "DK JAKARTA": "DKI JAKARTA",
        "DKI JAKARTA": "DKI JAKARTA",
        "DAERAH KHUSUS IBUKOTA JAKARTA": "DKI JAKARTA",

        "YOGYAKARTA": "DAERAH ISTIMEWA YOGYAKARTA",
        "DI YOGYAKARTA": "DAERAH ISTIMEWA YOGYAKARTA",
        "D I YOGYAKARTA": "DAERAH ISTIMEWA YOGYAKARTA",
        "DAERAH ISTIMEWA YOGYAKARTA":
            "DAERAH ISTIMEWA YOGYAKARTA",

        "BANGKA BELITUNG": "KEPULAUAN BANGKA BELITUNG",
        "KEP BANGKA BELITUNG": "KEPULAUAN BANGKA BELITUNG",
        "KEPULAUAN BANGKA BELITUNG":
            "KEPULAUAN BANGKA BELITUNG",

        "KEP RIAU": "KEPULAUAN RIAU",
        "KEPULAUAN RIAU": "KEPULAUAN RIAU",

        "NANGGROE ACEH DARUSSALAM": "ACEH",

        "IRIAN JAYA BARAT": "PAPUA BARAT",
    }

    return alias.get(hasil, hasil)


def normalisasi_kabupaten(teks):
    """
    Normalisasi nama kabupaten/kota.

    Tipe KABUPATEN/KOTA sengaja dihilangkan. Pada tahap matching,
    nama kabupaten dipadukan dengan nama provinsi untuk menghindari
    benturan nama wilayah antarprovinsi.
    """
    hasil = normalisasi_dasar(teks)

    if not hasil:
        return ""

    hasil = re.sub(
        r"^(KABUPATEN|KAB|KOTA ADMINISTRASI|KOTA ADM|KOTA)\s+",
        "",
        hasil
    ).strip()

    alias = {
        # Perubahan nama wilayah
        "TOBA SAMOSIR": "TOBA",
        "MALUKU TENGGARA BARAT": "KEPULAUAN TANIMBAR",
        "MAMUJU UTARA": "PASANGKAYU",

        # Variasi penulisan
        "GUNUNG KIDUL": "GUNUNGKIDUL",
        "PEMATANG SIANTAR": "PEMATANGSIANTAR",
        "PADANG SIDEMPUAN": "PADANGSIDIMPUAN",
        "PADANG SIDIMPUAN": "PADANGSIDIMPUAN",
        "DHARMAS RAYA": "DHARMASRAYA",
        "PARE PARE": "PAREPARE",
        "PANGKAJENE KEPULAUAN":
            "PANGKAJENE DAN KEPULAUAN",
        "SELAYAR": "KEPULAUAN SELAYAR",
        "BAU BAU": "BAUBAU",
        "YAPEN WAROPEN": "KEPULAUAN YAPEN",
        "POLIWALI MANDAR": "POLEWALI MANDAR",
        "TAMBRAW": "TAMBRAUW",
        "GUNUNG SITOLI": "GUNUNGSITOLI",
        "BANJAR BARU": "BANJARBARU",

        # DKI Jakarta
        "JAKARTA PUSAT": "JAKARTAPUSAT",
        "JAKARTA UTARA": "JAKARTAUTARA",
        "JAKARTA BARAT": "JAKARTABARAT",
        "JAKARTA SELATAN": "JAKARTASELATAN",
        "JAKARTA TIMUR": "JAKARTATIMUR",
        "KEPULAUAN SERIBU": "KEPULAUANSERIBU",

        # Sulawesi Utara
        "SIAU TAGULANDANG BIARO":
            "SIAUTAGULANDANGBIARO",
    }

    hasil = alias.get(hasil, hasil)

    # Menghilangkan spasi untuk toleransi variasi penulisan.
    hasil = hasil.replace(" ", "")

    return hasil


def buat_key_kabupaten(provinsi, kabupaten):
    """Membuat key gabungan provinsi dan kabupaten/kota."""
    provinsi_norm = normalisasi_provinsi(provinsi)
    kabupaten_norm = normalisasi_kabupaten(kabupaten)

    if not kabupaten_norm:
        return ""

    if provinsi_norm:
        return f"{provinsi_norm}||{kabupaten_norm}"

    return kabupaten_norm


# ============================================================
# 4. FUNGSI PENCARIAN SHEET DAN KOLOM
# ============================================================

def cari_sheet(sheet_names, target_sheet):
    """
    Cari sheet secara berurutan:
    1. Kecocokan persis setelah normalisasi
    2. Kecocokan parsial
    """
    target_normal = normalisasi_nama_sheet(target_sheet)

    for nama in sheet_names:
        if normalisasi_nama_sheet(nama) == target_normal:
            return nama

    for nama in sheet_names:
        if target_normal in normalisasi_nama_sheet(nama):
            return nama

    return None


def norm_kolom(teks):
    """Normalisasi sederhana nama kolom."""
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(teks).lower()
    )


def cari_kolom(df, kandidat_nama):
    """
    Mencari kolom berdasarkan satu atau beberapa kandidat nama.

    Kandidat pertama memiliki prioritas tertinggi.
    """
    if isinstance(kandidat_nama, str):
        kandidat_nama = [kandidat_nama]

    # Pertama: exact match.
    for target in kandidat_nama:
        target_norm = norm_kolom(target)

        for kolom in df.columns:
            if norm_kolom(kolom) == target_norm:
                return kolom

    # Kedua: partial match.
    for target in kandidat_nama:
        target_norm = norm_kolom(target)

        for kolom in df.columns:
            kolom_norm = norm_kolom(kolom)

            if target_norm and target_norm in kolom_norm:
                return kolom

    return None


# ============================================================
# 5. FLATTEN HEADER EXCEL
# ============================================================

def bersihkan_nilai_header(value):
    """Bersihkan satu nilai pada header."""
    if pd.isna(value):
        return None

    value = str(value).strip()

    if not value or value.lower() == "nan":
        return None

    value = re.sub(r"\s+", " ", value)

    return value


def buat_nama_kolom_unik(nama_kolom):
    """Tambahkan suffix pada nama kolom duplikat."""
    seen = {}
    hasil = []

    for nama in nama_kolom:
        if nama not in seen:
            seen[nama] = 0
            hasil.append(nama)
        else:
            seen[nama] += 1
            hasil.append(f"{nama}_{seen[nama]}")

    return hasil


def flatten_excel_header(
    file_bytes,
    sheet_name,
    header_start=HEADER_START,
    header_end=HEADER_END,
    data_start=DATA_START
):
    """
    Melakukan flatten header dua tingkat dari baris Excel 9-10.

    Catatan penting:
    merged cell pada header kelompok biasanya hanya mempunyai nilai
    pada sel paling kiri. Karena itu, baris header pertama di-forward
    fill ke kanan.
    """
    raw = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )

    if raw.empty:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak memiliki data."
        )

    if len(raw) <= header_end:
        raise ValueError(
            f"Header baris Excel 9-10 tidak ditemukan "
            f"pada sheet '{sheet_name}'."
        )

    header_rows = raw.iloc[
        header_start:header_end + 1
    ].copy()

    header_rows = header_rows.map(
        bersihkan_nilai_header
    )

    # Forward fill hanya diperlukan untuk mempertahankan
    # grup merged header ke kolom-kolom di sebelah kanan.
    header_rows = header_rows.ffill(axis=1)

    nama_kolom = []

    for col_idx in range(raw.shape[1]):
        parts = []

        for row_idx in range(len(header_rows)):
            value = header_rows.iloc[row_idx, col_idx]

            if value is None:
                continue

            value = str(value).strip()

            if value and value not in parts:
                parts.append(value)

        if parts:
            col_name = "_".join(parts)
        else:
            col_name = f"Kolom_{col_idx + 1}"

        col_name = re.sub(r"\s+", " ", col_name).strip()
        nama_kolom.append(col_name)

    nama_kolom = buat_nama_kolom_unik(nama_kolom)

    df = raw.iloc[data_start:].copy()
    df.columns = nama_kolom

    # Ubah sel kosong berupa string menjadi NA.
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    # Hapus baris dan kolom kosong.
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    # Hapus nama kolom yang ikut terbuang.
    df = df.reset_index(drop=True)

    return df


# ============================================================
# 6. MENYIMPAN DATAFRAME KE EXCEL
# ============================================================

def format_worksheet(ws, df):
    """Format dasar worksheet hasil flatten."""
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    for col_idx, col_name in enumerate(
        df.columns,
        start=1
    ):
        sample_values = df[col_name].head(200)

        panjang = [len(str(col_name))]

        panjang.extend([
            len(str(value))
            for value in sample_values
            if pd.notna(value)
        ])

        max_length = max(panjang) if panjang else 10
        col_letter = get_column_letter(col_idx)

        ws.column_dimensions[col_letter].width = min(
            max(max_length + 2, 10),
            40
        )


def dataframe_to_excel(hasil_sheet):
    """Simpan beberapa dataframe ke satu workbook Excel."""
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for nama_sheet, df in hasil_sheet.items():
            nama_sheet_output = nama_sheet[:31]

            df.to_excel(
                writer,
                sheet_name=nama_sheet_output,
                index=False
            )

            ws = writer.sheets[nama_sheet_output]
            format_worksheet(ws, df)

    output.seek(0)

    return output.getvalue()


def dataframe_dict_to_excel(hasil_sheet):
    """Simpan dictionary dataframe untuk laporan matching."""
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for nama_sheet, df in hasil_sheet.items():
            if df is None or df.empty:
                df = pd.DataFrame({
                    "Informasi": ["Tidak ada data."]
                })

            df.to_excel(
                writer,
                sheet_name=nama_sheet[:31],
                index=False
            )

            ws = writer.sheets[nama_sheet[:31]]
            format_worksheet(ws, df)

    output.seek(0)

    return output.getvalue()


# ============================================================
# 7. PROSES SATU FILE EXCEL
# ============================================================

def proses_satu_file(file_bytes, nama_file):
    """Memproses dua sheet target dari satu workbook."""
    hasil_sheet = {}
    log = []

    excel = pd.ExcelFile(
        BytesIO(file_bytes),
        engine="openpyxl"
    )

    sheet_names = excel.sheet_names

    daftar_target = [
        {
            "target": SHEET_ANALISIS,
            "output": "Tabel Analisis"
        },
        {
            "target": SHEET_KABUPATEN,
            "output": "Analisis Pusat by Kab"
        },
    ]

    for konfigurasi in daftar_target:
        target = konfigurasi["target"]
        output_name = konfigurasi["output"]

        sheet_ditemukan = cari_sheet(
            sheet_names,
            target
        )

        if sheet_ditemukan:
            try:
                df = flatten_excel_header(
                    file_bytes=file_bytes,
                    sheet_name=sheet_ditemukan
                )

                hasil_sheet[output_name] = df

                log.append({
                    "Nama File": nama_file,
                    "Sheet Target": target,
                    "Sheet Ditemukan": sheet_ditemukan,
                    "Status": "Berhasil",
                    "Baris": len(df),
                    "Kolom": len(df.columns),
                    "Keterangan": "Header berhasil di-flatten."
                })

            except Exception as error:
                log.append({
                    "Nama File":
