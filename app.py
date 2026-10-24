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
                    "Nama File": nama_file,
                    "Sheet Target": target,
                    "Sheet Ditemukan": sheet_ditemukan,
                    "Status": "Gagal",
                    "Baris": 0,
                    "Kolom": 0,
                    "Keterangan": str(error)
                })

        else:
            log.append({
                "Nama File": nama_file,
                "Sheet Target": target,
                "Sheet Ditemukan": "",
                "Status": "Tidak ditemukan",
                "Baris": 0,
                "Kolom": 0,
                "Keterangan": (
                    "Sheet tidak ditemukan. "
                    f"Sheet tersedia: {', '.join(sheet_names)}"
                )
            })

    if not hasil_sheet:
        raise ValueError(
            "Tidak ada sheet target yang berhasil diproses. "
            f"Sheet tersedia: {', '.join(sheet_names)}"
        )

    output_bytes = dataframe_to_excel(
        hasil_sheet
    )

    nama_output = (
        f"{Path(nama_file).stem}_Flatten_Header.xlsx"
    )

    return {
        "nama_input": nama_file,
        "nama_output": nama_output,
        "data": output_bytes,
        "log": log
    }


# ============================================================
# 8. PEMBACAAN NAMA FILE BOUNDARY DARI ZIP
# ============================================================

def nama_tanpa_extension(path):
    """Mengambil nama file tanpa extension GeoJSON/JSON."""
    nama = PurePosixPath(path).name

    nama = re.sub(
        r"\.(geojson|json)$",
        "",
        nama,
        flags=re.IGNORECASE
    )

    return nama


def identifikasi_level_boundary(path):
    """
    Tentukan level boundary berdasarkan nama folder/file.

    ADM1/provinsi = level provinsi
    ADM2/kabupaten/kota = level kabupaten/kota
    """
    path_norm = normalisasi_dasar(
        path.replace("/", " ")
    )

    token_provinsi = [
        "ADM1",
        "PROVINSI",
        "PROVINCE"
    ]

    token_kabupaten = [
        "ADM2",
        "KABUPATEN",
        "KAB KOTA",
        "KABKOT",
        "REGENCY",
        "CITY"
    ]

    if any(token in path_norm for token in token_kabupaten):
        return "kabupaten"

    if any(token in path_norm for token in token_provinsi):
        return "provinsi"

    return None


def pisahkan_nama_boundary_kabupaten(nama_file):
    """
    Membaca pola nama kabupaten dalam filename.

    Pola yang disarankan:
    PROVINSI__KABUPATEN.geojson

    Jika tidak emisah '__', nama provinsi tidak
    dapat dipastikan dan matching akan memakai nama kabupaten saja.
    """
    nama = nama_tanpa_extension(nama_file)

    nama = re.sub(
        r"^(ADM2[_\-\s]*)",
        "",
        nama,
        flags=re.IGNORECASE
    )

    if "__" in nama:
        bagian = nama.split("__", maxsplit=1)

        provinsi = bagian[0].replace("_", " ")
        kabupaten = bagian[1].replace("_", " ")

        return provinsi, kabupaten

    return "", nama.replace("_", " ")


def baca_daftar_boundary(daftar_file_boundary):
    """
    Membaca boundary dari:
    - Satu file GeoJSON
    - Beberapa file GeoJSON
    - Satu atau beberapa file ZIP berisi GeoJSON

    Format input:
    [
        {"name": "ACEH.geojson", "data": bytes},
        {"name": "boundary.zip", "data": bytes}
    ]
    """

    daftar_provinsi = []
    daftar_kabupaten = []
    tidak_teridentifikasi = []

    daftar_geojson = []

    # --------------------------------------------------------
    # 1. BACA FILE GEOJSON DAN ZIP
    # --------------------------------------------------------

    for file_item in daftar_file_boundary:

        nama_file = file_item["name"]
        file_bytes = file_item["data"]

        if nama_file.lower().endswith(".zip"):

            try:
                with zipfile.ZipFile(
                    BytesIO(file_bytes),
                    mode="r"
                ) as zip_file:

                    anggota = [
                        nama
                        for nama in zip_file.namelist()
                        if not nama.endswith("/")
                        and nama.lower().endswith(
                            (".geojson", ".json")
                        )
                        and "__macosx" not in nama.lower()
                    ]

                    for path in anggota:
                        daftar_geojson.append({
                            "Path": path,
                            "Nama File": nama_tanpa_extension(path),
                        })

            except zipfile.BadZipFile:
                raise ValueError(
                    f"File '{nama_file}' bukan ZIP yang valid."
                )

        elif nama_file.lower().endswith(
            (".geojson", ".json")
        ):

            daftar_geojson.append({
                "Path": nama_file,
                "Nama File": nama_tanpa_extension(nama_file),
            })

        else:
            tidak_teridentifikasi.append({
                "Path": nama_file,
                "Keterangan": "Format file tidak didukung.",
            })

    # --------------------------------------------------------
    # 2. IDENTIFIKASI LEVEL BOUNDARY
    # --------------------------------------------------------

    for item in daftar_geojson:

        path = item["Path"]
        nama = item["Nama File"]

        level = identifikasi_level_boundary(path)

        if level is None:
            level = identifikasi_level_boundary(nama)

        if level == "provinsi":

            nama_provinsi = re.sub(
                r"^ADM1[_\-\s]*",
                "",
                nama,
                flags=re.IGNORECASE
            ).replace("_", " ")

            daftar_provinsi.append({
                "Path": path,
                "Nama Boundary": nama_provinsi,
                "Key Provinsi": normalisasi_provinsi(
                    nama_provinsi
                ),
            })

        elif level == "kabupaten":

            provinsi, kabupaten = (
                pisahkan_nama_boundary_kabupaten(path)
            )

            # Jika filename langsung tidak memuat path folder,
            # gunakan nama file untuk memisahkan nama wilayah.
            if not kabupaten:
                provinsi, kabupaten = (
                    pisahkan_nama_boundary_kabupaten(nama)
                )

            daftar_kabupaten.append({
                "Path": path,
                "Provinsi Boundary": provinsi,
                "Kabupaten/Kota Boundary": kabupaten,
                "Key Provinsi": normalisasi_provinsi(provinsi),
                "Key Kabupaten": normalisasi_kabupaten(kabupaten),
                "Key Gabungan": buat_key_kabupaten(
                    provinsi,
                    kabupaten
                ),
            })

        else:
            tidak_teridentifikasi.append({
                "Path": path,
                "Keterangan": (
                    "Level boundary tidak dikenali. "
                    "Nama file atau folder harus mengandung "
                    "ADM1/provinsi atau ADM2/kabupaten/district."
                ),
            })

    # --------------------------------------------------------
    # 3. HASIL DATAFRAME
    # --------------------------------------------------------

    df_provinsi = pd.DataFrame(
        daftar_provinsi,
        columns=[
            "Path",
            "Nama Boundary",
            "Key Provinsi",
        ]
    )

    df_kabupaten = pd.DataFrame(
        daftar_kabupaten,
        columns=[
            "Path",
            "Provinsi Boundary",
            "Kabupaten/Kota Boundary",
            "Key Provinsi",
            "Key Kabupaten",
            "Key Gabungan",
        ]
    )

    df_tidak_teridentifikasi = pd.DataFrame(
        tidak_teridentifikasi,
        columns=[
            "Path",
            "Keterangan",
        ]
    )

    return (
        df_provinsi,
        df_kabupaten,
        df_tidak_teridentifikasi
    )


# ============================================================
# 9. MATCHING BOUNDARY
# ============================================================

def matching_provinsi_excel(
    nama_file,
    df_provinsi,
    boundary_provinsi
):
    """Matching nama provinsi Excel dengan ADM1."""
    laporan = []
    peringatan = []

    kolom_prov = cari_kolom(
        df_provinsi,
        [
            "Kabupaten / Kota / Provinsi",
            "Kabupaten/Kota/Provinsi",
            "Provinsi"
        ]
    )

    if kolom_prov is None:
        peringatan.append({
            "Nama File": nama_file,
            "Bagian": "Matching Provinsi",
            "Status": "Kolom tidak ditemukan",
            "Keterangan": (
                "Kolom wilayah provinsi tidak ditemukan. "
                f"Kolom tersedia: {list(df_provinsi.columns)}"
            )
        })

        return laporan, peringatan

    key_boundary = set(
        boundary_provinsi["Key Provinsi"]
        .dropna()
        .astype(str)
    )

    nilai_provinsi = (
        df_provinsi[kolom_prov]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
    )

    for nama_provinsi in nilai_provinsi:
        key = normalisasi_provinsi(nama_provinsi)

        if key in {"", "INDONESIA", "NASIONAL", "TOTAL"}:
            continue

        laporan.append({
            "Nama File": nama_file,
            "Kolom Sumber": kolom_prov,
            "Nama Provinsi Excel": nama_provinsi,
            "Key Normalisasi": key,
            "Status": (
                "Cocok"
                if key in key_boundary
                else "Tidak cocok"
            )
        })

    return laporan, peringatan


def matching_kabupaten_excel(
    nama_file,
    df_kabupaten,
    boundary_kabupaten
):
    """Matching nama kabupaten/kota Excel dengan ADM2."""
    laporan = []
    peringatan = []

    kolom_kab = cari_kolom(
        df_kabupaten,
        [
            "Kabupaten / Kota",
            "Kabupaten/Kota",
            "Kabupaten Kota",
            "Kab/Kota"
        ]
    )

    kolom_prov = cari_kolom(
        df_kabupaten,
        [
            "Provinsi",
            "Nama Provinsi"
        ]
    )

    if kolom_kab is None:
        peringatan.append({
            "Nama File": nama_file,
            "Bagian": "Matching Kabupaten/Kota",
            "Status": "Kolom tidak ditemukan",
            "Keterangan": (
                "Kolom kabupaten/kota tidak ditemukan. "
                f"Kolom tersedia: {list(df_kabupaten.columns)}"
            )
        })

        return laporan, peringatan

    kolom_dipakai = [kolom_kab]

    if kolom_prov:
        kolom_dipakai.append(kolom_prov)

    df_nama = (
        df_kabupaten[kolom_dipakai]
        .dropna(subset=[kolom_kab])
        .drop_duplicates()
    )

    key_gabungan_boundary = set(
        boundary_kabupaten["Key Gabungan"]
        .dropna()
        .astype(str)
    )

    key_kabupaten_boundary = set(
        boundary_kabupaten["Key Kabupaten"]
        .dropna()
        .astype(str)
    )

    for _, row in df_nama.iterrows():
        nama_kab = str(row[kolom_kab]).strip()

        nama_prov = ""

        if (
            kolom_prov
            and pd.notna(row[kolom_prov])
        ):
            nama_prov = str(row[kolom_prov]).strip()

        key_prov = normalisasi_provinsi(nama_prov)
        key_kab = normalisasi_kabupaten(nama_kab)

        if key_kab in {
            "",
            "INDONESIA",
            "NASIONAL",
            "TOTAL"
        }:
            continue

        key_gabungan = buat_key_kabupaten(
            nama_prov,
            nama_kab
        )

        # Prioritas matching:
        # 1. Key gabungan provinsi + kabupaten
        # 2. Nama kabupaten saja bila boundary tidak menyimpan provinsi
        if (
            key_prov
            and key_gabungan in key_gabungan_boundary
        ):
            status = "Cocok"
            metode = "Provinsi + kabupaten/kota"

        elif key_kab in key_kabupaten_boundary:
            status = "Cocok"
            metode = "Nama kabupaten/kota"

        else:
            status = "Tidak cocok"
            metode = "Tidak ditemukan"

        laporan.append({
            "Nama File": nama_file,
            "Provinsi Excel": nama_prov,
            "Kabupaten/Kota Excel": nama_kab,
            "Key Provinsi": key_prov,
            "Key Kabupaten": key_kab,
            "Key Gabungan": key_gabungan,
            "Status": status,
            "Metode Matching": metode
        })

    return laporan, peringatan


def buat_laporan_matching(
    hasil_flatten,
    boundary_provinsi,
    boundary_kabupaten
):
    """Proses matching seluruh workbook hasil Tahap 1."""
    laporan_provinsi = []
    laporan_kabupaten = []
    laporan_file = []

    for hasil in hasil_flatten:
        nama_file = hasil["nama_input"]
        excel_bytes = hasil["data"]

        try:
            excel = pd.ExcelFile(
                BytesIO(excel_bytes),
                engine="openpyxl"
            )

        except Exception as error:
            laporan_file.append({
                "Nama File": nama_file,
                "Bagian": "Pembacaan workbook",
                "Status": "Gagal",
                "Keterangan": str(error)
            })

            continue

        # ----------------------------------------------------
        # PROVINSI
        # ----------------------------------------------------

        if "Tabel Analisis" in excel.sheet_names:
            df_prov = pd.read_excel(
                BytesIO(excel_bytes),
                sheet_name="Tabel Analisis",
                engine="openpyxl"
            )

            hasil_prov, warning_prov = (
                matching_provinsi_excel(
                    nama_file=nama_file,
                    df_provinsi=df_prov,
                    boundary_provinsi=boundary_provinsi
                )
            )

            laporan_provinsi.extend(hasil_prov)
            laporan_file.extend(warning_prov)

        else:
            laporan_file.append({
                "Nama File": nama_file,
                "Bagian": "Matching Provinsi",
                "Status": "Sheet tidak ditemukan",
                "Keterangan": (
                    "Sheet hasil 'Tabel Analisis' "
                    "tidak tersedia."
                )
            })

        # ----------------------------------------------------
        # KABUPATEN/KOTA
        # ----------------------------------------------------

        if "Analisis Pusat by Kab" in excel.sheet_names:
            df_kab = pd.read_excel(
                BytesIO(excel_bytes),
                sheet_name="Analisis Pusat by Kab",
                engine="openpyxl"
            )

            hasil_kab, warning_kab = (
                matching_kabupaten_excel(
                    nama_file=nama_file,
                    df_kabupaten=df_kab,
                    boundary_kabupaten=boundary_kabupaten
                )
            )

            laporan_kabupaten.extend(hasil_kab)
            laporan_file.extend(warning_kab)

        else:
            laporan_file.append({
                "Nama File": nama_file,
                "Bagian": "Matching Kabupaten/Kota",
                "Status": "Sheet tidak ditemukan",
                "Keterangan": (
                    "Sheet hasil 'Analisis Pusat by Kab' "
                    "tidak tersedia."
                )
            })

    return {
        "provinsi": pd.DataFrame(laporan_provinsi),
        "kabupaten": pd.DataFrame(laporan_kabupaten),
        "peringatan": pd.DataFrame(laporan_file),
    }


# ============================================================
# 10. FUNGSI TAMPILAN
# ============================================================

def tampilkan_status_tahapan():
    """Status penyelesaian pada sidebar."""
    tahap_1_selesai = (
        len(st.session_state["hasil_flatten"]) > 0
    )

    tahap_2_selesai = (
        st.session_state["hasil_matching"] is not None
    )

    if tahap_1_selesai:
        st.sidebar.success("Tahap 1: Hasil tersedia")
    else:
        st.sidebar.info("Tahap 1: Belum diproses")

    if tahap_2_selesai:
        st.sidebar.success("Tahap 2: Hasil tersedia")
    else:
        st.sidebar.info("Tahap 2: Belum diproses")


def tampilkan_preview_dataframe(df, judul, maksimal_baris=100):
    """Preview dataframe maksimal sejumlah baris tertentu."""
    if df is None or df.empty:
        st.info(f"{judul}: tidak ada data.")
        return

    st.subheader(judul)

    if len(df) > maksimal_baris:
        st.caption(
            f"Menampilkan {maksimal_baris:,} dari "
            f"{len(df):,} baris."
        )

    st.dataframe(
        df.head(maksimal_baris),
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 11. SIDEBAR NAVIGATION
# ============================================================

st.sidebar.title("Surveilans MR")
st.sidebar.caption(
    "Pengolahan data dan dashboard peta"
)

TAHAP = st.sidebar.radio(
    "NAVIGASI TAHAP",
    [
        "1. Upload & Flatten",
        "2. Boundary & Matching",
        "3. Validasi Indikator",
        "4. Dashboard Peta",
        "5. Ekspor Hasil",
    ],
    key="navigasi_tahap"
)

st.sidebar.divider()
tampilkan_status_tahapan()

st.sidebar.divider()

if st.sidebar.button(
    "Reset Seluruh Proses",
    use_container_width=True
):
    for key, default_value in DEFAULT_SESSION_STATE.items():
        st.session_state[key] = default_value

    st.rerun()


# ============================================================
# 12. HALAMAN TAHAP 1
# ============================================================

def halaman_tahap_1():
    st.title("Dashboard Surveilans Campak-Rubella")
    st.subheader("Tahap 1 — Upload dan Flatten Header Excel")

    st.info(
        "Tahap ini memproses header Excel baris 9-10. "
        "Data dianggap mulai pada baris Excel 11."
    )

    st.markdown(
        """
        **Ketentuan file Excel**

        - Format file: `.xlsm` atau `.xlsx`
        - Maksimum ukuran: 50 MB per file
        - Header: baris Excel 9-10
        - Data: mulai baris Excel 11
        - Sheet target:
          1. `Tabel Analisis`
          2. `Tabel Analisis Pusat by Kab`
        """
    )

    uploaded_files = st.file_uploader(
        "Upload file Excel",
        type=["xlsm", "xlsx"],
        accept_multiple_files=True,
        help=(
            "Semua file akan diproses satu per satu. "
            f"Maksimum {MAX_FILE_SIZE_MB} MB per file."
        ),
        key="upload_excel_tahap_1"
    )

    if uploaded_files:
        daftar_file = []

        for file in uploaded_files:
            ukuran_mb = file.size / (1024 * 1024)

            daftar_file.append({
                "Nama File": file.name,
                "Ukuran (MB)": round(ukuran_mb, 2),
                "Status": (
                    "Siap diproses"
                    if ukuran_mb <= MAX_FILE_SIZE_MB
                    else "Melebihi batas ukuran"
                )
            })

        st.subheader("File yang Diunggah")

        st.dataframe(
            pd.DataFrame(daftar_file),
            use_container_width=True,
            hide_index=True
        )

    tombol_proses = st.button(
        "Proses dan Flatten Header",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files,
        key="proses_flatten"
    )

    if tombol_proses:
        file_valid = [
            file
            for file in uploaded_files
            if file.size <= MAX_FILE_SIZE_MB * 1024 * 1024
        ]

        file_terlalu_besar = [
            file.name
            for file in uploaded_files
            if file.size > MAX_FILE_SIZE_MB * 1024 * 1024
        ]

        if file_terlalu_besar:
            st.warning(
                "File berikut melebihi batas ukuran: "
                + ", ".join(file_terlalu_besar)
            )

        if not file_valid:
            st.error(
                "Tidak ada file yang memenuhi batas ukuran."
            )

        else:
            hasil_berhasil = []
            hasil_gagal = []

            progress = st.progress(0)
            status_text = st.empty()

            for nomor, file in enumerate(
                file_valid,
                start=1
            ):
                status_text.write(
                    f"Memproses {nomor}/{len(file_valid)}: "
                    f"{file.name}"
                )

                try:
                    hasil = proses_satu_file(
                        file_bytes=file.getvalue(),
                        nama_file=file.name
                    )

                    hasil_berhasil.append(hasil)

                except Exception as error:
                    hasil_gagal.append({
                        "Nama File": file.name,
                        "Jenis Error":
                            type(error).__name__,
                        "Detail Error": str(error)
                    })

                progress.progress(
                    nomor / len(file_valid)
                )

            status_text.success("Pemrosesan selesai.")

            st.session_state["hasil_flatten"] = (
                hasil_berhasil
            )

            st.session_state["hasil_flatten_gagal"] = (
                hasil_gagal
            )

            # Jika Tahap 1 diulang, hasil Tahap 2 harus dibuang.
            st.session_state["hasil_matching"] = None

            st.rerun()

    hasil_berhasil = st.session_state["hasil_flatten"]
    hasil_gagal = st.session_state[
        "hasil_flatten_gagal"
    ]

    if hasil_berhasil or hasil_gagal:
        st.divider()
        st.subheader("Hasil Tahap 1")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "File berhasil",
            len(hasil_berhasil)
        )

        col2.metric(
            "File gagal",
            len(hasil_gagal)
        )

        col3.metric(
            "Total file",
            len(hasil_berhasil) + len(hasil_gagal)
        )

    if hasil_berhasil:
        daftar_log = []

        for hasil in hasil_berhasil:
            daftar_log.extend(hasil["log"])

        tampilkan_preview_dataframe(
            pd.DataFrame(daftar_log),
            "Laporan Flatten",
            maksimal_baris=500
        )

        st.subheader("Download Hasil Excel")

        for nomor, hasil in enumerate(hasil_berhasil):
            with st.container(border=True):
                st.write(
                    f"**{hasil['nama_output']}**"
                )

                st.download_button(
                    label="Download Excel",
                    data=hasil["data"],
                    file_name=hasil["nama_output"],
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    key=(
                        f"download_flatten_{nomor}_"
                        f"{hasil['nama_output']}"
                    ),
                    use_container_width=True
                )

        zip_buffer = BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zip_file:

            for hasil in hasil_berhasil:
                zip_file.writestr(
                    hasil["nama_output"],
                    hasil["data"]
                )

        st.download_button(
            label="Download Semua Hasil Flatten (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Hasil_Flatten_Header.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key="download_semua_flatten"
        )

    if hasil_gagal:
        st.error(
            "Ada file yang gagal diproses."
        )

        tampilkan_preview_dataframe(
            pd.DataFrame(hasil_gagal),
            "Detail Error",
            maksimal_baris=500
        )

    elif hasil_berhasil:
        st.success(
            "Tahap 1 selesai. Buka Tahap 2 pada sidebar "
            "untuk melakukan boundary matching."
        )


# ============================================================
# 13. HALAMAN TAHAP 2
# ============================================================

def halaman_tahap_2():
    st.title("Tahap 2 — Boundary & Matching")

    st.write(
        "Tahap ini membaca daftar GeoJSON dari satu file ZIP, "
        "kemudian mencocokkan nama provinsi dan kabupaten/kota "
        "dengan hasil flatten Tahap 1."
    )

    if not st.session_state["hasil_flatten"]:
        st.warning(
            "Hasil Tahap 1 belum tersedia. "
            "Jalankan Upload & Flatten terlebih dahulu."
        )
        return

    st.info(
        "Pada Tahap 2, aplikasi baru memeriksa kecocokan nama. "
        "Geometri belum digabungkan dengan indikator."
    )

    st.markdown(
        """
        **Petunjuk boundary ZIP**

        - Masukkan GeoJSON provinsi dalam folder yang mengandung
          nama `adm1` atau `provinsi`.
        - Masukkan GeoJSON kabupaten/kota dalam folder yang
          mengandung nama `adm2`, `kabupaten`, atau `kota`.
        - Format filename ADM2 yang paling aman:
          `PROVINSI__KABUPATEN.geojson`.
        """
    )

    boundary_zip = st.file_uploader(
        "Upload ZIP boundary GeoJSON",
        type=["zip"],
        accept_multiple_files=False,
        key="upload_boundary_tahap_2"
    )

    if boundary_zip:
        ukuran_mb = boundary_zip.size / (1024 * 1024)

        col1, col2 = st.columns(2)

        col1.metric(
            "File boundary",
            boundary_zip.name
        )

        col2.metric(
            "Ukuran",
            f"{ukuran_mb:.2f} MB"
        )

        if ukuran_mb > MAX_BOUNDARY_SIZE_MB:
            st.error(
                f"Ukuran ZIP melebihi "
                f"{MAX_BOUNDARY_SIZE_MB} MB."
            )

    tombol_matching = st.button(
        "Jalankan Boundary & Matching",
        type="primary",
        use_container_width=True,
        disabled=boundary_zip is None,
        key="jalankan_matching"
    )

    if tombol_matching:
        ukuran_mb = boundary_zip.size / (1024 * 1024)

        if ukuran_mb > MAX_BOUNDARY_SIZE_MB:
            st.error(
                "ZIP terlalu besar untuk diproses."
            )
            return

        try:
            with st.spinner(
                "Membaca daftar boundary..."
            ):
                (
                    boundary_provinsi,
                    boundary_kabupaten,
                    boundary_tidak_teridentifikasi
                ) = baca_daftar_boundary(
                    boundary_zip.getvalue()
                )

            if boundary_provinsi.empty:
                st.warning(
                    "Boundary provinsi tidak ditemukan. "
                    "Pastikan folder atau filename mengandung "
                    "'adm1' atau 'provinsi'."
                )

            if boundary_kabupaten.empty:
                st.warning(
                    "Boundary kabupaten/kota tidak ditemukan. "
                    "Pastikan folder atau filename mengandung "
                    "'adm2', 'kabupaten', atau 'kota'."
                )

            with st.spinner(
                "Mencocokkan nama wilayah..."
            ):
                laporan = buat_laporan_matching(
                    hasil_flatten=st.session_state[
                        "hasil_flatten"
                    ],
                    boundary_provinsi=boundary_provinsi,
                    boundary_kabupaten=boundary_kabupaten
                )

            laporan["boundary_provinsi"] = (
                boundary_provinsi
            )

            laporan["boundary_kabupaten"] = (
                boundary_kabupaten
            )

            laporan[
                "boundary_tidak_teridentifikasi"
            ] = boundary_tidak_teridentifikasi

            st.session_state["hasil_matching"] = laporan
            st.session_state["boundary_zip_name"] = (
                boundary_zip.name
            )

            st.rerun()

        except zipfile.BadZipFile:
            st.error(
                "File yang diunggah bukan ZIP yang valid."
            )

        except Exception as error:
            st.error("Tahap 2 gagal dijalankan.")
            st.exception(error)

    hasil = st.session_state["hasil_matching"]

    if hasil is None:
        return

    st.divider()
    st.subheader("Hasil Tahap 2")

    provinsi = hasil["provinsi"]
    kabupaten = hasil["kabupaten"]
    peringatan = hasil["peringatan"]

    boundary_provinsi = hasil["boundary_provinsi"]
    boundary_kabupaten = hasil["boundary_kabupaten"]

    boundary_tidak_teridentifikasi = hasil[
        "boundary_tidak_teridentifikasi"
    ]

    st.caption(
        "Boundary aktif: "
        f"{st.session_state['boundary_zip_name']}"
    )

    col1, col2 = st.columns(2)

    col1.metric(
        "Boundary provinsi",
        len(boundary_provinsi)
    )

    col2.metric(
        "Boundary kabupaten/kota",
        len(boundary_kabupaten)
    )

    st.subheader("Ringkasan Matching")

    col1, col2, col3, col4 = st.columns(4)

    prov_cocok = (
        int((provinsi["Status"] == "Cocok").sum())
        if not provinsi.empty
        else 0
    )

    prov_tidak_cocok = (
        int(
            (provinsi["Status"] == "Tidak cocok").sum()
        )
        if not provinsi.empty
        else 0
    )

    kab_cocok = (
        int((kabupaten["Status"] == "Cocok").sum())
        if not kabupaten.empty
        else 0
    )

    kab_tidak_cocok = (
        int(
            (kabupaten["Status"] == "Tidak cocok").sum()
        )
        if not kabupaten.empty
        else 0
    )

    col1.metric("Provinsi cocok", prov_cocok)
    col2.metric(
        "Provinsi tidak cocok",
        prov_tidak_cocok
    )
    col3.metric("Kab/Kota cocok", kab_cocok)
    col4.metric(
        "Kab/Kota tidak cocok",
        kab_tidak_cocok
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "Matching Provinsi",
        "Matching Kabupaten/Kota",
        "Boundary Tidak Teridentifikasi",
        "Peringatan"
    ])

    with tab1:
        tampilkan_preview_dataframe(
            provinsi,
            "Detail Matching Provinsi",
            maksimal_baris=1_000
        )

    with tab2:
        tampilkan_preview_dataframe(
            kabupaten,
            "Detail Matching Kabupaten/Kota",
            maksimal_baris=2_000
        )

    with tab3:
        tampilkan_preview_dataframe(
            boundary_tidak_teridentifikasi,
            "Boundary Tidak Teridentifikasi",
            maksimal_baris=1_000
        )

    with tab4:
        tampilkan_preview_dataframe(
            peringatan,
            "Peringatan Pembacaan",
            maksimal_baris=1_000
        )

    file_laporan = dataframe_dict_to_excel({
        "Matching Provinsi": provinsi,
        "Matching Kabupaten": kabupaten,
        "Peringatan": peringatan,
        "Boundary Provinsi": boundary_provinsi,
        "Boundary Kabupaten": boundary_kabupaten,
        "Tidak Teridentifikasi":
            boundary_tidak_teridentifikasi
    })

    st.download_button(
        "Download Laporan Matching",
        data=file_laporan,
        file_name="Laporan_Matching_Tahap_2.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
        key="download_laporan_matching"
    )

    if prov_tidak_cocok == 0 and kab_tidak_cocok == 0:
        st.success(
            "Semua nama wilayah Excel berhasil dicocokkan "
            "dengan boundary."
        )
    else:
        st.warning(
            "Masih terdapat nama wilayah yang tidak cocok. "
            "Gunakan laporan matching untuk menambahkan alias "
            "yang diperlukan."
        )


# ============================================================
# 14. HALAMAN TAHAP BERIKUTNYA
# ============================================================

def halaman_belum_aktif(judul):
    st.title(judul)

    if not st.session_state["hasil_flatten"]:
        st.warning(
            "Tahap 1 belum selesai."
        )
        return

    if st.session_state["hasil_matching"] is None:
        st.warning(
            "Tahap 2 belum selesai."
        )
        return

    st.info(
        "Tahap ini belum diaktifkan. "
        "Hasil Tahap 1 dan Tahap 2 sudah tersedia "
        "untuk digunakan pada pengembangan berikutnya."
    )


# ============================================================
# 15. ROUTING HALAMAN
# ============================================================

if TAHAP == "1. Upload & Flatten":
    halaman_tahap_1()

elif TAHAP == "2. Boundary & Matching":
    halaman_tahap_2()

elif TAHAP == "3. Validasi Indikator":
    halaman_belum_aktif(
        "Tahap 3 — Validasi Indikator"
    )

elif TAHAP == "4. Dashboard Peta":
    halaman_belum_aktif(
        "Tahap 4 — Dashboard Peta"
    )

elif TAHAP == "5. Ekspor Hasil":
    halaman_belum_aktif(
        "Tahap 5 — Ekspor Hasil"
    )
