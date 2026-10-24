# ============================================================
# STREAMLIT APP: SURVEILANS MR - TAHAP 1
# UPLOAD DAN FLATTEN HEADER EXCEL
# ============================================================

import re
import zipfile
import traceback
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from openpyxl.utils import get_column_letter


# ============================================================
# 1. KONFIGURASI STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Dashboard Surveilans MR",
    page_icon="📊",
    layout="wide"
)

st.title("Dashboard Surveilans Campak-Rubella")

st.subheader("Tahap 1 — Flatten Header Excel")

st.info(
    "Tahap ini hanya memproses header Excel. "
    "Proses boundary, matching wilayah, validasi indikator, "
    "dan pembuatan peta akan dilakukan pada tahap berikutnya."
)


# ============================================================
# 2. KONFIGURASI PEMROSESAN
# ============================================================

SHEET_ANALISIS = "Tabel Analisis"

SHEET_KAKO = "Tabel Analisis Pusat by Kab"

HEADER_START = 8
HEADER_END = 9
DATA_START = 10

MAX_FILE_SIZE_MB = 50


st.markdown(
    """
    **Ketentuan file:**
    - Format: XLSM atau XLSX.
    - Header: baris Excel 9–10.
    - Data dimulai dari baris Excel 11.
    - Sheet yang diproses:
        1. `Tabel Analisis`
        2. `Tabel Analisis Pusat by Kab`
    """
)


# ============================================================
# 3. UPLOAD FILE
# ============================================================

uploaded_files = st.file_uploader(
    "Upload file Excel",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True,
    help=(
        "Upload file Excel yang akan diproses. "
        f"Maksimum {MAX_FILE_SIZE_MB} MB per file."
    )
)


# ============================================================
# 4. NORMALISASI NAMA SHEET
# ============================================================

def normalisasi_nama_sheet(nama):

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(nama).lower()
    )


# ============================================================
# 5. CARI SHEET
# ============================================================

def cari_sheet(sheet_names, target_sheet):

    target_normal = normalisasi_nama_sheet(
        target_sheet
    )

    # Pencarian nama yang sama setelah normalisasi.
    for nama in sheet_names:

        if normalisasi_nama_sheet(nama) == target_normal:
            return nama

    # Pencarian nama target sebagai bagian nama sheet.
    for nama in sheet_names:

        if target_normal in normalisasi_nama_sheet(nama):
            return nama

    return None


# ============================================================
# 6. FLATTEN HEADER
# ============================================================

def flatten_excel_header(
    file_bytes,
    sheet_name,
    header_start=HEADER_START,
    header_end=HEADER_END,
    data_start=DATA_START
):

    # Baca data mentah tanpa header.
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
            f"Header baris 9–10 tidak ditemukan "
            f"pada sheet '{sheet_name}'."
        )

    # Ambil dua baris header.
    header_rows = raw.iloc[
        header_start:header_end + 1
    ].copy()

    # Bersihkan nilai header.
    header_rows = header_rows.map(
        lambda x: (
            str(x).strip()
            if pd.notna(x) and str(x).strip() != ""
            else None
        )
    )

    # Isi header kelompok ke kanan.
    header_rows = header_rows.ffill(axis=1)

    # Gabungkan nama header.
    new_columns = []

    for col_idx in range(raw.shape[1]):

        parts = []

        for row_idx in range(len(header_rows)):

            value = header_rows.iloc[
                row_idx,
                col_idx
            ]

            if value is not None:

                value = str(value).strip()

                if value and (
                    not parts or value != parts[-1]
                ):

                    parts.append(value)

        if parts:

            col_name = "_".join(parts)

        else:

            col_name = f"Kolom_{col_idx + 1}"

        new_columns.append(col_name)

    # Pastikan nama kolom unik.
    seen = {}

    unique_columns = []

    for col in new_columns:

        if col not in seen:

            seen[col] = 0

            unique_columns.append(col)

        else:

            seen[col] += 1

            unique_columns.append(
                f"{col}_{seen[col]}"
            )

    # Ambil data setelah baris header.
    df = raw.iloc[data_start:].copy()

    df.columns = unique_columns

    # Hapus baris yang seluruhnya kosong.
    df = df.dropna(how="all")

    # Hapus kolom yang seluruh datanya kosong.
    df = df.dropna(
        axis=1,
        how="all"
    )

    return df.reset_index(drop=True)


# ============================================================
# 7. SIMPAN HASIL KE EXCEL
# ============================================================

def dataframe_to_excel(hasil_sheet):

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

            # Bekukan baris header.
            ws.freeze_panes = "A2"

            # Aktifkan filter.
            ws.auto_filter.ref = ws.dimensions

            # Atur lebar kolom berdasarkan sampel data.
            for col_idx, col_name in enumerate(
                df.columns,
                start=1
            ):

                sample_values = df[col_name].head(200)

                max_length = max(
                    [len(str(col_name))]
                    + [
                        len(str(value))
                        for value in sample_values
                        if pd.notna(value)
                    ]
                )

                col_letter = get_column_letter(
                    col_idx
                )

                ws.column_dimensions[
                    col_letter
                ].width = min(
                    max_length + 2,
                    40
                )

    output.seek(0)

    return output.getvalue()


# ============================================================
# 8. PROSES SATU FILE
# ============================================================

def proses_satu_file(file_bytes, nama_file):

    hasil_sheet = {}

    log = []

    # Baca daftar sheet.
    excel = pd.ExcelFile(
        BytesIO(file_bytes),
        engine="openpyxl"
    )

    sheet_names = excel.sheet_names

    # --------------------------------------------------------
    # SHEET TABEL ANALISIS
    # --------------------------------------------------------

    sheet_analisis = cari_sheet(
        sheet_names,
        SHEET_ANALISIS
    )

    if sheet_analisis:

        df_analisis = flatten_excel_header(
            file_bytes=file_bytes,
            sheet_name=sheet_analisis
        )

        hasil_sheet[
            "Tabel Analisis"
        ] = df_analisis

        log.append({
            "Nama File": nama_file,
            "Sheet": SHEET_ANALISIS,
            "Status": "Berhasil",
            "Baris": len(df_analisis),
            "Kolom": len(df_analisis.columns),
            "Keterangan": "Header berhasil di-flatten."
        })

    else:

        log.append({
            "Nama File": nama_file,
            "Sheet": SHEET_ANALISIS,
            "Status": "Tidak ditemukan",
            "Baris": 0,
            "Kolom": 0,
            "Keterangan": (
                "Sheet tidak ditemukan. "
                f"Sheet tersedia: {sheet_names}"
            )
        })

    # --------------------------------------------------------
    # SHEET TABEL ANALISIS PUSAT BY KAB
    # --------------------------------------------------------

    sheet_kako = cari_sheet(
        sheet_names,
        SHEET_KAKO
    )

    if sheet_kako:

        df_kako = flatten_excel_header(
            file_bytes=file_bytes,
            sheet_name=sheet_kako
        )

        hasil_sheet[
            "Analisis Pusat by Kab"
        ] = df_kako

        log.append({
            "Nama File": nama_file,
            "Sheet": SHEET_KAKO,
            "Status": "Berhasil",
            "Baris": len(df_kako),
            "Kolom": len(df_kako.columns),
            "Keterangan": "Header berhasil di-flatten."
        })

    else:

        log.append({
            "Nama File": nama_file,
            "Sheet": SHEET_KAKO,
            "Status": "Tidak ditemukan",
            "Baris": 0,
            "Kolom": 0,
            "Keterangan": (
                "Sheet tidak ditemukan. "
                f"Sheet tersedia: {sheet_names}"
            )
        })

    # --------------------------------------------------------
    # VALIDASI HASIL
    # --------------------------------------------------------

    if not hasil_sheet:

        raise ValueError(
            "Kedua sheet target tidak ditemukan. "
            f"Sheet yang tersedia: {sheet_names}"
        )

    # Buat file Excel hasil flatten.
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
# 9. TAMPILKAN FILE YANG DIUPLOAD
# ============================================================

if uploaded_files:

    st.subheader("File yang Diunggah")

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

    st.dataframe(
        pd.DataFrame(daftar_file),
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# 10. TOMBOL PEMROSESAN
# ============================================================

if st.button(
    "Proses dan Flatten Header",
    type="primary",
    use_container_width=True,
    disabled=not uploaded_files
):

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

                file_bytes = file.getvalue()

                hasil = proses_satu_file(
                    file_bytes=file_bytes,
                    nama_file=file.name
                )

                hasil_berhasil.append(hasil)

            except Exception as error:

                hasil_gagal.append({
                    "Nama File": file.name,
                    "Jenis Error": type(error).__name__,
                    "Detail Error": str(error)
                })

            progress.progress(
                nomor / len(file_valid)
            )

        status_text.write(
            "Pemrosesan selesai."
        )

        # Simpan hasil pemrosesan.
        st.session_state["hasil_flatten"] = (
            hasil_berhasil
        )

        st.session_state["hasil_flatten_gagal"] = (
            hasil_gagal
        )


# ============================================================
# 11. TAMPILKAN HASIL
# ============================================================

if "hasil_flatten" in st.session_state:

    hasil_berhasil = st.session_state[
        "hasil_flatten"
    ]

    hasil_gagal = st.session_state.get(
        "hasil_flatten_gagal",
        []
    )

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

    # --------------------------------------------------------
    # LAPORAN FLATTEN
    # --------------------------------------------------------

    if hasil_berhasil:

        st.subheader("Laporan Flatten")

        daftar_log = []

        for hasil in hasil_berhasil:

            daftar_log.extend(
                hasil["log"]
            )

        st.dataframe(
            pd.DataFrame(daftar_log),
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # DOWNLOAD PER FILE
        # ----------------------------------------------------

        st.subheader("Download Hasil Excel")

        for hasil in hasil_berhasil:

            with st.container(border=True):

                st.write(
                    f"**{hasil['nama_output']}**"
                )

                st.download_button(
                    label="Download Excel",
                    data=hasil["data"],
                    file_name=hasil["nama_output"],
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    key=f"download_{hasil['nama_output']}",
                    use_container_width=True
                )

        # ----------------------------------------------------
        # DOWNLOAD ZIP
        # ----------------------------------------------------

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
            label="Download Semua Hasil (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Hasil_Flatten_Header.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

    # --------------------------------------------------------
    # LAPORAN ERROR
    # --------------------------------------------------------

    if hasil_gagal:

        st.error(
            "Ada file yang gagal diproses. "
            "Periksa detail error di bawah."
        )

        st.dataframe(
            pd.DataFrame(hasil_gagal),
            use_container_width=True,
            hide_index=True
        )

    elif hasil_berhasil:

        st.success(
            "Pemrosesan flatten selesai. "
            "Silakan periksa dan unduh hasilnya."
        )

    # --------------------------------------------------------
    # RESET
    # --------------------------------------------------------

    if st.button("Reset Hasil"):

        st.session_state.pop(
            "hasil_flatten",
            None
        )

        st.session_state.pop(
            "hasil_flatten_gagal",
            None
        )

        st.rerun()


# ============================================================
# 12. INFORMASI TAHAP BERIKUTNYA
# ============================================================

st.divider()

st.caption(
    "Tahap berikutnya: boundary GeoJSON, matching wilayah, "
    "validasi indikator, dan dashboard peta."
)


# ============================================================
# SIDEBAR NAVIGATION + TAHAP 2
# ============================================================

import json
import re
import unicodedata
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError


# ============================================================
# 1. NAVIGASI TAHAP
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

tahap_selesai = (
    "hasil_flatten" in st.session_state
    and len(st.session_state["hasil_flatten"]) > 0
)

if tahap_selesai:
    st.sidebar.success("Tahap 1: Hasil tersedia")
else:
    st.sidebar.info("Tahap 1: Belum ada hasil")


# ============================================================
# 2. NORMALISASI NAMA WILAYAH
# ============================================================

def normalisasi_dasar(teks):

    if pd.isna(teks):
        return ""

    teks = str(teks).strip().upper()

    teks = unicodedata.normalize(
        "NFKD",
        teks
    )

    teks = "".join(
        karakter
        for karakter in teks
        if not unicodedata.combining(karakter)
    )

    teks = teks.replace("&", " DAN ")

    teks = re.sub(
        r"[^A-Z0-9]+",
        " ",
        teks
    )

    teks = re.sub(
        r"\s+",
        " ",
        teks
    ).strip()

    return teks


def normalisasi_provinsi(teks):

    hasil = normalisasi_dasar(teks)

    alias = {
        "JAKARTA": "DKI JAKARTA",
        "DK JAKARTA": "DKI JAKARTA",
        "DAERAH KHUSUS IBUKOTA JAKARTA": "DKI JAKARTA",

        "YOGYAKARTA": "DAERAH ISTIMEWA YOGYAKARTA",
        "DI YOGYAKARTA": "DAERAH ISTIMEWA YOGYAKARTA",
        "D I YOGYAKARTA": "DAERAH ISTIMEWA YOGYAKARTA",

        "BANGKA BELITUNG": "KEPULAUAN BANGKA BELITUNG",
        "KEP BANGKA BELITUNG": "KEPULAUAN BANGKA BELITUNG",

        "KEP RIAU": "KEPULAUAN RIAU",
        "NANGGROE ACEH DARUSSALAM": "ACEH",
    }

    return alias.get(hasil, hasil)


def normalisasi_kabupaten(teks):

    hasil = normalisasi_dasar(teks)

    hasil = re.sub(
        r"^KABUPATEN\s+",
        "",
        hasil
    )

    hasil = re.sub(
        r"^KAB\s+",
        "",
        hasil
    )

    alias = {
        "KOTA GUNUNG SITOLI": "KOTA GUNUNGSITOLI",
        "DHARMAS RAYA": "DHARMASRAYA",
        "KOTA PARE PARE": "KOTA PAREPARE",
        "PANGKAJENE KEPULAUAN": "PANGKAJENE DAN KEPULAUAN",
        "SELAYAR": "KEPULAUAN SELAYAR",
        "KOTA BAU BAU": "KOTA BAUBAU",
        "YAPEN WAROPEN": "KEPULAUAN YAPEN",
        "POLIWALI MANDAR": "POLEWALI MANDAR",
        "TAMBRAW": "TAMBRAUW",
    }

    return alias.get(hasil, hasil)


# ============================================================
# 3. AMBIL DAFTAR BOUNDARY DARI GITHUB
# ============================================================

@st.cache_data(ttl=86400, show_spinner=False)
def ambil_daftar_boundary():

    url = (
        "https://api.github.com/repos/"
        "dmxsan/indonesia-admin-boundaries/"
        "git/trees/main?recursive=1"
    )

    request = Request(
        url,
        headers={
            "User-Agent": "Streamlit-Surveilans-MR"
        }
    )

    with urlopen(request, timeout=60) as response:
        data = json.loads(
            response.read().decode("utf-8")
        )

    if data.get("truncated"):
        raise RuntimeError(
            "Daftar file GitHub terpotong. "
            "Daftar boundary belum dapat divalidasi."
        )

    files = [
        item["path"]
        for item in data.get("tree", [])
        if item.get("type") == "blob"
        and item["path"].lower().endswith(".geojson")
    ]

    prefix_provinsi = (
        "processed-data/02-provinces/province-only/"
    )

    prefix_kabupaten = (
        "processed-data/03-districts/"
    )

    files_provinsi = [
        path
        for path in files
        if path.startswith(prefix_provinsi)
    ]

    files_kabupaten = [
        path
        for path in files
        if path.startswith(prefix_kabupaten)
    ]

    if not files_provinsi:
        raise ValueError(
            "File boundary provinsi tidak ditemukan."
        )

    if not files_kabupaten:
        raise ValueError(
            "File boundary kabupaten/kota tidak ditemukan."
        )

    return files_provinsi, files_kabupaten


# ============================================================
# 4. IDENTIFIKASI KOLOM EXCEL
# ============================================================

def cari_kolom(df, nama_target):

    def norm_kolom(teks):
        return re.sub(
            r"[^a-z0-9]",
            "",
            str(teks).lower()
        )

    target = norm_kolom(nama_target)

    # Prioritaskan kecocokan persis.
    for kolom in df.columns:

        if norm_kolom(kolom) == target:
            return kolom

    # Jika nama header memiliki awalan tambahan.
    for kolom in df.columns:

        if target in norm_kolom(kolom):
            return kolom

    return None


# ============================================================
# 5. PROSES MATCHING
# ============================================================

def buat_laporan_matching(
    hasil_flatten,
    files_provinsi,
    files_kabupaten
):

    boundary_provinsi = {
        normalisasi_provinsi(
            path.split("/")[-1]
            .removesuffix(".geojson")
            .replace("_", " ")
        )
        for path in files_provinsi
    }

    boundary_kabupaten = {
        normalisasi_kabupaten(
            path.split("/")[-1]
            .removesuffix(".geojson")
            .replace("_", " ")
        )
        for path in files_kabupaten
    }

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
                "Status": "Gagal membaca hasil flatten",
                "Keterangan": str(error)
            })

            continue

        # ----------------------------------------------------
        # MATCHING PROVINSI
        # ----------------------------------------------------

        if "Tabel Analisis" in excel.sheet_names:

            df_prov = pd.read_excel(
                BytesIO(excel_bytes),
                sheet_name="Tabel Analisis",
                engine="openpyxl"
            )

            kolom_prov = cari_kolom(
                df_prov,
                "Kabupaten / Kota / Provinsi"
            )

            if kolom_prov is None:

                laporan_file.append({
                    "Nama File": nama_file,
                    "Status": "Kolom provinsi tidak ditemukan",
                    "Keterangan": (
                        "Periksa header Tabel Analisis."
                    )
                })

            else:

                nilai_prov = (
                    df_prov[kolom_prov]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
                )

                for nama in nilai_prov:

                    key = normalisasi_provinsi(nama)

                    if key in ("", "INDONESIA"):
                        continue

                    laporan_provinsi.append({
                        "Nama File": nama_file,
                        "Nama Provinsi Excel": nama,
                        "Nama Normalisasi": key,
                        "Status": (
                            "Cocok"
                            if key in boundary_provinsi
                            else "Tidak cocok"
                        )
                    })

        # ----------------------------------------------------
        # MATCHING KABUPATEN/KOTA
        # ----------------------------------------------------

        if "Analisis Pusat by Kab" in excel.sheet_names:

            df_kab = pd.read_excel(
                BytesIO(excel_bytes),
                sheet_name="Analisis Pusat by Kab",
                engine="openpyxl"
            )

            kolom_kab = cari_kolom(
                df_kab,
                "Kabupaten / Kota"
            )

            kolom_prov_kab = cari_kolom(
                df_kab,
                "Provinsi"
            )

            if kolom_kab is None:

                laporan_file.append({
                    "Nama File": nama_file,
                    "Status": "Kolom kabupaten/kota tidak ditemukan",
                    "Keterangan": (
                        "Periksa header Analisis Pusat by Kab."
                    )
                })

            else:

                kolom_yang_dibutuhkan = [kolom_kab]

                if kolom_prov_kab:
                    kolom_yang_dibutuhkan.append(
                        kolom_prov_kab
                    )

                df_nama = (
                    df_kab[kolom_yang_dibutuhkan]
                    .dropna(subset=[kolom_kab])
                    .drop_duplicates()
                )

                for _, row in df_nama.iterrows():

                    nama_kab = str(row[kolom_kab]).strip()

                    nama_prov = (
                        str(row[kolom_prov_kab]).strip()
                        if kolom_prov_kab
                        and pd.notna(row[kolom_prov_kab])
                        else ""
                    )

                    key_kab = normalisasi_kabupaten(
                        nama_kab
                    )

                    if key_kab in ("", "INDONESIA"):
                        continue

                    laporan_kabupaten.append({
                        "Nama File": nama_file,
                        "Provinsi Excel": nama_prov,
                        "Kabupaten/Kota Excel": nama_kab,
                        "Nama Normalisasi": key_kab,
                        "Status Nama Boundary": (
                            "Cocok"
                            if key_kab in boundary_kabupaten
                            else "Tidak cocok"
                        )
                    })

    return (
        laporan_provinsi,
        laporan_kabupaten,
        laporan_file,
        len(files_provinsi),
        len(files_kabupaten)
    )


# ============================================================
# 6. TAMPILKAN HALAMAN TAHAP 2
# ============================================================

if TAHAP == "2. Boundary & Matching":

    st.title("Tahap 2 — Boundary & Matching")

    st.write(
        "Tahap ini mengambil daftar GeoJSON dari repository "
        "boundary Indonesia dan mencocokkan nama wilayah "
        "dengan hasil flatten Tahap 1."
    )

    if not tahap_selesai:

        st.warning(
            "Hasil Tahap 1 belum tersedia pada sesi ini. "
            "Silakan buka Tahap 1 dan jalankan flatten terlebih dahulu."
        )

        st.stop()

    st.info(
        "Tahap ini memeriksa kecocokan nama berdasarkan daftar file "
        "boundary. Geometri belum diunduh atau diproses pada tahap ini."
    )

    if st.button(
        "Jalankan Boundary & Matching",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Mengambil daftar boundary dari GitHub..."
            ):

                (
                    files_provinsi,
                    files_kabupaten
                ) = ambil_daftar_boundary()

            st.success(
                "Daftar boundary berhasil diambil."
            )

            with st.spinner(
                "Mencocokkan nama wilayah..."
            ):

                (
                    laporan_provinsi,
                    laporan_kabupaten,
                    laporan_file,
                    jumlah_boundary_provinsi,
                    jumlah_boundary_kabupaten
                ) = buat_laporan_matching(
                    st.session_state["hasil_flatten"],
                    files_provinsi,
                    files_kabupaten
                )

            st.session_state["hasil_matching"] = {
                "provinsi": laporan_provinsi,
                "kabupaten": laporan_kabupaten,
                "file": laporan_file,
                "jumlah_boundary_provinsi": jumlah_boundary_provinsi,
                "jumlah_boundary_kabupaten": jumlah_boundary_kabupaten,
            }

        except (HTTPError, URLError, TimeoutError) as error:

            st.error(
                "Tidak dapat mengambil daftar boundary dari GitHub."
            )

            st.code(str(error))

        except Exception as error:

            st.error("Tahap 2 gagal dijalankan.")

            st.exception(error)

    # --------------------------------------------------------
    # TAMPILKAN HASIL MATCHING
    # --------------------------------------------------------

    if "hasil_matching" in st.session_state:

        hasil = st.session_state["hasil_matching"]

        provinsi = pd.DataFrame(hasil["provinsi"])

        kabupaten = pd.DataFrame(hasil["kabupaten"])

        file_error = pd.DataFrame(hasil["file"])

        st.divider()

        st.subheader("Ringkasan Boundary")

        col1, col2 = st.columns(2)

        col1.metric(
            "File boundary provinsi",
            hasil["jumlah_boundary_provinsi"]
        )

        col2.metric(
            "File boundary kabupaten/kota",
            hasil["jumlah_boundary_kabupaten"]
        )

        if not provinsi.empty:

            cocok = (
                provinsi["Status"] == "Cocok"
            ).sum()

            tidak_cocok = (
                provinsi["Status"] == "Tidak cocok"
            ).sum()

            st.subheader("Matching Provinsi")

            col1, col2 = st.columns(2)

            col1.metric("Cocok", int(cocok))

            col2.metric("Tidak cocok", int(tidak_cocok))

            st.dataframe(
                provinsi,
                use_container_width=True,
                hide_index=True
            )

        if not kabupaten.empty:

            cocok = (
                kabupaten["Status Nama Boundary"] == "Cocok"
            ).sum()

            tidak_cocok = (
                kabupaten["Status Nama Boundary"] == "Tidak cocok"
            ).sum()

            st.subheader("Matching Kabupaten/Kota")

            col1, col2 = st.columns(2)

            col1.metric("Cocok", int(cocok))

            col2.metric("Tidak cocok", int(tidak_cocok))

            st.dataframe(
                kabupaten,
                use_container_width=True,
                hide_index=True
            )

        if not file_error.empty:

            st.subheader("Peringatan Pembacaan File")

            st.dataframe(
                file_error,
                use_container_width=True,
                hide_index=True
            )

        # ----------------------------------------------------
        # DOWNLOAD LAPORAN MATCHING
        # ----------------------------------------------------

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            provinsi.to_excel(
                writer,
                sheet_name="Matching Provinsi",
                index=False
            )

            kabupaten.to_excel(
                writer,
                sheet_name="Matching Kabupaten",
                index=False
            )

            file_error.to_excel(
                writer,
                sheet_name="Peringatan",
                index=False
            )

        st.download_button(
            "Download Laporan Matching",
            data=output.getvalue(),
            file_name="Laporan_Matching_Tahap_2.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True
        )

    st.stop()


# ============================================================
# 7. TAHAP SELANJUTNYA
# ============================================================

if TAHAP in [
    "3. Validasi Indikator",
    "4. Dashboard Peta",
    "5. Ekspor Hasil"
]:

    st.title(TAHAP)

    st.info(
        "Tahap ini belum diaktifkan. "
        "Selesaikan dan verifikasi Tahap 2 terlebih dahulu."
    )

    st.stop()


# ============================================================
# JIKA TAHAP 1 DIPILIH
# KODE TAHAP 1 LAMA AKAN DILANJUTKAN DI BAWAH INI
# ============================================================
