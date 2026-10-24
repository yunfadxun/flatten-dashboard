# ============================================================
# DASHBOARD SURVEILANS CAMPAK-RUBELLA
# app.py
#
# TAHAP 1:
# - Upload Excel
# - Flatten header
# - Simpan hasil ke session_state
#
# TAHAP 2-3:
# - Dipanggil dari streamlit_tahap_2_3.py
# ============================================================

import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from openpyxl.utils import get_column_letter

from streamlit_tahap_2_3 import (
    init_state_tahap_2_3,
    halaman_tahap_2,
    halaman_tahap_3,
)


# ============================================================
# 1. KONFIGURASI HALAMAN
# ============================================================

st.set_page_config(
    page_title="Dashboard Surveilans Campak-Rubella",
    page_icon="🦠",
    layout="wide"
)


# ============================================================
# 2. INISIALISASI SESSION STATE
# ============================================================

init_state_tahap_2_3()


if "hasil_flatten" not in st.session_state:
    st.session_state["hasil_flatten"] = []


if "hasil_flatten_gagal" not in st.session_state:
    st.session_state["hasil_flatten_gagal"] = []


# ============================================================
# 3. KONFIGURASI TAHAP 1
# ============================================================

SHEET_TARGET = [
    "Tabel Analisis",
    "Tabel Analisis Pusat by Kab",
]

# Pandas menggunakan indeks mulai dari 0.
#
# 8  = Excel baris 9
# 9  = Excel baris 10
# 10 = Excel baris 11

BARIS_HEADER_1 = 8
BARIS_HEADER_2 = 9
BARIS_AWAL_DATA = 10

MAX_FILE_SIZE_MB = 50


# ============================================================
# 4. FUNGSI NORMALISASI NAMA SHEET
# ============================================================

def normalisasi_nama_sheet(nama):
    """
    Normalisasi nama sheet.

    Contoh:
    'Tabel Analisis'
    'tabel analisis'
    'Tabel  Analisis'

    akan lebih mudah dicocokkan.
    """

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(nama).lower()
    )


def cari_sheet(
    daftar_sheet_excel,
    nama_target
):
    """
    Cari sheet dengan urutan:

    1. Exact match setelah normalisasi.
    2. Partial match jika exact tidak ditemukan.
    """

    target = normalisasi_nama_sheet(
        nama_target
    )

    # --------------------------------------------------------
    # EXACT MATCH
    # --------------------------------------------------------

    for nama_sheet in daftar_sheet_excel:

        if (
            normalisasi_nama_sheet(
                nama_sheet
            )
            == target
        ):
            return nama_sheet

    # --------------------------------------------------------
    # PARTIAL MATCH
    # --------------------------------------------------------

    for nama_sheet in daftar_sheet_excel:

        nama_normal = (
            normalisasi_nama_sheet(
                nama_sheet
            )
        )

        if target in nama_normal:
            return nama_sheet

    return None


# ============================================================
# 5. MEMBERSIHKAN HEADER
# ============================================================

def bersihkan_teks(nilai):
    """
    Membersihkan isi header Excel.
    """

    if pd.isna(nilai):
        return ""

    teks = str(nilai).strip()

    teks = re.sub(
        r"[\r\n\t]+",
        " ",
        teks
    )

    teks = re.sub(
        r"\s+",
        " ",
        teks
    )

    return teks.strip()


# ============================================================
# 6. MEMBUAT NAMA KOLOM
# ============================================================

def buat_nama_kolom(
    header_1,
    header_2
):
    """
    Gabungkan header tingkat 1 dan header tingkat 2.

    Contoh:

    Performa Surveilans
    Measles Incidence Rate

    menjadi:

    Performa Surveilans_Measles Incidence Rate
    """

    h1 = bersihkan_teks(
        header_1
    )

    h2 = bersihkan_teks(
        header_2
    )

    if h1 and h2:

        if h1.casefold() == h2.casefold():
            nama = h1

        else:
            nama = f"{h1}_{h2}"

    elif h1:

        nama = h1

    elif h2:

        nama = h2

    else:

        nama = "Kolom_Tanpa_Nama"

    return nama.strip("_")


# ============================================================
# 7. MEMASTIKAN NAMA KOLOM UNIK
# ============================================================

def pastikan_nama_kolom_unik(
    daftar_kolom
):
    """
    Contoh:

    Kabupaten
    Kabupaten
    Kabupaten

    menjadi:

    Kabupaten
    Kabupaten_2
    Kabupaten_3
    """

    hasil = []
    penghitung = {}

    for nama in daftar_kolom:

        if not nama:
            nama = "Kolom_Tanpa_Nama"

        if nama not in penghitung:

            penghitung[nama] = 1
            hasil.append(nama)

        else:

            penghitung[nama] += 1

            hasil.append(
                f"{nama}_"
                f"{penghitung[nama]}"
            )

    return hasil


# ============================================================
# 8. FLATTEN HEADER EXCEL
# ============================================================

def flatten_excel_header(
    file_bytes,
    sheet_name
):
    """
    Struktur workbook:

    Excel baris 9:
        header tingkat pertama

    Excel baris 10:
        header tingkat kedua

    Excel baris 11:
        data pertama
    """

    df_raw = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )

    if df_raw.empty:

        raise ValueError(
            f"Sheet '{sheet_name}' kosong."
        )

    if len(df_raw) <= BARIS_HEADER_2:

        raise ValueError(
            f"Sheet '{sheet_name}' tidak "
            "memiliki header pada baris "
            "Excel 9–10."
        )

    # --------------------------------------------------------
    # AMBIL HEADER
    # --------------------------------------------------------

    header_df = df_raw.iloc[
        BARIS_HEADER_1:
        BARIS_HEADER_2 + 1
    ].copy()

    # --------------------------------------------------------
    # BERSIHKAN HEADER
    # --------------------------------------------------------

    header_df = header_df.map(
        lambda x: (
            bersihkan_teks(x)
            if pd.notna(x)
            else ""
        )
    )

    # --------------------------------------------------------
    # PENTING:
    # Forward-fill header tingkat pertama.
    #
    # Ini dibutuhkan karena merged cells Excel biasanya
    # hanya mempunyai nilai pada kolom pertama.
    # --------------------------------------------------------

    header_1 = (
        header_df.iloc[0]
        .replace("", pd.NA)
        .ffill()
        .fillna("")
        .tolist()
    )

    header_2 = (
        header_df.iloc[1]
        .fillna("")
        .tolist()
    )

    # --------------------------------------------------------
    # BUAT NAMA KOLOM
    # --------------------------------------------------------

    nama_kolom = []

    for h1, h2 in zip(
        header_1,
        header_2
    ):

        nama = buat_nama_kolom(
            h1,
            h2
        )

        nama_kolom.append(
            nama
        )

    nama_kolom = (
        pastikan_nama_kolom_unik(
            nama_kolom
        )
    )

    # --------------------------------------------------------
    # AMBIL DATA
    # --------------------------------------------------------

    df_data = df_raw.iloc[
        BARIS_AWAL_DATA:
    ].copy()

    df_data.columns = nama_kolom

    # --------------------------------------------------------
    # STRING KOSONG → NA
    # --------------------------------------------------------

    df_data = df_data.replace(
        r"^\s*$",
        pd.NA,
        regex=True
    )

    # --------------------------------------------------------
    # HAPUS BARIS SEPENUHNYA KOSONG
    # --------------------------------------------------------

    df_data = df_data.dropna(
        how="all"
    )

    # --------------------------------------------------------
    # HAPUS KOLOM SEPENUHNYA KOSONG
    #
    # Tetapi nama header yang valid dipertahankan apabila
    # memiliki data.
    # --------------------------------------------------------

    df_data = df_data.dropna(
        axis=1,
        how="all"
    )

    df_data = df_data.reset_index(
        drop=True
    )

    return df_data


# ============================================================
# 9. FORMAT WORKSHEET OUTPUT
# ============================================================

def format_worksheet(
    worksheet,
    df
):
    """
    Format dasar hasil flatten.
    """

    worksheet.freeze_panes = "A2"

    if len(df.columns) > 0:

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

    for kolom_idx, nama_kolom in enumerate(
        df.columns,
        start=1
    ):

        sample_values = (
            df[nama_kolom]
            .head(200)
        )

        panjang = [
            len(str(nama_kolom))
        ]

        panjang.extend([
            len(str(value))
            for value in sample_values
            if pd.notna(value)
        ])

        max_length = max(
            panjang
        )

        lebar = min(
            max(
                max_length + 2,
                12
            ),
            40
        )

        huruf_kolom = (
            get_column_letter(
                kolom_idx
            )
        )

        worksheet.column_dimensions[
            huruf_kolom
        ].width = lebar


# ============================================================
# 10. PROSES SATU FILE EXCEL
# ============================================================

def proses_file_excel(
    file_bytes,
    nama_file
):
    """
    Memproses seluruh sheet target.

    Output sengaja menggunakan struktur yang kompatibel
    dengan streamlit_tahap_2_3.py.
    """

    hasil_sheet = {}

    sheet_berhasil = []

    sheet_tidak_ditemukan = []

    log = []

    # --------------------------------------------------------
    # BACA WORKBOOK
    # --------------------------------------------------------

    try:

        excel_file = pd.ExcelFile(
            io.BytesIO(file_bytes),
            engine="openpyxl"
        )

        daftar_sheet_excel = (
            excel_file.sheet_names
        )

    except Exception as error:

        raise ValueError(
            "Tidak dapat membaca file Excel. "
            f"Detail: {error}"
        )

    # --------------------------------------------------------
    # PROSES TARGET SHEET
    # --------------------------------------------------------

    for sheet_target in SHEET_TARGET:

        sheet_asli = cari_sheet(
            daftar_sheet_excel,
            sheet_target
        )

        if sheet_asli is None:

            sheet_tidak_ditemukan.append(
                sheet_target
            )

            log.append({
                "Nama File":
                    nama_file,

                "Sheet Target":
                    sheet_target,

                "Sheet Ditemukan":
                    "",

                "Status":
                    "Tidak ditemukan",

                "Baris":
                    0,

                "Kolom":
                    0,
            })

            continue

        try:

            df_hasil = (
                flatten_excel_header(
                    file_bytes=file_bytes,
                    sheet_name=sheet_asli
                )
            )

            hasil_sheet[
                sheet_target
            ] = df_hasil

            sheet_berhasil.append(
                sheet_target
            )

            log.append({
                "Nama File":
                    nama_file,

                "Sheet Target":
                    sheet_target,

                "Sheet Ditemukan":
                    sheet_asli,

                "Status":
                    "Berhasil",

                "Baris":
                    len(df_hasil),

                "Kolom":
                    len(df_hasil.columns),
            })

        except Exception as error:

            log.append({
                "Nama File":
                    nama_file,

                "Sheet Target":
                    sheet_target,

                "Sheet Ditemukan":
                    sheet_asli,

                "Status":
                    "Gagal",

                "Baris":
                    0,

                "Kolom":
                    0,

                "Error":
                    str(error),
            })

    # --------------------------------------------------------
    # VALIDASI
    # --------------------------------------------------------

    if not hasil_sheet:

        raise ValueError(
            "Tidak ada sheet target yang "
            "berhasil diproses. "
            "Sheet tersedia: "
            + ", ".join(
                daftar_sheet_excel
            )
        )

    # --------------------------------------------------------
    # BUAT WORKBOOK HASIL
    # --------------------------------------------------------

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for (
            nama_sheet,
            df_hasil
        ) in hasil_sheet.items():

            # ----------------------------------------------
            # PENTING:
            #
            # streamlit_tahap_2_3.py mengharapkan:
            #
            # Tabel Analisis
            # Analisis Pusat by Kab
            #
            # Jadi sheet kedua harus diganti namanya.
            # ----------------------------------------------

            if (
                nama_sheet
                == "Tabel Analisis Pusat by Kab"
            ):

                nama_sheet_output = (
                    "Analisis Pusat by Kab"
                )

            else:

                nama_sheet_output = (
                    nama_sheet
                )

            nama_sheet_output = (
                nama_sheet_output[:31]
            )

            df_hasil.to_excel(
                writer,
                sheet_name=nama_sheet_output,
                index=False
            )

            worksheet = (
                writer.sheets[
                    nama_sheet_output
                ]
            )

            format_worksheet(
                worksheet,
                df_hasil
            )

    output.seek(0)

    excel_bytes = output.getvalue()

    # --------------------------------------------------------
    # FORMAT OUTPUT UNTUK TAHAP 2
    # --------------------------------------------------------

    return {

        # Format yang dipakai UI Tahap 1
        "nama_file":
            nama_file,

        "excel_bytes":
            excel_bytes,

        "data":
            hasil_sheet,

        "sheet_berhasil":
            sheet_berhasil,

        "sheet_tidak_ditemukan":
            sheet_tidak_ditemukan,

        # Format yang dibaca module Tahap 2–3
        "nama_input":
            nama_file,

        "nama_output":
            buat_nama_file_output(
                nama_file
            ),

        "log":
            log,
    }


# ============================================================
# 11. NAMA FILE OUTPUT
# ============================================================

def buat_nama_file_output(
    nama_file
):
    """
    Membuat nama output yang aman.
    """

    nama_dasar = Path(
        nama_file
    ).stem

    nama_dasar = re.sub(
        r'[\\/*?:"<>|]',
        "_",
        nama_dasar
    )

    return (
        f"{nama_dasar}_"
        f"Flattened.xlsx"
    )


# ============================================================
# 12. MEMBUAT ZIP
# ============================================================

def buat_zip_hasil(
    daftar_hasil
):
    """
    Gabungkan seluruh workbook hasil menjadi ZIP.
    """

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED
    ) as zip_file:

        nama_dipakai = set()

        for hasil in daftar_hasil:

            nama_output = (
                hasil["nama_output"]
            )

            nama_awal = Path(
                nama_output
            ).stem

            ekstensi = Path(
                nama_output
            ).suffix

            penghitung = 2

            while (
                nama_output
                in nama_dipakai
            ):

                nama_output = (
                    f"{nama_awal}_"
                    f"{penghitung}"
                    f"{ekstensi}"
                )

                penghitung += 1

            nama_dipakai.add(
                nama_output
            )

            zip_file.writestr(
                nama_output,
                hasil["excel_bytes"]
            )

    zip_buffer.seek(0)

    return zip_buffer.getvalue()


# ============================================================
# 13. RESET TAHAP BERIKUTNYA
# ============================================================

def reset_tahap_berikutnya():
    """
    Jika Tahap 1 diproses ulang, hasil Tahap 2 dan 3 lama
    tidak boleh digunakan lagi.
    """

    st.session_state[
        "hasil_boundary"
    ] = None

    st.session_state[
        "hasil_matching"
    ] = None

    st.session_state[
        "hasil_validasi_indikator"
    ] = None

    st.session_state[
        "payload_dashboard"
    ] = None


# ============================================================
# 14. HALAMAN TAHAP 1
# ============================================================

def halaman_tahap_1():

    st.title(
        "Dashboard Surveilans Campak-Rubella"
    )

    st.markdown(
        "### Tahap 1 — Upload dan Flatten Header Excel"
    )

    st.write(
        """
        Unggah satu atau beberapa file Excel surveilans
        campak-rubella.

        Aplikasi akan merapikan header baris Excel 9–10
        dan mengambil data mulai dari baris Excel 11.
        """
    )

    st.info(
        "Nama file sebaiknya mengandung tahun 4 digit, "
        "misalnya `MR02_Nas_2023.xlsx`, "
        "`MR02_Nas_2024.xlsx`, atau "
        "`MR02_Nas_2025.xlsx`. "
        "Tahap 2 menggunakan tahun pada nama file."
    )

    st.markdown(
        """
        **Sheet yang diproses:**

        1. `Tabel Analisis`
        2. `Tabel Analisis Pusat by Kab`
        """
    )

    st.divider()

    # ========================================================
    # UPLOAD
    # ========================================================

    st.subheader(
        "1. Upload File Excel"
    )

    uploaded_files = st.file_uploader(
        "Pilih satu atau beberapa file Excel",
        type=[
            "xlsx",
            "xlsm",
        ],
        accept_multiple_files=True,
        help=(
            f"Ukuran maksimum setiap file "
            f"{MAX_FILE_SIZE_MB} MB."
        ),
        key="upload_excel_tahap_1"
    )

    # ========================================================
    # INFORMASI FILE
    # ========================================================

    if uploaded_files:

        daftar_upload = []

        for file in uploaded_files:

            ukuran_mb = (
                file.size
                / (1024 * 1024)
            )

            daftar_upload.append({
                "Nama File":
                    file.name,

                "Ukuran (MB)":
                    round(
                        ukuran_mb,
                        2
                    ),

                "Status":
                    (
                        "Siap diproses"
                        if ukuran_mb
                        <= MAX_FILE_SIZE_MB
                        else
                        "Melebihi batas"
                    )
            })

        st.dataframe(
            pd.DataFrame(
                daftar_upload
            ),
            use_container_width=True,
            hide_index=True
        )

    # ========================================================
    # TOMBOL PROSES
    # ========================================================

    proses = st.button(
        "Proses File Excel",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files,
        key="proses_excel_tahap_1"
    )

    if proses:

        # ----------------------------------------------------
        # VALIDASI UKURAN
        # ----------------------------------------------------

        file_valid = [
            file
            for file in uploaded_files
            if (
                file.size
                <= MAX_FILE_SIZE_MB
                * 1024
                * 1024
            )
        ]

        file_besar = [
            file.name
            for file in uploaded_files
            if (
                file.size
                > MAX_FILE_SIZE_MB
                * 1024
                * 1024
            )
        ]

        if file_besar:

            st.warning(
                "File berikut tidak diproses "
                "karena terlalu besar: "
                + ", ".join(
                    file_besar
                )
            )

        if not file_valid:

            st.error(
                "Tidak ada file yang dapat diproses."
            )

            return

        # ----------------------------------------------------
        # PROSES
        # ----------------------------------------------------

        daftar_hasil = []
        daftar_error = []

        progress = st.progress(0)

        status = st.empty()

        for i, file in enumerate(
            file_valid,
            start=1
        ):

            status.info(
                f"Memproses "
                f"{file.name} "
                f"({i}/{len(file_valid)})..."
            )

            try:

                hasil = proses_file_excel(
                    file_bytes=(
                        file.getvalue()
                    ),
                    nama_file=(
                        file.name
                    )
                )

                daftar_hasil.append(
                    hasil
                )

            except Exception as error:

                daftar_error.append({

                    "Nama File":
                        file.name,

                    "Jenis Error":
                        type(
                            error
                        ).__name__,

                    "Detail Error":
                        str(
                            error
                        ),
                })

            progress.progress(
                i / len(file_valid)
            )

        # ----------------------------------------------------
        # SIMPAN SESSION STATE
        #
        # INI YANG DIBACA TAHAP 2
        # ----------------------------------------------------

        st.session_state[
            "hasil_flatten"
        ] = daftar_hasil

        st.session_state[
            "hasil_flatten_gagal"
        ] = daftar_error

        # Hapus hasil Tahap 2–3 lama.
        reset_tahap_berikutnya()

        status.success(
            "Pemrosesan selesai."
        )

        st.rerun()

    # ========================================================
    # AMBIL HASIL SESSION
    # ========================================================

    daftar_hasil = (
        st.session_state.get(
            "hasil_flatten",
            []
        )
    )

    daftar_error = (
        st.session_state.get(
            "hasil_flatten_gagal",
            []
        )
    )

    # ========================================================
    # TAMPILKAN HASIL
    # ========================================================

    if daftar_hasil:

        st.divider()

        st.success(
            f"Berhasil memproses "
            f"{len(daftar_hasil)} "
            f"file Excel."
        )

        # ----------------------------------------------------
        # METRIC
        # ----------------------------------------------------

        col1, col2, col3 = (
            st.columns(3)
        )

        col1.metric(
            "File berhasil",
            len(daftar_hasil)
        )

        col2.metric(
            "File gagal",
            len(daftar_error)
        )

        col3.metric(
            "Total file",
            (
                len(daftar_hasil)
                + len(daftar_error)
            )
        )

        # ----------------------------------------------------
        # RINGKASAN
        # ----------------------------------------------------

        st.subheader(
            "2. Ringkasan Pemrosesan"
        )

        data_ringkasan = []

        for hasil in daftar_hasil:

            jumlah_baris = sum(
                len(df)
                for df
                in hasil[
                    "data"
                ].values()
            )

            data_ringkasan.append({

                "Nama File":
                    hasil[
                        "nama_file"
                    ],

                "Sheet Berhasil":
                    len(
                        hasil[
                            "sheet_berhasil"
                        ]
                    ),

                "Total Baris Data":
                    jumlah_baris,

                "Status":
                    "Berhasil",
            })

        st.dataframe(
            pd.DataFrame(
                data_ringkasan
            ),
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # LOG
        # ----------------------------------------------------

        logs = []

        for hasil in daftar_hasil:

            logs.extend(
                hasil["log"]
            )

        if logs:

            with st.expander(
                "Lihat detail proses"
            ):

                st.dataframe(
                    pd.DataFrame(
                        logs
                    ),
                    use_container_width=True,
                    hide_index=True
                )

        # ----------------------------------------------------
        # DOWNLOAD ZIP
        # ----------------------------------------------------

        st.subheader(
            "3. Download Hasil"
        )

        zip_bytes = (
            buat_zip_hasil(
                daftar_hasil
            )
        )

        st.download_button(
            label=(
                "Download Semua Hasil (ZIP)"
            ),
            data=zip_bytes,
            file_name=(
                "Hasil_Flatten_Header_"
                "Surveilans_MR.zip"
            ),
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key="download_zip_tahap_1"
        )

        # ----------------------------------------------------
        # DOWNLOAD PER FILE
        # ----------------------------------------------------

        st.markdown(
            "#### Download Per File"
        )

        for i, hasil in enumerate(
            daftar_hasil
        ):

            with st.container(
                border=True
            ):

                col1, col2 = (
                    st.columns(
                        [3, 1]
                    )
                )

                with col1:

                    st.write(
                        f"**"
                        f"{hasil['nama_file']}"
                        f"**"
                    )

                    st.caption(
                        "Sheet berhasil: "
                        + ", ".join(
                            hasil[
                                "sheet_berhasil"
                            ]
                        )
                    )

                with col2:

                    st.download_button(
                        label=(
                            "Download Excel"
                        ),
                        data=(
                            hasil[
                                "excel_bytes"
                            ]
                        ),
                        file_name=(
                            hasil[
                                "nama_output"
                            ]
                        ),
                        mime=(
                            "application/"
                            "vnd.openxmlformats-"
                            "officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key=(
                            f"download_excel_"
                            f"{i}"
                        ),
                        use_container_width=True
                    )

        # ----------------------------------------------------
        # PREVIEW
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "4. Pratinjau Hasil"
        )

        nama_file_preview = (
            st.selectbox(
                "Pilih file",
                options=[
                    hasil[
                        "nama_file"
                    ]
                    for hasil
                    in daftar_hasil
                ],
                key=(
                    "preview_file_"
                    "tahap_1"
                )
            )
        )

        hasil_preview = next(
            hasil
            for hasil in daftar_hasil
            if (
                hasil[
                    "nama_file"
                ]
                == nama_file_preview
            )
        )

        daftar_sheet_preview = (
            list(
                hasil_preview[
                    "data"
                ].keys()
            )
        )

        sheet_preview = (
            st.selectbox(
                "Pilih sheet",
                options=(
                    daftar_sheet_preview
                ),
                key=(
                    "preview_sheet_"
                    "tahap_1"
                )
            )
        )

        df_preview = (
            hasil_preview[
                "data"
            ][
                sheet_preview
            ]
        )

        col1, col2 = (
            st.columns(2)
        )

        col1.metric(
            "Jumlah baris",
            f"{len(df_preview):,}"
        )

        col2.metric(
            "Jumlah kolom",
            f"{len(df_preview.columns):,}"
        )

        st.dataframe(
            df_preview.head(
                100
            ),
            use_container_width=True,
            height=450
        )

        if len(df_preview) > 100:

            st.caption(
                "Pratinjau dibatasi "
                "100 baris pertama."
            )

        # ----------------------------------------------------
        # STATUS NEXT STEP
        # ----------------------------------------------------

        st.success(
            "Tahap 1 selesai. "
            "Hasil sudah tersimpan dan siap "
            "digunakan oleh Tahap 2."
        )

    # ========================================================
    # ERROR
    # ========================================================

    if daftar_error:

        st.divider()

        st.subheader(
            "5. File yang Gagal Diproses"
        )

        st.dataframe(
            pd.DataFrame(
                daftar_error
            ),
            use_container_width=True,
            hide_index=True
        )

    # ========================================================
    # RESET TAHAP 1
    # ========================================================

    if daftar_hasil or daftar_error:

        st.divider()

        if st.button(
            "Reset Hasil Tahap 1",
            use_container_width=True,
            key="reset_tahap_1"
        ):

            st.session_state[
                "hasil_flatten"
            ] = []

            st.session_state[
                "hasil_flatten_gagal"
            ] = []

            reset_tahap_berikutnya()

            st.rerun()


# ============================================================
# 15. SIDEBAR
# ============================================================

st.sidebar.title(
    "Surveilans MR"
)

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


# ============================================================
# 16. STATUS SIDEBAR
# ============================================================

st.sidebar.divider()

if st.session_state.get(
    "hasil_flatten"
):

    st.sidebar.success(
        "Tahap 1 ✓"
    )

else:

    st.sidebar.info(
        "Tahap 1 belum selesai"
    )


if st.session_state.get(
    "hasil_matching"
) is not None:

    st.sidebar.success(
        "Tahap 2 ✓"
    )

else:

    st.sidebar.info(
        "Tahap 2 belum selesai"
    )


if st.session_state.get(
    "hasil_validasi_indikator"
) is not None:

    st.sidebar.success(
        "Tahap 3 ✓"
    )

else:

    st.sidebar.info(
        "Tahap 3 belum selesai"
    )


# ============================================================
# 17. ROUTING
# ============================================================

if TAHAP == "1. Upload & Flatten":

    halaman_tahap_1()


elif TAHAP == "2. Boundary & Matching":

    halaman_tahap_2()


elif TAHAP == "3. Validasi Indikator":

    halaman_tahap_3()


elif TAHAP == "4. Dashboard Peta":

    st.title(
        "Tahap 4 — Dashboard Peta"
    )

    if (
        st.session_state.get(
            "payload_dashboard"
        )
        is None
    ):

        st.warning(
            "Tahap 3 belum selesai. "
            "Jalankan Validasi Indikator "
            "terlebih dahulu."
        )

    else:

        st.success(
            "Boundary dan payload tersedia. "
            "Data siap digunakan untuk "
            "Dashboard Peta."
        )


elif TAHAP == "5. Ekspor Hasil":

    st.title(
        "Tahap 5 — Ekspor Hasil"
    )

    if (
        st.session_state.get(
            "payload_dashboard"
        )
        is None
    ):

        st.warning(
            "Tahap 3 belum selesai."
        )

    else:

        st.info(
            "Data Tahap 1–3 tersedia "
            "untuk proses ekspor."
        )
