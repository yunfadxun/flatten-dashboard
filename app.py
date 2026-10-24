# ============================================================
# STREAMLIT APP: FLATTEN HEADER MULTIPLE XLSM
# ============================================================

from pathlib import Path
from io import BytesIO
import re
import zipfile

import pandas as pd
import streamlit as st

from openpyxl.utils import get_column_letter


# ============================================================
# 1. KONFIGURASI STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Flatten Header XLSM",
    page_icon="📊",
    layout="wide"
)

st.title("Flatten Header Multiple XLSM")
st.write(
    "Upload beberapa file XLSM untuk meratakan header "
    "pada sheet Tabel Analisis dan Tabel Analisis Pusat by Kab."
)


# ============================================================
# 2. KONFIGURASI PEMROSESAN
# ============================================================

SHEET_ANALISIS = "Tabel Analisis"
SHEET_KAKO = "Tabel Analisis Pusat by Kab"

HEADER_START = 8
HEADER_END = 9
DATA_START = 10

# Batas ukuran upload per file dapat diatur di sini.
# Satuan MB.
MAX_FILE_SIZE_MB = 50


# ============================================================
# 3. FUNCTION: NORMALISASI NAMA SHEET
# ============================================================

def normalisasi_nama_sheet(nama):
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(nama).lower()
    )


# ============================================================
# 4. FUNCTION: CARI SHEET
# ============================================================

def cari_sheet(sheet_names, target_sheet):

    target_normal = normalisasi_nama_sheet(
        target_sheet
    )

    # Cari nama sheet yang sama setelah normalisasi.
    for nama in sheet_names:

        if normalisasi_nama_sheet(nama) == target_normal:
            return nama

    # Cari nama target sebagai bagian dari nama sheet.
    for nama in sheet_names:

        if target_normal in normalisasi_nama_sheet(nama):
            return nama

    return None


# ============================================================
# 5. FUNCTION: FLATTEN MULTI-ROW HEADER
# ============================================================

def flatten_excel_header(
    excel_file,
    sheet_name,
    header_start,
    header_end,
    data_start=None
):

    # Baca sheet tanpa menetapkan header.
    raw = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )

    if raw.empty:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak memiliki data."
        )

    if header_end >= len(raw):
        raise ValueError(
            f"Baris header {header_end + 1} tidak ditemukan "
            f"pada sheet '{sheet_name}'."
        )

    # Ambil baris header.
    header_rows = raw.iloc[
        header_start:header_end + 1
    ].copy()

    # Normalisasi isi header.
    header_rows = header_rows.map(
        lambda x: (
            str(x).strip()
            if pd.notna(x) and str(x).strip() != ""
            else None
        )
    )

    # Isi header kelompok secara horizontal.
    header_rows = header_rows.ffill(axis=1)

    # Gabungkan header setiap kolom.
    new_columns = []

    for col_idx in range(raw.shape[1]):

        parts = []

        for row_idx in range(len(header_rows)):

            value = header_rows.iloc[
                row_idx, col_idx
            ]

            if value is not None:

                value = str(value).strip()

                # Hindari pengulangan nama header
                # yang sama secara berurutan.
                if value and (
                    not parts or value != parts[-1]
                ):
                    parts.append(value)

        if parts:
            col_name = "_".join(parts)
        else:
            col_name = f"Kolom_{col_idx + 1}"

        new_columns.append(col_name)

    # Pastikan semua nama kolom unik.
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

    # Tentukan awal data.
    if data_start is None:
        data_start = header_end + 1

    # Ambil data setelah header.
    df = raw.iloc[data_start:].copy()

    # Terapkan nama kolom.
    df.columns = unique_columns

    # Hapus baris kosong.
    df = df.dropna(how="all")

    # Hapus kolom yang seluruh datanya kosong.
    df = df.dropna(axis=1, how="all")

    return df.reset_index(drop=True)


# ============================================================
# 6. FUNCTION: FORMAT EXCEL OUTPUT
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

            # Atur lebar kolom.
            for col_idx, col_cells in enumerate(
                ws.iter_cols(),
                start=1
            ):

                max_length = max(
                    (
                        len(str(cell.value))
                        for cell in col_cells
                        if cell.value is not None
                    ),
                    default=0
                )

                col_letter = get_column_letter(col_idx)

                ws.column_dimensions[
                    col_letter
                ].width = min(max_length + 2, 45)

    output.seek(0)

    return output.getvalue()


# ============================================================
# 7. FUNCTION: PROSES SATU FILE XLSM
# ============================================================

def proses_satu_file(file_bytes, nama_file):

    hasil_sheet = {}
    log = []

    # Baca nama sheet dari file yang di-upload.
    excel = pd.ExcelFile(
        BytesIO(file_bytes),
        engine="openpyxl"
    )

    sheet_names = excel.sheet_names

    # --------------------------------------------------------
    # PROSES TABEL ANALISIS
    # --------------------------------------------------------

    sheet_analisis = cari_sheet(
        sheet_names,
        SHEET_ANALISIS
    )

    if sheet_analisis:

        df_analisis = flatten_excel_header(
            excel_file=BytesIO(file_bytes),
            sheet_name=sheet_analisis,
            header_start=HEADER_START,
            header_end=HEADER_END,
            data_start=DATA_START
        )

        hasil_sheet["Tabel Analisis"] = df_analisis

        log.append({
            "Sheet": "Tabel Analisis",
            "Status": "Berhasil",
            "Baris": len(df_analisis),
            "Kolom": len(df_analisis.columns),
            "Keterangan": "Header berhasil di-flatten."
        })

    else:

        log.append({
            "Sheet": "Tabel Analisis",
            "Status": "Tidak ditemukan",
            "Baris": 0,
            "Kolom": 0,
            "Keterangan": "Sheet tidak ditemukan."
        })

    # --------------------------------------------------------
    # PROSES TABEL ANALISIS PUSAT BY KAB
    # --------------------------------------------------------

    sheet_kako = cari_sheet(
        sheet_names,
        SHEET_KAKO
    )

    if sheet_kako:

        df_kako = flatten_excel_header(
            excel_file=BytesIO(file_bytes),
            sheet_name=sheet_kako,
            header_start=HEADER_START,
            header_end=HEADER_END,
            data_start=DATA_START
        )

        hasil_sheet["Analisis Pusat by Kab"] = df_kako

        log.append({
            "Sheet": "Tabel Analisis Pusat by Kab",
            "Status": "Berhasil",
            "Baris": len(df_kako),
            "Kolom": len(df_kako.columns),
            "Keterangan": "Header berhasil di-flatten."
        })

    else:

        log.append({
            "Sheet": "Tabel Analisis Pusat by Kab",
            "Status": "Tidak ditemukan",
            "Baris": 0,
            "Kolom": 0,
            "Keterangan": "Sheet tidak ditemukan."
        })

    # Validasi.
    if not hasil_sheet:
        raise ValueError(
            "Kedua sheet target tidak ditemukan."
        )

    # Buat file Excel di memori.
    output_bytes = dataframe_to_excel(
        hasil_sheet
    )

    nama_output = (
        f"{Path(nama_file).stem}_Flatten_Header.xlsx"
    )

    return {
        "nama_output": nama_output,
        "data": output_bytes,
        "log": log,
        "jumlah_sheet": len(hasil_sheet)
    }


# ============================================================
# 8. UPLOAD FILE
# ============================================================

uploaded_files = st.file_uploader(
    "Upload file XLSM",
    type=["xlsm"],
    accept_multiple_files=True,
    help=(
        "Kamu bisa mengunggah beberapa file XLSM sekaligus. "
        f"Maksimum {MAX_FILE_SIZE_MB} MB per file."
    )
)


# ============================================================
# 9. TOMBOL PROSES
# ============================================================

if uploaded_files:

    st.subheader("File yang di-upload")

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

    file_terlalu_besar = [
        f.name
        for f in uploaded_files
        if f.size > MAX_FILE_SIZE_MB * 1024 * 1024
    ]

    if file_terlalu_besar:

        st.warning(
            f"Ada {len(file_terlalu_besar)} file "
            "yang melebihi batas ukuran."
        )

    if st.button(
        "Proses dan Flatten Header",
        type="primary",
        use_container_width=True
    ):

        file_valid = [
            f for f in uploaded_files
            if f.size <= MAX_FILE_SIZE_MB * 1024 * 1024
        ]

        if not file_valid:

            st.error(
                "Tidak ada file yang memenuhi batas ukuran."
            )

        else:

            hasil_berhasil = []
            hasil_gagal = []

            progress = st.progress(0)

            status_text = st.empty()

            # ------------------------------------------------
            # PROSES SELURUH FILE
            # ------------------------------------------------

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

                    hasil["nama_input"] = file.name

                    hasil_berhasil.append(hasil)

                except Exception as e:

                    hasil_gagal.append({
                        "Nama File": file.name,
                        "Error": str(e)
                    })

                progress.progress(
                    nomor / len(file_valid)
                )

            status_text.write("Pemrosesan selesai.")

            # Simpan hasil pada session state.
            st.session_state["hasil_flatten"] = (
                hasil_berhasil
            )

            st.session_state["hasil_flatten_gagal"] = (
                hasil_gagal
            )


# ============================================================
# 10. TAMPILKAN HASIL
# ============================================================

if "hasil_flatten" in st.session_state:

    hasil_berhasil = st.session_state[
        "hasil_flatten"
    ]

    hasil_gagal = st.session_state[
        "hasil_flatten_gagal"
    ]

    st.divider()

    st.subheader("Ringkasan Pemrosesan")

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
        "Total file diproses",
        len(hasil_berhasil) + len(hasil_gagal)
    )

    # --------------------------------------------------------
    # DETAIL HASIL
    # --------------------------------------------------------

    if hasil_berhasil:

        st.subheader("File Hasil Flatten")

        daftar_log = []

        for hasil in hasil_berhasil:

            for item in hasil["log"]:

                daftar_log.append({
                    "Nama File": hasil["nama_input"],
                    **item
                })

        st.dataframe(
            pd.DataFrame(daftar_log),
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # DOWNLOAD EXCEL PER FILE
        # ----------------------------------------------------

        st.subheader("Download Excel")

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
        # DOWNLOAD SELURUH FILE DALAM ZIP
        # ----------------------------------------------------

        zip_buffer = BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zipf:

            for hasil in hasil_berhasil:

                zipf.writestr(
                    hasil["nama_output"],
                    hasil["data"]
                )

        zip_buffer.seek(0)

        st.download_button(
            label="Download Semua Hasil (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Hasil_Flatten_Header.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

    # --------------------------------------------------------
    # FILE GAGAL
    # --------------------------------------------------------

    if hasil_gagal:

        st.subheader("File yang Gagal Diproses")

        st.dataframe(
            pd.DataFrame(hasil_gagal),
            use_container_width=True,
            hide_index=True
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
