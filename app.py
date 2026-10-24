import re
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Surveilans MR — Tahap 1", page_icon="🗺️", layout="wide")
st.title("Dashboard Surveilans Campak-Rubella")
st.subheader("Tahap 1 — Upload dan Flatten Header Excel")
st.info(
    "Versi ini sengaja hanya menjalankan flatten header terlebih dahulu. "
    "Proses boundary, matching, dan pembuatan peta belum dijalankan agar aplikasi Streamlit Cloud lebih ringan."
)

st.markdown(
    "**Format yang dibutuhkan:** 3 file Excel untuk tahun 2023, 2024, dan 2025. "
    "Pada setiap file, aplikasi mencari sheet `Tabel Analisis` dan `Tabel Analisis Pusat by Kab`, "
    "menggunakan header baris 9–10 (indeks Python 8–9) dan data mulai baris 11 (indeks Python 10)."
)

uploaded_files = st.file_uploader(
    "Unggah file Excel 2023, 2024, dan 2025",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Nama file harus memuat tahun 2023, 2024, atau 2025. File .xls tidak didukung pada versi ini."
)


def norm_sheet(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def find_sheet(sheet_names, target):
    normalized_target = norm_sheet(target)
    for name in sheet_names:
        if norm_sheet(name) == normalized_target:
            return name
    for name in sheet_names:
        if normalized_target in norm_sheet(name):
            return name
    return None


def flatten_sheet(file_bytes: bytes, sheet_name: str, header_start=8, header_end=9, data_start=10):
    raw = pd.read_excel(
        BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine="openpyxl"
    )
    if raw.empty:
        raise ValueError(f"Sheet '{sheet_name}' kosong.")
    if len(raw) <= header_end:
        raise ValueError(f"Sheet '{sheet_name}' tidak memiliki baris header 9–10.")

    header_rows = raw.iloc[header_start:header_end + 1].copy()
    header_rows = header_rows.applymap(
        lambda x: str(x).strip() if pd.notna(x) and str(x).strip() else None
    )
    header_rows = header_rows.ffill(axis=1)

    columns = []
    for col_idx in range(raw.shape[1]):
        parts = []
        for row_idx in range(len(header_rows)):
            value = header_rows.iloc[row_idx, col_idx]
            if value is not None:
                value = str(value).strip()
                if value and (not parts or value != parts[-1]):
                    parts.append(value)
        columns.append("_".join(parts) if parts else f"Kolom_{col_idx + 1}")

    seen = {}
    unique_columns = []
    for col in columns:
        if col not in seen:
            seen[col] = 0
            unique_columns.append(col)
        else:
            seen[col] += 1
            unique_columns.append(f"{col}_{seen[col]}")

    df = raw.iloc[data_start:].copy()
    df.columns = unique_columns
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def flatten_workbook(file_bytes: bytes, filename: str):
    xls = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    targets = [
        ("Tabel Analisis", "Tabel Analisis"),
        ("Tabel Analisis Pusat by Kab", "Analisis Pusat by Kab"),
    ]
    sheets = {}
    logs = []
    for target, output_name in targets:
        source_sheet = find_sheet(xls.sheet_names, target)
        if source_sheet is None:
            logs.append({"File": filename, "Sheet": target, "Status": "Tidak ditemukan", "Baris": 0, "Kolom": 0})
            continue
        df = flatten_sheet(file_bytes, source_sheet)
        sheets[output_name] = df
        logs.append({"File": filename, "Sheet": target, "Status": "Berhasil", "Baris": len(df), "Kolom": len(df.columns)})

    if not sheets:
        raise ValueError("Tidak ditemukan sheet 'Tabel Analisis' maupun 'Tabel Analisis Pusat by Kab'.")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
            ws = writer.sheets[safe_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            # Batasi lebar kolom tanpa memindai seluruh worksheet berulang kali.
            for idx, col_name in enumerate(df.columns, start=1):
                sample_values = df[col_name].head(200).tolist()
                max_len = max([len(str(col_name))] + [len(str(v)) for v in sample_values if pd.notna(v)])
                ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 40)
    return output.getvalue(), logs


if st.button("Jalankan Tahap 1 — Flatten Header", type="primary", disabled=not uploaded_files):
    if len(uploaded_files) != 3:
        st.error(f"Unggah tepat 3 file: 2023, 2024, dan 2025. Saat ini ada {len(uploaded_files)} file.")
        st.stop()

    years = {}
    for f in uploaded_files:
        matches = re.findall(r"202[345]", f.name)
        if not matches:
            st.error(f"Nama file '{f.name}' tidak memuat tahun 2023, 2024, atau 2025.")
            st.stop()
        year = matches[0]
        if year in years:
            st.error(f"Ada lebih dari satu file untuk tahun {year}: '{years[year]}' dan '{f.name}'.")
            st.stop()
        years[year] = f.name
    missing = sorted(set(["2023", "2024", "2025"]) - set(years))
    if missing:
        st.error("File untuk tahun berikut belum ada: " + ", ".join(missing))
        st.stop()

    zip_buffer = BytesIO()
    all_logs = []
    errors = []
    progress = st.progress(0, text="Menyiapkan proses...")
    ordered_files = sorted(uploaded_files, key=lambda f: re.findall(r"202[345]", f.name)[0])

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, uploaded in enumerate(ordered_files, start=1):
            progress.progress((i - 1) / len(ordered_files), text=f"Memproses {uploaded.name} ({i}/{len(ordered_files)})...")
            try:
                flattened_bytes, logs = flatten_workbook(uploaded.getvalue(), uploaded.name)
                out_name = re.sub(r"\.(xlsx|xlsm)$", "", uploaded.name, flags=re.I) + "_Flatten_Header.xlsx"
                zf.writestr(out_name, flattened_bytes)
                all_logs.extend(logs)
            except Exception as exc:
                errors.append({"File": uploaded.name, "Error": str(exc)})
            progress.progress(i / len(ordered_files), text=f"Selesai memeriksa {uploaded.name} ({i}/{len(ordered_files)})")

    st.session_state["flatten_logs"] = all_logs
    st.session_state["flatten_errors"] = errors
    st.session_state["flatten_zip"] = zip_buffer.getvalue() if not errors else None

if "flatten_logs" in st.session_state:
    st.markdown("### Hasil Tahap 1")
    if st.session_state.get("flatten_logs"):
        st.dataframe(pd.DataFrame(st.session_state["flatten_logs"]), use_container_width=True, hide_index=True)
    if st.session_state.get("flatten_errors"):
        st.error("Ada file yang gagal diproses. Periksa detail berikut sebelum melanjutkan.")
        st.dataframe(pd.DataFrame(st.session_state["flatten_errors"]), use_container_width=True, hide_index=True)
    elif st.session_state.get("flatten_zip"):
        st.success("Flatten selesai untuk ketiga file. Unduh dan periksa hasilnya sebelum masuk ke tahap matching.")
        st.download_button(
            "Unduh ZIP hasil flatten",
            data=st.session_state["flatten_zip"],
            file_name="Hasil_Flatten_Header_MR.zip",
            mime="application/zip",
            type="primary",
        )

st.caption("Tahap berikutnya (boundary, matching, validasi indikator, dan peta) akan ditambahkan setelah hasil flatten ini terverifikasi.")
