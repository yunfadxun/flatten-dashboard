import streamlit as st
from pathlib import Path
import traceback
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Dashboard Peta Surveilans MR", page_icon="🗺️", layout="wide")
st.title("Dashboard Peta Surveilans Campak-Rubella")
st.write("Unggah file Excel 2023, 2024, dan 2025. Aplikasi akan flatten header baris 9–10 terlebih dahulu, lalu membuat dashboard peta surveilans, laporan matching, dan ZIP.")

uploaded_files = st.file_uploader(
    "Pilih 3 file Excel (2023, 2024, 2025)",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    help="Nama file harus memuat tahun 2023, 2024, atau 2025. Header pada dua sheet target akan di-flatten terlebih dahulu sebelum dashboard dibuat."
)

def _normalisasi_nama_sheet(nama):
    import re
    return re.sub(r"[^a-z0-9]", "", str(nama).lower())


def _cari_sheet_flatten(sheet_names, target_sheet):
    target = _normalisasi_nama_sheet(target_sheet)
    for nama in sheet_names:
        if _normalisasi_nama_sheet(nama) == target:
            return nama
    for nama in sheet_names:
        if target in _normalisasi_nama_sheet(nama):
            return nama
    return None


def _flatten_header_excel(file_bytes, sheet_name, header_start=8, header_end=9, data_start=10):
    raw = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )
    if raw.empty:
        raise ValueError(f"Sheet '{sheet_name}' kosong.")
    if header_end >= len(raw):
        raise ValueError(f"Baris header yang diperlukan tidak ditemukan pada '{sheet_name}'.")

    header_rows = raw.iloc[header_start:header_end + 1].copy()
    header_rows = header_rows.map(
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


def _flatten_file_for_dashboard(file_bytes, filename):
    """Flatten kedua sheet target dan hasilkan workbook XLSX di memori."""
    from io import BytesIO
    import pandas as pd

    xls = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    hasil_sheet = {}
    logs = []

    targets = [
        ("Tabel Analisis", "Tabel Analisis"),
        ("Tabel Analisis Pusat by Kab", "Analisis Pusat by Kab"),
    ]

    for target, output_name in targets:
        found = _cari_sheet_flatten(xls.sheet_names, target)
        if found is None:
            logs.append(f"{target}: tidak ditemukan")
            continue
        df = _flatten_header_excel(
            file_bytes=file_bytes,
            sheet_name=found,
            header_start=8,
            header_end=9,
            data_start=10,
        )
        hasil_sheet[output_name] = df
        logs.append(f"{target}: {len(df):,} baris × {len(df.columns):,} kolom")

    if not hasil_sheet:
        raise ValueError("Kedua sheet target tidak ditemukan: Tabel Analisis dan Tabel Analisis Pusat by Kab.")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in hasil_sheet.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            ws = writer.sheets[sheet_name[:31]]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.utils import get_column_letter
            for col_idx, cells in enumerate(ws.iter_cols(), start=1):
                max_len = max((len(str(cell.value)) for cell in cells if cell.value is not None), default=0)
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 45)

    return output.getvalue(), logs


def display(obj):
    if isinstance(obj, pd.DataFrame):
        st.dataframe(obj, use_container_width=True)
    else:
        pass

if st.button("Flatten Header → Buat Dashboard Surveilans", type="primary"):
    if not uploaded_files:
        st.error("Unggah dulu tiga file Excel.")
        st.stop()
    if len(uploaded_files) != 3:
        st.error(f"Harus mengunggah tepat 3 file Excel; saat ini {len(uploaded_files)}.")
        st.stop()

    with st.spinner("Tahap 1: flatten header. Tahap 2: memproses boundary, matching wilayah, dan membuat dashboard..."):
        try:
            # Jalankan pipeline notebook yang sudah disesuaikan untuk Streamlit.
            # ============================================================
            # CELL 1 — LANGKAH 1-10
            # DOWNLOAD DAN PERSIAPAN BOUNDARY
            # ============================================================


            # ============================================================
            # LANGKAH 1 — INSTALL PACKAGE
            # ============================================================




            # ============================================================
            # LANGKAH 2 — IMPORT
            # ============================================================

            import os
            import re
            import json
            import html
            import shutil
            import zipfile
            import subprocess
            import unicodedata
            import warnings

            from pathlib import Path

            import numpy as np
            import pandas as pd
            import geopandas as gpd
            import plotly.io as pio




            warnings.filterwarnings("ignore")

            # Hilangkan batas ukuran object GeoJSON pada GDAL
            os.environ["OGR_GEOJSON_MAX_OBJ_SIZE"] = "0"


            # ============================================================
            # LANGKAH 3 — KONFIGURASI
            # ============================================================

            GITHUB_REPO_URL = (
                "https://github.com/dmxsan/"
                "indonesia-admin-boundaries.git"
            )

            FOLDER_REPO = Path.cwd() / "boundary_repo"

            FOLDER_PROVINSI = (
                FOLDER_REPO
                / "processed-data"
                / "02-provinces"
                / "province-only"
            )

            FOLDER_KABUPATEN = (
                FOLDER_REPO
                / "processed-data"
                / "03-districts"
            )

            FOLDER_OUTPUT = Path(
                str(Path.cwd() / "output_peta_mr")
            )

            FILE_DASHBOARD = (
                FOLDER_OUTPUT
                / "Dashboard_Peta_Surveilans_MR.html"
            )

            FILE_LAPORAN = (
                FOLDER_OUTPUT
                / "Laporan_Matching.xlsx"
            )

            FILE_ZIP = Path(
                str(Path.cwd() / "Peta_Surveilans_MR_2023_2025.zip")
            )

            TAHUN_LIST = [2023, 2024, 2025]

            SHEET_PROVINSI = "Tabel Analisis"
            SHEET_KABUPATEN = "Analisis Pusat by Kab"

            KOLOM_PROVINSI = (
                "Kabupaten / Kota / Provinsi"
            )

            KOLOM_KABUPATEN = (
                "Kabupaten / Kota"
            )

            KOLOM_PROVINSI_KABUPATEN = (
                "Provinsi"
            )

            WARNA_RENDAH = "#E8F1F1"
            WARNA_TENGAH = "#64BABA"
            WARNA_TINGGI = "#008C8C"
            WARNA_NODATA = "#D9D9D9"

            # Simplifikasi untuk visualisasi, satuan meter
            TOLERANCE_ADM1 = 1500
            TOLERANCE_ADM2 = 1200

            # False membuat HTML lebih kecil, tetapi butuh internet saat dibuka.
            # True membuat dashboard dapat dibuka offline.
            DASHBOARD_OFFLINE = True


            # ============================================================
            # LANGKAH 4 — FUNGSI NORMALISASI
            # ============================================================

            def teks_dasar(teks):
                if pd.isna(teks):
                    return ""

                teks = str(teks).strip().upper()
                teks = teks.replace("_", " ")

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
                hasil = teks_dasar(teks)

                alias = {
                    "JAKARTA":
                        "DKI JAKARTA",

                    "DK JAKARTA":
                        "DKI JAKARTA",

                    "DAERAH KHUSUS IBUKOTA JAKARTA":
                        "DKI JAKARTA",

                    "YOGYAKARTA":
                        "DAERAH ISTIMEWA YOGYAKARTA",

                    "DI YOGYAKARTA":
                        "DAERAH ISTIMEWA YOGYAKARTA",

                    "D I YOGYAKARTA":
                        "DAERAH ISTIMEWA YOGYAKARTA",

                    "BANGKA BELITUNG":
                        "KEPULAUAN BANGKA BELITUNG",

                    "KEP BANGKA BELITUNG":
                        "KEPULAUAN BANGKA BELITUNG",

                    "KEP RIAU":
                        "KEPULAUAN RIAU",

                    "NANGGROE ACEH DARUSSALAM":
                        "ACEH",
                }

                return alias.get(
                    hasil,
                    hasil
                )


            # Alias menyesuaikan nama Excel ke nama file dalam repo BIG
            ALIAS_KABUPATEN = {
                "KAB KEP SERIBU":
                    "ADM KEP SERIBU",

                "KABUPATEN KEPULAUAN SERIBU":
                    "ADM KEP SERIBU",

                "KEPULAUAN SERIBU":
                    "ADM KEP SERIBU",

                "KOTA GUNUNG SITOLI":
                    "KOTA GUNUNGSITOLI",

                "DHARMAS RAYA":
                    "DHARMASRAYA",

                "KOTA PARE PARE":
                    "KOTA PAREPARE",

                "PANGKAJENE KEPULAUAN":
                    "PANGKAJENE DAN KEPULAUAN",

                "SELAYAR":
                    "KEPULAUAN SELAYAR",

                "KOTA BAU BAU":
                    "KOTA BAUBAU",

                "YAPEN WAROPEN":
                    "KEPULAUAN YAPEN",

                "POLIWALI MANDAR":
                    "POLEWALI MANDAR",

                "TAMBRAW":
                    "TAMBRAUW",
            }


            def normalisasi_kabupaten_excel(teks):
                hasil = teks_dasar(teks)

                hasil = re.sub(
                    r"^KAB\s+",
                    "KABUPATEN ",
                    hasil
                )

                hasil = re.sub(
                    r"^KOTA ADM\s+",
                    "KOTA ",
                    hasil
                )

                hasil = re.sub(
                    r"^KOTA ADMINISTRASI\s+",
                    "KOTA ",
                    hasil
                )

                hasil = re.sub(
                    r"^KOTAMADYA\s+",
                    "KOTA ",
                    hasil
                )

                # Boundary repo tidak selalu memakai kata Kabupaten
                hasil = re.sub(
                    r"^KABUPATEN\s+",
                    "",
                    hasil
                )

                return ALIAS_KABUPATEN.get(
                    hasil,
                    hasil
                )


            def normalisasi_kabupaten_boundary(teks):
                hasil = teks_dasar(teks)

                hasil = re.sub(
                    r"^KABUPATEN\s+",
                    "",
                    hasil
                )

                return hasil


            def normalisasi_kolom(teks):
                return re.sub(
                    r"\s+",
                    " ",
                    str(teks).strip().casefold()
                )


            # ============================================================
            # LANGKAH 5 — CLONE REPOSITORY
            # ============================================================

            if FOLDER_REPO.exists():
                shutil.rmtree(FOLDER_REPO)

            print("=" * 75)
            print("CLONE REPOSITORY BOUNDARY")
            print("=" * 75)

            subprocess.run(
                [
                    "git",
                    "clone",
                    "--depth",
                    "1",
                    GITHUB_REPO_URL,
                    str(FOLDER_REPO),
                ],
                check=True
            )

            print("Repository berhasil di-clone.")


            # ============================================================
            # LANGKAH 6 — IDENTIFIKASI FILE BOUNDARY
            # ============================================================

            if not FOLDER_PROVINSI.exists():
                raise FileNotFoundError(
                    f"Folder tidak ditemukan: {FOLDER_PROVINSI}"
                )

            if not FOLDER_KABUPATEN.exists():
                raise FileNotFoundError(
                    f"Folder tidak ditemukan: {FOLDER_KABUPATEN}"
                )

            files_adm1 = sorted(
                FOLDER_PROVINSI.rglob("*.geojson")
            )

            files_adm2 = sorted(
                FOLDER_KABUPATEN.rglob("*.geojson")
            )

            print("\nBoundary ditemukan:")

            print(
                "Provinsi:",
                len(files_adm1)
            )

            print(
                "Kabupaten/Kota:",
                len(files_adm2)
            )

            if len(files_adm1) == 0:
                raise FileNotFoundError(
                    "Tidak ditemukan GeoJSON provinsi."
                )

            if len(files_adm2) == 0:
                raise FileNotFoundError(
                    "Tidak ditemukan GeoJSON kabupaten/kota."
                )


            # ============================================================
            # LANGKAH 7 — FUNGSI MEMBACA GEOJSON
            # ============================================================

            def baca_geojson_aman(path):
                try:
                    gdf = gpd.read_file(
                        path,
                        engine="pyogrio"
                    )

                except Exception:
                    gdf = gpd.read_file(path)

                if len(gdf) == 0:
                    return None

                if gdf.crs is None:
                    gdf = gdf.set_crs(
                        "EPSG:4326"
                    )
                else:
                    gdf = gdf.to_crs(
                        "EPSG:4326"
                    )

                gdf = gdf[
                    gdf.geometry.notna()
                    & ~gdf.geometry.is_empty
                ].copy()

                if len(gdf) == 0:
                    return None

                invalid = ~gdf.geometry.is_valid

                if invalid.any():
                    try:
                        gdf.loc[
                            invalid,
                            "geometry"
                        ] = (
                            gdf.loc[
                                invalid,
                                "geometry"
                            ].buffer(0)
                        )
                    except Exception:
                        pass

                return gdf


            # ============================================================
            # LANGKAH 8 — BACA 38 PROVINSI
            # ============================================================

            daftar_adm1 = []
            gagal_adm1 = []

            for nomor, path in enumerate(
                files_adm1,
                start=1
            ):
                try:
                    gdf = baca_geojson_aman(path)

                    if gdf is None:
                        gagal_adm1.append({
                            "File": path.name,
                            "Error": "Geometri kosong",
                        })
                        continue

                    nama_provinsi = (
                        path.stem
                        .replace("_", " ")
                        .strip()
                        .upper()
                    )

                    gdf["PROVINSI"] = nama_provinsi

                    daftar_adm1.append(
                        gdf[
                            [
                                "PROVINSI",
                                "geometry"
                            ]
                        ].copy()
                    )

                    if (
                        nomor % 10 == 0
                        or nomor == len(files_adm1)
                    ):
                        print(
                            f"Provinsi dibaca: "
                            f"{nomor}/{len(files_adm1)}"
                        )

                except Exception as error:
                    gagal_adm1.append({
                        "File": path.name,
                        "Error": str(error),
                    })


            if len(daftar_adm1) == 0:
                raise ValueError(
                    "Tidak ada boundary provinsi yang berhasil dibaca."
                )

            gdf_provinsi = gpd.GeoDataFrame(
                pd.concat(
                    daftar_adm1,
                    ignore_index=True
                ),
                geometry="geometry",
                crs="EPSG:4326"
            )

            gdf_provinsi["KEY_PROV"] = (
                gdf_provinsi["PROVINSI"]
                .apply(normalisasi_provinsi)
            )

            gdf_provinsi = (
                gdf_provinsi[
                    [
                        "PROVINSI",
                        "KEY_PROV",
                        "geometry"
                    ]
                ]
                .dissolve(
                    by="KEY_PROV",
                    as_index=False,
                    aggfunc="first"
                )
                .reset_index(drop=True)
            )

            print("\nJumlah boundary provinsi:", len(gdf_provinsi))

            if len(gdf_provinsi) != 38:
                raise ValueError(
                    f"Boundary provinsi harus 38, "
                    f"tetapi ditemukan {len(gdf_provinsi)}."
                )


            # ============================================================
            # LANGKAH 9 — BACA SELURUH KABUPATEN/KOTA
            # ============================================================

            daftar_adm2 = []
            gagal_adm2 = []

            for nomor, path in enumerate(
                files_adm2,
                start=1
            ):
                try:
                    gdf = baca_geojson_aman(path)

                    if gdf is None:
                        gagal_adm2.append({
                            "File": path.name,
                            "Error": "Geometri kosong",
                        })
                        continue

                    # Nama file digunakan karena struktur repo
                    # sudah satu file per kabupaten/kota.
                    nama_kabupaten = (
                        path.stem
                        .replace("_", " ")
                        .strip()
                    )

                    gdf["NAMA_KAB"] = nama_kabupaten

                    daftar_adm2.append(
                        gdf[
                            [
                                "NAMA_KAB",
                                "geometry"
                            ]
                        ].copy()
                    )

                    if (
                        nomor % 50 == 0
                        or nomor == len(files_adm2)
                    ):
                        print(
                            f"Kabupaten/Kota dibaca: "
                            f"{nomor}/{len(files_adm2)}"
                        )

                except Exception as error:
                    gagal_adm2.append({
                        "File": path.name,
                        "Error": str(error),
                    })


            if len(daftar_adm2) == 0:
                raise ValueError(
                    "Tidak ada boundary kabupaten/kota "
                    "yang berhasil dibaca."
                )

            gdf_kabupaten = gpd.GeoDataFrame(
                pd.concat(
                    daftar_adm2,
                    ignore_index=True
                ),
                geometry="geometry",
                crs="EPSG:4326"
            )

            gdf_kabupaten = (
                gdf_kabupaten[
                    gdf_kabupaten.geometry.notna()
                    & ~gdf_kabupaten.geometry.is_empty
                ]
                .reset_index(drop=True)
            )

            gdf_kabupaten["KEY_KAB"] = (
                gdf_kabupaten["NAMA_KAB"]
                .apply(normalisasi_kabupaten_boundary)
            )

            print(
                "Jumlah feature kabupaten/kota:",
                len(gdf_kabupaten)
            )


            # ============================================================
            # LANGKAH 10 — ASSIGN ADM2 KE 38 PROVINSI
            # ============================================================

            titik_adm2 = gdf_kabupaten[
                ["geometry"]
            ].copy()

            titik_adm2["ID_ADM2"] = (
                titik_adm2.index
            )

            titik_adm2["geometry"] = (
                titik_adm2.geometry.representative_point()
            )

            prov_join = gdf_provinsi[
                [
                    "PROVINSI",
                    "KEY_PROV",
                    "geometry"
                ]
            ].copy()

            hasil_join = gpd.sjoin(
                titik_adm2,
                prov_join,
                how="left",
                predicate="within"
            )

            hasil_join = (
                hasil_join
                .sort_values("ID_ADM2")
                .drop_duplicates(
                    subset="ID_ADM2",
                    keep="first"
                )
                .set_index("ID_ADM2")
            )

            gdf_kabupaten["PROVINSI"] = (
                gdf_kabupaten.index.map(
                    hasil_join["PROVINSI"]
                )
            )

            gdf_kabupaten["KEY_PROV"] = (
                gdf_kabupaten.index.map(
                    hasil_join["KEY_PROV"]
                )
            )

            jumlah_kosong = int(
                gdf_kabupaten["KEY_PROV"]
                .isna()
                .sum()
            )

            if jumlah_kosong > 0:
                raise ValueError(
                    f"Ada {jumlah_kosong} boundary ADM2 "
                    "belum mendapatkan provinsi."
                )

            gdf_kabupaten["KEY_MAP"] = (
                gdf_kabupaten["KEY_PROV"]
                + "|"
                + gdf_kabupaten["KEY_KAB"]
            )



            # ============================================================
            # CEK DUPLIKAT KEY ADM2
            # ============================================================

            # ============================================================
            # GABUNGKAN FEATURE ADM2 YANG MEMILIKI KEY_MAP SAMA
            # ============================================================

            duplikat_boundary = gdf_kabupaten[
                gdf_kabupaten["KEY_MAP"].duplicated(
                    keep=False
                )
            ].copy()

            if len(duplikat_boundary) > 0:

                print(
                    f"Ditemukan {len(duplikat_boundary)} feature "
                    "dengan KEY_MAP berulang."
                )

                print(
                    "Feature tersebut akan digabungkan "
                    "berdasarkan KEY_MAP."
                )

                display(
                    duplikat_boundary[
                        [
                            "PROVINSI",
                            "NAMA_KAB",
                            "KEY_KAB",
                            "KEY_MAP"
                        ]
                    ].sort_values(
                        "KEY_MAP"
                    )
                )


            # ============================================================
            # DISSOLVE
            # Satu KEY_MAP = satu wilayah administratif
            # ============================================================

            gdf_kabupaten = (
                gdf_kabupaten[
                    [
                        "PROVINSI",
                        "KEY_PROV",
                        "NAMA_KAB",
                        "KEY_KAB",
                        "KEY_MAP",
                        "geometry"
                    ]
                ]
                .dissolve(
                    by="KEY_MAP",
                    as_index=False,
                    aggfunc={
                        "PROVINSI": "first",
                        "KEY_PROV": "first",
                        "NAMA_KAB": "first",
                        "KEY_KAB": "first",
                    }
                )
                .reset_index(drop=True)
            )


            # ============================================================
            # VALIDASI SETELAH DISSOLVE
            # ============================================================

            duplikat_setelah = (
                gdf_kabupaten["KEY_MAP"]
                .duplicated()
                .sum()
            )

            print()
            print("=" * 75)
            print("HASIL SETELAH DISSOLVE")
            print("=" * 75)

            print(
                "Jumlah ADM2 final:",
                len(gdf_kabupaten)
            )

            print(
                "Provinsi ADM2:",
                gdf_kabupaten[
                    "KEY_PROV"
                ].nunique()
            )

            print(
                "KEY_MAP duplikat:",
                duplikat_setelah
            )

            print(
                "Geometri kosong:",
                gdf_kabupaten.geometry.isna().sum()
            )

            print(
                "Geometri invalid:",
                (~gdf_kabupaten.geometry.is_valid).sum()
            )


            if duplikat_setelah > 0:

                raise ValueError(
                    "Masih terdapat KEY_MAP duplikat "
                    "setelah dissolve."
                )

            else:

                print()
                print(
                    "✅ Boundary ADM2 sudah unik "
                    "berdasarkan KEY_MAP."
                )

            print("\n" + "=" * 75)
            print("CELL 1 SELESAI")
            print("=" * 75)

            print(
                "Provinsi:",
                len(gdf_provinsi)
            )

            print(
                "Kabupaten/Kota:",
                len(gdf_kabupaten)
            )

            print(
                "Provinsi dalam ADM2:",
                gdf_kabupaten["KEY_PROV"].nunique()
            )

            print(
                "Duplikat key ADM2:",
                gdf_kabupaten["KEY_MAP"]
                .duplicated()
                .sum()
            )

            # ============================================================
            # CELL 2 — LANGKAH 11-20
            # UPLOAD EXCEL, MATCHING, VALIDASI, DAN PAYLOAD
            #
            # PRASYARAT:
            # CELL 1 SUDAH BERHASIL
            # ============================================================


            # ============================================================
            # LANGKAH 11 — OPTIMASI GEOMETRY UNTUK DASHBOARD
            # ============================================================

            print("=" * 75)
            print("LANGKAH 11 — OPTIMASI GEOMETRY")
            print("=" * 75)


            # ------------------------------------------------------------
            # PROVINSI
            # ------------------------------------------------------------

            gdf_prov_plot = (
                gdf_provinsi
                .to_crs("EPSG:3857")
                .copy()
            )

            gdf_prov_plot["geometry"] = (
                gdf_prov_plot.geometry.simplify(
                    tolerance=TOLERANCE_ADM1,
                    preserve_topology=True
                )
            )

            gdf_prov_plot = (
                gdf_prov_plot
                .to_crs("EPSG:4326")
                .reset_index(drop=True)
            )

            gdf_prov_plot["MAP_ID"] = (
                gdf_prov_plot.index.astype(str)
            )

            gdf_prov_plot = gdf_prov_plot[
                [
                    "PROVINSI",
                    "KEY_PROV",
                    "MAP_ID",
                    "geometry"
                ]
            ].copy()


            # ------------------------------------------------------------
            # KABUPATEN/KOTA
            # ------------------------------------------------------------

            gdf_kab_plot = (
                gdf_kabupaten
                .to_crs("EPSG:3857")
                .copy()
            )

            if TOLERANCE_ADM2 > 0:

                gdf_kab_plot["geometry"] = (
                    gdf_kab_plot.geometry.simplify(
                        tolerance=TOLERANCE_ADM2,
                        preserve_topology=True
                    )
                )

            gdf_kab_plot = (
                gdf_kab_plot
                .to_crs("EPSG:4326")
                .reset_index(drop=True)
            )

            gdf_kab_plot["MAP_ID"] = (
                gdf_kab_plot.index.astype(str)
            )


            # Validasi KEY_MAP dari Cell 1
            if "KEY_MAP" not in gdf_kab_plot.columns:
                raise KeyError(
                    "KEY_MAP tidak ditemukan pada gdf_kab_plot. "
                    "Cell 1 belum selesai dengan benar."
                )


            gdf_kab_plot = gdf_kab_plot[
                [
                    "NAMA_KAB",
                    "PROVINSI",
                    "KEY_PROV",
                    "KEY_KAB",
                    "KEY_MAP",
                    "MAP_ID",
                    "geometry"
                ]
            ].copy()


            print(
                "Provinsi untuk plot:",
                len(gdf_prov_plot)
            )

            print(
                "Kab/Kota untuk plot:",
                len(gdf_kab_plot)
            )


            # ============================================================
            # LANGKAH 12 — BUAT GEOJSON SEKALI
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 12 — MEMBUAT GEOJSON")
            print("=" * 75)


            geojson_prov = json.loads(
                gdf_prov_plot.to_json()
            )

            geojson_kab = json.loads(
                gdf_kab_plot.to_json()
            )


            geojson_prov_json = json.dumps(
                geojson_prov,
                ensure_ascii=False,
                separators=(",", ":")
            )

            geojson_kab_json = json.dumps(
                geojson_kab,
                ensure_ascii=False,
                separators=(",", ":")
            )


            ukuran_prov_mb = (
                len(
                    geojson_prov_json.encode("utf-8")
                )
                / 1024**2
            )

            ukuran_kab_mb = (
                len(
                    geojson_kab_json.encode("utf-8")
                )
                / 1024**2
            )


            print(
                f"Ukuran GeoJSON Provinsi: "
                f"{ukuran_prov_mb:.2f} MB"
            )

            print(
                f"Ukuran GeoJSON Kab/Kota: "
                f"{ukuran_kab_mb:.2f} MB"
            )


            # ============================================================
            # LANGKAH 13 — UPLOAD 3 FILE EXCEL
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 13 — UPLOAD EXCEL")
            print("=" * 75)

            print(
                "Upload sekaligus:\n"
                "- Indonesia 2023 Week 52.xlsx\n"
                "- Indonesia 2024 Week 52.xlsx\n"
                "- Indonesia 2025 Week 53.xlsx"
            )


            # --------------------------------------------------------
            # FLATTEN HEADER DAHULU, BARU DIPAKAI DASHBOARD
            # --------------------------------------------------------
            uploaded = {}
            log_flatten = []

            for uploaded_file in uploaded_files:
                nama_asli = uploaded_file.name
                konten_asli = uploaded_file.getvalue()

                # Proses hanya XLSX/XLSM dengan openpyxl.
                if Path(nama_asli).suffix.lower() not in [".xlsx", ".xlsm"]:
                    raise ValueError(
                        f"Format {nama_asli} belum didukung untuk flatten header. "
                        "Gunakan .xlsx atau .xlsm."
                    )

                konten_flat, detail_flat = _flatten_file_for_dashboard(
                    konten_asli,
                    nama_asli
                )

                nama_flat = f"{Path(nama_asli).stem}_Flatten_Header.xlsx"
                uploaded[nama_flat] = konten_flat
                (Path.cwd() / nama_flat).write_bytes(konten_flat)

                log_flatten.append({
                    "File Asli": nama_asli,
                    "File Hasil Flatten": nama_flat,
                    "Detail": " | ".join(detail_flat),
                })

            st.success("Tahap 1/2 selesai: header berhasil di-flatten untuk file yang valid.")
            st.dataframe(pd.DataFrame(log_flatten), use_container_width=True, hide_index=True)


            # ============================================================
            # LANGKAH 14 — IDENTIFIKASI EXCEL BERDASARKAN TAHUN
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 14 — IDENTIFIKASI EXCEL")
            print("=" * 75)


            file_excel = {}

            for tahun in TAHUN_LIST:

                kandidat = [
                    Path.cwd() / nama
                    for nama in uploaded.keys()
                    if (
                        str(tahun) in nama
                        and
                        Path(nama).suffix.lower()
                        in [".xlsx", ".xlsm", ".xls"]
                    )
                ]

                if len(kandidat) != 1:

                    raise FileNotFoundError(
                        f"Tahun {tahun}: harus ditemukan "
                        f"tepat 1 file, tetapi ditemukan "
                        f"{len(kandidat)}."
                    )

                file_excel[tahun] = kandidat[0]


            for tahun, path in file_excel.items():

                print(
                    f"{tahun} -> {path.name}"
                )


            # ============================================================
            # LANGKAH 15 — FUNGSI KOLOM DAN NUMERIK
            # ============================================================

            def cari_kolom_indikator(
                df,
                nama_indikator
            ):

                # Persis
                if nama_indikator in df.columns:
                    return nama_indikator

                target = normalisasi_kolom(
                    nama_indikator
                )

                # Normalisasi persis
                kandidat = [
                    kolom
                    for kolom in df.columns
                    if normalisasi_kolom(kolom)
                    == target
                ]

                if len(kandidat) >= 1:
                    return kandidat[0]

                # Awalan
                kandidat_awalan = [
                    kolom
                    for kolom in df.columns
                    if normalisasi_kolom(kolom)
                    .startswith(target)
                ]

                if len(kandidat_awalan) == 1:
                    return kandidat_awalan[0]

                return None


            def konversi_satu_nilai(nilai):

                if pd.isna(nilai):
                    return np.nan

                if isinstance(
                    nilai,
                    (
                        int,
                        float,
                        np.integer,
                        np.floating
                    )
                ):
                    return float(nilai)

                teks = str(nilai).strip()

                if teks.casefold() in {
                    "",
                    "-",
                    "--",
                    "na",
                    "n/a",
                    "none",
                    "nan",
                    "null",
                }:
                    return np.nan

                teks = teks.replace("%", "")
                teks = teks.replace(" ", "")

                # Format:
                # 1.234,5
                # 1,234.5
                if (
                    "," in teks
                    and "." in teks
                ):

                    if teks.rfind(",") > teks.rfind("."):

                        teks = teks.replace(
                            ".",
                            ""
                        )

                        teks = teks.replace(
                            ",",
                            "."
                        )

                    else:

                        teks = teks.replace(
                            ",",
                            ""
                        )

                elif "," in teks:

                    bagian = teks.split(",")

                    if (
                        len(bagian) == 2
                        and len(bagian[1]) <= 4
                    ):

                        teks = teks.replace(
                            ",",
                            "."
                        )

                    else:

                        teks = teks.replace(
                            ",",
                            ""
                        )

                try:

                    return float(teks)

                except Exception:

                    return np.nan


            # ============================================================
            # LANGKAH 16 — BACA DAN STANDARDISASI EXCEL
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 16 — BACA EXCEL")
            print("=" * 75)


            data_provinsi = {}
            data_kabupaten = {}


            for tahun, path in file_excel.items():

                xls = pd.ExcelFile(path)


                # --------------------------------------------------------
                # VALIDASI SHEET
                # --------------------------------------------------------

                if SHEET_PROVINSI not in xls.sheet_names:

                    raise ValueError(
                        f"Sheet '{SHEET_PROVINSI}' "
                        f"tidak ditemukan pada {tahun}."
                    )


                if SHEET_KABUPATEN not in xls.sheet_names:

                    raise ValueError(
                        f"Sheet '{SHEET_KABUPATEN}' "
                        f"tidak ditemukan pada {tahun}."
                    )


                # --------------------------------------------------------
                # BACA
                # --------------------------------------------------------

                df_prov = pd.read_excel(
                    path,
                    sheet_name=SHEET_PROVINSI
                )

                df_kab = pd.read_excel(
                    path,
                    sheet_name=SHEET_KABUPATEN
                )


                # --------------------------------------------------------
                # RAPIIKAN HEADER
                # --------------------------------------------------------

                df_prov.columns = [
                    str(kolom).strip()
                    for kolom in df_prov.columns
                ]

                df_kab.columns = [
                    str(kolom).strip()
                    for kolom in df_kab.columns
                ]


                # --------------------------------------------------------
                # VALIDASI KOLOM WILAYAH
                # --------------------------------------------------------

                if KOLOM_PROVINSI not in df_prov.columns:

                    raise KeyError(
                        f"Kolom '{KOLOM_PROVINSI}' "
                        f"tidak ditemukan pada data "
                        f"provinsi {tahun}."
                    )


                if (
                    KOLOM_PROVINSI_KABUPATEN
                    not in df_kab.columns
                ):

                    raise KeyError(
                        f"Kolom "
                        f"'{KOLOM_PROVINSI_KABUPATEN}' "
                        f"tidak ditemukan pada "
                        f"data kabupaten {tahun}."
                    )


                if KOLOM_KABUPATEN not in df_kab.columns:

                    raise KeyError(
                        f"Kolom '{KOLOM_KABUPATEN}' "
                        f"tidak ditemukan pada "
                        f"data kabupaten {tahun}."
                    )


                # --------------------------------------------------------
                # PROVINSI
                # --------------------------------------------------------

                df_prov["KEY_PROV"] = (
                    df_prov[
                        KOLOM_PROVINSI
                    ]
                    .apply(
                        normalisasi_provinsi
                    )
                )


                # Buang nasional dan kosong
                df_prov = df_prov[
                    (df_prov["KEY_PROV"] != "")
                    &
                    (df_prov["KEY_PROV"] != "INDONESIA")
                ].copy()


                # --------------------------------------------------------
                # KABUPATEN/KOTA
                # --------------------------------------------------------

                df_kab["KEY_PROV"] = (
                    df_kab[
                        KOLOM_PROVINSI_KABUPATEN
                    ]
                    .apply(
                        normalisasi_provinsi
                    )
                )


                df_kab["KEY_KAB"] = (
                    df_kab[
                        KOLOM_KABUPATEN
                    ]
                    .apply(
                        normalisasi_kabupaten_excel
                    )
                )


                # IMPORTANT:
                # Buang Indonesia SEBELUM membuat KEY_MAP.
                df_kab = df_kab[
                    (df_kab["KEY_PROV"] != "")
                    &
                    (df_kab["KEY_KAB"] != "")
                    &
                    (df_kab["KEY_KAB"] != "INDONESIA")
                ].copy()


                # Sekarang buat KEY_MAP
                df_kab["KEY_MAP"] = (
                    df_kab["KEY_PROV"]
                    + "|"
                    + df_kab["KEY_KAB"]
                )


                # --------------------------------------------------------
                # VALIDASI DUPLIKAT EXCEL
                # --------------------------------------------------------

                duplikat_prov = (
                    df_prov[
                        "KEY_PROV"
                    ]
                    .duplicated()
                    .sum()
                )


                duplikat_kab = (
                    df_kab[
                        "KEY_MAP"
                    ]
                    .duplicated()
                    .sum()
                )


                if duplikat_prov > 0:

                    print(
                        f"PERINGATAN {tahun}: "
                        f"{duplikat_prov} "
                        f"KEY_PROV duplikat."
                    )


                if duplikat_kab > 0:

                    print(
                        f"PERINGATAN {tahun}: "
                        f"{duplikat_kab} "
                        f"KEY_MAP Excel duplikat."
                    )


                # --------------------------------------------------------
                # SIMPAN
                # --------------------------------------------------------

                data_provinsi[
                    tahun
                ] = df_prov


                data_kabupaten[
                    tahun
                ] = df_kab


                print(
                    f"{tahun}: "
                    f"{len(df_prov)} provinsi; "
                    f"{len(df_kab)} kab/kota"
                )


            # ============================================================
            # LANGKAH 17 — VALIDASI MATCHING WILAYAH
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 17 — VALIDASI MATCHING")
            print("=" * 75)


            key_boundary_prov = set(
                gdf_prov_plot[
                    "KEY_PROV"
                ]
            )


            key_boundary_kab = set(
                gdf_kab_plot[
                    "KEY_MAP"
                ]
            )


            laporan_matching = []


            for tahun in TAHUN_LIST:

                # --------------------------------------------------------
                # PROVINSI
                # --------------------------------------------------------

                prov_keys = (
                    data_provinsi[
                        tahun
                    ]["KEY_PROV"]
                    .dropna()
                    .drop_duplicates()
                    .tolist()
                )


                for key in sorted(prov_keys):

                    laporan_matching.append({
                        "Tahun":
                            tahun,

                        "Level":
                            "Provinsi",

                        "Key":
                            key,

                        "Match":
                            key in key_boundary_prov,
                    })


                # --------------------------------------------------------
                # KABUPATEN/KOTA
                # --------------------------------------------------------

                kab_keys = (
                    data_kabupaten[
                        tahun
                    ]["KEY_MAP"]
                    .dropna()
                    .drop_duplicates()
                    .tolist()
                )


                for key in sorted(kab_keys):

                    laporan_matching.append({
                        "Tahun":
                            tahun,

                        "Level":
                            "Kabupaten/Kota",

                        "Key":
                            key,

                        "Match":
                            key in key_boundary_kab,
                    })


            df_matching = pd.DataFrame(
                laporan_matching
            )


            ringkasan_matching = (
                df_matching
                .groupby(
                    [
                        "Tahun",
                        "Level",
                        "Match"
                    ]
                )
                .size()
                .reset_index(
                    name="Jumlah"
                )
            )


            display(
                ringkasan_matching
            )


            belum_match = (
                df_matching[
                    ~df_matching[
                        "Match"
                    ]
                ]
                .copy()
            )


            if len(belum_match) > 0:

                print()
                print(
                    "⚠️ Masih ada wilayah belum match:",
                    len(belum_match)
                )

                display(
                    belum_match
                )

            else:

                print()
                print(
                    "✅ Matching wilayah 100%."
                )


            # ============================================================
            # LANGKAH 18 — DIAGNOSTIK OTOMATIS WILAYAH BELUM MATCH
            # ============================================================

            from difflib import get_close_matches


            diagnostik_matching = []


            belum_match_kab = (
                belum_match[
                    belum_match[
                        "Level"
                    ] == "Kabupaten/Kota"
                ][
                    ["Key"]
                ]
                .drop_duplicates()
                .reset_index(drop=True)
            )


            for _, row in belum_match_kab.iterrows():

                key_excel = row["Key"]

                if "|" not in key_excel:
                    continue


                provinsi, nama_excel = (
                    key_excel.split(
                        "|",
                        1
                    )
                )


                kandidat_boundary = (
                    gdf_kab_plot.loc[
                        gdf_kab_plot[
                            "KEY_PROV"
                        ] == provinsi,
                        "KEY_KAB"
                    ]
                    .dropna()
                    .drop_duplicates()
                    .tolist()
                )


                kandidat = get_close_matches(
                    nama_excel,
                    kandidat_boundary,
                    n=3,
                    cutoff=0.25
                )


                diagnostik_matching.append({
                    "PROVINSI":
                        provinsi,

                    "NAMA_EXCEL":
                        nama_excel,

                    "KANDIDAT_1":
                        kandidat[0]
                        if len(kandidat) > 0
                        else "",

                    "KANDIDAT_2":
                        kandidat[1]
                        if len(kandidat) > 1
                        else "",

                    "KANDIDAT_3":
                        kandidat[2]
                        if len(kandidat) > 2
                        else "",
                })


            df_diagnostik_matching = pd.DataFrame(
                diagnostik_matching
            )


            if len(
                df_diagnostik_matching
            ) > 0:

                print()
                print("=" * 75)
                print(
                    "DIAGNOSTIK WILAYAH BELUM MATCH"
                )
                print("=" * 75)

                display(
                    df_diagnostik_matching
                )


            # ============================================================
            # LANGKAH 19 — IDENTIFIKASI INDIKATOR
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 19 — IDENTIFIKASI INDIKATOR")
            print("=" * 75)


            INDIKATOR_PILIHAN = [
                "Jumlah Suspek",

                "Jumlah Meninggal",

                "Jumlah Suspek dengan spesimen",

                "Performa Surveilans_Measles Incidence Rate (<1/1.000.000 population)",

                "Performa Surveilans_Rubella Incidence Rate (<1/1.000.000 population)",

                "Performa Surveilans_DISCARDED RATE (≥2/100.000 population)",

                "Performa Surveilans_Pos Rate Measles",

                "Performa Surveilans_Investigasi Adekuat (%)",

                "Performa Surveilans_Spesimen Adekuat (%)",

                "Performa Surveilans_Timeliness Specimen Result (%)",

                "Performa Surveilans_Laboratory Confirmation (%)",

                "Klasifikasi_Campak Positif",

                "Klasifikasi_Rubella Positif",

                "Klasifikasi_Konfirmasi Hubungan Epidemiologi Campak",

                "Klasifikasi_Konfirmasi Hubungan Epidemiologi Rubella",

                "Klasifikasi_Negatif",

                "Klasifikasi_Klinis",

                "Klasifikasi_Pending",
            ]


            indikator = []


            for nama_indikator in INDIKATOR_PILIHAN:

                tersedia = False


                for tahun in TAHUN_LIST:

                    kolom_prov = (
                        cari_kolom_indikator(
                            data_provinsi[
                                tahun
                            ],
                            nama_indikator
                        )
                    )


                    kolom_kab = (
                        cari_kolom_indikator(
                            data_kabupaten[
                                tahun
                            ],
                            nama_indikator
                        )
                    )


                    if (
                        kolom_prov is not None
                        or
                        kolom_kab is not None
                    ):

                        tersedia = True

                        break


                if tersedia:

                    indikator.append(
                        nama_indikator
                    )


            if len(indikator) == 0:

                raise ValueError(
                    "Tidak ada indikator yang "
                    "berhasil ditemukan."
                )


            for nomor, nama in enumerate(
                indikator,
                start=1
            ):

                print(
                    f"{nomor}. {nama}"
                )


            # ------------------------------------------------------------
            # TIPE INDIKATOR
            # ------------------------------------------------------------

            def tipe_indikator(nama):

                nama_norm = normalisasi_kolom(
                    nama
                )


                if any(
                    kata in nama_norm
                    for kata in [
                        "investigasi adekuat",
                        "spesimen adekuat",
                        "timeliness",
                        "laboratory confirmation",
                        "pos rate",
                        "persen",
                        "percentage",
                    ]
                ):

                    return "Persentase"


                if any(
                    kata in nama_norm
                    for kata in [
                        "incidence rate",
                        "discarded rate",
                    ]
                ):

                    return "Rate"


                return "Jumlah"


            # ============================================================
            # LANGKAH 20 — BUAT PAYLOAD DASHBOARD
            # ============================================================

            print("\n" + "=" * 75)
            print("LANGKAH 20 — MEMBUAT PAYLOAD")
            print("=" * 75)


            payload = {}

            log_indikator = []


            for tahun in TAHUN_LIST:

                for level in [
                    "Provinsi",
                    "Kabupaten/Kota"
                ]:


                    # ----------------------------------------------------
                    # TENTUKAN DATA DAN KEY
                    # ----------------------------------------------------

                    if level == "Provinsi":

                        df = data_provinsi[
                            tahun
                        ]

                        field_key = "KEY_PROV"

                        boundary_order = (
                            gdf_prov_plot[
                                "KEY_PROV"
                            ]
                            .tolist()
                        )


                    else:

                        df = data_kabupaten[
                            tahun
                        ]

                        field_key = "KEY_MAP"

                        boundary_order = (
                            gdf_kab_plot[
                                "KEY_MAP"
                            ]
                            .tolist()
                        )


                    # ----------------------------------------------------
                    # LOOP INDIKATOR
                    # ----------------------------------------------------

                    for nama_indikator in indikator:

                        kolom = cari_kolom_indikator(
                            df,
                            nama_indikator
                        )


                        key_payload = (
                            f"{tahun}|||"
                            f"{level}|||"
                            f"{nama_indikator}"
                        )


                        # ------------------------------------------------
                        # INDIKATOR TIDAK TERSEDIA
                        # ------------------------------------------------

                        if kolom is None:

                            log_indikator.append({
                                "Tahun":
                                    tahun,

                                "Level":
                                    level,

                                "Indikator":
                                    nama_indikator,

                                "Kolom sumber":
                                    "",

                                "Wilayah dengan data":
                                    0,

                                "Status":
                                    "Tidak tersedia",
                            })

                            continue


                        # ------------------------------------------------
                        # KONVERSI NILAI
                        # ------------------------------------------------

                        series_nilai = (
                            df[
                                kolom
                            ]
                            .apply(
                                konversi_satu_nilai
                            )
                        )


                        tabel_lookup = pd.DataFrame({
                            "KEY":
                                df[
                                    field_key
                                ],

                            "NILAI":
                                series_nilai,
                        })


                        # Buang key kosong
                        tabel_lookup = (
                            tabel_lookup[
                                tabel_lookup[
                                    "KEY"
                                ].notna()
                                &
                                (
                                    tabel_lookup[
                                        "KEY"
                                    ] != ""
                                )
                            ]
                            .copy()
                        )


                        # Jika ada duplicate,
                        # prioritaskan nilai non-null.
                        tabel_lookup[
                            "_IS_NULL"
                        ] = (
                            tabel_lookup[
                                "NILAI"
                            ]
                            .isna()
                            .astype(int)
                        )


                        tabel_lookup = (
                            tabel_lookup
                            .sort_values(
                                [
                                    "KEY",
                                    "_IS_NULL"
                                ]
                            )
                            .drop_duplicates(
                                subset="KEY",
                                keep="first"
                            )
                        )


                        lookup = dict(
                            zip(
                                tabel_lookup[
                                    "KEY"
                                ],
                                tabel_lookup[
                                    "NILAI"
                                ]
                            )
                        )


                        # ------------------------------------------------
                        # SUSUN SESUAI URUTAN BOUNDARY
                        # ------------------------------------------------

                        nilai_urut = [
                            lookup.get(
                                key,
                                np.nan
                            )
                            for key in boundary_order
                        ]


                        nilai_valid = [
                            float(nilai)
                            for nilai in nilai_urut
                            if pd.notna(nilai)
                        ]


                        jenis = tipe_indikator(
                            nama_indikator
                        )


                        # ------------------------------------------------
                        # SKALA WARNA
                        # ------------------------------------------------

                        if jenis == "Persentase":

                            zmin = 0.0
                            zmax = 100.0


                        elif len(
                            nilai_valid
                        ) > 0:

                            zmin = 0.0

                            # P95 supaya outlier tidak membuat
                            # sebagian besar wilayah terlalu pucat.
                            zmax = float(
                                np.nanpercentile(
                                    nilai_valid,
                                    95
                                )
                            )


                            if zmax <= 0:

                                zmax = float(
                                    max(
                                        nilai_valid
                                    )
                                )


                            if zmax <= 0:

                                zmax = 1.0


                        else:

                            zmin = 0.0
                            zmax = 1.0


                        # ------------------------------------------------
                        # SIMPAN PAYLOAD
                        # ------------------------------------------------

                        payload[
                            key_payload
                        ] = {
                            "values": [
                                None
                                if pd.isna(nilai)
                                else float(nilai)
                                for nilai in nilai_urut
                            ],

                            "zmin":
                                zmin,

                            "zmax":
                                zmax,

                            "type":
                                jenis,

                            "source_column":
                                kolom,
                        }


                        log_indikator.append({
                            "Tahun":
                                tahun,

                            "Level":
                                level,

                            "Indikator":
                                nama_indikator,

                            "Kolom sumber":
                                kolom,

                            "Wilayah dengan data":
                                len(nilai_valid),

                            "Status":
                                "Tersedia",
                        })


            df_log_indikator = pd.DataFrame(
                log_indikator
            )


            # ============================================================
            # RINGKASAN CELL 2
            # ============================================================

            print("\n" + "=" * 75)
            print("CELL 2 SELESAI")
            print("=" * 75)

            print(
                "Provinsi boundary:",
                len(gdf_prov_plot)
            )

            print(
                "Kab/Kota boundary:",
                len(gdf_kab_plot)
            )

            print(
                "Wilayah belum match:",
                len(belum_match)
            )

            print(
                "Indikator:",
                len(indikator)
            )

            print(
                "Kombinasi payload:",
                len(payload)
            )

            # ============================================================
            # DIAGNOSTIK NAMA KAB/KOTA YANG BELUM MATCH
            # ============================================================

            from difflib import get_close_matches


            # Buang |INDONESIA dari analisis
            belum_match_kab = df_matching[
                (df_matching["Level"] == "Kabupaten/Kota")
                & (~df_matching["Match"])
                & (~df_matching["Key"].str.endswith("|INDONESIA"))
            ].copy()


            # Karena hasilnya sama untuk 2023/2024/2025,
            # cukup ambil nama unik
            belum_match_unik = (
                belum_match_kab[
                    ["Key"]
                ]
                .drop_duplicates()
                .reset_index(drop=True)
            )


            # ------------------------------------------------------------
            # Cari kandidat hanya pada provinsi yang sama
            # ------------------------------------------------------------

            hasil_kandidat = []

            for _, row in belum_match_unik.iterrows():

                key_excel = row["Key"]

                provinsi, nama_excel = key_excel.split(
                    "|",
                    1
                )

                kandidat_boundary = (
                    gdf_kab_plot[
                        gdf_kab_plot[
                            "KEY_PROV"
                        ] == provinsi
                    ]["KEY_KAB"]
                    .dropna()
                    .unique()
                    .tolist()
                )

                kandidat = get_close_matches(
                    nama_excel,
                    kandidat_boundary,
                    n=5,
                    cutoff=0.25
                )

                hasil_kandidat.append({
                    "PROVINSI":
                        provinsi,

                    "NAMA_EXCEL":
                        nama_excel,

                    "KANDIDAT_1":
                        kandidat[0]
                        if len(kandidat) > 0
                        else "",

                    "KANDIDAT_2":
                        kandidat[1]
                        if len(kandidat) > 1
                        else "",

                    "KANDIDAT_3":
                        kandidat[2]
                        if len(kandidat) > 2
                        else "",

                    "KANDIDAT_4":
                        kandidat[3]
                        if len(kandidat) > 3
                        else "",

                    "KANDIDAT_5":
                        kandidat[4]
                        if len(kandidat) > 4
                        else "",
                })


            df_kandidat_missing = pd.DataFrame(
                hasil_kandidat
            )


            print("=" * 90)
            print("KABUPATEN/KOTA BELUM MATCH")
            print("=" * 90)

            print(
                "Jumlah unik belum match:",
                len(df_kandidat_missing)
            )

            display(
                df_kandidat_missing
            )

            # ============================================================
            # CELL 3
            # FIX DUPLIKAT BOUNDARY ADM2
            # GABUNGKAN MULTIPLE FEATURE DENGAN KEY_MAP YANG SAMA
            # ============================================================

            import geopandas as gpd
            import pandas as pd


            # ============================================================
            # 1. VALIDASI VARIABEL
            # ============================================================

            if "gdf_kabupaten" not in globals():
                raise NameError(
                    "gdf_kabupaten belum tersedia. "
                    "Jalankan Cell 1 dan Cell 2 terlebih dahulu."
                )

            if "KEY_MAP" not in gdf_kabupaten.columns:
                raise KeyError(
                    "Kolom KEY_MAP tidak ditemukan dalam gdf_kabupaten."
                )


            # ============================================================
            # 2. SIMPAN COPY SEBELUM DIPERBAIKI
            # ============================================================

            gdf_kabupaten_sebelum_dissolve = (
                gdf_kabupaten.copy()
            )


            # ============================================================
            # 3. IDENTIFIKASI DUPLIKAT
            # ============================================================

            mask_duplikat = (
                gdf_kabupaten["KEY_MAP"]
                .duplicated(keep=False)
            )

            gdf_duplikat = (
                gdf_kabupaten[
                    mask_duplikat
                ]
                .copy()
            )


            print("=" * 75)
            print("DUPLIKAT ADM2 SEBELUM DISSOLVE")
            print("=" * 75)

            print(
                "Jumlah feature awal:",
                len(gdf_kabupaten)
            )

            print(
                "Feature yang masuk kelompok duplikat:",
                len(gdf_duplikat)
            )

            print(
                "KEY_MAP duplikat unik:",
                gdf_duplikat["KEY_MAP"].nunique()
            )


            if len(gdf_duplikat) > 0:

                kolom_tampil = [
                    kolom
                    for kolom in [
                        "PROVINSI",
                        "NAMA_KAB",
                        "KEY_PROV",
                        "KEY_KAB",
                        "KEY_MAP"
                    ]
                    if kolom in gdf_duplikat.columns
                ]

                display(
                    gdf_duplikat[
                        kolom_tampil
                    ]
                    .sort_values(
                        "KEY_MAP"
                    )
                    .reset_index(drop=True)
                )

            else:

                print(
                    "✅ Tidak ditemukan duplicate KEY_MAP."
                )


            # ============================================================
            # 4. DISSOLVE BERDASARKAN KEY_MAP
            #
            # PENTING:
            # Jangan pakai drop_duplicates().
            #
            # Dissolve mempertahankan seluruh bagian geometri dan
            # menggabungkannya menjadi Polygon/MultiPolygon.
            # ============================================================

            kolom_agg = {}

            for kolom in [
                "PROVINSI",
                "NAMA_KAB",
                "KEY_PROV",
                "KEY_KAB"
            ]:
                if kolom in gdf_kabupaten.columns:
                    kolom_agg[kolom] = "first"


            gdf_kabupaten = (
                gdf_kabupaten
                .dissolve(
                    by="KEY_MAP",
                    as_index=False,
                    aggfunc=kolom_agg
                )
                .reset_index(drop=True)
            )


            # ============================================================
            # 5. PASTIKAN CRS
            # ============================================================

            if gdf_kabupaten.crs is None:

                gdf_kabupaten = (
                    gdf_kabupaten
                    .set_crs("EPSG:4326")
                )

            else:

                gdf_kabupaten = (
                    gdf_kabupaten
                    .to_crs("EPSG:4326")
                )


            # ============================================================
            # 6. VALIDASI DUPLIKAT SETELAH DISSOLVE
            # ============================================================

            jumlah_duplikat_sesudah = int(
                gdf_kabupaten[
                    "KEY_MAP"
                ]
                .duplicated()
                .sum()
            )


            print("\n" + "=" * 75)
            print("HASIL DISSOLVE ADM2")
            print("=" * 75)

            print(
                "Feature sebelum :",
                len(gdf_kabupaten_sebelum_dissolve)
            )

            print(
                "Feature sesudah :",
                len(gdf_kabupaten)
            )

            print(
                "Feature digabung:",
                (
                    len(gdf_kabupaten_sebelum_dissolve)
                    - len(gdf_kabupaten)
                )
            )

            print(
                "KEY_MAP duplikat:",
                jumlah_duplikat_sesudah
            )


            if jumlah_duplikat_sesudah > 0:

                raise ValueError(
                    "Masih terdapat KEY_MAP duplikat "
                    "setelah dissolve."
                )

            else:

                print()
                print(
                    "✅ Semua KEY_MAP boundary ADM2 "
                    "sekarang unik."
                )


            # ============================================================
            # 7. CEK GEOMETRI
            # ============================================================

            jumlah_geometry_kosong = int(
                (
                    gdf_kabupaten.geometry.isna()
                    | gdf_kabupaten.geometry.is_empty
                ).sum()
            )


            jumlah_invalid = int(
                (~gdf_kabupaten.geometry.is_valid)
                .sum()
            )


            print("\n" + "=" * 75)
            print("VALIDASI GEOMETRI")
            print("=" * 75)

            print(
                "Geometry kosong :",
                jumlah_geometry_kosong
            )

            print(
                "Geometry invalid:",
                jumlah_invalid
            )


            # ============================================================
            # 8. PERBAIKI GEOMETRY INVALID JIKA ADA
            # ============================================================

            if jumlah_invalid > 0:

                print(
                    "\nMemperbaiki geometry invalid..."
                )

                try:

                    gdf_kabupaten[
                        "geometry"
                    ] = (
                        gdf_kabupaten.geometry
                        .make_valid()
                    )

                except Exception:

                    gdf_kabupaten[
                        "geometry"
                    ] = (
                        gdf_kabupaten.geometry
                        .buffer(0)
                    )


            # ============================================================
            # 9. RESET ID UNTUK PLOTLY
            # ============================================================

            gdf_kabupaten = (
                gdf_kabupaten
                .reset_index(drop=True)
            )


            gdf_kabupaten[
                "MAP_ID"
            ] = (
                gdf_kabupaten.index
                .astype(str)
            )


            # ============================================================
            # 10. RINGKASAN FINAL CELL 3
            # ============================================================

            print("\n" + "=" * 75)
            print("BOUNDARY ADM2 SIAP")
            print("=" * 75)

            print(
                "Wilayah unik:",
                gdf_kabupaten[
                    "KEY_MAP"
                ].nunique()
            )

            print(
                "Jumlah feature:",
                len(gdf_kabupaten)
            )

            print(
                "Duplikat:",
                gdf_kabupaten[
                    "KEY_MAP"
                ].duplicated().sum()
            )

            print(
                "Provinsi:",
                gdf_kabupaten[
                    "KEY_PROV"
                ].nunique()
            )

            print("=" * 75)

            # ============================================================
            # CELL 2C
            # REBUILD BOUNDARY PLOT SETELAH DISSOLVE
            # ============================================================

            import json


            # ============================================================
            # 1. REBUILD KABUPATEN/KOTA PLOT
            # ============================================================

            gdf_kab_plot = (
                gdf_kabupaten
                .copy()
                .reset_index(drop=True)
            )


            # Optional simplification
            if TOLERANCE_ADM2 > 0:

                gdf_kab_plot = (
                    gdf_kab_plot
                    .to_crs("EPSG:3857")
                )

                gdf_kab_plot["geometry"] = (
                    gdf_kab_plot.geometry.simplify(
                        tolerance=TOLERANCE_ADM2,
                        preserve_topology=True
                    )
                )

                gdf_kab_plot = (
                    gdf_kab_plot
                    .to_crs("EPSG:4326")
                )


            # ============================================================
            # 2. RESET MAP ID
            # ============================================================

            gdf_kab_plot = (
                gdf_kab_plot
                .reset_index(drop=True)
            )

            gdf_kab_plot["MAP_ID"] = (
                gdf_kab_plot.index.astype(str)
            )


            # ============================================================
            # 3. KEEP KOLOM YANG DIPERLUKAN
            # ============================================================

            gdf_kab_plot = gdf_kab_plot[
                [
                    "PROVINSI",
                    "NAMA_KAB",
                    "KEY_PROV",
                    "KEY_KAB",
                    "KEY_MAP",
                    "MAP_ID",
                    "geometry",
                ]
            ].copy()


            # ============================================================
            # 4. VALIDASI FINAL
            # ============================================================

            assert (
                gdf_kab_plot["KEY_MAP"]
                .duplicated()
                .sum()
                == 0
            ), "Masih ada KEY_MAP duplikat."


            print("=" * 70)
            print("BOUNDARY KAB/KOTA FINAL")
            print("=" * 70)

            print(
                "Jumlah wilayah:",
                len(gdf_kab_plot)
            )

            print(
                "KEY_MAP unik:",
                gdf_kab_plot["KEY_MAP"].nunique()
            )

            print(
                "Duplikat:",
                gdf_kab_plot["KEY_MAP"]
                .duplicated()
                .sum()
            )


            # ============================================================
            # 5. REBUILD GEOJSON KABUPATEN/KOTA
            # ============================================================

            geojson_kab = json.loads(
                gdf_kab_plot.to_json()
            )

            geojson_kab_json = json.dumps(
                geojson_kab,
                ensure_ascii=False,
                separators=(",", ":")
            )


            ukuran_kab_mb = (
                len(
                    geojson_kab_json.encode("utf-8")
                )
                / 1024**2
            )


            print(
                f"Ukuran GeoJSON kab/kota: "
                f"{ukuran_kab_mb:.2f} MB"
            )


            # ============================================================
            # 6. PENTING:
            # REBUILD PAYLOAD JIKA PAYLOAD SUDAH DIBUAT SEBELUM DISSOLVE
            # ============================================================

            print()
            print(
                "✅ gdf_kab_plot dan geojson_kab "
                "sudah menggunakan boundary hasil dissolve."
            )

            print(
                "Jika payload dibuat sebelum Cell dissolve, "
                "jalankan ulang cell pembentukan payload sekarang."
            )

            boundary_keys_order = (
                gdf_kab_plot["KEY_MAP"].tolist()
            )

            # ============================================================
            # CELL 3 — LANGKAH 21-30
            # DASHBOARD HTML, LAPORAN, ZIP, DAN DOWNLOAD
            # ============================================================
            from plotly.offline import get_plotlyjs

            # ============================================================
            # LANGKAH 21 — SIAPKAN OUTPUT
            # ============================================================

            if FOLDER_OUTPUT.exists():
                shutil.rmtree(FOLDER_OUTPUT)

            FOLDER_OUTPUT.mkdir(
                parents=True,
                exist_ok=True
            )

            if FILE_ZIP.exists():
                FILE_ZIP.unlink()


            # ============================================================
            # LANGKAH 22 — SIMPAN LAPORAN
            # ============================================================

            with pd.ExcelWriter(
                FILE_LAPORAN,
                engine="openpyxl"
            ) as writer:
                df_matching.to_excel(
                    writer,
                    sheet_name="Matching Wilayah",
                    index=False
                )

                df_log_indikator.to_excel(
                    writer,
                    sheet_name="Indikator",
                    index=False
                )

                gdf_prov_plot[
                    [
                        "PROVINSI",
                        "KEY_PROV"
                    ]
                ].to_excel(
                    writer,
                    sheet_name="Boundary Provinsi",
                    index=False
                )

                gdf_kab_plot[
                    [
                        "PROVINSI",
                        "NAMA_KAB",
                        "KEY_MAP"
                    ]
                ].to_excel(
                    writer,
                    sheet_name="Boundary Kab Kota",
                    index=False
                )


            # ============================================================
            # LANGKAH 23 — DATA JAVASCRIPT
            # ============================================================

            payload_json = json.dumps(
                payload,
                ensure_ascii=False,
                separators=(",", ":")
            )

            locations_prov_json = json.dumps(
                gdf_prov_plot["MAP_ID"]
                .astype(str)
                .tolist(),
                ensure_ascii=False
            )

            locations_kab_json = json.dumps(
                gdf_kab_plot["MAP_ID"]
                .astype(str)
                .tolist(),
                ensure_ascii=False
            )

            nama_prov_json = json.dumps(
                gdf_prov_plot["PROVINSI"]
                .astype(str)
                .tolist(),
                ensure_ascii=False
            )

            nama_kab_json = json.dumps(
                (
                    gdf_kab_plot["NAMA_KAB"]
                    .astype(str)
                    + ", "
                    + gdf_kab_plot["PROVINSI"]
                    .astype(str)
                ).tolist(),
                ensure_ascii=False
            )


            # ============================================================
            # LANGKAH 24 — DROPDOWN
            # ============================================================

            opsi_tahun = "".join(
                f'<option value="{tahun}">{tahun}</option>'
                for tahun in TAHUN_LIST
            )

            opsi_level = (
                '<option value="Provinsi">Provinsi</option>'
                '<option value="Kabupaten/Kota">'
                'Kabupaten/Kota</option>'
            )

            opsi_indikator = "".join(
                (
                    f'<option value="{html.escape(nama)}">'
                    f'{html.escape(nama)}</option>'
                )
                for nama in indikator
            )


            # ============================================================
            # LANGKAH 25 — PLOTLY SCRIPT
            # ============================================================

            if DASHBOARD_OFFLINE:
                plotly_script = (
                    "<script>"
                    + get_plotlyjs()
                    + "</script>"
                )
            else:
                plotly_script = (
                    '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>'
                )


            # ============================================================
            # LANGKAH 26 — TEMPLATE DASHBOARD
            # ============================================================

            dashboard_template = """
            <!DOCTYPE html>
            <html lang="id">
            <head>
            <meta charset="UTF-8">
            <meta
                name="viewport"
                content="width=device-width, initial-scale=1.0"
            >
            <title>Dashboard Peta Surveilans MR</title>

            __PLOTLY_SCRIPT__

            <style>
            * {
                box-sizing: border-box;
            }

            body {
                margin: 0;
                font-family: Arial, sans-serif;
                background: #F3F7F7;
                color: #253238;
            }

            .header {
                padding: 22px 28px;
                color: white;
                background: linear-gradient(
                    135deg,
                    #006F6F,
                    #00A0A0
                );
            }

            .header h1 {
                margin: 0 0 6px;
                font-size: 25px;
            }

            .header p {
                margin: 0;
                opacity: 0.92;
            }

            .container {
                max-width: 1500px;
                margin: auto;
                padding: 18px;
            }

            .controls {
                display: grid;
                grid-template-columns:
                    minmax(120px, 0.6fr)
                    minmax(180px, 0.9fr)
                    minmax(300px, 2.5fr);
                gap: 14px;
                padding: 17px;
                background: white;
                border-radius: 12px;
                box-shadow: 0 3px 15px rgba(0,0,0,0.08);
            }

            label {
                display: block;
                margin-bottom: 7px;
                font-size: 13px;
                font-weight: bold;
                color: #36585B;
            }

            select {
                width: 100%;
                padding: 11px;
                border: 1px solid #B7CCCC;
                border-radius: 8px;
                background: white;
                font-size: 14px;
            }

            .summary {
                display: grid;
                grid-template-columns:
                    repeat(4, minmax(120px, 1fr));
                gap: 12px;
                margin-top: 14px;
            }

            .card {
                padding: 13px;
                background: white;
                border-radius: 10px;
                box-shadow: 0 3px 12px rgba(0,0,0,0.07);
            }

            .card-label {
                color: #577173;
                font-size: 12px;
            }

            .card-value {
                margin-top: 4px;
                color: #007878;
                font-size: 21px;
                font-weight: bold;
            }

            .status {
                margin-top: 12px;
                padding: 10px 13px;
                border-radius: 8px;
                background: #E7F5F5;
                color: #27595B;
                font-size: 13px;
            }

            .map-card {
                margin-top: 14px;
                padding: 6px;
                background: white;
                border-radius: 12px;
                box-shadow: 0 3px 15px rgba(0,0,0,0.08);
            }

            #map {
                width: 100%;
                min-height: 700px;
            }

            @media (max-width: 800px) {
                .controls,
                .summary {
                    grid-template-columns: 1fr;
                }

                .container {
                    padding: 10px;
                }
            }
            </style>
            </head>

            <body>

            <div class="header">
                <h1>Dashboard Peta Surveilans Campak-Rubella</h1>
                <p>Provinsi dan Kabupaten/Kota | 2023–2025</p>
            </div>

            <div class="container">

                <div class="controls">
                    <div>
                        <label for="tahun">Tahun</label>
                        <select id="tahun">
                            __OPSI_TAHUN__
                        </select>
                    </div>

                    <div>
                        <label for="level">Level wilayah</label>
                        <select id="level">
                            __OPSI_LEVEL__
                        </select>
                    </div>

                    <div>
                        <label for="indikator">Indikator</label>
                        <select id="indikator">
                            __OPSI_INDIKATOR__
                        </select>
                    </div>
                </div>

                <div class="summary">
                    <div class="card">
                        <div class="card-label">Total wilayah</div>
                        <div id="jumlahWilayah" class="card-value">-</div>
                    </div>

                    <div class="card">
                        <div class="card-label">Dengan data</div>
                        <div id="denganData" class="card-value">-</div>
                    </div>

                    <div class="card">
                        <div class="card-label">Tanpa data</div>
                        <div id="tanpaData" class="card-value">-</div>
                    </div>

                    <div class="card">
                        <div class="card-label">Nilai maksimum</div>
                        <div id="nilaiMaksimum" class="card-value">-</div>
                    </div>
                </div>

                <div id="status" class="status">
                    Memuat peta...
                </div>

                <div class="map-card">
                    <div id="map"></div>
                </div>

            </div>

            <script>
            const payload = __PAYLOAD__;

            const geoProv = __GEO_PROV__;
            const geoKab = __GEO_KAB__;

            const locationsProv = __LOC_PROV__;
            const locationsKab = __LOC_KAB__;

            const namesProv = __NAMA_PROV__;
            const namesKab = __NAMA_KAB__;

            const tahunEl = document.getElementById("tahun");
            const levelEl = document.getElementById("level");
            const indikatorEl = document.getElementById("indikator");
            const statusEl = document.getElementById("status");

            function makeKey() {
                return (
                    tahunEl.value
                    + "|||"
                    + levelEl.value
                    + "|||"
                    + indikatorEl.value
                );
            }

            function availableIndicators() {
                const result = new Set();

                Object.keys(payload).forEach(function(key) {
                    const parts = key.split("|||");

                    if (
                        parts[0] === tahunEl.value
                        && parts[1] === levelEl.value
                    ) {
                        result.add(parts[2]);
                    }
                });

                return result;
            }

            function updateIndicatorOptions() {
                const available = availableIndicators();
                let selectedValid = false;

                Array.from(indikatorEl.options)
                    .forEach(function(option) {
                        option.disabled = !available.has(
                            option.value
                        );

                        if (
                            option.selected
                            && !option.disabled
                        ) {
                            selectedValid = true;
                        }
                    });

                if (!selectedValid) {
                    const firstValid = Array.from(
                        indikatorEl.options
                    ).find(
                        option => !option.disabled
                    );

                    if (firstValid) {
                        indikatorEl.value =
                            firstValid.value;
                    }
                }
            }

            function formatNumber(value, type) {
                if (
                    value === null
                    || value === undefined
                    || Number.isNaN(value)
                ) {
                    return "Data tidak tersedia";
                }

                if (type === "Persentase") {
                    return value.toLocaleString(
                        "id-ID",
                        {
                            minimumFractionDigits: 1,
                            maximumFractionDigits: 1
                        }
                    ) + "%";
                }

                if (type === "Rate") {
                    return value.toLocaleString(
                        "id-ID",
                        {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2
                        }
                    );
                }

                return value.toLocaleString(
                    "id-ID",
                    {
                        maximumFractionDigits: 2
                    }
                );
            }

            function renderMap() {
                updateIndicatorOptions();

                const item = payload[makeKey()];

                if (!item) {
                    Plotly.purge("map");

                    statusEl.textContent =
                        "Indikator tidak tersedia.";

                    return;
                }

                const isProv =
                    levelEl.value === "Provinsi";

                const geojson =
                    isProv ? geoProv : geoKab;

                const locations =
                    isProv ? locationsProv : locationsKab;

                const names =
                    isProv ? namesProv : namesKab;

                const noDataLocations = [];
                const noDataCustom = [];

                const dataLocations = [];
                const dataZ = [];
                const dataCustom = [];

                item.values.forEach(
                    function(value, index) {
                        const formatted = formatNumber(
                            value,
                            item.type
                        );

                        if (value === null) {
                            noDataLocations.push(
                                locations[index]
                            );

                            noDataCustom.push([
                                names[index],
                                formatted
                            ]);

                        } else {
                            dataLocations.push(
                                locations[index]
                            );

                            dataZ.push(value);

                            dataCustom.push([
                                names[index],
                                formatted
                            ]);
                        }
                    }
                );

                const traces = [];

                if (noDataLocations.length > 0) {
                    traces.push({
                        type: "choropleth",
                        geojson: geojson,
                        featureidkey: "properties.MAP_ID",
                        locations: noDataLocations,
                        z: noDataLocations.map(() => 0),
                        zmin: 0,
                        zmax: 1,
                        colorscale: [
                            [0, "__WARNA_NODATA__"],
                            [1, "__WARNA_NODATA__"]
                        ],
                        customdata: noDataCustom,
                        hovertemplate:
                            "<b>%{customdata[0]}</b><br>"
                            + "%{customdata[1]}"
                            + "<extra></extra>",
                        marker: {
                            line: {
                                color: "white",
                                width: 0.35
                            }
                        },
                        showscale: false
                    });
                }

                if (dataLocations.length > 0) {
                    traces.push({
                        type: "choropleth",
                        geojson: geojson,
                        featureidkey: "properties.MAP_ID",
                        locations: dataLocations,
                        z: dataZ,
                        zmin: item.zmin,
                        zmax: item.zmax,
                        colorscale: [
                            [0.0, "__WARNA_RENDAH__"],
                            [0.5, "__WARNA_TENGAH__"],
                            [1.0, "__WARNA_TINGGI__"]
                        ],
                        customdata: dataCustom,
                        hovertemplate:
                            "<b>%{customdata[0]}</b><br>"
                            + indikatorEl.value
                            + ": %{customdata[1]}"
                            + "<extra></extra>",
                        marker: {
                            line: {
                                color: "white",
                                width: 0.35
                            }
                        },
                        colorbar: {
                            title: {
                                text: item.type
                            },
                            thickness: 15,
                            len: 0.68
                        },
                        showscale: true
                    });
                }

                const layout = {
                    title: {
                        text:
                            "<b>"
                            + indikatorEl.value
                            + "</b><br><sup>"
                            + levelEl.value
                            + " | "
                            + tahunEl.value
                            + "</sup>",
                        x: 0.5,
                        xanchor: "center"
                    },

                    geo: {
                        fitbounds: "locations",
                        visible: false
                    },

                    height: 700,

                    margin: {
                        l: 5,
                        r: 50,
                        t: 75,
                        b: 5
                    },

                    paper_bgcolor: "white",

                    font: {
                        family: "Arial",
                        size: 12,
                        color: "#253238"
                    }
                };

                const config = {
                    responsive: true,
                    displaylogo: false,
                    scrollZoom: true,
                    toImageButtonOptions: {
                        format: "png",
                        filename:
                            tahunEl.value
                            + "_"
                            + levelEl.value
                            + "_"
                            + indikatorEl.value,
                        height: 900,
                        width: 1400,
                        scale: 2
                    }
                };

                Plotly.react(
                    "map",
                    traces,
                    layout,
                    config
                );

                const validValues = item.values.filter(
                    value => value !== null
                );

                document.getElementById(
                    "jumlahWilayah"
                ).textContent = item.values.length;

                document.getElementById(
                    "denganData"
                ).textContent = validValues.length;

                document.getElementById(
                    "tanpaData"
                ).textContent = (
                    item.values.length
                    - validValues.length
                );

                document.getElementById(
                    "nilaiMaksimum"
                ).textContent = (
                    validValues.length > 0
                    ? formatNumber(
                        Math.max(...validValues),
                        item.type
                    )
                    : "-"
                );

                statusEl.textContent =
                    tahunEl.value
                    + " | "
                    + levelEl.value
                    + " | "
                    + indikatorEl.value
                    + " | Kolom sumber: "
                    + item.source_column;
            }

            tahunEl.addEventListener(
                "change",
                renderMap
            );

            levelEl.addEventListener(
                "change",
                renderMap
            );

            indikatorEl.addEventListener(
                "change",
                renderMap
            );

            renderMap();
            </script>
            </body>
            </html>
            """


            # ============================================================
            # LANGKAH 27 — ISI TEMPLATE
            # ============================================================

            dashboard_html = dashboard_template

            replacement = {
                "__PLOTLY_SCRIPT__":
                    plotly_script,

                "__OPSI_TAHUN__":
                    opsi_tahun,

                "__OPSI_LEVEL__":
                    opsi_level,

                "__OPSI_INDIKATOR__":
                    opsi_indikator,

                "__PAYLOAD__":
                    payload_json,

                "__GEO_PROV__":
                    geojson_prov_json,

                "__GEO_KAB__":
                    geojson_kab_json,

                "__LOC_PROV__":
                    locations_prov_json,

                "__LOC_KAB__":
                    locations_kab_json,

                "__NAMA_PROV__":
                    nama_prov_json,

                "__NAMA_KAB__":
                    nama_kab_json,

                "__WARNA_RENDAH__":
                    WARNA_RENDAH,

                "__WARNA_TENGAH__":
                    WARNA_TENGAH,

                "__WARNA_TINGGI__":
                    WARNA_TINGGI,

                "__WARNA_NODATA__":
                    WARNA_NODATA,
            }

            for token, value in replacement.items():
                dashboard_html = (
                    dashboard_html.replace(
                        token,
                        value
                    )
                )


            # ============================================================
            # LANGKAH 28 — SIMPAN DASHBOARD DAN README
            # ============================================================

            with open(
                FILE_DASHBOARD,
                "w",
                encoding="utf-8"
            ) as file:
                file.write(dashboard_html)

            readme = """
            DASHBOARD PETA SURVEILANS CAMPAK-RUBELLA
            ========================================

            1. Ekstrak ZIP.
            2. Buka Dashboard_Peta_Surveilans_MR.html.
            3. Gunakan Chrome, Edge, atau Firefox.
            4. Pilih tahun, level wilayah, dan indikator.

            Wilayah abu-abu berarti data tidak tersedia pada kombinasi
            tahun, level, dan indikator yang dipilih.

            Laporan_Matching.xlsx berisi:
            - validasi matching nama wilayah;
            - ketersediaan indikator;
            - daftar boundary yang digunakan.
            """

            with open(
                FOLDER_OUTPUT / "README.txt",
                "w",
                encoding="utf-8"
            ) as file:
                file.write(readme)


            # ============================================================
            # LANGKAH 29 — BUAT ZIP
            # ============================================================

            with zipfile.ZipFile(
                FILE_ZIP,
                mode="w",
                compression=zipfile.ZIP_DEFLATED,
                allowZip64=True
            ) as zip_file:
                for path in FOLDER_OUTPUT.rglob("*"):
                    if path.is_file():
                        zip_file.write(
                            path,
                            arcname=path.name
                        )

            ukuran_html = (
                FILE_DASHBOARD.stat().st_size
                / 1024**2
            )

            ukuran_zip = (
                FILE_ZIP.stat().st_size
                / 1024**2
            )


            # ============================================================
            # LANGKAH 30 — PREVIEW DAN DOWNLOAD
            # ============================================================

            print("\n" + "=" * 75)
            print("SELESAI")
            print("=" * 75)

            print(
                "Provinsi boundary:",
                len(gdf_prov_plot)
            )

            print(
                "Kab/Kota boundary:",
                len(gdf_kab_plot)
            )

            print(
                "Kombinasi data:",
                len(payload)
            )

            print(
                f"Ukuran dashboard: "
                f"{ukuran_html:.2f} MB"
            )

            print(
                f"Ukuran ZIP: "
                f"{ukuran_zip:.2f} MB"
            )

            print(
                "Dashboard:",
                FILE_DASHBOARD
            )

            print(
                "Laporan:",
                FILE_LAPORAN
            )

            st.success("Dashboard berhasil dibuat.")
            st.info(f"Dashboard: {FILE_DASHBOARD.name} | ZIP: {FILE_ZIP.name}")
            st.download_button("Download paket dashboard (ZIP)", data=FILE_ZIP.read_bytes(), file_name=FILE_ZIP.name, mime="application/zip")
            st.download_button("Download dashboard HTML", data=FILE_DASHBOARD.read_bytes(), file_name=FILE_DASHBOARD.name, mime="text/html")
            st.download_button("Download laporan matching Excel", data=FILE_LAPORAN.read_bytes(), file_name=FILE_LAPORAN.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Proses gagal: {e}")
            st.code(traceback.format_exc())
