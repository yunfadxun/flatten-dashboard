# -*- coding: utf-8 -*-
"""Tahap 2-3 Streamlit Surveilans MR.

Tempelkan file ini sebagai module, lalu pada app.py gunakan:
    from streamlit_tahap_2_3 import init_state_tahap_2_3, halaman_tahap_2, halaman_tahap_3

Prasyarat Tahap 1:
    st.session_state["hasil_flatten"] = [
        {"nama_input": "...2023...xlsx", "data": b"..."}, ...
    ]
Workbook hasil flatten harus memiliki sheet:
    - Tabel Analisis
    - Analisis Pusat by Kab
"""

import os
import re
import shutil
import subprocess
import unicodedata
from difflib import get_close_matches
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

try:
    import geopandas as gpd
except ImportError:
    gpd = None

os.environ.setdefault("OGR_GEOJSON_MAX_OBJ_SIZE", "0")

GITHUB_REPO_URL = "https://github.com/dmxsan/indonesia-admin-boundaries.git"
CACHE_ROOT = Path(".streamlit_cache")
REPO_DIR = CACHE_ROOT / "indonesia-admin-boundaries"
ADM1_DIR = REPO_DIR / "processed-data" / "02-provinces" / "province-only"
ADM2_DIR = REPO_DIR / "processed-data" / "03-districts"

SHEET_PROVINSI = "Tabel Analisis"
SHEET_KABUPATEN = "Analisis Pusat by Kab"
KOLOM_PROVINSI = "Kabupaten / Kota / Provinsi"
KOLOM_PROVINSI_KAB = "Provinsi"
KOLOM_KABUPATEN = "Kabupaten / Kota"
TOLERANCE_ADM1 = 1500
TOLERANCE_ADM2 = 1200

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


def init_state_tahap_2_3():
    defaults = {
        "hasil_boundary": None,
        "hasil_matching": None,
        "hasil_validasi_indikator": None,
        "payload_dashboard": None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def teks_dasar(teks):
    if teks is None or pd.isna(teks):
        return ""
    teks = str(teks).strip().upper().replace("_", " ")
    teks = unicodedata.normalize("NFKD", teks)
    teks = "".join(c for c in teks if not unicodedata.combining(c))
    teks = teks.replace("&", " DAN ")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", teks)).strip()


def normalisasi_provinsi(teks):
    hasil = re.sub(r"^(PROVINSI|PROPINSI|PROV)\s+", "", teks_dasar(teks)).strip()
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


ALIAS_KABUPATEN = {
    "KAB KEP SERIBU": "ADM KEP SERIBU",
    "KABUPATEN KEPULAUAN SERIBU": "ADM KEP SERIBU",
    "KEPULAUAN SERIBU": "ADM KEP SERIBU",
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


def normalisasi_kabupaten_excel(teks):
    hasil = teks_dasar(teks)
    hasil = re.sub(r"^KAB\s+", "KABUPATEN ", hasil)
    hasil = re.sub(r"^(KOTA ADM|KOTA ADMINISTRASI|KOTAMADYA)\s+", "KOTA ", hasil)
    hasil = re.sub(r"^KABUPATEN\s+", "", hasil)
    return ALIAS_KABUPATEN.get(hasil, hasil)


def normalisasi_kabupaten_boundary(teks):
    return re.sub(r"^KABUPATEN\s+", "", teks_dasar(teks)).strip()


def normalisasi_kolom(teks):
    return re.sub(r"\s+", " ", str(teks).strip().casefold())


def identifikasi_tahun(nama_file):
    hasil = re.findall(r"(?<!\d)(20\d{2})(?!\d)", str(nama_file))
    return int(hasil[0]) if hasil else None


def _run_git(args):
    proses = subprocess.run(args, capture_output=True, text=True, check=False)
    if proses.returncode != 0:
        raise RuntimeError(proses.stderr.strip() or proses.stdout.strip() or "Perintah Git gagal")
    return proses.stdout.strip()


def siapkan_repository_boundary(force_refresh=False):
    if shutil.which("git") is None:
        raise RuntimeError("Git tidak tersedia pada server Streamlit.")
    CACHE_ROOT.mkdir(parents=True, exist_ok=True)
    if force_refresh and REPO_DIR.exists():
        shutil.rmtree(REPO_DIR)
    if not (REPO_DIR / ".git").exists():
        if REPO_DIR.exists():
            shutil.rmtree(REPO_DIR)
        _run_git(["git", "clone", "--depth", "1", GITHUB_REPO_URL, str(REPO_DIR)])
        status = "Repository berhasil di-clone."
    else:
        # Tidak wajib pull setiap rerun. Tombol refresh memanggil force_refresh=True.
        status = "Repository lokal digunakan."
    if not ADM1_DIR.exists() or not ADM2_DIR.exists():
        raise FileNotFoundError("Struktur processed-data pada repository tidak sesuai konfigurasi.")
    files_adm1 = sorted(ADM1_DIR.rglob("*.geojson"))
    files_adm2 = sorted(ADM2_DIR.rglob("*.geojson"))
    if not files_adm1 or not files_adm2:
        raise FileNotFoundError("File GeoJSON ADM1 atau ADM2 tidak ditemukan.")
    return {"status_repo": status, "files_adm1": files_adm1, "files_adm2": files_adm2}


def baca_geojson_aman(path):
    try:
        gdf = gpd.read_file(path, engine="pyogrio")
    except Exception:
        gdf = gpd.read_file(path)
    if gdf.empty:
        return None
    gdf = gdf.set_crs("EPSG:4326") if gdf.crs is None else gdf.to_crs("EPSG:4326")
    gdf = gdf[gdf.geometry.notna() & ~gdf.geometry.is_empty].copy()
    if gdf.empty:
        return None
    invalid = ~gdf.geometry.is_valid
    if invalid.any():
        try:
            gdf.loc[invalid, "geometry"] = gdf.loc[invalid, "geometry"].make_valid()
        except Exception:
            gdf.loc[invalid, "geometry"] = gdf.loc[invalid, "geometry"].buffer(0)
    return gdf


def baca_boundary(files, level, progress=None):
    daftar, gagal = [], []
    total = len(files)
    for i, path in enumerate(files, 1):
        try:
            gdf = baca_geojson_aman(path)
            if gdf is None:
                raise ValueError("Geometri kosong")
            nama = path.stem.replace("_", " ").strip()
            if level == "Provinsi":
                gdf["PROVINSI"] = nama.upper()
                daftar.append(gdf[["PROVINSI", "geometry"]].copy())
            else:
                gdf["NAMA_KAB"] = nama
                daftar.append(gdf[["NAMA_KAB", "geometry"]].copy())
        except Exception as error:
            gagal.append({"Level": level, "File": path.name, "Error": str(error)})
        if progress:
            progress(i, total)
    if not daftar:
        raise ValueError(f"Tidak ada boundary {level} yang berhasil dibaca.")
    hasil = gpd.GeoDataFrame(pd.concat(daftar, ignore_index=True), geometry="geometry", crs="EPSG:4326")
    return hasil, pd.DataFrame(gagal)


def proses_boundary_lengkap(force_refresh=False):
    if gpd is None:
        raise ImportError("GeoPandas belum terpasang.")
    repo = siapkan_repository_boundary(force_refresh=force_refresh)
    bar1, txt1 = st.progress(0), st.empty()
    def p1(i, n):
        bar1.progress(i / n); txt1.caption(f"Membaca ADM1 {i}/{n}")
    prov, gagal1 = baca_boundary(repo["files_adm1"], "Provinsi", p1)
    prov["KEY_PROV"] = prov["PROVINSI"].apply(normalisasi_provinsi)
    prov = prov[["PROVINSI", "KEY_PROV", "geometry"]].dissolve(by="KEY_PROV", as_index=False, aggfunc="first")
    if len(prov) != 38:
        raise ValueError(f"Boundary provinsi harus 38, ditemukan {len(prov)}.")

    bar2, txt2 = st.progress(0), st.empty()
    def p2(i, n):
        bar2.progress(i / n); txt2.caption(f"Membaca ADM2 {i}/{n}")
    kab, gagal2 = baca_boundary(repo["files_adm2"], "Kabupaten/Kota", p2)
    kab["KEY_KAB"] = kab["NAMA_KAB"].apply(normalisasi_kabupaten_boundary)

    titik = kab[["geometry"]].copy()
    titik["ID_ADM2"] = titik.index
    titik["geometry"] = titik.geometry.representative_point()
    joined = gpd.sjoin(titik, prov[["PROVINSI", "KEY_PROV", "geometry"]], how="left", predicate="within")
    joined = joined.sort_values("ID_ADM2").drop_duplicates("ID_ADM2").set_index("ID_ADM2")
    kab["PROVINSI"] = kab.index.map(joined["PROVINSI"])
    kab["KEY_PROV"] = kab.index.map(joined["KEY_PROV"])
    missing = kab["KEY_PROV"].isna()
    if missing.any():
        contoh = kab.loc[missing, "NAMA_KAB"].head(10).tolist()
        raise ValueError(f"{int(missing.sum())} ADM2 belum mendapat provinsi. Contoh: {contoh}")
    kab["KEY_MAP"] = kab["KEY_PROV"] + "|" + kab["KEY_KAB"]
    dup_before = kab.loc[kab["KEY_MAP"].duplicated(keep=False), ["PROVINSI", "NAMA_KAB", "KEY_MAP"]].copy()
    awal = len(kab)
    kab = kab[["PROVINSI", "KEY_PROV", "NAMA_KAB", "KEY_KAB", "KEY_MAP", "geometry"]].dissolve(
        by="KEY_MAP", as_index=False,
        aggfunc={"PROVINSI": "first", "KEY_PROV": "first", "NAMA_KAB": "first", "KEY_KAB": "first"},
    ).reset_index(drop=True)
    if kab["KEY_MAP"].duplicated().any():
        raise ValueError("KEY_MAP masih duplikat setelah dissolve.")

    prov_plot = prov.to_crs("EPSG:3857").copy()
    prov_plot["geometry"] = prov_plot.geometry.simplify(TOLERANCE_ADM1, preserve_topology=True)
    prov_plot = prov_plot.to_crs("EPSG:4326").reset_index(drop=True)
    prov_plot["MAP_ID"] = prov_plot.index.astype(str)

    kab_plot = kab.to_crs("EPSG:3857").copy()
    kab_plot["geometry"] = kab_plot.geometry.simplify(TOLERANCE_ADM2, preserve_topology=True)
    kab_plot = kab_plot.to_crs("EPSG:4326").reset_index(drop=True)
    kab_plot["MAP_ID"] = kab_plot.index.astype(str)

    return {
        "repository": repo,
        "gdf_prov_plot": prov_plot[["PROVINSI", "KEY_PROV", "MAP_ID", "geometry"]],
        "gdf_kab_plot": kab_plot[["PROVINSI", "NAMA_KAB", "KEY_PROV", "KEY_KAB", "KEY_MAP", "MAP_ID", "geometry"]],
        "gagal_boundary": pd.concat([gagal1, gagal2], ignore_index=True),
        "duplikat_sebelum": dup_before,
        "feature_digabung": awal - len(kab),
    }


def _cari_sheet(sheet_names, target):
    norm = lambda x: re.sub(r"[^a-z0-9]", "", str(x).lower())
    t = norm(target)
    exact = [x for x in sheet_names if norm(x) == t]
    if exact:
        return exact[0]
    partial = [x for x in sheet_names if t in norm(x)]
    return partial[0] if partial else None


def _cari_kolom(df, target):
    norm = lambda x: re.sub(r"[^a-z0-9]", "", str(x).lower())
    t = norm(target)
    exact = [c for c in df.columns if norm(c) == t]
    if exact:
        return exact[0]
    part = [c for c in df.columns if t in norm(c)]
    return part[0] if len(part) == 1 else None


def baca_hasil_flatten():
    items = st.session_state.get("hasil_flatten", [])
    if not items:
        raise ValueError("Hasil Tahap 1 belum tersedia.")
    data_prov, data_kab, files, warnings = {}, {}, {}, []
    for item in items:
        nama = item.get("nama_input", "")
        tahun = identifikasi_tahun(nama)
        if tahun is None:
            raise ValueError(f"Tahun 20xx tidak ditemukan pada nama file: {nama}")
        if tahun in files:
            raise ValueError(f"Lebih dari satu file terdeteksi untuk tahun {tahun}.")
        xls = pd.ExcelFile(BytesIO(item["data"]), engine="openpyxl")
        s_prov = _cari_sheet(xls.sheet_names, SHEET_PROVINSI)
        s_kab = _cari_sheet(xls.sheet_names, SHEET_KABUPATEN)
        if not s_prov or not s_kab:
            raise ValueError(f"Sheet target tidak lengkap pada {nama}. Tersedia: {xls.sheet_names}")
        prov = pd.read_excel(BytesIO(item["data"]), sheet_name=s_prov, engine="openpyxl")
        kab = pd.read_excel(BytesIO(item["data"]), sheet_name=s_kab, engine="openpyxl")
        kp = _cari_kolom(prov, KOLOM_PROVINSI)
        kpk = _cari_kolom(kab, KOLOM_PROVINSI_KAB)
        kk = _cari_kolom(kab, KOLOM_KABUPATEN)
        if not kp or not kpk or not kk:
            raise KeyError(f"Kolom wilayah tidak lengkap pada {nama}. Provinsi={kp}, ProvinsiKab={kpk}, Kab={kk}")
        prov["KEY_PROV"] = prov[kp].apply(normalisasi_provinsi)
        prov = prov[(prov["KEY_PROV"] != "") & (prov["KEY_PROV"] != "INDONESIA")].copy()
        kab["KEY_PROV"] = kab[kpk].apply(normalisasi_provinsi)
        kab["KEY_KAB"] = kab[kk].apply(normalisasi_kabupaten_excel)
        kab = kab[(kab["KEY_PROV"] != "") & (kab["KEY_KAB"] != "") & (kab["KEY_KAB"] != "INDONESIA")].copy()
        kab["KEY_MAP"] = kab["KEY_PROV"] + "|" + kab["KEY_KAB"]
        if prov["KEY_PROV"].duplicated().any():
            warnings.append({"Tahun": tahun, "Level": "Provinsi", "Peringatan": "KEY_PROV duplikat"})
        if kab["KEY_MAP"].duplicated().any():
            warnings.append({"Tahun": tahun, "Level": "Kabupaten/Kota", "Peringatan": "KEY_MAP duplikat"})
        data_prov[tahun], data_kab[tahun], files[tahun] = prov, kab, nama
    return {"data_provinsi": data_prov, "data_kabupaten": data_kab, "tahun_list": sorted(files), "file_tahun": files, "peringatan": pd.DataFrame(warnings)}


def validasi_matching(data_excel, boundary):
    key_prov = set(boundary["gdf_prov_plot"]["KEY_PROV"])
    key_kab = set(boundary["gdf_kab_plot"]["KEY_MAP"])
    rows = []
    for tahun in data_excel["tahun_list"]:
        for key in sorted(data_excel["data_provinsi"][tahun]["KEY_PROV"].dropna().unique()):
            rows.append({"Tahun": tahun, "Level": "Provinsi", "Key": key, "Match": key in key_prov})
        for key in sorted(data_excel["data_kabupaten"][tahun]["KEY_MAP"].dropna().unique()):
            rows.append({"Tahun": tahun, "Level": "Kabupaten/Kota", "Key": key, "Match": key in key_kab})
    df = pd.DataFrame(rows)
    ringkasan = df.groupby(["Tahun", "Level", "Match"], dropna=False).size().reset_index(name="Jumlah")
    belum = df.loc[~df["Match"]].copy()
    diag = []
    unik = belum.loc[belum["Level"] == "Kabupaten/Kota", "Key"].drop_duplicates()
    for key in unik:
        if "|" not in key:
            continue
        prov, nama = key.split("|", 1)
        pilihan = boundary["gdf_kab_plot"].loc[boundary["gdf_kab_plot"]["KEY_PROV"] == prov, "KEY_KAB"].dropna().unique().tolist()
        dekat = get_close_matches(nama, pilihan, n=5, cutoff=0.25)
        row = {"PROVINSI": prov, "NAMA_EXCEL": nama}
        row.update({f"KANDIDAT_{i}": dekat[i-1] if len(dekat) >= i else "" for i in range(1, 6)})
        diag.append(row)
    return {"df_matching": df, "ringkasan_matching": ringkasan, "belum_match": belum, "diagnostik": pd.DataFrame(diag), "data_excel": data_excel}


def cari_kolom_indikator(df, nama):
    if nama in df.columns:
        return nama
    target = normalisasi_kolom(nama)
    exact = [c for c in df.columns if normalisasi_kolom(c) == target]
    if exact:
        return exact[0]
    prefix = [c for c in df.columns if normalisasi_kolom(c).startswith(target)]
    return prefix[0] if len(prefix) == 1 else None


def konversi_satu_nilai(nilai):
    if pd.isna(nilai):
        return np.nan
    if isinstance(nilai, (int, float, np.integer, np.floating)):
        return float(nilai)
    teks = str(nilai).strip().replace("%", "").replace(" ", "")
    if teks.casefold() in {"", "-", "--", "na", "n/a", "none", "nan", "null"}:
        return np.nan
    if "," in teks and "." in teks:
        teks = teks.replace(".", "").replace(",", ".") if teks.rfind(",") > teks.rfind(".") else teks.replace(",", "")
    elif "," in teks:
        bagian = teks.split(",")
        teks = teks.replace(",", ".") if len(bagian) == 2 and len(bagian[1]) <= 4 else teks.replace(",", "")
    try:
        return float(teks)
    except ValueError:
        return np.nan


def tipe_indikator(nama):
    n = normalisasi_kolom(nama)
    if any(x in n for x in ["investigasi adekuat", "spesimen adekuat", "timeliness", "laboratory confirmation", "pos rate", "persen", "percentage"]):
        return "Persentase"
    if any(x in n for x in ["incidence rate", "discarded rate"]):
        return "Rate"
    return "Jumlah"


def validasi_dan_buat_payload(data_excel, boundary):
    indikator, ketersediaan = [], []
    for nama in INDIKATOR_PILIHAN:
        tersedia = False
        for tahun in data_excel["tahun_list"]:
            cp = cari_kolom_indikator(data_excel["data_provinsi"][tahun], nama)
            ck = cari_kolom_indikator(data_excel["data_kabupaten"][tahun], nama)
            ketersediaan.append({"Tahun": tahun, "Indikator": nama, "Kolom Provinsi": cp or "", "Kolom Kabupaten/Kota": ck or ""})
            tersedia = tersedia or cp is not None or ck is not None
        if tersedia:
            indikator.append(nama)
    if not indikator:
        raise ValueError("Tidak ada indikator pilihan yang ditemukan.")

    payload, log = {}, []
    for tahun in data_excel["tahun_list"]:
        configs = [
            ("Provinsi", data_excel["data_provinsi"][tahun], "KEY_PROV", boundary["gdf_prov_plot"]["KEY_PROV"].tolist()),
            ("Kabupaten/Kota", data_excel["data_kabupaten"][tahun], "KEY_MAP", boundary["gdf_kab_plot"]["KEY_MAP"].tolist()),
        ]
        for level, df, field_key, order in configs:
            for nama in indikator:
                kolom = cari_kolom_indikator(df, nama)
                key = f"{tahun}|||{level}|||{nama}"
                if kolom is None:
                    log.append({"Tahun": tahun, "Level": level, "Indikator": nama, "Status": "Tidak tersedia", "Kolom sumber": "", "Wilayah dengan data": 0})
                    continue
                lookup_df = pd.DataFrame({"KEY": df[field_key], "NILAI": df[kolom].apply(konversi_satu_nilai)})
                lookup_df = lookup_df[lookup_df["KEY"].notna() & (lookup_df["KEY"] != "")].copy()
                lookup_df["_NULL"] = lookup_df["NILAI"].isna().astype(int)
                lookup_df = lookup_df.sort_values(["KEY", "_NULL"]).drop_duplicates("KEY")
                lookup = dict(zip(lookup_df["KEY"], lookup_df["NILAI"]))
                values = [lookup.get(k, np.nan) for k in order]
                valid = [float(v) for v in values if pd.notna(v)]
                jenis = tipe_indikator(nama)
                if jenis == "Persentase":
                    zmin, zmax = 0.0, 100.0
                elif valid:
                    zmin, zmax = 0.0, float(np.nanpercentile(valid, 95))
                    if zmax <= 0:
                        zmax = max(max(valid), 1.0)
                else:
                    zmin, zmax = 0.0, 1.0
                payload[key] = {"values": [None if pd.isna(v) else float(v) for v in values], "zmin": zmin, "zmax": zmax, "type": jenis, "source_column": kolom}
                log.append({"Tahun": tahun, "Level": level, "Indikator": nama, "Status": "Tersedia", "Kolom sumber": kolom, "Wilayah dengan data": len(valid)})
    return {"indikator": indikator, "laporan_ketersediaan": pd.DataFrame(ketersediaan), "log_indikator": pd.DataFrame(log), "tahun_list": data_excel["tahun_list"], "jumlah_payload": len(payload)}, payload


def dataframe_dict_to_excel(sheets):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if df is None or df.empty:
                df = pd.DataFrame({"Informasi": ["Tidak ada data"]})
            df.to_excel(writer, sheet_name=name[:31], index=False)
    out.seek(0)
    return out.getvalue()


def halaman_tahap_2():
    init_state_tahap_2_3()
    st.title("Tahap 2 — Boundary & Matching")
    if not st.session_state.get("hasil_flatten"):
        st.warning("Tahap 1 belum selesai.")
        return
    st.info("Boundary diambil otomatis dari GitHub dmxsan/indonesia-admin-boundaries. Tidak perlu upload GeoJSON.")
    c1, c2 = st.columns([3, 1])
    run = c1.button("Jalankan Boundary & Matching", type="primary", use_container_width=True)
    refresh = c2.button("Refresh GitHub", use_container_width=True)
    if run or refresh:
        try:
            data = baca_hasil_flatten()
            boundary = proses_boundary_lengkap(force_refresh=refresh)
            matching = validasi_matching(data, boundary)
            st.session_state["hasil_boundary"] = boundary
            st.session_state["hasil_matching"] = matching
            st.session_state["hasil_validasi_indikator"] = None
            st.session_state["payload_dashboard"] = None
            st.rerun()
        except Exception as error:
            st.error("Tahap 2 gagal.")
            st.exception(error)
    b = st.session_state.get("hasil_boundary")
    m = st.session_state.get("hasil_matching")
    if b is None or m is None:
        return
    a, c, d = st.columns(3)
    a.metric("Provinsi", len(b["gdf_prov_plot"]))
    c.metric("Kabupaten/Kota", len(b["gdf_kab_plot"]))
    d.metric("Belum match", len(m["belum_match"]))
    st.dataframe(m["ringkasan_matching"], use_container_width=True, hide_index=True)
    with st.expander("Wilayah belum match", expanded=not m["belum_match"].empty):
        st.dataframe(m["belum_match"], use_container_width=True, hide_index=True)
    with st.expander("Diagnostik kandidat nama"):
        st.dataframe(m["diagnostik"], use_container_width=True, hide_index=True)
    report = dataframe_dict_to_excel({"Matching": m["df_matching"], "Ringkasan": m["ringkasan_matching"], "Belum Match": m["belum_match"], "Diagnostik": m["diagnostik"], "Gagal Boundary": b["gagal_boundary"], "Duplikat ADM2": b["duplikat_sebelum"]})
    st.download_button("Download Laporan Tahap 2", report, "Laporan_Tahap_2.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)


def halaman_tahap_3():
    init_state_tahap_2_3()
    st.title("Tahap 3 — Validasi Indikator")
    b = st.session_state.get("hasil_boundary")
    m = st.session_state.get("hasil_matching")
    if b is None or m is None:
        st.warning("Tahap 2 belum selesai.")
        return
    if st.button("Validasi Indikator dan Buat Payload", type="primary", use_container_width=True):
        try:
            hasil, payload = validasi_dan_buat_payload(m["data_excel"], b)
            st.session_state["hasil_validasi_indikator"] = hasil
            st.session_state["payload_dashboard"] = payload
            st.rerun()
        except Exception as error:
            st.error("Tahap 3 gagal.")
            st.exception(error)
    hasil = st.session_state.get("hasil_validasi_indikator")
    payload = st.session_state.get("payload_dashboard")
    if hasil is None or payload is None:
        return
    c1, c2, c3 = st.columns(3)
    c1.metric("Indikator ditemukan", len(hasil["indikator"]))
    c2.metric("Payload", hasil["jumlah_payload"])
    c3.metric("Tahun", len(hasil["tahun_list"]))
    t1, t2 = st.tabs(["Ketersediaan", "Log payload"])
    with t1:
        st.dataframe(hasil["laporan_ketersediaan"], use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(hasil["log_indikator"], use_container_width=True, hide_index=True)
    report = dataframe_dict_to_excel({"Ketersediaan": hasil["laporan_ketersediaan"], "Log Payload": hasil["log_indikator"], "Indikator": pd.DataFrame({"Indikator": hasil["indikator"]})})
    st.download_button("Download Laporan Tahap 3", report, "Laporan_Tahap_3.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
