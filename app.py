# -*- coding: utf-8 -*-
"""Streamlit Dashboard Surveilans Campak-Rubella.

Tahap 1: Upload workbook asli dan flatten header baris Excel 9-10.
Tahap 2: Ambil/baca boundary Indonesia, validasi, assign ADM2 ke ADM1, dissolve.
Tahap 3: Standardisasi data, matching wilayah, diagnostik, indikator dan payload.
Tahap 4: Dashboard peta interaktif.
Tahap 5: Ekspor laporan, data flatten, GeoJSON, dan ZIP hasil.
"""

import json
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
import zipfile
from difflib import get_close_matches
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import geopandas as gpd
import plotly.graph_objects as go
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

os.environ["OGR_GEOJSON_MAX_OBJ_SIZE"] = "0"

st.set_page_config(page_title="Dashboard Surveilans MR", page_icon="🗺️", layout="wide")

# ============================================================
# 1. KONFIGURASI
# ============================================================
GITHUB_REPO_URL = "https://github.com/dmxsan/indonesia-admin-boundaries.git"
SHEET_ANALISIS = "Tabel Analisis"
SHEET_KABUPATEN_ASLI = "Tabel Analisis Pusat by Kab"
SHEET_KABUPATEN_FLAT = "Analisis Pusat by Kab"
HEADER_START, HEADER_END, DATA_START = 8, 9, 10
MAX_FILE_SIZE_MB = 100
TOLERANCE_ADM1_DEFAULT = 1500
TOLERANCE_ADM2_DEFAULT = 1200
WARNA_RENDAH = "#E8F1F1"
WARNA_TENGAH = "#64BABA"
WARNA_TINGGI = "#008C8C"
WARNA_NODATA = "#D9D9D9"
INDIKATOR_PILIHAN = [
    "Jumlah Suspek", "Jumlah Meninggal", "Jumlah Suspek dengan spesimen",
    "Performa Surveilans_Measles Incidence Rate (<1/1.000.000 population)",
    "Performa Surveilans_Rubella Incidence Rate (<1/1.000.000 population)",
    "Performa Surveilans_DISCARDED RATE (≥2/100.000 population)",
    "Performa Surveilans_Pos Rate Measles",
    "Performa Surveilans_Investigasi Adekuat (%)",
    "Performa Surveilans_Spesimen Adekuat (%)",
    "Performa Surveilans_Timeliness Specimen Result (%)",
    "Performa Surveilans_Laboratory Confirmation (%)",
    "Klasifikasi_Campak Positif", "Klasifikasi_Rubella Positif",
    "Klasifikasi_Konfirmasi Hubungan Epidemiologi Campak",
    "Klasifikasi_Konfirmasi Hubungan Epidemiologi Rubella",
    "Klasifikasi_Negatif", "Klasifikasi_Klinis", "Klasifikasi_Pending",
]

DEFAULTS = {
    "flatten_results": [], "flatten_errors": [], "boundary": None,
    "standardized": None, "matching": None, "payload": None,
    "indicator_log": None, "selected_years": [],
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# 2. NORMALISASI DASAR
# ============================================================
def normalisasi_dasar(teks):
    if teks is None or pd.isna(teks): return ""
    hasil = str(teks).strip().upper()
    if not hasil or hasil in {"NAN", "NONE", "NULL"}: return ""
    hasil = unicodedata.normalize("NFKD", hasil)
    hasil = "".join(c for c in hasil if not unicodedata.combining(c))
    hasil = hasil.replace("&", " DAN ")
    hasil = re.sub(r"[/_,.;:()\-]+", " ", hasil)
    hasil = re.sub(r"[^A-Z0-9 ]+", " ", hasil)
    return re.sub(r"\s+", " ", hasil).strip()

def normalisasi_nama_sheet(nama):
    return re.sub(r"[^a-z0-9]", "", str(nama).lower())

def normalisasi_provinsi(teks):
    hasil = normalisasi_dasar(teks)
    if not hasil: return ""
    hasil = re.sub(r"^(PROVINSI|PROPINSI|PROV)\s+", "", hasil).strip()
    alias = {
        "JAKARTA":"DKI JAKARTA", "DK JAKARTA":"DKI JAKARTA",
        "DAERAH KHUSUS IBUKOTA JAKARTA":"DKI JAKARTA",
        "YOGYAKARTA":"DAERAH ISTIMEWA YOGYAKARTA",
        "DI YOGYAKARTA":"DAERAH ISTIMEWA YOGYAKARTA",
        "D I YOGYAKARTA":"DAERAH ISTIMEWA YOGYAKARTA",
        "BANGKA BELITUNG":"KEPULAUAN BANGKA BELITUNG",
        "KEP BANGKA BELITUNG":"KEPULAUAN BANGKA BELITUNG",
        "KEP RIAU":"KEPULAUAN RIAU", "NANGGROE ACEH DARUSSALAM":"ACEH",
        "IRIAN JAYA BARAT":"PAPUA BARAT",
    }
    return alias.get(hasil, hasil)

ALIAS_KAB = {
    "TOBA SAMOSIR":"TOBA", "MALUKU TENGGARA BARAT":"KEPULAUAN TANIMBAR",
    "MAMUJU UTARA":"PASANGKAYU", "GUNUNG KIDUL":"GUNUNGKIDUL",
    "PEMATANG SIANTAR":"PEMATANGSIANTAR", "PADANG SIDEMPUAN":"PADANGSIDIMPUAN",
    "PADANG SIDIMPUAN":"PADANGSIDIMPUAN", "DHARMAS RAYA":"DHARMASRAYA",
    "PARE PARE":"PAREPARE", "PANGKAJENE KEPULAUAN":"PANGKAJENE DAN KEPULAUAN",
    "SELAYAR":"KEPULAUAN SELAYAR", "BAU BAU":"BAUBAU",
    "YAPEN WAROPEN":"KEPULAUAN YAPEN", "POLIWALI MANDAR":"POLEWALI MANDAR",
    "TAMBRAW":"TAMBRAUW", "GUNUNG SITOLI":"GUNUNGSITOLI",
    "BANJAR BARU":"BANJARBARU", "JAKARTA PUSAT":"JAKARTAPUSAT",
    "JAKARTA UTARA":"JAKARTAUTARA", "JAKARTA BARAT":"JAKARTABARAT",
    "JAKARTA SELATAN":"JAKARTASELATAN", "JAKARTA TIMUR":"JAKARTATIMUR",
    "KEPULAUAN SERIBU":"KEPULAUANSERIBU",
    "SIAU TAGULANDANG BIARO":"SIAUTAGULANDANGBIARO",
}
def normalisasi_kabupaten(teks):
    hasil = normalisasi_dasar(teks)
    if not hasil: return ""
    hasil = re.sub(r"^(KABUPATEN|KAB|KOTA ADMINISTRASI|KOTA ADM|KOTAMADYA|KOTA)\s+", "", hasil).strip()
    return ALIAS_KAB.get(hasil, hasil).replace(" ", "")

def buat_key_kabupaten(provinsi, kabupaten):
    p, k = normalisasi_provinsi(provinsi), normalisasi_kabupaten(kabupaten)
    return f"{p}|{k}" if p and k else k

def cari_sheet(names, target):
    t = normalisasi_nama_sheet(target)
    for n in names:
        if normalisasi_nama_sheet(n) == t: return n
    for n in names:
        if t in normalisasi_nama_sheet(n): return n
    return None

def norm_kolom(x): return re.sub(r"[^a-z0-9]", "", str(x).lower())
def cari_kolom(df, kandidat):
    kandidat = [kandidat] if isinstance(kandidat, str) else kandidat
    for target in kandidat:
        for col in df.columns:
            if norm_kolom(col) == norm_kolom(target): return col
    for target in kandidat:
        for col in df.columns:
            if norm_kolom(target) and norm_kolom(target) in norm_kolom(col): return col
    return None

# ============================================================
# 3. TAHAP 1 - FLATTEN EXCEL
# ============================================================
def bersihkan_nilai_header(v):
    if pd.isna(v): return None
    v = re.sub(r"\s+", " ", str(v).strip())
    return v if v and v.lower() != "nan" else None

def buat_nama_kolom_unik(cols):
    seen, out = {}, []
    for col in cols:
        n = seen.get(col, 0); out.append(col if n == 0 else f"{col}_{n}"); seen[col] = n + 1
    return out

def flatten_excel_header(file_bytes, sheet_name):
    raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None, engine="openpyxl")
    if raw.empty or len(raw) <= HEADER_END: raise ValueError("Header Excel baris 9-10 tidak ditemukan.")
    h = raw.iloc[HEADER_START:HEADER_END+1].map(bersihkan_nilai_header).ffill(axis=1)
    cols = []
    for j in range(raw.shape[1]):
        parts = []
        for i in range(len(h)):
            v = h.iloc[i, j]
            if v is not None and str(v).strip() and str(v).strip() not in parts: parts.append(str(v).strip())
        cols.append("_".join(parts) if parts else f"Kolom_{j+1}")
    df = raw.iloc[DATA_START:].copy(); df.columns = buat_nama_kolom_unik(cols)
    return df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

def format_worksheet(ws, df):
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    fill = PatternFill(fill_type="solid", fgColor="1F4E78"); font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]: cell.fill, cell.font = fill, font
    for i, c in enumerate(df.columns, 1):
        lengths = [len(str(c))] + [len(str(v)) for v in df[c].head(200) if pd.notna(v)]
        ws.column_dimensions[get_column_letter(i)].width = min(max(max(lengths)+2, 10), 40)

def dfs_to_excel(sheets):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        used = set()
        for name, df in sheets.items():
            base = re.sub(r"[\\/*?:\[\]]", "_", str(name))[:31] or "Sheet"
            final, n = base, 1
            while final in used: final = f"{base[:27]}_{n}"; n += 1
            used.add(final)
            (df if df is not None and not df.empty else pd.DataFrame({"Informasi":["Tidak ada data"]})).to_excel(writer, sheet_name=final, index=False)
            format_worksheet(writer.sheets[final], df if df is not None and not df.empty else pd.DataFrame({"Informasi":["Tidak ada data"]}))
    return out.getvalue()

def proses_excel(file_bytes, name):
    xls = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    hasil, log = {}, []
    for target, output in [(SHEET_ANALISIS, SHEET_ANALISIS), (SHEET_KABUPATEN_ASLI, SHEET_KABUPATEN_FLAT)]:
        found = cari_sheet(xls.sheet_names, target)
        if not found:
            log.append({"File":name,"Sheet":target,"Status":"Tidak ditemukan","Keterangan":", ".join(xls.sheet_names)}); continue
        try:
            df = flatten_excel_header(file_bytes, found); hasil[output] = df
            log.append({"File":name,"Sheet":found,"Status":"Berhasil","Baris":len(df),"Kolom":len(df.columns)})
        except Exception as e: log.append({"File":name,"Sheet":found,"Status":"Gagal","Keterangan":str(e)})
    if not hasil: raise ValueError("Tidak ada sheet target yang berhasil diproses.")
    return {"nama_input":name, "tahun":deteksi_tahun(name), "sheets":hasil,
            "data":dfs_to_excel(hasil), "nama_output":f"{Path(name).stem}_Flatten.xlsx", "log":log}

def deteksi_tahun(name):
    ys = re.findall(r"(?<!\d)(20\d{2})(?!\d)", str(name))
    return int(ys[-1]) if ys else None

# ============================================================
# 4. TAHAP 2 - BOUNDARY
# ============================================================
def baca_geojson_aman(path):
    try: gdf = gpd.read_file(path, engine="pyogrio")
    except Exception: gdf = gpd.read_file(path)
    if gdf.empty: return None
    gdf = gdf.set_crs("EPSG:4326") if gdf.crs is None else gdf.to_crs("EPSG:4326")
    gdf = gdf[gdf.geometry.notna() & ~gdf.geometry.is_empty].copy()
    if gdf.empty: return None
    invalid = ~gdf.geometry.is_valid
    if invalid.any():
        try: gdf.loc[invalid, "geometry"] = gdf.loc[invalid, "geometry"].make_valid()
        except Exception: gdf.loc[invalid, "geometry"] = gdf.loc[invalid, "geometry"].buffer(0)
    return gdf

def siapkan_repo(root, source, uploaded_zip=None):
    repo = root / "boundary_repo"
    if repo.exists(): shutil.rmtree(repo)
    if source == "GitHub":
        subprocess.run(["git","clone","--depth","1",GITHUB_REPO_URL,str(repo)], check=True, capture_output=True, text=True)
    else:
        repo.mkdir(parents=True)
        with zipfile.ZipFile(BytesIO(uploaded_zip), "r") as z: z.extractall(repo)
    candidates = [p for p in repo.rglob("processed-data") if p.is_dir()]
    base = candidates[0] if candidates else repo
    prov = base / "02-provinces" / "province-only"; kab = base / "03-districts"
    if not prov.exists() or not kab.exists():
        raise FileNotFoundError("Folder processed-data/02-provinces/province-only atau 03-districts tidak ditemukan.")
    return sorted(prov.rglob("*.geojson")), sorted(kab.rglob("*.geojson"))

def bangun_boundary(files_adm1, files_adm2, tol1, tol2, progress=None):
    d1, fail1 = [], []
    total = len(files_adm1)+len(files_adm2)
    for i,p in enumerate(files_adm1):
        try:
            g = baca_geojson_aman(p)
            if g is None: raise ValueError("Geometri kosong")
            name = p.stem.replace("_"," ").strip(); g["PROVINSI"] = name
            d1.append(g[["PROVINSI","geometry"]])
        except Exception as e: fail1.append({"File":p.name,"Error":str(e)})
        if progress: progress.progress((i+1)/max(total,1))
    if not d1: raise ValueError("Tidak ada boundary provinsi yang berhasil dibaca.")
    prov = gpd.GeoDataFrame(pd.concat(d1,ignore_index=True),geometry="geometry",crs="EPSG:4326")
    prov["KEY_PROV"] = prov["PROVINSI"].map(normalisasi_provinsi)
    prov = prov[["PROVINSI","KEY_PROV","geometry"]].dissolve(by="KEY_PROV",as_index=False,aggfunc="first")
    d2, fail2 = [], []
    for i,p in enumerate(files_adm2):
        try:
            g = baca_geojson_aman(p)
            if g is None: raise ValueError("Geometri kosong")
            g["NAMA_KAB"] = p.stem.replace("_"," ").strip(); d2.append(g[["NAMA_KAB","geometry"]])
        except Exception as e: fail2.append({"File":p.name,"Error":str(e)})
        if progress: progress.progress((len(files_adm1)+i+1)/max(total,1))
    if not d2: raise ValueError("Tidak ada boundary kabupaten/kota yang berhasil dibaca.")
    kab = gpd.GeoDataFrame(pd.concat(d2,ignore_index=True),geometry="geometry",crs="EPSG:4326")
    kab["KEY_KAB"] = kab["NAMA_KAB"].map(normalisasi_kabupaten)
    pts = kab[["geometry"]].copy(); pts["ID"] = pts.index; pts["geometry"] = pts.geometry.representative_point()
    joined = gpd.sjoin(pts, prov[["PROVINSI","KEY_PROV","geometry"]], how="left", predicate="within")
    joined = joined.sort_values("ID").drop_duplicates("ID").set_index("ID")
    kab["PROVINSI"] = kab.index.map(joined["PROVINSI"]); kab["KEY_PROV"] = kab.index.map(joined["KEY_PROV"])
    if kab["KEY_PROV"].isna().any(): raise ValueError(f"Ada {int(kab['KEY_PROV'].isna().sum())} ADM2 belum mendapatkan provinsi.")
    kab["KEY_MAP"] = kab["KEY_PROV"] + "|" + kab["KEY_KAB"]
    duplicates_before = kab[kab["KEY_MAP"].duplicated(False)][["PROVINSI","NAMA_KAB","KEY_MAP"]].copy()
    kab = kab[["PROVINSI","KEY_PROV","NAMA_KAB","KEY_KAB","KEY_MAP","geometry"]].dissolve(
        by="KEY_MAP",as_index=False,aggfunc={"PROVINSI":"first","KEY_PROV":"first","NAMA_KAB":"first","KEY_KAB":"first"})
    def optimize(g,tol):
        x=g.to_crs("EPSG:3857").copy()
        if tol>0: x["geometry"]=x.geometry.simplify(tol,preserve_topology=True)
        x=x.to_crs("EPSG:4326").reset_index(drop=True); x["MAP_ID"]=x.index.astype(str); return x
    prov_plot, kab_plot = optimize(prov,tol1), optimize(kab,tol2)
    return {"prov":prov,"kab":kab,"prov_plot":prov_plot,"kab_plot":kab_plot,
            "fail_prov":pd.DataFrame(fail1),"fail_kab":pd.DataFrame(fail2),"duplicates_before":duplicates_before}

# ============================================================
# 5. TAHAP 3 - STANDARDISASI, MATCHING, PAYLOAD
# ============================================================
def konversi_satu_nilai(v):
    if pd.isna(v): return np.nan
    if isinstance(v,(int,float,np.integer,np.floating)): return float(v)
    s=str(v).strip().replace("%","").replace(" ","")
    if s.casefold() in {"","-","--","na","n/a","none","nan","null"}: return np.nan
    if "," in s and "." in s:
        s = s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    elif "," in s:
        bits=s.split(","); s=s.replace(",",".") if len(bits)==2 and len(bits[1])<=4 else s.replace(",","")
    try: return float(s)
    except Exception: return np.nan

def cari_kolom_indikator(df,nama):
    if nama in df.columns:return nama
    target=re.sub(r"\s+"," ",nama.strip().casefold())
    exact=[c for c in df.columns if re.sub(r"\s+"," ",str(c).strip().casefold())==target]
    if exact:return exact[0]
    starts=[c for c in df.columns if re.sub(r"\s+"," ",str(c).strip().casefold()).startswith(target)]
    return starts[0] if len(starts)==1 else None

def standardisasi_semua(results):
    dp,dk,errors={}, {}, []
    for r in results:
        year=r["tahun"]
        if year is None: errors.append({"File":r["nama_input"],"Error":"Tahun tidak ditemukan pada nama file"}); continue
        if year in dp: errors.append({"File":r["nama_input"],"Error":f"Tahun {year} duplikat"}); continue
        try:
            p=r["sheets"][SHEET_ANALISIS].copy(); k=r["sheets"][SHEET_KABUPATEN_FLAT].copy()
            cp=cari_kolom(p,["Kabupaten / Kota / Provinsi","Provinsi"])
            ckp=cari_kolom(k,["Provinsi","Nama Provinsi"]); ckk=cari_kolom(k,["Kabupaten / Kota","Kabupaten/Kota","Kab/Kota"])
            if not cp or not ckp or not ckk: raise KeyError(f"Kolom wilayah gagal dideteksi: prov={cp}, prov_kab={ckp}, kab={ckk}")
            p["KEY_PROV"]=p[cp].map(normalisasi_provinsi); p=p[(p.KEY_PROV!="")&(p.KEY_PROV!="INDONESIA")].copy()
            k["KEY_PROV"]=k[ckp].map(normalisasi_provinsi); k["KEY_KAB"]=k[ckk].map(normalisasi_kabupaten)
            k=k[(k.KEY_PROV!="")&(k.KEY_KAB!="")&(k.KEY_KAB!="INDONESIA")].copy(); k["KEY_MAP"]=k.KEY_PROV+"|"+k.KEY_KAB
            dp[year],dk[year]=p,k
        except Exception as e: errors.append({"File":r["nama_input"],"Error":str(e)})
    if not dp: raise ValueError("Tidak ada file bertahun valid yang dapat distandardisasi.")
    return {"prov":dp,"kab":dk,"errors":pd.DataFrame(errors)}

def matching_dan_payload(std,boundary):
    bp=set(boundary["prov_plot"].KEY_PROV); bk=set(boundary["kab_plot"].KEY_MAP)
    match=[]
    for year in sorted(std["prov"]):
        for key in sorted(std["prov"][year].KEY_PROV.dropna().unique()): match.append({"Tahun":year,"Level":"Provinsi","Key":key,"Match":key in bp})
        for key in sorted(std["kab"][year].KEY_MAP.dropna().unique()): match.append({"Tahun":year,"Level":"Kabupaten/Kota","Key":key,"Match":key in bk})
    dm=pd.DataFrame(match)
    diagnostics=[]
    for key in dm[(dm.Level=="Kabupaten/Kota")&(~dm.Match)].Key.drop_duplicates():
        if "|" not in key: continue
        p,n=key.split("|",1); candidates=boundary["kab_plot"].loc[boundary["kab_plot"].KEY_PROV==p,"KEY_KAB"].dropna().unique().tolist()
        near=get_close_matches(n,candidates,n=5,cutoff=.25)
        diagnostics.append({"PROVINSI":p,"NAMA_EXCEL":n,**{f"KANDIDAT_{i+1}":near[i] if i<len(near) else "" for i in range(5)}})
    available=[]
    for ind in INDIKATOR_PILIHAN:
        if any(cari_kolom_indikator(std[x][y],ind) for y in std["prov"] for x in ["prov","kab"]): available.append(ind)
    payload,logs={},[]
    logs=[]
    for y in sorted(std["prov"]):
        for level, df, field, order in [
            ("Provinsi",std["prov"][y],"KEY_PROV",boundary["prov_plot"].KEY_PROV.tolist()),
            ("Kabupaten/Kota",std["kab"][y],"KEY_MAP",boundary["kab_plot"].KEY_MAP.tolist())]:
            for ind in available:
                col=cari_kolom_indikator(df,ind)
                if col is None: logs.append({"Tahun":y,"Level":level,"Indikator":ind,"Status":"Tidak tersedia","Kolom sumber":"","Wilayah dengan data":0}); continue
                t=pd.DataFrame({"KEY":df[field],"NILAI":df[col].map(konversi_satu_nilai)}); t=t[t.KEY.notna()&(t.KEY!="")].copy(); t["_NULL"]=t.NILAI.isna().astype(int)
                t=t.sort_values(["KEY","_NULL"]).drop_duplicates("KEY"); lookup=dict(zip(t.KEY,t.NILAI)); vals=[lookup.get(k,np.nan) for k in order]
                valid=[float(v) for v in vals if pd.notna(v)]; typ=tipe_indikator(ind)
                zmax=100.0 if typ=="Persentase" else (float(np.nanpercentile(valid,95)) if valid else 1.0); zmax=max(zmax,1.0)
                payload[f"{y}|||{level}|||{ind}"]={"values":[None if pd.isna(v) else float(v) for v in vals],"zmin":0.0,"zmax":zmax,"type":typ,"source_column":col}
                logs.append({"Tahun":y,"Level":level,"Indikator":ind,"Status":"Tersedia","Kolom sumber":col,"Wilayah dengan data":len(valid)})
    return dm,pd.DataFrame(diagnostics),payload,pd.DataFrame(logs),available

def tipe_indikator(n):
    x=n.casefold()
    if any(k in x for k in ["investigasi adekuat","spesimen adekuat","timeliness","laboratory confirmation","pos rate","persen","percentage"]): return "Persentase"
    if "incidence rate" in x or "discarded rate" in x:return "Rate"
    return "Jumlah"

# ============================================================
# 6. UI DAN ROUTING
# ============================================================
st.sidebar.title("Surveilans MR")
page=st.sidebar.radio("Tahapan",["1. Upload & Flatten","2. Boundary","3. Matching & Indikator","4. Dashboard Peta","5. Ekspor"])
st.sidebar.caption("Tahap harus dijalankan berurutan.")
for label,key in [("Tahap 1","flatten_results"),("Tahap 2","boundary"),("Tahap 3","payload")]:
    done = bool(st.session_state[key]) if key=="flatten_results" else st.session_state[key] is not None
    (st.sidebar.success if done else st.sidebar.info)(f"{label}: {'siap' if done else 'belum'}")
if st.sidebar.button("Reset seluruh proses",use_container_width=True):
    for k,v in DEFAULTS.items(): st.session_state[k]=v
    st.rerun()

st.title("Dashboard Surveilans Campak-Rubella")

if page=="1. Upload & Flatten":
    st.header("Tahap 1 — Upload dan Flatten Header")
    st.info("Upload workbook asli. Header Excel baris 9–10 akan digabung dan data dibaca mulai baris 11. Tahun harus ada pada nama file.")
    files=st.file_uploader("Upload Excel",type=["xlsx","xlsm"],accept_multiple_files=True)
    if files:
        st.dataframe(pd.DataFrame([{"Nama":f.name,"Ukuran MB":round(f.size/1024**2,2),"Tahun":deteksi_tahun(f.name)} for f in files]),hide_index=True,use_container_width=True)
    if st.button("Proses flatten",type="primary",use_container_width=True,disabled=not files):
        ok,err=[],[]; bar=st.progress(0)
        for i,f in enumerate(files):
            try:
                if f.size>MAX_FILE_SIZE_MB*1024**2: raise ValueError(f"File lebih dari {MAX_FILE_SIZE_MB} MB")
                ok.append(proses_excel(f.getvalue(),f.name))
            except Exception as e: err.append({"File":f.name,"Error":str(e)})
            bar.progress((i+1)/len(files))
        st.session_state.flatten_results=ok; st.session_state.flatten_errors=err
        st.session_state.boundary=st.session_state.standardized=st.session_state.matching=st.session_state.payload=None
        st.rerun()
    if st.session_state.flatten_results:
        logs=pd.DataFrame([x for r in st.session_state.flatten_results for x in r["log"]]); st.dataframe(logs,hide_index=True,use_container_width=True)
        for i,r in enumerate(st.session_state.flatten_results): st.download_button(f"Download {r['nama_output']}",r["data"],r["nama_output"],key=f"flat{i}")
    if st.session_state.flatten_errors: st.error("Ada file gagal"); st.dataframe(pd.DataFrame(st.session_state.flatten_errors),hide_index=True)

elif page=="2. Boundary":
    st.header("Tahap 2 — Persiapan Boundary")
    if not st.session_state.flatten_results: st.warning("Jalankan Tahap 1 dahulu."); st.stop()
    source=st.radio("Sumber boundary",["GitHub","Upload ZIP repository"],horizontal=True)
    z=st.file_uploader("Upload ZIP repository",type="zip") if source!="GitHub" else None
    c1,c2=st.columns(2); tol1=c1.number_input("Simplifikasi ADM1 (meter)",0,10000,TOLERANCE_ADM1_DEFAULT,100); tol2=c2.number_input("Simplifikasi ADM2 (meter)",0,10000,TOLERANCE_ADM2_DEFAULT,100)
    if st.button("Proses boundary",type="primary",use_container_width=True,disabled=(source!="GitHub" and z is None)):
        try:
            with tempfile.TemporaryDirectory() as td:
                files1,files2=siapkan_repo(Path(td),source,z.getvalue() if z else None)
                prog=st.progress(0); b=bangun_boundary(files1,files2,tol1,tol2,prog)
            st.session_state.boundary=b; st.session_state.standardized=st.session_state.matching=st.session_state.payload=None; st.rerun()
        except Exception as e: st.exception(e)
    if st.session_state.boundary:
        b=st.session_state.boundary; a,bm,c,d=st.columns(4); a.metric("ADM1",len(b["prov_plot"])); bm.metric("ADM2",len(b["kab_plot"])); c.metric("Gagal ADM1",len(b["fail_prov"])); d.metric("Gagal ADM2",len(b["fail_kab"]))
        if len(b["prov_plot"])!=38: st.warning(f"Boundary provinsi ditemukan {len(b['prov_plot'])}, bukan 38.")
        if not b["duplicates_before"].empty:
            st.caption("Feature ADM2 duplikat telah digabung dengan dissolve."); st.dataframe(b["duplicates_before"],hide_index=True,use_container_width=True)

elif page=="3. Matching & Indikator":
    st.header("Tahap 3 — Matching, Diagnostik, dan Indikator")
    if not st.session_state.boundary: st.warning("Jalankan Tahap 2 dahulu."); st.stop()
    if st.button("Jalankan matching dan bentuk payload",type="primary",use_container_width=True):
        try:
            std=standardisasi_semua(st.session_state.flatten_results)
            m,d,p,l,a=matching_dan_payload(std,st.session_state.boundary)
            st.session_state.standardized=std; st.session_state.matching={"matching":m,"diagnostics":d,"available":a}; st.session_state.payload=p; st.session_state.indicator_log=l; st.session_state.selected_years=sorted(std["prov"]); st.rerun()
        except Exception as e: st.exception(e)
    if st.session_state.matching:
        m=st.session_state.matching["matching"]; summary=m.groupby(["Tahun","Level","Match"]).size().reset_index(name="Jumlah")
        st.dataframe(summary,hide_index=True,use_container_width=True)
        missing=m[~m.Match]
        if missing.empty: st.success("Matching wilayah 100%.")
        else:
            st.warning(f"Masih ada {len(missing)} kombinasi wilayah-tahun belum match."); st.dataframe(missing,hide_index=True,use_container_width=True)
            if not st.session_state.matching["diagnostics"].empty: st.subheader("Kandidat nama terdekat"); st.dataframe(st.session_state.matching["diagnostics"],hide_index=True,use_container_width=True)
        st.subheader("Ketersediaan indikator"); st.dataframe(st.session_state.indicator_log,hide_index=True,use_container_width=True)

elif page=="4. Dashboard Peta":
    st.header("Tahap 4 — Dashboard Peta")
    if not st.session_state.payload: st.warning("Jalankan Tahap 3 dahulu."); st.stop()
    years=st.session_state.selected_years; c1,c2,c3=st.columns([1,1,3]); year=c1.selectbox("Tahun",years); level=c2.selectbox("Level",["Provinsi","Kabupaten/Kota"])
    inds=sorted({k.split("|||",2)[2] for k in st.session_state.payload if k.startswith(f"{year}|||{level}|||")})
    ind=c3.selectbox("Indikator",inds)
    item=st.session_state.payload.get(f"{year}|||{level}|||{ind}")
    b=st.session_state.boundary; g=b["prov_plot"].copy() if level=="Provinsi" else b["kab_plot"].copy()
    g["NILAI"]=item["values"]
    names=g["PROVINSI"] if level=="Provinsi" else g["NAMA_KAB"].astype(str)+", "+g["PROVINSI"].astype(str)
    colors=[[0,WARNA_RENDAH],[.5,WARNA_TENGAH],[1,WARNA_TINGGI]]
    fig=go.Figure()
    nodata=g.NILAI.isna()
    if nodata.any(): fig.add_choropleth(geojson=json.loads(g.to_json()),locations=g.loc[nodata,"MAP_ID"],z=[0]*nodata.sum(),featureidkey="properties.MAP_ID",colorscale=[[0,WARNA_NODATA],[1,WARNA_NODATA]],showscale=False,text=names[nodata],hovertemplate="<b>%{text}</b><br>Data tidak tersedia<extra></extra>",marker_line_color="white",marker_line_width=.35)
    if (~nodata).any(): fig.add_choropleth(geojson=json.loads(g.to_json()),locations=g.loc[~nodata,"MAP_ID"],z=g.loc[~nodata,"NILAI"],featureidkey="properties.MAP_ID",zmin=item["zmin"],zmax=item["zmax"],colorscale=colors,text=names[~nodata],hovertemplate=f"<b>%{{text}}</b><br>{ind}: %{{z}}<extra></extra>",marker_line_color="white",marker_line_width=.35,colorbar_title=item["type"])
    fig.update_geos(fitbounds="locations",visible=False); fig.update_layout(height=720,margin=dict(l=0,r=0,t=60,b=0),title=f"{ind}<br><sup>{level} | {year}</sup>")
    valid=g.NILAI.dropna(); a,bm,c,d=st.columns(4); a.metric("Total wilayah",len(g)); bm.metric("Dengan data",len(valid)); c.metric("Tanpa data",int(nodata.sum())); d.metric("Nilai maksimum",f"{valid.max():,.2f}" if len(valid) else "-")
    st.plotly_chart(fig,use_container_width=True,config={"displaylogo":False,"scrollZoom":True,"toImageButtonOptions":{"format":"png","scale":2}})
    st.caption(f"Kolom sumber: {item['source_column']}")

else:
    st.header("Tahap 5 — Ekspor Hasil")
    if not st.session_state.payload: st.warning("Jalankan Tahap 3 dahulu."); st.stop()
    m=st.session_state.matching; b=st.session_state.boundary
    report=dfs_to_excel({"Matching Wilayah":m["matching"],"Diagnostik":m["diagnostics"],"Indikator":st.session_state.indicator_log,"Boundary Provinsi":pd.DataFrame(b["prov_plot"].drop(columns="geometry")),"Boundary Kab Kota":pd.DataFrame(b["kab_plot"].drop(columns="geometry"))})
    st.download_button("Download laporan Excel",report,"Laporan_Matching_dan_Indikator.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    zipout=BytesIO()
    with zipfile.ZipFile(zipout,"w",zipfile.ZIP_DEFLATED,allowZip64=True) as z:
        z.writestr("Laporan_Matching_dan_Indikator.xlsx",report)
        z.writestr("Boundary_Provinsi.geojson",b["prov_plot"].to_json())
        z.writestr("Boundary_Kabupaten_Kota.geojson",b["kab_plot"].to_json())
        z.writestr("Payload_Dashboard.json",json.dumps(st.session_state.payload,ensure_ascii=False,separators=(",",":")))
        for r in st.session_state.flatten_results: z.writestr(r["nama_output"],r["data"])
    st.download_button("Download seluruh hasil (ZIP)",zipout.getvalue(),"Hasil_Streamlit_Surveilans_MR.zip",mime="application/zip",type="primary",use_container_width=True)
