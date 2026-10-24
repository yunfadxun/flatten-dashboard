# ============================================================
# FLATTEN HEADER MULTIPLE XLSM
# ============================================================

!pip install -q pandas openpyxl

import re
import zipfile
from pathlib import Path
from google.colab import files
import pandas as pd

# ============================================================
# 1. UPLOAD MULTIPLE FILES
# ============================================================

print("Upload semua file XLSM yang ingin diproses.")

uploaded = files.upload()

if not uploaded:
  raise ValueError("Tidak ada file yang di-upload.")

file_input_list = [
    Path(nama_file)
    for nama_file in uploaded.keys()
    if nama_file.lower().endswith(".xlsm")
]

if not file_input_list:
  raise ValueError("Tidak ditemukan file XLSM.")

print(f"\nTotal file: {len(file_input_list)}")

for file_input in file_input_list:
  print(f"- {file_input.name}")


# ============================================================
# 2. KONFIGURASI
# ============================================================

SHEET_ANALISIS = "Tabel Analisis"
SHEET_KAKO = "Tabel Analisis Pusat by Kab"

HEADER_START = 8  # Baris ke-9 di Excel (index 0-based)
HEADER_END = 9  # Baris ke-10 di Excel
DATA_START = 10  # Baris ke-11 di Excel (Awal Data)

FOLDER_OUTPUT = Path("hasil_flatten_header")
FOLDER_OUTPUT.mkdir(parents=True, exist_ok=True)

FILE_ZIP = Path("Hasil_Flatten_Header.zip")


# ============================================================
# 3. FUNGSI: FLATTEN MULTI-ROW HEADER
# ============================================================


def flatten_excel_header(
    excel_file, sheet_name, header_start, header_end, data_start=None
):
  # Baca Excel tanpa header
  raw = pd.read_excel(
      excel_file, sheet_name=sheet_name, header=None, engine="openpyxl"
  )

  # Ambil baris header
  header_rows = raw.iloc[header_start : header_end + 1].copy()

  # Normalisasi header
  header_rows = header_rows.map(
      lambda x: (
          str(x).strip() if pd.notna(x) and str(x).strip() != "" else None
      )
  )

  # Isi header kelompok secara horizontal (merged cells)
  header_rows = header_rows.ffill(axis=1)

  # Gabungkan header setiap kolom
  new_columns = []

  for col_idx in range(raw.shape[1]):
    parts = []

    for row_idx in range(len(header_rows)):
      value = header_rows.iloc[row_idx, col_idx]

      if value is not None:
        value = str(value).strip()

        if value and (not parts or value != parts[-1]):
          parts.append(value)

    if parts:
      col_name = "_".join(parts)
    else:
      col_name = f"Kolom_{col_idx + 1}"

    new_columns.append(col_name)

  # Pastikan nama kolom unik
  seen = {}
  unique_columns = []

  for col in new_columns:
    if col not in seen:
      seen[col] = 0
      unique_columns.append(col)
    else:
      seen[col] += 1
      unique_columns.append(f"{col}_{seen[col]}")

  # Tentukan baris pertama data
  if data_start is None:
    data_start = header_end + 1

  # Buat DataFrame
  df = raw.iloc[data_start:].copy()
  df.columns = unique_columns

  # Hapus baris dan kolom yang seratus persen kosong
  df = df.dropna(how="all")
  df = df.dropna(axis=1, how="all")

  return df.reset_index(drop=True)


# ============================================================
# 4. FUNGSI: CARI SHEET
# ============================================================


def normalisasi_nama_sheet(nama):
  return re.sub(r"[^a-z0-9]", "", str(nama).lower())


def cari_sheet(excel_file, target_sheet):
  sheet_names = pd.ExcelFile(excel_file, engine="openpyxl").sheet_names

  target_normal = normalisasi_nama_sheet(target_sheet)

  for nama in sheet_names:
    if normalisasi_nama_sheet(nama) == target_normal:
      return nama

  for nama in sheet_names:
    if target_normal in normalisasi_nama_sheet(nama):
      return nama

  return None


# ============================================================
# 5. PROSES SELURUH FILE
# ============================================================

hasil_berhasil = []
hasil_gagal = []

for nomor, excel_file in enumerate(file_input_list, start=1):
  print("\n" + "=" * 65)
  print(f"FILE {nomor}/{len(file_input_list)}")
  print(f"Nama: {excel_file.name}")
  print("=" * 65)

  try:
    hasil_sheet = {}

    # --- TABEL ANALISIS ---
    sheet_analisis = cari_sheet(excel_file, SHEET_ANALISIS)

    if sheet_analisis:
      df_analisis = flatten_excel_header(
          excel_file=excel_file,
          sheet_name=sheet_analisis,
          header_start=HEADER_START,
          header_end=HEADER_END,
          data_start=DATA_START,
      )
      hasil_sheet["Tabel Analisis"] = df_analisis
      print(f"[BERHASIL] Tabel Analisis: {df_analisis.shape}")
    else:
      print("[TIDAK DITEMUKAN] Tabel Analisis")

    # --- TABEL ANALISIS PUSAT BY KAB ---
    sheet_kako = cari_sheet(excel_file, SHEET_KAKO)

    if sheet_kako:
      df_kako = flatten_excel_header(
          excel_file=excel_file,
          sheet_name=sheet_kako,
          header_start=HEADER_START,
          header_end=HEADER_END,
          data_start=DATA_START,
      )
      hasil_sheet["Analisis Pusat by Kab"] = df_kako
      print(f"[BERHASIL] Tabel Analisis Pusat by Kab: {df_kako.shape}")
    else:
      print("[TIDAK DITEMUKAN] Tabel Analisis Pusat by Kab")

    if not hasil_sheet:
      raise ValueError("Kedua sheet target tidak ditemukan.")

    # --- SIMPAN HASIL KE EXCEL ---
    nama_output = f"{excel_file.stem}_Flatten_Header.xlsx"
    file_output = FOLDER_OUTPUT / nama_output

    with pd.ExcelWriter(file_output, engine="openpyxl") as writer:
      for nama_sheet, df in hasil_sheet.items():
        df.to_excel(writer, sheet_name=nama_sheet[:31], index=False)

        ws = writer.sheets[nama_sheet[:31]]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col_cells in ws.columns:
          col_letter = col_cells[0].column_letter
          max_length = max(
              (
                  len(str(cell.value))
                  for cell in col_cells
                  if cell.value is not None
              ),
              default=0,
          )
          ws.column_dimensions[col_letter].width = min(max_length + 2, 45)

    hasil_berhasil.append(str(file_output))
    print(f"\n[OUTPUT] {file_output}")

  except Exception as e:
    hasil_gagal.append({"file": excel_file.name, "error": str(e)})
    print(f"\n[GAGAL] {excel_file.name}")
    print(f"Alasan: {e}")


# ============================================================
# 6. KEMAS SELURUH OUTPUT KE ZIP
# ============================================================

if hasil_berhasil:
  with zipfile.ZipFile(FILE_ZIP, "w", zipfile.ZIP_DEFLATED) as zipf:
    for file_output in hasil_berhasil:
      file_output = Path(file_output)
      zipf.write(file_output, arcname=file_output.name)


# ============================================================
# 7. RINGKASAN & DOWNLOAD
# ============================================================

print("\n" + "=" * 65)
print("RINGKASAN PROSES")
print("=" * 65)

print(f"Total file input : {len(file_input_list)}")
print(f"Berhasil         : {len(hasil_berhasil)}")
print(f"Gagal            : {len(hasil_gagal)}")

if hasil_gagal:
  print("\nFile yang gagal:")
  for item in hasil_gagal:
    print(f"\n- {item['file']}")
    print(f"  Error: {item['error']}")

if hasil_berhasil:
  print("\nFile output:")
  for file_output in hasil_berhasil:
    print(f"- {Path(file_output).name}")

  print("\nMengunduh hasil dalam ZIP...")
  files.download(str(FILE_ZIP))
else:
  print("\nTidak ada file yang berhasil diproses.")
