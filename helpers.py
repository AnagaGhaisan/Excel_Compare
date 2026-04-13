import os
import re
import math
import openpyxl
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

# File paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DRAFT_TEMPLATE_PATH = os.path.join(BASE_DIR, "static/template/Draft Output.xlsx")

# Helper function to check file extensions
def allowed_file(filename, allowed_extensions):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed_extensions


def extract_no_faktur_from_description(desc: str, voucher_cat: str) -> str | None:
    if pd.isna(desc):
        return None

    s = str(desc).strip()
    parts = [p.strip() for p in s.split("/")]

    # 1. CARI VOUCHER MENGGUNAKAN REGEX PINTAR
    voucher_part = None
    for p in parts:
        # (?i) = Kebal terhadap huruf besar/kecil
        if re.search(r'(?i)(?:[A-Z]{2,5}-\d+[A-Z0-9_-]*|[A-Z]{2,5}\d+-[A-Z0-9_-]+|[A-Z]{2,5}\d{5,})', p):
            voucher_part = p
            break

    # 2. LOGIKA FALLBACK (Jika formatnya sangat aneh sehingga regex gagal)
    if not voucher_part:
        if str(voucher_cat).strip() == "GL-JV":
            voucher_part = parts if len(parts) >= 1 else None
        else:
            voucher_part = parts if len(parts) >= 2 else (parts if parts else None)
            
    # 3. KEMBALIKAN JIKA KATEGORI GL-JV
    if str(voucher_cat).strip() == "GL-JV":
        return voucher_part

# --- 4. LOGIKA KHUSUS DIGUNGGUNG: Deteksi Nominal di K3 ---
    if len(parts) >= 3 and voucher_part:
        # HARUS pakai parts karena kita mau membersihkan nominal uangnya, BUKAN vouchernya
        clean_num = parts[2].replace(" ", "").replace(",", "").replace(".", "")

        # PERBAIKAN PAMUNGKAS: 
        # 1. Tidak boleh diawali nol (Mencegah Nomor HP)
        # 2. Panjang angka minimal 4 digit (Mencegah ID Toko/Area seperti "318" ikut tergabung)
        if clean_num.isdigit() and not clean_num.startswith("0") and len(clean_num) > 4:  # <--- UBAH JADI > 4
            try:
                v = float(clean_num)
                formatted_num = f"{int(v)}" if v.is_integer() else f"{v}"
                return f"{voucher_part}/ {formatted_num}"
            except ValueError:
                try:
                    v = _parse_id_number(parts)
                    if pd.notna(v):
                        formatted_num = f"{int(v)}" if float(v).is_integer() else f"{v}"
                        return f"{voucher_part}/ {formatted_num}"
                except:
                    pass

    return voucher_part

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalisasi nama kolom: uppercase, spasi/punctuation -> underscore, rapihin underscore.
    Contoh: 'No Voucher' -> 'NO_VOUCHER'
    """

    def clean(col):
        col = str(col).strip().upper()
        col = re.sub(r"[^A-Z0-9]+", "_", col)
        col = re.sub(r"_+", "_", col).strip("_")
        return col

    df = df.copy()
    df.columns = [clean(c) for c in df.columns]
    return df


def _parse_id_number(x):
    """
    Aman untuk angka dengan format Indonesia:
    - '77.597.727' -> 77597727
    - '7.759.773'  -> 7759773
    - '1.234,56'   -> 1234.56
    """
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)

    s = str(x).strip()
    if not s:
        return np.nan

    s = s.replace(" ", "")
    # hapus ribuan ".", ubah desimal "," jadi "."
    s = s.replace(".", "")
    s = s.replace(",", ".")
    # handle (123) -> -123
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]

    try:
        v = float(s)
        return -v if neg else v
    except:
        return np.nan


def calculate_net(row):
    debit = float(row.get("Debit Amount", 0))
    credit = float(row.get("Credit Amount", 0))
    acc_name = str(row.get("Account Name", "")).strip()

    # Pendapatan: -Debit + Credit
    if acc_name in [
        "Interest Bank Income",
        "Other Income",
        "Rental Income",
        "Repair Service Income",
        "Sales",
        "Sales Price Protection",
    ]:
        return -debit + credit

    # Beban: Debit - Credit
    elif acc_name in ["POP Expense", "Promotion Gift"]:
        return debit - credit

    # Sales Return: -Debit - Credit (Sesuai -AA10-AB10)
    elif acc_name == "Sales Return":
        return -(debit - credit)

    return 0


def compare_files(
    k3_sheets: dict, coretax_sheets_1: dict, coretax_sheets_2: dict, output_dir: str
) -> str:
    # k3_sheets, coretax_sheets_1, coretax_sheets_2 are already dicts of {sheet_name: DataFrame}
    # from pd.read_excel(..., sheet_name=None) in app.py — no need to re-read.

    # 1) Concatenate all K3 sheets (keep original column names for later use)
    k3 = pd.concat(k3_sheets.values(), ignore_index=True)
    print(f"K3 combined shape: {k3.shape}, columns: {list(k3.columns)}")

    # 2) BARU terapkan ekstraksi No. Faktur & Nett pada variabel 'k3'
    k3["No Faktur (key)"] = k3.apply(
        lambda row: extract_no_faktur_from_description(
            row.get("Description", ""), row.get("Voucher Category", "")
        ),
        axis=1,
    )
    # Bersihkan spasi agar tidak meleset saat merge
    k3["No Faktur (key)"] = k3["No Faktur (key)"].astype(str).str.strip()
    k3["Nett"] = k3.apply(calculate_net, axis=1)

    # 2) Normalize columns for Coretax (biar NO VOUCHER / DOC_NO kebaca konsisten)
    coretax_1 = pd.concat(
        [_normalize_columns(sheet_data) for sheet_data in coretax_sheets_1.values()],
        ignore_index=True,
    )
    coretax_2 = pd.concat(
        [_normalize_columns(sheet_data) for sheet_data in coretax_sheets_2.values()],
        ignore_index=True,
    )

    # Pastikan key jadi NO_VOUCHER
    if "DOC_NO" in coretax_1.columns and "NO_VOUCHER" not in coretax_1.columns:
        coretax_1 = coretax_1.rename(columns={"DOC_NO": "NO_VOUCHER"})
    if "DOC_NO" in coretax_2.columns and "NO_VOUCHER" not in coretax_2.columns:
        coretax_2 = coretax_2.rename(columns={"DOC_NO": "NO_VOUCHER"})

    if "NO_VOUCHER" not in coretax_1.columns:
        raise ValueError(
            "Coretax Digunggung: kolom DOC_NO / NO VOUCHER tidak ditemukan."
        )
    if "NO_VOUCHER" not in coretax_2.columns:
        raise ValueError(
            "Coretax Tidak Digunggung: kolom DOC_NO / NO VOUCHER tidak ditemukan."
        )

    # 3) Harmonize DPP/PPN + CUSTOMER + status
    # --- Digunggung: AMOUNT_BEF_TAX = DPP, TAX_AMOUNT = PPN, CUSTOMER_NAME = CUSTOMER
    if "DPP" not in coretax_1.columns and "AMOUNT_BEF_TAX" in coretax_1.columns:
        coretax_1["DPP"] = coretax_1["AMOUNT_BEF_TAX"]
    if "PPN" not in coretax_1.columns and "TAX_AMOUNT" in coretax_1.columns:
        coretax_1["PPN"] = coretax_1["TAX_AMOUNT"]

    coretax_1["CUSTOMER"] = (
        coretax_1["CUSTOMER_NAME"] if "CUSTOMER_NAME" in coretax_1.columns else None
    )
    coretax_1["FP_STATUS"] = "FP Digunggung"

    # --- Tidak Digunggung: DPP = DPP, PPN = PPN, NAMA_PEMBELI = CUSTOMER
    if "DPP" not in coretax_2.columns and "AMOUNT_BEF_TAX" in coretax_2.columns:
        coretax_2["DPP"] = coretax_2["AMOUNT_BEF_TAX"]
    if "PPN" not in coretax_2.columns and "TAX_AMOUNT" in coretax_2.columns:
        coretax_2["PPN"] = coretax_2["TAX_AMOUNT"]

    if "NAMA_PEMBELI" in coretax_2.columns:
        coretax_2["CUSTOMER"] = coretax_2["NAMA_PEMBELI"]
    elif "CUSTOMER_NAME" in coretax_2.columns:
        coretax_2["CUSTOMER"] = coretax_2["CUSTOMER_NAME"]
    else:
        coretax_2["CUSTOMER"] = None

    coretax_2["FP_STATUS"] = "FP Tidak Digunggung"

    if "DEPT" in coretax_1.columns:
        coretax_1["CUSTOMER"] = coretax_1["DEPT"]
    elif "CUSTOMER_NAME" in coretax_1.columns:
        coretax_1["CUSTOMER"] = coretax_1["CUSTOMER_NAME"]
    else:
        coretax_1["CUSTOMER"] = None

    coretax_1["FP_STATUS"] = "FP Digunggung"

    # 4) Bersihin key + convert angka
    for df in (coretax_1, coretax_2):
        df["NO_VOUCHER"] = df["NO_VOUCHER"].astype(str).str.strip()
        if "DPP" in df.columns:
            df["DPP"] = df["DPP"].apply(_parse_id_number)
        else:
            df["DPP"] = 0.0
        if "PPN" in df.columns:
            df["PPN"] = df["PPN"].apply(_parse_id_number)
        else:
            df["PPN"] = 0.0

    # ----- LOGIKA KHUSUS DIGUNGGUNG: GABUNG NO_VOUCHER + AMOUNT -----
    # 1. Cari yang duplikat HANYA di Coretax 1 (Digunggung)
    coretax_1_counts = coretax_1["NO_VOUCHER"].value_counts()
    dup_vouchers_c1 = coretax_1_counts[coretax_1_counts > 1].index.tolist()

    # Fungsi format angka (hapus koma .0 jika integer biar match)
    def format_amount(val):
        if pd.isna(val):
            return "0"
        val_float = abs(float(val))  # Pakai absolute biar minus/plus tetap ketemu
        return f"{int(val_float)}" if val_float.is_integer() else f"{val_float}"

    # 2. Modifikasi NO_VOUCHER di Coretax 1 (KANAN)
    is_dup_c1 = coretax_1["NO_VOUCHER"].isin(dup_vouchers_c1)
    if is_dup_c1.any():
        coretax_1.loc[is_dup_c1, "NO_VOUCHER"] = (
            coretax_1.loc[is_dup_c1, "NO_VOUCHER"]
            + "/ "
            + coretax_1.loc[is_dup_c1, "DPP"].apply(format_amount)
        )

    # 3. Modifikasi No Faktur (key) di GL K3 (KIRI) AGAR BISA MATCH!
    is_dup_k3 = k3["No Faktur (key)"].isin(dup_vouchers_c1)
    if is_dup_k3.any():

        def get_gl_amount(row):
            # Ambil nominal dari GL sesuai rumus Difference kamu
            if str(row.get("Account Name", "")).strip() == "Sales Return":
                return float(row.get("Debit Amount", 0))
            else:
                return float(row.get("Nett", 0))

        # Terapkan format angka yang sama dan gabungkan ke No Faktur GL
        gl_amounts = k3[is_dup_k3].apply(get_gl_amount, axis=1).apply(format_amount)
        k3.loc[is_dup_k3, "No Faktur (key)"] = (
            k3.loc[is_dup_k3, "No Faktur (key)"] + "/ " + gl_amounts
        )
    # ----------------------------------------------------------------

    # 5) Combine Coretax (ambil kolom penting aja)
    keep_cols_1 = [
        c
        for c in ["NO_VOUCHER", "VOUCHER_NO", "DPP", "PPN", "CUSTOMER", "FP_STATUS"]
        if c in coretax_1.columns
    ]
    keep_cols_2 = [
        c
        for c in ["NO_VOUCHER", "VOUCHER_NO", "DPP", "PPN", "CUSTOMER", "FP_STATUS"]
        if c in coretax_2.columns
    ]
    coretax_combined = pd.concat(
        [coretax_1[keep_cols_1], coretax_2[keep_cols_2]], ignore_index=True
    )

    coretax_combined = coretax_combined.drop_duplicates(
        subset=["NO_VOUCHER"], keep="first"
    )

    # 6) kalau NO_VOUCHER muncul beberapa kali, DPP/PPN dijumlah, CUSTOMER diambil first non-null, status digabung unik
    def join_unique(series):
        vals = [v for v in series.dropna().astype(str).tolist() if v.strip()]
        return "; ".join(sorted(set(vals))) if vals else None

    agg_map = {
        "DPP": "sum",
        "PPN": "sum",
        "CUSTOMER": "first",  # Takes first non-null value for CUSTOMER
        "FP_STATUS": join_unique,
    }
    if "VOUCHER_NO" in coretax_combined.columns:
        agg_map["VOUCHER_NO"] = "first"

    # Aggregate the combined coretax data
    coretax_agg = coretax_combined.groupby("NO_VOUCHER", as_index=False).agg(agg_map)

    # 8) Debugging step: Check columns in Coretax_2
    print("Columns in Coretax_2:", coretax_2.columns)

    # 9) Debugging: Check if 'NO FP MODIF' exists in coretax_2
    if "NO FP MODIF" in coretax_2.columns:
        print("NO FP MODIF exists in Coretax_2.")
    else:
        print("NO FP MODIF NOT found in Coretax_2")

    # 10) Merge
    merged = pd.merge(
        k3,
        coretax_agg,
        left_on="No Faktur (key)",
        right_on="NO_VOUCHER",
        how="left",
        indicator=True,
    )

    # 11) Compute Difference based on account type
    merged["Debit Amount"] = pd.to_numeric(
        merged["Debit Amount"], errors="coerce"
    ).fillna(0)
    merged["Credit Amount"] = pd.to_numeric(
        merged["Credit Amount"], errors="coerce"
    ).fillna(0)

    # Apply the Net calculation based on the account type
    merged["Net"] = merged.apply(calculate_net, axis=1)

    merged["DPP"] = pd.to_numeric(merged["DPP"], errors="coerce").fillna(0)
    merged["PPN"] = pd.to_numeric(merged["PPN"], errors="coerce").fillna(0)

    # Logika baru untuk menghitung Difference
    def calculate_difference(row):
        return float(row["Net"]) - float(
            row["DPP"]
        )  # Menggunakan float, tanpa pembulatan

    # Terapkan fungsi ke kolom Difference
    merged["Difference"] = merged.apply(calculate_difference, axis=1)

    # 12) Keterangan + Customer (langsung dari kolom kanonik)
    merged["Keterangan (Digunggung/Tidak Digunngung)"] = merged["FP_STATUS"]
    merged.loc[
        merged["_merge"] != "both", "Keterangan (Digunggung/Tidak Digunngung)"
    ] = "Tidak ada di Coretax"

    merged["Customer"] = merged["CUSTOMER"]
    merged.loc[merged["_merge"] != "both", "Customer"] = None

    # Debugging: Check if 'NO_FP_MODIF' exists in coretax_2
    print(
        "Cek apakah 'NO_FP_MODIF' ada di coretax_2:", "NO_FP_MODIF" in coretax_2.columns
    )
    if "NO_FP_MODIF" in coretax_2.columns:
        print(
            coretax_2["NO_FP_MODIF"].head()
        )  # Menampilkan beberapa nilai untuk memastikan kolom ada
    else:
        print("Kolom 'NO_FP_MODIF' tidak ditemukan di Coretax_2")

    print("Setelah merge, kolom di merged:", merged.columns)

    # Jika kolom 'NO_FP_MODIF' ada di coretax_2, tambahkan ke merged
    if "NO_FP_MODIF" in coretax_2.columns:
        # PENTING: Drop duplicate agar tidak terjadi ledakan data (Cartesian product) saat merge
        coretax_2_unique = coretax_2[["NO_VOUCHER", "NO_FP_MODIF"]].drop_duplicates(
            subset=["NO_VOUCHER"]
        )
        merged = pd.merge(
            merged,
            coretax_2_unique,
            left_on="No Faktur (key)",
            right_on="NO_VOUCHER",
            how="left",
            suffixes=("", "_from_coretax2"),
        )
        print("Setelah merge NO_FP_MODIF, kolom di merged:", merged.columns)
    else:
        merged["NO_FP_MODIF"] = None  # Atur sebagai None jika kolom tidak ada

    # 13) Before filling NaN, convert categorical columns to string type
    for column in merged.columns:
        if (
            merged[column].dtype.name == "category"
        ):  # Check if the column is categorical
            merged[column] = merged[column].astype(str)
    # Ensure column headers are strings
    merged.columns = merged.columns.astype(str)
    # Now, proceed with other operations
    merged = merged.astype(str).fillna("-")

    # Cek jika file Excel ada
    if not os.path.exists(DRAFT_TEMPLATE_PATH):
        raise FileNotFoundError(
            "Draft Output.xlsx tidak ditemukan. Taruh file itu 1 folder dengan app.py"
        )

    # Load template Excel
    wb = load_workbook(DRAFT_TEMPLATE_PATH)
    ws = wb.active

    # --- 14) INISIALISASI TOTAL & VARIABEL (Sebelum Loop) ---
    debit_total = credit_total = net_total = balance_total = 0
    dpp_total = ppn_total = difference_total = 0

    template_ws = ws  # Sheet pertama (template)
    start_row = 5  # Data mulai baris 5
    MAX_ROWS_PER_SHEET = 500_000

    # Helper: copy header rows (1-4) from template to a new sheet
    def _copy_header_rows(src_ws, dst_ws, up_to_row=4):
        for row_idx in range(1, up_to_row + 1):
            for col_idx in range(1, src_ws.max_column + 1):
                src_cell = src_ws.cell(row_idx, col_idx)
                dst_cell = dst_ws.cell(row_idx, col_idx)
                dst_cell.value = src_cell.value
                if src_cell.has_style:
                    dst_cell.font = src_cell.font.copy()
                    dst_cell.border = src_cell.border.copy()
                    dst_cell.fill = src_cell.fill.copy()
                    dst_cell.number_format = src_cell.number_format
                    dst_cell.alignment = src_cell.alignment.copy()

    def _sanitize_sheet_name(name: str) -> str:
        """Bersihkan nama sheet Excel (max 31 karakter, tanpa karakter terlarang)."""
        for ch in ['[', ']', ':', '*', '?', '/', '\\']:
            name = name.replace(ch, '_')
        return name.strip()[:31]

    # --- PERBAIKAN: PRE-CALCULATE DATA UNTUK MENCEGAH LOOPING LAMBAT ---
    print("Menyiapkan dictionary untuk mempercepat proses perhitungan...")

    # 1. Jadikan Set agar pencarian 'in' berjalan sekejap mata
    coretax_voucher_set = set(coretax_agg["NO_VOUCHER"].dropna().astype(str))

    # 2. Hitung total Net dan Debit dari awal, jangan hitung di dalam loop
    merged["Net_Numeric"] = pd.to_numeric(merged["Net"], errors="coerce").fillna(0)
    merged["Debit_Numeric"] = pd.to_numeric(
        merged["Debit Amount"], errors="coerce"
    ).fillna(0)

    grouped_totals = (
        merged.groupby("NO_VOUCHER")
        .agg(
            total_gl_net=("Net_Numeric", "sum"), total_gl_debit=("Debit_Numeric", "sum")
        )
        .to_dict("index")
    )
    # -------------------------------------------------------------------

    # --- 15) TULIS DATA PER AKUN (PISAH SHEET, SPLIT >500K BARIS) ---
    processed_coretax = set()  # Untuk melacak voucher mana yang sudah muncul data Coretax-nya
    sheet_row_counts = {}
    first_sheet = True

    # Group by Account Name, pertahankan urutan kemunculan
    account_groups = merged.groupby("Account Name", sort=False)

    for account_name, group_df in account_groups:
        group_df = group_df.reset_index(drop=True)
        total_rows = len(group_df)
        n_chunks = max(1, math.ceil(total_rows / MAX_ROWS_PER_SHEET))
        safe_name = _sanitize_sheet_name(str(account_name))

        for chunk_idx in range(n_chunks):
            chunk_start = chunk_idx * MAX_ROWS_PER_SHEET
            chunk_end = min(chunk_start + MAX_ROWS_PER_SHEET, total_rows)
            chunk_df = group_df.iloc[chunk_start:chunk_end]

            # Tentukan nama sheet
            if n_chunks == 1:
                sheet_title = safe_name
            else:
                suffix = f"_{chunk_idx + 1}"
                sheet_title = safe_name[: 31 - len(suffix)] + suffix

            # Buat / ambil worksheet
            if first_sheet:
                current_ws = template_ws
                current_ws.title = sheet_title
                first_sheet = False
            else:
                current_ws = wb.create_sheet(title=sheet_title)
                _copy_header_rows(template_ws, current_ws)

            sheet_row_counts[current_ws.title] = 0

            for local_i in range(len(chunk_df)):
                r = start_row + local_i
                row = chunk_df.iloc[local_i]

                v_bal = _parse_id_number(row.get("Balance", 0))

                # --- PERBAIKAN: Bersihkan NaN agar tidak dianggap teks "nan" ---
                voucher_no = str(row.get("NO_VOUCHER", "-")).strip()
                if voucher_no.lower() in ["nan", "none", "", "-"]:
                    voucher_no = "-"

                # --- SISI GL (KIRI): TETAP FLOAT (MENYIMPAN KOMA) ---
                row_debit = float(row.get("Debit Amount", 0))
                row_credit = float(row.get("Credit Amount", 0))
                row_net = float(row.get("Net", 0))

                # Selalu tambahkan data GL ke subtotal
                debit_total += row_debit
                credit_total += row_credit
                net_total += row_net
                balance_total += v_bal

                # Cek apakah voucher valid ini ada di Coretax
                voucher_no_in_coretax = (voucher_no != "-") and (
                    voucher_no in coretax_voucher_set
                )

                # --- LOGIKA PENENTUAN STATUS & DIFFERENCE ---
                row_diff_formula = "" # Variabel untuk menampung rumus Excel dinamis

                if voucher_no == "-":
                    # KONDISI 1: DATA KOSONG
                    row_dpp = 0.0
                    row_ppn = 0.0

                    if row["Account Name"] == "Sales Return":
                        row_diff = -(float(row_debit) + row_dpp)
                        # Rumus Excel: -(Debit + DPP)
                        row_diff_formula = f"=-(G{r} + O{r})" 
                    else:
                        row_diff = float(row_net)
                        # Rumus Excel: Net
                        row_diff_formula = f"=I{r}" 

                    status = "Tidak ada di Coretax"
                    difference_total += row_diff

                elif voucher_no not in processed_coretax:
                    # KONDISI 2: BARIS PERTAMA YANG MATCH
                    row_dpp = float(row.get("DPP", 0))
                    row_ppn = float(row.get("PPN", 0))

                    if voucher_no_in_coretax:
                        totals = grouped_totals.get(voucher_no, {})
                        total_gl_net = float(totals.get("total_gl_net", 0))
                        total_gl_debit = float(totals.get("total_gl_debit", 0))

                        if row["Account Name"] == "Sales Return":
                            row_diff = -(float(total_gl_debit) + row_dpp)
                            # Karena total_gl_debit adalah gabungan banyak baris, angkanya kita print mati, tapi DPP tetap referensi sel O
                            row_diff_formula = f"=-({total_gl_debit} + O{r})"
                        elif row["Account Name"] in ["Repair Service Income", "Sales Price Protection", "Sales"]:
                            row_diff = float(total_gl_net) - row_dpp
                            row_diff_formula = f"={total_gl_net} - O{r}"
                        else:
                            if total_gl_net == 0:
                                row_diff = float(row_net - row_dpp)
                                # Gunakan Net baris ini dikurangi DPP
                                row_diff_formula = f"=I{r} - O{r}"
                            else:
                                row_diff = float(total_gl_net - row_dpp)
                                row_diff_formula = f"={total_gl_net} - O{r}"

                        status = "Unique"
                    else:
                        if row["Account Name"] == "Sales Return":
                            row_diff = -(float(row_debit) + row_dpp)
                            row_diff_formula = f"=-(G{r} + O{r})"
                        elif row["Account Name"] in ["Repair Service Income", "Sales Price Protection", "Sales"]:
                            row_diff = float(row_net) - row_dpp
                            row_diff_formula = f"=I{r} - O{r}"
                        else:
                            row_diff = float(row_net)
                            row_diff_formula = f"=I{r}"
                        status = "Tidak ada di Coretax"

                    dpp_total += row_dpp
                    ppn_total += row_ppn
                    difference_total += row_diff
                    processed_coretax.add(voucher_no)

                else:
                    # KONDISI 3: BARIS LANJUTAN
                    row_dpp = 0.0
                    row_ppn = 0.0
                    row_diff = 0
                    row_diff_formula = "=0"

                    if voucher_no_in_coretax:
                        status = "Unique"
                    else:
                        status = "Tidak ada di Coretax"

                # --- PROSES CETAK KE SHEET ---
                current_ws.cell(r, 1).value = row.get("Account No.")
                current_ws.cell(r, 2).value = row.get("Account Name")
                current_ws.cell(r, 3).value = row.get("Date")
                current_ws.cell(r, 4).value = row.get("Voucher Category")
                current_ws.cell(r, 5).value = row.get("Voucher No.")
                current_ws.cell(r, 6).value = row.get("Description")
                current_ws.cell(r, 7).value = row_debit
                current_ws.cell(r, 8).value = row_credit
                current_ws.cell(r, 9).value = row_net
                current_ws.cell(r, 10).value = row.get("Direction")
                current_ws.cell(r, 11).value = v_bal

                current_ws.cell(r, 13).value = voucher_no if voucher_no != "-" else None
                current_ws.cell(r, 14).value = (
                    row.get("NO_FP_MODIF")
                    if str(row.get("NO_FP_MODIF")) not in ["nan", "None"]
                    else None
                )
                current_ws.cell(r, 15).value = row_dpp
                current_ws.cell(r, 16).value = row_ppn
                
                # Cetak formula yang sudah dipilih secara cerdas oleh Python
                current_ws.cell(r, 17).value = row_diff_formula
                current_ws.cell(r, 18).value = (
                    row.get("Customer")
                    if str(row.get("Customer")) not in ["nan", "None"]
                    else None
                )
                current_ws.cell(r, 19).value = row.get(
                    "Keterangan (Digunggung/Tidak Digunngung)"
                )
                current_ws.cell(r, 20).value = status

                sheet_row_counts[current_ws.title] += 1

        print(f"Akun '{account_name}': {total_rows} baris -> {n_chunks} sheet")

    # --- 16) CETAK HASIL SUBTOTAL DINAMIS (Mengikuti Filter) ---
    # Mapping kolom: 7=G, 8=H, 9=I, 11=K, 15=O, 16=P, 17=Q
    # SUBTOTAL(109, ...) menjumlahkan baris yang terlihat (mengikuti filter & hidden rows).
    col_letter_map = {7: "G", 8: "H", 9: "I", 11: "K", 15: "O", 16: "P", 17: "Q"}

    bold_font = Font(bold=True)

    # Terapkan formula subtotal ke setiap sheet, mengikuti jumlah data pada sheet tersebut
    for ws_name, row_count in sheet_row_counts.items():
        target_ws = wb[ws_name]
        if row_count <= 0:
            # Tidak ada data pada sheet ini
            for col_idx in col_letter_map:
                cell = target_ws.cell(3, col_idx)
                cell.value = 0
                cell.font = bold_font
            continue

        last_row = start_row + row_count - 1
        for col_idx, col_letter in col_letter_map.items():
            cell = target_ws.cell(3, col_idx)
            cell.value = (
                f"=SUBTOTAL(109, {col_letter}{start_row}:{col_letter}{last_row})"
            )
            cell.font = bold_font

    # --- 17) FORMAT RANGE JADI EXCEL TABLE DI TIAP SHEET ---
    header_row = 4
    first_col = "A"
    last_col = "T"
    table_style = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    table_idx = 1
    for ws_name, row_count in sheet_row_counts.items():
        if row_count <= 0:
            continue

        target_ws = wb[ws_name]
        last_data_row = start_row + row_count - 1
        table_ref = f"{first_col}{header_row}:{last_col}{last_data_row}"
        table_name = f"DataTable{table_idx}"
        table_idx += 1

        tab = Table(displayName=table_name, ref=table_ref)
        tab.tableStyleInfo = table_style
        target_ws.add_table(tab)

    print(f"Selesai! Data ditulis ke {len(sheet_row_counts)} sheet. Total baris: {len(merged)}")

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Save the workbook (preserves template formatting + all sheets)
    out_name = f"Draft_Updated_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_path = os.path.join(output_dir, out_name)
    wb.save(out_path)
    print(f"Saved output to {out_path}")

    # Dapatkan daftar nama sheet dari workbook yang baru disimpan
    sheet_names = wb.sheetnames

    return out_path, out_name, sheet_names


def delete_all_uploaded_files(app):
    try:
        # Cek jika direktori upload ada
        if os.path.exists(app.config["UPLOAD_FOLDER"]):
            # Hapus semua file dalam folder upload
            for filename in os.listdir(app.config["UPLOAD_FOLDER"]):
                file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)  # Hapus file
                    print(f"Uploaded file deleted: {file_path}")
    except Exception as e:
        print(f"Error deleting uploaded files: {e}")
