import os
import re
import math
import uuid
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

FORMULA_CREDIT_MINUS_DEBIT = "credit_minus_debit"
FORMULA_DEBIT_MINUS_CREDIT = "debit_minus_credit"
FORMULA_NEGATIVE_DEBIT_MINUS_CREDIT = "negative_debit_minus_credit"

FORMULA_ALIASES = {
    FORMULA_CREDIT_MINUS_DEBIT: FORMULA_CREDIT_MINUS_DEBIT,
    FORMULA_DEBIT_MINUS_CREDIT: FORMULA_DEBIT_MINUS_CREDIT,
    FORMULA_NEGATIVE_DEBIT_MINUS_CREDIT: FORMULA_NEGATIVE_DEBIT_MINUS_CREDIT,
    "-debit + credit": FORMULA_CREDIT_MINUS_DEBIT,
    "debit - credit": FORMULA_DEBIT_MINUS_CREDIT,
    "-(debit - credit)": FORMULA_NEGATIVE_DEBIT_MINUS_CREDIT,
}

ACCOUNT_FORMULA_DEFAULTS = {
    "Interest Bank Income": FORMULA_CREDIT_MINUS_DEBIT,
    "Other Income": FORMULA_CREDIT_MINUS_DEBIT,
    "Rental Income": FORMULA_CREDIT_MINUS_DEBIT,
    "Repair Service Income": FORMULA_CREDIT_MINUS_DEBIT,
    "Sales": FORMULA_CREDIT_MINUS_DEBIT,
    "Sales Price Protection": FORMULA_CREDIT_MINUS_DEBIT,
    "POP Expense": FORMULA_DEBIT_MINUS_CREDIT,
    "Promotion Gift": FORMULA_DEBIT_MINUS_CREDIT,
    "Sales Return": FORMULA_NEGATIVE_DEBIT_MINUS_CREDIT,
}

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
        if re.search(
            r"(?i)(?:[A-Z]{2,5}-\d+[A-Z0-9_-]*|[A-Z]{2,5}\d+-[A-Z0-9_-]+|[A-Z]{2,5}\d{5,})",
            p,
        ):
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
        if (
            clean_num.isdigit() and not clean_num.startswith("0") and len(clean_num) > 4
        ):  # <--- UBAH JADI > 4
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


def _pick_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    return None


def _ensure_column_from_aliases(
    df: pd.DataFrame, target: str, candidates: list[str], required: bool = False
) -> pd.DataFrame:
    df = df.copy()
    source = _pick_existing_column(df, candidates)

    if source is None:
        if required:
            raise ValueError(
                f"Kolom untuk '{target}' tidak ditemukan. Coba salah satu: {', '.join(candidates)}"
            )
        if target not in df.columns:
            df[target] = None
        return df

    if source != target:
        df[target] = df[source]

    return df


def _prepare_k3_dataframe(k3_sheets: dict) -> pd.DataFrame:
    k3 = pd.concat(k3_sheets.values(), ignore_index=True)
    k3 = _normalize_columns(k3)
    k3 = _ensure_column_from_aliases(
        k3, "ACCOUNT_NO", ["ACCOUNT_NO", "ACCOUNT_NO_"], required=True
    )
    k3 = _ensure_column_from_aliases(
        k3, "ACCOUNT_NAME", ["ACCOUNT_NAME"], required=True
    )
    k3 = _ensure_column_from_aliases(k3, "DATE", ["DATE"], required=True)
    k3 = _ensure_column_from_aliases(
        k3, "VOUCHER_CATEGORY", ["VOUCHER_CATEGORY"], required=True
    )
    k3 = _ensure_column_from_aliases(k3, "VOUCHER_NO", ["VOUCHER_NO"], required=True)
    k3 = _ensure_column_from_aliases(k3, "DESCRIPTION", ["DESCRIPTION"], required=True)
    k3 = _ensure_column_from_aliases(
        k3, "DEBIT_AMOUNT", ["DEBIT_AMOUNT"], required=True
    )
    k3 = _ensure_column_from_aliases(
        k3, "CREDIT_AMOUNT", ["CREDIT_AMOUNT"], required=True
    )
    k3 = _ensure_column_from_aliases(k3, "DIRECTION", ["DIRECTION"], required=False)
    k3 = _ensure_column_from_aliases(k3, "BALANCE", ["BALANCE"], required=False)
    return k3


def normalize_formula_key(formula_key: str | None) -> str | None:
    if formula_key is None:
        return None

    normalized = FORMULA_ALIASES.get(str(formula_key).strip())
    return normalized


def _get_default_formula_for_account(account_name: str, direction: str = "") -> str:
    if account_name in ACCOUNT_FORMULA_DEFAULTS:
        return ACCOUNT_FORMULA_DEFAULTS[account_name]

    direction_normalized = str(direction).strip().lower()
    if direction_normalized in {"credit", "kredit", "cr"}:
        return FORMULA_CREDIT_MINUS_DEBIT
    if direction_normalized in {"debit", "dr"}:
        return FORMULA_DEBIT_MINUS_CREDIT

    return FORMULA_DEBIT_MINUS_CREDIT


def normalize_account_formula_map(account_formula_map: dict | None) -> dict[str, str]:
    normalized_map = {}
    if not account_formula_map:
        return normalized_map

    for account_name, formula_key in account_formula_map.items():
        normalized_formula = normalize_formula_key(formula_key)
        if normalized_formula:
            normalized_map[str(account_name).strip()] = normalized_formula

    return normalized_map


def get_gl_account_options(k3_sheets: dict) -> list[dict[str, str]]:
    k3 = _prepare_k3_dataframe(k3_sheets)
    options = []

    for account_name in k3["ACCOUNT_NAME"].dropna().astype(str).str.strip().unique():
        if not account_name or account_name.lower() in {"nan", "none"}:
            continue

        account_rows = k3[k3["ACCOUNT_NAME"].astype(str).str.strip() == account_name]
        direction = ""
        if "DIRECTION" in account_rows.columns:
            direction_candidates = (
                account_rows["DIRECTION"].dropna().astype(str).str.strip().tolist()
            )
            direction = next((value for value in direction_candidates if value), "")

        options.append(
            {
                "account_name": account_name,
                "direction": direction,
                "default_formula": _get_default_formula_for_account(
                    account_name, direction
                ),
            }
        )

    return options


def _calculate_net_by_formula(debit: float, credit: float, formula_key: str) -> float:
    if formula_key == FORMULA_CREDIT_MINUS_DEBIT:
        return -debit + credit
    if formula_key == FORMULA_NEGATIVE_DEBIT_MINUS_CREDIT:
        return -(debit - credit)
    return debit - credit


def _calculate_difference_from_net(net_value: float, dpp_value: float) -> float:
    return float(net_value) - float(dpp_value)


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


def calculate_net(row, account_formula_map: dict | None = None):
    debit = float(row.get("DEBIT_AMOUNT", row.get("Debit Amount", 0)))
    credit = float(row.get("CREDIT_AMOUNT", row.get("Credit Amount", 0)))
    acc_name = str(row.get("ACCOUNT_NAME", row.get("Account Name", ""))).strip()
    direction = str(row.get("DIRECTION", row.get("Direction", ""))).strip()

    normalized_formula_map = account_formula_map or {}
    formula_key = normalized_formula_map.get(acc_name) or _get_default_formula_for_account(
        acc_name, direction
    )
    return _calculate_net_by_formula(debit, credit, formula_key)


def compare_files(
    k3_sheets: dict,
    coretax_sheets_1: dict,
    coretax_sheets_2: dict,
    output_dir: str,
    progress_callback=None,
    account_formula_map: dict | None = None,
) -> str:
    # k3_sheets, coretax_sheets_1, coretax_sheets_2 are already dicts of {sheet_name: DataFrame}
    # from pd.read_excel(..., sheet_name=None) in app.py — no need to re-read.
    def _emit_progress(progress: int, message: str):
        if progress_callback is None:
            return
        try:
            progress_callback(progress, message)
        except Exception as callback_error:
            # Progress reporting failure should never stop comparison process.
            print(f"Progress callback error: {callback_error}")

    _emit_progress(0, "Preparing comparison data...")

    normalized_formula_map = normalize_account_formula_map(account_formula_map)
    k3 = _prepare_k3_dataframe(k3_sheets)
    print(f"K3 combined shape: {k3.shape}, columns: {list(k3.columns)}")
    _emit_progress(8, "Normalizing GL source data...")

    # 2) BARU terapkan ekstraksi No. Faktur & Nett pada variabel 'k3'
    k3["No Faktur (key)"] = k3.apply(
        lambda row: extract_no_faktur_from_description(
            row.get("DESCRIPTION", ""), row.get("VOUCHER_CATEGORY", "")
        ),
        axis=1,
    )
    # Bersihkan spasi agar tidak meleset saat merge
    k3["No Faktur (key)"] = k3["No Faktur (key)"].astype(str).str.strip()
    k3["Nett"] = k3.apply(
        lambda row: calculate_net(row, normalized_formula_map), axis=1
    )

    # 2) Normalize columns for Coretax (biar NO VOUCHER / DOC_NO kebaca konsisten)
    coretax_1 = pd.concat(
        [_normalize_columns(sheet_data) for sheet_data in coretax_sheets_1.values()],
        ignore_index=True,
    )
    coretax_2 = pd.concat(
        [_normalize_columns(sheet_data) for sheet_data in coretax_sheets_2.values()],
        ignore_index=True,
    )
    _emit_progress(18, "Normalizing Coretax data...")

    # Pastikan key jadi NO_VOUCHER
    coretax_1 = _ensure_column_from_aliases(
        coretax_1, "NO_VOUCHER", ["NO_VOUCHER", "DOC_NO"], required=True
    )
    coretax_2 = _ensure_column_from_aliases(
        coretax_2, "NO_VOUCHER", ["NO_VOUCHER", "DOC_NO"], required=True
    )

    # 3) Harmonize DPP/PPN + CUSTOMER + status
    # --- Digunggung
    coretax_1 = _ensure_column_from_aliases(
        coretax_1, "DPP", ["DPP", "AMOUNT_BEF_TAX"], required=False
    )
    coretax_1 = _ensure_column_from_aliases(
        coretax_1, "PPN", ["PPN", "TAX_AMOUNT"], required=False
    )
    coretax_1 = _ensure_column_from_aliases(
        coretax_1, "CUSTOMER", ["CUSTOMER", "DEPT", "CUSTOMER_NAME", "NPWP_NAME_DOC"]
    )
    coretax_1 = _ensure_column_from_aliases(
        coretax_1, "FP_STATUS", ["FP_STATUS", "TAX_STATUS"]
    )
    coretax_1["FP_STATUS"] = coretax_1["FP_STATUS"].fillna("FP Digunggung")

    # --- Tidak Digunggung
    coretax_2 = _ensure_column_from_aliases(
        coretax_2, "DPP", ["DPP", "AMOUNT_BEF_TAX"], required=False
    )
    coretax_2 = _ensure_column_from_aliases(
        coretax_2, "PPN", ["PPN", "TAX_AMOUNT"], required=False
    )
    coretax_2 = _ensure_column_from_aliases(
        coretax_2, "CUSTOMER", ["CUSTOMER", "NAMA_PEMBELI", "CUSTOMER_NAME"]
    )
    coretax_2 = _ensure_column_from_aliases(
        coretax_2, "FP_STATUS", ["FP_STATUS", "STATUS_FAKTUR", "TAX_STATUS"]
    )
    coretax_2["FP_STATUS"] = coretax_2["FP_STATUS"].fillna("FP Tidak Digunggung")

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
            # Ambil nominal GL sesuai formula akun yang dipilih user
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
    _emit_progress(30, "Merging Coretax datasets...")

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

    # 8) Tentukan kolom nomor faktur pajak yang tersedia di Coretax
    print("Columns in Coretax_2:", coretax_2.columns)
    nomor_faktur_pajak_col = None
    for candidate in ["NOMOR_FAKTUR_PAJAK", "NO_FP_MODIF"]:
        if candidate in coretax_2.columns:
            nomor_faktur_pajak_col = candidate
            break

    if nomor_faktur_pajak_col:
        print(f"Kolom nomor faktur pajak yang dipakai: {nomor_faktur_pajak_col}")
    else:
        print("Kolom nomor faktur pajak tidak ditemukan di Coretax_2")

    # 10) Merge
    merged = pd.merge(
        k3,
        coretax_agg,
        left_on="No Faktur (key)",
        right_on="NO_VOUCHER",
        how="left",
        indicator=True,
    )
    _emit_progress(42, "Building reconciliation results...")

    # 11) Compute Difference based on account type
    merged["Debit Amount"] = pd.to_numeric(
        merged["DEBIT_AMOUNT"], errors="coerce"
    ).fillna(0)
    merged["Credit Amount"] = pd.to_numeric(
        merged["CREDIT_AMOUNT"], errors="coerce"
    ).fillna(0)

    # Apply the Net calculation based on the account type
    merged["Net"] = merged.apply(
        lambda row: calculate_net(row, normalized_formula_map), axis=1
    )

    merged["DPP"] = pd.to_numeric(merged["DPP"], errors="coerce").fillna(0)
    merged["PPN"] = pd.to_numeric(merged["PPN"], errors="coerce").fillna(0)

    # Difference mengikuti formula Net akun yang dipilih user
    def calculate_difference(row):
        return _calculate_difference_from_net(row["Net"], row["DPP"])

    # Terapkan fungsi ke kolom Difference
    merged["Difference"] = merged.apply(calculate_difference, axis=1)

    # 12) Keterangan + Customer (langsung dari kolom kanonik)
    merged["Keterangan (Digunggung/Tidak Digunngung)"] = merged["FP_STATUS"]
    merged.loc[
        merged["_merge"] != "both", "Keterangan (Digunggung/Tidak Digunngung)"
    ] = "Tidak ada di Coretax"

    merged["Customer"] = merged["CUSTOMER"]
    merged.loc[merged["_merge"] != "both", "Customer"] = None

    # Debugging: tampilkan beberapa data nomor faktur pajak bila kolom tersedia
    if nomor_faktur_pajak_col:
        print(coretax_2[nomor_faktur_pajak_col].head())
    else:
        print("Kolom nomor faktur pajak tidak ditemukan di coretax_2")

    print("Setelah merge, kolom di merged:", merged.columns)

    # Jika kolom nomor faktur pajak ada di coretax_2, tambahkan ke merged
    if nomor_faktur_pajak_col:
        # PENTING: Drop duplicate agar tidak terjadi ledakan data (Cartesian product) saat merge
        coretax_2_unique = coretax_2[
            ["NO_VOUCHER", nomor_faktur_pajak_col]
        ].drop_duplicates(subset=["NO_VOUCHER"])
        merged = pd.merge(
            merged,
            coretax_2_unique,
            left_on="No Faktur (key)",
            right_on="NO_VOUCHER",
            how="left",
            suffixes=("", "_from_coretax2"),
        )
        merged = merged.rename(columns={nomor_faktur_pajak_col: "NOMOR_FAKTUR_PAJAK"})
        print("Setelah merge NOMOR_FAKTUR_PAJAK, kolom di merged:", merged.columns)
    else:
        merged["NOMOR_FAKTUR_PAJAK"] = None  # Atur sebagai None jika kolom tidak ada

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
    _emit_progress(50, "Preparing output workbook...")

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
        for ch in ["[", "]", ":", "*", "?", "/", "\\"]:
            name = name.replace(ch, "_")
        return name.strip()[:31]

    def _ensure_table_headers_are_strings(
        worksheet, header_row_idx: int, first_col_idx: int, last_col_idx: int
    ):
        seen_headers = set()

        for col_idx in range(first_col_idx, last_col_idx + 1):
            cell = worksheet.cell(header_row_idx, col_idx)
            header_value = cell.value

            if header_value is None or str(header_value).strip() == "":
                header_text = f"Column_{col_idx}"
            else:
                header_text = str(header_value).strip()

            original_header_text = header_text
            suffix = 1
            while header_text in seen_headers:
                suffix += 1
                header_text = f"{original_header_text}_{suffix}"

            cell.value = header_text
            seen_headers.add(header_text)

    # --- PERBAIKAN: PRE-CALCULATE DATA UNTUK MENCEGAH LOOPING LAMBAT ---
    print("Menyiapkan dictionary untuk mempercepat proses perhitungan...")

    # 1. Jadikan Set agar pencarian 'in' berjalan sekejap mata
    coretax_voucher_set = set(coretax_agg["NO_VOUCHER"].dropna().astype(str))
    duplicate_voucher_set = set(
        merged["NO_VOUCHER"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda series: ~series.isin(["", "-", "nan", "None", "none"])]
        .value_counts()
        .loc[lambda counts: counts > 1]
        .index
    )

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
    processed_coretax = (
        set()
    )  # Untuk melacak voucher mana yang sudah muncul data Coretax-nya
    duplicate_group_map = {
        voucher: f"Duplicate {idx}"
        for idx, voucher in enumerate(sorted(duplicate_voucher_set), start=1)
    }
    sheet_row_counts = {}
    first_sheet = True

    # Group by Account Name, pertahankan urutan kemunculan
    account_groups = merged.groupby("ACCOUNT_NAME", sort=False)

    total_rows_to_write = len(merged)
    rows_written = 0
    last_emitted_progress = 52

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

                v_bal = _parse_id_number(row.get("BALANCE", 0))

                # --- PERBAIKAN: Bersihkan NaN agar tidak dianggap teks "nan" ---
                voucher_no = str(row.get("NO_VOUCHER", "-")).strip()
                if voucher_no.lower() in ["nan", "none", "", "-"]:
                    voucher_no = "-"

                # --- SISI GL (KIRI): TETAP FLOAT (MENYIMPAN KOMA) ---
                row_debit = float(row.get("Debit Amount", 0))
                row_credit = float(row.get("Credit Amount", 0))
                row_net = float(row.get("Net", 0))
                display_row_net = row_net

                # Selalu tambahkan data GL ke subtotal
                debit_total += row_debit
                credit_total += row_credit
                net_total += row_net
                balance_total += v_bal

                # Cek apakah voucher valid ini ada di Coretax
                voucher_no_in_coretax = (voucher_no != "-") and (
                    voucher_no in coretax_voucher_set
                )
                voucher_no_is_duplicate = voucher_no in duplicate_voucher_set

                def build_match_status():
                    if not voucher_no_is_duplicate:
                        return "Unique"

                    return duplicate_group_map.get(voucher_no, "Duplicate")

                # --- LOGIKA PENENTUAN STATUS & DIFFERENCE ---
                row_diff_formula = ""  # Variabel untuk menampung rumus Excel dinamis

                if voucher_no == "-":
                    # KONDISI 1: DATA KOSONG
                    row_dpp = 0.0
                    row_ppn = 0.0

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
                        if voucher_no_is_duplicate:
                            display_row_net = total_gl_net if total_gl_net != 0 else row_net
                        if total_gl_net == 0:
                            row_diff = _calculate_difference_from_net(row_net, row_dpp)
                            row_diff_formula = f"=I{r} - O{r}"
                        else:
                            row_diff = _calculate_difference_from_net(
                                total_gl_net, row_dpp
                            )
                            row_diff_formula = f"={total_gl_net} - O{r}"

                        status = build_match_status()
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
                    if voucher_no_is_duplicate:
                        display_row_net = 0

                    if voucher_no_in_coretax:
                        status = build_match_status()
                    else:
                        status = "Tidak ada di Coretax"

                # --- PROSES CETAK KE SHEET ---
                current_ws.cell(r, 1).value = row.get("ACCOUNT_NO")
                current_ws.cell(r, 2).value = row.get("ACCOUNT_NAME")
                current_ws.cell(r, 3).value = row.get("DATE")
                current_ws.cell(r, 4).value = row.get("VOUCHER_CATEGORY")
                current_ws.cell(r, 5).value = row.get("VOUCHER_NO")
                current_ws.cell(r, 6).value = row.get("DESCRIPTION")
                current_ws.cell(r, 7).value = row_debit
                current_ws.cell(r, 8).value = row_credit
                current_ws.cell(r, 9).value = display_row_net
                current_ws.cell(r, 10).value = row.get("DIRECTION")
                current_ws.cell(r, 11).value = v_bal

                current_ws.cell(r, 13).value = voucher_no if voucher_no != "-" else None
                current_ws.cell(r, 14).value = (
                    row.get("NOMOR_FAKTUR_PAJAK")
                    if str(row.get("NOMOR_FAKTUR_PAJAK")) not in ["nan", "None"]
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
                rows_written += 1

                if total_rows_to_write > 0 and (
                    rows_written == total_rows_to_write or rows_written % 2000 == 0
                ):
                    dynamic_progress = 52 + int(
                        (rows_written / total_rows_to_write) * 36
                    )
                    dynamic_progress = min(88, dynamic_progress)
                    if dynamic_progress > last_emitted_progress:
                        last_emitted_progress = dynamic_progress
                        _emit_progress(
                            dynamic_progress,
                            f"Writing reconciliation rows ({rows_written}/{total_rows_to_write})...",
                        )

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

    _emit_progress(90, "Applying sheet table formatting...")

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
        _ensure_table_headers_are_strings(target_ws, header_row, 1, 20)
        last_data_row = start_row + row_count - 1
        table_ref = f"{first_col}{header_row}:{last_col}{last_data_row}"
        table_name = f"DataTable{table_idx}"
        table_idx += 1

        tab = Table(displayName=table_name, ref=table_ref)
        tab.tableStyleInfo = table_style
        target_ws.add_table(tab)

    print(
        f"Selesai! Data ditulis ke {len(sheet_row_counts)} sheet. Total baris: {len(merged)}"
    )

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    _emit_progress(96, "Saving comparison workbook...")

    # Save the workbook (preserves template formatting + all sheets)
    unique_suffix = uuid.uuid4().hex[:8]
    out_name = (
        f"Draft_Updated_Output_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{unique_suffix}.xlsx"
    )
    out_path = os.path.join(output_dir, out_name)
    wb.save(out_path)
    print(f"Saved output to {out_path}")
    _emit_progress(100, "Comparison file generated.")

    # Dapatkan daftar nama sheet dari workbook yang baru disimpan
    sheet_names = wb.sheetnames

    return out_path, out_name, sheet_names
