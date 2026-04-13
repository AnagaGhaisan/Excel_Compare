import os
import pandas as pd
import numpy as np # <--- TAMBAHKAN INI
import openpyxl
import re
from flask import Blueprint, request, redirect, url_for, current_app
from werkzeug.utils import secure_filename
from helpers import allowed_file, delete_all_uploaded_files
import uuid

recap_bp = Blueprint("recap_bp", __name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "static/template/Ekualisasi_Darft_Output.xlsx")

# Catatan: Fungsi calculate_row_nett dihapus karena kita gunakan Numpy Select di bawah

def process_recap_2_files(source_path, ppn_path, output_dir):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template tidak ditemukan di path: {TEMPLATE_PATH}")

    try:
        # --- 1. AMBIL DATA DARI REKAP PPN ---
        df_ppn_raw = pd.read_excel(ppn_path, header=None)
        ppn_monthly_values = {}
        
        # Cari baris yang mengandung 'jumlah penyerahan' dengan cara vectorized (lebih cepat)
        mask = df_ppn_raw.apply(lambda row: row.astype(str).str.strip().str.lower().eq("jumlah penyerahan").any(), axis=1)
        if mask.any():
            target_row = df_ppn_raw[mask].iloc
            for m in range(1, 13):
                val = target_row.iloc[5 + m]
                ppn_monthly_values[m] = float(val) if pd.notnull(val) else 0

        # --- 2. HITUNG DATA DARI DRAFT GL ---
        xls_gl = pd.ExcelFile(source_path)
        all_gl_dfs = []
        for sheet in xls_gl.sheet_names:
            df_raw = pd.read_excel(xls_gl, sheet_name=sheet, header=None)
            
            # Cari baris header tanpa iterrows (Sangat Cepat)
            header_mask = df_raw.isin(['Account Name']).any(axis=1)
            
            if header_mask.any():
                h_row = header_mask.idxmax()
                df = df_raw.iloc[h_row + 1 :].copy()
                df.columns = [str(x).strip() for x in df_raw.iloc[h_row]]
                df = df.loc[:, ~df.columns.duplicated()] # Perbaikan warning pandas .loc
                all_gl_dfs.append(df)

        if not all_gl_dfs:
            raise ValueError("Tidak ada satupun sheet yang memiliki kolom 'Account Name'.")

        df_combined = pd.concat(all_gl_dfs, ignore_index=True)
        df_combined["Date"] = pd.to_datetime(df_combined["Date"], dayfirst=True, errors="coerce")
        df_combined["Month"] = df_combined["Date"].dt.month
        
        # --- OPTIMASI KALKULASI NETT MASSAL (MENGGANTIKAN calculate_row_nett) ---
        df_combined['Debit Amount'] = pd.to_numeric(df_combined['Debit Amount'], errors='coerce').fillna(0)
        df_combined['Credit Amount'] = pd.to_numeric(df_combined['Credit Amount'], errors='coerce').fillna(0)
        acc_names = df_combined['Account Name'].astype(str).str.strip()

        # Tentukan kondisi akun
        cond_income = acc_names.isin([
            "Interest Bank Income", "Other Income", "Other Operating Income", 
            "Rental Income", "Repair Service Income", "Sales", "Sales Price Protection"
        ])
        cond_expense = acc_names.isin(["POP Expense", "Promotion Gift"])
        cond_return = acc_names == "Sales Return"

        # Terapkan rumus secara massal menggunakan numpy
        df_combined["Calculated_Nett"] = np.select(
            [cond_income, cond_expense, cond_return],
            [
                -df_combined['Debit Amount'] + df_combined['Credit Amount'],
                df_combined['Debit Amount'] - df_combined['Credit Amount'],
                -(df_combined['Debit Amount'] - df_combined['Credit Amount'])
            ],
            default=0
        )
        # ------------------------------------------------------------------------

        summary_month = df_combined.groupby(["Account Name", "Month"])["Calculated_Nett"].sum().to_dict()
        summary_total_acc = df_combined.groupby("Account Name")["Calculated_Nett"].sum().to_dict()
        summary_all_month_total = df_combined.groupby("Month")["Calculated_Nett"].sum().to_dict()

        # --- 3. ISI TEMPLATE ---
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # Update Bagian Atas & Total AH20
        total_gl_year = 0
        for r in range(8, 20):
            acc_cell = ws.cell(row=r, column=3)
            if acc_cell.value:
                acc_name = str(acc_cell.value).strip()
                val = summary_total_acc.get(acc_name, 0)
                ws.cell(row=r, column=5).value = val
                ws.cell(row=r, column=5).number_format = "#,##0"
                total_gl_year += val
        ws.cell(row=20, column=34).value = total_gl_year

        # Grid Bulanan & Baris TOTAL (Baris 39)
        header_row_gl = 23
        month_start_row = 25
        total_row_idx = 39
        account_to_col = {}

        for c in range(1, ws.max_column + 1):
            h_val = ws.cell(row=header_row_gl, column=c).value
            if h_val and "cfm." in str(h_val):
                match = re.search(r"-\s*(.+)$", str(h_val))
                if match:
                    account_to_col[match.group(1).strip()] = c + 1

        grand_ppn = 0
        grand_gl_total = 0
        grand_selisih = 0
        acc_column_totals = {acc: 0 for acc in account_to_col.keys()}

        for m in range(1, 13):
            t_row = month_start_row + (m - 1)
            m_ppn = ppn_monthly_values.get(m, 0)
            ws.cell(row=t_row, column=3).value = m_ppn
            ws.cell(row=t_row, column=3).number_format = "#,##0"
            grand_ppn += m_ppn

            for acc, t_col in account_to_col.items():
                val = summary_month.get((acc, m), 0)
                ws.cell(row=t_row, column=t_col).value = val
                ws.cell(row=t_row, column=t_col).number_format = "#,##0"
                acc_column_totals[acc] += val

            m_gl_total = summary_all_month_total.get(m, 0)
            ws.cell(row=t_row, column=32).value = m_gl_total
            ws.cell(row=t_row, column=32).number_format = "#,##0"
            grand_gl_total += m_gl_total

            m_selisih = m_ppn - m_gl_total
            ws.cell(row=t_row, column=34).value = m_selisih
            ws.cell(row=t_row, column=34).number_format = "#,##0"
            grand_selisih += m_selisih
            if m_selisih != 0:
                ws.cell(row=t_row, column=34).font = openpyxl.styles.Font(
                    color="FF0000", bold=True
                )

        # ISI BARIS TOTAL (BARIS 39)
        ws.cell(row=total_row_idx, column=2).font = openpyxl.styles.Font(bold=True)
        ws.cell(row=total_row_idx, column=3).value = grand_ppn
        ws.cell(row=total_row_idx, column=3).number_format = "#,##0"

        for acc, t_col in account_to_col.items():
            ws.cell(row=total_row_idx, column=t_col).value = acc_column_totals[acc]
            ws.cell(row=total_row_idx, column=t_col).number_format = "#,##0"

        ws.cell(row=total_row_idx, column=32).value = grand_gl_total
        ws.cell(row=total_row_idx, column=32).number_format = "#,##0"
        ws.cell(row=total_row_idx, column=34).value = grand_selisih
        ws.cell(row=total_row_idx, column=34).number_format = "#,##0"

        for c in + list(account_to_col.values()):
            ws.cell(row=total_row_idx, column=c).font = openpyxl.styles.Font(bold=True)

        out_name = f"Final_Ekualisasi_{os.path.basename(source_path)}"
        out_path = os.path.join(output_dir, out_name)
        wb.save(out_path)

        return out_name, wb.sheetnames

    except Exception as e:
        raise ValueError(f"Proses Gagal: {str(e)}")

@recap_bp.route("/upload_recap", methods=["POST"])
def upload_recap():
    if "k3_file" not in request.files or "ppn_file" not in request.files:
        return "Kesalahan: File Draft (GL) dan Rekap PPN wajib diupload", 400

    source_file = request.files["k3_file"]
    ppn_file = request.files["ppn_file"]

    if source_file.filename == "" or ppn_file.filename == "":
        return "Kesalahan: Nama file tidak boleh kosong", 400

    if source_file and ppn_file:
        unique_id = str(uuid.uuid4())[:8]
        s_name = f"{unique_id}_{secure_filename(source_file.filename)}"
        p_name = f"{unique_id}_{secure_filename(ppn_file.filename)}"

        s_path = os.path.join(current_app.config["UPLOAD_FOLDER"], s_name)
        p_path = os.path.join(current_app.config["UPLOAD_FOLDER"], p_name)

        source_file.save(s_path)
        ppn_file.save(p_path)

        output_dir = current_app.config["OUTPUT_RECAP_FOLDER"]
        os.makedirs(output_dir, exist_ok=True)

        try:
            out_name, sheets = process_recap_2_files(s_path, p_path, output_dir)

            # --- BAGIAN PENGHAPUSAN DINONAKTIFKAN (DI-COMMENT) ---
            # delete_all_uploaded_files(current_app)

            return redirect(
                url_for(
                    "show_comparison",
                    updated_file=out_name,
                    mode="recap",
                    sheets=",".join(sheets),
                )
            )
        except Exception as e:
            return f"Terjadi kesalahan saat memproses: {str(e)}", 500

    return "Format file tidak valid", 400