import os
import pandas as pd
import numpy as np
import re
import uuid
import threading
import time
import json
from urllib.parse import urlencode
from flask import (
    Blueprint,
    request,
    redirect,
    url_for,
    current_app,
    jsonify,
    Response,
    stream_with_context,
)
from werkzeug.utils import secure_filename
from helpers import allowed_file

recap_bp = Blueprint("recap_bp", __name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

RECAP_JOBS = {}
RECAP_JOBS_LOCK = threading.Lock()
RECAP_JOB_TTL_SECONDS = 1800


def _delete_recap_uploaded_files(file_paths):
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")


def _cleanup_recap_jobs():
    now = time.time()
    with RECAP_JOBS_LOCK:
        expired_job_ids = [
            job_id
            for job_id, job in RECAP_JOBS.items()
            if job.get("status") in {"done", "error"}
            and (now - job.get("updated_at", now)) > RECAP_JOB_TTL_SECONDS
        ]
        for job_id in expired_job_ids:
            RECAP_JOBS.pop(job_id, None)


def _create_recap_job(job_id):
    with RECAP_JOBS_LOCK:
        RECAP_JOBS[job_id] = {
            "job_id": job_id,
            "status": "queued",
            "progress": 0,
            "message": "Job dibuat.",
            "error": None,
            "redirect_url": None,
            "updated_at": time.time(),
        }


def _update_recap_job(
    job_id, progress=None, status=None, message=None, error=None, redirect_url=None
):
    with RECAP_JOBS_LOCK:
        job = RECAP_JOBS.get(job_id)
        if not job:
            return

        if progress is not None:
            job["progress"] = max(0, min(100, int(progress)))
        if status is not None:
            job["status"] = status
        if message is not None:
            job["message"] = message
        if error is not None:
            job["error"] = error
        if redirect_url is not None:
            job["redirect_url"] = redirect_url
        job["updated_at"] = time.time()


def _get_recap_job(job_id):
    with RECAP_JOBS_LOCK:
        job = RECAP_JOBS.get(job_id)
        return dict(job) if job else None


def _build_recap_redirect_url(out_name, sheet_list):
    query = urlencode(
        {
            "updated_file": out_name,
            "mode": "recap",
            "sheets": ",".join(sheet_list),
        }
    )
    return f"/comparison?{query}"


def _process_recap_job(job_id, source_path, ppn_path, output_dir):
    def _on_progress(recap_progress, recap_message):
        bounded = max(0, min(100, int(recap_progress)))
        _update_recap_job(
            job_id,
            progress=bounded,
            status="processing",
            message=recap_message,
        )

    _update_recap_job(
        job_id,
        progress=5,
        status="processing",
        message="File terunggah. Memulai proses recap...",
    )

    try:
        out_name, sheet_list = process_recap_2_files(
            source_path,
            ppn_path,
            output_dir,
            progress_callback=_on_progress,
        )
        redirect_url = _build_recap_redirect_url(out_name, sheet_list)
        _update_recap_job(
            job_id,
            progress=100,
            status="done",
            message="Recap selesai.",
            redirect_url=redirect_url,
        )
    except Exception as e:
        print(f"Error in recap job {job_id}: {e}")
        _update_recap_job(
            job_id,
            progress=100,
            status="error",
            message="Proses recap gagal.",
            error=str(e),
        )
    finally:
        _delete_recap_uploaded_files([source_path, ppn_path])


def process_recap_2_files(source_path, ppn_path, output_dir, progress_callback=None):
    def _emit_progress(progress: int, message: str):
        if progress_callback is None:
            return
        try:
            progress_callback(progress, message)
        except Exception as callback_error:
            print(f"Progress callback error: {callback_error}")

    try:
        _emit_progress(10, "Membaca file Rekap PPN...")
        # Daftar urutan akun default agar posisinya di Excel tetap sesuai template
        template_accounts = [
            "Sales",
            "Repair Service Income",
            "Sales Return",
            "Sales Price Protection",
            "POP Expense",
            "Promotion Gift",
            "Interest Bank Income",
            "Other Income",
            "Rental Income",
        ]
        
        month_labels = [
            "1. Januari", "2. Februari", "3. Maret", "4. April",
            "5. Mei", "6. Juni", "7. Juli", "8. Agustus",
            "9. September", "10. Oktober", "11. Nopember", "12. Desember",
        ]

        # --- 1. AMBIL DATA DARI REKAP PPN ---
        df_ppn_raw = pd.read_excel(ppn_path, header=None)
        ppn_monthly_values = {}
        
        # Cari baris yang mengandung 'jumlah penyerahan'
        mask = df_ppn_raw.apply(lambda row: row.astype(str).str.strip().str.lower().eq("jumlah penyerahan").any(), axis=1)
        if mask.any():
            target_row = df_ppn_raw[mask].iloc[0]
            for m in range(1, 13):
                val = target_row.iloc[5 + m]
                ppn_monthly_values[m] = float(val) if pd.notnull(val) else 0

        _emit_progress(22, "Mengambil data jumlah penyerahan Rekap PPN...")

        # --- 2. HITUNG DATA DARI DRAFT GL ---
        _emit_progress(30, "Membaca file Draft GL...")
        xls_gl = pd.ExcelFile(source_path)
        all_gl_dfs = []
        for sheet in xls_gl.sheet_names:
            df_raw = pd.read_excel(xls_gl, sheet_name=sheet, header=None)
            
            header_mask = df_raw.isin(['Account Name']).any(axis=1)
            
            if header_mask.any():
                h_row = header_mask.idxmax()
                df = df_raw.iloc[h_row + 1 :].copy()
                df.columns = [str(x).strip() for x in df_raw.iloc[h_row]]
                df = df.loc[:, ~df.columns.duplicated()] 
                all_gl_dfs.append(df)

        if not all_gl_dfs:
            raise ValueError("Tidak ada satupun sheet yang memiliki kolom 'Account Name'.")

        _emit_progress(42, "Menggabungkan sheet GL...")
        df_combined = pd.concat(all_gl_dfs, ignore_index=True)
        df_combined["Date"] = pd.to_datetime(df_combined["Date"], dayfirst=True, errors="coerce")
        df_combined["Month"] = df_combined["Date"].dt.month
        
        # --- MAPPING DINAMIS ACCOUNT NO & ACCOUNT NAME ---
        # Mencari kolom yang mengandung kata 'Account No' atau 'Account Num'
        acc_no_col = next((col for col in df_combined.columns if 'account no' in str(col).lower() or 'account num' in str(col).lower()), None)
        account_display_map = {}
        
        if acc_no_col:
            # Ambil pasangan Account Name dan Account No yang unik
            for _, row in df_combined.drop_duplicates(subset=['Account Name']).iterrows():
                acc_name = str(row.get('Account Name', '')).strip()
                acc_no = str(row.get(acc_no_col, '')).strip()
                
                # Cek jika Account No valid dan bukan 'nan'
                if acc_no and acc_no.lower() not in ['nan', 'none', '']:
                    account_display_map[acc_name] = f"{acc_no} - {acc_name}"
                else:
                    account_display_map[acc_name] = acc_name
        
        # --- KALKULASI NETT MASSAL ---
        df_combined['Debit Amount'] = pd.to_numeric(df_combined['Debit Amount'], errors='coerce').fillna(0)
        df_combined['Credit Amount'] = pd.to_numeric(df_combined['Credit Amount'], errors='coerce').fillna(0)
        acc_names = df_combined['Account Name'].astype(str).str.strip()

        cond_income = acc_names.isin([
            "Interest Bank Income", "Other Income", "Other Operating Income", 
            "Rental Income", "Repair Service Income", "Sales", "Sales Price Protection"
        ])
        cond_expense = acc_names.isin(["POP Expense", "Promotion Gift"])
        cond_return = acc_names == "Sales Return"

        df_combined["Calculated_Nett"] = np.select(
            [cond_income, cond_expense, cond_return],
            [
                -df_combined['Debit Amount'] + df_combined['Credit Amount'],
                df_combined['Debit Amount'] - df_combined['Credit Amount'],
                -(df_combined['Debit Amount'] - df_combined['Credit Amount'])
            ],
            default=0
        )

        _emit_progress(52, "Menghitung ringkasan per akun dan bulan...")
        summary_month = df_combined.groupby(["Account Name", "Month"])["Calculated_Nett"].sum().to_dict()
        summary_total_acc = df_combined.groupby("Account Name")["Calculated_Nett"].sum().to_dict()
        
        # Urutan akun mengikuti template, akun lain tetap ditambahkan di belakang.
        extra_accounts = sorted(
            [
                acc for acc in summary_total_acc.keys()
                if acc not in template_accounts and summary_total_acc[acc] != 0
            ]
        )
        # Saring template_accounts yang benar-benar ada transaksinya agar tidak error
        valid_template_accounts = [acc for acc in template_accounts if acc in summary_total_acc and summary_total_acc[acc] != 0]
        unique_accounts = valid_template_accounts + extra_accounts

        # Siapkan Data untuk Grid Bulanan
        monthly_data = []
        grand_ppn = 0
        grand_gl = 0
        grand_selisih = 0
        
        for m in range(1, 13):
            m_ppn = ppn_monthly_values.get(m, 0)
            row_data = {'Bulan': m, 'Rekap PPN': m_ppn}
            grand_ppn += m_ppn
            
            m_gl_total = 0
            for acc in unique_accounts:
                val = summary_month.get((acc, m), 0)
                row_data[acc] = val
                m_gl_total += val
                
            row_data['Total GL'] = m_gl_total
            grand_gl += m_gl_total
            
            m_selisih = m_ppn - m_gl_total
            row_data['Selisih'] = m_selisih
            grand_selisih += m_selisih
            
            monthly_data.append(row_data)
            
        df_monthly = pd.DataFrame(monthly_data)
        
        # Siapkan Baris Total
        total_row = {'Bulan': 'TOTAL', 'Rekap PPN': grand_ppn}
        for acc in unique_accounts:
            total_row[acc] = summary_total_acc.get(acc, 0)
        total_row['Total GL'] = grand_gl
        total_row['Selisih'] = grand_selisih
        
        df_monthly = pd.concat([df_monthly, pd.DataFrame([total_row])], ignore_index=True)

        _emit_progress(62, "Menyiapkan grid bulanan dan selisih...")

        # --- 3. BUAT FILE EXCEL SECARA DINAMIS TANPA TEMPLATE ---
        out_name = f"Final_Ekualisasi_{os.path.basename(source_path)}"
        if not out_name.endswith('.xlsx'):
            out_name = os.path.splitext(out_name)[0] + '.xlsx'
            
        out_path = os.path.join(output_dir, out_name)

        _emit_progress(72, "Membuat dan menulis file Excel Summary...")
        with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Summary Ekualisasi')
            writer.sheets['Summary Ekualisasi'] = worksheet
            worksheet.hide_gridlines(2)

            title_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'
            })
            section_format = workbook.add_format({
                'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter', 'border': 1
            })
            section_right_format = workbook.add_format({
                'bold': True, 'font_size': 12, 'align': 'right', 'valign': 'vcenter', 'border': 1
            })
            text_format = workbook.add_format({'valign': 'vcenter'})
            summary_idx_format = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'border': 1})
            summary_text_format = workbook.add_format({'valign': 'vcenter', 'border': 1})
            summary_num_format = workbook.add_format({'num_format': '#,##0', 'valign': 'vcenter', 'border': 1})
            left_line_text = workbook.add_format({'valign': 'vcenter', 'border': 1})
            num_format = workbook.add_format({'num_format': '#,##0', 'valign': 'vcenter'})
            bold_text = workbook.add_format({'bold': True, 'valign': 'vcenter'})
            bold_num_format = workbook.add_format({'bold': True, 'num_format': '#,##0', 'valign': 'vcenter'})
            total_line_text = workbook.add_format({'bold': True, 'valign': 'vcenter', 'border': 1})
            total_line_num = workbook.add_format({'bold': True, 'num_format': '#,##0', 'valign': 'vcenter', 'border': 1})
            header_group_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True
            })
            header_sub_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1
            })
            border_text = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
            border_num = workbook.add_format({'border': 1, 'num_format': '#,##0', 'valign': 'vcenter'})
            border_num_red = workbook.add_format({
                'border': 1, 'bold': True, 'font_color': 'red', 'num_format': '#,##0', 'valign': 'vcenter'
            })
            note_label = workbook.add_format({'valign': 'vcenter', 'border': 1})
            note_num = workbook.add_format({'num_format': '#,##0', 'valign': 'vcenter', 'border': 1})
            note_label_bold = workbook.add_format({'bold': True, 'valign': 'vcenter', 'border': 1})
            note_num_bold = workbook.add_format({'bold': True, 'num_format': '#,##0', 'valign': 'vcenter', 'border': 1})

            width_map = {
                0: 13, 1: 31, 2: 26, 3: 4.8, 4: 21, 5: 14.6, 6: 4.3, 7: 21, 8: 14.6, 9: 4.3,
                10: 21, 11: 14.6, 12: 4.3, 13: 21, 14: 14.6, 15: 4.3, 16: 21, 17: 14.6, 18: 4.3,
                19: 21, 20: 14.6, 21: 4.3, 22: 21, 23: 14.6, 24: 4.3, 25: 21, 26: 14.6, 27: 4.3,
                28: 21, 29: 14.6, 30: 4.3, 31: 23.5, 32: 17.6, 33: 3.5
            }
            for col_idx, width in width_map.items():
                worksheet.set_column(col_idx, col_idx, width)

            worksheet.merge_range(0, 0, 0, 33, 'Ekualisasi Peredaran Usaha dengan DPP Penyerahan PPN', title_format)
            worksheet.merge_range(1, 0, 1, 33, 'PT. WORLD INNOVATIVE TELECOMMUNICATION', title_format)
            worksheet.merge_range(6, 1, 6, 2, 'Peredaran Usaha cfm. SPT Tahunan PPh Badan', section_format)
            worksheet.write(7, 1, 'terdiri dari:', bold_text)

            total_all_acc = 0
            for i, acc in enumerate(unique_accounts, start=1):
                row_idx = 7 + i
                acc_total = summary_total_acc.get(acc, 0)
                worksheet.write_number(row_idx, 0, i, summary_idx_format)
                
                # Menggunakan Nama Display Gabungan (Account No + Name)
                display_name = account_display_map.get(acc, acc)
                worksheet.write(row_idx, 1, display_name, summary_text_format)
                
                worksheet.write(row_idx, 2, acc, summary_text_format)
                worksheet.write_blank(row_idx, 3, None, summary_text_format)
                worksheet.write_number(row_idx, 4, acc_total, summary_num_format)
                total_all_acc += acc_total

            total_top_row = 7 + len(unique_accounts) + 2
            worksheet.merge_range(
                total_top_row, 1, total_top_row, 31,
                'Total Peredaran Usaha cfm. SPT Tahunan PPh Badan (Lihat Laporan Keuangan)',
                section_right_format
            )
            worksheet.write_blank(total_top_row, 32, None, section_right_format)
            worksheet.write_number(total_top_row, 33, total_all_acc, total_line_num)

            # Header grid bulanan.
            ppn_label_col = 1
            ppn_amount_col = 2
            block_starts = [4 + 3 * idx for idx in range(len(unique_accounts))]
            total_gl_col = block_starts[-1] + 3 if block_starts else 4
            selisih_col = total_gl_col + 2

            worksheet.merge_range(22, ppn_label_col, 23, ppn_amount_col, 'Penyerahan (lokal & Ekspor) cfm. SPT Masa PPN (DPP)', header_group_format)

            for acc, start_col in zip(unique_accounts, block_starts):
                # Menampilkan Gabungan Account No & Name di Header Bulanan
                display_name = account_display_map.get(acc, acc)
                worksheet.merge_range(
                    22, start_col, 22, start_col + 1,
                    f"{display_name}",
                    header_group_format
                )
                worksheet.write(23, start_col, 'Month', header_sub_format)
                worksheet.write(23, start_col + 1, 'Amount', header_sub_format)

            worksheet.write(22, total_gl_col, 'TOTAL', header_group_format)
            worksheet.write(22, selisih_col, 'Selisih', header_group_format)

            month_start_row = 24
            for month_no, month_label in enumerate(month_labels, start=1):
                row_idx = month_start_row + month_no - 1
                monthly_row = df_monthly.iloc[month_no - 1]

                worksheet.write(row_idx, ppn_label_col, f'{month_label} (SPT Masa)', border_text)
                worksheet.write_number(row_idx, ppn_amount_col, monthly_row['Rekap PPN'], border_num)

                total_month_gl = 0
                for acc, start_col in zip(unique_accounts, block_starts):
                    acc_val = monthly_row.get(acc, 0)
                    worksheet.write(row_idx, start_col, month_label, border_text)
                    worksheet.write_number(row_idx, start_col + 1, acc_val, border_num)
                    total_month_gl += acc_val

                worksheet.write_number(row_idx, total_gl_col, total_month_gl, border_num)
                selisih_val = monthly_row['Selisih']
                if selisih_val != 0:
                    worksheet.write_number(row_idx, selisih_col, selisih_val, border_num_red)
                else:
                    worksheet.write_number(row_idx, selisih_col, selisih_val, border_num)

            total_monthly_row = month_start_row + 14
            worksheet.write(total_monthly_row, 1, 'Total Penyerahan', total_line_text)
            worksheet.write_number(total_monthly_row, ppn_amount_col, grand_ppn, total_line_num)

            for acc, start_col in zip(unique_accounts, block_starts):
                worksheet.write_blank(total_monthly_row, start_col, None, total_line_text)
                worksheet.write_number(total_monthly_row, start_col + 1, summary_total_acc.get(acc, 0), total_line_num)

            worksheet.write_number(total_monthly_row, total_gl_col, grand_gl, total_line_num)
            if grand_selisih != 0:
                worksheet.write_number(total_monthly_row, selisih_col, grand_selisih, border_num_red)
            else:
                worksheet.write_number(total_monthly_row, selisih_col, grand_selisih, total_line_num)

            worksheet.write(total_monthly_row + 2, 1, 'Adjustment Audit:', note_label_bold)
            worksheet.write(total_monthly_row + 3, 1, 'RE Recon (Autior KAP)', note_label)
            worksheet.write_number(total_monthly_row + 3, ppn_amount_col, 0, note_num)
            worksheet.write(total_monthly_row + 4, 1, 'Adjustment (Auditor KAP)', note_label)
            worksheet.write_number(total_monthly_row + 4, ppn_amount_col, 0, note_num)
            worksheet.write(total_monthly_row + 6, 1, 'Total', note_label_bold)
            worksheet.write_number(total_monthly_row + 6, ppn_amount_col, grand_ppn, note_num_bold)
            worksheet.merge_range(total_monthly_row + 8, 1, total_monthly_row + 8, 2, 'Selisih Kotor Ekualisasi Peredaran Usaha', section_format)
            worksheet.write_number(total_monthly_row + 8, 4, grand_selisih, total_line_num)

            adjustment_labels = [
                'Penyerahan terutang PPN Selain Pendapatan',
                'Penjualan Asset',
                'Adjustment karena terjadi kesalahan penerbitan invoice/sales retur',
                'Selisih dari Faktur Pajak Akibat Adjustment',
                'Adjustment NR secara Sistem',
                'Retur Non-PKP',
                'Penyerahan Non Objek PPN',
                'Retur Barang atas Transaksi FP 07',
                'Nota Retur Tidak Dapat Diimpor',
                'Nota Retur & Sales Retur Beda Waktu',
                'Selisih Kurs Transaksi Export',
                'Selisih Adjustment Sales Cut Off 2021',
                'Pendapatan lain-lain',
                'Adjustment Audit',
            ]
            adj_row_start = total_monthly_row + 11
            for idx, label in enumerate(adjustment_labels):
                worksheet.write(adj_row_start + idx, 31, label, note_label)
                worksheet.write_blank(adj_row_start + idx, 32, None, note_label)
                worksheet.write_number(adj_row_start + idx, 33, 0, note_num)

            total_other_income_row = adj_row_start + len(adjustment_labels)
            worksheet.merge_range(total_other_income_row, 31, total_other_income_row, 32, 'Total Pendapatan Lainnya', section_format)
            worksheet.write_number(total_other_income_row, 33, 0, note_num_bold)

            final_row = total_other_income_row + 2
            worksheet.merge_range(final_row, 1, final_row + 2, 32, 'Selisih Bersih Ekualisasi Peredaran Usaha PT. WIT tahun 2021', section_format)
            worksheet.write_number(final_row, 33, grand_selisih, total_line_num)

            worksheet.merge_range(
                final_row + 3, 1, final_row + 3, 33,
                'Menurut hemat kami, Selisih Bersih Ekualisasi senilai Rp. ..........',
                bold_text
            )

        _emit_progress(96, "Menyimpan file hasil recap...")
        return out_name, ["Summary Ekualisasi"]

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


@recap_bp.route("/upload_recap/start", methods=["POST"])
def start_upload_recap():
    _cleanup_recap_jobs()

    if "k3_file" not in request.files or "ppn_file" not in request.files:
        return jsonify({"error": "No file part"}), 400

    source_file = request.files["k3_file"]
    ppn_file = request.files["ppn_file"]

    if source_file.filename == "" or ppn_file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    allowed_ext = current_app.config.get("ALLOWED_EXTENSIONS", set())
    if not (
        source_file
        and allowed_file(source_file.filename, allowed_ext)
        and ppn_file
        and allowed_file(ppn_file.filename, allowed_ext)
    ):
        return jsonify({"error": "Invalid file type"}), 400

    unique_id = str(uuid.uuid4())[:8]
    job_id = str(uuid.uuid4())

    s_name = f"{unique_id}_{secure_filename(source_file.filename)}"
    p_name = f"{unique_id}_{secure_filename(ppn_file.filename)}"

    upload_folder = current_app.config["UPLOAD_FOLDER"]
    s_path = os.path.join(upload_folder, s_name)
    p_path = os.path.join(upload_folder, p_name)

    output_dir = current_app.config["OUTPUT_RECAP_FOLDER"]
    os.makedirs(output_dir, exist_ok=True)

    try:
        source_file.save(s_path)
        ppn_file.save(p_path)
    except Exception as e:
        _delete_recap_uploaded_files([s_path, p_path])
        return jsonify({"error": f"Failed to save uploaded files: {str(e)}"}), 500

    _create_recap_job(job_id)
    _update_recap_job(
        job_id,
        progress=2,
        status="processing",
        message="File terunggah. Memulai proses recap...",
    )

    worker = threading.Thread(
        target=_process_recap_job,
        args=(job_id, s_path, p_path, output_dir),
        daemon=True,
    )
    worker.start()

    return jsonify({"job_id": job_id}), 202


@recap_bp.route("/upload_recap/progress/<job_id>", methods=["GET"])
def stream_recap_progress(job_id):
    def event_stream():
        last_payload = None

        while True:
            job = _get_recap_job(job_id)
            if not job:
                payload = {
                    "job_id": job_id,
                    "status": "error",
                    "progress": 100,
                    "message": "Progress session not found.",
                    "error": "Job not found or already expired.",
                    "redirect_url": None,
                }
                yield f"data: {json.dumps(payload)}\n\n"
                break

            payload = {
                "job_id": job["job_id"],
                "status": job["status"],
                "progress": job["progress"],
                "message": job["message"],
                "error": job["error"],
                "redirect_url": job["redirect_url"],
            }
            payload_text = json.dumps(payload)

            if payload_text != last_payload:
                yield f"data: {payload_text}\n\n"
                last_payload = payload_text
            else:
                yield ": keep-alive\n\n"

            if job["status"] in {"done", "error"}:
                break

            time.sleep(0.5)

        _cleanup_recap_jobs()

    response = Response(
        stream_with_context(event_stream()),
        mimetype="text/event-stream",
    )
    response.headers["Cache-Control"] = "no-cache"
    response.headers["X-Accel-Buffering"] = "no"
    return response
