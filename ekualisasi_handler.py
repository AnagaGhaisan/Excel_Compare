import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook

def proses_ekualisasi(file_bupot, file_voucher, file_template, file_output):
    # 1. BACA DATA
    # Bukti Potong di skip 2 baris (mulai baris 3)
    df_bupot = pd.read_excel(file_bupot, header=2)
    # Voucher dibaca normal
    df_voucher = pd.read_excel(file_voucher)

    # 2. PEMBERSIHAN KEY (No. SPF)
    if 'No. SPF' in df_bupot.columns:
        df_bupot['No. SPF'] = df_bupot['No. SPF'].astype(str).str.strip().str.upper()
    else:
        raise ValueError("Kolom 'No. SPF' tidak ditemukan di Rincian Bukti Potong!")
        
    if 'No. SPF' in df_voucher.columns:
        df_voucher['No. SPF'] = df_voucher['No. SPF'].astype(str).str.strip().str.upper()
    else:
        raise ValueError("Kolom 'No. SPF' tidak ditemukan di Voucher Category / GL!")

    # 3. GABUNGKAN DATA (MERGE)
    # Menggunakan suffix agar kolom bernama sama (contoh 'Area') tidak bertubrukan
    merged = pd.merge(df_bupot, df_voucher, on='No. SPF', how='outer', suffixes=('_BUPOT', '_GL'))

    # Fungsi pembantu untuk mengambil nama kolom yang benar
    def get_col(col_name, is_gl=False):
        suffix = '_GL' if is_gl else '_BUPOT'
        suffixed_col = f"{col_name}{suffix}"
        if suffixed_col in merged.columns:
            return suffixed_col
        return col_name

    # 4. HITUNG SELISIH (DIFFERENCE) - DIPERBAIKI (SANGAT AMAN)
    dpp_col = get_col('DASAR PENGENAAN PAJAK (Rp)', is_gl=False)
    vc_col = get_col('Amount Voucher Category', is_gl=True)
    
    # Pastikan kolomnya benar-benar ada sebagai Series sebelum di-numeric-kan
    if dpp_col not in merged.columns: merged[dpp_col] = 0
    if vc_col not in merged.columns: merged[vc_col] = 0
    
    merged[dpp_col] = pd.to_numeric(merged[dpp_col], errors='coerce').fillna(0)
    merged[vc_col] = pd.to_numeric(merged[vc_col], errors='coerce').fillna(0)
    merged['Difference_Calc'] = merged[dpp_col] - merged[vc_col]

    pph_col = get_col('PAJAK PENGHASILAN', is_gl=False)
    pph_vc_col = get_col('PPh amount voucher category', is_gl=True)

    if pph_col not in merged.columns: merged[pph_col] = 0
    if pph_vc_col not in merged.columns: merged[pph_vc_col] = 0

    merged[pph_col] = pd.to_numeric(merged[pph_col], errors='coerce').fillna(0)
    merged[pph_vc_col] = pd.to_numeric(merged[pph_vc_col], errors='coerce').fillna(0)
    merged['Diff_Calc'] = merged[pph_col] - merged[pph_vc_col]

    # 5. BERSIHKAN NaN AGAR OPENPYXL TIDAK ERROR (Sama seperti di helpers.py)
    merged.fillna("", inplace=True)

    # 6. BUKA TEMPLATE EXCEL MENGGUNAKAN OPENPYXL
    if not os.path.exists(file_template):
        raise FileNotFoundError(f"Template tidak ditemukan di: {file_template}")

    wb = load_workbook(file_template)
    ws = wb.active

    # Mulai cetak data di baris ke-4 (karena 1-3 adalah header template)
    start_row = 4

    # 7. SATU LOOP UNTUK SEMUA
    for i in range(len(merged)):
        r = start_row + i
        row = merged.iloc[i]

        # Fungsi pembantu untuk extract data dengan aman per baris
        def get_val(base_col_name, is_gl=False):
            actual_col = get_col(base_col_name, is_gl)
            if actual_col in merged.columns:
                return row.get(actual_col, "")
            return ""

        # SISI KIRI: BUKTI POTONG
        ws.cell(r, 1).value = row.get('No. SPF', '')
        ws.cell(r, 2).value = get_val('Source.Name')
        ws.cell(r, 3).value = get_val('NO.')
        ws.cell(r, 4).value = get_val('NOMOR BUKTI POTONG')
        ws.cell(r, 5).value = get_val('NIK/NPWP')
        ws.cell(r, 6).value = get_val('NAMA')
        ws.cell(r, 7).value = get_val('TANGGAL BUKTI POTONG')
        ws.cell(r, 8).value = get_val('JENIS PAJAK')
        ws.cell(r, 9).value = get_val('KODE OBJEK PAJAK')
        ws.cell(r, 10).value = get_val('OBJEK PAJAK')
        ws.cell(r, 11).value = row.get(dpp_col, 0)
        ws.cell(r, 12).value = get_val('TINGKAT (%)')
        ws.cell(r, 13).value = get_val('TARIF')
        ws.cell(r, 14).value = row.get(pph_col, 0)
        ws.cell(r, 15).value = get_val('FASILITAS PERPAJAKAN')
        ws.cell(r, 16).value = get_val('UANG PERSEDIAAN / PEMBAYARAN LANGSUNG (UNTUK WP INSTANSI PEMERINTAH DENGAN DANA APBN)')
        ws.cell(r, 17).value = get_val('NITKU / NOMOR IDENTITAS SUBUNIT ORGANISASI')
        ws.cell(r, 18).value = get_val('NITKU')
        ws.cell(r, 19).value = get_val('Area')
        ws.cell(r, 20).value = get_val('STATUS')
        ws.cell(r, 21).value = get_val('KAP-KJS')
        ws.cell(r, 22).value = get_val('REFERENSI')

        # KOLOM TENGAH (Pemisah 'x')
        ws.cell(r, 23).value = ""

        # SISI KANAN: VOUCHER CATEGORY (GL)
        ws.cell(r, 24).value = get_val('Expense Account', True)
        ws.cell(r, 25).value = row.get(vc_col, 0)
        ws.cell(r, 26).value = row.get('Difference_Calc', 0)
        ws.cell(r, 27).value = get_val('Nama PT', True)
        ws.cell(r, 28).value = get_val('Area', True)
        ws.cell(r, 29).value = get_val('Keterangan', True)
        ws.cell(r, 30).value = row.get(pph_vc_col, 0)
        ws.cell(r, 31).value = row.get('Diff_Calc', 0)

    # 8. SIMPAN FILE
    # Menyimpan langsung file hasil modifikasi openpyxl ke target output
    wb.save(file_output)
    print(f"Selesai! File berhasil disimpan di: {file_output}")