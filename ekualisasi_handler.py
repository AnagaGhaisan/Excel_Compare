import pandas as pd
import numpy as np

def proses_ekualisasi(file_bupot, file_voucher, file_output):
    # Baca data (Skip 2 baris awal untuk Rincian Bukti Potong)
    df_bupot = pd.read_excel(file_bupot, header=2)
    df_voucher = pd.read_excel(file_voucher)

    # Pembersihan Key (No. SPF)
    if 'No. SPF' in df_bupot.columns and 'No. SPF' in df_voucher.columns:
        df_bupot['No. SPF'] = df_bupot['No. SPF'].astype(str).str.strip().str.upper()
        df_voucher['No. SPF'] = df_voucher['No. SPF'].astype(str).str.strip().str.upper()
    else:
        raise ValueError("Kolom 'No. SPF' tidak ditemukan di salah satu file!")

    if 'Area' in df_voucher.columns:
        df_voucher = df_voucher.rename(columns={'Area': 'Area_GL'})

    # Merge Data
    df_merged = pd.merge(df_bupot, df_voucher, on='No. SPF', how='outer')

    # Hitung Selisih
    if 'DASAR PENGENAAN PAJAK (Rp)' in df_merged.columns and 'Amount Voucher Category' in df_merged.columns:
        dpp = pd.to_numeric(df_merged['DASAR PENGENAAN PAJAK (Rp)'], errors='coerce').fillna(0)
        amount_vc = pd.to_numeric(df_merged['Amount Voucher Category'], errors='coerce').fillna(0)
        df_merged['Difference'] = dpp - amount_vc

    if 'PAJAK PENGHASILAN' in df_merged.columns and 'PPh amount voucher category' in df_merged.columns:
        pph = pd.to_numeric(df_merged['PAJAK PENGHASILAN'], errors='coerce').fillna(0)
        pph_vc = pd.to_numeric(df_merged['PPh amount voucher category'], errors='coerce').fillna(0)
        df_merged['Diff'] = pph - pph_vc

    # Susun Kolom Sesuai Format Output
    kolom_target = [
        'No. SPF', 'Source.Name', 'NO.', 'NOMOR BUKTI POTONG', 'NIK/NPWP', 'NAMA',
        'TANGGAL BUKTI POTONG', 'JENIS PAJAK', 'KODE OBJEK PAJAK', 'OBJEK PAJAK',
        'DASAR PENGENAAN PAJAK (Rp)', 'TINGKAT (%)', 'TARIF', 'PAJAK PENGHASILAN',
        'FASILITAS PERPAJAKAN', 'UANG PERSEDIAAN / PEMBAYARAN LANGSUNG (UNTUK WP INSTANSI PEMERINTAH DENGAN DANA APBN)',
        'NITKU / NOMOR IDENTITAS SUBUNIT ORGANISASI', 'NITKU', 'Area', 'STATUS',
        'KAP-KJS', 'REFERENSI', 'x', 'Expense Account', 'Amount Voucher Category',
        'Difference', 'Nama PT', 'Area_GL', 'Keterangan', 'PPh amount voucher category', 'Diff'
    ]

    for col in kolom_target:
        if col not in df_merged.columns:
            df_merged[col] = np.nan

    df_final = df_merged[kolom_target]
    df_final = df_final.rename(columns={'Area_GL': 'Area'})

    # Export ke Excel
    with pd.ExcelWriter(file_output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Ekualisasi PPh 23')