import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =========================
# UTIL FUNCTIONS
# =========================

def ensure_column(df, col):
    if col not in df.columns:
        raise ValueError(f"Kolom '{col}' tidak ditemukan di file")

def normalisasi_teks(series):
    return series.astype(str)\
        .replace(r'\r+|\n+|\t+', ' ', regex=True)\
        .str.strip().str.upper()

def clean_numeric(value):
    if pd.isna(value):
        return 0.0

    str_val = str(value).strip()

    if str_val in ["", "-", "None"]:
        return 0.0

    # handle (1,000) => -1000
    if str_val.startswith("(") and str_val.endswith(")"):
        str_val = "-" + str_val[1:-1]

    str_val = (
        str_val.replace('Rp', '')
        .replace(' ', '')
        .replace('.', '')
        .replace(',', '.')
    )

    try:
        return float(str_val)
    except:
        return 0.0


# =========================
# MAIN PROCESS
# =========================

def proses_ekualisasi(file_bupot, file_voucher, file_template, file_output):

    print("🚀 Mulai proses ekualisasi...")

    # =========================
    # 1. LOAD DATA
    # =========================
    df_bupot = pd.read_excel(file_bupot, header=2)
    df_voucher = pd.read_excel(file_voucher)

    df_bupot.columns = df_bupot.columns.str.strip()
    df_voucher.columns = df_voucher.columns.str.strip()
    
    print("Kolom BUPOT:")
    print(df_bupot.columns.tolist())

    print("\nKolom VOUCHER:")
    print(df_voucher.columns.tolist())

    # =========================
    # 2. VALIDASI KOLOM
    # =========================
    ensure_column(df_bupot, 'No. SPF')
    ensure_column(df_voucher, 'No. SPF')

    # =========================
    # 3. NORMALISASI KEY
    # =========================
    df_bupot['No. SPF'] = normalisasi_teks(df_bupot['No. SPF'])
    df_voucher['No. SPF'] = normalisasi_teks(df_voucher['No. SPF'])

    # =========================
    # 4. DEFINE KOLOM
    # =========================
    col_dpp_bupot = 'DASAR PENGENAAN PAJAK (Rp)'
    col_pph_bupot = 'PAJAK PENGHASILAN'

    col_dpp_gl = 'Amount (Original Currency)'
    col_acc_gl = 'Full Account Name'
    col_pph_gl = 'PPh amount voucher category'
    col_desc_gl = 'Description'

    # =========================
    # 5. NORMALISASI ANGKA
    # =========================
    for col in [col_dpp_bupot, col_pph_bupot]:
        ensure_column(df_bupot, col)
        df_bupot[col] = df_bupot[col].apply(clean_numeric)

    ensure_column(df_voucher, col_dpp_gl)
    ensure_column(df_voucher, col_acc_gl)
    ensure_column(df_voucher, col_desc_gl)

    df_voucher[col_dpp_gl] = df_voucher[col_dpp_gl].apply(clean_numeric)

    if col_pph_gl in df_voucher.columns:
        df_voucher[col_pph_gl] = df_voucher[col_pph_gl].apply(clean_numeric)
    else:
        df_voucher[col_pph_gl] = 0.0

    # =========================
    # 6. AGGREGATE VOUCHER
    # =========================
    df_bupot.rename(columns={"Area": "Area_Bupot"}, inplace=True)
    df_voucher.rename(columns={"Area": "Area_GL"}, inplace=True)

    agg_dict = {
        col_dpp_gl: 'sum',
        col_pph_gl: 'sum',
        col_acc_gl: 'first',
        col_desc_gl: 'first'
    }

    if 'Nama PT' in df_voucher.columns:
        agg_dict['Nama PT'] = 'first'
    
    # Gunakan nama baru 'Area_GL' di sini
    if 'Area_GL' in df_voucher.columns:
        agg_dict['Area_GL'] = 'first'

    df_voucher_agg = df_voucher.groupby('No. SPF', as_index=False).agg(agg_dict)

    print(f"📊 Data Bupot: {len(df_bupot)}")
    print(f"📊 Data Voucher: {len(df_voucher)}")
    print(f"📊 Setelah Aggregasi: {len(df_voucher_agg)}")
    
    df_bupot.rename(columns={"Area": "Area_Bupot"}, inplace=True)
    df_voucher.rename(columns={"Area": "Area_GL"}, inplace=True)
    

    # =========================
    # 7. MERGE
    # =========================
    merged = pd.merge(df_bupot, df_voucher_agg, on='No. SPF', how='left')

    print(f"📊 Setelah Merge: {len(merged)}")
    print("Kolom merged:")
    print(merged.columns.tolist())

    # =========================
    # 8. HITUNG SELISIH
    # =========================
    merged['Difference_Calc'] = merged[col_dpp_bupot].fillna(0) - merged[col_dpp_gl].fillna(0)
    merged['Diff_PPh_Calc'] = merged[col_pph_bupot].fillna(0) - merged[col_pph_gl].fillna(0)

    # pisahkan object & numeric biar aman
    merged_obj = merged.select_dtypes(include='object').fillna("")
    merged_num = merged.select_dtypes(exclude='object').fillna(0)
    merged = pd.concat([merged_obj, merged_num], axis=1)

    # =========================
    # 9. WRITE TO EXCEL
    # =========================
    wb = load_workbook(file_template, data_only=True)

    # HAPUS external links (anti error Excel)
    if hasattr(wb, "_external_links"):
        wb._external_links = []

    ws = wb.active

    start_row = 4
    num_format = '#,##0'
    red_fill = PatternFill(start_color="FFC7CE", fill_type="solid")

    for i, row in merged.iterrows():
        r = start_row + i

        # =========================
        # BUPOT (KIRI)
        # =========================
        ws.cell(r, 1).value = row.get('No. SPF', '')                     # A
        ws.cell(r, 2).value = row.get('Source.Name', '')                 # B
        ws.cell(r, 3).value = row.get('NO.', '')                         # C
        ws.cell(r, 4).value = row.get('NOMOR BUKTI POTONG', '')          # D
        ws.cell(r, 5).value = row.get('NIK/NPWP', '')                    # E
        ws.cell(r, 6).value = row.get('NAMA', '')                        # F
        ws.cell(r, 7).value = row.get('TANGGAL BUKTI POTONG', '')        # G
        ws.cell(r, 8).value = row.get('JENIS PAJAK', '')                 # H
        ws.cell(r, 9).value = row.get('KODE OBJEK PAJAK', '')            # I
        ws.cell(r, 10).value = row.get('OBJEK PAJAK', '')                # J

        c11 = ws.cell(r, 11)
        c11.value = row.get('DASAR PENGENAAN PAJAK (Rp)', 0)             # K
        c11.number_format = num_format

        ws.cell(r, 12).value = row.get('TINGKAT (%)', '')                # L
        ws.cell(r, 13).value = row.get('TARIF', '')                      # M

        c14 = ws.cell(r, 14)
        c14.value = row.get('PAJAK PENGHASILAN', 0)                      # N
        c14.number_format = num_format
        
        ws.cell(r, 15).value = row.get('FASILITAS PERPAJAKAN', '')   # O
        ws.cell(r, 16).value = row.get('UANG PERSEDIAAN / PEMBAYARAN LANGSUNG (UNTUK WP INSTANSI PEMERINTAH DENGAN DANA APBN)', '')  # P
        ws.cell(r, 17).value = row.get('NITKU / NOMOR IDENTITAS SUBUNIT ORGANISASI', '')  # Q
        ws.cell(r, 18).value = row.get('NITKU', '')  # R
        ws.cell(r, 19).value = row.get('Area_Bupot', '')
        ws.cell(r, 20).value = row.get('STATUS', '') # T
        ws.cell(r, 21).value = row.get('KAP-KJS', '') # U
        ws.cell(r, 22).value = row.get('REFERENSI', '')                  # V

        # =========================
        # GL (KANAN)
        # =========================
        ws.cell(r, 24).value = row.get('Full Account Name', '')          # X

        c25 = ws.cell(r, 25)
        c25.value = row.get('Amount (Original Currency)', 0)             # Y
        c25.number_format = num_format

        c26 = ws.cell(r, 26)
        c26.value = row.get('Difference_Calc', 0)                        # Z
        c26.number_format = num_format

        ws.cell(r, 27).value = row.get('Nama PT', '')                    # AA
        ws.cell(r, 28).value = row.get('Area_GL', '')     # kanan
        ws.cell(r, 29).value = row.get('Description', '')                # AC

        c30 = ws.cell(r, 30)
        c30.value = row.get('PPh amount voucher category', 0)            # AD
        c30.number_format = num_format

        c31 = ws.cell(r, 31)
        c31.value = row.get('Diff_PPh_Calc', 0)                          # AE
        c31.number_format = num_format

        # =========================
        # HIGHLIGHT SELISIH
        # =========================
        if row.get('Difference_Calc', 0) != 0:
            c26.fill = red_fill

        if row.get('Diff_PPh_Calc', 0) != 0:
            c31.fill = red_fill

    # =========================
    # 10. AUTO FILTER & WIDTH
    # =========================
    last_row = start_row + len(merged) - 1

    if last_row >= 3:
        ws.auto_filter.ref = f"A3:AE{last_row}"

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[col_letter].width = max_length + 2

    # =========================
    # 11. SAVE
    # =========================
    wb.save(file_output)

    print("✅ Selesai! File tersimpan di:", file_output)