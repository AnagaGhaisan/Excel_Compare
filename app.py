from flask import Flask, request, render_template, send_file, redirect, url_for
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
from helpers import compare_files, allowed_file, delete_all_uploaded_files
from recap_handler import recap_bp
import pandas as pd
from waitress import serve
import uuid


app = Flask(__name__)
CORS(app)
app.register_blueprint(recap_bp)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["ALLOWED_EXTENSIONS"] = {"xls", "xlsx"}

# 1. Definisikan jalur folder
OUTPUT_COMPARE_DIR = os.path.join(BASE_DIR, "outputs", "compare")
OUTPUT_RECAP_DIR = os.path.join(BASE_DIR, "outputs", "recap")

# 3. Masukkan ke dalam config agar bisa dibaca oleh recap_handler.py
app.config["OUTPUT_COMPARE_FOLDER"] = os.path.join(OUTPUT_DIR, "compare")
app.config["OUTPUT_RECAP_FOLDER"] = os.path.join(OUTPUT_DIR, "recap")
app.config["UPLOAD_FOLDER"] = os.path.join(BASE_DIR, "uploads")

# 2. Pastikan folder fisik dibuat di server
os.makedirs(app.config["OUTPUT_COMPARE_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_RECAP_FOLDER"], exist_ok=True)
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# Homepage route to upload files
@app.route("/")
def home():
    return render_template("index.html")  # Landing page with two card options


@app.route("/filecompare")
def file_compare():
    return render_template("FileCompare/index.html")


@app.route("/filerecap")
def file_recap():
    return render_template("FileRecap/index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if (
        "k3_file" not in request.files
        or "coretax_file_1" not in request.files
        or "coretax_file_2" not in request.files
    ):
        return "No file part"

    k3_file = request.files["k3_file"]
    coretax_file_1 = request.files["coretax_file_1"]
    coretax_file_2 = request.files["coretax_file_2"]

    # Check if files have been selected
    if (
        k3_file.filename == ""
        or coretax_file_1.filename == ""
        or coretax_file_2.filename == ""
    ):
        return "No selected file"

    # Check if files are valid
    if (
        k3_file
        and allowed_file(k3_file.filename, app.config["ALLOWED_EXTENSIONS"])
        and coretax_file_1
        and allowed_file(coretax_file_1.filename, app.config["ALLOWED_EXTENSIONS"])
        and coretax_file_2
        and allowed_file(coretax_file_2.filename, app.config["ALLOWED_EXTENSIONS"])
    ):
        unique_id = str(uuid.uuid4())[:8]

        # 3. Tambahkan unique_id ke nama file
        k3_filename = f"{unique_id}_{secure_filename(k3_file.filename)}"
        coretax_filename_1 = f"{unique_id}_{secure_filename(coretax_file_1.filename)}"
        coretax_filename_2 = f"{unique_id}_{secure_filename(coretax_file_2.filename)}"

        # 4. Gabungkan dengan path folder upload
        k3_file_path = os.path.join(app.config["UPLOAD_FOLDER"], k3_filename)
        coretax_file_path_1 = os.path.join(
            app.config["UPLOAD_FOLDER"], coretax_filename_1
        )
        coretax_file_path_2 = os.path.join(
            app.config["UPLOAD_FOLDER"], coretax_filename_2
        )

        # 5. Simpan file fisik
        k3_file.save(k3_file_path)
        coretax_file_1.save(coretax_file_path_1)
        coretax_file_2.save(coretax_file_path_2)

        # 1) Read all sheets for K3 and Coretax files
        try:
            k3_sheets = pd.read_excel(
                k3_file_path, sheet_name=None, header=1
            )  # Ensure header is read from row 2
            coretax_sheets_1 = pd.read_excel(
                coretax_file_path_1, sheet_name=None, header=1
            )  # Read with header in row 2
            coretax_sheets_2 = pd.read_excel(
                coretax_file_path_2, sheet_name=None, header=1
            )  # Read with header in row 2

            print("K3 sheets:", k3_sheets.keys())
            print("Coretax 1 sheets:", coretax_sheets_1.keys())
            print("Coretax 2 sheets:", coretax_sheets_2.keys())

            # Debug: Check the shapes of the sheets after reading
            for sheet_name, sheet_data in k3_sheets.items():
                print(f"K3 - {sheet_name} shape:", sheet_data.shape)
            for sheet_name, sheet_data in coretax_sheets_1.items():
                print(f"Coretax 1 - {sheet_name} shape:", sheet_data.shape)
            for sheet_name, sheet_data in coretax_sheets_2.items():
                print(f"Coretax 2 - {sheet_name} shape:", sheet_data.shape)

        except Exception as e:
            print(f"Error reading files: {e}")
            return "Error reading the Excel files."

        # ... (bagian upload file di atas tetap sama) ...

        # 2. Definisikan variabel output_dir agar tidak error
        output_dir = app.config["OUTPUT_COMPARE_FOLDER"]

        # 3. Jalankan proses perbandingan (Cukup panggil SATU kali saja)
        full_path, file_name, sheet_list = compare_files(
            k3_sheets, coretax_sheets_1, coretax_sheets_2, output_dir
        )

        # 4. HAPUS BAGIAN INI: delete_all_uploaded_files(app)
        # Agar file di folder uploads dan outputs tidak hilang (menjadi database)
        delete_all_uploaded_files(app)

        # 5. Redirect dengan menyertakan mode='compare'
        return redirect(
            url_for(
                "show_comparison",
                updated_file=file_name,
                mode="compare",  # Menandai ini mode compare
                sheets=",".join(sheet_list),
                page=1,
            )
        )

    return "Invalid file type"


@app.route("/comparison", methods=["GET"])
def show_comparison():
    # 1. Ambil data dari URL
    filename = request.args.get("updated_file")
    mode = request.args.get("mode", "compare")
    sheets_raw = request.args.get("sheets", "")
    sheet_list = sheets_raw.split(",") if sheets_raw else []

    # 2. Tentukan folder berdasarkan mode (SUDAH BENAR)
    if mode == "recap":
        folder_path = app.config["OUTPUT_RECAP_FOLDER"]
    else:
        folder_path = app.config["OUTPUT_COMPARE_FOLDER"]

    # 3. Gabungkan folder dengan nama file (SUDAH BENAR)
    file_path = os.path.join(folder_path, filename)

    # --- MASALAH ADA DI SINI ---
    # file_path = os.path.join(BASE_DIR, 'outputs', filename)  <-- ### HAPUS BARIS INI ###
    # Baris di atas harus dihapus karena membuat sistem mencari di /outputs/ saja,
    # bukan di /outputs/recap/ atau /outputs/compare/

    # 4. Ambil parameter pendukung
    current_sheet = request.args.get(
        "sheet_name", sheet_list[0] if sheet_list else None
    )
    page = request.args.get("page", 1, type=int)
    rows_per_page = 50

    # 5. Cek keberadaan file
    if not os.path.exists(file_path):
        return (
            f"File tidak ditemukan di: {file_path}. Pastikan mode={mode} sudah benar.",
            404,
        )

    # --- LOGIKA MODE RECAP ---
    if mode == "recap":
        xls = pd.ExcelFile(file_path)
        all_tables = []
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name)
            html_content = df.to_html(
                classes="table table-hover table-bordered", index=False, na_rep=""
            )
            all_tables.append({"name": name, "content": html_content})

        return render_template(
            "comparison.html",
            all_tables=all_tables,
            mode=mode,
            filename=filename,
            updated_file=filename,
        )

    # --- LOGIKA MODE COMPARE (Pagination & Pilih Sheet) ---
    else:
        # Baca sheet yang dipilih saja
        merged_df = pd.read_excel(file_path, sheet_name=current_sheet)

        total_pages = (len(merged_df) // rows_per_page) + (
            1 if len(merged_df) % rows_per_page != 0 else 0
        )
        start_row = (page - 1) * rows_per_page
        end_row = start_row + rows_per_page
        page_data = merged_df[start_row:end_row]

        table_html = page_data.to_html(
            classes="table table-hover table-striped table-bordered",
            index=False,
            na_rep="",
        )

        # Build Pagination Controls
        base_url = f"/comparison?updated_file={filename}&sheets={sheets_raw}&mode=compare&sheet_name={current_sheet}"

        pagination_html = ""
        if page > 1:
            pagination_html += f'<a href="{base_url}&page={page-1}" class="btn btn-sm btn-secondary me-1">Previous</a>'

        for p in range(max(1, page - 2), min(total_pages, page + 2) + 1):
            if p == page:
                pagination_html += (
                    f' <span class="btn btn-sm btn-primary me-1">{p}</span>'
                )
            else:
                pagination_html += f' <a href="{base_url}&page={p}" class="btn btn-sm btn-outline-secondary me-1">{p}</a>'

        if page < total_pages:
            pagination_html += f' <a href="{base_url}&page={page+1}" class="btn btn-sm btn-secondary">Next</a>'

        return render_template(
            "comparison.html",
            table_html=table_html,
            pagination_html=pagination_html,
            filename=filename,
            sheets=sheet_list,
            current_sheet=current_sheet,
            updated_file=filename,
            mode=mode,
            page=page,
        )


@app.route("/download/<filename>")
def download_file(filename):
    # Ambil mode dari parameter URL (recap atau compare)
    mode = request.args.get("mode", "compare")

    # Tentukan folder berdasarkan mode
    if mode == "recap":
        folder_path = app.config["OUTPUT_RECAP_FOLDER"]
    else:
        folder_path = app.config["OUTPUT_COMPARE_FOLDER"]

    file_path = os.path.join(folder_path, filename)

    # Cek apakah file benar-benar ada sebelum dikirim
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return f"File tidak ditemukan di database {mode}: {file_path}", 404


if __name__ == "__main__":
    # Ubah 'True' menjadi 'False' jika ingin pindah ke mode production
    DEBUG_MODE = True

    if DEBUG_MODE:
        print("Running in DEBUG mode...")
        app.run(host="0.0.0.0", port=8000, debug=True)
    else:
        print("Running in PRODUCTION mode (Waitress)...")
        serve(app, host="0.0.0.0", port=8000, threads=6)
