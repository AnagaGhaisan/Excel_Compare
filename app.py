from flask import (
    Flask,
    request,
    render_template,
    send_file,
    redirect,
    url_for,
    send_from_directory,
    Response,
    jsonify,
    stream_with_context,
)
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
from functools import lru_cache
from helpers import compare_files, allowed_file, get_gl_account_options
from recap_handler import recap_bp
import pandas as pd
from waitress import serve
import uuid
import threading
import time
import json
from urllib.parse import urlencode
# IMPORT FUNGSI DARI FILE LAIN (Pastikan nama file dan fungsi sesuai)
from ekualisasi_handler import proses_ekualisasi


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
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 86400
app.config["CLEANUP_STALE_OUTPUT_FILES"] = None

# 2. Pastikan folder fisik dibuat di server
os.makedirs(app.config["OUTPUT_COMPARE_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_RECAP_FOLDER"], exist_ok=True)
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

COMPARE_JOBS = {}
COMPARE_JOBS_LOCK = threading.Lock()
COMPARE_JOB_TTL_SECONDS = 1800
UPLOAD_FILE_TTL_SECONDS = 7200
OUTPUT_FILE_TTL_SECONDS = 365 * 24 * 60 * 60
COMPARE_PREPARATION_PROGRESS_MAX = 20
COMPARE_PROCESS_PROGRESS_MAX = 99


def _delete_uploaded_files(file_paths):
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")


def _cleanup_stale_uploaded_files(upload_folder, max_age_seconds):
    now = time.time()

    try:
        for filename in os.listdir(upload_folder):
            file_path = os.path.join(upload_folder, filename)
            if not os.path.isfile(file_path):
                continue

            try:
                file_age_seconds = now - os.path.getmtime(file_path)
            except OSError:
                continue

            if file_age_seconds > max_age_seconds:
                try:
                    os.remove(file_path)
                    print(f"Removed stale upload file: {file_path}")
                except Exception as e:
                    print(f"Error deleting stale upload file {file_path}: {e}")
    except FileNotFoundError:
        return


def _cleanup_stale_output_files(output_folders, max_age_seconds):
    now = time.time()

    for output_folder in output_folders:
        try:
            for filename in os.listdir(output_folder):
                file_path = os.path.join(output_folder, filename)
                if not os.path.isfile(file_path):
                    continue

                try:
                    file_age_seconds = now - os.path.getmtime(file_path)
                except OSError:
                    continue

                if file_age_seconds > max_age_seconds:
                    try:
                        os.remove(file_path)
                        print(f"Removed stale output file: {file_path}")
                    except Exception as e:
                        print(f"Error deleting stale output file {file_path}: {e}")
        except FileNotFoundError:
            continue


def _cleanup_compare_jobs():
    now = time.time()
    with COMPARE_JOBS_LOCK:
        expired_job_ids = [
            job_id
            for job_id, job in COMPARE_JOBS.items()
            if job.get("status") in {"done", "error"}
            and (now - job.get("updated_at", now)) > COMPARE_JOB_TTL_SECONDS
        ]
        for job_id in expired_job_ids:
            COMPARE_JOBS.pop(job_id, None)


def _create_compare_job(job_id):
    with COMPARE_JOBS_LOCK:
        COMPARE_JOBS[job_id] = {
            "job_id": job_id,
            "status": "queued",
            "progress": 0,
            "message": "Job created.",
            "error": None,
            "redirect_url": None,
            "updated_at": time.time(),
        }


def _update_compare_job(
    job_id, progress=None, status=None, message=None, error=None, redirect_url=None
):
    with COMPARE_JOBS_LOCK:
        job = COMPARE_JOBS.get(job_id)
        if not job:
            return

        if progress is not None:
            bounded_progress = max(0, min(100, int(progress)))
            job["progress"] = bounded_progress
        if status is not None:
            job["status"] = status
        if message is not None:
            job["message"] = message
        if error is not None:
            job["error"] = error
        if redirect_url is not None:
            job["redirect_url"] = redirect_url
        job["updated_at"] = time.time()


def _get_compare_job(job_id):
    with COMPARE_JOBS_LOCK:
        job = COMPARE_JOBS.get(job_id)
        return dict(job) if job else None


def _build_comparison_redirect_url(file_name, sheet_list):
    query = urlencode(
        {
            "updated_file": file_name,
            "mode": "compare",
            "sheets": ",".join(sheet_list),
            "page": 1,
        }
    )
    return f"/comparison?{query}"


def _map_compare_stage_progress(compare_progress):
    bounded_progress = max(0, min(100, int(compare_progress)))
    progress_span = COMPARE_PROCESS_PROGRESS_MAX - COMPARE_PREPARATION_PROGRESS_MAX
    return COMPARE_PREPARATION_PROGRESS_MAX + int(
        (bounded_progress / 100) * progress_span
    )


def _parse_account_formula_map(raw_value):
    if not raw_value:
        return {}

    try:
        parsed = json.loads(raw_value)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid account formula configuration: {e}") from e

    if not isinstance(parsed, dict):
        raise ValueError("Account formula configuration must be a JSON object.")

    return parsed


def _process_comparison_job(
    job_id, k3_file_path, coretax_file_path_1, coretax_file_path_2, account_formula_map
):
    _update_compare_job(
        job_id,
        progress=5,
        status="processing",
        message="Reading GL file...",
    )

    try:
        k3_sheets = pd.read_excel(k3_file_path, sheet_name=None, header=1)
        _update_compare_job(
            job_id,
            progress=10,
            status="processing",
            message="Reading Coretax Digunggung file...",
        )

        coretax_sheets_1 = pd.read_excel(coretax_file_path_1, sheet_name=None, header=1)
        _update_compare_job(
            job_id,
            progress=15,
            status="processing",
            message="Reading Coretax Tidak Digunggung file...",
        )

        coretax_sheets_2 = pd.read_excel(coretax_file_path_2, sheet_name=None, header=1)
        _update_compare_job(
            job_id,
            progress=COMPARE_PREPARATION_PROGRESS_MAX,
            status="processing",
            message="Comparing and matching data...",
        )

        def _on_compare_progress(compare_progress, compare_message):
            _update_compare_job(
                job_id,
                progress=_map_compare_stage_progress(compare_progress),
                status="processing",
                message=compare_message,
            )

        _, file_name, sheet_list = compare_files(
            k3_sheets,
            coretax_sheets_1,
            coretax_sheets_2,
            app.config["OUTPUT_COMPARE_FOLDER"],
            progress_callback=_on_compare_progress,
            account_formula_map=account_formula_map,
        )

        redirect_url = _build_comparison_redirect_url(file_name, sheet_list)
        _update_compare_job(
            job_id,
            progress=100,
            status="done",
            message="Comparison completed successfully.",
            redirect_url=redirect_url,
        )
    except Exception as e:
        print(f"Error in comparison job {job_id}: {e}")
        _update_compare_job(
            job_id,
            progress=100,
            status="error",
            message="Comparison failed.",
            error=str(e),
        )
    finally:
        _delete_uploaded_files([k3_file_path, coretax_file_path_1, coretax_file_path_2])


@lru_cache(maxsize=64)
def _read_excel_sheet_cached(file_path, sheet_name, modified_time):
    return pd.read_excel(file_path, sheet_name=sheet_name)


@lru_cache(maxsize=16)
def _read_excel_file_cached(file_path, modified_time):
    return pd.ExcelFile(file_path)


def _get_file_modified_time(file_path):
    try:
        return os.path.getmtime(file_path)
    except OSError:
        return 0.0


@lru_cache(maxsize=256)
def _static_asset_version(relative_path):
    file_path = os.path.join(app.static_folder, relative_path)
    return str(int(os.path.getmtime(file_path)))


@app.context_processor
def inject_asset_helpers():
    def asset_url(filename):
        try:
            version = _static_asset_version(filename)
        except OSError:
            version = None
        return url_for("static", filename=filename, v=version)

    return {"asset_url": asset_url}


app.config["UPLOAD_FILE_TTL_SECONDS"] = UPLOAD_FILE_TTL_SECONDS
app.config["OUTPUT_FILE_TTL_SECONDS"] = OUTPUT_FILE_TTL_SECONDS
app.config["CLEANUP_STALE_OUTPUT_FILES"] = _cleanup_stale_output_files
_cleanup_stale_uploaded_files(
    app.config["UPLOAD_FOLDER"], app.config["UPLOAD_FILE_TTL_SECONDS"]
)
_cleanup_stale_output_files(
    [
        app.config["OUTPUT_COMPARE_FOLDER"],
        app.config["OUTPUT_RECAP_FOLDER"],
    ],
    app.config["OUTPUT_FILE_TTL_SECONDS"],
)


@app.after_request
def set_static_cache_headers(response):
    if request.path.startswith("/static/"):
        response.cache_control.public = True
        response.cache_control.max_age = 31536000
        response.cache_control.immutable = True
    return response

# Homepage route to upload files
@app.route("/")
def home():
    return render_template("index.html")  # Landing page with two card options


@app.route("/filecompare")
def file_compare():
    return render_template("PPN/FileCompare/index.html")


@app.route("/filerecap")
def file_recap():
    return render_template("PPN/FileRecap/index.html")

@app.route('/download-template/<template_type>')
def download_template(template_type):
    # Mapping parameter ke nama file asli di folder static/excel_template
    templates = {
        "gl": "Gl Template.xlsx",
        "digunggung": "Digungung Template.xlsx",
        "tidak-digunggung": "Tidak Digungung Template.xlsx",
        "ppn-recap": "Rekap_PPN.xlsx"
    }
    
    filename = templates.get(template_type)
    if not filename:
        return "Template tidak ditemukan", 404
        
    # Pastikan folder name sesuai: static/excel_template
    template_dir = os.path.join(app.root_path, 'static', 'excel_template')
    
    # Memastikan file benar-benar ada sebelum dikirim
    if not os.path.exists(os.path.join(template_dir, filename)):
        return f"File {filename} tidak ditemukan di folder static/excel_template", 404
    
    return send_from_directory(template_dir, filename, as_attachment=True)


@app.route("/upload/start", methods=["POST"])
def start_upload_file():
    _cleanup_compare_jobs()
    _cleanup_stale_uploaded_files(
        app.config["UPLOAD_FOLDER"], app.config["UPLOAD_FILE_TTL_SECONDS"]
    )
    _cleanup_stale_output_files(
        [
            app.config["OUTPUT_COMPARE_FOLDER"],
            app.config["OUTPUT_RECAP_FOLDER"],
        ],
        app.config["OUTPUT_FILE_TTL_SECONDS"],
    )

    if (
        "k3_file" not in request.files
        or "coretax_file_1" not in request.files
        or "coretax_file_2" not in request.files
    ):
        return jsonify({"error": "No file part"}), 400

    k3_file = request.files["k3_file"]
    coretax_file_1 = request.files["coretax_file_1"]
    coretax_file_2 = request.files["coretax_file_2"]
    try:
        account_formula_map = _parse_account_formula_map(
            request.form.get("account_formulas")
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    if (
        k3_file.filename == ""
        or coretax_file_1.filename == ""
        or coretax_file_2.filename == ""
    ):
        return jsonify({"error": "No selected file"}), 400

    if not (
        k3_file
        and allowed_file(k3_file.filename, app.config["ALLOWED_EXTENSIONS"])
        and coretax_file_1
        and allowed_file(coretax_file_1.filename, app.config["ALLOWED_EXTENSIONS"])
        and coretax_file_2
        and allowed_file(coretax_file_2.filename, app.config["ALLOWED_EXTENSIONS"])
    ):
        return jsonify({"error": "Invalid file type"}), 400

    unique_id = str(uuid.uuid4())[:8]
    job_id = str(uuid.uuid4())

    k3_filename = f"{unique_id}_{secure_filename(k3_file.filename)}"
    coretax_filename_1 = f"{unique_id}_{secure_filename(coretax_file_1.filename)}"
    coretax_filename_2 = f"{unique_id}_{secure_filename(coretax_file_2.filename)}"

    k3_file_path = os.path.join(app.config["UPLOAD_FOLDER"], k3_filename)
    coretax_file_path_1 = os.path.join(app.config["UPLOAD_FOLDER"], coretax_filename_1)
    coretax_file_path_2 = os.path.join(app.config["UPLOAD_FOLDER"], coretax_filename_2)

    try:
        k3_file.save(k3_file_path)
        coretax_file_1.save(coretax_file_path_1)
        coretax_file_2.save(coretax_file_path_2)
    except Exception as e:
        _delete_uploaded_files([k3_file_path, coretax_file_path_1, coretax_file_path_2])
        return jsonify({"error": f"Failed to save uploaded files: {str(e)}"}), 500

    _create_compare_job(job_id)
    _update_compare_job(
        job_id,
        progress=3,
        status="processing",
        message="Files uploaded. Starting comparison...",
    )

    worker = threading.Thread(
        target=_process_comparison_job,
        args=(
            job_id,
            k3_file_path,
            coretax_file_path_1,
            coretax_file_path_2,
            account_formula_map,
        ),
        daemon=True,
    )
    worker.start()

    return jsonify({"job_id": job_id}), 202


@app.route("/upload/progress/<job_id>", methods=["GET"])
def stream_upload_progress(job_id):
    def event_stream():
        last_payload = None

        while True:
            job = _get_compare_job(job_id)
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

        _cleanup_compare_jobs()

    response = Response(
        stream_with_context(event_stream()),
        mimetype="text/event-stream",
    )
    response.headers["Cache-Control"] = "no-cache"
    response.headers["X-Accel-Buffering"] = "no"
    return response


@app.route("/upload/accounts", methods=["POST"])
def upload_account_options():
    if "k3_file" not in request.files:
        return jsonify({"error": "GL file is required."}), 400

    k3_file = request.files["k3_file"]
    if k3_file.filename == "":
        return jsonify({"error": "GL file is required."}), 400

    if not allowed_file(k3_file.filename, app.config["ALLOWED_EXTENSIONS"]):
        return jsonify({"error": "Invalid GL file type."}), 400

    try:
        k3_sheets = pd.read_excel(k3_file, sheet_name=None, header=1)
        accounts = get_gl_account_options(k3_sheets)
        return jsonify({"accounts": accounts}), 200
    except Exception as e:
        return jsonify({"error": f"Failed to read GL accounts: {str(e)}"}), 400


@app.route("/upload", methods=["POST"])
def upload_file():
    _cleanup_stale_output_files(
        [
            app.config["OUTPUT_COMPARE_FOLDER"],
            app.config["OUTPUT_RECAP_FOLDER"],
        ],
        app.config["OUTPUT_FILE_TTL_SECONDS"],
    )

    if (
        "k3_file" not in request.files
        or "coretax_file_1" not in request.files
        or "coretax_file_2" not in request.files
    ):
        return "No file part"

    k3_file = request.files["k3_file"]
    coretax_file_1 = request.files["coretax_file_1"]
    coretax_file_2 = request.files["coretax_file_2"]
    try:
        account_formula_map = _parse_account_formula_map(
            request.form.get("account_formulas")
        )
    except ValueError as e:
        return str(e), 400

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
            k3_sheets,
            coretax_sheets_1,
            coretax_sheets_2,
            output_dir,
            account_formula_map=account_formula_map,
        )

        _delete_uploaded_files([k3_file_path, coretax_file_path_1, coretax_file_path_2])

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
        modified_time = _get_file_modified_time(file_path)
        xls = _read_excel_file_cached(file_path, modified_time)
        all_tables = []
        for name in xls.sheet_names:
            df = _read_excel_sheet_cached(file_path, name, modified_time)
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
        modified_time = _get_file_modified_time(file_path)
        merged_df = _read_excel_sheet_cached(file_path, current_sheet, modified_time)

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
    
 
# --- ROUTE FLASK PPH 23 ---
@app.route('/ekualisasi-pph23', methods=['GET', 'POST'])
def ekualisasi_pph23_route():
    if request.method == 'POST':
        _cleanup_stale_output_files(
            [
                app.config["OUTPUT_COMPARE_FOLDER"],
                app.config["OUTPUT_RECAP_FOLDER"],
            ],
            app.config["OUTPUT_FILE_TTL_SECONDS"],
        )

        # 1. Cek ketersediaan file
        if 'file_bupot' not in request.files or 'file_voucher' not in request.files:
            return "No file part", 400
            
        file_bupot = request.files['file_bupot']
        file_voucher = request.files['file_voucher']
        
        # 2. Cek apakah file kosong
        if file_bupot.filename == '' or file_voucher.filename == '':
            return "No selected file", 400
            
        # 3. Validasi ekstensi dengan fungsi allowed_file dari helpers.py
        if (
            file_bupot and allowed_file(file_bupot.filename, app.config["ALLOWED_EXTENSIONS"]) and
            file_voucher and allowed_file(file_voucher.filename, app.config["ALLOWED_EXTENSIONS"])
        ):
            # 4. Buat Unique ID seperti PPN
            unique_id = str(uuid.uuid4())[:8]
            
            bupot_filename = f"{unique_id}_{secure_filename(file_bupot.filename)}"
            voucher_filename = f"{unique_id}_{secure_filename(file_voucher.filename)}"
            
            bupot_path = os.path.join(app.config["UPLOAD_FOLDER"], bupot_filename)
            voucher_path = os.path.join(app.config["UPLOAD_FOLDER"], voucher_filename)
            
            # 5. Path output & template
            output_filename = f"Hasil_Ekualisasi_PPH23_{unique_id}.xlsx"
            output_path = os.path.join(app.config["OUTPUT_COMPARE_FOLDER"], output_filename)
            template_path = os.path.join(BASE_DIR, "static", "template", "Format Output.xlsx")
            
            # 6. Simpan file fisik ke server
            file_bupot.save(bupot_path)
            file_voucher.save(voucher_path)
            
            try:
                # 7. Jalankan pemrosesan
                proses_ekualisasi(bupot_path, voucher_path, template_path, output_path)
                
                _delete_uploaded_files([bupot_path, voucher_path])
                
                # 9. Kembalikan file hasil
                return send_file(output_path, as_attachment=True, download_name='Hasil_Ekualisasi_PPH23.xlsx')
            
            except Exception as e:
                print(f"Error Ekualisasi PPH23: {e}")
                return f"Terjadi kesalahan saat memproses data: {str(e)}", 500
                
        return "Invalid file type", 400
                
    # Tampilkan antarmuka jika method GET
    return render_template('PPH/index.html')

if __name__ == "__main__":
    # Ubah 'True' menjadi 'False' jika ingin pindah ke mode production
    DEBUG_MODE = True

    if DEBUG_MODE:
        print("Running in DEBUG mode...")
        app.run(host="0.0.0.0", port=8000, debug=True)
    else:
        print("Running in PRODUCTION mode (Waitress)...")
        serve(app, host="0.0.0.0", port=8000, threads=6)
