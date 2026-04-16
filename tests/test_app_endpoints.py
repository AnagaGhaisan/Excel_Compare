import io
import json
import os
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

from app import app
from app import COMPARE_JOBS, COMPARE_JOBS_LOCK
from recap_handler import RECAP_JOBS, RECAP_JOBS_LOCK


class FlaskEndpointTests(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        app.config["TESTING"] = True
        app.config["WTF_CSRF_ENABLED"] = False
        app.config["UPLOAD_FOLDER"] = os.path.join(self.tmpdir.name, "uploads")
        app.config["OUTPUT_COMPARE_FOLDER"] = os.path.join(self.tmpdir.name, "outputs", "compare")
        app.config["OUTPUT_RECAP_FOLDER"] = os.path.join(self.tmpdir.name, "outputs", "recap")
        os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
        os.makedirs(app.config["OUTPUT_COMPARE_FOLDER"], exist_ok=True)
        os.makedirs(app.config["OUTPUT_RECAP_FOLDER"], exist_ok=True)
        self.client = app.test_client()
        with COMPARE_JOBS_LOCK:
            COMPARE_JOBS.clear()
        with RECAP_JOBS_LOCK:
            RECAP_JOBS.clear()

    def tearDown(self):
        self.tmpdir.cleanup()

    @patch("app.render_template", return_value="HOME_OK")
    def test_home(self, _):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"HOME_OK")

    @patch("app.render_template", return_value="COMPARE_PAGE")
    def test_file_compare_page(self, _):
        response = self.client.get("/filecompare")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"COMPARE_PAGE")

    @patch("app.render_template", return_value="RECAP_PAGE")
    def test_file_recap_page(self, _):
        response = self.client.get("/filerecap")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"RECAP_PAGE")

    @patch("app.send_from_directory")
    @patch("app.os.path.exists", return_value=True)
    def test_download_template_success(self, _, mock_send):
        mock_send.return_value = app.response_class("FILE_OK", status=200)
        response = self.client.get("/download-template/gl")
        self.assertEqual(response.status_code, 200)
        mock_send.assert_called_once()

    def test_download_template_invalid_type(self):
        response = self.client.get("/download-template/not-exist")
        self.assertEqual(response.status_code, 404)
        self.assertIn(b"Template tidak ditemukan", response.data)

    @patch("app.os.path.exists", return_value=False)
    def test_download_template_missing_file(self, _):
        response = self.client.get("/download-template/gl")
        self.assertEqual(response.status_code, 404)
        self.assertIn(b"tidak ditemukan", response.data)

    def test_upload_missing_file_part(self):
        response = self.client.post("/upload", data={}, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"No file part", response.data)

    @patch("app._delete_uploaded_files")
    @patch("app.compare_files")
    @patch("app.pd.read_excel")
    @patch("app.uuid.uuid4")
    def test_upload_success_redirects_to_comparison(
        self, mock_uuid, mock_read_excel, mock_compare_files, _
    ):
        mock_uuid.return_value = "12345678-aaaa-bbbb-cccc-ddddeeeeffff"
        mock_read_excel.return_value = {"Sheet1": pd.DataFrame([{"x": 1}])}
        mock_compare_files.return_value = (
            "/tmp/out.xlsx",
            "Draft_Updated_Output.xlsx",
            ["Sheet1", "Sheet2"],
        )

        data = {
            "k3_file": (io.BytesIO(b"dummy"), "k3.xlsx"),
            "coretax_file_1": (io.BytesIO(b"dummy"), "core1.xlsx"),
            "coretax_file_2": (io.BytesIO(b"dummy"), "core2.xlsx"),
        }
        response = self.client.post("/upload", data=data, content_type="multipart/form-data")

        self.assertEqual(response.status_code, 302)
        self.assertIn("/comparison?", response.location)
        self.assertIn("updated_file=Draft_Updated_Output.xlsx", response.location)
        self.assertIn("mode=compare", response.location)
        self.assertIn("sheets=Sheet1,Sheet2", response.location)

    @patch("app.threading.Thread")
    @patch("app.uuid.uuid4")
    def test_upload_start_returns_job_id_and_accepted(self, mock_uuid, mock_thread):
        mock_uuid.side_effect = [
            "11223344-aaaa-bbbb-cccc-ddddeeeeffff",
            "99887766-aaaa-bbbb-cccc-ddddeeeeffff",
        ]
        mock_worker = MagicMock()
        mock_thread.return_value = mock_worker

        data = {
            "k3_file": (io.BytesIO(b"dummy"), "k3.xlsx"),
            "coretax_file_1": (io.BytesIO(b"dummy"), "core1.xlsx"),
            "coretax_file_2": (io.BytesIO(b"dummy"), "core2.xlsx"),
        }
        response = self.client.post(
            "/upload/start", data=data, content_type="multipart/form-data"
        )

        self.assertEqual(response.status_code, 202)
        payload = response.get_json()
        self.assertIn("job_id", payload)
        self.assertEqual(payload["job_id"], "99887766-aaaa-bbbb-cccc-ddddeeeeffff")
        mock_thread.assert_called_once()
        mock_worker.start.assert_called_once()

    def test_upload_progress_missing_job_returns_error_event(self):
        response = self.client.get("/upload/progress/non-existent-job")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/event-stream")

        body = response.get_data(as_text=True)
        data_lines = [
            line.replace("data: ", "", 1)
            for line in body.splitlines()
            if line.startswith("data: ")
        ]
        self.assertGreaterEqual(len(data_lines), 1)

        payload = json.loads(data_lines[0])
        self.assertEqual(payload["job_id"], "non-existent-job")
        self.assertEqual(payload["status"], "error")
        self.assertEqual(payload["progress"], 100)
        self.assertIn("Job not found", payload["error"])

    @patch("app.render_template", return_value="COMPARISON_OK")
    @patch("app.pd.read_excel")
    @patch("app.os.path.exists", return_value=True)
    def test_comparison_compare_mode(self, _, mock_read_excel, mock_render):
        mock_read_excel.return_value = pd.DataFrame({"A": [1, 2, 3]})
        response = self.client.get(
            "/comparison?updated_file=test.xlsx&mode=compare&sheets=Sheet1&sheet_name=Sheet1&page=1"
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"COMPARISON_OK")
        mock_render.assert_called_once()

    @patch("app.render_template", return_value="RECAP_COMPARISON_OK")
    @patch("app.pd.read_excel")
    @patch("app.pd.ExcelFile")
    @patch("app.os.path.exists", return_value=True)
    def test_comparison_recap_mode(self, _, mock_excel_file, mock_read_excel, mock_render):
        mock_excel = MagicMock()
        mock_excel.sheet_names = ["Summary1", "Summary2"]
        mock_excel_file.return_value = mock_excel
        mock_read_excel.return_value = pd.DataFrame({"A": [1]})

        response = self.client.get("/comparison?updated_file=recap.xlsx&mode=recap")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"RECAP_COMPARISON_OK")
        mock_render.assert_called_once()

    @patch("app.send_file")
    @patch("app.os.path.exists", return_value=True)
    def test_download_compare_success(self, _, mock_send):
        mock_send.return_value = app.response_class("DOWNLOAD_OK", status=200)
        response = self.client.get("/download/result.xlsx?mode=compare")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"DOWNLOAD_OK")

    @patch("app.os.path.exists", return_value=False)
    def test_download_missing_file(self, _):
        response = self.client.get("/download/notfound.xlsx?mode=recap")
        self.assertEqual(response.status_code, 404)
        self.assertIn(b"File tidak ditemukan di database recap", response.data)

    @patch("app.render_template", return_value="PPH_PAGE")
    def test_ekualisasi_pph23_get(self, _):
        response = self.client.get("/ekualisasi-pph23")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"PPH_PAGE")

    @patch("app.send_file")
    @patch("app._delete_uploaded_files")
    @patch("app.proses_ekualisasi")
    @patch("app.uuid.uuid4")
    def test_ekualisasi_pph23_post_success(self, mock_uuid, _, __, mock_send):
        mock_uuid.return_value = "abcdef12-aaaa-bbbb-cccc-ddddeeeeffff"
        mock_send.return_value = app.response_class("EKUALISASI_OK", status=200)

        data = {
            "file_bupot": (io.BytesIO(b"bupot"), "bupot.xlsx"),
            "file_voucher": (io.BytesIO(b"voucher"), "voucher.xlsx"),
        }
        response = self.client.post(
            "/ekualisasi-pph23", data=data, content_type="multipart/form-data"
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"EKUALISASI_OK")

    def test_upload_recap_missing_file(self):
        response = self.client.post("/upload_recap", data={}, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 400)
        self.assertIn(b"wajib diupload", response.data)

    @patch("recap_handler.threading.Thread")
    @patch("recap_handler.uuid.uuid4")
    def test_upload_recap_start_returns_job_id(self, mock_uuid, mock_thread):
        mock_uuid.side_effect = [
            "11223344-aaaa-bbbb-cccc-ddddeeeeffff",
            "88776655-aaaa-bbbb-cccc-ddddeeeeffff",
        ]
        mock_worker = MagicMock()
        mock_thread.return_value = mock_worker

        data = {
            "k3_file": (io.BytesIO(b"gl"), "draft_gl.xlsx"),
            "ppn_file": (io.BytesIO(b"ppn"), "rekap_ppn.xlsx"),
        }
        response = self.client.post(
            "/upload_recap/start", data=data, content_type="multipart/form-data"
        )

        self.assertEqual(response.status_code, 202)
        payload = response.get_json()
        self.assertIn("job_id", payload)
        self.assertEqual(payload["job_id"], "88776655-aaaa-bbbb-cccc-ddddeeeeffff")
        mock_thread.assert_called_once()
        mock_worker.start.assert_called_once()

    def test_upload_recap_progress_missing_job_returns_error_event(self):
        response = self.client.get("/upload_recap/progress/non-existent-job")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "text/event-stream")

        body = response.get_data(as_text=True)
        data_lines = [
            line.replace("data: ", "", 1)
            for line in body.splitlines()
            if line.startswith("data: ")
        ]
        self.assertGreaterEqual(len(data_lines), 1)

        payload = json.loads(data_lines[0])
        self.assertEqual(payload["job_id"], "non-existent-job")
        self.assertEqual(payload["status"], "error")
        self.assertEqual(payload["progress"], 100)
        self.assertIn("Job not found", payload["error"])

    @patch("recap_handler.process_recap_2_files")
    @patch("recap_handler.uuid.uuid4")
    def test_upload_recap_success_redirect(self, mock_uuid, mock_process):
        mock_uuid.return_value = "11223344-aaaa-bbbb-cccc-ddddeeeeffff"
        mock_process.return_value = ("Final_Ekualisasi_test.xlsx", ["Summary Ekualisasi"])

        data = {
            "k3_file": (io.BytesIO(b"gl"), "draft_gl.xlsx"),
            "ppn_file": (io.BytesIO(b"ppn"), "rekap_ppn.xlsx"),
        }
        response = self.client.post(
            "/upload_recap", data=data, content_type="multipart/form-data"
        )

        self.assertEqual(response.status_code, 302)
        self.assertIn("/comparison?", response.location)
        self.assertIn("updated_file=Final_Ekualisasi_test.xlsx", response.location)
        self.assertIn("mode=recap", response.location)


if __name__ == "__main__":
    unittest.main()
