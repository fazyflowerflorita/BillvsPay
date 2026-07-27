from __future__ import annotations

from dataclasses import dataclass
from email.parser import BytesParser
from email.policy import default
from io import BytesIO
import json
import mimetypes
import os
import shutil
import sys
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from reconcile import create_default_rules_workbook, reconcile


ROOT = Path(__file__).resolve().parents[1]
STATIC_DIR = Path(__file__).resolve().parent / "static"
OUTPUT_DIR = ROOT / "output"
CONFIG_DIR = ROOT / "config"
TMP_DIR = ROOT / "tmp"
UPLOAD_DIR = ROOT / "uploads"
RULES_PATH = CONFIG_DIR / "business_rules.xlsx"


@dataclass
class UploadedFile:
    filename: str
    file: BytesIO


def parse_multipart_form(headers, body: bytes) -> dict[str, UploadedFile]:
    content_type = headers.get("Content-Type", "")
    message_bytes = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8") + body
    message = BytesParser(policy=default).parsebytes(message_bytes)
    form: dict[str, UploadedFile] = {}

    if not message.is_multipart():
        return form

    for part in message.iter_parts():
        disposition = part.get_content_disposition()
        if disposition != "form-data":
            continue
        field_name = part.get_param("name", header="content-disposition")
        filename = part.get_filename() or ""
        if not field_name:
            continue
        payload = part.get_payload(decode=True) or b""
        form[field_name] = UploadedFile(filename=filename, file=BytesIO(payload))

    return form


class ReconciliationHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.serve_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/static/"):
            file_path = STATIC_DIR / unquote(parsed.path.removeprefix("/static/"))
            self.serve_file(file_path)
            return
        if parsed.path.startswith("/download/"):
            file_name = Path(unquote(parsed.path.removeprefix("/download/"))).name
            file_path = OUTPUT_DIR / file_name
            self.serve_file(file_path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True)
            return
        self.send_error(404, "Not found")

    def do_POST(self) -> None:
        if self.path != "/reconcile":
            self.send_error(404, "Not found")
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        form = parse_multipart_form(self.headers, self.rfile.read(content_length))
        if "payroll" not in form or "billing" not in form:
            self.send_json({"error": "Upload both Payroll and Billing Excel reports."}, status=400)
            return

        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        upload_id = uuid.uuid4().hex
        payroll_path = UPLOAD_DIR / f"{upload_id}_payroll.xlsx"
        billing_path = UPLOAD_DIR / f"{upload_id}_billing.xlsx"
        try:
            self.save_upload(form["payroll"], payroll_path)
            self.save_upload(form["billing"], billing_path)
            result = reconcile(payroll_path, billing_path, RULES_PATH, OUTPUT_DIR)
        except Exception as exc:
            self.send_json({"error": str(exc)}, status=500)
            return

        output_path = Path(result["output_path"])
        self.send_json(
            {
                "file": output_path.name,
                "downloadUrl": f"/download/{output_path.name}",
                "summary": result["summary"],
                "exceptionSummary": result["exception_summary"],
                "reconciledCount": result["reconciled_count"],
                "exceptionCount": result["exception_count"],
            }
        )

    def save_upload(self, field: UploadedFile, target: Path) -> None:
        with target.open("wb") as handle:
            shutil.copyfileobj(field.file, handle)

    def serve_file(self, file_path: Path, content_type: str | None = None, as_attachment: bool = False) -> None:
        try:
            resolved = file_path.resolve()
            allowed_roots = [STATIC_DIR.resolve(), OUTPUT_DIR.resolve()]
            if not any(resolved == root or root in resolved.parents for root in allowed_roots):
                self.send_error(403, "Forbidden")
                return
            if not resolved.exists() or not resolved.is_file():
                self.send_error(404, "Not found")
                return
            content_type = content_type or mimetypes.guess_type(resolved.name)[0] or "application/octet-stream"
            data = resolved.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            if as_attachment:
                self.send_header("Content-Disposition", f'attachment; filename="{resolved.name}"')
            self.end_headers()
            self.wfile.write(data)
        except OSError as exc:
            self.send_error(500, str(exc))

    def send_json(self, payload: dict, status: int = 200) -> None:
        data = json.dumps(payload, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> int:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    if not RULES_PATH.exists():
        create_default_rules_workbook(RULES_PATH)

    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT") or (sys.argv[1] if len(sys.argv) > 1 else 8765))
    server = ThreadingHTTPServer((host, port), ReconciliationHandler)
    display_host = "127.0.0.1" if host in {"0.0.0.0", ""} else host
    print(f"Bill vs Pay Reconciliation running at http://{display_host}:{port}")
    print("Press Ctrl+C to stop.")
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
