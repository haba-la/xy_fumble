from __future__ import annotations

import argparse
import uuid
from pathlib import Path

from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

from docx_to_html import convert_docx_to_temp_html

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "runtime" / "uploads"
HTML_TEMP_DIR = BASE_DIR / "runtime" / "html_temp"
ALLOWED_EXTENSIONS = {".docx"}

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
HTML_TEMP_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


@app.get("/")
def index():
    return render_template("upload.html")


@app.post("/upload")
def upload_docx():
    file_storage = request.files.get("docx_file")
    if file_storage is None or not file_storage.filename:
        return render_template("upload.html", error="Please select a DOCX file."), 400

    raw_filename = file_storage.filename.strip()
    ext = Path(raw_filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return render_template("upload.html", error="Only .docx files are supported."), 400

    safe_filename = secure_filename(raw_filename)
    base_stem = Path(safe_filename).stem if safe_filename else "uploaded_docx"

    source_name = f"{base_stem}_{uuid.uuid4().hex}{ext}"
    source_path = UPLOAD_DIR / source_name
    file_storage.save(source_path)

    target_name = f"{Path(source_name).stem}.html"
    target_path = HTML_TEMP_DIR / target_name

    output_path = convert_docx_to_temp_html(
        docx_path=source_path,
        output_html_path=target_path,
    )

    return render_template(
        "upload.html",
        success="DOCX converted successfully. HTML temp file is ready for editing.",
        html_file=output_path.name,
        html_path=str(output_path),
    )


@app.get("/html/<path:filename>")
def get_html(filename: str):
    return send_from_directory(HTML_TEMP_DIR, filename)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run DOCX to HTML upload web app")
    parser.add_argument("--host", default="127.0.0.1", help="Server host")
    parser.add_argument("--port", type=int, default=8060, help="Server port")
    parser.add_argument("--debug", action="store_true", help="Enable Flask debug mode")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
