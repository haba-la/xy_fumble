from __future__ import annotations

import argparse
import uuid
from pathlib import Path
from typing import Any, Dict

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from docx_html_layout_sync.layout_sync import sync_docx_layout_to_html
from docx_to_html_temp.docx_to_html import convert_docx_to_temp_html
from html_consistency_checker.check_html_consistency import run_html_consistency_check
from html_to_docx_sync.html_to_docx import html_to_docx

BASE_DIR = Path(__file__).resolve().parent
RUNTIME_DIR = BASE_DIR / "runtime" / "portal_jobs"
JOBS_DIR = RUNTIME_DIR / "jobs"
ALLOWED_EXTENSIONS = {".docx"}

JOBS_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "web_portal" / "templates"),
    static_folder=str(BASE_DIR / "web_portal" / "static"),
)


def _ensure_docx(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _artifact_paths(job_dir: Path) -> Dict[str, Path]:
    return {
        "source_docx": job_dir / "source_upload.docx",
        "template_docx": job_dir / "template_upload.docx",
        "temp_html": job_dir / "source_temp.html",
        "synced_html": job_dir / "synced_layout.html",
        "profile_json": job_dir / "layout_profile.json",
        "synced_docx": job_dir / "synced_layout.docx",
        "consistency_json": job_dir / "consistency_result.json",
    }


def _save_upload(file_storage, target_path: Path) -> None:
    target_path.parent.mkdir(parents=True, exist_ok=True)
    file_storage.save(target_path)


@app.get("/")
def index():
    return render_template("layout_portal.html")


@app.post("/api/process")
def process_files():
    source_file = request.files.get("source_docx")
    template_file = request.files.get("template_docx")

    if source_file is None or not source_file.filename:
        return jsonify({"ok": False, "error": "请上传待修改文档 DOCX。"}), 400
    if template_file is None or not template_file.filename:
        return jsonify({"ok": False, "error": "请上传模板文档 DOCX。"}), 400
    if not _ensure_docx(source_file.filename) or not _ensure_docx(template_file.filename):
        return jsonify({"ok": False, "error": "仅支持 .docx 文件。"}), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    paths = _artifact_paths(job_dir)

    source_name = secure_filename(source_file.filename) or "source.docx"
    template_name = secure_filename(template_file.filename) or "template.docx"

    try:
        _save_upload(source_file, paths["source_docx"])
        _save_upload(template_file, paths["template_docx"])

        convert_docx_to_temp_html(
            docx_path=paths["source_docx"],
            output_html_path=paths["temp_html"],
        )

        sync_docx_layout_to_html(
            docx_path=str(paths["template_docx"]),
            html_path=str(paths["temp_html"]),
            output_html_path=str(paths["synced_html"]),
            profile_output_path=str(paths["profile_json"]),
        )

        html_to_docx(
            paths["synced_html"],
            paths["synced_docx"],
            template_docx_path=paths["template_docx"],
        )

        consistency_error = ""
        consistency_result: Dict[str, Any] = {
            "has_inconsistency": None,
            "summary": "一致性检查未执行",
            "issues": [],
        }

        try:
            consistency_result = run_html_consistency_check(
                html_path=paths["synced_html"],
                output_path=paths["consistency_json"],
            )
        except Exception as exc:
            consistency_error = str(exc)

    except Exception as exc:
        return jsonify({"ok": False, "error": f"处理失败: {exc}"}), 500

    return jsonify(
        {
            "ok": True,
            "job_id": job_id,
            "files": {
                "source": source_name,
                "template": template_name,
            },
            "preview_url": f"/preview/{job_id}",
            "downloads": {
                "profile_json": f"/download/{job_id}/profile_json",
                "synced_html": f"/download/{job_id}/synced_html",
                "synced_docx": f"/download/{job_id}/synced_docx",
            },
            "consistency": consistency_result,
            "consistency_error": consistency_error,
        }
    )


@app.get("/preview/<job_id>")
def preview_html(job_id: str):
    target = _artifact_paths(JOBS_DIR / job_id)["synced_html"]
    if not target.exists():
        return "预览文件不存在", 404
    return send_file(target, mimetype="text/html; charset=utf-8")


@app.get("/download/<job_id>/<artifact>")
def download_file(job_id: str, artifact: str):
    mapping = {
        "profile_json": ("profile_json", "application/json"),
        "synced_html": ("synced_html", "text/html; charset=utf-8"),
        "synced_docx": (
            "synced_docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ),
    }

    if artifact not in mapping:
        return "不支持的导出类型", 404

    key, mime = mapping[artifact]
    target = _artifact_paths(JOBS_DIR / job_id)[key]
    if not target.exists():
        return "文件不存在", 404

    return send_file(target, as_attachment=True, mimetype=mime, download_name=target.name)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run document layout portal")
    parser.add_argument("--host", default="127.0.0.1", help="Server host")
    parser.add_argument("--port", type=int, default=8090, help="Server port")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
