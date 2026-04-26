import os
import re
import json
import io
import time
from pathlib import Path

import requests
from flask import Flask, jsonify, request, send_from_directory, send_file

from llm_template_generator import generate_template_from_text, save_template_to_file
from docx_formatter import generate_formatting_instructions, format_docx_bytes


# ===== 配置区（建议用环境变量）=====
# 注意：os.getenv 的第一个参数必须是“环境变量名”，不是 API Key 本身。
CODE_API_KEY = "1c3e83b7-e80f-457b-9d9f-c313c8a3070e"
API_KEY = (os.getenv("ARK_API_KEY") or CODE_API_KEY).strip()
MODEL = os.getenv("ARK_MODEL", "doubao-seed-2-0-lite-260215")
ARK_URL = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
LLM_TIMEOUT_SECONDS = 300
LLM_RETRY_TIMES = 3
# ================================

BASE_DIR = Path(__file__).resolve().parent
PORTAL_DIR = BASE_DIR / "app_portal"
CHAT_UI_FILE = "chat_ui.html"
FIXED_ENTRY_FILE = "studio.html"
TEMPLATES_DIR = Path("/root/文档规则/templates")

app = Flask(__name__)


def _safe_json_load(file_path: Path):
    try:
        return json.loads(file_path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _list_templates_meta():
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    items = []
    for p in sorted(TEMPLATES_DIR.glob("*.json")):
        data = _safe_json_load(p)
        if not isinstance(data, dict):
            continue
        items.append(
            {
                "file_name": p.name,
                "name": data.get("name", p.stem),
                "description": data.get("description", ""),
                "rules_count": len(data.get("rules", {})) if isinstance(data.get("rules", {}), dict) else 0,
            }
        )
    return items


def _get_template_by_name(template_name: str):
    for p in TEMPLATES_DIR.glob("*.json"):
        data = _safe_json_load(p)
        if isinstance(data, dict) and data.get("name") == template_name:
            return p, data
    guess = TEMPLATES_DIR / f"{_safe_template_filename(template_name)}.json"
    if guess.exists():
        data = _safe_json_load(guess)
        if isinstance(data, dict):
            return guess, data
    return None, None


def _safe_template_filename(name: str) -> str:
    clean = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", name.strip())
    return clean or "template"


def call_llm(user_prompt: str, system_prompt: str = "你是一个专业的文案助手。") -> str:
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}",
    }

    data = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.7,
        "max_tokens": 2000,
    }

    last_error = None
    for attempt in range(1, LLM_RETRY_TIMES + 1):
        try:
            resp = requests.post(ARK_URL, headers=headers, json=data, timeout=LLM_TIMEOUT_SECONDS)
            if resp.status_code != 200:
                raise RuntimeError(f"上游接口错误: {resp.status_code} - {resp.text}")

            payload = resp.json()
            return payload["choices"][0]["message"]["content"]
        except Exception as exc:
            last_error = exc
            if attempt < LLM_RETRY_TIMES:
                time.sleep(2 * attempt)

    raise RuntimeError(f"调用大模型失败(重试{LLM_RETRY_TIMES}次): {last_error}")


@app.route("/", methods=["GET"])
def index():
    # 固定入口页面，避免不同路由打开不同UI
    return send_from_directory(BASE_DIR, FIXED_ENTRY_FILE)


@app.route("/chat", methods=["GET"])
def chat_page():
    # 与首页保持一致
    return send_from_directory(BASE_DIR, FIXED_ENTRY_FILE)


@app.route("/chat.html", methods=["GET"])
def chat_html_page():
    # 历史路径统一指向固定入口
    return send_from_directory(BASE_DIR, FIXED_ENTRY_FILE)


@app.route("/index.html", methods=["GET"])
def index_html_page():
    # 历史路径统一指向固定入口
    return send_from_directory(BASE_DIR, FIXED_ENTRY_FILE)


@app.route("/studio", methods=["GET"])
def studio_page():
    return send_from_directory(BASE_DIR, FIXED_ENTRY_FILE)


@app.route("/<path:filename>", methods=["GET"])
def serve_files(filename: str):
    # 统一托管根目录与 app_portal 目录中的页面/样式
    target = BASE_DIR / filename
    if target.is_file():
        return send_from_directory(BASE_DIR, filename)

    portal_target = PORTAL_DIR / filename
    if portal_target.is_file():
        return send_from_directory(PORTAL_DIR, filename)

    return jsonify({"ok": False, "error": f"文件不存在: {filename}"}), 404


@app.route("/api/chat", methods=["POST"])
def chat():
    body = request.get_json(silent=True) or {}
    message = (body.get("message") or "").strip()
    system_prompt = (body.get("system_prompt") or "你是一个专业的文案助手。").strip()

    if not message:
        return jsonify({"ok": False, "error": "message 不能为空"}), 400

    if not API_KEY:
        return jsonify({"ok": False, "error": "请先在 demo.py 或 ARK_API_KEY 环境变量中配置 API Key"}), 500

    try:
        reply = call_llm(message, system_prompt=system_prompt)
        return jsonify({"ok": True, "reply": reply})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/template/generate", methods=["POST"])
def generate_template():
    """根据自然语言规则描述生成排版模板，可选保存到本地JSON文件。"""
    try:
        body = request.get_json(silent=True) or {}
        format_text = (body.get("format_text") or "").strip()
        template_name = (body.get("template_name") or "").strip()
        template_description = (body.get("template_description") or "").strip()
        should_save = bool(body.get("save", False))

        if not format_text:
            return jsonify({"ok": False, "error": "format_text 不能为空"}), 400
        if not template_name:
            return jsonify({"ok": False, "error": "template_name 不能为空"}), 400
        if not API_KEY:
            return jsonify({"ok": False, "error": "请先在 demo.py 或 ARK_API_KEY 环境变量中配置 API Key"}), 500

        ok, result = generate_template_from_text(
            format_text=format_text,
            template_name=template_name,
            template_description=template_description,
            llm_call=call_llm,
        )
        if not ok:
            return jsonify({"ok": False, **result}), 500

        payload = {"ok": True, "template": result}

        if should_save:
            filename = _safe_template_filename(template_name) + ".json"
            output_path = BASE_DIR / "templates" / filename
            save_template_to_file(result, output_path)
            payload["saved_path"] = str(output_path)

        return jsonify(payload)
    except Exception as exc:
        return jsonify({"ok": False, "error": f"服务端异常: {exc}"}), 500


@app.route("/api/templates", methods=["GET"])
def list_templates():
    try:
        return jsonify({"ok": True, "templates": _list_templates_meta()})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"读取模板失败: {exc}"}), 500


@app.route("/api/templates/upload", methods=["POST"])
def upload_template_json():
    """上传json模板并设置名称/描述。"""
    try:
        f = request.files.get("template_file")
        if not f:
            return jsonify({"ok": False, "error": "template_file 不能为空"}), 400

        template_name = (request.form.get("template_name") or "").strip()
        template_description = (request.form.get("template_description") or "").strip()
        if not template_name:
            return jsonify({"ok": False, "error": "template_name 不能为空"}), 400

        raw = f.read()
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception as exc:
            return jsonify({"ok": False, "error": f"模板JSON解析失败: {exc}"}), 400

        rules = payload.get("rules", payload)
        if not isinstance(rules, dict):
            return jsonify({"ok": False, "error": "模板格式错误：需要 rules 对象"}), 400

        template = {
            "name": template_name,
            "description": template_description or template_name,
            "rules": rules,
        }

        file_name = _safe_template_filename(template_name) + ".json"
        out_path = TEMPLATES_DIR / file_name
        save_template_to_file(template, out_path)

        return jsonify({"ok": True, "template": template, "saved_path": str(out_path)})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"上传模板失败: {exc}"}), 500


@app.route("/api/templates/save", methods=["POST"])
def save_template_direct():
    """直接保存模板JSON（用于编辑后保存）。"""
    try:
        body = request.get_json(silent=True) or {}
        template_name = (body.get("template_name") or "").strip()
        template_description = (body.get("template_description") or "").strip()
        rules = body.get("rules")

        if not template_name:
            return jsonify({"ok": False, "error": "template_name 不能为空"}), 400
        if not isinstance(rules, dict):
            return jsonify({"ok": False, "error": "rules 必须是对象"}), 400

        template = {
            "name": template_name,
            "description": template_description or template_name,
            "rules": rules,
        }
        out_path = TEMPLATES_DIR / (_safe_template_filename(template_name) + ".json")
        save_template_to_file(template, out_path)
        return jsonify({"ok": True, "template": template, "saved_path": str(out_path)})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"保存模板失败: {exc}"}), 500


@app.route("/api/templates/<template_name>", methods=["GET"])
def get_template(template_name: str):
    try:
        _, data = _get_template_by_name(template_name)
        if not data:
            return jsonify({"ok": False, "error": "模板不存在"}), 404
        return jsonify({"ok": True, "template": data})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"读取模板失败: {exc}"}), 500


@app.route("/api/docx/format", methods=["POST"])
def format_docx_with_template():
    """上传docx并选择模板排版，返回排版后的docx文件。"""
    try:
        f = request.files.get("docx_file")
        template_name = (request.form.get("template_name") or "").strip()

        if not f:
            return jsonify({"ok": False, "error": "docx_file 不能为空"}), 400
        if not template_name:
            return jsonify({"ok": False, "error": "template_name 不能为空"}), 400
        if not API_KEY:
            return jsonify({"ok": False, "error": "请先配置 API Key"}), 500

        filename_lower = (f.filename or "").lower()
        if not filename_lower.endswith(".docx"):
            return jsonify({"ok": False, "error": "仅支持 .docx 文件"}), 400

        _, template_data = _get_template_by_name(template_name)
        if not template_data:
            return jsonify({"ok": False, "error": f"模板不存在: {template_name}"}), 404

        rules = template_data.get("rules", {})
        source_bytes = f.read()

        from docx import Document

        doc = Document(io.BytesIO(source_bytes))
        paragraphs = [p.text for p in doc.paragraphs]

        ok, instructions_or_error = generate_formatting_instructions(paragraphs, rules, call_llm)
        if not ok:
            return jsonify({"ok": False, **instructions_or_error}), 500

        ok2, output_or_error = format_docx_bytes(source_bytes, instructions_or_error)
        if not ok2:
            return jsonify({"ok": False, **output_or_error}), 500

        out_name = (Path(f.filename).stem or "document") + "_formatted.docx"
        return send_file(
            io.BytesIO(output_or_error),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as exc:
        return jsonify({"ok": False, "error": f"排版失败: {exc}"}), 500


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="127.0.0.1", port=port, debug=True)
