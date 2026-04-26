import json
import re
from pathlib import Path
from typing import Callable, Dict, Tuple


_ALIGNMENT_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "justify",
    "justified": "justify",
    "左对齐": "left",
    "左": "left",
    "居中": "center",
    "居中对齐": "center",
    "右对齐": "right",
    "右": "right",
    "两端对齐": "justify",
    "两端": "justify",
}


def _build_prompt(format_text: str) -> str:
    return """
你是一个专业的文档格式解析助手。请分析以下文本中的格式要求，提取出排版规则，并输出结构化 JSON。\n
文本内容：
<text>
__FORMAT_TEXT__
</text>

要求：
1. 字体名使用常见中文字体，例如：宋体、黑体、楷体、仿宋。
2. 字号使用中文字号，例如：小二、三号、四号、小四、五号。
3. alignment 仅允许：left、center、right、justify。
4. bold 必须是 true 或 false。
5. line_spacing 建议返回数值（1.0, 1.25, 1.5, 2.0 等）。

请仅返回 JSON，格式示例：
{
  "rules": {
    "标题": {
      "font": "黑体",
      "size": "小二",
      "bold": true,
      "alignment": "center",
      "line_spacing": 1.5
    },
    "正文": {
      "font": "宋体",
      "size": "小四",
      "bold": false,
      "alignment": "justify",
      "line_spacing": 1.5
    }
  }
}
""".replace("__FORMAT_TEXT__", format_text).strip()


def _extract_json(text: str) -> str:
    text = text.strip()

    code_fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", text, re.IGNORECASE)
    if code_fence:
        return code_fence.group(1).strip()

    brace_count = 0
    start = -1
    for i, ch in enumerate(text):
        if ch == "{":
            if brace_count == 0:
                start = i
            brace_count += 1
        elif ch == "}":
            brace_count -= 1
            if brace_count == 0 and start != -1:
                return text[start : i + 1]

    return text


def _normalize_alignment(value: str) -> str:
    if value is None:
        return "left"
    return _ALIGNMENT_MAP.get(str(value).strip().lower(), _ALIGNMENT_MAP.get(str(value).strip(), "left"))


def _normalize_rule(rule: Dict) -> Dict:
    normalized = dict(rule or {})
    normalized["font"] = str(normalized.get("font", "宋体"))
    normalized["size"] = str(normalized.get("size", "小四"))
    normalized["bold"] = bool(normalized.get("bold", False))
    normalized["alignment"] = _normalize_alignment(normalized.get("alignment", "left"))

    try:
        normalized["line_spacing"] = float(normalized.get("line_spacing", 1.5))
    except (TypeError, ValueError):
        normalized["line_spacing"] = 1.5

    return normalized


def _normalize_rules(rules: Dict) -> Dict:
    normalized = {}
    for key, value in (rules or {}).items():
        normalized[str(key)] = _normalize_rule(value if isinstance(value, dict) else {})
    return normalized


def generate_template_from_text(
    format_text: str,
    template_name: str,
    template_description: str,
    llm_call: Callable[[str, str], str],
) -> Tuple[bool, Dict]:
    if not format_text.strip():
        return False, {"error": "format_text 不能为空"}
    if not template_name.strip():
        return False, {"error": "template_name 不能为空"}

    prompt = _build_prompt(format_text)
    system_prompt = "你是一个严谨的文档排版规则抽取器。输出必须是可解析的 JSON。"

    try:
        response_text = llm_call(prompt, system_prompt=system_prompt)
        print(f"LLM *******************原始返回内容：{response_text}")
    except Exception as exc:
        return False, {"error": f"调用大模型失败: {exc}"}

    try:
        raw_json = _extract_json(response_text)
        parsed = json.loads(raw_json)
    except Exception as exc:
        return False, {"error": f"解析大模型返回内容失败: {exc}", "raw": response_text}

    rules = parsed.get("rules", parsed)
    if not isinstance(rules, dict):
        return False, {"error": "模型返回的 rules 不是对象", "raw": parsed}

    template = {
        "name": template_name,
        "description": template_description or template_name,
        "rules": _normalize_rules(rules),
    }
    return True, template


def save_template_to_file(template: Dict, output_file: Path) -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
