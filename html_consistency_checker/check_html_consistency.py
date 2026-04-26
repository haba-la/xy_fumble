import argparse
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, Optional


MODULE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = MODULE_DIR.parent
ROOT_DIR = PROJECT_DIR.parent

DEFAULT_HTML_PATH = PROJECT_DIR / "tempt_files" / "output_synced.html"
DEFAULT_PROMPT_PATH = MODULE_DIR / "prompt_consistency.txt"
DEMO_PY_PATH = ROOT_DIR / "文档规则" / "demo.py"
DEFAULT_OUTPUT_PATH = PROJECT_DIR / "runtime" / "consistency_result.json"


def _load_demo_module(demo_file: Path):
	if not demo_file.exists():
		raise FileNotFoundError(f"找不到模型配置文件: {demo_file}")

	demo_dir = str(demo_file.parent)
	if demo_dir not in sys.path:
		sys.path.insert(0, demo_dir)

	spec = importlib.util.spec_from_file_location("doc_rules_demo", str(demo_file))
	if spec is None or spec.loader is None:
		raise RuntimeError(f"无法加载模块: {demo_file}")

	module = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(module)
	return module


def _extract_json(text: str):
	text = text.strip()

	# 优先直接解析
	try:
		return json.loads(text)
	except Exception:
		pass

	# 兼容 ```json ... ``` 包裹
	fenced = re.search(r"```(?:json)?\s*(\{[\s\S]*\})\s*```", text, flags=re.IGNORECASE)
	if fenced:
		try:
			return json.loads(fenced.group(1))
		except Exception:
			pass

	# 兜底提取首个 JSON 对象
	first_brace = text.find("{")
	last_brace = text.rfind("}")
	if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
		candidate = text[first_brace : last_brace + 1]
		return json.loads(candidate)

	raise ValueError("模型返回内容中未找到可解析的 JSON")


def _build_user_prompt(prompt_text: str, html_text: str) -> str:
	return (
		f"{prompt_text}\n\n"
		"以下是需要审查的 HTML 内容，请只按上面的 JSON 结构输出结果：\n"
		"```html\n"
		f"{html_text}\n"
		"```\n"
	)


def run_html_consistency_check(
	html_path: Path,
	prompt_path: Optional[Path] = None,
	output_path: Optional[Path] = None,
	demo_py_path: Optional[Path] = None,
) -> Dict[str, Any]:
	"""Run consistency check for visible HTML text and optionally persist JSON output."""
	use_prompt_path = prompt_path or DEFAULT_PROMPT_PATH
	use_output_path = output_path or DEFAULT_OUTPUT_PATH
	use_demo_py = demo_py_path or DEMO_PY_PATH

	if not html_path.exists():
		raise FileNotFoundError(f"HTML 文件不存在: {html_path}")
	if not use_prompt_path.exists():
		raise FileNotFoundError(f"Prompt 文件不存在: {use_prompt_path}")

	html_text = html_path.read_text(encoding="utf-8")
	prompt_text = use_prompt_path.read_text(encoding="utf-8")

	demo_module = _load_demo_module(use_demo_py)
	if not hasattr(demo_module, "call_llm"):
		raise AttributeError(f"{use_demo_py} 中未找到 call_llm 函数")

	model_output = demo_module.call_llm(
		user_prompt=_build_user_prompt(prompt_text, html_text),
		system_prompt="你是一个严谨的文本一致性审查助手。",
	)

	result_json = _extract_json(model_output)

	if use_output_path:
		use_output_path.parent.mkdir(parents=True, exist_ok=True)
		use_output_path.write_text(
			json.dumps(result_json, ensure_ascii=False, indent=2),
			encoding="utf-8",
		)

	return result_json


def main():
	parser = argparse.ArgumentParser(description="使用大模型审查 HTML 文本内容一致性")
	parser.add_argument("--html", default=str(DEFAULT_HTML_PATH), help="待审查 HTML 文件路径")
	parser.add_argument("--prompt", default=str(DEFAULT_PROMPT_PATH), help="审查提示词文件路径")
	parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH), help="输出 JSON 文件路径")
	args = parser.parse_args()

	html_path = Path(args.html)
	prompt_path = Path(args.prompt)
	output_path = Path(args.output)
	result_json = run_html_consistency_check(
		html_path=html_path,
		prompt_path=prompt_path,
		output_path=output_path,
	)

	print(json.dumps(result_json, ensure_ascii=False, indent=2))
	print(f"\n结果已保存到: {output_path}")


if __name__ == "__main__":
	main()
