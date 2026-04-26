from __future__ import annotations

import argparse
import tempfile
from pathlib import Path
from typing import Optional


def _import_mammoth():
    try:
        import mammoth  # type: ignore

        return mammoth
    except ImportError as exc:
        raise ImportError(
            "Missing dependency 'mammoth'. Install it with: pip install mammoth"
        ) from exc


def convert_docx_to_temp_html(
    docx_path: Path,
    output_html_path: Optional[Path] = None,
    temp_dir: Optional[Path] = None,
) -> Path:
    """Convert a DOCX file into an HTML file for later editing."""
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    mammoth = _import_mammoth()

    with docx_path.open("rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)

    html_body = result.value
    html_text = (
        "<!doctype html>\n"
        "<html lang=\"en\">\n"
        "<head>\n"
        "  <meta charset=\"utf-8\" />\n"
        "  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n"
        "  <title>Temporary Converted HTML</title>\n"
        "</head>\n"
        "<body>\n"
        f"{html_body}\n"
        "</body>\n"
        "</html>\n"
    )

    if output_html_path is None:
        if temp_dir is not None:
            temp_dir.mkdir(parents=True, exist_ok=True)
            html_file = tempfile.NamedTemporaryFile(
                prefix="docx_converted_",
                suffix=".html",
                dir=temp_dir,
                delete=False,
                mode="w",
                encoding="utf-8",
            )
        else:
            html_file = tempfile.NamedTemporaryFile(
                prefix="docx_converted_",
                suffix=".html",
                delete=False,
                mode="w",
                encoding="utf-8",
            )

        with html_file:
            html_file.write(html_text)
        return Path(html_file.name)

    output_html_path.parent.mkdir(parents=True, exist_ok=True)
    output_html_path.write_text(html_text, encoding="utf-8")
    return output_html_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert DOCX to temporary HTML")
    parser.add_argument("--docx", type=Path, required=True, help="Input DOCX path")
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional output HTML path. If omitted, a temp HTML file is created.",
    )
    parser.add_argument(
        "--temp-dir",
        type=Path,
        default=None,
        help="Optional temp directory when --output is omitted.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output = convert_docx_to_temp_html(
        docx_path=args.docx,
        output_html_path=args.output,
        temp_dir=args.temp_dir,
    )
    print(f"Temporary HTML saved to: {output}")


if __name__ == "__main__":
    main()
