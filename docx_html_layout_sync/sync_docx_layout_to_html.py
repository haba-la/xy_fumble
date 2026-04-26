"""CLI entry point for syncing DOCX layout styles into HTML."""

from __future__ import annotations

import argparse
from pathlib import Path

from layout_sync import sync_docx_layout_to_html


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read layout from DOCX and apply it to HTML via injected CSS.",
    )
    parser.add_argument("--docx", required=True, type=Path, help="Source DOCX template path")
    parser.add_argument("--html", required=True, type=Path, help="Input HTML path")
    parser.add_argument("--output", required=True, type=Path, help="Output HTML path")
    parser.add_argument(
        "--profile-output",
        type=Path,
        default=None,
        help="Optional JSON path to save extracted DOCX layout profile",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.docx.exists():
        raise FileNotFoundError(f"DOCX file not found: {args.docx}")
    if not args.html.exists():
        raise FileNotFoundError(f"HTML file not found: {args.html}")

    result = sync_docx_layout_to_html(
        docx_path=str(args.docx),
        html_path=str(args.html),
        output_html_path=str(args.output),
        profile_output_path=str(args.profile_output) if args.profile_output else None,
    )
    print(f"Styled HTML saved to: {result['output_html']}")
    if result["profile_json"]:
        print(f"Extracted profile saved to: {result['profile_json']}")


if __name__ == "__main__":
    main()

