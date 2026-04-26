"""Sync DOCX layout styles into HTML output."""

from .layout_sync import (
    build_css_from_profile,
    extract_docx_layout_profile,
    inject_css_into_html,
    sync_docx_layout_to_html,
)

__all__ = [
    "build_css_from_profile",
    "extract_docx_layout_profile",
    "inject_css_into_html",
    "sync_docx_layout_to_html",
]

