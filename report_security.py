"""Safe formatting helpers for generated reports."""

import html
import json
import re


def escape_report_text(value):
    return html.escape(str(value), quote=True)


def report_download_name(title):
    safe_title = re.sub(r"[^A-Za-z0-9._-]+", "_", str(title)).strip("._") or "receipt"
    return json.dumps(f"{safe_title}_Receipt.png")