"""Monthly financial closing rules."""

from collections import defaultdict
from datetime import datetime


def _month_key(record):
    recorded_at = record.get("recorded_at")
    if recorded_at:
        try:
            return datetime.fromisoformat(recorded_at).strftime("%Y-%m")
        except ValueError:
            pass
    try:
        return datetime.strptime(record.get("date", ""), "%b %d, %Y").strftime("%Y-%m")
    except ValueError:
        return ""


def month_summary(state, month):
    totals = defaultdict(float)
    for record in state.get("records", []):
        if record.get("type") in {"material", "expense", "excess"} and _month_key(record) == month:
            totals[record.get("type")] += float(record.get("amount", 0) or 0)
    for record in state.get("labor_records", []):
        if _month_key(record) == month:
            totals["labor"] += float(record.get("net", 0) or 0)
    for record in state.get("payroll_expenses", []):
        if _month_key(record) == month:
            totals["payroll"] += float(record.get("price", 0) or 0)
    totals["spent"] = totals["material"] + totals["expense"] + totals["labor"] + totals["payroll"]
    return dict(totals)


def is_month_closed(closes, month):
    return bool(closes.get(month, {}).get("closed"))


def close_month(closes, month, summary, closed_at, closed_by="Project Manager", note=""):
    if is_month_closed(closes, month):
        raise ValueError(f"Month {month} is already closed.")
    result = dict(closes)
    result[month] = {
        "closed": True,
        "closed_at": closed_at,
        "closed_by": closed_by,
        "note": note.strip(),
        "summary": {key: float(value) for key, value in summary.items()},
    }
    return result
