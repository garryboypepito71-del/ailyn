"""Deterministic parsing helpers for OCR receipt text."""

import re


def parse_scanned_receipt(text):
    """Extract conservative material-entry suggestions from OCR text."""
    amount_pattern = r"(?:PHP|₱)\s*([0-9][0-9,]*(?:\.\d{1,2})?)"
    currency_amounts = [float(value.replace(",", "")) for value in re.findall(amount_pattern, text, re.IGNORECASE)]
    all_amounts = [float(value.replace(",", "")) for value in re.findall(r"\b[0-9][0-9,]*\.\d{2}\b", text)]
    amounts = currency_amounts or all_amounts
    quantity_match = re.search(
        r"\b(?:qty|quantity)\s*[:x-]?\s*(\d+)\b|\b(\d+)\s*(?:pcs?|pieces?|units?|x)\b",
        text,
        re.IGNORECASE,
    )
    quantity = int(next(value for value in quantity_match.groups() if value)) if quantity_match else 1
    delivery_match = re.search(r"delivery\D+([0-9][0-9,]*(?:\.\d{1,2})?)", text, re.IGNORECASE)
    delivery = float(delivery_match.group(1).replace(",", "")) if delivery_match else 0.0
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    name = next(
        (line for line in lines if not re.search(r"(?:php|₱|total|qty|quantity|delivery|receipt|invoice|date)", line, re.IGNORECASE)),
        "",
    )
    total = amounts[-1] if amounts else 0.0
    price = total / quantity if quantity > 0 else total
    return {"name": name[:120], "price": price, "qty": quantity, "delivery": delivery}