"""Pure business rules for payroll calculations."""

FULL_DAY_RATES = {
    "Carpenter": 700.0,
    "Mason": 700.0,
    "Laborer": 500.0,
    "Electrician": 900.0,
    "Plumber": 900.0,
    "Painter": 650.0,
    "Welder": 850.0,
    "Foreman": 1000.0,
}

TIER_TABLE = {"full_day": 1.0, "half_day": 0.5, "partial_day": 0.25}


def get_partial_rate(rate, worked_days):
    """Return pay for the fractional part of a workday."""
    fractional_days = float(worked_days) % 1
    if fractional_days <= 0:
        return 0.0
    return float(rate) * fractional_days


def calculate_labor_pay(worked_days, role):
    """Return gross pay and its full-day/fractional breakdown."""
    days = float(worked_days)
    if days < 0:
        raise ValueError("Worked days cannot be negative.")
    if role not in FULL_DAY_RATES:
        raise ValueError(f"Unknown labor role: {role}")
    rate = FULL_DAY_RATES[role]
    full_pay = int(days) * rate
    partial_pay = get_partial_rate(rate, days)
    return full_pay + partial_pay, full_pay, partial_pay