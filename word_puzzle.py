"""Offline word and arithmetic puzzle engine for the client dashboard."""

import random

WORD_PUZZLES = (
    ("BUDGET", "The approved amount available for project work."),
    ("MATERIAL", "A physical item used to build the project."),
    ("WORKER", "A person who performs construction work."),
    ("RECEIPT", "Proof that a purchase was made."),
    ("PLANNER", "A person or tool that organizes future work."),
    ("PROJECT", "The complete construction effort being managed."),
    ("PAYROLL", "The list and total of worker payments."),
    ("PROGRESS", "How much work has been completed."),
    ("FOUNDATION", "The structural base of a building."),
    ("MASONRY", "Construction using brick, block, or stone."),
    ("CARPENTER", "A skilled worker who builds with timber."),
    ("FOREMAN", "The person supervising a work crew."),
    ("BLUEPRINT", "A technical drawing used to guide construction."),
    ("CONCRETE", "A strong mixture used for foundations and slabs."),
    ("DELIVERY", "Transport of materials to the worksite."),
    ("SUPPLIER", "A business that provides goods or materials."),
    ("INVOICE", "A document requesting payment for goods or services."),
    ("LEDGER", "A record of financial transactions."),
    ("BALANCE", "Money remaining after costs are deducted."),
    ("SCHEDULE", "The planned timing of project work."),
    ("MILESTONE", "A major checkpoint in project progress."),
    ("INSPECTION", "A formal check of work quality or compliance."),
    ("SAFETY", "Protection from hazards at the worksite."),
    ("PLUMBING", "The system that carries water and drainage."),
    ("ELECTRICAL", "The system that distributes power through a building."),
    ("ROOFING", "The work of installing and protecting a roof."),
    ("PERMIT", "Official approval required before certain work begins."),
    ("ESTIMATE", "An expected cost or quantity calculated in advance."),
    ("VARIATION", "An approved change to the original work or cost."),
    ("CASHFLOW", "Money moving into and out of the project."),
    ("CLOSEOUT", "The final process of completing and handing over work."),
    ("WARRANTY", "A promise to repair or replace qualifying defects."),
)


def normalize_guess(guess):
    return "".join(str(guess).split()).upper()


def _scramble(answer, seed):
    letters = list(answer)
    random.Random(seed).shuffle(letters)
    scrambled = "".join(letters)
    return scrambled if scrambled != answer else answer[1:] + answer[:1]


def puzzle_for_level(level, seed=0):
    index = (max(0, int(level) - 1) + int(seed)) % len(WORD_PUZZLES)
    answer, hint = WORD_PUZZLES[index]
    return {"answer": answer, "scramble": _scramble(answer, int(seed) + int(level) * 97), "hint": hint}


def check_guess(level, guess, seed=0):
    puzzle = puzzle_for_level(level, seed)
    return normalize_guess(guess) == puzzle["answer"]


def math_problem(level, seed=0):
    """Return a repeatable arithmetic problem from easy to hardest."""
    rng = random.Random(int(seed) + int(level) * 1009)
    difficulty = max(1, int(level))
    if difficulty == 8:
        left, right, divisor = rng.randint(20, 100), rng.randint(2, 20), rng.randint(2, 8)
        return f"({left} - {right}) x {divisor}", float((left - right) * divisor)
    scale = min(1000, 100 + difficulty * 35)
    left, right, extra = rng.randint(20, scale), rng.randint(2, scale // 3), rng.randint(2, scale // 4)
    if difficulty % 4 == 1:
        return f"{left} + {right} x {extra}", float(left + right * extra)
    if difficulty % 4 == 2:
        return f"({left} - {right}) / {extra}", float((left - right) / extra)
    if difficulty % 4 == 3:
        return f"{left} x {right} - {extra}", float(left * right - extra)
    return f"({left} + {right}) x {extra}", float((left + right) * extra)
    if difficulty == 2:
        left, right = rng.randint(10, 60), rng.randint(1, 20)
        return f"{left} - {right}", float(left - right)
    if difficulty == 3:
        left, right = rng.randint(2, 12), rng.randint(2, 12)
        return f"{left} x {right}", float(left * right)
    if difficulty == 4:
        divisor, quotient = rng.randint(2, 12), rng.randint(2, 15)
        return f"{divisor * quotient} / {divisor}", float(quotient)
    if difficulty == 5:
        left, right, extra = rng.randint(2, 20), rng.randint(2, 12), rng.randint(1, 20)
        return f"{left} x {right} + {extra}", float(left * right + extra)
    if difficulty == 6:
        left, right, divisor = rng.randint(10, 80), rng.randint(2, 30), rng.randint(2, 8)
        return f"({left} + {right}) / {divisor}", (left + right) / divisor
    if difficulty == 7:
        base, exponent = rng.randint(2, 9), rng.randint(2, 3)
        extra = rng.randint(5, 30)
        return f"{base}^{exponent} + {extra}", float(base ** exponent + extra)
    left, right, divisor = rng.randint(20, 100), rng.randint(2, 20), rng.randint(2, 8)
    return f"({left} - {right}) x {divisor}", float((left - right) * divisor)


def check_math_answer(expected, answer):
    try:
        return abs(float(answer) - float(expected)) < 1e-9
    except (TypeError, ValueError):
        return False
