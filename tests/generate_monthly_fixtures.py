"""Synthetic monthly inputs; never imports application code or real business data."""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

import generate_fixtures as daily

MARKER = ".cashrec-monthly-test-fixture.json"
DATES = [datetime(2026, 10, 19), daily.PREVIOUS, daily.CURRENT]
SCENARIOS = ("good", "no-previous", "pending-uncached", "pending-wrong-total", "pending-duplicate", "pending-missing-entity")
PENDING = {
    "2026.10.19": {"TMA": (800, 750, 50), "TMG": (1700, 1580, 120)},
    "2026.10.22": {"TMA": (900, 820, 80), "TMG": (1800, 1650, 150)},
    "2026.10.25": {"TMA": (1000, 900, 100), "TMG": (2000, 1800, 200)},
}
SETTLEMENT_TYPES = ["SETTLE", "Posting TF TD", "MISC CIMB TF TD", "CFD TF TD", "Rebate TF TD", "Pending TF TD"]


def settlement_inputs(path: Path, day: datetime) -> None:
    book = load_workbook(path)
    sheet = book["BANK BD"]
    for index, category in enumerate(SETTLEMENT_TYPES, 1):
        amount = 10 + index if day == daily.CURRENT else index
        sheet.append([amount, None, "Synthetic TMA " + category, category, "USD", "TMA-Retail", amount])
    tmg_amount = -7 if day == daily.CURRENT else -3
    sheet.append([None, -tmg_amount, "Synthetic TMG SETTLE", "SETTLE", "USD", "TMG-Retail", tmg_amount])
    sheet.append([999, None, "Synthetic TMG excluded Posting", "Posting TF TD", "USD", "TMG-Retail", 999])
    sheet.append([None, 8, "Synthetic unrelated retail", "Unrelated adjustment", "USD", "TMA-Retail", -8])
    book.save(path)


def pending_snapshot(path: Path, day: datetime, scenario: str) -> None:
    book = load_workbook(path)
    sheet = book["PSP On The Way"]
    # Replace only the synthetic pending sheet, preserving the transaction source.
    book.remove(sheet)
    sheet = book.create_sheet("PSP On The Way", 0)
    sheet.append([
        "PSP Gateway", "Date", "Currency", "RATE", "Settle Amount", "PSP Settle USD", "Jur", None,
        "Date", "Bank PSP Settle Amount", "Bank currency", "RATE", "BANK SETTLE USD", "Jur", None,
        None, "Branch", "TOTAL PSP SETTLE USD", "TOTAL BANK SETTLE USD", "Total PSP On the way USD",
    ])
    values = PENDING[day.strftime("%Y.%m.%d")]
    # Entity rows deliberately move; MT4 regulator matching must not reinterpret them.
    branches = [("TTG", (9900, 900, 9000)), ("TMG", values["TMG"]), ("TMNZ", (8000, 7000, 1000)), ("TMA", values["TMA"])]
    if day == daily.CURRENT and scenario == "pending-missing-entity":
        branches = [row for row in branches if row[0] != "TMA"]
    if day == daily.CURRENT and scenario == "pending-duplicate":
        branches.append(("TMA", values["TMA"]))
    caches = {}
    for branch, (settled, received, pending) in branches:
        row = sheet.max_row + 1
        # Historical detail dates are not a date attached to the Q:T cumulative summary.
        detail = [None, datetime(2026, 10, 1), None, None, None, None, "legacy ASIC/VFSC detail", None,
                  datetime(2026, 10, 2), None, None, None, None, None, None, None]
        sheet.append(detail + [branch, settled, received, f"=R{row}-S{row}"])
        if not (day == daily.CURRENT and scenario == "pending-uncached" and branch == "TMA"):
            caches[f"T{row}"] = pending + (1 if day == daily.CURRENT and scenario == "pending-wrong-total" and branch == "TMA" else 0)
    book.save(path)
    daily.cache_formula_numbers(path, "PSP On The Way", caches)


def generate(root: Path, scenario: str) -> None:
    if root.exists() and any(root.iterdir()) and not (root / MARKER).exists():
        raise ValueError("Refusing to overwrite a non-fixture directory.")
    if scenario == "no-previous" and root.exists() and any((root / d.strftime("%Y.%m.%d")).exists() for d in DATES[:-1]):
        raise ValueError("Use a fresh directory for no-previous; earlier inputs are preserved.")
    root.mkdir(parents=True, exist_ok=True)
    days = [daily.CURRENT] if scenario == "no-previous" else DATES
    for day in days:
        daily.write_day(root, day)
        needed = root / day.strftime("%Y.%m.%d") / "Needed files"
        if day == DATES[0]:
            psp = needed / "PSP+BANK/PSP Balance Summary.xlsx"
            book = load_workbook(psp)
            for column in (5, 6):
                book["OCT.2026"].cell(1, column).value = day
            book.save(psp)
        settlement_inputs(needed / "PSP+BANK/BANK BD.xlsx", day)
        pending_snapshot(needed / "PSP+BANK/BANK BD.xlsx", day, scenario)
    daily.write_mapping(root / "db/tmc/account_mapping.xlsx")
    oracle = Path(__file__).parent / "fixtures/monthly_expected.json"
    (root / "monthly_expected.json").write_text(oracle.read_text(encoding="utf-8"), encoding="utf-8")
    expected = json.loads((Path(__file__).parent / "fixtures/expected.json").read_text(encoding="utf-8"))
    transactions = expected["sheets"]["TransactionExchange"]
    for index, category in enumerate(SETTLEMENT_TYPES, 1):
        transactions["rows"].append({"Source":"BANK BD", "Comment":"Synthetic TMA " + category,
            "Rational":category, "Branch":"TMA", "Type":"Retail", "Currency":"USD", "Amount":10 + index, "USD Equivalent":10 + index})
    transactions["rows"].extend([
        {"Source":"BANK BD", "Comment":"Synthetic TMG SETTLE", "Rational":"SETTLE", "Branch":"TMG", "Type":"Retail", "Currency":"USD", "Amount":-7, "USD Equivalent":-7},
        {"Source":"BANK BD", "Comment":"Synthetic TMG excluded Posting", "Rational":"Posting TF TD", "Branch":"TMG", "Type":"Retail", "Currency":"USD", "Amount":999, "USD Equivalent":999},
        {"Source":"BANK BD", "Comment":"Synthetic unrelated retail", "Rational":"Unrelated adjustment", "Branch":"TMA", "Type":"Retail", "Currency":"USD", "Amount":-8, "USD Equivalent":-8},
    ])
    transactions["total_usd"] = 1447
    transactions["source_totals_usd"]["BANK BD"] = 1197
    (root / "expected.json").write_text(json.dumps(expected, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (root / MARKER).write_text(json.dumps({"synthetic_only": True, "scenario": scenario, "dates": [d.strftime("%Y.%m.%d") for d in days]}, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--scenario", choices=SCENARIOS, default="good")
    args = parser.parse_args()
    generate(args.output.resolve(), args.scenario)
    print(json.dumps({"synthetic_only": True, "scenario": args.scenario, "workspace": str(args.output.resolve()), "business_acceptance_executed": False}))


if __name__ == "__main__":
    main()
