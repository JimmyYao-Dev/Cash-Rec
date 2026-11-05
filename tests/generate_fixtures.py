"""Generate synthetic, independently specified inputs for installed-UI acceptance.

This script never imports Cash Rec code. It is a developer/test tool, not an
application entry point or an end-user runtime dependency.
"""

from __future__ import annotations

import argparse
import csv
import json
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook


PREVIOUS = datetime(2026, 10, 22)
CURRENT = datetime(2026, 10, 25)
MARKER = ".cashrec-test-fixture.json"
SCENARIOS = (
    "good", "ambiguous-regulator", "zero-rate", "missing-rate",
    "bank-conflict", "missing-psp-date", "missing-input", "unknown-bank",
    "no-previous", "corrupt-bank-bd",
)
EQUITY_HEADERS = [
    "dt_report", "report_jurisdiction", "login", "is_ib", "group", "user_id",
    "equity", "EQUITY_USD", "positive_equity", "currency", "jurisdiction", "is_cimb",
]
TRANSACTION_HEADERS = [
    "date", "user_id", "crm_transaction_no", "trading_server_transaction_no",
    "trading_server_type", "trading_server", "trading_server_login", "fund_amount",
    "trading_server_currency", "payment_amount", "payment_currency", "exchange_rate",
    "surcharge_amount", "comment", "accounting_comment", "jurisdiction",
    "payment_channel", "mapping_deal", "mapping_deal_currency",
]
BANK_HEADERS = [
    "Account Name", "Account Number", "Currency", "Debits", "No Of Debits",
    "Credits", "No Of Credits", "EOD balance",
]
BANK_MAPPING_HEADERS = [
    "Legacy Code", "Account Name", "Account Number", "Company", "Branch",
    "Account Type", "Settle Type", "Client Type", "Bank", "Source File",
]
PSP_MAPPING_HEADERS = ["Source", "PSP", "Currency", "Company", "Branch", "Source File"]


def write_csv(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)


def save_book(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.remove(book.active)
    for name, rows in sheets.items():
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
        sheet.freeze_panes = "A2"
    book.save(path)


def equity_rows(day: datetime) -> list[list[object]]:
    current = day == CURRENT
    # TMA/Standard/USD is 40+48=88 currently versus 80 previously.
    # Jurisdiction, group, equity, EQUITY_USD and is_cimb are deliberate decoys.
    data = [
        [day + timedelta(days=1), "ASICGL", "SYN-E1", False, "synthetic-va", "SYN-U1", 4000, 5000, 40 if current else 80, "USD", "VFSC", True],
        [day - timedelta(days=2), " AsIcGL ", "SYN-E2", False, "synthetic", "SYN-U2", 4800, 6000, 48 if current else 0, "USD", "FMA", False],
        [day, "ASIC", "SYN-E3", True, "synthetic", "SYN-U3", 2000, 2000, 20 if current else 10, "USD", "FMA", False],
        [day, "ASIC", "SYN-E4", True, "synthetic", "SYN-U4", 12500, 15000, 125 if current else 100, "AUD", "VFSC", False],
        [day, "VFSCSEA", "SYN-E5", False, "synthetic", "SYN-U5", 20000, 20000, 200 if current else 100, "USC", "ASIC", False],
        [day, " vFsCCN ", "SYN-E6", False, "synthetic", "SYN-U6", 3750, 3750, 37.5 if current else 25, "EUR", "ASIC", False],
        [day, "FMA", "SYN-EXCLUDED1", False, "synthetic", "SYN-X1", 999, 999, 999, "USD", "ASIC", False],
        [day, "", "SYN-EXCLUDED2", False, "synthetic", "SYN-X2", 888, 888, 888, "USD", "ASIC", False],
    ]
    return [[v.isoformat(sep=" ") if isinstance(v, datetime) else str(v).lower() if isinstance(v, bool) else v for v in row] for row in data]


def transaction_rows(day: datetime) -> list[list[object]]:
    examples = [
        ("SYN-T1", day - timedelta(days=2), 100, "USD", "NAB synthetic wire", "Deposit", "ASICGL", "Withdrawal"),
        ("SYN-T2", day - timedelta(days=1), -30, "USD", "Gateway-A withdrawal", "Withdrawal", "ASICGL", "Deposit"),
        ("SYN-T3", day, 125, "AUD", "Gateway-B deposit", "Deposit", "VFSCSEA", "Withdrawal"),
        ("SYN-T4", day - timedelta(days=2), -25, "EUR", "Other bank withdrawal", "Withdrawal", "VFSC", "Deposit"),
        ("SYN-T5", day, 100, "USD", "Gateway equal-amount distinct transaction", "Deposit", " ASIC ", "Withdrawal"),
        ("SYN-T6", day, 999, "USD", "Excluded other regulator", "Deposit", "FMA", "Deposit"),
        ("SYN-T7", day, 888, "USD", "Excluded empty regulator", "Deposit", "", "Deposit"),
    ]
    rows = []
    for identity, source_day, amount, currency, comment, direction, regulator, decoy in examples:
        rows.append([
            source_day.strftime("%Y-%m-%d"), "SYN-U", identity, identity, "MT4", "SYNTHETIC",
            "SYN-LOGIN", amount, currency, amount * 10, "CNY", 10, 0, comment,
            direction, regulator, "Synthetic Gateway", decoy, "USD",
        ])
    return rows


def write_bank(path: Path, records: list[tuple[str, str, str, float]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        # Exactly seven pre-header rows, including a misleading report total.
        handle.write("sep=,\n")
        writer.writerows([
            ["SYNTHETIC ACCOUNT BALANCE REPORT"], ["Generated for automated acceptance"],
            ["Total balance", 999999], ["No real customer or bank data"], [], [], BANK_HEADERS,
        ])
        for name, number, currency, balance in records:
            writer.writerow([name, number, currency, 0, 0, 0, 0, balance])


def write_psp(path: Path) -> None:
    old, current, future = PREVIOUS, CURRENT, CURRENT + timedelta(days=1)
    header_dates = [None, None, None, None, old, old, current, current, future, future]
    fields = ["Source", "WAY", "CURR", "Rate", "Amount", "Eq USD", "Amount", "Eq USD", "Amount", "Eq USD"]
    # Eq USD and adjacent dates intentionally disagree with daily rate conversion.
    rows = [header_dates, fields,
        ["PSPAU", "Gateway-A", "USD", 7, 80, 8000, 88, 8800, 99999, 99999],
        ["PSPGL", "Gateway-B", "AUD", 7, 100, 10000, 125, 12500, 99999, 99999],
        ["PSPGL", "NewGateway", "EUR", 7, 25, 2500, 37.5, 3750, 99999, 99999],
        ["PSPGL", "Shared", "USD", 7, 10, 1000, 20, 2000, 99999, 99999],
        ["PSPAU", "Shared", "USD", 7, 10, 1000, 15, 1500, 99999, 99999],
        ["PSPGL-TTG", "Excluded", "USD", 7, 9999, 9999, 9999, 9999, 9999, 9999],
        ["PSPGL", "Total", "USD", 7, 99999, 99999, 99999, 99999, 99999, 99999],
    ]
    save_book(path, {"SEP.2026": [[None] * 4 + [datetime(2026, 9, 30)] * 2, fields[:6], ["PSPGL", "WrongMonth", "USD", 1, 99999, 99999]], "OCT.2026": rows})


def write_mapping(path: Path, conflict: bool = False, supplement_psp: bool = False) -> None:
    rows = [BANK_MAPPING_HEADERS,
        ["SYN-B1", "Synthetic Retail", "001-1001", "TMC", "ANZO", "Clients", "NA", "Retail", "NAB", "NAB AU"],
        ["SYN-B1", "Synthetic Retail Alias", "001 1001", "TMC", "ANZO", "Clients", "NA", "Retail", "NAB", "NAB AU"],
        ["SYN-B2", "Synthetic Settle", "001-1002", "TMC", "TMA", "Settle", "Others", "NA", "NAB", "NAB AU"],
        ["SYN-B3", "Synthetic Wholesale", "001-1003", "TMC", "TMA", "Clients", "NA", "Wholesale", "NAB", "NAB AU"],
        ["SYN-B4", "Synthetic GL Wholesale", "009-0001", "TMC", "TMG", "Clients", "NA", "Wholesale", "NAB", "NAB GL"],
        ["SYN-B5", "Synthetic GL Insto", "009-0002", "TMC", "TMG", "Clients", "NA", "Insto", "NAB", "NAB GL"],
    ]
    if conflict:
        rows.append(["SYN-CONFLICT", "Conflicting class", "001-1001", "TMC", "ANZO", "Settle", "Others", "NA", "NAB", "NAB AU"])
    psp = [PSP_MAPPING_HEADERS,
        ["PSPAU", "Gateway-A", "USD", "TMC", "TMG", "PSP Balance Summary"],
        ["PSPGL", "Gateway-B", "AUD", "TMC", "TMG", "PSP Balance Summary"],
        ["PSPGL", "Shared", "USD", "TMC", "TMG", "PSP Balance Summary"],
        ["PSPAU", "Shared", "USD", "TMC", "TMA", "PSP Balance Summary"],
    ]
    if supplement_psp:
        psp.append(["PSPGL", "NewGateway", "EUR", "TMC", "TMG", "PSP Balance Summary"])
    save_book(path, {"BANK MAPPING": rows, "PSP MAPPING": psp, "TP MAPPING": [["Trading Pool", "Company", "Branch", "Amount", "DEP", "Currency", "Source File"]]})


def write_day(root: Path, day: datetime) -> None:
    needed = root / day.strftime("%Y.%m.%d") / "Needed files"
    current = day == CURRENT
    write_csv(needed / "MT4/Risk/Accounting Equity Report.csv", EQUITY_HEADERS, equity_rows(day))
    write_csv(needed / "MT4/MT4/Transaction Summary Report.csv", TRANSACTION_HEADERS, transaction_rows(day))
    write_bank(needed / "Bank/NAB AU.csv", [
        ("Synthetic Retail", "001-1001", "USD", 880 if current else 800),
        ("Synthetic Settle", "001-1002", "AUD", 125 if current else 100),
        ("Synthetic Wholesale", "001-1003", "USD", 10),
    ])
    write_bank(needed / "Bank/NAB GL.csv", [
        ("Synthetic GL Wholesale", "009-0001", "USD", 400 if current else 350),
        ("Synthetic GL Insto", "009-0002", "EUR", 25 if current else 12.5),
    ])
    write_psp(needed / "PSP+BANK/PSP Balance Summary.xlsx")
    bank_rows = [
        ["Credit", "Debit", "Comment", "Rational", "Currency", "Company", "Amount"],
        [25, None, "Synthetic retail transfer", "Adjustment", "USD", "TMG-Retail", 25],
        [None, 10, "Synthetic wholesale withdrawal", "Withdrawal", "EUR", "TMA-Wholesale", -10],
        [50, None, "Synthetic promotion", "promote", "AUD", "TMG-promote", 50],
        [None, 20, "PSP CN Commission", "PSP CN Commission", "USD", "TMA-PSP", -20],
        [None, 5, "PSP Commission", "PSP Commission", "USD", "TMG-PSP", -5],
        [999, None, "Excluded TTG", "Deposit", "USD", "TTG-Retail", 999],
        [888, None, "Excluded unknown", "Deposit", "USD", "UNKNOWN-Retail", 888],
        [100, None, "Synthetic BANK BD PSP deposit", "Deposit", "USD", "TMG-PSP", 100],
    ]
    pending = [["PSP Gateway", "Date", "Currency", "RATE", "Settle Amount", "PSP Settle USD", "Jur", None, "Date", "Bank PSP Settle Amount", "Bank currency", "RATE", "BANK SETTLE USD", "Jur", None, None, "Branch", "TOTAL PSP SETTLE USD", "TOTAL BANK SETTLE USD", "Total PSP On the way USD"]]
    for branch, psp, bank in [("TMA", 1000, 900), ("TMG", 2000, 1800), ("TTG", 9999, 0)]:
        pending.append([None] * 16 + [branch, psp, bank, psp - bank])
    save_book(needed / "PSP+BANK/BANK BD.xlsx", {"PSP On The Way": pending, "BANK BD": bank_rows})
    save_book(needed / "Rate/VFSC Rate.xlsx", {"Sheet1": [
        ["Date", "Symb", "Currency", "Units per USD", "USD per Unit", "Notes"],
        [day, "USD", "USD US Dollar", 1, 1, "synthetic"],
        [day, "AUD", "AUD Australian Dollar", 1.25, 0.8, "synthetic"],
        [day, "EUR", "EUR Euro", 1.25, 0.8, "synthetic"],
        [day, "USC", "USC US Cents", 100, 0.01, "synthetic units"],
    ]})


def apply_scenario(root: Path, scenario: str) -> None:
    needed = root / CURRENT.strftime("%Y.%m.%d") / "Needed files"
    if scenario == "ambiguous-regulator":
        with (needed / "MT4/Risk/Accounting Equity Report.csv").open("a", encoding="utf-8", newline="") as handle:
            csv.writer(handle).writerow(["2026-10-26", "ASIC-VFSC", "SYN-AMBIGUOUS", "false", "synthetic", "SYN-X", 99, 99, 99, "USD", "ASIC", "false"])
    elif scenario in {"zero-rate", "missing-rate"}:
        path = needed / "Rate/VFSC Rate.xlsx"
        book = load_workbook(path)
        if scenario == "zero-rate":
            book.active["D3"] = 0
        else:
            book.active.delete_rows(3)
        book.save(path)
    elif scenario == "missing-psp-date":
        path = needed / "PSP+BANK/PSP Balance Summary.xlsx"
        book = load_workbook(path)
        book["OCT.2026"]["G1"] = CURRENT + timedelta(days=2)
        book["OCT.2026"]["H1"] = CURRENT + timedelta(days=2)
        book.save(path)
    elif scenario == "missing-input":
        (needed / "Bank/NAB GL.csv").unlink()
    elif scenario == "unknown-bank":
        with (needed / "Bank/NAB AU.csv").open("a", encoding="utf-8", newline="") as handle:
            csv.writer(handle).writerow(["Unknown synthetic account", "001-9999", "USD", 0, 0, 0, 0, 12])
    elif scenario == "corrupt-bank-bd":
        (needed / "PSP+BANK/BANK BD.xlsx").write_text("Deliberately invalid synthetic workbook", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path(".scratch/implementation/fixture-workspace"))
    parser.add_argument("--scenario", choices=SCENARIOS, default="good")
    parser.add_argument("--supplement-psp", action="store_true", help="Add the previously unmapped synthetic PSP to mapping.")
    args = parser.parse_args()
    root = args.output.resolve()
    if root.exists() and any(root.iterdir()) and not (root / MARKER).exists():
        parser.error("Refusing to overwrite a non-fixture directory (fixture marker absent).")
    root.mkdir(parents=True, exist_ok=True)
    if args.scenario == "no-previous" and (root / PREVIOUS.strftime("%Y.%m.%d")).exists():
        parser.error("Use a fresh output directory for no-previous; existing earlier inputs are preserved.")
    days = [CURRENT] if args.scenario == "no-previous" else [PREVIOUS, CURRENT]
    for day in days:
        write_day(root, day)
    write_mapping(root / "db/tmc/account_mapping.xlsx", args.scenario == "bank-conflict", args.supplement_psp)
    apply_scenario(root, args.scenario)
    oracle = Path(__file__).resolve().parent / "fixtures/expected.json"
    (root / "expected.json").write_text(oracle.read_text(encoding="utf-8"), encoding="utf-8")
    (root / MARKER).write_text(json.dumps({"synthetic": True, "scenario": args.scenario, "days": [d.strftime("%Y.%m.%d") for d in days]}, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"fixture": str(root), "scenario": args.scenario, "days": len(days), "synthetic_only": True}))


if __name__ == "__main__":
    main()
