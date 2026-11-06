"""Synthetic monthly inputs; never imports application code or real business data."""
from __future__ import annotations

import argparse
import csv
import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

import generate_fixtures as daily

MARKER = ".cashrec-monthly-test-fixture.json"
DATES = [datetime(2026, 10, 19), daily.PREVIOUS, daily.CURRENT]
SCENARIOS = ("good", "no-previous", "pending-uncached", "pending-wrong-total", "pending-duplicate", "pending-missing-entity", "flow-source-collision", "flows-mixed", "comment-keywords")
PENDING = {
    "2026.10.19": {"TMA": (800, 750, 50), "TMG": (1700, 1580, 120)},
    "2026.10.22": {"TMA": (900, 820, 80), "TMG": (1800, 1650, 150)},
    "2026.10.25": {"TMA": (1000, 900, 100), "TMG": (2000, 1800, 200)},
}
SETTLEMENT_TYPES = ["SETTLE", "Posting TF TD", "MISC CIMB TF TD", "CFD TF TD", "Rebate TF TD", "Pending TF TD"]
PRIORITY_CASES = [
    ("Synthetic priority posting-fee", "tf fee", 301, "Bank Fee", "TMA"),
    ("Synthetic priority psp settle", "psp settle", 302, "PSP Settle In", "TMA"),
    ("Synthetic priority pending-withdrawal", "tf withdrawal", -303, "BW WD", "TMA"),
    ("Synthetic priority miscellaneous decoy", "unrelated", 304, "Other", "TMA"),
    ("Synthetic TMG CFD", "internal tf", 305, "CFD TF TD", "TMG"),
    ("Synthetic psp settlement override", "psp settle fee", 306, "PSP Settle In", "TMA"),
    ("Synthetic ordinary movement", "internal tf", 307, "Bank TF", "TMA"),
    ("Synthetic deposit withdrawal priority", "dep with", -308, "BW WD", "TMA"),
]
COMMENT_KEYWORD_CASES = [
    ("Keyword posting instruction", "memo", 2, "Posting TF TD", "TMA"),
    ("Keyword misc instruction", "memo", 3, "MISC CIMB TF TD", "TMA"),
    ("Keyword cfd instruction", "memo", 5, "CFD TF TD", "TMA"),
    ("Keyword rebate instruction", "memo", 7, "Rebate TF TD", "TMA"),
    ("Keyword pending instruction", "memo", 11, "Pending TF TD", "TMA"),
    ("Keyword SETTLE instruction", "memo", -9, "Settle", "TMA"),
    ("Keyword SETTLE TMG instruction", "memo", -13, "Settle", "TMG"),
    ("Keyword posting TMG excluded", "memo", 17, "Posting TF TD", "TMG"),
    ("Keyword posting fee override", "fee", 19, "Bank Fee", "TMA"),
    ("Keyword SETTLE PSP override", "psp settle", 23, "PSP Settle In", "TMA"),
]


def settlement_inputs(path: Path, day: datetime) -> None:
    book = load_workbook(path)
    sheet = book["BANK BD"]
    for index, category in enumerate(SETTLEMENT_TYPES, 1):
        amount = 10 + index if day == daily.CURRENT else index
        rational = "settlement payout" if category == "SETTLE" else "internal tf"
        sheet.append([amount, None, "Synthetic TMA " + category, rational, "USD", "TMA-Retail", amount])
    tmg_amount = -7 if day == daily.CURRENT else -3
    sheet.append([None, -tmg_amount, "Synthetic TMG SETTLE", "settlement payout", "USD", "TMG-Retail", tmg_amount])
    sheet.append([999, None, "Synthetic TMG excluded Posting", "internal tf", "USD", "TMG-Retail", 999])
    sheet.append([None, 8, "Synthetic unrelated retail", "Unrelated adjustment", "USD", "TMA-Retail", -8])
    for comment, rational, amount, expected_type, branch in PRIORITY_CASES:
        sheet.append([amount if amount >= 0 else None, -amount if amount < 0 else None,
                      comment, rational, "USD", branch + "-Retail", amount])
    book.save(path)


def flow_collision_inputs(needed: Path, mixed: bool = False) -> None:
    path = needed / "PSP+BANK/BANK BD.xlsx"
    book = load_workbook(path)
    sheet = book["BANK BD"]
    for row in sheet.iter_rows(min_row=2):
        if row[2].value == "Synthetic BANK BD PSP deposit":
            row[2].value = "Gateway-B deposit"
    sheet.append([0, None, "Zero BANK BD Wholesale", "Deposit", "USD", "TMG-Wholesale", 0])
    if mixed:
        sheet.append([8, None, "Wholesale positive offset", "Deposit", "USD", "TMA-Wholesale", 8])
        sheet.append([0, None, "Wholesale zero offset", "Deposit", "USD", "TMA-Wholesale", 0])
    book.save(path)
    path = needed / "MT4/MT4/Transaction Summary Report.csv"
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle))
    record = {"date":"2026-10-01", "user_id":"SYN-ZERO", "crm_transaction_no":"ZERO-1",
              "trading_server_transaction_no":"ZERO-MT4-1", "trading_server_type":"MT4",
              "trading_server":"Synthetic", "trading_server_login":"000000",
              "fund_amount":0, "trading_server_currency":"USD", "payment_amount":999,
              "payment_currency":"AUD", "exchange_rate":1, "surcharge_amount":0,
              "comment":"Zero PSP source collision", "accounting_comment":"Deposit",
              "jurisdiction":"Australian Securities and Investments Commission ASIC",
              "payment_channel":"Synthetic", "mapping_deal":"DECOY", "mapping_deal_currency":"EUR"}
    rows.append([record.get(header, "") for header in rows[0]])
    if mixed:
        for index, (comment, amount, kind) in enumerate([
            ("Duplicate same-amount PSP deposit", 100, "Deposit"),
            ("Additional PSP deposit", 5, "Deposit"),
            ("Zero PSP withdrawal", 0, "Withdrawal"),
            ("NAB synthetic withdrawal", -10, "Withdrawal"),
        ], 1):
            extra = dict(record)
            extra.update({"crm_transaction_no":f"FLOW-MIX-{index}", "trading_server_transaction_no":f"MIX-ID-{index}",
                          "fund_amount":amount, "comment":comment, "accounting_comment":kind})
            rows.append([extra.get(header, "") for header in rows[0]])
    daily.write_csv(path, rows[0], rows[1:])


def comment_keyword_inputs(path: Path) -> None:
    book=load_workbook(path)
    sheet=book["BANK BD"]
    for comment, rational, amount, expected_type, branch in COMMENT_KEYWORD_CASES:
        sheet.append([amount if amount >= 0 else None, -amount if amount < 0 else None,
                      comment, rational, "USD", branch + "-Retail", amount])
    book.save(path)


def write_comment_policy_oracles(root: Path) -> None:
    """Write expectations only; existing synthetic input bytes are never changed."""
    marker=json.loads((root/MARKER).read_text(encoding="utf-8"))
    if marker.get("synthetic_only") is not True:
        raise ValueError("Policy oracles require an explicitly marked synthetic workspace.")
    added=marker["scenario"] == "comment-keywords"
    monthly=json.loads((root/"monthly_expected.json").read_text(encoding="utf-8"))
    monthly["settle_pnl"]={"TMA":404 if added else 385,"TMG":-20 if added else -7,"Total":384 if added else 378}
    monthly["breakdown"].update({"To settle pnl (TMA)":monthly["settle_pnl"]["TMA"],"To Settle adj (TMG)":monthly["settle_pnl"]["TMG"]})
    monthly["breakdown"]["Check"]=monthly["breakdown"]["Diff"]-monthly["settle_pnl"]["TMA"]-monthly["settle_pnl"]["TMG"]+25
    monthly["policy"]="User-confirmed Comment keywords without Rational tf prerequisite"
    (root/"monthly_expected_comment_policy.json").write_text(json.dumps(monthly,indent=2)+"\n",encoding="utf-8")
    expected=json.loads((root/"expected.json").read_text(encoding="utf-8"))
    row=next(r for r in expected["sheets"]["TransactionExchange"]["rows"] if r["Comment"]=="Synthetic priority miscellaneous decoy")
    row["Transaction Type"]="MISC CIMB TF TD"
    (root/"expected_comment_policy.json").write_text(json.dumps(expected,indent=2)+"\n",encoding="utf-8")
    historical=root/"historical_expected.json"
    if historical.exists():
        old=json.loads(historical.read_text(encoding="utf-8"))
        old["settle_pnl"]={"TMA":325,"TMG":-3,"Total":322}
        (root/"historical_expected_comment_policy.json").write_text(json.dumps(old,indent=2)+"\n",encoding="utf-8")


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
        if scenario in {"flow-source-collision", "flows-mixed"}:
            flow_collision_inputs(needed, scenario == "flows-mixed")
        if scenario == "comment-keywords":
            comment_keyword_inputs(needed / "PSP+BANK/BANK BD.xlsx")
        pending_snapshot(needed / "PSP+BANK/BANK BD.xlsx", day, scenario)
    daily.write_mapping(root / "db/tmc/account_mapping.xlsx")
    oracle = Path(__file__).parent / "fixtures/monthly_expected.json"
    (root / "monthly_expected.json").write_text(oracle.read_text(encoding="utf-8"), encoding="utf-8")
    if scenario == "flows-mixed":
        monthly = json.loads(oracle.read_text(encoding="utf-8"))
        monthly["flows"] = {
            "TMA":{"PSP DEP":205,"PSP WITH":-30,"NET PSP":175,"NET BW":90,"NET WS":0},
            "TMG":{"PSP DEP":100,"PSP WITH":-20,"NET PSP":80,"NET BW":0,"NET WS":0},
            "Total":{"PSP DEP":305,"PSP WITH":-50,"NET PSP":255,"NET BW":90,"NET WS":0},
        }
        monthly["breakdown"].update({"MT4":345,"WS DEP":0,"Diff":-82,"Check":-131})
        (root / "monthly_expected.json").write_text(json.dumps(monthly, indent=2) + "\n", encoding="utf-8")
    expected = json.loads((Path(__file__).parent / "fixtures/expected.json").read_text(encoding="utf-8"))
    transactions = expected["sheets"]["TransactionExchange"]
    for index, category in enumerate(SETTLEMENT_TYPES, 1):
        transactions["rows"].append({"Source":"BANK BD", "Comment":"Synthetic TMA " + category,
            "Rational":"settlement payout" if category == "SETTLE" else "internal tf",
            "Transaction Type":"Settle" if category == "SETTLE" else category,
            "Branch":"TMA", "Type":"Retail", "Currency":"USD", "Amount":10 + index, "USD Equivalent":10 + index})
    transactions["rows"].extend([
        {"Source":"BANK BD", "Comment":"Synthetic TMG SETTLE", "Rational":"settlement payout", "Transaction Type":"Settle", "Branch":"TMG", "Type":"Retail", "Currency":"USD", "Amount":-7, "USD Equivalent":-7},
        {"Source":"BANK BD", "Comment":"Synthetic TMG excluded Posting", "Rational":"internal tf", "Transaction Type":"Posting TF TD", "Branch":"TMG", "Type":"Retail", "Currency":"USD", "Amount":999, "USD Equivalent":999},
        {"Source":"BANK BD", "Comment":"Synthetic unrelated retail", "Rational":"Unrelated adjustment", "Branch":"TMA", "Type":"Retail", "Currency":"USD", "Amount":-8, "USD Equivalent":-8},
    ])
    for comment, rational, amount, expected_type, branch in PRIORITY_CASES:
        transactions["rows"].append({"Source":"BANK BD", "Comment":comment, "Rational":rational,
            "Transaction Type":expected_type, "Branch":branch, "Type":"Retail", "Currency":"USD", "Amount":amount, "USD Equivalent":amount})
    transactions["total_usd"] = 2661
    transactions["source_totals_usd"]["BANK BD"] = 2411
    if scenario in {"flow-source-collision", "flows-mixed"}:
        row = next(row for row in transactions["rows"] if row["Source"] == "BANK BD" and row["Comment"] == "Synthetic BANK BD PSP deposit")
        row["Comment"] = "Gateway-B deposit"
        transactions["rows"].extend([
            {"Source":"Transaction reports", "Comment":"Zero PSP source collision", "Branch":"TMA", "Type":"PSP", "Transaction Type":"PSP Dep", "Currency":"USD", "Amount":0, "USD Equivalent":0},
            {"Source":"BANK BD", "Comment":"Zero BANK BD Wholesale", "Branch":"TMG", "Type":"Wholesale", "Transaction Type":"Deposit", "Currency":"USD", "Amount":0, "USD Equivalent":0},
        ])
        if scenario == "flows-mixed":
            for comment, amount, transaction_type, type_name in [
                ("Duplicate same-amount PSP deposit",100,"PSP Dep","PSP"),
                ("Additional PSP deposit",5,"PSP Dep","PSP"),
                ("Zero PSP withdrawal",0,"PSP WITH","PSP"),
                ("NAB synthetic withdrawal",-10,"BW WD","Retail"),
            ]:
                row={"Source":"Transaction reports", "Comment":comment, "Branch":"TMA", "Type":type_name,
                     "Transaction Type":transaction_type, "Currency":"USD", "Amount":amount, "USD Equivalent":amount}
                if type_name == "Retail":row["Exchange Type"]="MT4 BW"
                transactions["rows"].append(row)
            for comment, amount in [("Wholesale positive offset",8),("Wholesale zero offset",0)]:
                transactions["rows"].append({"Source":"BANK BD", "Comment":comment, "Branch":"TMA", "Type":"Wholesale",
                    "Transaction Type":"Deposit", "Currency":"USD", "Amount":amount, "USD Equivalent":amount})
            transactions["total_usd"] = 2764
            transactions["source_totals_usd"] = {"Transaction reports":345,"BANK BD":2419}
    if scenario == "comment-keywords":
        for comment,rational,amount,expected_type,branch in COMMENT_KEYWORD_CASES:
            transactions["rows"].append({"Source":"BANK BD","Comment":comment,"Rational":rational,
                "Transaction Type":expected_type,"Branch":branch,"Type":"Retail","Currency":"USD","Amount":amount,"USD Equivalent":amount})
        transactions["total_usd"]=2726
        transactions["source_totals_usd"]["BANK BD"]=2476
    (root / "expected.json").write_text(json.dumps(expected, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (root / MARKER).write_text(json.dumps({"synthetic_only": True, "scenario": scenario, "dates": [d.strftime("%Y.%m.%d") for d in days]}, indent=2) + "\n", encoding="utf-8")
    write_comment_policy_oracles(root)
    if scenario == "comment-keywords":
        for original,sidecar in [("monthly_expected.json","monthly_expected_comment_policy.json"),("expected.json","expected_comment_policy.json")]:
            (root/original).write_text((root/sidecar).read_text(encoding="utf-8"),encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--scenario", choices=SCENARIOS, default="good")
    parser.add_argument("--oracles-only", action="store_true", help="Write new-policy sidecar expectations without changing existing marked input files.")
    args = parser.parse_args()
    if args.oracles_only:
        write_comment_policy_oracles(args.output.resolve())
    else:
        generate(args.output.resolve(), args.scenario)
    actual_scenario=json.loads((args.output.resolve()/MARKER).read_text(encoding="utf-8"))["scenario"]
    print(json.dumps({"synthetic_only": True, "scenario": actual_scenario, "workspace": str(args.output.resolve()), "input_files_modified":not args.oracles_only, "business_acceptance_executed": False}))


if __name__ == "__main__":
    main()
