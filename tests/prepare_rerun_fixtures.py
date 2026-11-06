"""Prepare independent synthetic rerun inputs, never invoking application code."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
from pathlib import Path

from openpyxl import load_workbook

import generate_monthly_fixtures as monthly
from excel_style_signature import cell_style_signature


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def expected_22(workspace: Path) -> dict:
    result = read_json(workspace / "monthly_expected_comment_policy.json")
    result.update(business_date="2026-10-22", previous_date="2026-10-19")
    result["balances"] = {
        "TMA": {"Insto": 0, "Retail": 800, "Wholesale": 10, "PSP": 90, "PSP Pending": 80, "Settle": 80},
        "TMG": {"Insto": 10, "Retail": 0, "Wholesale": 350, "PSP": 110, "PSP Pending": 150, "Settle": 0},
        "Total": {"Insto": 10, "Retail": 800, "Wholesale": 360, "PSP": 200, "PSP Pending": 230, "Settle": 80},
    }
    result["diffs"] = {
        entity: {field: (30 if entity != "Total" else 60) if field == "PSP Pending" else 0
                 for field in ["Insto", "Retail", "Wholesale", "PSP", "PSP Pending", "Settle"]}
        for entity in ["TMA", "TMG", "Total"]
    }
    result["previous_balances"] = json.loads(json.dumps(result["balances"]))
    result["previous_balances"]["TMA"]["PSP Pending"] = 50
    result["previous_balances"]["TMG"]["PSP Pending"] = 120
    result["previous_balances"]["Total"]["PSP Pending"] = 170
    result["settle_pnl"] = {"TMA": 325, "TMG": -3, "Total": 322}
    result["breakdown"].update({"Settle Bal": 80, "Client Bank Move": 60, "Diff": -182,
                               "To settle pnl (TMA)": 325, "To Settle adj (TMG)": -3, "Check": -479})
    return result


def mutate(workspace: Path, setup: Path) -> dict:
    marker = read_json(workspace / monthly.MARKER)
    proof = read_json(setup)
    if not marker.get("synthetic_only") or not proof.get("Passed") or not proof.get("OwnNewInstance"):
        raise ValueError("Rerun mutation requires a marked fixture and successful actual owned-Excel setup.")
    if Path(proof["Workspace"]).resolve() != workspace:
        raise ValueError("The Excel setup belongs to a different synthetic workspace.")
    daily_path = workspace / "2026.10.25/CashRec_Base.xlsx"
    monthly_path = workspace / "monthly/CashRec_Breakdown_2026.10.xlsx"
    for path, field in [(daily_path, "DailySHA256"), (monthly_path, "MonthlySHA256")]:
        if hashlib.sha256(path.read_bytes()).hexdigest().upper() != proof[field]:
            raise ValueError("The saved Excel precondition changed before source mutation.")
    book = load_workbook(daily_path, data_only=False)
    sheet = book["Equity"]; headers = {cell.value: cell.column for cell in sheet[1]}
    saved_dep_formats = {"/".join(str(sheet.cell(r, headers[k]).value) for k in ["Branch", "Retail Type", "Currency"]):
        sheet.cell(r, headers["DEP"]).number_format for r in range(2, sheet.max_row + 1)}
    saved_dep_styles = {"/".join(str(sheet.cell(r, headers[k]).value) for k in ["Branch", "Retail Type", "Currency"]):
        cell_style_signature(sheet.cell(r, headers["DEP"])) for r in range(2, sheet.max_row + 1)}
    book.close()
    book = load_workbook(monthly_path, data_only=False)
    saved_monthly_formats = {book["Break down"].cell(r, 1).value.strftime("%Y-%m-%d"):
        {letter: book["Break down"][f"{letter}{r}"].number_format for letter in ["G", "L", "M", "N", "O", "P"]}
        for r in range(2, book["Break down"].max_row + 1) if hasattr(book["Break down"].cell(r, 1).value, "strftime")}
    saved_monthly_styles = {book["Break down"].cell(r, 1).value.strftime("%Y-%m-%d"):
        {field: cell_style_signature(book["Break down"].cell(r, column))
         for column, field in [(7, "To LP"), (12, "Bank Fee"), (13, "Bank fee-conversion"), (14, "Time Diff")]}
        for r in range(2, book["Break down"].max_row + 1) if hasattr(book["Break down"].cell(r, 1).value, "strftime")}
    book.close()
    # New IB/EUR key has an explicit zero prior amount, not an assumed missing-key zero.
    for day in ["2026.10.19", "2026.10.22", "2026.10.25"]:
        path = workspace / day / "Needed files/MT4/Risk/Accounting Equity Report.csv"
        with path.open(encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle); headers = reader.fieldnames; rows = list(reader)
        rows = [r for r in rows if r["login"] != "SYN-NEW-IB-EUR"]
        if day == "2026.10.25":
            rows = [r for r in rows if r["login"] != "SYN-E6"]
            next(r for r in rows if r["login"] == "SYN-E1")["positive_equity"] = "60"
            next(r for r in rows if r["login"] == "SYN-E4")["positive_equity"] = "150"
        added = dict(rows[0]); added.update(dt_report=day.replace(".", "-"), report_jurisdiction="VFSC",
            login="SYN-NEW-IB-EUR", is_ib="true", currency="EUR", positive_equity="12.5" if day.endswith("25") else "0")
        rows.append(added)
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers); writer.writeheader(); writer.writerows(reversed(rows))
    path = workspace / "2026.10.25/Needed files/Bank/NAB AU.csv"
    text = path.read_text(encoding="utf-8-sig")
    rows = list(csv.reader(text.splitlines()))
    header = next(i for i, row in enumerate(rows) if "Account Number" in row and "EOD balance" in row)
    account = rows[header].index("Account Number"); balance = rows[header].index("EOD balance")
    match = [r for r in rows[header+1:] if len(r) > balance and r[account].replace("-", "").replace(" ", "") == "0011001"]
    assert len(match) == 1
    match[0][balance] = "910"
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle).writerows(rows)
    path = workspace / "2026.10.25/Needed files/PSP+BANK/BANK BD.xlsx"
    book = load_workbook(path)
    changes = {"Synthetic TMA SETTLE": (21, None, 21), "Synthetic TMG SETTLE": (None, 11, -11), "PSP CN Commission": (None, 25, -25)}
    seen = set()
    for row in book["BANK BD"].iter_rows(min_row=2):
        if row[2].value in changes:
            comment = row[2].value; credit, debit, amount = changes[comment]
            row[0].value = credit; row[1].value = debit; row[6].value = amount; seen.add(comment)
    assert seen == changes.keys()
    book.save(path); book.close()
    # Saving a synthetic workbook with openpyxl does not calculate its existing Pending formulas.
    caches = {}
    pending = load_workbook(path, data_only=False)["PSP On The Way"]
    for row in pending.iter_rows(min_row=2):
        if row[19].data_type == "f":
            caches[f"T{row[0].row}"] = row[17].value - row[18].value
    monthly.daily.cache_formula_numbers(path, "PSP On The Way", caches)
    mapping = workspace / "db/tmc/account_mapping.xlsx"
    book = load_workbook(mapping)
    if not any(r[0].value == "PSPGL" and r[1].value == "NewGateway" and r[2].value == "EUR" for r in book["PSP MAPPING"].iter_rows(min_row=2)):
        book["PSP MAPPING"].append(["PSPGL", "NewGateway", "EUR", "TMC", "TMG", "PSP Balance Summary"])
    book.save(mapping); book.close()
    expected = read_json(workspace / "expected_comment_policy.json")
    equity = expected["sheets"]["Equity"]
    equity["rows"] = [r for r in equity["rows"] if not (r["Branch"] == "TMG" and r["Currency"] == "EUR")]
    manual = {r["Key"]: r for r in proof["ManualDEP"]}
    for row in equity["rows"]:
        key = "/".join(str(row[k]) for k in ["Branch", "Retail Type", "Currency"])
        row["DEP"] = manual[key]["Value"]; row["DEP Number Format"] = saved_dep_formats[key]
        row["DEP Cell Style"] = saved_dep_styles[key]
        if key == "TMA/Standard/USD": row.update(Amount=108, **{"USD Equivalent": 108, "Movement": 28})
        if key == "TMA/IB/AUD": row.update(Amount=150, **{"USD Equivalent": 120, "Movement": 40})
    equity["rows"].append({"Branch": "TMG", "Retail Type": "IB", "Currency": "EUR", "Amount": 12.5,
        "USD Equivalent": 10, "Prev USD Equivalent": 0, "Movement": 10, "DEP": 0})
    equity.update(total_usd=260, total_previous_usd=171, total_movement=89)
    bank = next(r for r in expected["sheets"]["Bank"]["rows"] if r["Account Number"] == "0011001")
    bank.update({"EOD Balance": 910, "USD Equivalent": 910, "Movement": 110})
    expected["sheets"]["Bank"].update(total_usd=1440, total_movement=190)
    transactions = expected["sheets"]["TransactionExchange"]
    for row in transactions["rows"]:
        if row["Comment"] in changes: row.update(Amount=changes[row["Comment"]][2], **{"USD Equivalent": changes[row["Comment"]][2]})
    transactions.update(total_usd=2662, source_totals_usd={"Transaction reports": 250, "BANK BD": 2412})
    next(r for r in expected["sheets"]["PSP"]["rows"] if r["PSP"] == "NewGateway")["Company"] = "TMC"
    write_json(workspace / "rerun_expected.json", expected)
    report = read_json(workspace / "monthly_expected_comment_policy.json")
    report["balances"]["TMA"]["Retail"] = 910; report["balances"]["Total"]["Retail"] = 910
    report["diffs"]["TMA"]["Retail"] = 110; report["diffs"]["Total"]["Retail"] = 110
    report["settle_pnl"] = {"TMA": 395, "TMG": -11, "Total": 384}
    report["breakdown"].update({"Client Bank Move": 293, "Diff": 51, "To settle pnl (TMA)": 395,
        "To Settle adj (TMG)": -11, "To Commission": -30, "Check": -295.5})
    report["manual_values"] = {"To LP": -5, "Bank Fee": 0, "Bank fee-conversion": None, "Time Diff": -2.5}
    report["breakdown_number_formats"] = saved_monthly_formats["2026-10-25"]
    report["manual_cell_styles"] = saved_monthly_styles["2026-10-25"]
    report["allow_sheet_reorder"] = True
    write_json(workspace / "rerun_monthly_expected.json", report)
    history = expected_22(workspace)
    history["manual_values"] = {"To LP": 0, "Bank Fee": -3, "Bank fee-conversion": None, "Time Diff": None}
    history["breakdown"]["Check"] = -476; history["breakdown_number_formats"] = saved_monthly_formats["2026-10-22"]
    history["manual_cell_styles"] = saved_monthly_styles["2026-10-22"]
    history["allow_sheet_reorder"] = True
    write_json(workspace / "rerun_history_expected.json", history)
    contract = {"synthetic_only": True, "business_acceptance_executed": False, "unmatched_expected": {
        "Date": "2026.10.25", "Sheet": "Equity", "KeyContains": ["TMG", "Standard", "EUR"], "Field": "DEP", "Value": 7,
        "NumberFormat": saved_dep_formats["TMG/Standard/EUR"]},
        "new_key": "TMG/IB/EUR", "new_key_DEP": 0, "expected_current_check": -295.5,
        "expected_retained_history_check": -476, "mutation": "same-date automatic changes plus Equity reorder/add/remove and PSP mapping completion"}
    write_json(workspace / "rerun_contract.json", contract)
    return contract


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--phase", choices=["init", "mutate"], required=True)
    parser.add_argument("--excel-setup", type=Path)
    args = parser.parse_args(); workspace = args.workspace.resolve()
    if args.phase == "init":
        monthly.generate(workspace, "good"); write_json(workspace / "initial_22_expected.json", expected_22(workspace))
        result = {"synthetic_only": True, "business_acceptance_executed": False, "workspace": str(workspace)}
    else:
        if args.excel_setup is None: parser.error("--excel-setup is required for mutate")
        result = mutate(workspace, args.excel_setup)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
