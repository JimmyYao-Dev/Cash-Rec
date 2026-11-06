"""Inspect GUI transaction evidence, recoverable backups and unmatched manual CSV."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
from decimal import Decimal
from pathlib import Path


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def verify(ui_path: Path, setup_path: Path, contract_path: Path) -> dict:
    ui, setup, contract = load(ui_path), load(setup_path), load(contract_path)
    if not ui.get("Passed") or not setup.get("Passed") or not setup.get("OwnNewInstance") or not contract.get("synthetic_only"):
        raise AssertionError("Successful real GUI and owned-Excel synthetic evidence is required.")
    run = ui["Run"]; transaction = run["Transaction"]
    if transaction["Status"] != "Committed" or not run["Complete"]:
        raise AssertionError("The normal rerun did not commit both complete outputs.")
    committed = {str(Path(path).resolve()).casefold() for path in transaction["CommittedOutputs"]}
    if not all(str(Path(run[field]).resolve()).casefold() in committed for field in ["DailyOutput", "MonthlyOutput"]):
        raise AssertionError("The successful transaction does not identify both committed workbooks.")
    backups = [{"path": path, "sha256": sha256(Path(path))} for path in transaction["Backups"]]
    hashes = {backup["sha256"] for backup in backups}
    if not {setup["DailySHA256"], setup["MonthlySHA256"]}.issubset(hashes):
        raise AssertionError("The pre-overwrite Excel-edited daily/monthly bytes are not both recoverable from backups.")
    unmatched = Path(run["UnmatchedManualPath"])
    if not unmatched.read_bytes().startswith(b"\xef\xbb\xbf"):
        raise AssertionError("The unmatched manual CSV must include a UTF-8 BOM for desktop Excel.")
    with unmatched.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle); records = list(reader)
        required = {"Date", "Sheet", "Key", "Field", "Value", "NumberFormat", "Reason"}
        if not required.issubset(reader.fieldnames or []):
            raise AssertionError("The unmatched manual file is missing traceable value/key/format fields.")
    if run["UnmatchedManualCount"] != 1 or len(records) != 1:
        raise AssertionError("Only the removed manual Equity key should be unmatched in this fixture.")
    row = records[0]; wanted = contract["unmatched_expected"]
    for field in ["Date", "Sheet", "Field", "NumberFormat"]:
        if row[field] != wanted[field]:
            raise AssertionError("The unmatched manual identity or original format differs: " + field)
    if not all(part in row["Key"] for part in wanted["KeyContains"]) or not row["Reason"]:
        raise AssertionError("The unmatched manual key or explanation is missing.")
    if Decimal(row["Value"]) != Decimal(str(wanted["Value"])):
        raise AssertionError("The removed manual DEP value was changed or lost.")
    return {"passed": True, "entry": "Read-only inspection of installed-GUI transaction outputs", "backups": backups,
            "unmatched_csv": str(unmatched), "unmatched_count": len(records), "removed_DEP": row["Value"],
            "old_excel_edited_output_hashes_recoverable": True, "transaction_status": transaction["Status"]}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ui-evidence", type=Path, required=True)
    parser.add_argument("--excel-setup", type=Path, required=True)
    parser.add_argument("--contract", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = verify(args.ui_evidence, args.excel_setup, args.contract); status = 0
    except Exception as error:
        result = {"passed": False, "error": str(error)}; status = 1
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False)); return status


if __name__ == "__main__":
    raise SystemExit(main())
