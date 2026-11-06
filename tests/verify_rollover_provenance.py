"""Read installed-GUI rollover artifacts and prove the prior month stayed unchanged."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

from openpyxl import load_workbook


def normalized_date(value) -> str:
    return str(value).replace(".", "-")[:10]


def verify(workspace: Path, date: str, initialization: bool) -> dict:
    precondition = json.loads((workspace/"rollover_source_precondition.json").read_text(encoding="utf-8"))
    if not precondition.get("synthetic_only") or precondition["scenario"] != "good":
        raise AssertionError("A valid marked synthetic month-end source precondition is required.")
    previous = Path(precondition["prior_monthly"])
    digest = hashlib.sha256(previous.read_bytes()).hexdigest().upper()
    if digest != precondition["prior_monthly_sha256"]:
        raise AssertionError("The previous monthly file was overwritten or changed during rollover.")
    current = workspace/"monthly"/("CashRec_Breakdown_"+date[:7].replace("-", ".")+".xlsx")
    book = load_workbook(current, data_only=False)
    try:
        properties = {prop.name: prop.value for prop in book.custom_doc_props}
        if normalized_date(properties.get("CashRecMonthlyBaselineDate")) != "2026-10-31":
            raise AssertionError("The new monthly workbook does not identify the actual retained month-end baseline.")
        for sheet in book:
            for row in sheet.iter_rows(min_row=2 if sheet.title == "Break down" else 3):
                value = row[0].value
                if value is not None and not hasattr(value, "strftime"):
                    raise AssertionError("A trailing total or explanatory record was copied into the new monthly data.")
            if sheet.max_row >= 80:
                raise AssertionError("Trailing style-only/total rows were carried into the new monthly workbook.")
    finally:
        book.close()
    provenance_path = workspace/"monthly/.cashrec"/(date.replace("-", ".")+".json")
    document = json.loads(provenance_path.read_text(encoding="utf-8-sig"))
    records = [record for record in document["Provenance"] if record.get("Kind") == "Monthly rollover"]
    if len(records) != (1 if initialization else 0):
        raise AssertionError("The monthly file was initialized the wrong number of times for this task.")
    if records:
        record = records[0]
        if normalized_date(record["BaselineDate"]) != "2026-10-31" or Path(record["SourcePath"]).resolve() != previous.resolve():
            raise AssertionError("Rollover selected a different source file or business date.")
        if record["SourceSHA256"] != digest or not record["Selection"]:
            raise AssertionError("Rollover did not retain a verifiable source hash and selection reason.")
    return {"passed": True, "entry": "Read-only installed-GUI rollover artifacts", "prior_monthly_unchanged": True,
            "prior_monthly_sha256": digest, "retained_baseline": "2026-10-31", "initializations_this_task": len(records),
            "provenance": str(provenance_path), "trailing_nonbusiness_rows_not_copied": True}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--business-date", required=True)
    parser.add_argument("--expect-initialization", action="store_true")
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = verify(args.workspace.resolve(), args.business_date, args.expect_initialization); status = 0
    except Exception as error:
        result = {"passed": False, "error": str(error)}; status = 1
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False)); return status


if __name__ == "__main__":
    raise SystemExit(main())
