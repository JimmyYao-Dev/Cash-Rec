"""Read actual installed range-GUI diagnostics against an independent day contract."""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path


def day(value) -> str:
    return str(value).replace(".", "-")[:10]


def codes(records) -> set[str]:
    return {record["Code"] for record in records if record.get("Code")}


def verify(evidence: Path, expected: Path) -> dict:
    ui = json.loads(evidence.read_text(encoding="utf-8-sig")); oracle = json.loads(expected.read_text(encoding="utf-8"))
    if not ui.get("Passed") or not oracle.get("synthetic_only"):
        raise AssertionError("Matching installed real range-GUI evidence and a synthetic oracle are required.")
    batch = ui["Batch"]
    if not batch["Batch"] or batch["Operation"] != oracle["operation"]:
        raise AssertionError("The UI did not invoke a batch generation operation.")
    if day(batch["StartDate"]) != oracle["start_date"] or day(batch["EndDate"]) != oracle["end_date"]:
        raise AssertionError("The batch substituted a different date range.")
    if batch["Status"] != oracle["status"] or bool(batch["Complete"]) != (oracle["status"] == "Success"):
        raise AssertionError("The range status or overall completion contradicts its independent expectation.")
    actual = batch["Days"]; wanted = oracle["days"]
    if [day(row["Date"]) for row in actual] != [row["date"] for row in wanted]:
        raise AssertionError("The batch did not preserve every inclusive date once in ascending order.")
    counts = Counter(row["Status"] for row in actual)
    for status in ["Success", "Missing", "Failed", "Partial"]:
        if batch["Counts"].get(status, 0) != counts.get(status, 0):
            raise AssertionError("The batch status counts do not match its actual per-day results.")
    if oracle.get("required_batch_code") not in (None, *codes(batch.get("Diagnostics", []))):
        raise AssertionError("The invalid-range diagnostic code was omitted.")
    output_paths = set()
    for record, target in zip(actual, wanted):
        if record["Status"] != target["status"]:
            raise AssertionError("Wrong per-day status: " + target["date"])
        if "blocked" in target and bool(record["Blocked"]) != target["blocked"]:
            raise AssertionError("Unavailable prior basis was counted as an unblocked success.")
        if target.get("required_code") not in (None, *codes(record.get("Diagnostics", []))):
            raise AssertionError("The per-day missing-input/basis reason is not explicitly coded.")
        for field in ["DailyOutput", "MonthlyOutput"]:
            value = record.get(field)
            if value:
                path = Path(value)
                if not path.is_file(): raise AssertionError("A published per-day output does not exist.")
                output_paths.add(str(path.resolve()).casefold())
                if field == "DailyOutput" and path.parent.name != target["date"].replace("-", "."):
                    raise AssertionError("A daily output was stored under a neighboring business date.")
        if target["status"] in {"Success", "Partial"}:
            if not record.get("DailyOutput") or not record.get("MonthlyOutput"):
                raise AssertionError("A completed/partial day lost its usable report locations.")
            diagnostic = Path(record["DiagnosticPath"])
            if not diagnostic.is_file(): raise AssertionError("The generated day lacks its task diagnostic.")
            task = json.loads(diagnostic.read_text(encoding="utf-8-sig"))
            if bool(task["Complete"]) != (target["status"] == "Success"):
                raise AssertionError("Per-day batch status disagrees with the installed single-task result.")
        elif record.get("DailyOutput") or record.get("MonthlyOutput"):
            raise AssertionError("A missing/failed day was falsely published as an output.")
    published = {str(Path(path).resolve()).casefold() for path in batch["PublishedOutputs"]}
    if not output_paths.issubset(published) or any(not Path(path).is_file() for path in batch["PublishedOutputs"]):
        raise AssertionError("The batch output list hides usable outputs or points to missing files.")
    if not wanted and published:
        raise AssertionError("A reversed range published outputs despite executing no days.")
    if not Path(batch["ResultsDirectory"]).is_dir():
        raise AssertionError("The range's result folder cannot be opened.")
    return {"passed": True, "entry": "Independent read of installed real range-GUI diagnostics", "dates": [row["date"] for row in wanted],
            "status": batch["Status"], "counts": dict(counts), "inclusive_ascending_dates_verified": True,
            "published_main_output_count": len(output_paths), "all_usable_outputs_exist": True}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ui-evidence", type=Path, required=True); parser.add_argument("--expected", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True); args = parser.parse_args()
    try: result = verify(args.ui_evidence, args.expected); status = 0
    except Exception as error: result = {"passed": False, "error": str(error)}; status = 1
    args.report.parent.mkdir(parents=True, exist_ok=True); args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False)); raise SystemExit(status)
