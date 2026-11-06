"""Prepare only marked synthetic rerun input; never invoke application code."""
from __future__ import annotations

import argparse
import io
import json
import posixpath
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from xlsx_xml import numeric_cells


def prepare(edited: Path, workspace: Path, business_date: str) -> dict:
    marker=workspace/".cashrec-monthly-test-fixture.json"
    if not marker.exists() or not json.loads(marker.read_text())["synthetic_only"]:
        raise ValueError("Destination must be an explicitly marked synthetic workspace.")
    setup=json.loads((edited.parent/"excel-monthly-result.json").read_text(encoding="utf-8-sig"))
    original=Path(setup["Workbook"])
    if not setup["Passed"] or not setup["OwnNewInstance"] or Path(setup["EditedCopy"]).resolve()!=edited:
        raise ValueError("Source must be a successful owned-Excel edited evidence copy.")
    if not (original.parent.parent/".cashrec-monthly-test-fixture.json").exists():
        raise ValueError("Source evidence must originate from a marked synthetic workspace.")
    oracle_path=workspace/"monthly_expected_comment_policy.json"
    oracle=json.loads(oracle_path.read_text(encoding="utf-8"))
    book=load_workbook(edited,data_only=False)
    rows={}
    for name in ["TMA","TMG","Total","Break down"]:
        sheet=book[name]
        rows[name]=next(r for r in range(2,sheet.max_row+1) if str(sheet.cell(r,1).value).startswith(business_date))
    row=rows["Break down"]
    assert [book["Break down"][f"{c}{row}"].value for c in ["G","L","M","N","O"]]==[-5,0,None,-2.5,None]
    oracle["manual_values"]={"To LP":-5,"Bank Fee":0,"Bank fee-conversion":None,"Time Diff":-2.5}
    oracle["breakdown_number_formats"]={c:book["Break down"][f"{c}{row}"].number_format for c in ["G","L","M","N","O","P"]}
    oracle["breakdown"]["Check"]=(oracle["breakdown"]["Diff"]-oracle["settle_pnl"]["TMA"]
        -oracle["settle_pnl"]["TMG"]-oracle["breakdown"]["To Settle (WS)"]-oracle["breakdown"]["To Commission"]+7.5)
    book.close()
    cells={name:{f"{c}{rows[name]}":value for c,value in [("Q",666),("S",777),("U",888)]} for name in ["TMA","TMG","Total"]}
    cells["Break down"]={f"O{row}":17}
    ns={"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    target=workspace/"monthly"/("CashRec_Breakdown_"+datetime.fromisoformat(business_date).strftime("%Y.%m")+".xlsx")
    if target.resolve()==edited:raise ValueError("Edited evidence copy must be preserved.")
    target.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(io.BytesIO(edited.read_bytes())) as archive:
        wb=ET.fromstring(archive.read("xl/workbook.xml"));rels=ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        parts={}
        for node in wb.findall("s:sheets/s:sheet",ns):
            rel=node.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            path=next(r.attrib["Target"] for r in rels if r.attrib["Id"]==rel)
            part=path.lstrip("/") if path.startswith("/") else posixpath.normpath("xl/"+path)
            parts[part]=node.attrib["name"]
        with ZipFile(target,"w") as output:
            for item in archive.infolist():
                data=archive.read(item.filename);name=parts.get(item.filename)
                if name in cells:
                    data=numeric_cells(data,cells[name])
                output.writestr(item,data)
    oracle_path.write_text(json.dumps(oracle,indent=2)+"\n",encoding="utf-8")
    result={"synthetic_only":True,"precondition_only":True,"source_owned_excel_copy":str(edited),"destination":str(target),"stale_cells":cells,"manual_values":oracle["manual_values"],"business_acceptance_executed":False}
    (workspace/"stale-cells-precondition.json").write_text(json.dumps(result,indent=2),encoding="utf-8")
    return result


if __name__=="__main__":
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--edited-workbook",type=Path,required=True);parser.add_argument("--workspace",type=Path,required=True)
    parser.add_argument("--business-date",default="2026-10-25");args=parser.parse_args()
    print(json.dumps(prepare(args.edited_workbook.resolve(),args.workspace.resolve(),args.business_date),indent=2))
