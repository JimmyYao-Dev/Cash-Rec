"""Patch small synthetic worksheet fragments without rewriting namespace scopes."""
from __future__ import annotations

import io
import re
from decimal import Decimal
from xml.etree import ElementTree as ET

PREFIX = r"(?:[A-Za-z_][\w.-]*:)?"
SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def validate_compatibility_prefixes(data: bytes) -> None:
    pending, scopes = [], [{"xml": "http://www.w3.org/XML/1998/namespace"}]
    compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    for event, node in ET.iterparse(io.BytesIO(data), events=("start-ns", "start", "end")):
        if event == "start-ns":
            pending.append(node)
        elif event == "start":
            scope = {**scopes[-1], **dict(pending)}; pending.clear(); scopes.append(scope)
            names = node.attrib.get("{" + compatibility + "}Ignorable", "").split()
            if node.tag == "{" + compatibility + "}Choice":
                names += node.attrib.get("Requires", "").split()
            if any(name not in scope for name in names):
                raise ValueError("Synthetic XML preparation lost a compatibility namespace prefix declaration.")
        else:
            scopes.pop()


def _element(tag: str, address: str) -> re.Pattern:
    return re.compile(rf'<(?P<prefix>{PREFIX}){tag}\b(?=[^>]*\br\s*=\s*["\']{re.escape(address)}["\'])'
                      rf'(?P<attrs>[^>]*?)(?:\s*/>|>(?P<body>.*?)</(?P=prefix){tag}\s*>)', re.DOTALL)


def _column(address: str) -> int:
    column = 0
    for letter in re.match(r"[A-Z]+", address).group():
        column = column * 26 + ord(letter) - ord("A") + 1
    return column


def numeric_cells(data: bytes, values: dict[str, int | float | Decimal], *, formula_cache: bool = False) -> bytes:
    """Literal writes clear formulas; cache writes preserve the exact formula/style XML."""
    validate_compatibility_prefixes(data)
    document = ET.fromstring(data)
    cells = {cell.attrib["r"]: cell for cell in document.findall(f".//{{{SPREADSHEET}}}sheetData/{{{SPREADSHEET}}}row/{{{SPREADSHEET}}}c")}
    text = data.decode("utf-8-sig")
    for address, value in values.items():
        if not re.fullmatch(r"[A-Z]+[1-9][0-9]*", address) or isinstance(value, bool) or not Decimal(str(value)).is_finite():
            raise ValueError("Synthetic numeric cell address/value is invalid.")
        original = cells.get(address)
        if formula_cache and (original is None or original.find(f"{{{SPREADSHEET}}}f") is None):
            raise ValueError("Cache target must be an existing formula cell: " + address)
        matches = list(_element("c", address).finditer(text))
        if len(matches) > 1:
            raise ValueError("Duplicate synthetic cell address: " + address)
        if matches:
            match = matches[0]; prefix = match.group("prefix")
            attrs = re.sub(r'\s+t\s*=\s*(?:"[^"]*"|\'[^\']*\')', "", match.group("attrs"))
            body = match.group("body") or ""
            number = f"<{prefix}v>{value}</{prefix}v>"
            if formula_cache:
                body = re.sub(rf'<{re.escape(prefix)}v\b[^>]*(?:/>|>.*?</{re.escape(prefix)}v\s*>)', "", body, flags=re.DOTALL)
                formula = re.search(rf'<{re.escape(prefix)}f\b[^>]*(?:/>|>.*?</{re.escape(prefix)}f\s*>)', body, flags=re.DOTALL)
                if formula is None:
                    raise ValueError("Synthetic formula prefix differs from its cell prefix: " + address)
                body = body[:formula.end()] + number + body[formula.end():]
            else:
                body = number
            replacement = f"<{prefix}c{attrs}>{body}</{prefix}c>"
            text = text[:match.start()] + replacement + text[match.end():]
        else:
            row_number = re.search(r"[0-9]+", address).group()
            row = _element("row", row_number).search(text)
            if row is None:
                raise ValueError("Synthetic business row is missing: " + address)
            prefix = row.group("prefix"); body = row.group("body") or ""
            cell = f'<{prefix}c r="{address}"><{prefix}v>{value}</{prefix}v></{prefix}c>'
            later = next((match for match in re.finditer(rf'<{PREFIX}c\b[^>]*\br=["\']([A-Z]+[0-9]+)["\']', body)
                          if _column(match.group(1)) > _column(address)), None)
            position = later.start() if later else len(body)
            body = body[:position] + cell + body[position:]
            text = text[:row.start()] + f'<{prefix}row{row.group("attrs")}>{body}</{prefix}row>' + text[row.end():]
    result = text.encode("utf-8")
    validate_compatibility_prefixes(result)
    return result
