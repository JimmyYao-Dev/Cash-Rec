"""Exercise both actual XML consumers without requiring application code or Excel."""
import unittest
from xml.etree import ElementTree as ET

from xlsx_xml import numeric_cells, validate_compatibility_prefixes

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SOURCE = b'''<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="urn:synthetic:x14ac" xmlns:xr="urn:synthetic:xr" mc:Ignorable="x14ac xr"><x:sheetData><x:row r="3"><x:c r="A3" s="2" t="str"><x:f>B3-(-2)</x:f><x:v>0</x:v><x:extLst/></x:c><x:c r="B3" s="4"><x:v>1</x:v></x:c><x:c r="E3"/></x:row></x:sheetData><mc:AlternateContent><mc:Choice xmlns:local="urn:synthetic:local" Requires="local"/></mc:AlternateContent></x:worksheet>'''


class WorksheetPatchTests(unittest.TestCase):
    def test_literal_replaces_formula_preserves_style_and_prefix_scopes(self):
        result = numeric_cells(SOURCE, {"A3": 666})
        cell = ET.fromstring(result).find(f".//{S}c[@r='A3']")
        self.assertEqual(cell.attrib, {"r": "A3", "s": "2"})
        self.assertIsNone(cell.find(S + "f"))
        self.assertEqual(cell.find(S + "v").text, "666")
        self.assertIn(b'xmlns:x14ac="urn:synthetic:x14ac"', result)
        self.assertIn(b'xmlns:local="urn:synthetic:local" Requires="local"', result)
        validate_compatibility_prefixes(result)

    def test_cache_retains_exact_formula_style_and_other_cell_content(self):
        result = numeric_cells(SOURCE, {"A3": -12.5}, formula_cache=True)
        cell = ET.fromstring(result).find(f".//{S}c[@r='A3']")
        self.assertIn(b'<x:f>B3-(-2)</x:f>', result)
        self.assertEqual(cell.attrib, {"r": "A3", "s": "2"})
        self.assertEqual(cell.find(S + "v").text, "-12.5")
        self.assertEqual([node.tag for node in cell], [S + "f", S + "v", S + "extLst"])
        self.assertIn(b'<x:c r="B3" s="4"><x:v>1</x:v></x:c>', result)
        validate_compatibility_prefixes(result)

    def test_new_literal_cell_is_inserted_in_column_order(self):
        result = numeric_cells(SOURCE, {"D3": 17, "C3": 0})
        self.assertEqual([cell.attrib["r"] for cell in ET.fromstring(result).find(f".//{S}row")], ["A3", "B3", "C3", "D3", "E3"])

    def test_invalid_prefix_and_nonformula_cache_are_explicit_errors(self):
        with self.assertRaises(ValueError):
            numeric_cells(SOURCE.replace(b' xmlns:xr="urn:synthetic:xr"', b""), {"A3": 1})
        with self.assertRaises(ValueError):
            numeric_cells(SOURCE, {"B3": 1}, formula_cache=True)


if __name__ == "__main__":
    unittest.main()
