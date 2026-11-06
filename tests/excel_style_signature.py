"""Semantic signature of visible/manual cell styles; never compares workbook-local IDs."""


def color_signature(color):
    if color is None:
        return None
    return {"type": color.type, "value": color.value, "tint": color.tint}


def cell_style_signature(cell):
    return {"font_name": cell.font.name, "font_size": cell.font.sz, "bold": cell.font.b,
            "italic": cell.font.i, "font_color": color_signature(cell.font.color),
            "fill_pattern": cell.fill.patternType, "fill_color": color_signature(cell.fill.fgColor),
            "bottom_border": cell.border.bottom.style, "bottom_color": color_signature(cell.border.bottom.color),
            "horizontal": cell.alignment.horizontal, "vertical": cell.alignment.vertical,
            "wrap_text": cell.alignment.wrap_text, "locked": cell.protection.locked,
            "hidden": cell.protection.hidden, "number_format": cell.number_format}
