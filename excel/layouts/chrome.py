"""Kopfleiste (Farbfläche Z. 1–3) und Schrittzeile der Leitfaden-Seiten.

Die anklickbaren Reiter und die Schritt-Leiste setzt navigation.py nach der Neuberechnung als Formen darüber.
"""
import re

from openpyxl.styles import Border
from openpyxl.utils import get_column_letter

from design_pro import ACC, TEAL, fill

NAV_MIN_PX = 1480  # Mindestbreite der Kopfleiste (Platz für alle Reiter)


def col_px(ws, col):
    d = ws.column_dimensions.get(get_column_letter(col))
    w = d.width if d is not None and d.customWidth and d.width else 8.43
    return int(w * 7 + 5)


def masthead(ws, title_text):
    """Kopfleiste als Farbfläche; Reiter und Logo-Schriftzug setzt navigation.py als Formen darüber."""
    brand = [c for c in ws[2] if c.value == "MM HOLDING"]
    last_col = brand[0].column if brand else ws.max_column
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= 2 <= mr.max_row:
            if brand and mr.min_col == brand[0].column:
                last_col = mr.max_col
            ws.unmerge_cells(str(mr))
    for c in ws[2]:
        c.value = None
        c.hyperlink = None
    px, col = 0, 1
    while col <= last_col or px < NAV_MIN_PX:
        px += col_px(ws, col)
        col += 1
    last_col = col - 1
    for r, h in ((1, 6), (2, 33), (3, 3)):
        ws.row_dimensions[r].height = h
        for cc in range(1, last_col + 1):
            c = ws.cell(r, cc)
            c.fill = fill(ACC if r == 3 else TEAL)
            c.border = Border()


def stepper_row(ws):
    """Leitfaden-Seiten: Punktreihe (●○○) durch freie Zeile für die Schritt-Leiste ersetzen."""
    if not re.match(r"S\d\d ", ws.title):
        return
    for c in ws[8]:
        if isinstance(c.value, str) and set(c.value) <= set("●○"):
            c.value = None
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 10




def apply(wb):
    for ws in wb.worksheets:
        if ws.title == "Dashboard":
            continue
        masthead(ws, ws.title)
        stepper_row(ws)
