"""Kopfleiste (Farbfläche Z. 1–3), Monogramm-Entfernung und Zeilen für Schritt- und Sprungleisten.

Die anklickbaren Reiter, die Schritt-Leiste und die Sprungleisten setzt navigation.py nach der Neuberechnung
als Formen darüber. Hier wird nur die Zellfläche vorbereitet (Farben, Zeilenhöhen) – keine Werte, keine Formeln.
"""
import re

from openpyxl.styles import Border

from core import ACCENT, NAVY, col_px, fill

# Die Reiterleiste endet am Inhaltsrand der Schritt-Seiten (Ende Spalte I ≈ 1 354 px, navigation.nav_frame).
# Die Navy-Fläche reicht mindestens bis zur nächsten Spaltengrenze dahinter (Schritt-Seiten: Ende Spalte J).
NAV_MIN_PX = 1368
JUMP_ROW_SHEETS = ("Eingaben", "Diagramme")  # Zeile 8 trägt die Sprungleiste (navigation.jump_bar)


def masthead(ws):
    """Kopfleiste als Farbfläche: Z. 1 (6 pt) und 2 (33 pt) Navy, Z. 3 (3 pt) Akzentlinie."""
    brand = [c for c in ws[2] if c.value == "MM HOLDING"]
    last_col = brand[0].column if brand else ws.max_column
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= 2 <= mr.max_row:
            if brand and mr.min_col == brand[0].column:
                last_col = mr.max_col
            ws.unmerge_cells(str(mr))
    for c in ws[2]:
        if not (isinstance(c.value, str) and c.value.startswith("=")):
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
            c.fill = fill(ACCENT if r == 3 else NAVY)
            c.border = Border()


def remove_monogram(ws):
    """P3-01: das kleine Monogramm-Bild in A1/A2 entfällt – die Marke „MM HOLDING“ (Form) genügt."""
    keep = []
    for img in ws._images:
        a = getattr(img.anchor, "_from", None)
        small = (getattr(img, "width", 0) or 0) <= 64 and (getattr(img, "height", 0) or 0) <= 64
        if a is not None and a.col == 0 and a.row <= 1 and small:
            continue
        keep.append(img)
    ws._images = keep


def stepper_row(ws):
    """Leitfaden-Seiten: Punktreihe (●○○) entfernen, Zeile 8 als Schritt-Leiste (30 pt), Zeile 9 Abstand."""
    if not re.match(r"S\d\d ", ws.title):
        return
    for c in ws[8]:
        if isinstance(c.value, str) and set(c.value) <= set("●○ "):
            c.value = None
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = 10


def jump_row(ws):
    """Eingaben/Diagramme: Zeile 8 als ruhige Zeile für die Sprungleiste (22 px Reiter, mittig)."""
    if ws.title not in JUMP_ROW_SHEETS:
        return
    if all(c.value is None for c in ws[8]):
        ws.row_dimensions[8].height = 30


def apply(wb):
    for ws in wb.worksheets:
        remove_monogram(ws)
        if ws.title == "Dashboard":
            continue
        masthead(ws)
        stepper_row(ws)
        jump_row(ws)
