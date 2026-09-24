"""Kopfleiste (Farbfläche Z. 1–3), Monogramm-Entfernung und Zeilen für Schritt- und Sprungleisten.

Die anklickbaren Reiter, die Schritt-Leiste und die Sprungleisten setzt navigation.py nach der Neuberechnung
als Formen darüber. Hier wird nur die Zellfläche vorbereitet (Farben, Zeilenhöhen) – keine Werte, keine Formeln.
"""
import re

from openpyxl.styles import Border

from openpyxl.utils import column_index_from_string, get_column_letter

from core import ACCENT, CONTENT_EDGE, H_GAP, NAV_MIN_PX, NAVY, NOFILL, col_px, fill

# Die Reiter (Formen, navigation.py) enden auf jedem Blatt bei 1 200 px; die Navy-Fläche der Kopfleiste reicht
# mindestens bis core.NAV_MIN_PX (1 368 px, hinter dem Inhaltsrand der Schritt-Seiten) bzw. bis zur letzten
# Inhaltsspalte breiterer Blätter (core.CONTENT_EDGE – einzige Quelle, auch für navigation.py).
JUMP_ROW_SHEETS = ("Eingaben", "Diagramme")  # Zeile 8 trägt die Sprungleiste (navigation.jump_bar)
BAND_TO = CONTENT_EDGE                       # Kompatibilität für Altaufrufer


def band_end_col(ws):
    """Letzte Spalte der Kopfleiste (Zellfläche).

    - Mindestens alle Spalten, deren rechte Kante innerhalb von NAV_MIN_PX liegt; den Rest bis genau NAV_MIN_PX
      ergänzt navigation.band_extension als Fläche (so endet die Kopfleiste auf jedem Blatt an derselben Kante,
      auch wenn hinter dem Inhalt eine sehr breite Spalte folgt).
    - Breiter Inhalt (core.CONTENT_EDGE, Diagramme): bis zur letzten Inhaltsspalte – spaltenbasiert, weil die Blatt-Module
      die Spaltenbreiten erst nach diesem Modul setzen.
    Früher reichte die Fläche bis zur Spalte der alten Marke (z. B. Start bis U = 1 914 px bei 731 px Inhalt)."""
    last = column_index_from_string(CONTENT_EDGE[ws.title]) if ws.title in CONTENT_EDGE else 1
    for ch in getattr(ws, "_charts", []):
        to = getattr(ch.anchor, "to", None)
        if to is not None:
            last = max(last, to.col + (1 if to.colOff else 0))
    px, c = 0, 1
    while c < 400:
        w = col_px(ws, get_column_letter(c))          # ausgeblendete Spalten = 0 px (core.col_px)
        if px + w > NAV_MIN_PX + 6:
            break
        px += w
        c += 1
    return max(last, c - 1)


def masthead(ws):
    """Kopfleiste als Farbfläche: Z. 1 (6 pt) und 2 (33 pt) Navy, Z. 3 (3 pt) Akzentlinie – so breit wie
    Reiterleiste bzw. Inhalt; dahinter bleibt die Kopfzone weiß."""
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= 2 <= mr.max_row:
            ws.unmerge_cells(str(mr))
    for c in ws[2]:
        if not (isinstance(c.value, str) and c.value.startswith("=")):
            c.value = None
        c.hyperlink = None
    last_col = band_end_col(ws)
    for r, h in ((1, 6), (2, 33), (3, 3)):
        ws.row_dimensions[r].height = h
        for cc in range(1, max(last_col, ws.max_column) + 1):
            c = ws.cell(r, cc)
            c.fill = fill(ACCENT if r == 3 else NAVY) if cc <= last_col else NOFILL
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
    """Leitfaden-Seiten: Punktreihe (●○○) entfernen, Zeile 8 als Schritt-Leiste (30 pt), Zeile 9 = eine
    Standard-Fuge (core.H_GAP) bis zum ersten Abschnittskopf."""
    if not re.match(r"S\d\d ", ws.title):
        return
    for c in ws[8]:
        if isinstance(c.value, str) and set(c.value) <= set("●○ "):
            c.value = None
    ws.row_dimensions[8].height = 30
    ws.row_dimensions[9].height = H_GAP


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
