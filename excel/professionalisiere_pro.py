"""Überträgt das Design-System aus „Cockpit – Entwurf v4“ auf das Immobilien-Kalkulationstool Pro.

Aufruf:  python excel/professionalisiere_pro.py QUELLE.xlsx ZIEL.xlsx

Rechenlogik, Formeln, Namen, Blattschutz, Datenüberprüfungen und Diagramme bleiben
unverändert – es werden ausschließlich Schriften, Farben, Rahmen, Eingabe-Markierung,
Reiterfarben, Diagrammfarben und das Logo vereinheitlicht.
"""
import io
import re
import sys
import zipfile
from copy import copy

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font, PatternFill, Side
from PIL import Image

# --------------------------------------------------------------------------- Design-Tokens (Entwurf v4)
TEAL, TEAL_MID, ACC = "0E4A4F", "2E7F80", "4FA3A5"
INK, MUTED, MUTED2 = "1A1D21", "5B6068", "9AA0A6"
ON_DARK, ON_DARK_2 = "FFFFFF", "CFE3E3"
RED, AMB, GRN, BLUE = "B42318", "B54708", "1F7A4D", "2F6FB0"
INPUT_BG, INPUT_LINE = "EAF1FA", "B9CFE8"

FILL_MAP = {
    "F4F0E8": "FFFFFF",   # beiger Seitenhintergrund → weiß
    "0F1C16": TEAL,       # Banner, Tabellenköpfe
    "EBE6DB": "E6F0F0",   # Zwischenüberschriften, Summenzeilen
    "E2DCCE": "F2F6F6",   # Sekundär-Schaltflächen, Kopfzeilen
    "A08A55": TEAL_MID,   # Primär-Schaltflächen / Akzentleisten
    "FFFFFF": "FFFFFF",
}
FONT_MAP = {
    "0C0F0D": INK, "0F1C16": TEAL, "173026": TEAL, "6B685E": MUTED, "B9B4A6": MUTED2,
    "F4F0E8": ON_DARK, "EBE6DB": ON_DARK_2, "FFFFFF": ON_DARK,
    "4E7A62": TEAL_MID,   # Verknüpfungen
    "A08A55": TEAL_MID,   # Akzent-/Eyebrow-Texte
    "C6A45C": AMB,        # Hinweis/Warnung bzw. Label auf dunklem Grund
    "6F8F7A": GRN,        # erledigt / ✓
    "8C3B2E": RED,        # kritisch
}
# Schriftfarben auf dunklem Hintergrund (Teal)
ON_DARK_MAP = {INK: ON_DARK, TEAL: ON_DARK, MUTED: ON_DARK_2, MUTED2: ON_DARK_2, TEAL_MID: ON_DARK_2,
               AMB: "F5C98B", GRN: "9FD8B8", RED: "F4B3A8"}
BORDER_MAP = {
    ("hair", "D6D0C2"): ("hair", "E1E3E6"), ("thin", "D6D0C2"): ("thin", "E1E3E6"), ("dashed", "D6D0C2"): ("dashed", "DADCDF"),
    ("hair", "B9B4A6"): ("hair", "C9CDD2"), ("thin", "0F1C16"): ("thin", TEAL),
    ("thin", "A08A55"): ("thin", ACC), ("hair", "A08A55"): ("hair", ACC), ("medium", "A08A55"): ("medium", TEAL),
}
# Ampel-Schriftfarben der bedingten Formatierung (Hellvarianten stehen nur auf dunklen Kacheln)
CF_FONT_MAP = {"9DB4A6": "9FD8B8", "E0A898": "F4B3A8", "C6A45C": "F5C98B", "4E7A62": GRN, "8C3B2E": RED, "8C7748": AMB}
FONT_NAME_MAP = {"Inter": "Aptos", "Calibri": "Aptos", "Fraunces": "Aptos Display", "IBM Plex Mono": "Consolas"}
DARK_FILLS = {TEAL, TEAL_MID}
HINT_CELLS = []

GROUP = {"ein": "C4C8CE", "inp": BLUE, "aus": TEAL, "ber": "6E9C9E", "bank": "7F9BB8", "anh": "DADCDF"}
TAB_GROUP = {"Start": "ein", "Leitfaden": "ein", "Cockpit": "aus", "Diagramme": "aus", "Eingaben": "inp",
             "Steuern": "ber", "Projektion": "ber", "Finanzierung": "ber", "AfA-Vergleich": "ber", "Sensitivität": "ber",
             "Bankgespräch": "bank", "Haushaltsrechnung": "bank", "Vermögensaufstellung": "bank",
             "Hinweise": "anh", "Konfiguration": "anh"}

CHART_COLOR_MAP = {
    "0f1c16": TEAL, "173026": TEAL_MID, "a08a55": ACC, "c6a45c": BLUE, "6f8f7a": "6E9C9E",
    "b9b4a6": "A9B8C6", "e2dcce": "EDEEF0", "6b685e": MUTED, "8c3b2e": RED, "4f81bd": BLUE, "7f9aa6": "7F9BB8",
}


def rgb6(color):
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return None
    return color.rgb[-6:].upper()


def restyle_cell(c):
    fill_new = None
    if c.fill is not None and c.fill.fill_type == "solid":
        old = rgb6(c.fill.fgColor)
        if old in FILL_MAP:
            fill_new = FILL_MAP[old]
            c.fill = PatternFill("solid", start_color=fill_new, end_color=fill_new)

    f = c.font
    name = FONT_NAME_MAP.get(f.name, f.name)
    old_col = rgb6(f.color)
    col = FONT_MAP.get(old_col, old_col) if old_col else None
    if fill_new in DARK_FILLS:
        col = ON_DARK_2 if old_col == "C6A45C" else ON_DARK_MAP.get(col or INK, col)
    if old_col == "8C3B2E" and isinstance(c.value, str) and c.value.startswith("="):
        col = INK  # Hinweislisten neutral – Warnungen (⚠) färbt die bedingte Formatierung
        HINT_CELLS.append(c.coordinate)
    nf = copy(f)
    nf.name = name
    if col:
        nf.color = col
    c.font = nf

    b = c.border
    if any(s is not None and s.style for s in (b.left, b.right, b.top, b.bottom)):
        def side(s):
            if s is None or not s.style:
                return s
            key = (s.style, rgb6(s.color))
            st, colr = BORDER_MAP.get(key, (s.style, rgb6(s.color)))
            return Side(style=st, color=colr)
        c.border = Border(left=side(b.left), right=side(b.right), top=side(b.top), bottom=side(b.bottom),
                          diagonal=b.diagonal, diagonalUp=b.diagonalUp, diagonalDown=b.diagonalDown)

    # Eingabefelder (entsperrte Zellen) im einheitlichen Eingabestil
    if c.protection is not None and c.protection.locked is False:
        c.fill = PatternFill("solid", start_color=INPUT_BG, end_color=INPUT_BG)
        nf = copy(c.font)
        nf.color = BLUE
        nf.bold = True
        c.font = nf
        line = Side(style="thin", color=INPUT_LINE)
        c.border = Border(left=line, right=line, top=line, bottom=line)


def restyle_workbook(src, tmp):
    wb = load_workbook(src)
    for ws in wb.worksheets:
        for cf in list(ws.conditional_formatting):
            for rule in cf.rules:
                d = rule.dxf
                if d is not None and d.font is not None and d.font.color is not None:
                    old = rgb6(d.font.color)
                    if old in CF_FONT_MAP:
                        d.font.color = CF_FONT_MAP[old]
                if d is not None and d.fill is not None:
                    for attr in ("fgColor", "bgColor"):
                        colr = getattr(d.fill, attr)
                        old = rgb6(colr) if colr is not None else None
                        if old in FILL_MAP:
                            setattr(d.fill, attr, FILL_MAP[old])
        HINT_CELLS.clear()
        for c in list(ws._cells.values()):
            restyle_cell(c)
        for ref in HINT_CELLS:
            ws.conditional_formatting.add(ref, FormulaRule(formula=[f'LEFT({ref},1)="⚠"'], font=Font(color=RED, bold=True)))
        grp = TAB_GROUP.get(ws.title, "ein")
        ws.sheet_properties.tabColor = GROUP[grp]
        ws.sheet_view.showGridLines = False
    # Standardschrift für leere Zellen
    wb._named_styles["Normal"].font = Font(name="Aptos", size=11)
    wb.properties.title = "Immobilien-Kalkulation"
    wb.properties.creator = "MM Holding GmbH"
    wb.save(tmp)


def recolor_png(data):
    im = Image.open(io.BytesIO(data)).convert("RGBA")
    px = im.load()
    targets = [((15, 28, 22), (14, 74, 79)), ((160, 138, 85), (207, 227, 227)), ((198, 164, 92), (207, 227, 227)),
               ((244, 240, 232), (255, 255, 255))]
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            best = min(targets, key=lambda t: (t[0][0] - r) ** 2 + (t[0][1] - g) ** 2 + (t[0][2] - b) ** 2)
            dist = sum((best[0][i] - (r, g, b)[i]) ** 2 for i in range(3))
            if dist < 60 ** 2:
                px[x, y] = (*best[1], a)
    out = io.BytesIO()
    im.save(out, "PNG")
    return out.getvalue()


def postprocess(tmp, dst):
    """Diagrammfarben, Diagrammschrift, Theme-Schriften und Logo im Paket anpassen."""
    pat = re.compile(r'srgbClr val="([0-9A-Fa-f]{6})"')
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/charts/chart") and item.filename.endswith(".xml"):
                s = data.decode("utf-8")
                s = pat.sub(lambda m: f'srgbClr val="{CHART_COLOR_MAP.get(m.group(1).lower(), m.group(1))}"', s)
                s = s.replace('typeface="Inter"', 'typeface="Aptos"').replace('typeface="Arial"', 'typeface="Aptos"')
                data = s.encode("utf-8")
            elif item.filename == "xl/theme/theme1.xml":
                s = data.decode("utf-8")
                s = re.sub(r'(<a:majorFont><a:latin typeface=")[^"]*"', r'\1Aptos Display"', s)
                s = re.sub(r'(<a:minorFont><a:latin typeface=")[^"]*"', r'\1Aptos"', s)
                data = s.encode("utf-8")
            elif item.filename.startswith("xl/media/") and item.filename.endswith(".png"):
                data = recolor_png(data)
            zout.writestr(item, data)


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    tmp = dst + ".tmp.xlsx"
    restyle_workbook(src, tmp)
    postprocess(tmp, dst)
    import os
    os.remove(tmp)
    print(f"gespeichert: {dst}")
