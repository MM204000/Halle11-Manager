"""Premium-Redesign des Immobilien-Kalkulationstools Pro (Design-System nach „Cockpit – Entwurf v4“).

Aufruf:  python excel/design_pro.py QUELLE.xlsx ZIEL.xlsx

Die Rechenlogik bleibt unangetastet: Formeln, Namen, Blattschutz, Datenüberprüfungen und
Diagrammdaten werden nicht verändert. Umgestaltet werden ausschließlich die Darstellung
(Kopfleiste, Typografie, Farben, Tabellen, Kacheln, Schaltflächen, Hinweise, Zahlenformate,
Diagramm-Styling, Logo) – komponentenweise anhand der Stilklassen der Vorlage.
"""
import io
import os
import re
import sys
import zipfile
from copy import copy

from lxml import etree
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Protection
from openpyxl.workbook.defined_name import DefinedName
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import importlib  # noqa: E402
import traceback  # noqa: E402

import dashboard  # noqa: E402

# Blatt-Layout-Module (excel/layouts/<name>.py, je Modul eine Funktion apply(wb)) in dieser Reihenfolge
LAYOUTS = ["chrome", "start", "leitfaden", "steps", "cockpit", "calc", "forms", "bank", "sensitivity", "diagramme"]
STRICT = os.environ.get("DESIGN_STRICT") == "1"
MODULE_ERRORS = []


def run_hook(label, fn, *args):
    try:
        return fn(*args)
    except Exception:  # Module werden parallel entwickelt – ein defektes Modul bricht den Build nur im STRICT-Modus
        MODULE_ERRORS.append(label)
        print(f"FEHLER in {label}:", file=sys.stderr)
        traceback.print_exc()
        if STRICT:
            raise

# ============================================================================ Design-Tokens / Farbthemen
# Rollen: TEAL = Primärfarbe (Banner, Titel), TEAL_MID = Sekundär, ACC = Akzent, TEAL_L/TEAL_XL = helle Flächen.
THEMES = {
    "blau": dict(
        TEAL="0B2A4A", TEAL_MID="1D4F8A", ACC="4A86C8", TEAL_L="E7EEF7", TEAL_XL="F3F7FC", SUM_BG="F3F7FC",
        ON_DARK_2="C8D7EB", ON_DARK_ACC="9CBBE2",
        INPUT_BG="FFF5D6", INPUT_LINE="E6CB77", INPUT_FG="1D4F8A",
        GROUP={"ein": "C4C8CE", "inp": "E6B940", "aus": "0B2A4A", "ber": "4A86C8", "bank": "8FA9C9", "anh": "DADCDF"},
        SERIES=["0B2A4A", "4A86C8", "9CBBE2", "1D4F8A", "C5CDD8", "6E7F96"],
        PIE_MAIN="2D63A6",
        LOGO_DARK=(11, 42, 74), LOGO_LIGHT=(156, 187, 226),
        INPUT_WORD="Gelb", LINK_WORD="Blau",
    ),
}
INK, INK2, MUTED, MUTED2 = "1A1D21", "3A3F45", "5B6068", "8A9099"
HEAD_TINT = "EEF3FA"
LINE, LINE2, HEAD_BG = "E6E8EB", "D5D9DE", "F3F4F5"
ON_DARK = "FFFFFF"
RED, AMB, GRN = "B42318", "B54708", "1F7A4D"
SANS, DISPLAY = "Calibri", "Calibri"

TAB_GROUP = {"Cockpit": "aus", "Diagramme": "aus", "Eingaben": "inp", "Steuern": "ber", "Projektion": "ber",
             "Finanzierung": "ber", "AfA-Vergleich": "ber", "Sensitivität": "ber", "Bankgespräch": "bank",
             "Haushaltsrechnung": "bank", "Vermögensaufstellung": "bank", "Hinweise": "anh", "Konfiguration": "anh"}
FORM_SHEETS = ("Eingaben", "Konfiguration")  # Navy-Formularbänder (P2-04)
GROUP_LABEL = {"ein": "Leitfaden", "inp": "Eingaben", "aus": "Auswertung", "ber": "Berechnung", "bank": "Bank",
               "anh": "Anhang"}

# Vorlagenfarben (Pro 1)
O_BG, O_DARK, O_SAND, O_SAND2, O_BRASS, O_WHITE = "F4F0E8", "0F1C16", "EBE6DB", "E2DCCE", "A08A55", "FFFFFF"

CF_FONT_MAP = {"9DB4A6": GRN, "E0A898": RED, "C6A45C": AMB, "4E7A62": GRN, "8C3B2E": RED, "8C7748": AMB}
SCALE_MAP = {"E9C6BB": "F6D3CD", "C3D8CA": "CDE8D8", "FFFFFF": "FFFFFF"}
NUMFMT_MAP = {
    "0.00%": "0.00 %", "0.0%": "0.0 %", "0%": "0 %", "0.00\\x": '0.00"×"', '0.0"-fach"': '0.0"-fach"',
    '\\+0%;\\-0%;"Basis"': '+0 %;-0 %;"Basis"', '\\+0.0%;\\-0.0%;"Basis"': '+0.0 %;-0.0 %;"Basis"',
}


def apply_theme(name):
    """Setzt die Farbrollen des gewählten Themas als Modulkonstanten."""
    t = THEMES[name]
    g = globals()
    g.update({k: v for k, v in t.items()})
    g["ON_DARK_ACC"] = t["ON_DARK_ACC"]
    g["TEXT_MAP_LIGHT"] = {"0C0F0D": INK, "6B685E": MUTED, "B9B4A6": MUTED2, "4E7A62": t["TEAL_MID"],
                           "A08A55": t["TEAL_MID"], "C6A45C": AMB, "6F8F7A": GRN, "8C3B2E": RED, "173026": t["TEAL"],
                           "F4F0E8": INK, "EBE6DB": MUTED}
    g["TEXT_MAP_DARK"] = {"F4F0E8": ON_DARK, "EBE6DB": t["ON_DARK_2"], "B9B4A6": t["ON_DARK_2"], "A08A55": t["ON_DARK_ACC"],
                          "C6A45C": t["ON_DARK_ACC"], "0C0F0D": ON_DARK, "0F1C16": ON_DARK, "6B685E": t["ON_DARK_2"]}
    ser = t["SERIES"]
    g["SERIES_MAP"] = {"0f1c16": ser[0], "173026": ser[3], "a08a55": ser[1], "c6a45c": ser[2], "6f8f7a": "3E8E6A",
                       "6b685e": ser[5], "b9b4a6": ser[4], "e2dcce": "E6E8EB", "8c3b2e": RED, "4f81bd": ser[1],
                       "7f9aa6": ser[2], "ffffff": "FFFFFF"}
    g["TEXT_FIXES"] = {
        "Weiß mit Messinglinie – ausfüllen": f"{t['INPUT_WORD']} hinterlegt – hier eingeben",
        "Grün – Eingabe im Leitfaden": f"{t['LINK_WORD']} – übernommen aus dem Leitfaden",
        "grün / messing / rot (Konfiguration)": "grün / gelb / rot (Schwellen: Konfiguration)",
    }
    iw, lw = t["INPUT_WORD"], t["LINK_WORD"]
    g["TEXT_REGEX"] = [(re.compile(r"(?<![A-Za-zÄÖÜäöüß])'([^'\n]{2,160}?)'(?![A-Za-zÄÖÜäöüß])"), "„\\1“"),  # P3-03
                       (re.compile(r"^Rechtsform / Investor$"), "Rechtsform"),
                       (re.compile(r"Weiße Felder"), f"{iw} hinterlegte Felder"),
                       (re.compile(r"weiße Felder"), f"{iw.lower()} hinterlegte Felder"),
                       (re.compile(r"\(weiß\)"), f"({iw.lower()})"),
                       (re.compile(r"Grün geschriebene"), f"{lw} geschriebene"),
                       (re.compile(r"grün geschriebene"), f"{lw.lower()} geschriebene")]


apply_theme("blau")


def fix_text(v):
    if not isinstance(v, str):
        return v
    v = TEXT_FIXES.get(v, v)
    for pat, rep in TEXT_REGEX:
        v = pat.sub(rep, v)
    return v


def rgb(color):
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return None
    return color.rgb[-6:].upper()


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


NOFILL = PatternFill(fill_type=None)


def side(style, color):
    return Side(style=style, color=color)


def font_like(f, **kw):
    nf = copy(f)
    for k, v in kw.items():
        setattr(nf, k, v)
    return nf


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


# ============================================================================ Klassifizierung
class Orig:
    """Unveränderliche Sicht auf den Originalstil einer Zelle."""
    __slots__ = ("fill", "fname", "size", "bold", "italic", "color", "locked", "link", "value")

    def __init__(self, c):
        self.fill = rgb(c.fill.fgColor) if c.fill is not None and c.fill.fill_type == "solid" else None
        self.fname = c.font.name
        self.size = c.font.sz or 11
        self.bold = bool(c.font.b)
        self.italic = bool(c.font.i)
        self.color = rgb(c.font.color)
        self.locked = c.protection.locked is not False
        self.link = c.hyperlink is not None
        self.value = c.value


def is_tile_label(o):
    return o.fill == O_DARK and o.color == "C6A45C" and o.fname == "Inter"


def is_tile_value(o):
    return o.fill == O_DARK and o.fname == "Fraunces" and o.size >= 12 and o.color == "F4F0E8"


def is_rule_header(o):
    return o.fill == O_DARK and o.fname == "Inter" and o.bold and o.size <= 9 and o.color == "F4F0E8"


def is_year_header(o):
    return o.fill == O_DARK and o.fname == "Inter" and o.bold and o.size >= 9.5 and o.color == "F4F0E8"


# ============================================================================ Zellen umgestalten
def style_text(c, o, dark):
    """Schrift nach Rolle."""
    f = c.font
    size, bold, name, color = o.size, o.bold, SANS, None
    if o.fname == "Fraunces":
        name = DISPLAY if o.size >= 12 else SANS
        size = {34: 30, 20: 22, 18: 22, 16: 20, 14: 16, 13: 13, 12: 12, 11: 11}.get(int(o.size), o.size)
        bold = o.size >= 11
        color = TEAL if (o.size >= 13 and not dark) else None
    elif o.fname == "IBM Plex Mono":
        size = max(o.size, 8)
        bold = o.color == O_BRASS
    tmap = TEXT_MAP_DARK if dark else TEXT_MAP_LIGHT
    if color is None and o.color:
        if o.color == O_DARK and not dark:
            color = TEAL if (bold or size >= 11) else INK
        else:
            color = tmap.get(o.color, ON_DARK if dark else INK)
    elif color is None:
        color = ON_DARK if dark else INK
    c.font = font_like(f, name=name, sz=size, b=bold, i=False, color=color, u=None if o.link else f.u)


def restyle_borders(c):
    b = c.border
    if not any(s is not None and s.style for s in (b.left, b.right, b.top, b.bottom)):
        return

    def m(s, vertical=False):
        if s is None or not s.style:
            return s
        col = rgb(s.color)
        if col in ("D6D0C2",):
            return side("hair" if s.style == "hair" else s.style, LINE)
        if col == "B9B4A6":
            return side("hair", LINE2)
        if col == O_DARK:
            return side(s.style, LINE2 if vertical else TEAL)
        if col == O_BRASS:
            return side(s.style, TEAL if s.style == "medium" else ACC)
        return s
    c.border = Border(left=m(b.left, True), right=m(b.right, True), top=m(b.top), bottom=m(b.bottom))


def restyle_cell(ws, c, o, ctx):
    hero = ctx["hero"](c.row)
    lb = c.border.left
    callout = lb is not None and lb.style == "medium" and rgb(lb.color) == O_BRASS
    restyle_borders(c)
    if callout:
        c.fill = fill(TEAL_XL)
        c.border = Border(left=side("thick", TEAL))
        if o.value is not None:
            style_text(c, o, False)
            c.font = font_like(c.font, sz=min(o.size, 9.5), color=INK2 if not o.bold else TEAL)
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        return
    # Vorlagen ohne Blattschutz: Eingabefelder erkennen (weiß, dunkle Schrift, Rahmen unten)
    if ctx["template"] and o.fill == O_WHITE and o.fname == "Inter" and o.color == O_DARK and not o.bold \
            and o.value is not None and not is_formula(o.value):
        o.locked = False

    if ctx["template"] and o.fill == O_WHITE and o.color == "4E7A62" and is_formula(o.value):
        c.protection = Protection(locked=False)  # verknüpft, darf überschrieben werden
    # ---- Eingabefelder
    if not o.locked:
        if ctx["template"]:
            c.protection = Protection(locked=False)
        style_text(c, o, False)
        c.fill = fill(INPUT_BG)
        c.font = font_like(c.font, color=INPUT_FG, b=True)
        ln = side("thin", INPUT_LINE)
        c.border = Border(left=ln, right=ln, top=ln, bottom=ln)
        return

    f = o.fill
    # ---- Dunkle Flächen
    if f == O_DARK:
        if hero:
            c.fill = fill(TEAL)
            style_text(c, o, True)
            if o.fname == "Fraunces" and o.size >= 30:
                c.font = font_like(c.font, name=DISPLAY, sz=30, b=True)
            return
        if c.row in ctx["tile_rows"]:
            label_row = c.row in ctx["tile_label_rows"]
            c.fill = fill(TEAL if label_row else TEAL_XL)
            if is_tile_label(o):
                c.font = Font(name=SANS, sz=8, b=True, color=ON_DARK)
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=False, shrink_to_fit=False)
            elif is_tile_value(o):
                c.font = Font(name=DISPLAY, sz=20 if o.size >= 16 else 12.5, b=True, color=TEAL)
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1, shrink_to_fit=False)
            return
        if c.row in ctx["rule_rows"] and c.column in ctx["rule_rows"][c.row]:
            run = ctx["rule_rows"][c.row]
            if ws.title in FORM_SHEETS:  # Formular-Band: einzige Überschriftenebene, Navy-Vollfläche
                c.fill = fill(TEAL)
                c.border = Border()
                if o.value is not None:
                    c.font = Font(name=SANS, sz=10, b=True, color=ON_DARK)
            else:  # L1b Tabellenabschnitt
                c.fill = fill(TEAL_L)
                c.border = Border(left=side("thick", ACC) if c.column == min(run) else None,
                                  bottom=side("thin", TEAL_MID))
                if o.value is not None:
                    c.font = Font(name=SANS, sz=9, b=True, color=TEAL)
            if o.value is not None:
                c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            return
        c.fill = fill(TEAL)
        style_text(c, o, True)
        if is_year_header(o):
            c.font = Font(name=SANS, sz=9, b=True, color=ON_DARK)
        return

    # ---- Messing (Primär-Schaltflächen / Akzentlinien)
    if f == O_BRASS:
        if o.value is not None or o.link:
            c.fill = fill(TEAL)
            c.font = Font(name=SANS, sz=10, b=True, color=ON_DARK)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border()
        else:
            c.fill = fill(ACC if ws.row_dimensions[c.row].height and ws.row_dimensions[c.row].height <= 4 else TEAL)
        return

    # ---- Sand (Tabellenköpfe, Summen, Ergebnisse, Sekundär-Schaltflächen)
    if f in (O_SAND, O_SAND2):
        style_text(c, o, False)
        if o.link:  # Sekundär-Schaltfläche
            c.fill = NOFILL
            c.font = Font(name=SANS, sz=10, b=True, color=TEAL)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ln = side("thin", TEAL)
            c.border = Border(left=ln, right=ln, top=ln, bottom=ln)
        elif o.fname == "Fraunces":  # Bild-Platzhalter
            c.fill = fill("F7F8F9")
            c.font = Font(name=SANS, sz=9, color=MUTED)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            d = side("dashed", LINE2)
            c.border = Border(left=d, right=d, top=d, bottom=d)
        elif o.bold and o.size < 10:  # Spaltenkopf
            c.fill = fill(HEAD_TINT)
            c.font = Font(name=SANS, sz=8, b=True, color=TEAL_MID)
            if isinstance(o.value, str) and not is_formula(o.value) and len(o.value) <= 40:
                if c.data_type == "s" and o.value.upper() != o.value:
                    c.value = o.value.upper()
            c.border = Border(top=c.border.top, bottom=side("thin", ACC))
        elif f == O_SAND2 and o.bold:  # Ergebnis
            c.fill = fill(TEAL_L)
            c.font = font_like(c.font, b=True, color=TEAL)
        elif f == O_SAND and o.color == "4E7A62" and not o.bold:  # Verknüpfung
            c.fill = NOFILL
        else:  # Summe
            c.fill = fill(SUM_BG)
        return

    # ---- Hell (Seite / Karten)
    if f in (O_BG, O_WHITE, None):
        c.fill = NOFILL
        style_text(c, o, False)
        if o.link and o.fname == "Inter" and o.bold:
            c.font = font_like(c.font, color=TEAL)
        if o.fname == "IBM Plex Mono" and o.color == O_BRASS and o.size >= 9:  # Schrittnummern
            c.font = Font(name=DISPLAY, sz=14, b=True, color=ACC)
        if o.fname == "IBM Plex Mono" and o.color == "6B685E":  # Firmenzeile
            c.font = Font(name=SANS, sz=8, color=MUTED)
        if o.italic and o.size <= 8:  # Haftungshinweis
            c.font = Font(name=SANS, sz=8, color=MUTED)
        if o.color == "8C3B2E" and is_formula(o.value):  # Hinweislisten neutral, ⚠ per Bedingung
            c.font = font_like(c.font, color=INK)
            ctx["hints"].append(c.coordinate)


def analyse_rows(ws, origs, hero_rows):
    """Kachelzeilen und Linien-Überschriften (Laufweite) je Zeile bestimmen."""
    tile_rows, rule_rows, tile_label_rows = set(), {}, set()
    by_row = {}
    for (r, col), o in origs.items():
        by_row.setdefault(r, {})[col] = o
    for r, cells in by_row.items():
        if r <= 3 or r in hero_rows:
            continue
        if any(is_tile_label(o) or is_tile_value(o) for o in cells.values()):
            tile_rows.add(r)
            if any(is_tile_label(o) for o in cells.values()):
                tile_label_rows.add(r)
            continue
        # zusammenhängende dunkle Läufe
        cols = sorted(col for col, o in cells.items() if o.fill == O_DARK)
        runs, cur = [], []
        for col in cols:
            if cur and col != cur[-1] + 1:
                runs.append(cur)
                cur = []
            cur.append(col)
        if cur:
            runs.append(cur)
        for run in runs:
            os_ = [cells[col] for col in run]
            if any(is_year_header(o) for o in os_):
                continue
            if any(is_rule_header(o) for o in os_):
                rule_rows.setdefault(r, set()).update(run)
    return tile_rows, rule_rows, tile_label_rows


def tile_gaps(ws, tile_rows, tile_label_rows):
    """Weiße Fugen zwischen Kacheln und teal Oberkante auf Label-Zeilen."""
    gap = side("thick", "FFFFFF")
    for r in sorted(tile_rows):
        tint = TEAL if r in tile_label_rows else TEAL_XL
        cells = [c for c in ws[r] if c.fill is not None and c.fill.fill_type == "solid" and rgb(c.fill.fgColor) == tint]
        if not cells:
            continue
        label_row = False
        anchors = {}
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= r <= mr.max_row:
                for col in range(mr.min_col, mr.max_col + 1):
                    anchors[col] = mr.max_col
        for c in cells:
            right_edge = anchors.get(c.column, c.column) == c.column
            nxt = ws.cell(r, c.column + 1)
            nxt_tile = nxt.fill is not None and nxt.fill.fill_type == "solid" and rgb(nxt.fill.fgColor) == tint
            c.border = Border(top=side("thick", TEAL) if label_row else None,
                              right=gap if (right_edge and nxt_tile) else None)


def sheet_title(ws):
    m = re.match(r"S(\d\d) (.+)", ws.title)
    if m:
        return f"Leitfaden  ›  Schritt {int(m.group(1))} von 12  ·  {m.group(2)}"
    grp = TAB_GROUP.get(ws.title, "ein")
    if ws.title in ("Start", "Leitfaden"):
        return ws.title
    return f"{GROUP_LABEL[grp]}  ›  {ws.title}"


# ============================================================================ Abschnittsüberschriften
def section_rules(ws):
    """Unterstrich unter Abschnittsüberschriften (Calibri 13) über die Blockbreite."""
    merged = {}
    for mr in ws.merged_cells.ranges:
        merged[(mr.min_row, mr.min_col)] = mr
    for c in list(ws._cells.values()):
        if c.value is None or c.row <= 3:
            continue
        if c.font.name == DISPLAY and c.font.sz == 13 and rgb(c.font.color) == TEAL:
            mr = merged.get((c.row, c.column))
            cols = range(mr.min_col, mr.max_col + 1) if mr else [c.column]
            for k, col in enumerate(cols):
                cc = ws.cell(c.row, col)
                cc.fill = fill(TEAL_L)
                cc.border = Border(left=side("thick", ACC) if k == 0 else None, bottom=side("thin", ACC))
            c.font = font_like(c.font, sz=12.5)
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[c.row].height = max(ws.row_dimensions[c.row].height or 15, 24)


def normalise_merged(ws):
    """Füllung der Ankerzelle auf ganzen Verbundbereich übertragen (saubere Flächen)."""
    for mr in ws.merged_cells.ranges:
        a = ws.cell(mr.min_row, mr.min_col)
        for r in range(mr.min_row, mr.max_row + 1):
            for col in range(mr.min_col, mr.max_col + 1):
                if (r, col) != (mr.min_row, mr.min_col):
                    ws.cell(r, col).fill = copy(a.fill)


# ============================================================================ Arbeitsmappe
def design_workbook(src, tmp):
    wb = load_workbook(src)
    for ws in wb.worksheets:
        hero_rows = set(range(4, 23)) if ws.title == "Start" else set()
        origs = {(c.row, c.column): Orig(c) for c in ws._cells.values()}
        tile_rows, rule_rows, tile_label_rows = analyse_rows(ws, origs, hero_rows)
        ctx = {"hero": lambda r, h=hero_rows: r in h, "tile_rows": tile_rows, "tile_label_rows": tile_label_rows,
               "rule_rows": rule_rows, "hints": [],
               "template": ws.title in ("Bankgespräch", "Haushaltsrechnung", "Vermögensaufstellung")}

        for c in list(ws._cells.values()):
            o = origs.get((c.row, c.column))
            if o is None:
                continue
            if isinstance(c.value, str) and c.data_type == "s":
                new = fix_text(c.value)
                if new != c.value:
                    c.value = new
                    c.data_type = "s"
            restyle_cell(ws, c, o, ctx)
            if c.number_format in NUMFMT_MAP:
                c.number_format = NUMFMT_MAP[c.number_format]

        # P1-05: Stil nur im Lauf der Überschrift (nie die ganze Zeile); Höhe nur, wenn der Rest der Zeile leer ist
        for r, run in rule_rows.items():
            rest_empty = all(c.value is None for c in ws[r] if c.column not in run)
            if rest_empty:
                ws.row_dimensions[r].height = 22 if ws.title in FORM_SHEETS else 20
            else:
                ws.row_dimensions[r].height = 18
            for cc in run:
                c = ws.cell(r, cc)
                if c.value is not None:
                    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in ws._cells.values():
            if c.row > 3 and c.value is not None and c.row not in rule_rows and c.row not in tile_rows:
                al = c.alignment
                if al.vertical in (None, "bottom"):
                    c.alignment = Alignment(horizontal=al.horizontal, vertical="center", wrap_text=al.wrap_text,
                                            indent=al.indent, shrink_to_fit=al.shrink_to_fit, text_rotation=al.text_rotation)
        normalise_merged(ws)
        tile_gaps(ws, tile_rows, tile_label_rows)
        section_rules(ws)

        for ref in ctx["hints"]:
            ws.conditional_formatting.add(ref, FormulaRule(formula=[f'LEFT({ref},1)="⚠"'], font=Font(color=RED, bold=True)))
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                d = rule.dxf
                if d is not None and d.font is not None:
                    d.font.name = None
                if d is not None and d.font is not None and d.font.color is not None:
                    old = rgb(d.font.color)
                    if old in CF_FONT_MAP:
                        d.font.color = CF_FONT_MAP[old]
                if rule.type == "colorScale" and rule.colorScale is not None:
                    for col in rule.colorScale.color:
                        old = (col.rgb or "")[-6:].upper()
                        if old in SCALE_MAP:
                            col.rgb = "FF" + SCALE_MAP[old]

        ws.sheet_properties.tabColor = GROUP[TAB_GROUP.get(ws.title, "ein")]
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 100
        if ws.freeze_panes is None and ws.title != "Start":
            ws.freeze_panes = "A4"
        ws.oddHeader.left.text = None
        ws.oddFooter.left.text = "MM Holding GmbH · Immobilien-Kalkulation · &A"
        ws.oddFooter.center.text = None
        ws.oddFooter.right.text = "Seite &P von &N"
        for part in (ws.oddFooter.left, ws.oddFooter.right):
            part.size = 8
            part.font = "Calibri,Regular"
            part.color = MUTED2

    # ---- globale Regeln (früh) → Blatt-Layouts → globale Regeln (final) → Dashboard → Druck
    import global_rules
    run_hook("global_rules.early", global_rules.early, wb)
    for name in LAYOUTS:
        try:
            mod = importlib.import_module(f"layouts.{name}")
        except Exception:
            MODULE_ERRORS.append(f"layouts.{name} (Import)")
            traceback.print_exc()
            if STRICT:
                raise
            continue
        run_hook(f"layouts.{name}", mod.apply, wb)
    run_hook("global_rules.final", global_rules.final, wb)
    run_hook("dashboard.build", dashboard.build, wb,
             {k: globals()[k] for k in ("TEAL", "TEAL_MID", "ACC", "TEAL_L", "TEAL_XL", "ON_DARK_2", "ON_DARK_ACC")})
    run_hook("global_rules.page_setup", global_rules.page_setup, wb)
    # Beim Öffnen: jedes Blatt oben links, Cursor in der ersten Eingabe-/Inhaltszeile
    for w in wb.worksheets:
        w.sheet_view.topLeftCell = "A1"
        anchor = w.freeze_panes or "A1"
        for sel in w.sheet_view.selection:
            sel.activeCell = anchor
            sel.sqref = anchor
    # Standardschrift Calibri 11 → Excel-Ziffernbreite 7 px; alle Zellen tragen ihre eigene Calibri-Schrift
    wb._named_styles["Normal"].font = Font(name="Calibri", sz=11)
    wb.properties.title = "Immobilien-Kalkulation"
    wb.properties.creator = "MM Holding GmbH"
    wb.properties.subject = "Kauf, Finanzierung, Cashflow, Steuern und Exit vermieteter Immobilien"
    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w.title == "Start"
    wb.save(tmp)


def recolor_png(data):
    im = Image.open(io.BytesIO(data)).convert("RGBA")
    px = im.load()
    targets = [((15, 28, 22), LOGO_DARK), ((160, 138, 85), LOGO_LIGHT), ((198, 164, 92), LOGO_LIGHT),
               ((244, 240, 232), (255, 255, 255))]
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b, al = px[x, y]
            if al == 0:
                continue
            best = min(targets, key=lambda t: sum((t[0][i] - (r, g, b)[i]) ** 2 for i in range(3)))
            if sum((best[0][i] - (r, g, b)[i]) ** 2 for i in range(3)) < 60 ** 2:
                px[x, y] = (*best[1], al)
    out = io.BytesIO()
    im.save(out, "PNG")
    return out.getvalue()


def postprocess(tmp, dst):
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/theme/theme1.xml":
                s = data.decode("utf-8")
                s = re.sub(r'(<a:majorFont><a:latin typeface=")[^"]*"', rf'\1{DISPLAY}"', s)
                s = re.sub(r'(<a:minorFont><a:latin typeface=")[^"]*"', rf'\1{SANS}"', s)
                data = s.encode("utf-8")
            elif item.filename.startswith("xl/media/") and item.filename.endswith(".png"):
                data = recolor_png(data)
            zout.writestr(item, data)


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    apply_theme(sys.argv[3] if len(sys.argv) > 3 else "blau")
    tmp = dst + ".tmp.xlsx"
    design_workbook(src, tmp)
    postprocess(tmp, dst)
    os.remove(tmp)
    if MODULE_ERRORS:
        print("Module mit Fehlern:", ", ".join(MODULE_ERRORS), file=sys.stderr)
    print(f"gespeichert: {dst}")
