"""Gemeinsame Design-Bausteine des Immobilien-Kalkulationstools (Blau-Weiß).

Alle Layout-Module verwenden AUSSCHLIESSLICH diese Tokens und Komponenten, damit jedes Blatt
dieselbe visuelle Sprache spricht. Die Bausteine verändern nur Darstellung (Stil, Zahlenformat,
Zeilenhöhe, Verbund, statische Beschriftungen) – niemals Formeln oder Werte bestehender Rechenzellen.
"""
import math
import re

from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

# =============================================================================== Farben
NAVY = "0B2A4A"      # Primär: Kopfleiste, Titel, Kachel-Label, Jahreskopf, Schlüsselergebnis
BLUE = "1D4F8A"      # Sekundär: Primär-Button, Links, verknüpfte Werte, L2-Überschriften
ACCENT = "4A86C8"    # Akzent: Linien, Akzentkanten, Schrittnummern
SKY = "9CBBE2"
MIST = "C8D7EB"
TINT = "E7EEF7"      # L1-Band, Ergebniszeile
TINT_XL = "F3F7FC"   # Kachelfläche, Blockergebnis, Info-Zeilen
HEAD = "EEF3FA"      # Spaltenkopf
INK = "1A1D21"
INK2 = "3A3F45"
MUTED = "5B6068"     # Sekundärtext (bis 9 pt immer diese Farbe)
MUTED2 = "8A9099"    # nur Linien, inaktive Felder, Text ab 11 pt
LINE = "E6E8EB"      # Haarlinie Datenzeile
LINE2 = "D5D9DE"     # Trennlinie / inaktiver Rahmen
WHITE = "FFFFFF"
GREEN, AMBER, RED = "1F7A4D", "B54708", "B42318"
GREEN_BG, AMBER_BG, RED_BG = "EAF5EF", "FDF3E7", "FBECEB"
INPUT_BG, INPUT_LINE, INPUT_FG = "FFF5D6", "E6CB77", BLUE
INACTIVE_BG, INACTIVE_FG = "F3F4F6", MUTED2

SANS, DISPLAY = "Calibri", "Calibri"

# =============================================================================== Typo-Skala und Zeilenraster
# 8 · 9 · 10 · 12,5 · 16 · 20 · 22 (Hero 30, Schrittnummer 14). Keine Zellschrift unter 8 pt.
T_MICRO, T_SMALL, T_BODY, T_H3, T_H2, T_KPI, T_H1, T_HERO = 8, 9, 10, 12.5, 16, 20, 22, 30
H_BAND, H_BAND_B, H_HEAD, H_ROW, H_ROW2, H_ROW3 = 24, 20, 20, 18, 30, 42
H_STEP_ROW, H_STEP_ROW2 = 22, 32            # Datenzeilen der Schritt-Seiten (1- / 2-zeilig)
H_TILE_LABEL, H_TILE_VALUE, H_TILE_SUB, H_BUTTON = 18, 30, 16, 26

# =============================================================================== Zahlenformate
NUMFMT = {
    "eur": '#,##0" €";-#,##0" €";"–"',
    "eur_zero": '#,##0" €";-#,##0" €";0" €"',
    "eur2": '#,##0.00" €";-#,##0.00" €";"–"',
    "num": '#,##0;-#,##0;"–"',
    "int": "#,##0",
    "pct1": '0.0 %;-0.0 %;"–"',
    "pct2": '0.00 %;-0.00 %;"–"',
    "dscr": "0.00",
    "mult1": '0.0"×"',
    "mult2": '0.00"×"',
    "year": "0",
    "years": '0" Jahre"',
    "date": "DD.MM.YYYY",
    "qm": '#,##0" m²"',
    "eur_qm": '#,##0.00" €/m²"',
    "tax_effect": '"Zahlung "#,##0" €";"Erstattung "#,##0" €";"–"',
    "yesno": '[=1]"ja";[=0]"nein";0',
}

# =============================================================================== KPI-Spezifikation (einzige Quelle)
# rule: "ampel" = grün ab green, gelb ab yellow, sonst rot (Namen aus Konfiguration);
#       "cf"    = rot < 0, gelb wenn Jahr 1 ≥ 0 aber Jahr 2 < 0, sonst grün.
KPI = {
    "GI": dict(label="Gesamtinvestition", fmt=NUMFMT["eur"]),
    "EK": dict(label="Eigenkapitalbedarf inkl. Reserve", fmt=NUMFMT["eur"]),
    "RATE": dict(label="Rate an die Bank / Monat", fmt=NUMFMT["eur"]),
    "CF": dict(label="Cashflow n. St. / Monat (Jahr 1)", fmt=NUMFMT["eur"], rule="cf"),
    "CFV": dict(label="Cashflow v. St. / Monat (Jahr 1)", fmt=NUMFMT["eur"], rule="neg"),
    "BMR": dict(label="Bruttomietrendite", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_BMR_gruen", yellow="Ampel_BMR_gelb"),
    "NMR": dict(label="Nettomietrendite", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_NMR_gruen", yellow="Ampel_NMR_gelb"),
    "DSCR": dict(label="DSCR (Jahr 1)", fmt=NUMFMT["dscr"], rule="ampel", green="Ampel_DSCR_gruen", yellow="Ampel_DSCR_gelb"),
    "EKR": dict(label="EK-Rendite Jahr 1", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_EKR_gruen", yellow="Ampel_EKR_gelb"),
    "IRR": dict(label="IRR n. St.", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_IRR_gruen", yellow="Ampel_IRR_gelb"),
    "FAKTOR": dict(label="Kaufpreisfaktor", fmt=NUMFMT["mult1"]),
    "MULT": dict(label="Eigenkapital-Multiple", fmt=NUMFMT["mult2"]),
    "RF": dict(label="Rechtsform", fmt="@"),
}
CF_YEAR2 = "INDEX(Projektion!$D$37:$AQ$37,2)"
RECHTSFORM_SHORT = 'CHOOSE(Rechtsform_Idx,"Privat","vv-GmbH","GmbH")'

# Link-Ziele: erste Zelle des scrollbaren Bereichs (A1 liegt im fixierten Bereich)
LINK_TARGET = {"Projektion": "D11", "Finanzierung": "D11", "Steuern": "D4", "AfA-Vergleich": "D4"}


def link_target(sheet):
    if re.match(r"S\d\d ", sheet):
        return "D12"
    return LINK_TARGET.get(sheet, "A4")


# =============================================================================== Primitive
def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


NOFILL = PatternFill(fill_type=None)


def side(style, color):
    return Side(style=style, color=color)


def font(size=T_BODY, bold=False, color=INK, name=SANS, italic=False, underline=None):
    return Font(name=name, sz=size, b=bold, i=italic, color=color, u=underline)


def align(h="left", v="center", indent=0, wrap=False):
    """Einzug > 0 nur mit horizontal left/right (Excel ignoriert ihn bei 'general')."""
    if indent and h not in ("left", "right", "distributed"):
        h = "left"
    return Alignment(horizontal=h, vertical=v, indent=indent, wrap_text=wrap, shrink_to_fit=False)


def col(c):
    return c if isinstance(c, int) else column_index_from_string(c)


def L(c):
    return c if isinstance(c, str) else get_column_letter(c)


def rng(c1, r1, c2=None, r2=None):
    c2 = c1 if c2 is None else c2
    r2 = r1 if r2 is None else r2
    return f"{L(c1)}{r1}:{L(c2)}{r2}"


def iter_cells(ws, c1, r1, c2, r2):
    for r in range(r1, r2 + 1):
        for cc in range(col(c1), col(c2) + 1):
            yield ws.cell(r, cc)


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def set_text(cell, text):
    """Statischen Text setzen (auch mit führendem '=' als Text, nie als Formel)."""
    cell.value = text
    if isinstance(text, str):
        cell.data_type = "s"


def safe_merge(ws, c1, r1, c2, r2):
    """Verbinden, sofern nicht bereits identisch verbunden; überlappende Verbünde vorher lösen."""
    target = rng(c1, r1, c2, r2)
    if col(c1) == col(c2) and r1 == r2:
        return
    for mr in list(ws.merged_cells.ranges):
        if str(mr) == target:
            return
        if not (mr.max_row < r1 or mr.min_row > r2 or mr.max_col < col(c1) or mr.min_col > col(c2)):
            ws.unmerge_cells(str(mr))
    ws.merge_cells(target)


def style_range(ws, c1, r1, c2, r2, fnt=None, fil=None, brd=None, aln=None, fmt=None):
    for c in iter_cells(ws, c1, r1, c2, r2):
        if fnt is not None:
            c.font = fnt
        if fil is not None:
            c.fill = fil
        if brd is not None:
            c.border = brd
        if aln is not None:
            c.alignment = aln
        if fmt is not None:
            c.number_format = fmt


def set_height(ws, row, pt):
    ws.row_dimensions[row].height = pt


# =============================================================================== Maße
def col_px(ws, c):
    """Excel-Pixel einer Spalte (Standardschrift Calibri 11 → Ziffernbreite 7 px)."""
    d = ws.column_dimensions.get(L(c))
    w = d.width if (d is not None and d.width and (d.customWidth or d.width != 13)) else 8.43
    return int(((256 * w + int(128 / 7)) / 256) * 7)


def span_px(ws, c1, c2):
    return sum(col_px(ws, c) for c in range(col(c1), col(c2) + 1))


def text_px(text, size=T_BODY, bold=False):
    """Näherung der Laufweite von Calibri (px bei 96 dpi)."""
    if text is None:
        return 0
    s = str(text)
    em = size * 96 / 72
    narrow = sum(1 for ch in s if ch in "iljtfrI.,:;!'|() ·")
    wide = sum(1 for ch in s if ch in "MWmw@%€ÄÖÜ")
    caps = sum(1 for ch in s if ch.isupper())
    units = len(s) * 0.49 - narrow * 0.22 + wide * 0.25 + caps * 0.08
    return units * em * (1.07 if bold else 1.0)


def lines_needed(text, width_px, size=T_BODY, bold=False):
    if not text:
        return 1
    total = 0
    for para in str(text).split("\n"):
        w = text_px(para, size, bold)
        total += max(1, math.ceil(w / max(width_px - 10, 20)))
    return total


def fit_row_height(ws, row, c1, c2, base=H_ROW, line_pt=13, pad=8, max_lines=None):
    """Zeilenhöhe nach umbrechendem Inhalt der Zeile (verbundene Bereiche berücksichtigt)."""
    merged = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_row == row == mr.max_row:
            merged[mr.min_col] = mr.max_col
    need = 1
    for cc in range(col(c1), col(c2) + 1):
        cell = ws.cell(row, cc)
        if cell.value is None or is_formula(cell.value):
            continue
        end = merged.get(cc, cc)
        if not (cell.alignment and cell.alignment.wrap_text):
            continue
        sz = cell.font.sz or T_BODY
        n = lines_needed(cell.value, span_px(ws, cc, end), sz, bool(cell.font.b))
        need = max(need, n)
    if max_lines:
        need = min(need, max_lines)
    ws.row_dimensions[row].height = base if need == 1 else need * line_pt + pad
    return need


# =============================================================================== Überschriften
def band_l1(ws, row, c1, c2, title=None, right=None, height=H_BAND):
    """L1 Abschnitt: E7EEF7, Akzentkante links thick, Unterkante thin, 12,5 pt Display fett Navy."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.fill = fill(TINT)
        c.border = Border(left=side("thick", ACCENT) if k == 0 else None, bottom=side("thin", ACCENT))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_H3, True, NAVY, DISPLAY)
    first.alignment = align("left", "center", 1)
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(T_MICRO, False, BLUE)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def band_l1b(ws, row, c1, c2, title=None, right=None, height=H_BAND_B):
    """L1b Tabellenabschnitt: E7EEF7, 9 pt fett Versalien Navy, Akzentkante links, Unterkante Blau."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.fill = fill(TINT)
        c.border = Border(left=side("thick", ACCENT) if k == 0 else None, bottom=side("thin", BLUE))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_SMALL, True, NAVY)
    first.alignment = align("left", "center", 1)
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(T_MICRO, False, BLUE)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def band_form(ws, row, c1, c2, title=None, right=None, height=H_BAND_B + 2):
    """Formular-Band (nur Eingaben/Konfiguration): Navy-Vollfläche, 10 pt fett weiß."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = fill(NAVY)
        c.border = Border()
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_BODY, True, WHITE)
    first.alignment = align("left", "center", 1)
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(8.5 if False else T_MICRO, False, SKY)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def subhead_l2(ws, row, c1, c2, title=None, height=H_HEAD):
    """L2 Unterblock: 8,5→8/9 pt fett Versalien Blau, ohne Füllung, Unterkante thin Akzent."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = NOFILL
        c.border = Border(bottom=side("thin", ACCENT))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_SMALL, True, BLUE)
    first.alignment = align("left", "center", 1)
    set_height(ws, row, height)


def table_head(ws, row, c1, c2, labels=None, height=H_HEAD):
    """Spaltenkopf über volle Tabellenbreite: EEF3FA, 8 pt fett Versalien Blau, Unterkante thin Akzent.
    labels: {spalte: (text, 'left'|'right'|'center')} – Köpfe folgen der Ausrichtung ihrer Werte."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = fill(HEAD)
        c.border = Border(bottom=side("thin", ACCENT))
        c.font = font(T_MICRO, True, BLUE)
        if c.value is not None and c.alignment.horizontal not in ("right", "center"):
            c.alignment = align("left", "center", 1)
    for cc, spec in (labels or {}).items():
        text, h = spec if isinstance(spec, tuple) else (spec, "left")
        c = ws.cell(row, col(cc))
        set_text(c, text)
        c.alignment = align(h, "center", 1 if h in ("left", "right") else 0)
    set_height(ws, row, height)


def hairline(ws, row, c1, c2):
    for c in iter_cells(ws, c1, row, c2, row):
        b = c.border
        c.border = Border(left=None, right=None, top=b.top, bottom=side("hair", LINE))


def total(ws, row, c1, c2, level=2):
    """Summenstufen: 1 Zwischensumme (fett, hair oben), 2 Blockergebnis (F3F7FC, thin Navy oben),
    3 Schlüsselergebnis (E7EEF7, thin oben, double unten, fett Navy)."""
    for c in iter_cells(ws, c1, row, c2, row):
        f = c.font
        if level == 1:
            c.fill = NOFILL
            c.border = Border(top=side("hair", LINE2), bottom=side("hair", LINE))
            c.font = font(f.sz or T_BODY, True, INK)
        elif level == 2:
            c.fill = fill(TINT_XL)
            c.border = Border(top=side("thin", NAVY), bottom=side("hair", LINE))
            c.font = font(f.sz or T_BODY, True, NAVY)
        else:
            c.fill = fill(TINT)
            c.border = Border(top=side("thin", NAVY), bottom=side("double", NAVY))
            c.font = font(f.sz or T_BODY, True, NAVY)
    if level == 3:
        set_height(ws, row, max(ws.row_dimensions[row].height or 0, H_ROW))


def memo(ws, row, c1, c2):
    """Nachrichtliche Zeile: kursiv 9 pt grau, Einzug 1."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.font = font(T_SMALL, False, MUTED, italic=True)
        c.fill = NOFILL


# =============================================================================== Ampel
def add_ampel(ws, ref, kpi_key, value_ref=None, with_fill=False):
    """Bedingte Schriftfarbe (und optional Tönung) nach KPI-Spezifikation.
    ref: Zielbereich (z. B. 'C12:D13'); value_ref: absolute Zelle/Name mit dem Wert (Default: erste Zelle von ref)."""
    spec = KPI[kpi_key]
    rule = spec.get("rule")
    if not rule:
        return
    v = value_ref or ref.split(":")[0].replace("$", "")
    if not re.match(r"^[A-Za-z_]", v) or re.match(r"^[A-Z]{1,3}\d+$", v):
        vv = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", v) if re.match(r"^[A-Z]{1,3}\d+$", v) else v
    else:
        vv = v

    def r_(formula, fg, bg):
        kw = dict(font=Font(color=fg, bold=True), stopIfTrue=True)
        if with_fill:
            kw["fill"] = fill(bg)
        ws.conditional_formatting.add(ref, FormulaRule(formula=[formula], **kw))

    if rule == "ampel":
        r_(f"AND(ISNUMBER({vv}),{vv}<{spec['yellow']})", RED, RED_BG)
        r_(f"AND(ISNUMBER({vv}),{vv}<{spec['green']})", AMBER, AMBER_BG)
        r_(f"ISNUMBER({vv})", GREEN, GREEN_BG)
    elif rule == "cf":
        r_(f"{vv}<0", RED, RED_BG)
        r_(f"{CF_YEAR2}<0", AMBER, AMBER_BG)
        r_(f"ISNUMBER({vv})", GREEN, GREEN_BG)
    elif rule == "neg":
        r_(f"{vv}<0", RED, RED_BG)


# =============================================================================== Komponenten
def kpi_tile(ws, c1, c2, label_row, value_row, sub_row=None, label=None, value=None, sub=None,
             kpi=None, fmt=None, value_rows=1, gap_right=True, value_size=T_KPI):
    """Kachel: Label-Streifen Navy (8 pt fett weiß, Satzschreibung), Wert 20 pt Display fett Navy auf F3F7FC,
    optionale Unterzeile 8 pt grau. value=None lässt einen vorhandenen Wert/Formel unverändert."""
    spec = KPI.get(kpi, {}) if kpi else {}
    last_value_row = value_row + value_rows - 1
    rows = [label_row] + list(range(value_row, last_value_row + 1)) + ([sub_row] if sub_row else [])
    for r in rows:
        for c in iter_cells(ws, c1, r, c2, r):
            c.fill = fill(NAVY if r == label_row else TINT_XL)
            c.border = Border(right=side("thick", WHITE) if (gap_right and col(c.column) == col(c2)) else None)
    safe_merge(ws, c1, label_row, c2, label_row)
    lab = ws.cell(label_row, col(c1))
    if label is not None or spec.get("label"):
        if not is_formula(lab.value) or label is not None:
            set_text(lab, label if label is not None else spec["label"])
    lab.font = font(T_MICRO, True, WHITE)
    lab.alignment = align("left", "center", 1)
    set_height(ws, label_row, H_TILE_LABEL)
    safe_merge(ws, c1, value_row, c2, last_value_row)
    val = ws.cell(value_row, col(c1))
    if value is not None:
        val.value = value
    val.font = font(value_size, True, NAVY, DISPLAY)
    val.alignment = align("left", "center", 1)
    f = fmt or spec.get("fmt")
    if f:
        val.number_format = f
    if value_rows == 1:
        set_height(ws, value_row, H_TILE_VALUE)
    if kpi and spec.get("rule"):
        add_ampel(ws, rng(c1, value_row, c2, last_value_row), kpi)
    if sub_row:
        safe_merge(ws, c1, sub_row, c2, sub_row)
        s = ws.cell(sub_row, col(c1))
        if sub is not None:
            s.value = sub
        s.font = font(T_MICRO, False, MUTED)
        s.alignment = align("left", "center", 1)
        set_height(ws, sub_row, H_TILE_SUB)


def button(ws, c1, row, c2, text, target_sheet=None, kind="primary", target_cell=None, tooltip=None):
    """Primär: 1D4F8A gefüllt, 10 pt fett weiß. Sekundär: weiß, Rahmen thin 1D4F8A, 10 pt fett Blau.
    Textlink: ohne Rahmen, 9,5→10 pt Blau. Jeder Link endet mit ' ›' oder beginnt mit '‹ '."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.border = Border()
        c.fill = NOFILL
    safe_merge(ws, c1, row, c2, row)
    c = ws.cell(row, col(c1))
    set_text(c, text)
    ln = side("thin", BLUE)
    if kind == "primary":
        style_range(ws, c1, row, c2, row, fil=fill(BLUE))
        c.font = font(T_BODY, True, WHITE)
        c.alignment = align("center", "center")
    elif kind == "secondary":
        for cc in iter_cells(ws, c1, row, c2, row):
            cc.border = Border(top=ln, bottom=ln, left=ln if cc.column == col(c1) else None,
                               right=ln if cc.column == col(c2) else None)
        c.font = font(T_BODY, True, BLUE)
        c.alignment = align("center", "center")
    else:
        c.font = font(T_BODY, False, BLUE)
        c.alignment = align("left", "center", 1)
    if target_sheet:
        loc = f"'{target_sheet}'!{target_cell or link_target(target_sheet)}"
        c.hyperlink = Hyperlink(ref=c.coordinate, location=loc, display=text, tooltip=tooltip)
    set_height(ws, row, H_BUTTON if kind != "link" else max(ws.row_dimensions[row].height or 0, H_ROW))


def text_link(cell, text, target_sheet, target_cell=None, size=T_BODY, bold=True, tooltip=None):
    set_text(cell, text)
    cell.font = font(size, bold, BLUE)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{target_sheet}'!{target_cell or link_target(target_sheet)}",
                               display=text, tooltip=tooltip)


def callout(ws, c1, head_row, c2, body_r1, body_r2, title="Einordnung", status=None, status_col=None):
    """Einordnung: ein Akzentbalken (thick 4A86C8) für Kopf und Text; Kopf 9 pt fett Navy,
    Text 9,5→10 pt INK2 auf F3F7FC, umbrechend, vertikal zentriert."""
    for r in range(head_row, body_r2 + 1):
        for c in iter_cells(ws, c1, r, c2, r):
            c.fill = fill(TINT_XL)
            c.border = Border(left=side("thick", ACCENT) if c.column == col(c1) else None,
                              bottom=side("thin", ACCENT) if r == head_row else None)
    h = ws.cell(head_row, col(c1))
    if title is not None:
        set_text(h, title)
    h.font = font(T_SMALL, True, NAVY)
    h.alignment = align("left", "center", 1)
    set_height(ws, head_row, H_HEAD)
    if status is not None:
        sc = ws.cell(head_row, col(status_col or c2))
        sc.value = status
        sc.font = font(T_MICRO, True, MUTED)
        sc.alignment = align("right", "center", 1)
    safe_merge(ws, c1, body_r1, c2, body_r2)
    b = ws.cell(body_r1, col(c1))
    b.font = font(T_SMALL, False, INK2)
    b.alignment = align("left", "center", 1, wrap=True)


def input_style(cell, state="required"):
    """Eingabezustände: required (gelb, Rahmen thin), optional (gelb, Rahmen dashed),
    inactive (grau), linked (weiß, dashed Akzent, Blau normal)."""
    if state in ("required", "optional"):
        ln = side("thin" if state == "required" else "dashed", INPUT_LINE)
        cell.fill = fill(INPUT_BG)
        cell.font = font(cell.font.sz or T_BODY, True, INPUT_FG)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)
        cell.protection = Protection(locked=False)
    elif state == "inactive":
        ln = side("thin", LINE2)
        cell.fill = fill(INACTIVE_BG)
        cell.font = font(cell.font.sz or T_BODY, False, INACTIVE_FG)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)
    elif state == "linked":
        ln = side("dashed", ACCENT)
        cell.fill = NOFILL
        cell.font = font(cell.font.sz or T_BODY, False, BLUE)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)


def inactive_when(ws, ref, condition):
    """Eingabefeld bedingt inaktiv darstellen (bleibt editierbar)."""
    ws.conditional_formatting.add(ref, FormulaRule(formula=[condition], stopIfTrue=True,
                                                   font=Font(color=INACTIVE_FG, bold=False),
                                                   fill=fill(INACTIVE_BG),
                                                   border=Border(left=side("thin", LINE2), right=side("thin", LINE2),
                                                                 top=side("thin", LINE2), bottom=side("thin", LINE2))))


def page_header(ws, c1, c2, eyebrow, title, subtitle=None, context=None, context_col=None):
    """Seitenkopf Z. 5–7: Eyebrow (8 pt fett Versalien Blau), H1 (22 pt Display fett Navy), Untertitel (10 pt grau).
    context: (formel_name, formel_adresse) rechts oben in context_col."""
    e, t, s = ws.cell(5, col(c1)), ws.cell(6, col(c1)), ws.cell(7, col(c1))
    set_text(e, eyebrow.upper())
    e.font = font(T_MICRO, True, BLUE)
    e.alignment = align("left", "bottom")
    if title is not None:
        if not is_formula(t.value) or not is_formula(title):
            t.value = title
    t.font = font(T_H1, True, NAVY, DISPLAY)
    t.alignment = align("left", "center")
    set_height(ws, 5, 16)
    set_height(ws, 6, 30)
    if subtitle is not None:
        s.value = subtitle
    if s.value is not None:
        s.font = font(T_BODY, False, MUTED)
        s.alignment = align("left", "top")
        set_height(ws, 7, 22)
    if context:
        cc = col(context_col or c2)
        n, a_ = ws.cell(6, cc), ws.cell(7, cc)
        n.value, a_.value = context
        n.font = font(T_BODY, True, NAVY)
        a_.font = font(T_SMALL, False, MUTED)
        n.alignment = align("right", "center", 1)
        a_.alignment = align("right", "top", 1)


FOOTER_1 = "Keine Gewähr für die Richtigkeit der Angaben · ersetzt keine Rechts-, Steuer- oder Finanzberatung · Rechtsstand September 2026"
FOOTER_2 = "MM Holding GmbH · Kornhausgasse 4 · 88250 Weingarten · Amtsgericht Ulm HRB 750729"


def footer(ws, row, c1, c2):
    """Fuß: Oberkante thin LINE2, zwei Zeilen 8 pt grau."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.border = Border(top=side("thin", LINE2))
        c.fill = NOFILL
    a1, a2 = ws.cell(row, col(c1)), ws.cell(row + 1, col(c1))
    set_text(a1, FOOTER_1)
    set_text(a2, FOOTER_2)
    for c in (a1, a2):
        c.font = font(T_MICRO, False, MUTED)
        c.alignment = align("left", "center")
    set_height(ws, row, 18)
    set_height(ws, row + 1, 14)


def hide_rows(ws, r1, r2):
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].hidden = True


def hide_cols(ws, c1, c2):
    for cc in range(col(c1), col(c2) + 1):
        ws.column_dimensions[L(cc)].hidden = True
