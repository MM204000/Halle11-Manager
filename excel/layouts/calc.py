"""Rechenblätter Steuern, Projektion, Finanzierung, AfA-Vergleich (Agent D).

Jahrestabellen (P1-12, P2-04, P2-22, P3-04, P2-08, P2-10, P2-12):
  zweizeiliger Jahreskopf (Navy „Jahr n“ + Blau „Kalenderjahr“, doppelte Kopfzeile ausgeblendet), L1b-Bänder,
  drei Summenstufen, „davon“-Zeilen und Memo-Zeilen, Haarlinie über B:AQ, 5-Jahres-Raster, Einheiten zentriert,
  Verkaufsjahr / Zinswechsel / negative Ergebnisse / inaktive Blöcke per bedingter Formatierung.
Steuern (P2-23): technische Indizes ausgeblendet, Ja/Nein-Schalter, Exit-Ergebnisband, Sprunglinks.
AfA-Vergleich (P2-28, P1-01): Kurznamen + Rechtsgrundlage, zulässige Variante sichtbar, Kacheln, Jahresköpfe,
  Hinweise mit Stichwort, Diagramm-Anker im Raster.

Nur Darstellung: keine Formel, auf die etwas verweist, wird verändert; neue Formeln sind reine Anzeigeformeln
in leeren Zellen (bzw. Beschriftungen, auf die nichts verweist).
"""
from copy import copy

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.hyperlink import Hyperlink

import core as K
from core import (ACCENT, AMBER, BLUE, GREEN, GREEN_BG, INK, INK2, LINE, LINE2, MUTED, MUTED2, NAVY, NOFILL, NUMFMT,
                  RED, SANS, T_BODY, T_MICRO, T_SMALL, TINT, TINT_XL, WHITE, align, fill, font, side)

# ------------------------------------------------------------------------------------------------ Konstanten
FIRST, LAST = 4, 43                 # Jahresspalten D … AQ (Jahr 1 … 40)
GRID = (8, 13, 18, 23, 28, 33, 38)  # H, M, R, W, AB, AG, AL: rechte Kante nach Jahr 5, 10, … 35
FADED = "D0D5DD"                    # entfallene Werte (Darlehen II ungenutzt, nach Volltilgung, KSt bei Privat)
H_DATA, H_KEY, H_HEAD, H_BANDROW, H_GAP = 16, 18, 20, 20, 12
YEAR_FMT = '"Jahr "0'
EUR, NUM, PCT1, PCT2 = NUMFMT["eur"], NUMFMT["num"], NUMFMT["pct1"], NUMFMT["pct2"]
YESNO = NUMFMT["yesno"]


# ------------------------------------------------------------------------------------------------ Hilfen
def _col(c):
    return K.col(c)


def _ungroup_cols(ws):
    """Gruppierte Spaltenbreiten (<col min max>) in Einzelspalten auflösen, damit Breiten je Spalte setzbar sind."""
    dims = ws.column_dimensions
    for key, d in list(dims.items()):
        lo, hi = d.min, d.max
        if lo and hi and hi > lo:
            for ci in range(lo + 1, hi + 1):
                letter = get_column_letter(ci)
                nd = ColumnDimension(ws, index=letter, width=d.width, hidden=d.hidden, customWidth=d.customWidth,
                                     outlineLevel=d.outlineLevel, collapsed=d.collapsed, min=ci, max=ci)
                nd._style = copy(d._style)
                dims[letter] = nd
            d.max = lo


def _width(ws, letter, w):
    d = ws.column_dimensions[letter]
    ci = _col(letter)
    d.min, d.max = ci, ci
    d.width = w


def _rich(parts):
    """parts: [(text, size, bold, color, italic)] → CellRichText (Calibri)."""
    blocks = []
    for text, size, bold, color, italic in parts:
        blocks.append(TextBlock(InlineFont(rFont=SANS, sz=size, b=bold, i=italic, color=color), text))
    return CellRichText(blocks)


def _cf(ws, ref, formula, fnt=None, fil=None):
    kw = {}
    if fnt is not None:
        kw["font"] = fnt
    if fil is not None:
        kw["fill"] = fil
    ws.conditional_formatting.add(ref, FormulaRule(formula=[formula], **kw))


def cf_rules(ws, ref, rules, base, year_cond=None):
    """Schrift-Regeln [(bedingung, Font)] für einen Bereich, danach eine Auffangregel mit der Grundfarbe.

    - Mit year_cond steht vor jeder Regel eine Kombination (Schrift + Verkaufsjahr-Fläche), damit die
      Spaltenmarkierung auch dort durchläuft, wo nur die erste zutreffende Regel greift (LibreOffice).
    - Die Auffangregel (immer wahr, nur Grundfarbe) ändert in Excel nichts; LibreOffice verliert sonst in
      Zellen mit nicht zutreffender Schrift-Regel den Einzug (Vorschau zeigt versetzte Zahlen)."""
    for cond, fnt in rules:
        if year_cond:
            _cf(ws, ref, f"AND({cond},{year_cond})", fnt, fill(TINT))
        _cf(ws, ref, cond, fnt)
    if year_cond:
        _cf(ws, ref, year_cond, base, fill(TINT))
    _cf(ws, ref, "TRUE", base)


def _label(ws, row, text, col=2):
    c = ws.cell(row, col)
    K.set_text(c, text)
    return c


def _clear_row(ws, row, c1, c2):
    for c in K.iter_cells(ws, c1, row, c2, row):
        c.fill = NOFILL
        c.border = Border()


def _link(ws, c1, row, c2, text, target):
    """Sprunglink innerhalb des Blatts: verbundener Textlink, 9 pt fett Blau."""
    for c in K.iter_cells(ws, c1, row, c2, row):
        c.fill = NOFILL
        c.border = Border()
    K.safe_merge(ws, c1, row, c2, row)
    c = ws.cell(row, _col(c1))
    K.set_text(c, text)
    c.font = font(T_SMALL, True, BLUE)
    c.alignment = align("left", "center", 1)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{ws.title}'!{target}", display=text,
                            tooltip=f"Springen zu {text.lstrip('↓ ').strip()}")


# ------------------------------------------------------------------------------------------------ Seitenkopf
def header(ws, title, subtitle, eyebrow="BERECHNUNG", legend=None, eyebrow_raw=None):
    """P2-05: Eyebrow / H1 / Untertitel. Texte bleiben im fixierten Bereich B:C (keine Überläufe über die
    Fixierlinie); die Legende steht rechts davon ab Spalte D, auf Höhe des Untertitels."""
    ws.row_dimensions[4].height = 8
    K.page_header(ws, "B", "C", eyebrow, title, subtitle)
    if eyebrow_raw:
        K.set_text(ws["B5"], eyebrow_raw)
    for r in (5, 6, 7):
        for c in K.iter_cells(ws, 3, r, LAST, r):
            if c.coordinate not in ("B5", "B6", "B7"):
                c.border = Border()
                c.fill = NOFILL
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 22
    if legend:
        c = ws["D7"]
        c.value = _rich(legend)
        c.font = font(T_SMALL, False, MUTED)
        c.alignment = align("left", "top", 1)


def legend_parts(*items):
    """Legende: [(text, farbe|None, fett)] → Rich-Text-Teile (9 pt) mit „ · “ als Trenner."""
    parts = []
    for i, (text, color, bold) in enumerate(items):
        if i:
            parts.append(("   ·   ", T_SMALL, False, MUTED2, False))
        parts.append((text, T_SMALL, bold, color or MUTED, False))
    return parts


# ------------------------------------------------------------------------------------------------ Jahrestabelle
def year_head(ws, r_year, r_dup, r_cal):
    """Zweizeiliger Kopf: r_year „Jahr n“ (Navy, 9 pt fett weiß), r_dup ausgeblendet, r_cal Kalenderjahr (Blau)."""
    for c in K.iter_cells(ws, 2, r_year, LAST, r_year):
        c.fill = fill(NAVY)
        c.border = Border()
        if c.column >= FIRST:
            c.number_format = YEAR_FMT
            c.font = font(T_SMALL, True, WHITE)
            c.alignment = align("right", "center", 1)
    K.set_text(ws.cell(r_year, 2), "Jahr der Kalkulation")
    ws.cell(r_year, 2).font = font(T_SMALL, True, WHITE)
    ws.cell(r_year, 2).alignment = align("left", "center", 1)
    ws.cell(r_year, 3).value = None
    ws.row_dimensions[r_year].height = H_HEAD
    K.hide_rows(ws, r_dup, r_dup)
    for c in K.iter_cells(ws, 2, r_cal, LAST, r_cal):
        c.fill = fill(BLUE)
        c.border = Border(bottom=side("medium", ACCENT))
        if c.column >= FIRST:
            c.number_format = NUMFMT["year"]
            c.font = font(T_SMALL, False, WHITE)
            c.alignment = align("right", "center", 1)
    K.set_text(ws.cell(r_cal, 2), "Kalenderjahr")
    ws.cell(r_cal, 2).font = font(T_SMALL, False, WHITE)
    ws.cell(r_cal, 2).alignment = align("left", "center", 1)
    ws.row_dimensions[r_cal].height = H_HEAD


def band(ws, row, title, c2=LAST, years_from=None, right=None):
    """L1b-Tabellenabschnitt über die volle Tabellenbreite; optional Kalenderjahre (=D$42) in D:AQ."""
    for c in K.iter_cells(ws, 2, row, c2, row):
        c.font = font(T_SMALL, True, NAVY)
    if isinstance(title, str) and title.startswith("="):
        K.band_l1b(ws, row, 2, c2, None, right=right)
        ws.cell(row, 2).value = title
    else:
        K.band_l1b(ws, row, 2, c2, title, right=right)
    if years_from:
        for cc in range(FIRST, LAST + 1):
            c = ws.cell(row, cc)
            L_ = get_column_letter(cc)
            c.value = f"={L_}${years_from}"
            c.number_format = NUMFMT["year"]
            c.font = font(T_MICRO, True, BLUE)
            c.alignment = align("right", "center", 1)
    ws.row_dimensions[row].height = H_BANDROW


def data_row(ws, row, kind="data", fmt=NUM, label=None, unit=True):
    """kind: data · davon · memo · t1 · t2 · t3 (Summenstufen nach core.total)."""
    if label is not None:
        _label(ws, row, label)
    b, cu = ws.cell(row, 2), ws.cell(row, 3)
    for c in K.iter_cells(ws, 2, row, LAST, row):
        c.fill = NOFILL
        c.border = Border(bottom=side("hair", LINE))
        if c.column >= FIRST:
            c.font = font(T_BODY, False, INK)
            c.alignment = align("right", "center", 1)
            c.number_format = fmt
    b.font = font(T_BODY, False, INK)
    b.alignment = align("left", "center", 1)
    if unit:
        cu.font = font(T_SMALL, False, MUTED)
        cu.alignment = align("center", "center")
    h = H_DATA
    if kind == "davon":
        b.font = font(T_SMALL, False, MUTED)
        b.alignment = align("left", "center", 2)
    elif kind == "memo":
        K.memo(ws, row, 2, LAST)
        for c in K.iter_cells(ws, 2, row, LAST, row):
            c.border = Border(bottom=side("hair", LINE))
        b.alignment = align("left", "center", 1)
    elif kind in ("t1", "t2", "t3"):
        K.total(ws, row, 2, LAST, {"t1": 1, "t2": 2, "t3": 3}[kind])
        if unit:
            cu.font = font(T_SMALL, False, MUTED)
        if kind == "t3":
            h = H_KEY
    ws.row_dimensions[row].height = h


def memo_gap(ws, row):
    """Erste Memo-Zeile einer Gruppe: +4 pt Abstand nach oben (Text unten ausgerichtet)."""
    ws.row_dimensions[row].height = H_DATA + 4
    for c in K.iter_cells(ws, 2, row, LAST, row):
        al = c.alignment
        c.alignment = align(al.horizontal or "left", "bottom", al.indent or 0)


def gap(ws, row, h=H_GAP):
    _clear_row(ws, row, 2, LAST)
    ws.row_dimensions[row].height = h


def grid(ws, rows):
    """5-Jahres-Raster: rechte Kante thin D5D9DE an H, M, R, W, AB, AG, AL (Köpfe, Bänder, Daten)."""
    for r in rows:
        for cc in GRID:
            c = ws.cell(r, cc)
            b = c.border
            c.border = Border(left=b.left, top=b.top, bottom=b.bottom, right=side("thin", LINE2))


def footer(ws, row, merge_to="M", line_to=LAST):
    """Fuß auf dem Tabellenraster: Oberkante über die Tabellenbreite, Text über B:merge_to verbunden."""
    for r in (row, row + 1):
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row == r:
                ws.unmerge_cells(str(mr))
        for c in K.iter_cells(ws, 3, r, line_to, r):
            c.value = None
    K.footer(ws, row, 2, line_to)
    K.safe_merge(ws, "B", row, merge_to, row)
    K.safe_merge(ws, "B", row + 1, merge_to, row + 1)
    ws.cell(row, 2).alignment = align("left", "center")
    ws.cell(row + 1, 2).alignment = align("left", "center")


def year_marks(ws, yr_row, head_ref, body_ref, band_refs=()):
    """Verkaufsjahr: Spalte E7EEF7, im Kopf 4A86C8 (bedingte Formatierung, nur Anzeige)."""
    L0 = get_column_letter(FIRST)
    cond = f"{L0}${yr_row}=Haltedauer"
    _cf(ws, head_ref, cond, None, fill(ACCENT))
    for ref in band_refs:
        _cf(ws, ref, cond, Font(color=WHITE, bold=True), fill(ACCENT))
    return cond


def negatives(ws, rows, year_cond):
    """Rot B42318 für negative Werte in Ergebniszeilen; rows: {zeile: grundfarbe}."""
    for r, base in rows.items():
        cf_rules(ws, f"D{r}:AQ{r}", [(f"D{r}<0", Font(color=RED))], Font(color=base), year_cond)


# ------------------------------------------------------------------------------------------------ Projektion
def projektion(ws):
    _ungroup_cols(ws)
    _width(ws, "C", 12)
    header(ws, "Projektion", "Miete, Kosten, Cashflow und Vermögen über 40 Jahre · Beträge in € pro Jahr",
           legend=legend_parts(("Jahr 1 = erste 12 Monate ab Kaufdatum", None, False),
                               ("hellblau hinterlegt = Verkaufsjahr", BLUE, True),
                               ("rot = negatives Ergebnis", RED, True),
                               ("Steuer: + Zahlung / – Erstattung", None, False)))
    year_head(ws, 8, 9, 10)
    gap(ws, 11, 10)
    rows = {
        12: ("band", "MIETEINNAHMEN"),
        13: ("data", None), 14: ("data", None), 15: ("t2", None),
        16: ("memo", "Umlagefähige Betriebskosten (Durchlaufposten)"), 17: ("memo", "Warmmiete"),
        18: ("gap",),
        19: ("band", "NICHT UMLAGEFÄHIGE BEWIRTSCHAFTUNGSKOSTEN (ZAHLUNGSABFLUSS)"),
        20: ("data", None), 21: ("data", "Zuführung Erhaltungsrücklage WEG"), 22: ("data", None),
        23: ("data", None), 24: ("data", None), 25: ("data", None), 26: ("t2", None),
        27: ("gap",),
        28: ("band", "CASHFLOW"),
        29: ("t2", "= Einnahmenüberschuss vor Finanzierung (NOI)"),
        30: ("davon", "davon Zinsen"), 31: ("davon", "davon Tilgung (inkl. Sondertilgung)"),
        32: ("t1", "– Kapitaldienst"), 33: ("t2", None),
        34: ("memo", "Info: Steuerliches Ergebnis"), 35: ("data", None), 36: ("t3", None),
        37: ("data", None), 38: ("memo", None),
        39: ("gap",),
        40: ("band", "VERMÖGENSENTWICKLUNG"),
        41: ("data", None), 42: ("data", "– Restschuld Jahresende"), 43: ("t3", "= Nettovermögen"),
        44: ("memo", "Vermögenszuwachs im Jahr"), 45: ("data", None), 46: ("memo", None),
        47: ("gap",),
        48: ("band", "EXIT-CASHFLOW (BASIS FÜR DIE EIGENKAPITALRENDITE / IRR)"),
        49: ("data", None),
        50: ("gap",),
    }
    fmts = {45: PCT1}
    render(ws, rows, fmts, memo_first=(16, 34, 38, 44, 46))
    # Zeile 48/49: „Jahr 0“ als Etikett im Band, Wert in C rechts
    c48 = ws["C48"]
    K.set_text(c48, "Jahr 0")
    c48.font = font(T_MICRO, True, BLUE)
    c48.alignment = align("right", "center", 1)
    ws["B49"].value = _rich([("Cashflow inkl. Nettoerlös im Verkaufsjahr", T_BODY, False, INK, False),
                             ("   Jahr 0 = Eigenkapital", T_SMALL, False, MUTED, False)])
    c49 = ws["C49"]
    c49.number_format = NUM
    c49.font = font(T_BODY, False, INK)
    c49.alignment = align("right", "center", 1)
    footer(ws, 51)
    table_rows = [8, 10] + [r for r, s in rows.items() if s[0] != "gap"]
    grid(ws, table_rows)
    ws.freeze_panes = "D11"
    # Bedingte Formatierung (nur Anzeige)
    ycond = year_marks(ws, 9, "D8:AQ10", None)
    negatives(ws, {33: NAVY, 36: NAVY, 38: MUTED, 43: NAVY}, ycond)
    _cf(ws, "D49:AQ49", ycond, Font(bold=True, color=NAVY), fill(TINT))
    _cf(ws, "D49:AQ49", "TRUE", Font(color=INK))
    _cf(ws, "D11:AQ49", ycond, None, fill(TINT))


def render(ws, rows, fmts, memo_first=()):
    for r, spec in rows.items():
        kind = spec[0]
        if kind == "gap":
            gap(ws, r)
        elif kind == "band":
            band(ws, r, spec[1], years_from=spec[2] if len(spec) > 2 else None)
        else:
            data_row(ws, r, kind, fmts.get(r, NUM), spec[1])
    for r in memo_first:
        memo_gap(ws, r)


# ------------------------------------------------------------------------------------------------ Finanzierung
def finanzierung(ws):
    _ungroup_cols(ws)
    _width(ws, "C", 12)
    header(ws, "Finanzierung", "Tilgungsverlauf Darlehen I und II · Annuitätendarlehen mit konstanter Rate",
           legend=legend_parts(("nach der Zinsbindung gilt der Anschlusszins (fett)", None, False),
                               ("hellblau hinterlegt = Verkaufsjahr", BLUE, True),
                               ("grau = entfällt (getilgt bzw. nicht genutzt)", MUTED2, True)))
    year_head(ws, 8, 9, 10)
    gap(ws, 11, 10)
    rows = {
        12: ("band", "DARLEHEN I"),
        13: ("data", None), 14: ("data", None), 15: ("data", None), 16: ("data", None),
        17: ("data", "– Reguläre Tilgung"), 18: ("data", "– Sondertilgung"),
        19: ("t2", "= Restschuld Jahresende"), 20: ("t1", "Kapitaldienst (Zins + Tilgung + Sondertilgung)"),
        21: ("gap",),
        22: ("band", '="DARLEHEN II"&IF(Darlehen_II=0,"  ·  NICHT GENUTZT","")'),
        23: ("data", None), 24: ("data", None), 25: ("data", None), 26: ("data", None),
        27: ("data", "– Reguläre Tilgung"), 28: ("data", "– Sondertilgung"),
        29: ("data", "– Tilgungszuschuss (KfW, kein Zahlungsabfluss)"),
        30: ("t2", "= Restschuld Jahresende"), 31: ("t1", "Kapitaldienst (Zins + Tilgung + Sondertilgung)"),
        32: ("gap",),
        33: ("band", "SUMME DARLEHEN"),
        34: ("data", None), 35: ("data", None), 36: ("data", None),
        37: ("t1", "= Kapitaldienst gesamt"), 38: ("t3", "= Restschuld Jahresende"),
        39: ("memo", None), 40: ("memo", None),
        41: ("gap",),
    }
    fmts = {14: PCT2, 24: PCT2, 39: PCT2}
    render(ws, rows, fmts, memo_first=(39,))
    footer(ws, 42)
    table_rows = [8, 10] + [r for r, s in rows.items() if s[0] != "gap"]
    grid(ws, table_rows)
    ws.freeze_panes = "D11"
    ycond = year_marks(ws, 9, "D8:AQ10", None)
    faded = Font(color=FADED, bold=False)
    switch = Font(bold=True, color=BLUE)          # Zinswechsel nach der Zinsbindung
    # entfallene Werte grau (Darlehen I nach Volltilgung; Darlehen II ungenutzt)
    for r in range(14, 19):
        rules = [("D$13=0", faded)] + ([("AND(ISNUMBER(C14),D14<>C14)", switch)] if r == 14 else [])
        cf_rules(ws, f"D{r}:AQ{r}", rules, Font(color=INK), ycond)
    for r in range(23, 32):
        rules = [("Darlehen_II=0", faded)] + ([("AND(ISNUMBER(C24),D24<>C24)", switch)] if r == 24 else [])
        cf_rules(ws, f"D{r}:AQ{r}", rules, Font(color={30: NAVY}.get(r, INK)), ycond)
    _cf(ws, "D11:AQ40", ycond, None, fill(TINT))


# ------------------------------------------------------------------------------------------------ Steuern
STEUERN_LABELS = {
    17: "Wohngebäude", 21: "Sofortige Verlustverrechnung", 23: "Disagio sofort abziehbar",
    28: "Degressive AfA aktiv", 30: "AfA-Bemessungsgrundlage Gebäude (ohne Sanierungsanteil)",
    32: "Sonder-AfA § 7b aktiv", 35: "Renovierung als Herstellungskosten aktivieren",
    85: "Veräußerungsgewinn (Preis – Verkaufskosten – Buchwert)",
    86: "Veräußerungsgewinn steuerpflichtig", 94: "Gesamtertrag nach Steuern",
    98: "Summe Gebäude-AfA Jahre 1–10 (Standard nach Baujahr)",
    99: "Steuereffekt der gewählten AfA-Variante (10 Jahre)",
}
STEUERN_NOTES = {
    24: "Fallback auf „Automatisch“, wenn Voraussetzungen fehlen (siehe Prüfhinweise im Cockpit)",
    88: "nur bei Option „Verlustvortrag“ bzw. GmbH; Mindestbesteuerung (§ 10d Abs. 2 EStG) nicht modelliert",
    94: "Cashflows + Nettoerlös – Eigenkapital",
    99: "gegenüber Standard-AfA · positiv = Steuerersparnis (Liquiditätsvorteil, keine endgültige Ersparnis)",
}
STEUERN_FMT = {
    17: YESNO, 18: PCT2, 19: PCT2, 20: PCT2, 21: YESNO, 22: '[=1]0" Jahr";0" Jahre"', 23: YESNO,
    25: PCT2, 26: PCT2, 27: "0.0", 28: YESNO, 31: '#,##0" €/m²"', 32: YESNO, 35: YESNO,
    78: NUMFMT["year"], 86: YESNO, 87: PCT2, 95: NUMFMT["mult2"], 96: PCT1,
}


def kv_rows(ws, rows, last="M"):
    """Parameter-/Exit-Block: Beschriftung B, Wert C rechts, Hinweis E:M grau; Haarlinie, keine Seitenrahmen."""
    for r in rows:
        if r in STEUERN_LABELS:
            _label(ws, r, STEUERN_LABELS[r])
        if r in STEUERN_NOTES:
            K.set_text(ws.cell(r, 5), STEUERN_NOTES[r])
        for c in K.iter_cells(ws, 2, r, last, r):
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE))
        b, v, n = ws.cell(r, 2), ws.cell(r, 3), ws.cell(r, 5)
        b.font = font(T_BODY, False, INK)
        b.alignment = align("left", "center", 1)
        v.font = font(T_BODY, False, INK)
        v.alignment = align("right", "center", 1)
        v.number_format = STEUERN_FMT.get(r, EUR)
        n.font = font(T_SMALL, False, MUTED)
        n.alignment = align("left", "center", 1)
        ws.row_dimensions[r].height = H_DATA


def result_band(ws, row, top=False, last="M"):
    """Exit-Ergebnis: F3F7FC über B:M, Akzentkante links, Wertzelle E7EEF7 in 11 pt fett Navy, 20 pt."""
    for c in K.iter_cells(ws, 2, row, last, row):
        c.fill = fill(TINT_XL)
        c.border = Border(left=side("thick", ACCENT) if c.column == 2 else None,
                          top=side("thin", NAVY) if top else None, bottom=side("hair", LINE))
    b, v = ws.cell(row, 2), ws.cell(row, 3)
    b.font = font(T_BODY, True, NAVY)
    v.fill = fill(TINT)
    v.font = font(11, True, NAVY)
    ws.cell(row, 5).fill = fill(TINT_XL)
    ws.row_dimensions[row].height = 20


def steuern(ws):
    _ungroup_cols(ws)
    _width(ws, "C", 12)
    header(ws, "Steuern & AfA", "Abgeleitete Parameter, AfA-Verlauf und steuerliches Ergebnis",
           legend=legend_parts(("hellblau hinterlegt = Verkaufsjahr", BLUE, True),
                               ("rot = negatives Ergebnis", RED, True),
                               ("Steuer: + Zahlung / – Erstattung (Verrechnung mit anderen Einkünften)", None, False)))
    ws.row_dimensions[5].height = 18
    _link(ws, "D", 6, "E", "↓ Jahrestabelle", "B40")
    _link(ws, "F", 6, "G", "↓ Exit-Ergebnis", "B77")
    # ---- Parameterblock
    band(ws, 8, "ABGELEITETE STEUERLICHE PARAMETER", c2=13, right="technische Indizes ausgeblendet")
    K.hide_rows(ws, 9, 16)
    K.hide_rows(ws, 24, 24)
    kv = [r for r in range(17, 38) if r != 24]
    kv_rows(ws, kv)
    for c in K.iter_cells(ws, 2, 37, 13, 37):
        c.border = Border(bottom=side("thin", NAVY))
    gap(ws, 38, 12)
    gap(ws, 39, 12)
    # ---- Jahrestabelle
    year_head(ws, 40, 41, 42)
    gap(ws, 43, 10)
    rows = {
        44: ("band", "ABSCHREIBUNG (AfA)", 42),
        45: ("data", None), 46: ("data", None), 47: ("data", None), 48: ("data", None), 49: ("data", None),
        50: ("data", None), 51: ("t1", None), 52: ("memo", None), 53: ("memo", None), 54: ("data", None),
        55: ("data", None), 56: ("t2", None), 57: ("memo", None),
        58: ("gap",),
        59: ("band", "STEUERLICHES ERGEBNIS AUS DEM OBJEKT", 42),
        60: ("data", None),
        61: ("davon", "davon Bewirtschaftungskosten (steuerlich)"), 62: ("davon", "davon Schuldzinsen"),
        63: ("davon", "davon Finanzierungsnebenkosten und Disagio"),
        64: ("davon", "davon Erhaltungsaufwand (ggf. verteilt, § 82b EStDV)"),
        65: ("davon", "davon Abschreibungen"),
        66: ("t1", "– Werbungskosten / Betriebsausgaben gesamt"), 67: ("t3", "= Steuerliches Ergebnis"),
        68: ("memo", None), 69: ("data", None), 70: ("memo", None),
        71: ("data", "Körperschaftsteuersatz im Kalenderjahr (nur GmbH)"), 72: ("data", "Angewendeter Steuersatz (Grenzsatz bzw. KSt + Soli + GewSt)"),
        73: ("t3", "= Steuer (+ Zahlung / – Erstattung)"), 74: ("memo", None),
    }
    render(ws, rows, {71: PCT2, 72: PCT2}, memo_first=(52, 57, 68, 70, 74))
    table_rows = [40, 42] + [r for r, s in rows.items() if s[0] != "gap"]
    grid(ws, table_rows)
    gap(ws, 75, 12)
    gap(ws, 76, 12)
    # ---- Exit-Block
    band(ws, 77, "VERKAUFSSZENARIO (EXIT) · STEUER UND GESAMTRENDITE", c2=13)
    kv_rows(ws, range(78, 100))
    result_band(ws, 91, top=True)
    for r in (94, 95, 96):
        result_band(ws, r)
    for c in K.iter_cells(ws, 2, 99, 13, 99):
        b = c.border
        c.border = Border(left=b.left, bottom=side("thin", NAVY))
    gap(ws, 100, 12)
    footer(ws, 101, line_to=13)
    # ---- Bedingte Formatierung
    ycond = "D$41=Haltedauer"
    _cf(ws, "D40:AQ42", ycond, None, fill(ACCENT))
    for ref in ("D44:AQ44", "D59:AQ59"):
        _cf(ws, ref, ycond, Font(color=WHITE, bold=True), fill(ACCENT))
        _cf(ws, ref, "TRUE", Font(color=BLUE))
    cf_rules(ws, "D71:AQ71", [("Rechtsform_Idx=1", Font(color=FADED))], Font(color=INK), ycond)
    negatives(ws, {67: NAVY, 69: INK}, ycond)
    _cf(ws, "D43:AQ74", ycond, None, fill(TINT))


# ------------------------------------------------------------------------------------------------ AfA-Vergleich
AFA_VARIANTS = {
    19: ("Linear 2,0 %", "§ 7 Abs. 4 S. 1 Nr. 2b EStG · Fertigstellung 1925–2022"),
    20: ("Linear 2,5 %", "§ 7 Abs. 4 S. 1 Nr. 2c EStG · Fertigstellung vor 1925"),
    21: ("Linear 3,0 %", "§ 7 Abs. 4 S. 1 Nr. 2a EStG · ab 2023 / BV-Wirtschaftsgebäude"),
    22: ("Gutachten (RND)", "§ 7 Abs. 4 S. 2 EStG · tatsächliche Restnutzungsdauer"),
    23: ("Degressiv 5 %", "§ 7 Abs. 5a EStG · Wohngebäude, Baubeginn 10/2023–9/2029"),
    24: ("Linear 3 % + § 7b", "5 % p. a. in Jahren 1–4, danach Restwert-AfA (§ 7a Abs. 9)"),
    25: ("Degressiv + § 7b", "kombiniert; die Sonder-AfA mindert den Restwert"),
    26: ("Im Modell angewendet", "lt. Blatt „Steuern“ · inkl. Sonder-AfA, ohne § 7h/7i"),
}
AFA_BASICS = {
    9: "AfA-Basis Gebäude (ohne § 7h/7i-Anteil)",
    10: "Sonder-AfA § 7b – Basis (max. 4.000 €/m²)",
    11: "Sonder-AfA § 7b p. a. (5 %)",
    12: "Steuersatz Jahr 1 (angewendet)",
    13: "Kalkulationszins für den Barwert",
    14: "Restnutzungsdauer lt. Gutachten",
    15: "Baujahr · Objektart",
}
AFA_HINTS = {
    47: ("Liquiditätseffekt: ", "Eine höhere AfA in den ersten Jahren ist ein Liquiditäts- und Zinsvorteil "
         "(Steuerstundung), keine endgültige Ersparnis – bei Verkauf innerhalb der 10-Jahres-Frist (Privat) bzw. "
         "immer (GmbH) mindert die kumulierte AfA den Buchwert und erhöht den steuerpflichtigen Veräußerungsgewinn."),
    48: ("Prüfumfang: ", "Die Spalte „Anwendbar“ prüft nur die im Tool erfassten Kriterien (Baujahr, Objektart, "
         "Baukosten je m²). Nicht geprüft werden: Baubeginn im Zeitfenster (§ 7 Abs. 5a), Bauantrag/Bauanzeige, "
         "Effizienzhaus-40-Standard mit QNG-Siegel (§ 7b), Anschaffung bis Ende des Fertigstellungsjahres (§ 7b)."),
    49: ("Kombination: ", "Degressive AfA und § 7b sind kombinierbar (§ 7b Abs. 1 S. 1 i. V. m. § 7 Abs. 5a EStG); "
         "im Tool wird die Sonder-AfA vereinfachend vom Restwert abgezogen (konservativ). Ein Wechsel von degressiv "
         "zu linear (Restwert / Restnutzungsdauer) erfolgt automatisch, sobald vorteilhaft – innerhalb der ersten "
         "10 Jahre bei 3 % typisierter Nutzungsdauer nicht."),
    50: ("§ 7h/7i: ", "Erhöhte Absetzungen nach § 7h / § 7i EStG (9 % / 7 %) betreffen nur den bescheinigten "
         "Sanierungsanteil und sind hier nicht als Variante dargestellt; sie werden im Blatt „Steuern“ zusätzlich "
         "zur regulären AfA der Altsubstanz berechnet."),
}
AFA_LAST = 17  # Q
EMU_PT = 12700


def afa(ws):
    _ungroup_cols(ws)
    _width(ws, "B", 40)
    _width(ws, "C", 30)
    _width(ws, "P", 15)
    _width(ws, "Q", 15)
    header(ws, "AfA-Vergleich", "Abschreibungsvarianten für dieses Objekt in den Jahren 1–10",
           eyebrow_raw="BERECHNUNG  ›  STEUERN & AfA",
           legend=legend_parts(("grün = für dieses Objekt anwendbar", GREEN, True),
                               ("orange = nur mit Gutachten", AMBER, True),
                               ("grau = nicht anwendbar", MUTED2, True)))
    ws.row_dimensions[5].height = 18
    # ---- Grundlagen + Kacheln
    band(ws, 8, "GRUNDLAGEN", c2=AFA_LAST)
    fmts = {9: EUR, 10: EUR, 11: EUR, 12: PCT2, 13: PCT2, 14: '0" Jahre";;"– (kein Gutachten)"', 15: "@"}
    for r in range(9, 16):
        _label(ws, r, AFA_BASICS[r])
        _clear_row(ws, r, 2, AFA_LAST)
        for c in K.iter_cells(ws, 2, r, 3, r):
            c.border = Border(bottom=side("hair", LINE))
        b, v = ws.cell(r, 2), ws.cell(r, 3)
        b.font = font(T_BODY, False, INK)
        b.alignment = align("left", "center", 1)
        v.font = font(T_BODY, False, INK)
        v.alignment = align("right", "center", 1)
        v.number_format = fmts[r]
        ws.row_dimensions[r].height = K.H_ROW
    ws["C15"].value = '=Baujahr&" · "&Objektart'
    tiles(ws)
    gap(ws, 16, 12)
    # ---- Variantentabelle
    band(ws, 17, "AfA-BETRÄGE JE JAHR (€)", c2=AFA_LAST,
         right="Steuerersparnis = AfA × Steuersatz Jahr 1 · Barwert mit dem Kalkulationszins")
    heads = {"B": ("VARIANTE", "left"), "C": ("ANWENDBAR FÜR DIESES OBJEKT?", "left"),
             "N": ("Σ JAHRE 1–4", "right"), "O": ("Σ JAHRE 1–10", "right"),
             "P": ("STEUERERSPARNIS\n1–10 NOMINAL", "right"), "Q": ("BARWERT\nSTEUERERSPARNIS", "right")}
    for k in range(10):
        heads[get_column_letter(4 + k)] = (f"Jahr {k + 1}", "right")
    K.table_head(ws, 18, 2, AFA_LAST, heads, height=30)
    for cc in range(2, AFA_LAST + 1):
        c = ws.cell(18, cc)
        c.alignment = align(c.alignment.horizontal, "center", 1, wrap=cc >= 16)
    for r in range(19, 27):
        name, basis = AFA_VARIANTS[r]
        key = r == 26
        b = ws.cell(r, 2)
        b.value = _rich([(name, T_BODY, True, NAVY if key else INK, False),
                         ("\n" + basis, T_MICRO, False, MUTED, False)])
        for c in K.iter_cells(ws, 2, r, AFA_LAST, r):
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE))
            if c.column >= 4:
                c.font = font(T_BODY, c.column in (15, 17), INK)
                c.alignment = align("right", "center", 1)
                c.number_format = NUM
        b.font = font(T_BODY, True, INK)
        b.alignment = align("left", "center", 1, wrap=True)
        cst = ws.cell(r, 3)
        cst.font = font(T_SMALL, False, MUTED)
        cst.alignment = align("left", "center", 1, wrap=True)
        ws.row_dimensions[r].height = 30
    K.total(ws, 26, 2, AFA_LAST, 3)
    ws.row_dimensions[26].height = 30
    ws["C26"].font = font(T_SMALL, True, NAVY)
    ws["C26"].alignment = align("left", "center", 1, wrap=True)
    gap(ws, 27, 14)
    # ---- Nebenrechnungen
    for head_row, title, first in ((28, "NEBENRECHNUNG: RESTWERT GEBÄUDE AM JAHRESENDE (€)", 29),
                                   (37, "NEBENRECHNUNG: KUMULIERTE AfA JE VARIANTE (€)", 38)):
        heads = {"B": (title, "left")}
        for k in range(10):
            heads[get_column_letter(4 + k)] = (f"Jahr {k + 1}", "right")
        K.table_head(ws, head_row, 2, AFA_LAST, heads)
        for i in range(7):
            r = first + i
            if head_row == 28:
                _label(ws, r, AFA_VARIANTS[19 + i][0])
            else:
                ws.cell(r, 2).value = f"=$S{19 + i}"
            for c in K.iter_cells(ws, 2, r, AFA_LAST, r):
                c.fill = NOFILL
                c.border = Border(bottom=side("hair", LINE))
                c.font = font(T_SMALL, False, MUTED)
                if c.column >= 4:
                    c.alignment = align("right", "center", 1)
                    c.number_format = NUM
            ws.cell(r, 2).alignment = align("left", "center", 1)
            ws.row_dimensions[r].height = H_DATA
    gap(ws, 36, 14)
    gap(ws, 45, 14)
    # ---- Hinweise (Satzbreite B:C, fettes Stichwort)
    band(ws, 46, "HINWEISE ZUR INTERPRETATION", c2=AFA_LAST)
    for r, (kw, text) in AFA_HINTS.items():
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row == r:
                ws.unmerge_cells(str(mr))
        _clear_row(ws, r, 2, AFA_LAST)
        K.safe_merge(ws, "B", r, "C", r)
        c = ws.cell(r, 2)
        c.value = _rich([(kw, T_SMALL, True, NAVY, False), (text, T_SMALL, False, INK2, False)])
        c.font = font(T_SMALL, False, INK2)
        c.alignment = align("left", "center", 1, wrap=True)
        n = K.lines_needed(kw + text, K.span_px(ws, "B", "C") - 12, 9 * 1.12)
        ws.row_dimensions[r].height = n * 12 + 10
    gap(ws, 51, 14)
    # ---- Diagramme
    K.band_l1(ws, 52, 2, AFA_LAST, "Diagramme")
    for r in range(53, 70):
        ws.row_dimensions[r].height = 15
    place_charts(ws)
    gap(ws, 70, 15)
    K.hide_rows(ws, 71, 72)
    footer(ws, 73, merge_to="Q", line_to=AFA_LAST)
    K.hide_cols(ws, "R", "S")
    # ---- Bedingte Formatierung: zulässige Variante sichtbar
    _cf(ws, "C19:C25", 'LEFT($C19,2)="Ja"', Font(color=GREEN, bold=True), fill(GREEN_BG))
    _cf(ws, "C19:C25", 'LEFT($C19,4)="Nur "', Font(color=AMBER, bold=True))
    cf_rules(ws, "C19:C25", [('LEFT($C19,4)="Nein"', Font(color=MUTED2, bold=False))], Font(color=MUTED))
    cf_rules(ws, "D19:Q25", [('LEFT($C19,4)="Nein"', Font(color=MUTED2, bold=False))], Font(color=INK))


def tiles(ws):
    """Kacheln rechts neben den Grundlagen (reine Verweise): angewendete Methode, AfA 1–10, Barwert."""
    specs = (("E", "H", "Im Modell angewendet", "=C26", None, "Methode lt. Blatt „Steuern“", K.T_H3),
             ("J", "M", "AfA Jahre 1–10 (angewendet)", "=O26", EUR, "Summe der Gebäude-AfA", K.T_KPI),
             ("O", "Q", "Barwert der Steuerersparnis", "=Q26", EUR, "Jahre 1–10, abgezinst", K.T_KPI))
    for c1, c2, lab, val, fmt, sub, size in specs:
        K.kpi_tile(ws, c1, c2, 9, 10, 12, label=lab, value=val, sub=sub, fmt=fmt, value_rows=2, gap_right=False,
                   value_size=size)
        v = ws[f"{c1}10"]
        v.alignment = align("left", "center", 1, wrap=size < K.T_KPI)
        ws.row_dimensions[12].height = K.H_ROW
    ws.row_dimensions[9].height = K.H_ROW


def place_charts(ws):
    """chart33 (Balken) exakt B53 bis Ende C (Fixierlinie frei), chart34 (Linien) E53 bis Ende Q;
    Höhe 3.059.280 EMU wie auf „Diagramme“ (16 Zeilen à 15 pt + Rest)."""
    rest = 3059280 - 16 * 15 * EMU_PT
    for ch in ws._charts:
        kind = type(ch).__name__
        if kind.startswith("Bar"):
            c1, c2 = 1, 3
        else:
            c1, c2 = 4, 17
        ch.anchor = TwoCellAnchor(editAs="twoCell", _from=AnchorMarker(col=c1, colOff=0, row=52, rowOff=0),
                                  to=AnchorMarker(col=c2, colOff=0, row=68, rowOff=rest))


# ------------------------------------------------------------------------------------------------ Einstieg
def apply(wb):
    for name, fn in (("Projektion", projektion), ("Finanzierung", finanzierung), ("Steuern", steuern),
                     ("AfA-Vergleich", afa)):
        if name in wb.sheetnames:
            fn(wb[name])
