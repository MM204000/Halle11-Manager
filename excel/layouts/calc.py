"""Rechenblätter Steuern, Projektion, Finanzierung, AfA-Vergleich (Agent D) – Runde 2: EINE Komponentensprache.

Alle Bausteine kommen aus core.py (CORE_API.md):
  Seitenkopf     C.page_header (Überzeile „REITER  ›  BLATT“, H1 22 pt, Untertitel 10 pt) + Legende mit Marken ● ■
  Abschnitte     C.section(level=1) – E7EEF7, Akzentkante, 12,5 pt Titelschreibung, Meta rechts 8 pt
                 C.section(level=2) – Tabellenköpfe (EEF3FA, 8 pt Versalien)
  Summenstufen   C.sum_row('deduct' | 'sub' | 'result')                                     (P1-14)
  Kacheln        C.tile(variant='light')                                                     (P1-11)
  Links          C.link_loc / C.link_row, „↑ Übersicht“ als Zell-Link                        (P3-03)
  Formate        C.NUMFMT (pct1/pct2/years_n/status_dot …)                                   (P2-12)
  CF-Abschluss   C.cf_close                                                                  (LibreOffice-Einzug)
Blattspezifisch: Jahreskopf (Navy/Blau), 5-Jahres-Raster, Verkaufsjahr-Spalte, Rot-Regel für Ergebniszeilen
(P2-01), „entfällt“ = „–“ in 98A2B3 kursiv inkl. Beschriftung (P2-02), Exit-Zeile Jahr 0 (P2-19),
Gliederung auf Steuern (P3-04), gleichmäßiger Zeilenrhythmus 16 pt (P3-05).

Nur Darstellung: keine Formel, auf die etwas verweist, wird verändert; neue Formeln sind reine Anzeigeformeln
in leeren Zellen (bzw. Beschriftungen, auf die nichts verweist).
"""
from copy import copy

from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Border, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import Outline

import core as C
from core import (ACCENT, AMBER, BLUE, GREEN, INK, INK2, LINE, LINE2, MUTED, NAVY, NOFILL, NUMFMT, RED,
                  T_BODY, T_MICRO, T_SMALL, TINT, WHITE, align, fill, font, side)

# ------------------------------------------------------------------------------------------------ Konstanten
FIRST, LAST = 4, 43                 # Jahresspalten D … AQ (Jahr 1 … 40)
GRID = (8, 13, 18, 23, 28, 33, 38)  # H, M, R, W, AB, AG, AL: rechte Kante nach Jahr 5, 10, … 35
FADED = "98A2B3"                    # „entfällt“ (Darlehen II ungenutzt, nach Volltilgung, KSt bei Privat) – P2-02
H_DATA, H_YEAR, H_GAP = 16, 20, C.H_GAP
YEAR_FMT = '"Jahr "0'
EUR, NUM, PCT1, PCT2 = NUMFMT["eur"], NUMFMT["num"], NUMFMT["pct1"], NUMFMT["pct2"]
YESNO = NUMFMT["yesno"]
DASH_FMT = '"–";"–";"–";"–"'        # bedingtes Zahlenformat „entfällt“
# P24: Steuersätze überall zweistellig (46,26 %). Variante mit führendem „#“ = gleiche Anzeige wie NUMFMT["pct2"];
# global_rules.numfmt_catalog erkennt sie nicht als „schlichtes Prozentformat“ und setzt sie nicht auf pct1 zurück.
PCT2K = '#0.00 %;"−"#0.00 %;"–"'
EYEBROW = {"Projektion": "PROJEKTION  ›  40 JAHRE", "Finanzierung": "FINANZIERUNG  ›  TILGUNGSPLAN",
           "Steuern": "STEUERN  ›  STEUERN & AFA", "AfA-Vergleich": "STEUERN  ›  AFA-VERGLEICH"}


# ------------------------------------------------------------------------------------------------ Hilfen
def _col(c):
    return C.col(c)


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


def _rule(ws, ref, formula, fnt=None, fil=None, numfmt=None, border=None):
    """Bedingte Formatierung (Ausdruck) mit optionalem bedingtem Zahlenformat (dxf numFmt)."""
    nf = NumberFormat(numFmtId=300, formatCode=numfmt) if numfmt else None
    dxf = DifferentialStyle(font=fnt, fill=fil, numFmt=nf, border=border)
    r = Rule(type="expression", dxf=dxf)
    r.formula = [formula]
    ws.conditional_formatting.add(ref, r)


def cf_rules(ws, ref, rules, base, year_cond=None, year_fill=TINT):
    """Regeln [(bedingung, Font, numfmt|None[, fill])] für einen Bereich, danach eine Auffangregel mit der Grundfarbe.

    - Mit year_cond steht vor jeder Regel eine Kombination (Regel + Verkaufsjahr-Fläche), damit die
      Spaltenmarkierung auch dort durchläuft, wo nur die erste zutreffende Regel greift (LibreOffice).
    - Die Auffangregel (immer wahr, nur Grundfarbe) ändert in Excel nichts; LibreOffice verliert sonst in
      Zellen mit nicht zutreffender Regel den Einzug (Vorschau zeigt versetzte Zahlen)."""
    for spec in rules:
        cond, fnt = spec[0], spec[1]
        nf = spec[2] if len(spec) > 2 else None
        fl = spec[3] if len(spec) > 3 else None
        if year_cond:
            _rule(ws, ref, f"AND({cond},{year_cond})", fnt, fill(year_fill), nf)
        _rule(ws, ref, cond, fnt, fl, nf)
    if year_cond:
        _rule(ws, ref, year_cond, base, fill(year_fill))
    _rule(ws, ref, "TRUE", base)


def _label(ws, row, text, col=2):
    c = ws.cell(row, col)
    C.set_text(c, text)
    return c


def _clear_row(ws, row, c1, c2):
    for c in C.iter_cells(ws, c1, row, c2, row):
        c.fill = NOFILL
        c.border = Border()


def _link(ws, c1, row, c2, text, target):
    """Sprunglink innerhalb des Blatts: verbundener Textlink, 9 pt fett Blau (Zell-Hyperlink)."""
    for c in C.iter_cells(ws, c1, row, c2, row):
        c.fill = NOFILL
        c.border = Border()
    C.safe_merge(ws, c1, row, c2, row)
    c = ws.cell(row, _col(c1))
    C.set_text(c, text)
    c.font = font(T_SMALL, True, BLUE)
    c.alignment = align("left", "center", 1)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=C.link_loc(ws.title, target), display=text,
                            tooltip=f"Springen zu {text.lstrip('↓↑ ').strip()}")


def _calibri(ws):
    """Sicherheitsnetz: Vorlagen-Schriften (Inter/Fraunces) in Band-, Lücken- und Leerzellen auf Calibri."""
    for c in list(ws._cells.values()):
        f = c.font
        if f is not None and f.name not in (C.SANS, None):
            nf = copy(f)
            nf.name = C.SANS
            c.font = nf


def _up_link(cell, color=BLUE):
    """„↑ Übersicht“ (P43): immer in Spalte C des Abschnittsbands (fixierter Bereich, bleibt beim Scrollen sichtbar),
    8,5 pt wie die Band-Meta, rechtsbündig, Sprung an den Blattanfang."""
    C.text_link(cell, "↑ Übersicht", cell.parent.title, size=C.T_LABEL, bold=False,
                tooltip="Zurück an den Seitenanfang")
    cell.font = font(C.T_LABEL, False, color)
    cell.alignment = align("right", "center", 1)


# ------------------------------------------------------------------------------------------------ Seitenkopf
def header(ws, eyebrow, title, subtitle, legend=None, link=None, edge="M"):
    """Seitenkopf-Vorlage P04 (C.page_header): Z. 5 Eyebrow „BEREICH  ›  SEITE“ (rechts die Unterreiter-Formen),
    Z. 6 H1 22 pt und rechts bündig der Rückfall-Link „Eingaben dazu: Sxx … ›“ (P37/P05), Z. 7 Untertitel und rechts
    bündig die Legende. Rechte Kante = `edge`: auf den 40-Jahres-Blättern Spalte M (Jahr-10-Rasterlinie = rechte Kante
    der Parameter-/Exit-Blöcke ≈ 1 375 px ≈ Reiterleiste), auf dem AfA-Vergleich Spalte Q (Inhaltskante)."""
    last = max(LAST, _col(edge))
    for r in (5, 6, 7):
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row == r and mr.min_col >= 4:
                ws.unmerge_cells(str(mr))
        for c in C.iter_cells(ws, 3, r, last, r):
            c.border = Border()
            c.fill = NOFILL
            if c.column >= 4 and not C.is_formula(c.value):
                c.value = None
    C.page_header(ws, "B", "C", eyebrow, title, subtitle)
    ws.row_dimensions[4].height = 12
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 22
    ws["B7"].alignment = align("left", "center")
    if legend is not None:
        C.safe_merge(ws, "D", 7, edge, 7)
        ws.cell(7, 4).value = legend
        ws.cell(7, 4).font = font(C.T_LABEL, False, MUTED)
        ws.cell(7, 4).alignment = align("right", "center")
    if link:
        text, target = link
        C.safe_merge(ws, "D", 6, edge, 6)
        c = ws.cell(6, 4)
        C.text_link(c, text, target, size=T_BODY, bold=True, tooltip=f"Eingaben zu diesem Blatt: {target}")
        c.alignment = align("right", "center")


def legend(*items, size=T_SMALL):
    """Legende als Rich-Text: items = (text, marke|None, markenfarbe) – Marke ● / ■ in Farbe, Text 5B6068."""
    parts = []
    for i, (text, mark, color) in enumerate(items):
        if i:
            parts.append(("   ·   ", size, False, MUTED))
        if mark:
            parts.append((mark + " ", size, True, color))
        parts.append((text, size, False, MUTED))
    return C.rich(parts)


def year_legend(*extra):
    """Kopflegende der Jahrestabellen (8,5 pt, rechtsbündig in Z. 7)."""
    return legend(("Verkaufsjahr", "■", ACCENT), ("negatives Ergebnis", "●", RED), *extra, size=C.T_LABEL)


# ------------------------------------------------------------------------------------------------ Jahrestabelle
def year_head(ws, r_year, r_dup, r_cal):
    """Zweizeiliger Kopf: r_year „Jahr n“ (Navy, 9 pt fett weiß), r_dup ausgeblendet, r_cal Kalenderjahr (Blau)."""
    for c in C.iter_cells(ws, 2, r_year, LAST, r_year):
        c.fill = fill(NAVY)
        c.border = Border()
        if c.column >= FIRST:
            c.number_format = YEAR_FMT
            c.font = font(T_SMALL, True, WHITE)
            c.alignment = align("right", "center", 1)
    C.set_text(ws.cell(r_year, 2), "Jahr der Kalkulation")
    ws.cell(r_year, 2).font = font(T_SMALL, True, WHITE)
    ws.cell(r_year, 2).alignment = align("left", "center", 1)
    ws.cell(r_year, 3).value = None
    ws.row_dimensions[r_year].height = H_YEAR
    C.hide_rows(ws, r_dup, r_dup)
    for c in C.iter_cells(ws, 2, r_cal, LAST, r_cal):
        c.fill = fill(BLUE)
        c.border = Border(bottom=side("medium", ACCENT))
        if c.column >= FIRST:
            c.number_format = NUMFMT["year"]
            c.font = font(T_SMALL, False, WHITE)
            c.alignment = align("right", "center", 1)
    C.set_text(ws.cell(r_cal, 2), "Kalenderjahr")
    ws.cell(r_cal, 2).font = font(T_SMALL, False, WHITE)
    ws.cell(r_cal, 2).alignment = align("left", "center", 1)
    ws.row_dimensions[r_cal].height = H_YEAR


def band(ws, row, title, c2=LAST, years_from=None, meta=None, extra=None):
    """Abschnittskopf Ebene 1 (C.section) über die Tabellenbreite; optional Kalenderjahre (=D$42) in D:AQ
    als Ersatzkopf (9 pt fett Navy, P3-04). Formel-Titel (Darlehen II) bleiben Formeln."""
    formula = isinstance(title, str) and title.startswith("=")
    C.section(ws, row, "B", c2, None if formula else title, level=1, meta=meta, extra=extra)
    if formula:
        ws.cell(row, 2).value = title
    if years_from:
        for cc in range(FIRST, LAST + 1):
            c = ws.cell(row, cc)
            c.value = f"={get_column_letter(cc)}${years_from}"
            c.number_format = NUMFMT["year"]
            c.font = font(T_MICRO, False, MUTED)     # P43: dezente Scroll-Orientierung, kein dritter Kopf
            c.alignment = align("right", "center", 1)


STAGE = {"sub": "sub", "result": "result", "deduct": "deduct"}


def data_row(ws, row, kind="data", fmt=NUM, label=None, unit=True):
    """kind: data · davon · memo · deduct · sub · result (Summenstufen über C.sum_row, P1-14)."""
    if label is not None:
        _label(ws, row, label)
    b, cu = ws.cell(row, 2), ws.cell(row, 3)
    for c in C.iter_cells(ws, 2, row, LAST, row):
        c.fill = NOFILL
        c.border = Border(bottom=side("hair", LINE))
        if c.column >= FIRST:
            c.font = font(T_BODY, False, INK)
            c.alignment = align("right", "center", 1)
            c.number_format = fmt
    b.font = font(T_BODY, False, INK)
    b.alignment = align("left", "center", 1)
    if kind == "davon":
        b.font = font(T_SMALL, False, MUTED)
        b.alignment = align("left", "center", 2)
    elif kind == "memo":
        C.memo(ws, row, 2, LAST)
        b.alignment = align("left", "center", 1)
    elif kind in STAGE:
        # Negativ-Rot setzt negatives() zusammen mit der Verkaufsjahr-Markierung (LibreOffice wertet nur die erste Regel)
        C.sum_row(ws, row, "B", get_column_letter(LAST), STAGE[kind], neg=False)
    if unit:
        cu.font = font(T_SMALL, False, MUTED, italic=(kind == "memo"))
        cu.alignment = align("center", "center")
    ws.row_dimensions[row].height = C.H_ROW if kind == "result" else H_DATA


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
        for c in C.iter_cells(ws, 3, r, line_to, r):
            c.value = None
    C.footer(ws, row, 2, line_to)
    C.safe_merge(ws, "B", row, merge_to, row)
    C.safe_merge(ws, "B", row + 1, merge_to, row + 1)
    ws.cell(row, 2).alignment = align("left", "center")
    ws.cell(row + 1, 2).alignment = align("left", "center")


def year_marks(ws, head_ref, cond, band_refs=()):
    """Verkaufsjahr (P3-05): Kopf 4A86C8 mit Schrift 0B2A4A fett, Bänder ebenso, Tabellenspalte E7EEF7."""
    _rule(ws, head_ref, cond, Font(color=NAVY, bold=True), fill(ACCENT))
    for ref in band_refs:
        _rule(ws, ref, cond, Font(color=NAVY, bold=True), fill(ACCENT))
        _rule(ws, ref, "TRUE", Font(color=MUTED, bold=False))
    return cond


def negatives(ws, rows, year_cond, finals=()):
    """P2-01/P15: Rot B42318 für negative Werte in Ergebnis-/Kumulzeilen; rows: {zeile: grundfarbe}.
    Endsummen (finals, Fläche E7EEF7 = Markierungsfarbe): Verkaufsjahr in C8D7EB und fett, damit die Spalte durchläuft."""
    for r, base in rows.items():
        fin = r in finals
        cf_rules(ws, f"D{r}:AQ{r}", [(f"D{r}<0", Font(color=RED, bold=True if fin else None))],
                 Font(color=base), year_cond, year_fill=C.MIST if fin else TINT)


def render(ws, rows, fmts):
    for r, spec in rows.items():
        kind = spec[0]
        if kind == "gap":
            gap(ws, r)
        elif kind == "band":
            band(ws, r, spec[1], years_from=spec[2] if len(spec) > 2 else None,
                 extra=spec[3] if len(spec) > 3 else None)
        else:
            data_row(ws, r, kind, fmts.get(r, NUM), spec[1])


# ------------------------------------------------------------------------------------------------ Projektion
def projektion(ws):
    _ungroup_cols(ws)
    _width(ws, "C", 12)
    header(ws, ("Berechnung", "Projektion"), "Projektion",
           "Miete, Kosten, Cashflow und Vermögen über 40 Jahre  ·  Beträge in € pro Jahr",
           legend=year_legend(("Steuer: + Zahlung / – Erstattung", None, None),
                              ("Jahr 1 = erste 12 Monate ab Kauf", None, None)),
           link=("Eingaben dazu: S11 Prognose & Exit  ›", "S11 Prognose & Exit"))
    year_head(ws, 8, 9, 10)
    # P29: Spalte C trägt im Kopf „Jahr 0 / Kauf“ – Kopf der Exit-Spalte C49 (Eigenkapitaleinsatz)
    for r, text, bg in ((8, "Jahr 0", NAVY), (10, "Kauf", BLUE)):
        c = ws.cell(r, 3)
        C.set_text(c, text)
        c.font = font(T_SMALL, r == 8, WHITE)
        c.alignment = align("right", "center", 1)
    gap(ws, 11, 9)
    rows = {
        12: ("band", "Mieteinnahmen"),
        13: ("data", None), 14: ("data", None), 15: ("sub", None),
        16: ("memo", "Umlagefähige Betriebskosten (Durchlaufposten)"), 17: ("memo", "Warmmiete"),
        18: ("gap",),
        19: ("band", "Bewirtschaftungskosten", None, "· nicht umlagefähig, Zahlungsabfluss"),
        20: ("data", None), 21: ("data", "Zuführung Erhaltungsrücklage WEG"), 22: ("data", None),
        23: ("data", None), 24: ("data", None), 25: ("data", None), 26: ("sub", None),
        27: ("gap",),
        28: ("band", "Cashflow"),
        29: ("sub", "= Einnahmenüberschuss vor Finanzierung (NOI)"),
        30: ("davon", "Zinsen"), 31: ("davon", "Tilgung (inkl. Sondertilgung)"),
        32: ("sub", "– Kapitaldienst (Zinsen + Tilgung inkl. Sondertilgung)"), 33: ("sub", None),
        34: ("memo", "Info: Steuerliches Ergebnis"), 35: ("data", None), 36: ("result", None),
        37: ("memo", None), 38: ("memo", None),
        39: ("gap",),
        40: ("band", "Vermögensentwicklung"),
        41: ("data", None), 42: ("data", "– Restschuld Jahresende"), 43: ("result", "= Nettovermögen"),
        44: ("memo", "Vermögenszuwachs im Jahr"), 45: ("data", None), 46: ("memo", None),
        47: ("gap",),
        48: ("band", "Exit-Cashflow", None, "· Basis für die Eigenkapitalrendite (IRR)"),
        49: ("result", "= Exit-Cashflow (Jahr 0 = Eigenkapitaleinsatz)"),
        50: ("gap",),
    }
    render(ws, rows, {45: PCT1})
    # P12: Eigenkapitalrendite als hervorgehobene Kennzahl zwischen den Memo-Zeilen – 10 pt fett 1D4F8A, ohne Fläche
    for c in C.iter_cells(ws, 2, 45, LAST, 45):
        c.font = font(T_BODY, True, BLUE)
    ws.cell(45, 3).font = font(T_SMALL, False, MUTED)
    # Z. 48/49 (P29): C48 frei für den Rücksprung, C49 wie die Jahreszellen (rechtsbündig, Einzug 1)
    ws["C49"].number_format = NUM
    ws["C49"].font = font(T_BODY, True, NAVY)
    ws["C49"].alignment = align("right", "center", 1)
    for r in (12, 19, 28, 40, 48):
        _up_link(ws.cell(r, 3))
    footer(ws, 51)
    grid(ws, [8, 10] + [r for r, s in rows.items() if s[0] != "gap"])
    ws.freeze_panes = "D11"
    # Bedingte Formatierung (nur Anzeige)
    ycond = year_marks(ws, "D8:AQ10", "D$9=Haltedauer")
    negatives(ws, {15: NAVY, 29: NAVY, 33: NAVY, 34: MUTED, 36: NAVY, 37: MUTED, 38: MUTED, 43: NAVY, 44: MUTED,
                   49: NAVY}, ycond, finals=(36, 43, 49))
    C.neg_red(ws, "C49")
    _rule(ws, "D11:AQ49", ycond, None, fill(TINT))


# ------------------------------------------------------------------------------------------------ Finanzierung
KD_LABEL = "Kapitaldienst (Zinsen + Tilgung inkl. Sondertilgung)"   # mappenweit gleiche Beschriftung (P12)


def finanzierung(ws):
    _ungroup_cols(ws)
    _width(ws, "C", 12)
    header(ws, ("Berechnung", "Tilgungsplan"), "Tilgungsplan",
           "Tilgungsverlauf Darlehen I und II  ·  Annuitätendarlehen mit konstanter Rate",
           legend=legend(("Verkaufsjahr", "■", ACCENT), ("Anschlusszins nach der Zinsbindung", "●", BLUE),
                         ("– entfällt (getilgt bzw. nicht genutzt)", None, None), size=C.T_LABEL),
           link=("Eingaben dazu: S07 Finanzierung  ›", "S07 Finanzierung"))
    year_head(ws, 8, 9, 10)
    gap(ws, 11, 9)
    rows = {
        12: ("band", "Darlehen I"),
        13: ("data", None), 14: ("data", None), 15: ("data", None), 16: ("data", None),
        17: ("data", "– Reguläre Tilgung"), 18: ("data", "– Sondertilgung"),
        19: ("sub", "= Restschuld Jahresende"), 20: ("sub", "= " + KD_LABEL),
        21: ("gap",),
        22: ("band", '="Darlehen II"&IF(Darlehen_II=0,"  ·  nicht genutzt","")'),
        23: ("data", None), 24: ("data", None), 25: ("data", None), 26: ("data", None),
        27: ("data", "– Reguläre Tilgung"), 28: ("data", "– Sondertilgung"),
        29: ("data", "– Tilgungszuschuss (KfW, kein Zahlungsabfluss)"),
        30: ("sub", "= Restschuld Jahresende"), 31: ("sub", "= " + KD_LABEL),
        32: ("gap",),
        33: ("band", "Summe Darlehen"),
        34: ("data", None), 35: ("data", None), 36: ("data", None),
        37: ("sub", "= Kapitaldienst gesamt (Zinsen + Tilgung inkl. Sondertilgung)"),
        38: ("result", "= Restschuld Jahresende"),
        39: ("memo", None), 40: ("memo", None),
        41: ("gap",),
    }
    render(ws, rows, {14: PCT2, 24: PCT2, 39: PCT2})
    # zwei aufeinanderfolgende Summenzeilen bilden einen Block: die zweite ohne eigene kräftige Oberlinie (P12)
    for r in (20, 31):
        for c in C.iter_cells(ws, 2, r, LAST, r):
            c.border = Border(top=side("hair", LINE), bottom=side("hair", LINE))
    for c in C.iter_cells(ws, 2, 38, LAST, 38):
        c.border = Border(top=side("hair", LINE), bottom=c.border.bottom)
    for r in (12, 22, 33):
        _up_link(ws.cell(r, 3))
    # P43: Darlehen II ungenutzt → Zeilen 23–31 als Gliederungsgruppe, beim Build eingeklappt; Hinweis im Band
    d2 = _input_value(ws.parent, "Darlehen_II")
    for r in range(23, 32):
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].hidden = d2 == 0
    ws.row_dimensions[32].collapsed = d2 == 0
    hint = ws.cell(22, KV_LAST)
    hint.value = ('=IF(Darlehen_II=0,"▸ ausgeblendet  ·  über das „+“ der Gliederung links einblenden",'
                  '"▸ über die Gliederung links ein-/ausklappbar")')
    hint.font = font(C.T_LABEL, False, BLUE)
    hint.alignment = align("right", "center", 1)
    ws.sheet_format.outlineLevelRow = max(ws.sheet_format.outlineLevelRow or 0, 1)
    footer(ws, 42)
    grid(ws, [8, 10] + [r for r, s in rows.items() if s[0] != "gap"])
    ws.freeze_panes = "D11"
    ycond = year_marks(ws, "D8:AQ10", "D$9=Haltedauer")
    faded = Font(color=FADED, bold=False, italic=True)
    switch = Font(bold=True, color=BLUE)          # Zinswechsel nach der Zinsbindung
    # P2-02/P43: entfallene Werte „–“ in 98A2B3 kursiv, nicht fett – Darlehen I nach Volltilgung (Z. 13–20),
    # Summenblock nach Volltilgung (Z. 34–40), Darlehen II ungenutzt (Z. 23–31)
    for r in range(13, 21):
        rules = [("D$13=0", faded, DASH_FMT)] + ([("AND(ISNUMBER(C14),D14<>C14)", switch)] if r == 14 else [])
        cf_rules(ws, f"D{r}:AQ{r}", rules, Font(color=NAVY if r in (19, 20) else INK), ycond)
    for r in range(23, 32):
        rules = [("Darlehen_II=0", faded, DASH_FMT, fill(WHITE) if r in (30, 31) else None)]
        if r == 24:
            rules.append(("AND(ISNUMBER(C24),D24<>C24)", switch))
        cf_rules(ws, f"D{r}:AQ{r}", rules, Font(color=NAVY if r in (30, 31) else INK), ycond)
    for r in range(34, 41):
        cf_rules(ws, f"D{r}:AQ{r}", [("D$34=0", faded, DASH_FMT)],
                 Font(color=NAVY if r in (37, 38) else (MUTED if r > 38 else INK)), ycond,
                 year_fill=C.MIST if r == 38 else TINT)
    # Beschriftung B:C der inaktiven Zeilen ebenso zurücknehmen (Summenzeilen dann weiß, nicht fett)
    _rule(ws, "B23:C29", "Darlehen_II=0", Font(color=FADED, bold=False, italic=True))
    _rule(ws, "B30:C31", "Darlehen_II=0", Font(color=FADED, bold=False, italic=True), fill(WHITE))
    _rule(ws, "D11:AQ40", ycond, None, fill(TINT))


def _input_value(wb, name):
    """Wert einer Eingabe hinter einem Namen beim Build (nur Konstanten; Formeln → None)."""
    try:
        dn = wb.defined_names.get(name)
        sheet, ref = next(iter(dn.destinations))
        v = wb[sheet][ref.replace("$", "")].value
        return None if C.is_formula(v) else v
    except Exception:
        return None


# ------------------------------------------------------------------------------------------------ Steuern
STEUERN_LABELS = {
    17: "Wohngebäude", 21: "Sofortige Verlustverrechnung", 23: "Disagio sofort abziehbar",
    27: "Typisierte Nutzungsdauer",
    28: "Degressive AfA aktiv", 30: "AfA-Bemessungsgrundlage Gebäude (ohne Sanierungsanteil)",
    32: "Sonder-AfA § 7b aktiv", 35: "Renovierung als Herstellungskosten aktivieren",
    78: "Haltedauer",
    85: "Veräußerungsgewinn (Preis – Verkaufskosten – Buchwert)",
    86: "Veräußerungsgewinn steuerpflichtig",
    91: "= Nettoerlös nach Steuern und Darlehensablösung",
    94: "= Gesamtertrag nach Steuern",
    96: "IRR n. St.  ·  Eigenkapitalrendite p. a.",
    97: "AfA-Vergleich:  Summe Gebäude-AfA Jahre 1–10 (gewählte Variante)",
    98: "AfA-Vergleich:  Summe Gebäude-AfA Jahre 1–10 (Standard nach Baujahr)",
    99: "AfA-Vergleich:  Steuereffekt der gewählten Variante (10 Jahre)",
}
STEUERN_NOTES = {
    24: "Fallback auf „Automatisch“, wenn Voraussetzungen fehlen (siehe Prüfhinweise im Cockpit)",
    88: "nur bei Option „Verlustvortrag“ bzw. GmbH; Mindestbesteuerung (§ 10d Abs. 2 EStG) nicht modelliert",
    94: "Cashflows + Nettoerlös – Eigenkapital",
    99: "gegenüber Standard-AfA · positiv = Steuerersparnis (Liquiditätsvorteil, keine endgültige Ersparnis)",
}
STEUERN_FMT = {
    17: YESNO, 18: PCT2K, 19: PCT2K, 20: PCT2K, 21: YESNO, 22: NUMFMT["years_n"], 23: YESNO,
    25: PCT2K, 26: PCT2K, 27: NUMFMT["years_n"], 28: YESNO, 31: C.typo_minus('#,##0" €/m²";-#,##0" €/m²";"–"'),
    32: YESNO, 35: YESNO, 78: NUMFMT["years_n"], 86: YESNO, 87: PCT2K, 95: NUMFMT["mult2"], 96: PCT1,
}
KV_LAST = 13   # M


def kv_rows(ws, rows, last=KV_LAST):
    """Parameter-/Exit-Block: Beschriftung B, Wert C rechts, Hinweis E:M grau; Haarlinie, keine Seitenrahmen."""
    for r in rows:
        if r in STEUERN_LABELS:
            _label(ws, r, STEUERN_LABELS[r])
        if r in STEUERN_NOTES:
            C.set_text(ws.cell(r, 5), STEUERN_NOTES[r])
        for c in C.iter_cells(ws, 2, r, last, r):
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


def kv_sum(ws, row, stage, top=True):
    """Summenstufe im Parameter-/Exit-Block (B:M); Hinweis in E bleibt 9 pt grau."""
    C.sum_row(ws, row, "B", get_column_letter(KV_LAST), stage)
    n = ws.cell(row, 5)
    n.font = font(T_SMALL, False, MUTED)
    if not top:
        for c in C.iter_cells(ws, 2, row, KV_LAST, row):
            b = c.border
            c.border = Border(top=side("hair", LINE), bottom=b.bottom)
    ws.row_dimensions[row].height = C.H_ROW


def _outline(ws, r1, r2, level=1):
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].outlineLevel = max(ws.row_dimensions[r].outlineLevel or 0, level)


def steuern(ws):
    _ungroup_cols(ws)
    _width(ws, "C", 12)
    header(ws, ("Steuern", "Steuer-Tabelle"), "Steuer-Tabelle",
           "Abgeleitete Parameter, AfA-Verlauf und steuerliches Ergebnis über 40 Jahre",
           legend=year_legend(("Steuer: + Zahlung / – Erstattung (Verrechnung mit anderen Einkünften)", None, None)),
           link=("Eingaben dazu: S09 Steuern  ›", "S09 Steuern"))
    # ---- Parameterblock (P3-04: per Gliederung einklappbar; technische Indizes Ebene 2, bleiben verborgen)
    band(ws, 8, "Abgeleitete steuerliche Parameter", c2=KV_LAST, extra="· über die Gliederung links einklappbar")
    # Sprunglinks im Band (statt schwebender Textlinks neben dem Titel): rechtsbündig wie die Band-Meta
    for c1, c2, text, row in (("I", "K", "↓ Jahrestabelle", 40), ("L", "M", "↓ Verkaufsszenario", 77)):
        C.safe_merge(ws, c1, 8, c2, 8)
        c = ws[f"{c1}8"]
        C.text_link(c, text, ws.title, C.link_row(ws.title, row), size=C.T_LABEL, bold=False,
                    tooltip=f"Springen zu {text[2:]}")
        c.font = font(C.T_LABEL, False, BLUE)
        c.alignment = align("right", "center", 1)
    C.hide_rows(ws, 9, 16)
    C.hide_rows(ws, 24, 24)
    kv_rows(ws, [r for r in range(17, 38) if r != 24])
    gap(ws, 38, H_GAP)
    C.hide_rows(ws, 39, 39)
    _outline(ws, 9, 37, 1)
    _outline(ws, 9, 16, 2)
    _outline(ws, 24, 24, 2)
    # ---- Jahrestabelle
    year_head(ws, 40, 41, 42)
    gap(ws, 43, 9)
    rows = {
        44: ("band", "Abschreibung (AfA)", 42),
        45: ("data", None), 46: ("data", None), 47: ("data", None), 48: ("data", None), 49: ("data", None),
        50: ("data", None), 51: ("sub", None), 52: ("memo", None), 53: ("memo", None), 54: ("data", None),
        55: ("data", None), 56: ("sub", None), 57: ("memo", None),
        58: ("gap",),
        59: ("band", "Steuerliches Ergebnis aus dem Objekt", 42),
        60: ("data", None),
        61: ("davon", "Bewirtschaftungskosten (steuerlich)"), 62: ("davon", "Schuldzinsen"),
        63: ("davon", "Finanzierungsnebenkosten und Disagio"),
        64: ("davon", "Erhaltungsaufwand (ggf. verteilt, § 82b EStDV)"),
        65: ("davon", "Abschreibungen"),
        66: ("sub", "– Werbungskosten / Betriebsausgaben gesamt"), 67: ("sub", "= Steuerliches Ergebnis"),
        68: ("memo", None), 69: ("data", None), 70: ("memo", None),
        71: ("data", "Körperschaftsteuersatz im Kalenderjahr (nur GmbH)"),
        72: ("data", "Angewendeter Steuersatz (Grenzsatz bzw. KSt + Soli + GewSt)"),
        73: ("result", "= Steuer (+ Zahlung / – Erstattung)"), 74: ("memo", None),
    }
    render(ws, rows, {71: PCT2K, 72: PCT2})
    grid(ws, [40, 42] + [r for r, s in rows.items() if s[0] != "gap"])
    for r in (44, 59):
        _up_link(ws.cell(r, 3))
    gap(ws, 75, H_GAP)
    C.hide_rows(ws, 76, 76)
    # ---- Exit-Block
    band(ws, 77, "Verkaufsszenario (Exit)", c2=KV_LAST, extra="· Steuer und Gesamtrendite")
    _up_link(ws.cell(77, 3))
    kv_rows(ws, range(78, 100))
    kv_sum(ws, 91, "sub")
    kv_sum(ws, 94, "sub")
    kv_sum(ws, 95, "sub", top=False)
    kv_sum(ws, 96, "final")
    # P12: AfA-Vergleich (Z. 97–99) als nachrichtlicher Unterabschnitt unter der Endsumme – Memo-Stil, Luft darüber
    for r in (97, 98, 99):
        C.sum_row(ws, r, "B", get_column_letter(KV_LAST), "memo")
        ws.cell(r, 3).alignment = align("right", "center", 1)
        ws.cell(r, 5).font = font(T_SMALL, False, MUTED, italic=True)
    ws.row_dimensions[97].height = C.H_ROW2
    for cc in range(2, KV_LAST + 1):
        ws.cell(97, cc).alignment = align(ws.cell(97, cc).alignment.horizontal or "left", "bottom",
                                          ws.cell(97, cc).alignment.indent)
    _outline(ws, 78, 99, 1)
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True, applyStyles=False,
                                            showOutlineSymbols=True)
    ws.sheet_format.outlineLevelRow = 2
    gap(ws, 100, H_GAP)
    footer(ws, 101, line_to=KV_LAST)
    # ---- Bedingte Formatierung
    ycond = "D$41=Haltedauer"
    year_marks(ws, "D40:AQ42", ycond, ("D44:AQ44", "D59:AQ59"))
    faded = Font(color=FADED, bold=False, italic=True)
    cf_rules(ws, "D71:AQ71", [("Rechtsform_Idx=1", faded, DASH_FMT)], Font(color=INK), ycond)
    _rule(ws, "B71:C71", "Rechtsform_Idx=1", faded)
    negatives(ws, {67: NAVY, 69: INK}, ycond)
    # Z. 73: Steuer-Sicht laut Beschriftung – negative Werte (Erstattung) bewusst NICHT rot (P07)
    cf_rules(ws, "D73:AQ73", [], Font(color=NAVY), ycond, year_fill=C.MIST)
    _rule(ws, "D43:AQ74", ycond, None, fill(TINT))
    # Exit-Block: negative Ergebnisse rot (P2-01, C92 kumulierter Cashflow)
    C.neg_red(ws, "C92")


# ------------------------------------------------------------------------------------------------ AfA-Vergleich
AFA_VARIANTS = {
    19: ("Linear 2,0 %", "§ 7 Abs. 4 S. 1 Nr. 2b EStG · Fertigstellung 1925–2022"),
    20: ("Linear 2,5 %", "§ 7 Abs. 4 S. 1 Nr. 2c EStG · Fertigstellung vor 1925"),
    21: ("Linear 3,0 %", "§ 7 Abs. 4 S. 1 Nr. 2a EStG · Fertigstellung ab 2023"),
    22: ("Gutachten (RND)", "§ 7 Abs. 4 S. 2 EStG · tatsächliche Restnutzungsdauer"),
    23: ("Degressiv 5 %", "§ 7 Abs. 5a EStG · Baubeginn 10/2023–9/2029"),
    24: ("Linear 3 % + § 7b", "5 % p. a. in Jahren 1–4, danach Restwert-AfA"),
    25: ("Degressiv + § 7b", "kombiniert; die Sonder-AfA mindert den Restwert"),
    26: ("Im Modell angewendet", "lt. Steuer-Tabelle · inkl. Sonder-AfA, ohne § 7h/7i"),
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
         "Sanierungsanteil und sind hier nicht als Variante dargestellt; sie werden in der Steuer-Tabelle zusätzlich "
         "zur regulären AfA der Altsubstanz berechnet."),
}
AFA_LAST = 17  # Q
EMU_PT = 12700
# Kurzform der angewendeten Methode für die Kachel (reine Anzeigeformel, P1-11)
AFA_SHORT = ('=IF(Is_Degressiv=1,"Degressiv "&FIXED(Degressiv_Satz*100,1)&" %",'
             'IF(Methode_eff=6,"Gutachten ","Linear ")&FIXED(AfA_Satz_linear*100,1)&" %")'
             '&IF(SonderAfA_aktiv=1," + § 7b","")')


def afa(ws):
    _ungroup_cols(ws)
    # Spaltenraster (P10/P28): B:C Beschriftung (C breit genug für „Automatisch nach Baujahr (§ 7 Abs. 4 EStG)“),
    # D:O einheitlich 74 px (Jahre, Σ-Spalten, Kachelrinnen I und N gleich breit), P:Q je 111 px → alle drei Kacheln 296 px
    _width(ws, "B", 37)
    _width(ws, "C", 34)
    for cc in range(4, 16):
        _width(ws, get_column_letter(cc), 10.57)
    _width(ws, "P", 15.86)
    _width(ws, "Q", 15.86)
    header(ws, ("Steuern", "AfA-Vergleich"), "AfA-Vergleich",
           "Abschreibungsvarianten für dieses Objekt in den Jahren 1–10  ·  Beträge in €",
           link=("Eingaben dazu: S10 Abschreibung  ›", "S10 Abschreibung"), edge="Q")
    # ---- Grundlagen + Kacheln
    band(ws, 8, "Grundlagen", c2=AFA_LAST)
    fmts = {9: EUR, 10: EUR, 11: EUR, 12: PCT2, 13: PCT2, 14: '0" Jahre";;"– (kein Gutachten)"', 15: "@"}
    for r in range(9, 16):
        _label(ws, r, AFA_BASICS[r])
        _clear_row(ws, r, 2, AFA_LAST)
        for c in C.iter_cells(ws, 2, r, 3, r):
            c.border = Border(bottom=side("hair", LINE))
        b, v = ws.cell(r, 2), ws.cell(r, 3)
        b.font = font(T_BODY, False, INK)
        b.alignment = align("left", "center", 1)
        v.font = font(T_BODY, False, INK)
        v.alignment = align("right", "center", 1)
        v.number_format = fmts[r]
        ws.row_dimensions[r].height = C.H_ROW
    ws["C15"].value = '=Baujahr&" · "&Objektart'
    tiles(ws)
    gap(ws, 16, H_GAP)
    # ---- Variantentabelle
    band(ws, 17, "AfA-Beträge je Jahr", c2=AFA_LAST,
         meta="Steuerersparnis = AfA × Steuersatz Jahr 1 · Barwert mit dem Kalkulationszins")
    # Versalien nur für die Textköpfe – „Jahr n“ bleibt in Satzschreibung (Kategorien des Liniendiagramms)
    heads = {"B": ("VARIANTE", "left"), "C": ("ANWENDBAR FÜR DIESES OBJEKT?", "left"),
             "N": ("Σ JAHRE\n1–4", "right"), "O": ("Σ JAHRE\n1–10", "right"),
             "P": ("STEUERERSPARNIS\n1–10 NOMINAL", "right"), "Q": ("BARWERT\nSTEUERERSPARNIS", "right")}
    for k in range(10):
        heads[get_column_letter(4 + k)] = (f"Jahr {k + 1}", "right")
    C.section(ws, 18, "B", get_column_letter(AFA_LAST), level=2, labels=heads, height=C.H_ROW2, caps=False)
    for cc in range(2, AFA_LAST + 1):
        c = ws.cell(18, cc)
        c.alignment = align(c.alignment.horizontal, "center", 1, wrap=cc >= 14)
    for r in range(19, 27):
        name, basis = AFA_VARIANTS[r]
        key = r == 26
        b = ws.cell(r, 2)
        b.value = C.rich([(name, T_BODY, True, NAVY if key else INK), ("\n" + basis, T_MICRO, False, MUTED)])
        for c in C.iter_cells(ws, 2, r, AFA_LAST, r):
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE))
            if c.column >= 4:
                c.font = font(T_BODY, c.column in (15, 17), INK)
                c.alignment = align("right", "center", 1)
                c.number_format = NUM
        b.font = font(T_BODY, True, INK)
        b.alignment = align("left", "center", 1, wrap=True)
        cst = ws.cell(r, 3)
        cst.font = font(T_SMALL, True, MUTED)
        cst.alignment = align("left", "center", 1, wrap=True)
        cst.number_format = NUMFMT["status_dot"]
        ws.row_dimensions[r].height = C.H_ROW2
    C.sum_row(ws, 26, "B", get_column_letter(AFA_LAST), "result")
    ws.row_dimensions[26].height = C.H_ROW2
    ws["C26"].font = font(T_SMALL, True, NAVY)
    ws["C26"].alignment = align("left", "center", 1, wrap=True)
    ws["C26"].number_format = "General"
    # Legende der Spalte „Anwendbar“ als Fußnote direkt unter der Tabelle (P1-18)
    _clear_row(ws, 27, 2, AFA_LAST)
    ws["B27"].value = legend(("anwendbar", "●", GREEN), ("nur mit Gutachten", "●", AMBER),
                             ("nicht anwendbar (Beträge nur zum Vergleich)", "●", FADED), size=T_MICRO)
    ws["B27"].font = font(T_MICRO, False, MUTED)
    ws["B27"].alignment = align("left", "top", 1)
    ws.row_dimensions[27].height = 22
    # ---- Nebenrechnungen
    for head_row, title, first in ((28, "NEBENRECHNUNG: RESTWERT GEBÄUDE AM JAHRESENDE (€)", 29),
                                   (37, "NEBENRECHNUNG: KUMULIERTE AFA JE VARIANTE (€)", 38)):
        heads = {"B": (title, "left")}
        for k in range(10):
            heads[get_column_letter(4 + k)] = (f"Jahr {k + 1}", "right")
        C.section(ws, head_row, "B", get_column_letter(AFA_LAST), level=2, labels=heads, caps=False)
        for i in range(7):
            r = first + i
            if head_row == 28:
                _label(ws, r, AFA_VARIANTS[19 + i][0])
            else:
                ws.cell(r, 2).value = f"=$S{19 + i}"
            for c in C.iter_cells(ws, 2, r, AFA_LAST, r):
                c.fill = NOFILL
                c.border = Border(bottom=side("hair", LINE))
                c.font = font(T_SMALL, False, MUTED)
                if c.column >= 4:
                    c.alignment = align("right", "center", 1)
                    c.number_format = NUM
            ws.cell(r, 2).alignment = align("left", "center", 1)
            ws.row_dimensions[r].height = H_DATA
    gap(ws, 36, H_GAP)
    gap(ws, 45, H_GAP)
    # ---- Hinweise: zwei Spalten (B:C | E:Q), je Absatz ein fettes Stichwort (P3-10)
    band(ws, 46, "Hinweise zur Interpretation", c2=AFA_LAST)
    for r in AFA_HINTS:
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row == r:
                ws.unmerge_cells(str(mr))
        _clear_row(ws, r, 2, AFA_LAST)
        for c in C.iter_cells(ws, 2, r, AFA_LAST, r):
            c.value = None
    # P33: zwei gleich breite Spaltengruppen mit einer Rinnenspalte; Zeilenhöhe aus dem längeren Block (P27-Raster)
    split = min(range(5, AFA_LAST - 3), key=lambda k: abs(C.span_px(ws, 2, k) - C.span_px(ws, k + 2, AFA_LAST)))
    layout = ((47, 2, split, 47), (47, split + 2, AFA_LAST, 49), (48, 2, split, 48), (48, split + 2, AFA_LAST, 50))
    need = {}
    for row, c1, c2, key in layout:
        kw, text = AFA_HINTS[key]
        C.safe_merge(ws, c1, row, c2, row)
        c = ws.cell(row, c1)
        c.value = C.rich([(kw, T_SMALL, True, NAVY), (text, T_SMALL, False, INK2)])
        c.font = font(T_SMALL, False, INK2)
        c.alignment = align("left", "top", 1, wrap=True)
        n = C.lines_needed_metric(kw + text, C.span_px(ws, c1, c2), T_SMALL, False, 1)
        need[row] = max(need.get(row, 1), n)
    for row, n in need.items():
        ws.row_dimensions[row].height = C.text_row_height(n)
    gap(ws, 49, H_GAP)
    C.hide_rows(ws, 50, 51)
    # ---- Diagramme
    band(ws, 52, "Diagramme", c2=AFA_LAST, extra="· Jahre 1–10")
    for r in range(53, 70):
        ws.row_dimensions[r].height = 15
    place_charts(ws)
    gap(ws, 70, C.H_ROW)
    # P19: leere Variante als Fußnote unter den Diagrammen statt unkommentierter Lücke (reine Anzeigeformel)
    fn = ws["B70"]
    fn.value = ('=IF(N(C14)=0,"Gutachten (RND): kein Restnutzungsdauer-Gutachten erfasst – Variante ohne Balken '
                'und Linie","")')
    fn.font = font(T_MICRO, False, MUTED, italic=True)
    fn.alignment = align("left", "center", 1)
    C.hide_rows(ws, 71, 72)
    footer(ws, 73, merge_to="Q", line_to=AFA_LAST)
    C.hide_cols(ws, "R", "S")
    # ---- Bedingte Formatierung: Anwendbarkeit als Statuswort (● in Statusfarbe, keine Flächen)
    cf_rules(ws, "C19:C25", [('LEFT($C19,2)="Ja"', Font(color=GREEN, bold=True)),
                             ('LEFT($C19,4)="Nur "', Font(color=AMBER, bold=True)),
                             ('LEFT($C19,4)="Nein"', Font(color=FADED, bold=False))], Font(color=MUTED))
    cf_rules(ws, "D19:Q25", [('LEFT($C19,4)="Nein"', Font(color=FADED, bold=False))], Font(color=INK))


def tiles(ws):
    """Kacheln rechts neben den Grundlagen (C.tile, light – Private Banking, P10/P11): drei gleich breite Kacheln
    (je 296 px, Rinnenspalten I und N gleich breit), gestreckt über Z. 9–15, damit sie mit der Grundlagenliste
    abschließen. Kopfstreifen Z. 9 · Wert Z. 10–14 · Fußzeile Z. 15 (Kontext)."""
    specs = (("E", "H", "Im Modell angewendet", AFA_SHORT, "General", "=C26"),
             ("J", "M", "AfA Jahre 1–10 (angewendet)", "=O26", EUR, "Summe der Gebäude-AfA lt. Steuer-Tabelle"),
             ("O", "Q", "Barwert der Steuerersparnis", "=Q26", EUR, "Jahre 1–10, abgezinst mit dem Kalkulationszins"))
    for r in range(9, 16):
        ws.row_dimensions[r].height = C.H_ROW
    for c1, c2, lab, val, fmt, sub in specs:
        C.tile(ws, c1, c2, 9, 10, 15, label=lab, value=val, sub=sub, fmt=fmt, variant="light", value_rows=5,
               gap_right=False, gap=False, status=None, neg=False)
    for r in range(10, 15):
        ws.row_dimensions[r].height = C.H_ROW


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
            _calibri(wb[name])
            C.cf_close(wb[name])
