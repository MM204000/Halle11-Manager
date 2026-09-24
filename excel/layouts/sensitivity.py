"""Blatt „Sensitivität“: lesbare Matrizen, Regel-Ampeln, Erklärungen neben ihren Objekten, Nebenrechnung eingeklappt.

Stil: ruhiges Rechenblatt (feine Linien, dezente Akzente). Nur Darstellung – Werte in C24:C51 bleiben Zahlen
(die Formeln nutzen $C24 …), sie erhalten lediglich ein sprechendes Zahlenformat.
"""
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.properties import Outline

import core as C

SHEET = "Sensitivität"
FIRST, LAST = 65, 115                      # Nebenrechnung (eingeklappt)
ROW_H = 16                                 # Datenzeilen des Blatts
MATRIX_ROWFMT = '"Miete +"0 %;"Miete -"0 %;"Miete Basis"'
MATRIX_COLFMT = '+0.0 %;-0.0 %;"Basis"'
HEAD_PCT = "0.0 %;-0.0 %;0.0 %"            # Kopfwerte: 0 % ist ein echter Wert, kein „–“

# Status-Töne der Matrizen (kräftiger als die Ampel-Hintergründe der Kacheln, damit die Heatmap lesbar ist)
M_RED, M_AMB, M_GRN = "F6D3CD", "FDEBD3", "D5ECDD"

READ_BOXES = [
    # (Kopfzeile, Titel, Text)
    (9, "Grenzen",
     "Die Jahr-1-Größen sind lineare Näherungen bei konstantem Steuersatz. Nicht abgebildet: Wechselwirkungen mit "
     "dem Mietausfall, Progressionseffekte, Vorfälligkeitsentschädigung und die Ausschüttungsbelastung der GmbH."),
    (21, "Miete × Sollzins",
     "Jede Zelle zeigt den Cashflow nach Steuern pro Monat im ersten Jahr, wenn die Miete um den Zeilenwert und "
     "der Sollzins des Hauptdarlehens um den Spaltenwert abweicht. Grün = Überschuss, rot = Zuschuss."),
    (32, "DSCR",
     "Kapitaldienstdeckung = Einnahmenüberschuss (NOI) geteilt durch die Rate. Banken erwarten meist 1,1 bis 1,2; "
     "unter 1,0 trägt die Bonität des Investors den Kredit."),
    (44, "IRR-Matrix",
     "Für jede Kombination wird die vollständige Zahlungsreihe über die Haltedauer neu gerechnet (ausgeblendete "
     "Nebenrechnung, Zeilen 65–115): Die Mietänderung wirkt in jedem Jahr mit der Mietsteigerung fort, die "
     "Wertsteigerung bestimmt Verkaufspreis, Verkaufskosten und Steuer beim Verkauf."),
]
LESEHILFE = ("Die Break-even-Miete ist die Nettokaltmiete, bei der der Cashflow im ersten Jahr genau null ist – "
             "alles darüber ist Puffer. Der maximale Sollzins zeigt, wie viel Zinsanstieg die Miete bei gleicher "
             "Tilgung verkraftet.\n¹ Bei voller Steuerwirkung: Ein steuerlicher Verlust wird sofort mit anderen "
             "Einkünften verrechnet (Privatperson). In der GmbH oder mit Verlustvortrag wirkt die Steuer erst, "
             "wenn das Objekt Gewinne erzielt.")
LABELS = {
    "C10": "Break-even-Miete für Cashflow vor Steuern = 0",
    "C11": "Puffer zur aktuellen Miete",
    "C12": "Break-even-Miete für Cashflow nach Steuern = 0 ¹",
    "C13": "Puffer zur aktuellen Miete",
    "C14": "Max. gewichteter Sollzins für Cashflow vor Steuern ≥ 0",
    "C15": "Aktueller gewichteter Sollzins",
    "C18": "Zinsänderungsrisiko: +1 %-Punkt Anschlusszins (p. a.)",
    "C20": "Jahr 1: Cashflow und Kapitaldienstdeckung – Miete × Sollzins",
    "C21": "CASHFLOW NACH STEUERN (€ / MONAT)",
    "C32": "KAPITALDIENSTDECKUNG (DSCR = NOI / KAPITALDIENST)",
    "C44": "IRR NACH STEUERN P. A. – VOLLSTÄNDIGE NEUBERECHNUNG DER ZAHLUNGSREIHE",
    "C53": "VERKAUF NACH DER HALTEDAUER JE WERTSTEIGERUNG (€, BASIS-MIETE)",
}


# =============================================================================== Helfer
def _unmerge(ws, r1, r2, c1, c2):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= r2 and mr.max_row >= r1 and mr.min_col <= C.col(c2) and mr.max_col >= C.col(c1):
            ws.unmerge_cells(str(mr))


def _clear(ws, c1, r1, c2, r2, values=True):
    for c in C.iter_cells(ws, c1, r1, c2, r2):
        if values and not C.is_formula(c.value):
            c.value = None
        c.fill = C.NOFILL
        c.border = Border()


def _drop_cf(ws, *refs):
    cf = ws.conditional_formatting
    for key in list(cf._cf_rules):
        if str(key.sqref) in refs:
            del cf._cf_rules[key]


def _rich(*parts):
    """Legende: (Text, Farbe, fett) – farbige ■ in Statusfarben, Rest 8 pt grau."""
    return CellRichText([TextBlock(InlineFont(rFont=C.SANS, sz=C.T_MICRO, b=b, color=col), t) for t, col, b in parts])


def _legend(ws, coord, items):
    parts = []
    for i, (sq, col, text) in enumerate(items):
        parts.append((("   " if i else "") + sq + " ", col, True))
        parts.append((text, C.MUTED, False))
    cell = ws[coord]
    cell.value = _rich(*parts)
    cell.font = C.font(C.T_MICRO, False, C.MUTED)
    cell.alignment = C.align("right", "center")


def _caption(ws, coord, text):
    cell = ws[coord]
    C.set_text(cell, text)
    cell.font = C.font(C.T_MICRO, False, C.MUTED)
    cell.alignment = C.align("left", "center", 1)


def _status_rules(ws, ref, first, green, yellow):
    """Regel-Ampel mit Füllung: < gelb rot · < grün amber · sonst grün (nur Zahlen)."""
    for cond, fg, bg in ((f"{first}<{yellow}", C.RED, M_RED), (f"{first}<{green}", C.AMBER, M_AMB),
                         ("TRUE", C.GREEN, M_GRN)):
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f"AND(ISNUMBER({first}),{cond})"], stopIfTrue=True,
                                                       font=Font(color=fg), fill=C.fill(bg)))


def _matrix(ws, head, r1, r2, c_last, corner):
    """Matrix: Kopfzeile (Spaltenwerte), Zeilenköpfe „Miete ±x %“ mit Achsentrenner, Datenzellen mit Haarlinie."""
    C.table_head(ws, head, "C", c_last)
    corner_cell = ws[f"C{head}"]
    C.set_text(corner_cell, corner)
    corner_cell.font = C.font(C.T_MICRO, True, C.BLUE)
    corner_cell.alignment = C.align("right", "center", 1)
    corner_cell.border = Border(bottom=C.side("thin", C.ACCENT), right=C.side("thin", C.ACCENT))
    for c in C.iter_cells(ws, "D", head, c_last, head):
        c.alignment = C.align("right", "center", 1)
    C.set_height(ws, head, C.H_HEAD)
    for r in range(r1, r2 + 1):
        lab = ws[f"C{r}"]
        lab.number_format = MATRIX_ROWFMT
        lab.font = C.font(C.T_SMALL, True, C.BLUE)
        lab.fill = C.NOFILL
        lab.alignment = C.align("left", "center", 1)
        lab.border = Border(bottom=C.side("hair", C.LINE), right=C.side("thin", C.ACCENT))
        for c in C.iter_cells(ws, "D", r, c_last, r):
            c.font = C.font(C.T_BODY, False, C.INK)
            c.fill = C.NOFILL
            c.alignment = C.align("right", "center", 1)
            c.border = Border(bottom=C.side("hair", C.LINE))
        C.set_height(ws, r, ROW_H)


def _base_box(cell):
    m = C.side("medium", C.NAVY)
    cell.border = Border(left=m, right=m, top=m, bottom=m)
    cell.font = C.font(C.T_BODY, True, C.NAVY)


# =============================================================================== Aufbau
def apply(wb):
    if SHEET not in wb.sheetnames:
        return
    ws = wb[SHEET]
    ws.sheet_view.showGridLines = False

    # ---- Raster: C 62 → 46, Matrixspalten 11, Erklärspalte L:M
    for k, w in {"B": 3, "C": 46, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 3, "L": 44,
                 "M": 16}.items():
        ws.column_dimensions[k].width = w

    # ---- Seitenkopf: alte Titel in B entfernen (Kopf steht in C5:C7, linke Kante = Inhaltskante C)
    for coord in ("B5", "B6", "B7"):
        cell = ws[coord]
        if not C.is_formula(cell.value):
            cell.value = None

    for coord, text in LABELS.items():
        C.set_text(ws[coord], text)

    # ---- Break-even und Grenzwerte (Z. 8–18) + Lesehilfe F9:J18
    C.band_l1(ws, 8, "C", "J")
    _unmerge(ws, 9, 18, "F", "J")
    _clear(ws, "F", 9, "J", 18)
    for r in range(9, 19):
        lab, val = ws[f"C{r}"], ws[f"D{r}"]
        key = r in (10, 12, 14)
        sub = r in (11, 13, 15)
        lab.font = C.font(C.T_BODY, key, C.NAVY if key else C.INK)
        lab.alignment = C.align("left", "center", 2 if sub else 1)
        val.font = C.font(C.T_BODY, key, C.NAVY if key else C.INK)
        val.alignment = C.align("right", "center", 1)
        for c in (lab, val):
            c.fill = C.fill(C.TINT_XL) if key else C.NOFILL
            c.border = Border(bottom=C.side("hair", C.LINE))
        C.set_height(ws, r, ROW_H)
    ws["D14"].number_format = C.NUMFMT["pct2"]
    ws["D15"].number_format = C.NUMFMT["pct2"]
    C.callout(ws, "F", 9, "J", 10, 18, title="Lesehilfe")
    ws["F10"].value = LESEHILFE
    ws["F10"].alignment = C.align("left", "top", 1, wrap=True)
    C.set_height(ws, 9, ROW_H)
    for ref in ("D13",):
        ws.conditional_formatting.add(ref, FormulaRule(formula=["$D$13<0"], stopIfTrue=True,
                                                       font=Font(color=C.RED, bold=True), fill=C.fill(C.RED_BG)))
    C.set_height(ws, 19, 14)

    # ---- Jahr 1: Cashflow und DSCR (Band C20:J20, L2-Unterköpfe, Legendenzeilen)
    C.band_l1(ws, 20, "C", "J")
    for row in (21, 32):
        _clear(ws, "C", row, "J", row, values=False)
        C.subhead_l2(ws, row, "C", "J")
    for row in (22, 33):
        _clear(ws, "C", row, "J", row)
        C.set_height(ws, row, 18)
    _caption(ws, "C22", "Zeilen: Nettokaltmiete (Abweichung)  ·  Spalten: Sollzins Darlehen I (Δ %-Punkte)")
    _caption(ws, "C33", "Zeilen: Nettokaltmiete (Abweichung)  ·  Spalten: Sollzins Darlehen I (Δ %-Punkte)")
    _matrix(ws, 23, 24, 30, "J", "Miete ↓  ·  Sollzins →")
    _matrix(ws, 34, 35, 41, "J", "Miete ↓  ·  Sollzins →")
    for c in list(C.iter_cells(ws, "D", 23, "J", 23)) + list(C.iter_cells(ws, "D", 34, "J", 34)):
        c.number_format = MATRIX_COLFMT
    C.set_height(ws, 31, 14)
    C.set_height(ws, 42, 14)

    # Cashflow: Skala rot → weiß (0) → grün; die Enden werden symmetrisch gehalten, damit positive Werte nie rot sind
    _drop_cf(ws, "D24:J30", "D35:J41", "D47:H51")
    ws.conditional_formatting.add("D24:J30", ColorScaleRule(
        start_type="formula", start_value="MIN(-1,MIN($D$24:$J$30))", start_color="F2B8AE",
        mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="formula", end_value="MAX(1,MAX($D$24:$J$30))", end_color="A9D8BC"))
    # DSCR als Regel-Ampel (Schwellen aus der Konfiguration)
    _status_rules(ws, "D35:J41", "D35", "Ampel_DSCR_gruen", "Ampel_DSCR_gelb")
    # Basisfall hervorheben
    _base_box(ws["G27"])
    _base_box(ws["G38"])

    _legend(ws, "J22", [("■", C.RED, "Zuschuss"), ("■", C.GREEN, "Überschuss"), ("▣", C.NAVY, "aktuelle Annahme")])
    _legend(ws, "J33", [("■", C.RED, "< 1,0 kritisch"), ("■", C.AMBER, "1,0–1,2 knapp"),
                        ("■", C.GREEN, "≥ 1,2 bankfähig"), ("▣", C.NAVY, "aktuelle Annahme")])

    # ---- IRR-Matrix (C43:H43) und Verkauf
    _unmerge(ws, 43, 43, "C", "J")
    _clear(ws, "I", 43, "J", 43)
    C.safe_merge(ws, "C", 43, "H", 43)
    c43 = ws["C43"]                         # reine Anzeigeformel (Bandtitel), kürzer: passt in C:H
    if C.is_formula(c43.value) and "Haltedauer" in c43.value:
        c43.value = '="Eigenkapitalrendite bei Verkauf nach "&Haltedauer&" Jahren – Miete × Wertsteigerung"'
    C.band_l1(ws, 43, "C", "H")
    _clear(ws, "I", 44, "J", 63, values=False)
    for row in (44, 53):
        _clear(ws, "C", row, "H", row, values=False)
        C.subhead_l2(ws, row, "C", "H")
    _clear(ws, "C", 45, "H", 45)
    C.set_height(ws, 45, 18)
    _caption(ws, "C45", "Zeilen: Nettokaltmiete (Abweichung)  ·  Spalten: Wertsteigerung p. a.")
    _matrix(ws, 46, 47, 51, "H", "Miete ↓  ·  Wertsteigerung →")
    for c in C.iter_cells(ws, "D", 46, "H", 46):
        c.number_format = HEAD_PCT
    for c in C.iter_cells(ws, "D", 47, "H", 51):
        c.number_format = C.NUMFMT["pct2"]
    # aktuelle Wertsteigerungsannahme: Kopf blau/weiß, Zelle der Basis-Miete umrahmt
    ws.conditional_formatting.add("D46:H46", FormulaRule(
        formula=["ABS(D46-Wertsteigerung)<0.00005"], stopIfTrue=True,
        font=Font(color=C.WHITE, bold=True), fill=C.fill(C.BLUE)))
    # Basiszelle: Ampelfarbe + Rahmen in EINER Regel (Excel und LibreOffice werten dann identisch aus)
    nb = C.side("thin", C.NAVY)
    base = "ABS(D$46-Wertsteigerung)<0.00005"
    for cond, fg, bg in (("D49<Ampel_IRR_gelb", C.RED, M_RED), ("D49<Ampel_IRR_gruen", C.AMBER, M_AMB),
                         ("TRUE", C.GREEN, M_GRN)):
        ws.conditional_formatting.add("D49:H49", FormulaRule(
            formula=[f"AND({base},ISNUMBER(D49),{cond})"], stopIfTrue=True, font=Font(bold=True, color=fg),
            fill=C.fill(bg), border=Border(left=nb, right=nb, top=nb, bottom=nb)))
    irr_g = "Ampel_IRR_gruen"
    irr_y = "Ampel_IRR_gelb"
    _status_rules(ws, "D47:H51", "D47", irr_g, irr_y)
    g, y = _thresholds(wb)
    _legend(ws, "H45", [("■", C.RED, f"< {y} %"), ("■", C.AMBER, f"{y}–{g} %"), ("■", C.GREEN, f"≥ {g} %"),
                        ("▣", C.NAVY, "aktuelle Annahme")])
    C.set_height(ws, 52, 14)

    # Verkaufstabelle
    C.table_head(ws, 54, "C", "H")
    head = ws["C54"]
    C.set_text(head, "WERTSTEIGERUNG P. A.  →")
    head.alignment = C.align("right", "center", 1)
    for c in C.iter_cells(ws, "D", 54, "H", 54):
        c.alignment = C.align("right", "center", 1)
        c.number_format = HEAD_PCT
    for r in range(55, 64):
        _clear(ws, "C", r, "H", r, values=False)
        lab = ws[f"C{r}"]
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", 1)
        for c in C.iter_cells(ws, "D", r, "H", r):
            c.font = C.font(C.T_BODY, False, C.INK)
            c.alignment = C.align("right", "center", 1)
        C.hairline(ws, r, "C", "H")
        C.set_height(ws, r, ROW_H)
    C.total(ws, 59, "C", "H", level=2)
    C.total(ws, 62, "C", "H", level=3)
    ws.conditional_formatting.add("D62:H62", FormulaRule(formula=["D62<0"], stopIfTrue=True,
                                                         font=Font(color=C.RED, bold=True), fill=C.fill(C.RED_BG)))
    for c in C.iter_cells(ws, "D", 63, "H", 63):
        c.number_format = C.NUMFMT["mult2"]

    # ---- „Wie lesen?“ neben den Objekten (L8 bündig mit C8, Kästen auf Höhe ihrer Tabelle)
    _unmerge(ws, 8, 63, "L", "M")
    _clear(ws, "L", 8, "M", 63)
    C.safe_merge(ws, "L", 8, "M", 8)
    C.band_l1(ws, 8, "L", "M", "Wie lesen?")
    width = C.span_px(ws, "L", "M")
    last_body = 0
    for head_row, title, text in READ_BOXES:
        lines = C.lines_needed(text, width - 12, C.T_SMALL)
        need = lines * 12 + 8
        r, acc = head_row + 1, 0
        while acc < need:
            acc += ws.row_dimensions[r].height or 15
            r += 1
        C.callout(ws, "L", head_row, "M", head_row + 1, r - 1, title=title)
        body = ws[f"L{head_row + 1}"]
        body.value = text
        body.alignment = C.align("left", "center", 1, wrap=True)
        last_body = r - 1
    # callout setzt die Kopfzeile auf 20 pt – Kopfzeilen, die zugleich Tabellenzeilen sind, zurück ins Raster
    C.set_height(ws, 9, ROW_H)
    C.set_height(ws, 21, C.H_HEAD)
    C.set_height(ws, 32, C.H_HEAD)
    C.set_height(ws, 44, C.H_HEAD)

    # ---- Diagramm IRR je Wertsteigerung: unter dem IRR-Kasten bis Ende M, unten bündig mit Z. 63
    for ch in ws._charts:
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = C.col("L") - 1, 0, last_body + 1, 0
        a.to.col, a.to.colOff, a.to.row, a.to.rowOff = C.col("M"), 0, 63, 0

    # ---- Nebenrechnung einklappen (Gliederung), Hinweis außerhalb der Gruppe
    for r in range(FIRST, LAST + 1):
        ws.row_dimensions[r].outline_level = 1
        ws.row_dimensions[r].hidden = True
    if ws.sheet_properties.outlinePr is None:
        ws.sheet_properties.outlinePr = Outline(summaryBelow=True, summaryRight=True)
    else:
        ws.sheet_properties.outlinePr.summaryBelow = True
    note = ws["C64"]
    C.set_text(note, "Nebenrechnung zur IRR-Matrix ausgeblendet (Zeilen 65–115) – "
                     "über das Gliederungssymbol [+] am linken Rand einblenden.")
    note.font = C.font(C.T_SMALL, False, C.MUTED, italic=True)
    note.alignment = C.align("left", "center", 1)
    C.set_height(ws, 64, 22)
    for r in range(90, 115):
        cell = ws[f"C{r}"]
        if isinstance(cell.value, str):
            v = cell.value.replace("Miete +0%", "Miete Basis").replace("%", " %").replace(" | ", " · ")
            C.set_text(cell, v)
    # Jahresspalten der Nebenrechnung (N:AS) mit einklappen – rechts vom Blatt bleibt nichts Halbfertiges sichtbar
    ws.column_dimensions.group("N", "AS", hidden=True, outline_level=1)
    for c in C.iter_cells(ws, "D", 88, "AR", 88):
        c.number_format = C.NUMFMT["year"]


def _thresholds(wb):
    """Aktuelle IRR-Schwellen (Konfiguration) für die statische Legende, deutsch formatiert."""
    try:
        dn_g, dn_y = wb.defined_names["Ampel_IRR_gruen"], wb.defined_names["Ampel_IRR_gelb"]
        vals = []
        for dn in (dn_g, dn_y):
            (sheet, ref), = list(dn.destinations)
            vals.append(float(wb[sheet][ref.replace("$", "")].value))
        return tuple(f"{v * 100:.0f}".replace(".", ",") for v in vals)
    except Exception:
        return "6", "3"
