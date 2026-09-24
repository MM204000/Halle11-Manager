"""Blatt „Sensitivität“: lesbare Matrizen, gedämpfte Heatmaps, Erklärungen im Rail, Nebenrechnung eingeklappt.

Stil: ruhiges Rechenblatt (feine Linien, dezente Akzente). Komponenten ausschließlich aus core.py:
section (Ebene 1/2, „↑ Übersicht“), callout_box (Rail-Erklärung je Abschnitt), sum_row (Summenstufen),
status_cf / status_pill / neg_red (Statusfarben-Disziplin), NUMFMT (Zahlenformat-Katalog).
Nur Darstellung – Werte in C24:C51 bleiben Zahlen (die Formeln nutzen $C24 …), sie erhalten lediglich ein
sprechendes Zahlenformat. Jeder Hauptabschnitt spannt C:J, jede Rail-Box beginnt in der Zeile ihres Abschnittskopfs.
"""
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.properties import Outline

import core as C

SHEET = "Sensitivität"
FIRST, LAST = 65, 115                      # Nebenrechnung (eingeklappt)
ROW_H = 16                                 # Datenzeilen des Blatts
MATRIX_ROWFMT = '"Miete +"0 %;"Miete −"0 %;"Miete Basis"'
MATRIX_COLFMT = '+0.0 %;"−"0.0 %;"Basis"'
HEAD_PCT = "0.0 %;-0.0 %;0.0 %"            # Kopfwerte: 0 % ist ein echter Wert, kein „–“

# Gedämpfte Heatmap-Töne (P2-03): DSCR innerhalb „kritisch“ abgestuft, darüber knapp / bankfähig;
# IRR in drei flachen Stufen; Schrift in allen Matrizen 1A1D21.
DSCR_MIN, DSCR_MID, DSCR_MAX = "F4C7C3", "FBE3E0", "FDF3F2"
DSCR_AMB, DSCR_GRN = "FEF0C7", "DCEFE3"
Q_RED, Q_AMB, Q_GRN = "FBE9E7", "FDF1E3", "E8F3EC"
CF_NEG, CF_POS = "F6D6D1", "CFE8DA"         # Cashflow-Skala (entsättigt): Zuschuss ← 0 (weiß) → Überschuss

LABELS = {
    "C10": "Break-even-Miete (Cashflow vor Steuern = 0)",
    "C11": "Puffer zur aktuellen Miete",
    "C12": "Break-even-Miete (Cashflow nach Steuern = 0) ¹",
    "C13": "Puffer zur aktuellen Miete",
    "C14": "Max. Sollzins (Cashflow vor Steuern ≥ 0)",
    "C15": "Aktueller gewichteter Sollzins",
    "C18": "Zinsänderungsrisiko: +1 %-Punkt Anschlusszins p. a.",
    "C20": "Jahr 1: Cashflow und Kapitaldienstdeckung – Miete × Sollzins",
    "C21": "Cashflow nach Steuern (€ / Monat)",
    "C32": "Kapitaldienstdeckung (DSCR = NOI / Kapitaldienst)",
    "C44": "IRR nach Steuern p. a. – vollständige Neuberechnung der Zahlungsreihe",
    "C53": "Verkauf nach der Haltedauer je Wertsteigerung (€, Basis-Miete)",
}
# Einordnung je Zeile (G, 9 pt grau) – ersetzt die frühere Lesehilfe-Box F9:J18 (P2-04)
NOTES = {
    9: "Ausgangswert aller Break-even-Größen",
    10: "Ab dieser Miete tragen sich Bewirtschaftung und Kapitaldienst",
    11: "Abstand der aktuellen Miete zur Break-even-Miete",
    12: "Wie oben, zusätzlich mit Steuerwirkung des Objekts",
    13: "Abstand zur Break-even-Miete nach Steuern",
    14: "Höchster Zins, bei dem der Cashflow vor Steuern ≥ 0 bleibt",
    15: "Gewichteter Zins aller Darlehen im ersten Jahr",
    16: "Kaufpreis, bei dem die Zielrendite erreicht wäre",
    17: "Kaufpreis für die Mindestrendite",
    18: "Mehrbelastung auf die Restschuld am Ende der Zinsbindung",
}
RAIL = [
    # (Abschnittszeile, Titel, Text)
    (8, "Lesehilfe und Grenzen",
     "¹ Bei voller Steuerwirkung: Ein steuerlicher Verlust wird sofort mit anderen Einkünften verrechnet "
     "(Privatperson). In der GmbH oder mit Verlustvortrag wirkt die Steuer erst, wenn das Objekt Gewinne erzielt.\n\n"
     "Grenzen: Die Jahr-1-Größen sind lineare Näherungen bei konstantem Steuersatz. Nicht abgebildet sind "
     "Wechselwirkungen mit dem Mietausfall, Progressionseffekte, Vorfälligkeitsentschädigung und die "
     "Ausschüttungsbelastung der GmbH."),
    (20, "So lesen Sie die Matrizen",
     "Cashflow: Jede Zelle zeigt den Cashflow nach Steuern pro Monat im ersten Jahr, wenn die Miete um den "
     "Zeilenwert und der Sollzins des Hauptdarlehens um den Spaltenwert abweicht. Grün = Überschuss, rot = Zuschuss.\n\n"
     "DSCR: Einnahmenüberschuss (NOI) geteilt durch den Kapitaldienst. Banken erwarten meist 1,1 bis 1,2; unter 1,0 "
     "trägt die Bonität des Investors den Kredit. Die Rottöne werden heller, je näher der Wert an 1,0 liegt."),
    (43, "IRR-Matrix",
     "Für jede Kombination wird die vollständige Zahlungsreihe über die Haltedauer neu gerechnet (ausgeblendete "
     "Nebenrechnung, Zeilen 65–115): Die Mietänderung wirkt in jedem Jahr mit der Mietsteigerung fort, die "
     "Wertsteigerung bestimmt Verkaufspreis, Verkaufskosten und Steuer beim Verkauf."),
]
SECTIONS = (8, 20, 43)


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


def _legend(ws, coord, items):
    """Legende rechtsbündig: (Stufe|None, Text) – Punkte in Statusfarbe, „▢“ Navy für die aktuelle Annahme."""
    parts = []
    for i, (lvl, text) in enumerate(items):
        if i:
            parts.append(("   ·   ", C.T_MICRO, False, C.MUTED))
        glyph, color = ("▢", C.NAVY) if lvl is None else (C.STATUS_DOT, C.STATUS_COLORS[lvl][0])
        parts += [(glyph + " ", C.T_MICRO, True, color), (text, C.T_MICRO, False, C.MUTED)]
    cell = ws[coord]
    cell.value = C.rich(parts)
    cell.font = C.font(C.T_MICRO, False, C.MUTED)
    cell.alignment = C.align("right", "center")


def _caption(ws, coord, text):
    cell = ws[coord]
    C.set_text(cell, text)
    cell.font = C.font(C.T_MICRO, False, C.MUTED)
    cell.alignment = C.align("left", "center", 1)


def _matrix(ws, head, r1, r2, c_last, corner, base_row):
    """Matrix: Kopfzeile C:J (Spaltenwerte), Zeilenköpfe rechtsbündig an der Heatmap, Datenzellen mit Haarlinie.
    Basis-Zeile: Kopf E7EEF7 fett Navy (P2-03)."""
    C.section(ws, head, "C", "J", None, level=2, variant="fill", caps=False)
    corner_cell = ws[f"C{head}"]
    C.set_text(corner_cell, corner)
    corner_cell.font = C.font(C.T_MICRO, True, C.BLUE)
    corner_cell.alignment = C.align("right", "center", 1)
    corner_cell.border = Border(bottom=C.side("thin", C.ACCENT), right=C.side("thin", C.ACCENT))
    for c in C.iter_cells(ws, "D", head, "J", head):
        c.font = C.font(C.T_MICRO, True, C.BLUE)
        c.alignment = C.align("right", "center", 1)
    C.set_height(ws, head, C.H_HEAD)
    for r in range(r1, r2 + 1):
        base = r == base_row
        lab = ws[f"C{r}"]
        lab.number_format = MATRIX_ROWFMT
        lab.font = C.font(C.T_SMALL, True, C.NAVY if base else C.BLUE)
        lab.fill = C.fill(C.TINT) if base else C.NOFILL
        lab.alignment = C.align("right", "center", 1)
        lab.border = Border(bottom=C.side("hair", C.LINE), right=C.side("thin", C.ACCENT))
        for c in C.iter_cells(ws, "D", r, "J", r):
            c.font = C.font(C.T_BODY, False, C.INK)
            c.fill = C.NOFILL
            c.alignment = C.align("right", "center", 1)
            c.border = Border(bottom=C.side("hair", C.LINE))   # Linien bis J – gleiche rechte Kante je Abschnitt
        C.set_height(ws, r, ROW_H)


def _base_head(cell):
    cell.fill = C.fill(C.TINT)
    cell.font = C.font(C.T_MICRO, True, C.NAVY)


def _base_box(cell):
    m = C.side("medium", C.NAVY)
    cell.border = Border(left=m, right=m, top=m, bottom=m)
    cell.font = C.font(C.T_BODY, True, C.INK)


def _fill_rule(ws, ref, formula, color, **kw):
    ws.conditional_formatting.add(ref, FormulaRule(formula=[formula], stopIfTrue=True, fill=C.fill(color), **kw))


def _uebersicht(ws, row):
    """P3-03: „↑ Übersicht“ rechtsbündig im Abschnittskopf (Zell-Link auf den Seitenanfang)."""
    cell = ws[f"J{row}"]
    C.text_link(cell, "↑ Übersicht", ws.title, size=C.T_MICRO, bold=False)
    cell.alignment = C.align("right", "center", 1)


def _thresholds(wb, kpi, digits=0, factor=100):
    """Aktuelle Schwellen (Konfiguration) für die statische Legende, deutsch formatiert: (grün, gelb)."""
    try:
        vals = []
        for suffix in ("gruen", "gelb"):
            (sheet, ref), = list(wb.defined_names[f"Ampel_{kpi}_{suffix}"].destinations)
            vals.append(float(wb[sheet][ref.replace("$", "")].value))
        return tuple(f"{v * factor:.{digits}f}".replace(".", ",") for v in vals)
    except Exception:
        return None


# =============================================================================== Aufbau
def apply(wb):
    if SHEET not in wb.sheetnames:
        return
    ws = wb[SHEET]
    ws.sheet_view.showGridLines = False

    # ---- Raster: C 46, Matrixspalten 11, Erklärspalte L:M
    for k, w in {"B": 3, "C": 46, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 3, "L": 44,
                 "M": 16}.items():
        ws.column_dimensions[k].width = w

    # ---- Seitenkopf (P1-18): Überzeile = Reitername; alte Titel in B entfernen (linke Kante = Inhaltskante C)
    for coord in ("B5", "B6", "B7"):
        cell = ws[coord]
        if not C.is_formula(cell.value):
            cell.value = None
    C.set_text(ws["C5"], "SENSITIVITÄT")
    ws["C5"].font = C.font(C.T_MICRO, True, C.BLUE)
    ws["C5"].alignment = C.align("left", "bottom")

    for coord, text in LABELS.items():
        C.set_text(ws[coord], text)

    # ---- Rail frei machen (die Boxen entstehen unten neu, je Abschnitt genau eine)
    _unmerge(ws, 8, 63, "L", "M")
    _clear(ws, "L", 8, "M", 63)

    # ---- Abschnitt 1: Break-even und Grenzwerte (C:J) – Wert + Status + Einordnung je Zeile
    _unmerge(ws, 8, 18, "C", "J")
    _clear(ws, "E", 9, "J", 18)
    C.section(ws, 8, "C", "J")
    _uebersicht(ws, 8)
    for r in range(9, 19):
        lab, val = ws[f"C{r}"], ws[f"D{r}"]
        key = r in (10, 12, 14)
        sub = r in (11, 13, 15)
        lab.font = C.font(C.T_BODY, key, C.NAVY if key else C.INK)
        lab.alignment = C.align("left", "center", 2 if sub else 1)
        val.font = C.font(C.T_BODY, key, C.NAVY if key else C.INK)
        val.alignment = C.align("right", "center", 1)
        for c in C.iter_cells(ws, "C", r, "J", r):
            c.fill = C.NOFILL
            c.border = Border(bottom=C.side("hair", C.LINE))
        note = ws[f"E{r}"]
        C.set_text(note, NOTES.get(r))
        note.font = C.font(C.T_SMALL, False, C.MUTED)
        note.alignment = C.align("left", "center", 1)
        C.set_height(ws, r, C.H_ROW)
    ws["D14"].number_format = C.NUMFMT["pct2"]
    ws["D15"].number_format = C.NUMFMT["pct2"]
    # Status (P1-10): Puffer < 0 bzw. Grenzzins < Ist → rot; Wert neutral sonst; Status als Chip in E
    _drop_cf(ws, "D11", "D13", "D14", "D62", "D62:H62")
    status = {
        11: ([("AND(ISNUMBER($D$11),$D$11<0)", "red"), ("ISNUMBER($D$11)", "green")],
             {"red": "Zuschuss nötig", "green": "Puffer vorhanden"}),
        13: ([("AND(ISNUMBER($D$13),$D$13<0)", "red"), ("ISNUMBER($D$13)", "green")],
             {"red": "Zuschuss nötig", "green": "Puffer vorhanden"}),
        14: ([("AND(ISNUMBER($D$14),$D$14<$D$15)", "red"), ("ISNUMBER($D$14)", "green")],
             {"red": "unter Ist-Zins", "green": "über Ist-Zins"}),
    }
    for r, (conds, words) in status.items():
        C.status_cf(ws, f"D{r}", conds[:1])
        C.safe_merge(ws, "I", r, "J", r)
        C.status_pill(ws, ws[f"I{r}"], conditions=conds, style="chip", words=words, h="right")
    C.set_height(ws, 19, C.H_GAP)

    # ---- Abschnitt 2: Jahr 1 – Cashflow und DSCR (C:J)
    _unmerge(ws, 20, 20, "C", "J")
    C.section(ws, 20, "C", "J")
    _uebersicht(ws, 20)
    for row in (21, 32):
        _clear(ws, "D", row, "J", row)
        C.section(ws, row, "C", "J", level=2, variant="line")
    for row in (22, 33):
        _clear(ws, "C", row, "J", row)
        C.set_height(ws, row, C.H_ROW)
    _caption(ws, "C22", "Zeilen: Nettokaltmiete (Abweichung)  ·  Spalten: Sollzins Darlehen I (Δ %-Punkte)")
    _caption(ws, "C33", "Zeilen: Nettokaltmiete (Abweichung)  ·  Spalten: Sollzins Darlehen I (Δ %-Punkte)")
    _matrix(ws, 23, 24, 30, "J", "Miete ↓  ·  Sollzins →", 27)
    _matrix(ws, 34, 35, 41, "J", "Miete ↓  ·  Sollzins →", 38)
    for c in list(C.iter_cells(ws, "D", 23, "J", 23)) + list(C.iter_cells(ws, "D", 34, "J", 34)):
        c.number_format = MATRIX_COLFMT
    for c in C.iter_cells(ws, "D", 35, "J", 41):
        c.number_format = C.NUMFMT["dscr"]
    _base_head(ws["G23"])
    _base_head(ws["G34"])
    C.set_height(ws, 31, C.H_GAP)
    C.set_height(ws, 42, C.H_GAP)

    _drop_cf(ws, "D24:J30", "D35:J41", "D46:H46", "D47:H51", "D49:H49")
    # Cashflow: entsättigte Skala rot → weiß (0) → grün; Enden symmetrisch, damit positive Werte nie rot sind
    ws.conditional_formatting.add("D24:J30", ColorScaleRule(
        start_type="formula", start_value="MIN(-1,MIN($D$24:$J$30))", start_color=CF_NEG,
        mid_type="num", mid_value=0, mid_color="FFFFFF",
        end_type="formula", end_value="MAX(1,MAX($D$24:$J$30))", end_color=CF_POS))
    # DSCR: ≥ grün → DCEFE3, ≥ gelb → FEF0C7 (je Stop), darunter Rot-Abstufung bis 1,0 (P2-03)
    _fill_rule(ws, "D35:J41", "AND(ISNUMBER(D35),D35>=Ampel_DSCR_gruen)", DSCR_GRN)
    _fill_rule(ws, "D35:J41", "AND(ISNUMBER(D35),D35>=Ampel_DSCR_gelb)", DSCR_AMB)
    ws.conditional_formatting.add("D35:J41", ColorScaleRule(
        start_type="min", start_color=DSCR_MIN, mid_type="percentile", mid_value=50, mid_color=DSCR_MID,
        end_type="max", end_color=DSCR_MAX))
    _base_box(ws["G27"])
    _base_box(ws["G38"])

    d = _thresholds(wb, "DSCR", 1, 1) or ("1,2", "1,0")
    _legend(ws, "J22", [("red", "Zuschuss"), ("green", "Überschuss"), (None, "aktuelle Annahme")])
    _legend(ws, "J33", [("red", f"< {d[1]}× kritisch"), ("amber", f"{d[1]}–{d[0]}× knapp"),
                        ("green", f"≥ {d[0]}× bankfähig"), (None, "aktuelle Annahme")])

    # ---- Abschnitt 3: IRR-Matrix und Verkauf (C:J, P2-04)
    _unmerge(ws, 43, 43, "C", "J")
    _clear(ws, "I", 43, "J", 43)
    C.safe_merge(ws, "C", 43, "I", 43)
    c43 = ws["C43"]                         # reine Anzeigeformel (Bandtitel)
    if C.is_formula(c43.value) and "Haltedauer" in c43.value:
        c43.value = '="Eigenkapitalrendite bei Verkauf nach "&Haltedauer&" Jahren – Miete × Wertsteigerung"'
    C.section(ws, 43, "C", "J")
    _uebersicht(ws, 43)
    _clear(ws, "I", 44, "J", 63, values=False)
    for row in (44, 53):
        _clear(ws, "D", row, "J", row, values=False)
        C.section(ws, row, "C", "J", level=2, variant="line")
    _clear(ws, "C", 45, "J", 45)
    C.set_height(ws, 45, C.H_ROW)
    _caption(ws, "C45", "Zeilen: Nettokaltmiete (Abweichung)  ·  Spalten: Wertsteigerung p. a.")
    _matrix(ws, 46, 47, 51, "H", "Miete ↓  ·  Wertsteigerung →", 49)
    for c in C.iter_cells(ws, "D", 46, "H", 46):
        c.number_format = HEAD_PCT
    for c in C.iter_cells(ws, "D", 47, "H", 51):
        c.number_format = C.NUMFMT["pct1"]
    # aktuelle Wertsteigerungsannahme: Kopf E7EEF7 fett Navy (keine Navy-Vollfläche mehr)
    ws.conditional_formatting.add("D46:H46", FormulaRule(
        formula=["ABS(D46-Wertsteigerung)<0.00005"], stopIfTrue=True,
        font=Font(color=C.NAVY, bold=True), fill=C.fill(C.TINT)))
    # Basiszelle (Basis-Miete × aktuelle Wertsteigerung): Stufenfläche + Navy-Rahmen in EINER Regel, alles in EINEM
    # Bereich D47:H51 (LibreOffice wertet je Zelle nur einen Bereich aus) – Basis-Zeile = Zeilenwert 0 in Spalte C
    nb = C.side("medium", C.NAVY)
    base = "ISNUMBER(D47),$C47=0,ABS(D$46-Wertsteigerung)<0.00005"   # flaches AND (LibreOffice: kein AND in AND)
    for cond, bg in (("D47<Ampel_IRR_gelb", Q_RED), ("D47<Ampel_IRR_gruen", Q_AMB), ("TRUE", Q_GRN)):
        ws.conditional_formatting.add("D47:H51", FormulaRule(
            formula=[f"AND({base},{cond})"], stopIfTrue=True, font=Font(bold=True),
            fill=C.fill(bg), border=Border(left=nb, right=nb, top=nb, bottom=nb)))
    _fill_rule(ws, "D47:H51", "AND(ISNUMBER(D47),D47<Ampel_IRR_gelb)", Q_RED)
    _fill_rule(ws, "D47:H51", "AND(ISNUMBER(D47),D47<Ampel_IRR_gruen)", Q_AMB)
    _fill_rule(ws, "D47:H51", "ISNUMBER(D47)", Q_GRN)
    g, y = _thresholds(wb, "IRR") or ("6", "3")
    _legend(ws, "J45", [("red", f"< {y} %"), ("amber", f"{y}–{g} %"), ("green", f"≥ {g} %"),
                        (None, "aktuelle Annahme")])
    C.set_height(ws, 52, C.H_GAP)

    # Verkaufstabelle (Kopf und Linien bis J)
    C.section(ws, 54, "C", "J", None, level=2, variant="fill", caps=False)
    head = ws["C54"]
    C.set_text(head, "Wertsteigerung p. a.  →")
    head.font = C.font(C.T_MICRO, True, C.BLUE)
    head.alignment = C.align("right", "center", 1)
    for c in C.iter_cells(ws, "D", 54, "H", 54):
        c.alignment = C.align("right", "center", 1)
        c.number_format = HEAD_PCT
    for r in range(55, 64):
        _clear(ws, "C", r, "J", r, values=False)
        lab = ws[f"C{r}"]
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", 1)
        for c in C.iter_cells(ws, "D", r, "J", r):
            c.font = C.font(C.T_BODY, False, C.INK)
            c.alignment = C.align("right", "center", 1)
        C.hairline(ws, r, "C", "J")
        C.set_height(ws, r, ROW_H)
    C.sum_row(ws, 59, "C", "J", "sub")
    C.sum_row(ws, 62, "C", "J", "result", value_from="D", neg=True)
    for c in C.iter_cells(ws, "D", 63, "H", 63):
        c.number_format = C.NUMFMT["mult2"]

    # ---- Rail: je Abschnitt genau eine Erklär-Box, Kopf in der Zeile des Abschnittskopfs (P2-04)
    width = C.span_px(ws, "L", "M")
    last_body = 0
    for head_row, title, text in RAIL:
        need = C.px_pt(C.lines_needed_metric(text, width, C.T_SMALL, False, 1) * C.line_pt(C.T_SMALL) + 12)
        r, acc = head_row + 1, 0
        while acc < need:
            acc += ws.row_dimensions[r].height or 15
            r += 1
        C.callout_box(ws, "L", head_row, "M", head_row + 1, r - 1, title=title, text=text, pill=False, fit=None)
        C.set_height(ws, head_row, C.H_BAND)       # Kopf so hoch wie der Abschnittskopf daneben
        last_body = r - 1

    # ---- Diagramm IRR je Wertsteigerung: bündig unter der IRR-Box bis Ende M, unten bündig mit Z. 63 (P2-09)
    for ch in ws._charts:
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = C.col("L") - 1, 0, last_body, 0
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
    C.cf_close(ws)
