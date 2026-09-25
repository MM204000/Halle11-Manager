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
HEAD_PCT = '0.0 %;"−"0.0 %;0.0 %'          # Kopfwerte: 0 % ist ein echter Wert, kein „–“

# Heatmaps (Runde 3, P21/P41; Runde 5 „Midnight & Gold“): Statusfarben nie als Fläche – die Fläche zeigt nur die
# Höhe des Werts in einer ruhigen Sequenz aus den core-Tokens (niedrig TINT_XL warm-weiß → ICE → zarter Nachtblau-
# Ton, 16 % BLUE auf Weiß). Der Status steckt allein in der Schriftfarbe (Rot B42318 darauf ≥ 5,0:1).
# Die aktuelle Annahme (Basis-Zeile/-Spalte) ist mappenweit GOLD_BG hinterlegt, die Basiszelle navy umrahmt.
SEQ_LO, SEQ_MID, SEQ_HI = C.TINT_XL, C.ICE, C.mix(C.BLUE, C.WHITE, 0.84)
BASE_BG = C.GOLD_BG

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
    "C44": "IRR nach Steuern p. a. (volle Neuberechnung)",
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
     "Zeilenwert und der Sollzins des Hauptdarlehens um den Spaltenwert abweicht. Rote Schrift = Zuschuss "
     "(Cashflow unter 0 €).\n\n"
     "DSCR: Einnahmenüberschuss (NOI) geteilt durch den Kapitaldienst. Unter der Ampel-Schwelle „prüfen“ trägt die "
     "Bonität des Investors einen Teil des Kapitaldiensts; die Schwellen stehen in der Legende über der Matrix.\n\n"
     "Farbton: Je kräftiger der Blauton, desto höher der Wert – unabhängig vom Status. Goldton und Rahmen "
     "markieren die aktuelle Annahme (Miete Basis, Sollzins wie eingegeben)."),
    (43, "IRR-Matrix",
     "Für jede Kombination wird die vollständige Zahlungsreihe über die Haltedauer neu gerechnet: Die Mietänderung wirkt in jedem Jahr mit der Mietsteigerung fort, die "
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


def _heat(ws, ref, conds):
    """Heatmap in EINEM Bereich (LibreOffice wertet je Zelle nur einen Bereich aus): zuerst die Statusschrift
    (erste zutreffende Regel), danach die Blau-Sequenz nach Wertgröße."""
    for cond, lvl in conds:
        ws.conditional_formatting.add(ref, FormulaRule(formula=[cond], stopIfTrue=False,
                                                       font=Font(color=C.STATUS_COLORS[lvl][0])))
    ws.conditional_formatting.add(ref, ColorScaleRule(
        start_type="min", start_color=SEQ_LO, mid_type="percentile", mid_value=50, mid_color=SEQ_MID,
        end_type="max", end_color=SEQ_HI))


def _legend_formula(ws, coord, formula):
    """Schwellen-Legende als Anzeigeformel (eine Formulierung, aus den Ampel-Namen der Konfiguration)."""
    cell = ws[coord]
    cell.value = formula
    cell.font = C.font(C.T_LABEL, False, C.MUTED)
    cell.alignment = C.align("right", "center")


def _matrix(ws, head, r1, r2, c_last, corner, base_row):
    """Matrix: Kopfzeile C:J (Spaltenwerte), Zeilenköpfe rechtsbündig an der Heatmap, Datenzellen mit Haarlinie.
    Basis-Zeile: Kopf GOLD_BG fett Navy (P2-03, Runde 5)."""
    C.section(ws, head, "C", c_last, None, level=2, variant="fill", caps=False)
    corner_cell = ws[f"C{head}"]
    C.set_text(corner_cell, corner)
    corner_cell.font = C.font(C.T_LABEL, True, C.BLUE)
    corner_cell.alignment = C.align("right", "center", 1)
    corner_cell.border = Border(bottom=C.side("thin", C.ACCENT), right=C.side("thin", C.ACCENT))
    for c in C.iter_cells(ws, "D", head, c_last, head):
        c.font = C.font(C.T_LABEL, True, C.BLUE)
        c.alignment = C.align("right", "center", 1)
    C.set_height(ws, head, C.H_HEAD)
    for r in range(r1, r2 + 1):
        base = r == base_row
        lab = ws[f"C{r}"]
        lab.number_format = MATRIX_ROWFMT
        lab.font = C.font(C.T_SMALL, True, C.NAVY if base else C.BLUE)
        lab.fill = C.fill(BASE_BG) if base else C.NOFILL
        lab.alignment = C.align("right", "center", 1)
        lab.border = Border(bottom=C.side("hair", C.LINE), right=C.side("thin", C.ACCENT))
        for c in C.iter_cells(ws, "D", r, c_last, r):
            c.font = C.font(C.T_BODY, False, C.INK)
            c.fill = C.NOFILL
            c.alignment = C.align("right", "center", 1)
            c.border = Border(bottom=C.side("hair", C.LINE))   # P30: Linien enden an der letzten Datenspalte
        C.set_height(ws, r, ROW_H)


def _base_head(cell):
    cell.fill = C.fill(BASE_BG)
    cell.font = C.font(C.T_LABEL, True, C.NAVY)


def _base_box(cell):
    m = C.side("medium", C.NAVY)
    cell.border = Border(left=m, right=m, top=m, bottom=m)
    cell.font = C.font(C.T_BODY, True, C.INK)


def _fill_rule(ws, ref, formula, color, **kw):
    ws.conditional_formatting.add(ref, FormulaRule(formula=[formula], stopIfTrue=True, fill=C.fill(color), **kw))


def _uebersicht(ws, row, c="J"):
    """P3-03: „↑ Übersicht“ rechtsbündig im Abschnittskopf (Zell-Link auf den Seitenanfang)."""
    cell = ws[f"{c}{row}"]
    C.text_link(cell, "↑ Übersicht", ws.title, size=C.T_MICRO, bold=False)
    cell.alignment = C.align("right", "center", 1)


def _nav_row(ws, row, back, nxt):
    """Leerzeile · Buttons (secondary / primary) · Leerzeile · Seitenfuß (aus global_rules.early hierher verlegt)."""
    old = next((c.row for c in ws._cells.values() if isinstance(c.value, str) and c.value == C.FOOTER_1), None)
    if old is not None and old != row + 2:
        for r in (old, old + 1):
            for c in C.iter_cells(ws, "A", r, "M", r):
                if not C.is_formula(c.value):
                    c.value = None
                c.border = Border()
            ws.row_dimensions[r].height = None
    C.footer(ws, row + 2, "C", "M")
    C.set_height(ws, row - 1, C.H_GAP)
    for c in C.iter_cells(ws, "C", row, "M", row):     # Reihe frei machen (alte Button-Lage L:M)
        c.hyperlink = None
        c.fill = C.NOFILL
        c.border = Border()
        if not C.is_formula(c.value):
            c.value = None
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= row <= mr.max_row:
            ws.unmerge_cells(str(mr))
    C.btn_row(ws, row, [dict(c1=back[0], c2=back[1], text=back[2], target=back[3], kind="secondary"),
                        dict(c1=nxt[0], c2=nxt[1], text=nxt[2], target=nxt[3], kind="primary")])
    C.set_height(ws, row + 1, C.H_GAP)


# =============================================================================== Aufbau
def apply(wb):
    if SHEET not in wb.sheetnames:
        return
    ws = wb[SHEET]
    ws.sheet_view.showGridLines = False

    # ---- Raster: C 46, Matrixspalten 11, Erklärspalte L:M
    for k, w in {"A": 1.5, "B": 3, "C": 46, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 3, "L": 44,
                 "M": 16}.items():
        ws.column_dimensions[k].width = w

    # ---- Seitenkopf (P04): Z. 5 Eyebrow „BERECHNUNG › SENSITIVITÄT“ · Z. 6 Titel + Objekt rechts ·
    #      Z. 7 Untertitel + Adresse/„Erstellt für“ rechts – rechte Kante = Inhaltskante M (wie S01–S12)
    #      P39: A + B = 4,5 – Titel und Inhalt beginnen an derselben x-Position wie auf den B-Blättern
    for coord in ("B5", "B6", "B7"):
        cell = ws[coord]
        if not C.is_formula(cell.value):
            cell.value = None
    C.page_header(ws, "C", "M", ("Berechnung", "Sensitivität"), "Sensitivität", None,
                  context=("=Obj_Name",
                           '=Obj_Adresse&IFERROR(IF(Erstellt_fuer="","","  ·  Erstellt für "&Erstellt_fuer),"")'))

    for coord, text in LABELS.items():
        C.set_text(ws[coord], text)
    # P21: keine Farbwörter in Beschriftungen – Ziel- bzw. Mindestwert statt „(grün)“/„(gelb)“ (reine Anzeigeformeln)
    for coord, old, new in (("C16", "(grün)", "(Zielwert)"), ("C17", "(gelb)", "(Mindestwert)")):
        v = ws[coord].value
        if isinstance(v, str) and old in v:
            ws[coord].value = v.replace(old, new)

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
    # P21: nur „erfüllt / prüfen / kritisch“ – die Erläuterung steht grau in der Einordnungsspalte E
    status = {
        11: [("AND(ISNUMBER($D$11),$D$11<0)", "red"), ("ISNUMBER($D$11)", "green")],
        13: [("AND(ISNUMBER($D$13),$D$13<0)", "red"), ("ISNUMBER($D$13)", "green")],
        14: [("AND(ISNUMBER($D$14),$D$14<$D$15)", "red"), ("ISNUMBER($D$14)", "green")],
    }
    for r, conds in status.items():
        C.status_cf(ws, f"D{r}", conds[:1])
        C.safe_merge(ws, "I", r, "J", r)
        C.status_pill(ws, ws[f"I{r}"], conditions=conds, style="chip", h="right")
    C.set_height(ws, 19, C.H_GAP)

    # ---- Abschnitt 2: Jahr 1 – Cashflow und DSCR (C:J)
    _unmerge(ws, 20, 20, "C", "J")
    C.section(ws, 20, "C", "J")
    _uebersicht(ws, 20)
    for row in (21, 32):
        _clear(ws, "D", row, "J", row)
        C.section(ws, row, "C", "J", level=2, variant="line")
    for row in (22, 33):                    # Legende steht rechts im Unterabschnittskopf, hier nur eine Fuge
        _clear(ws, "C", row, "J", row)
        C.set_height(ws, row, 6)
    corner = "Miete ↓  ·  Sollzins Darlehen I (Δ %-Pkt.) →"
    _matrix(ws, 23, 24, 30, "J", corner, 27)
    _matrix(ws, 34, 35, 41, "J", corner, 38)
    for c in list(C.iter_cells(ws, "D", 23, "J", 23)) + list(C.iter_cells(ws, "D", 34, "J", 34)):
        c.number_format = MATRIX_COLFMT
    for c in C.iter_cells(ws, "D", 35, "J", 41):
        c.number_format = C.NUMFMT["dscr"]
    _base_head(ws["G23"])
    _base_head(ws["G34"])
    C.set_height(ws, 31, C.H_GAP)
    C.set_height(ws, 42, C.H_GAP)

    _drop_cf(ws, "D24:J30", "D35:J41", "D46:H46", "D47:H51", "D49:H49")
    # Cashflow: Blau-Sequenz, rote Schrift nur für Zuschuss (< 0 €)
    _heat(ws, "D24:J30", [("AND(ISNUMBER(D24),D24<0)", "red")])
    # DSCR: Blau-Sequenz, Schrift rot unter „prüfen“-Schwelle, amber bis „erfüllt“, darüber neutral
    _heat(ws, "D35:J41", [("AND(ISNUMBER(D35),D35<Ampel_DSCR_gelb)", "red")])
    _base_box(ws["G27"])
    _base_box(ws["G38"])

    # Legenden rechts im Unterabschnittskopf; Schwellen live aus den Ampel-Namen der Konfiguration (P21)
    _legend_formula(ws, "J21", "Rote Schrift = Zuschuss (Cashflow < 0 €)   ·   □ aktuelle Annahme")
    _legend_formula(ws, "J32", '="Rote Schrift = kritisch (< "&FIXED(Ampel_DSCR_gelb,2)&"×)   ·   "&'
                    + C.threshold_text("DSCR") + '&"   ·   □ aktuelle Annahme"')

    # ---- Abschnitt 3: IRR-Matrix und Verkauf (C:J, P2-04)
    _unmerge(ws, 43, 43, "C", "J")
    _clear(ws, "H", 43, "J", 43)
    C.safe_merge(ws, "C", 43, "G", 43)
    c43 = ws["C43"]                         # reine Anzeigeformel (Bandtitel)
    if C.is_formula(c43.value) and "Haltedauer" in c43.value:
        c43.value = '="Eigenkapitalrendite bei Verkauf nach "&Haltedauer&" Jahren – Miete × Wertsteigerung"'
    C.section(ws, 43, "C", "H")             # P3-14: Blockkopf so breit wie IRR-Matrix und Verkaufstabelle (I:J Weißraum)
    _uebersicht(ws, 43, "H")
    _clear(ws, "I", 44, "J", 63, values=False)
    for row in (44, 53):                    # P30: Unterabschnitte und Tabellen enden an der letzten Datenspalte H
        _clear(ws, "D", row, "J", row, values=False)
        C.section(ws, row, "C", "H", level=2, variant="line")
    _clear(ws, "C", 45, "J", 45)
    C.set_height(ws, 45, 6)
    _matrix(ws, 46, 47, 51, "H", "Miete ↓  ·  Wertsteigerung p. a. →", 49)
    for c in C.iter_cells(ws, "D", 46, "H", 46):
        c.number_format = HEAD_PCT
    for c in C.iter_cells(ws, "D", 47, "H", 51):
        c.number_format = C.NUMFMT["pct1"]
    # aktuelle Wertsteigerungsannahme: Kopf GOLD_BG fett Navy (keine Navy-Vollfläche mehr)
    ws.conditional_formatting.add("D46:H46", FormulaRule(
        formula=["ABS(D46-Wertsteigerung)<0.00005"], stopIfTrue=True,
        font=Font(color=C.NAVY, bold=True), fill=C.fill(BASE_BG)))
    # Basiszelle (Basis-Miete × aktuelle Wertsteigerung): Navy-Rahmen + fett, Schrift in Statusfarbe – alles in EINEM
    # Bereich D47:H51 (LibreOffice wertet je Zelle nur einen Bereich aus); Basis-Zeile = Zeilenwert 0 in Spalte C
    nb = C.side("medium", C.NAVY)
    box = Border(left=nb, right=nb, top=nb, bottom=nb)
    base = "ISNUMBER(D47),$C47=0,ABS(D$46-Wertsteigerung)<0.00005"   # flaches AND (LibreOffice: kein AND in AND)
    for cond, color in (("D47<Ampel_IRR_gelb", C.RED), ("TRUE", C.INK)):
        ws.conditional_formatting.add("D47:H51", FormulaRule(
            formula=[f"AND({base},{cond})"], stopIfTrue=True, font=Font(bold=True, color=color), border=box))
    _heat(ws, "D47:H51", [("AND(ISNUMBER(D47),D47<Ampel_IRR_gelb)", "red")])
    _legend_formula(ws, "H44", '="Rote Schrift = kritisch (< "&FIXED(Ampel_IRR_gelb*100,1)&" %)   ·   "&'
                    + C.threshold_text("IRR") + '&"   ·   □ aktuelle Annahme"')
    C.set_height(ws, 52, C.H_GAP)

    # Verkaufstabelle (Kopf und Linien bis H – P30)
    C.section(ws, 54, "C", "H", None, level=2, variant="fill", caps=False)
    head = ws["C54"]
    C.set_text(head, "Wertsteigerung p. a.  →")
    head.font = C.font(C.T_LABEL, True, C.BLUE)
    head.alignment = C.align("right", "center", 1)
    for c in C.iter_cells(ws, "D", 54, "H", 54):
        c.alignment = C.align("right", "center", 1)
        c.number_format = HEAD_PCT
    for r in range(55, 64):
        _clear(ws, "C", r, "J", r, values=False)
        lab = ws[f"C{r}"]
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", 1)
        for c in C.iter_cells(ws, "D", r, "H", r):
            c.font = C.font(C.T_BODY, False, C.INK)
            c.alignment = C.align("right", "center", 1)
        C.hairline(ws, r, "C", "H")
        C.set_height(ws, r, ROW_H)
    C.sum_row(ws, 59, "C", "H", "sub")
    C.sum_row(ws, 62, "C", "H", "final", value_from="D")
    # P15: negative Ergebnis- und Kumulwerte der Herleitung rot (u. a. kumulierter Cashflow D60:H60)
    for r in (55, 56, 57, 58, 60, 61):
        C.neg_red(ws, f"D{r}:H{r}")
    for c in C.iter_cells(ws, "D", 63, "H", 63):
        c.number_format = C.NUMFMT["mult2"]

    # ---- Rail: je Abschnitt genau eine Erklär-Box, Kopf in der Zeile des Abschnittskopfs (P2-04)
    # P30: Die Boxen schließen unten mit ihrem Abschnitt ab (Z. 18 bzw. 41); die IRR-Box endet nach ihrem Text,
    # darunter folgt das Diagramm bis zum Tabellenende (Z. 63).
    width = C.span_px(ws, "L", "M")
    last_body = 0
    for (head_row, title, text), fixed_end in zip(RAIL, (18, 41, None)):
        need = C.px_pt(C.lines_needed_metric(text, width, C.T_SMALL, False, 1) * C.line_pt(C.T_SMALL) + 12)
        r, acc = head_row + 1, 0
        while acc < need:
            acc += ws.row_dimensions[r].height or 15
            r += 1
        end = max(fixed_end or 0, r - 1)
        C.callout_box(ws, "L", head_row, "M", head_row + 1, end, title=title, text=text, pill=False, fit=None)
        C.set_height(ws, head_row, C.H_BAND)       # Kopf so hoch wie der Abschnittskopf daneben
        last_body = end

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
    # P3-14: ein neutraler Hinweis ohne Zeilennummern. Excel blendet Gliederungen auf geschützten Blättern nicht ein –
    # deshalb der Hinweis auf den (kennwortlosen) Blattschutz.
    C.set_text(note, "Die vollständige Nebenrechnung ist ausgeblendet. Nach „Blattschutz aufheben“ (ohne Kennwort) "
                     "lässt sie sich über [+] am linken Rand einblenden.")
    note.font = C.font(C.T_SMALL, False, C.MUTED, italic=True)
    note.alignment = C.align("left", "center", 1)
    C.set_height(ws, 64, C.H_ROW)
    # ---- Buttonzeile am Seitenende (P22): „‹ Zurück: Finanzierung“ links, „Weiter: Bankgespräch ›“ rechts bündig M
    # P1-03: Zurück (C) und Weiter (G:J) gleich breit im Raster der Hauptspalte; „Tilgungsplan“ = Reitername
    _nav_row(ws, 117, ("C", "C", "‹  Zurück: Tilgungsplan", "Finanzierung"),
             ("G", "J", "Weiter: Bankgespräch  ›", "Bankgespräch"))
    for row in ws.iter_rows(min_row=1, max_row=64):
        for c in row:
            if c.number_format == "@" and C.is_formula(c.value):
                c.number_format = "General"
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
