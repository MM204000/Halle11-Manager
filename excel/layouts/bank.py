"""Bankgespräch (druckfertiges A4-Dokument), Haushaltsrechnung, Vermögensaufstellung.

Stil: ruhiges „Private Banking“ – viel Weißraum, feine Linien, eine Text- und eine Wertkante je Block.
Nur Darstellung: Formeln der Vorlage bleiben unverändert, geändert werden ausschließlich reine Anzeigeformeln
(Kopfzeilen, Fußnoten), statische Beschriftungen, Stile, Zahlenformate, Höhen/Breiten und Diagrammanker.
"""
import datetime
import re

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.datavalidation import DataValidation

import core as C

BANK, HH, VA = "Bankgespräch", "Haushaltsrechnung", "Vermögensaufstellung"


# =============================================================================== lokale Helfer
def _unmerge(ws, r1, r2, c1=1, c2=200):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= r2 and mr.max_row >= r1 and mr.min_col <= C.col(c2) and mr.max_col >= C.col(c1):
            ws.unmerge_cells(str(mr))


def _frame(ws, c1, r1, c2, r2, style="dashed", color=C.LINE2):
    """Rahmen nur außen um einen (verbundenen) Bereich, ohne Füllung."""
    ln = C.side(style, color)
    for r in range(r1, r2 + 1):
        for cc in range(C.col(c1), C.col(c2) + 1):
            cell = ws.cell(r, cc)
            cell.fill = C.NOFILL
            cell.border = Border(left=ln if cc == C.col(c1) else None, right=ln if cc == C.col(c2) else None,
                                 top=ln if r == r1 else None, bottom=ln if r == r2 else None)


def _input_box(ws, c1, row, c2, h="left"):
    """Eingabefeld über mehrere Spalten: gelb, Rahmen nur außen, Einzug 1."""
    ln = C.side("thin", C.INPUT_LINE)
    for cc in range(C.col(c1), C.col(c2) + 1):
        cell = ws.cell(row, cc)
        C.input_style(cell, "required")
        cell.border = Border(left=ln if cc == C.col(c1) else None, right=ln if cc == C.col(c2) else None, top=ln,
                             bottom=ln)
    first = ws.cell(row, C.col(c1))
    first.font = C.font(C.T_BODY, True, C.INPUT_FG)
    first.alignment = C.align(h, "center", 1)


def _anchor(chart, c1, r1, c2, r2):
    """Diagramm exakt auf die Zellgrenzen c1/r1 … Ende c2/r2 (ohne Versatz)."""
    a = chart.anchor
    a._from.col, a._from.colOff, a._from.row, a._from.rowOff = C.col(c1) - 1, 0, r1 - 1, 0
    a.to.col, a.to.colOff, a.to.row, a.to.rowOff = C.col(c2), 0, r2, 0


def _label(cell, indent=1, size=C.T_BODY, bold=False, color=C.INK, wrap=False):
    cell.font = C.font(size, bold, color)
    cell.alignment = C.align("left", "center", indent, wrap=wrap)


def _value(cell, fmt=None, size=C.T_BODY, bold=False, color=C.INK, h="right", indent=1):
    cell.font = C.font(size, bold, color)
    cell.alignment = C.align(h, "center", indent)
    if fmt:
        cell.number_format = fmt


def _hair(ws, row, c1, c2):
    for c in C.iter_cells(ws, c1, row, c2, row):
        c.fill = C.NOFILL
        c.border = Border(bottom=C.side("hair", C.LINE))


def _fmt_sub(formula, old, new):
    return formula.replace(old, new) if isinstance(formula, str) else formula


def _link(ws, coord, text, target, h="right", size=C.T_SMALL, v="center"):
    cell = ws[coord]
    C.text_link(cell, text, target, size=size, bold=True)
    cell.alignment = C.align(h, v, 0)


# =============================================================================== Bankgespräch
BANK_PANELS = (("B", "C"), ("E", "F"), ("H", "I"))
BANK_BULLETS = [
    "• Drucken Sie diese Übersicht (A4, Hochformat) und legen Sie sie der Finanzierungsanfrage bei. Alle Zahlen "
    "kommen aus der Kalkulation – Änderungen im Leitfaden oder auf dem Blatt „Eingaben“ aktualisieren die "
    "Übersicht automatisch.",
    "• Für die Bonitätsprüfung erwarten Banken eine Selbstauskunft. Die Vorlagen „Haushaltsrechnung“ und "
    "„Vermögensaufstellung“ sind vorbereitet und fließen in den Block „Sicherheiten & Bonität“ ein.",
    "• Argumentieren Sie mit Kapitaldienstdeckung (DSCR), Eigenkapitalquote, Beleihungsauslauf und dem Cashflow "
    "nach Steuern – das sind die Kennzahlen, auf die Kreditentscheider zuerst schauen.",
    "• Fotos des Objekts (Außenansicht, Wohnräume, Lageplan) erhöhen die Aussagekraft: über Einfügen › Bilder "
    "einfügen und in einen der drei gestrichelten Rahmen oben ziehen.",
]
BANK_LABELS = {
    "H20": "Nettokaltmiete Soll pro Monat",
    "H25": "Nettokaltmiete Ist (nach Ausfall)",
    "H30": "= Cashflow nach Steuern ¹",
    "E36": "Zu versteuerndes Einkommen",
    "E37": "Haushaltsüberschuss p. a.",
    "E38": "Nettovermögen",
}
BANK_NOTE = ("¹ Einschließlich Steuereffekt des Objekts (Steuererstattung bzw. -zahlung).   Alle Werte stammen aus "
             "der Kalkulation (Blätter Eingaben, Finanzierung, Projektion, Steuern). Prognosewerte sind Annahmen und "
             "keine Zusicherung. Haushaltsrechnung und Vermögensaufstellung liegen als separate Vorlagen bei.")
TEXT_VALUES = ("C20", "C21", "F21", "F22", "F26", "C37", "C40", "F36")
PCT1 = ("I22", "C36", "C38", "C39", "C41", "F29", "F30")


def bank(ws):
    # ---- Raster: drei Spalten Beschriftung + Wert, schmale Stege (Druck A4 hoch)
    widths = {"A": 4.5, "B": 27, "C": 18, "D": 1.5, "E": 27, "F": 18, "G": 1.5, "H": 28, "I": 18, "J": 2.5,
              "K": 48}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    ws.sheet_view.showGridLines = False

    # ---- Briefkopf: Titel links auf Weiß, Stand rechts, Akzentlinie; Objekt / Erstellt für; Metazeile
    _unmerge(ws, 5, 7, "B", "I")
    for c in C.iter_cells(ws, "B", 5, "I", 7):
        c.fill = C.NOFILL
        c.border = Border()
    title = ws["B5"]
    C.set_text(title, "Investitionsübersicht für das Bankgespräch")
    C.safe_merge(ws, "B", 5, "I", 5)
    title.font = C.font(C.T_KPI, True, C.NAVY, C.DISPLAY)
    title.alignment = C.align("left", "bottom")
    for c in C.iter_cells(ws, "B", 5, "I", 5):
        c.border = Border(bottom=C.side("medium", C.ACCENT))

    C.safe_merge(ws, "B", 6, "F", 6)
    ws["B6"].font = C.font(C.T_BODY, True, C.NAVY)
    ws["B6"].alignment = C.align("left", "center")
    made_for = ws["H6"]
    made_for.value = '=IFERROR(IF(Erstellt_fuer="","","Erstellt für "&Erstellt_fuer),"")'
    C.safe_merge(ws, "H", 6, "I", 6)
    made_for.font = C.font(C.T_BODY, True, C.NAVY)
    made_for.alignment = C.align("right", "center")

    meta = ws["B7"]
    meta.value = ('="Kaufdatum (geplant): "&TEXT(DAY(Kaufdatum),"00")&"."&TEXT(MONTH(Kaufdatum),"00")&"."'
                  '&YEAR(Kaufdatum)&"  ·  Erwerb als: "&CHOOSE(Rechtsform_Idx,"Privatperson",'
                  '"Vermögensverwaltende GmbH","GmbH")&"  ·  Objektart: "&Objektart&"  ·  Baujahr "&Baujahr')
    C.safe_merge(ws, "B", 7, "F", 7)
    meta.font = C.font(C.T_SMALL, False, C.MUTED)
    meta.alignment = C.align("left", "center")
    stand = ws["H7"]                       # rechts oben liegt die Formen-Unterleiste → Stand unter „Erstellt für“
    stand.value = "=TODAY()"
    stand.number_format = '"Stand: "DD.MM.YYYY'
    C.safe_merge(ws, "H", 7, "I", 7)
    stand.font = C.font(C.T_SMALL, False, C.MUTED)
    stand.alignment = C.align("right", "center")
    for r, h in {4: 12, 5: 34, 6: 22, 7: 16, 8: 12}.items():
        C.set_height(ws, r, h)

    # ---- Foto-Rahmen: nur gestrichelte Linie, ohne Anweisungstext (der steht im Hinweis-Panel)
    for c1, c2 in BANK_PANELS:
        ws[f"{c1}9"].value = None
        _frame(ws, c1, 9, c2, 15)
    for r in range(9, 16):
        C.set_height(ws, r, 10.5)
    C.set_height(ws, 16, 12)

    # ---- Abschnitte (L1) und Unterköpfe (L2)
    for row in (17, 32):
        for c1, c2 in BANK_PANELS:
            C.band_l1(ws, row, c1, c2)
        C.set_height(ws, row + 1, 6)
    for row, panels in ((19, BANK_PANELS), (24, BANK_PANELS), (27, [("E", "F")])):
        for c1, c2 in panels:
            C.subhead_l2(ws, row, c1, c2, height=C.H_ROW)

    for coord, text in BANK_LABELS.items():
        C.set_text(ws[coord], text)
    ws["F36"].value = _fmt_sub(ws["F36"].value, '"GmbH – Jahresabschluss"', '"lt. Jahresabschluss"')
    ws["C40"].value = _fmt_sub(ws["C40"].value, '" J. / "', '" Jahre / "')

    # ---- Datenzeilen: eine Text- und eine Wertkante je Panel, Haarlinie, alles 10 pt 1A1D21
    heads = {(19, "B"), (19, "E"), (19, "H"), (24, "B"), (24, "E"), (24, "H"), (27, "E")}
    for r in list(range(20, 31)) + list(range(34, 42)):
        C.set_height(ws, r, C.H_ROW)
        for c1, c2 in BANK_PANELS:
            if (r, c1) in heads:
                continue
            lab, val = ws[f"{c1}{r}"], ws[f"{c2}{r}"]
            if lab.value is None and val.value is None:
                continue
            _hair(ws, r, c1, c2)
            _label(lab)
            if val.coordinate != lab.coordinate and not (r == 41 and c1 == "E"):
                _value(val)
    for r in (19, 24):
        C.set_height(ws, r, C.H_ROW)
    # Beleihungsobjekt (Text über E:F; reine Anzeigeformel, kürzeres Präfix, damit auch längere Adressen passen)
    ws["E41"].value = _fmt_sub(ws["E41"].value, '"Beleihungsobjekt: "', '"Objekt: "')
    _label(ws["E41"])

    for coord in PCT1:
        ws[coord].number_format = C.NUMFMT["pct1"]
    ws["C35"].number_format = C.NUMFMT["dscr"]
    ws["I23"].number_format = C.NUMFMT["mult1"]
    ws["I21"].number_format = C.NUMFMT["eur2"]

    # Summen: Gesamtinvestition / Eigenkapital (Blockergebnis), CF v. St. (Zwischensumme), CF n. St. (Schlüssel)
    C.total(ws, 29, "B", "C", level=2)
    C.total(ws, 28, "E", "F", level=2)
    C.total(ws, 29, "H", "I", level=1)
    C.total(ws, 30, "H", "I", level=3)
    for coord in ("B29", "E28", "H29", "H30"):
        ws[coord].alignment = C.align("left", "center", 1)
    for coord in ("C29", "F28", "I29", "I30"):
        ws[coord].alignment = C.align("right", "center", 1)

    # Ampel wie auf Dashboard/Cockpit (dieselben Schwellen aus der Konfiguration)
    for coord, kpi in (("C35", "DSCR"), ("C36", "NMR"), ("C39", "IRR"), ("C41", "EKR"), ("I22", "BMR"),
                       ("I29", "CFV"), ("I30", "CF")):
        C.add_ampel(ws, coord, kpi)

    # ---- Diagrammzeile: Band, Kreis B:C, Säulen E:I (Zeilen 44–57)
    C.band_l1(ws, 43, "B", "I")
    C.set_height(ws, 42, 12)
    for r in range(44, 58):
        C.set_height(ws, r, 15)
    charts = sorted(ws._charts, key=lambda ch: ch.anchor._from.col)
    if len(charts) >= 2:
        _anchor(charts[0], "B", 44, "C", 57)
        _anchor(charts[1], "E", 44, "I", 57)
    for r, h in {58: 10, 59: 4, 60: 4}.items():
        C.set_height(ws, r, h)

    # ---- Annahmen und Hinweise (Fußnote ¹ zum Cashflow nach Steuern)
    b61 = ws["B61"]
    if isinstance(b61.value, str):
        b61.value = b61.value.replace(' | ', '  ·  ')
    C.set_text(ws["B63"], BANK_NOTE)
    for coord, (r1, r2) in (("B61", (61, 62)), ("B63", (63, 64))):
        cell = ws[coord]
        cell.font = C.font(C.T_SMALL, False, C.MUTED)
        cell.alignment = C.align("left", "top", 0, wrap=True)
        n = C.lines_needed(cell.value if coord == "B63" else "x" * 330, C.span_px(ws, "B", "I"), C.T_SMALL)
        per = max(12, (n * 11.5 + 4) / 2)
        C.set_height(ws, r1, per)
        C.set_height(ws, r2, per)

    # ---- Hinweis-Panel rechts (nur Bildschirm, außerhalb des Druckbereichs)
    _unmerge(ws, 5, 30, "K", "K")
    for r in range(5, 31):
        cell = ws[f"K{r}"]
        cell.value = None
        cell.fill = C.NOFILL
        cell.border = Border()
        cell.hyperlink = None
    body = "\n\n".join(BANK_BULLETS)
    # Kasten endet bündig mit dem oberen Tabellenblock (Z. 30); nur bei sehr langem Text tiefer
    need = C.lines_needed(body, C.col_px(ws, "K"), C.T_SMALL) * 11.5 + 8
    r, acc = 18, 0
    while (acc < need or r <= 30) and r < 42:
        acc += ws.row_dimensions[r].height or 15
        r += 1
    last = r - 1
    C.callout(ws, "K", 17, "K", 18, last, title="Bessere Konditionen im Bankgespräch")
    C.set_height(ws, 17, 24)
    C.set_height(ws, 18, 6)
    ws["K18"].value = body
    ws["K18"].alignment = C.align("left", "top", 1, wrap=True)
    for i, (text, target) in enumerate((("Haushaltsrechnung ausfüllen  ›", HH),
                                        ("Vermögensaufstellung ausfüllen  ›", VA),
                                        ("Alle Diagramme  ›", "Diagramme"))):
        cell = ws[f"K{max(last + 2, 34) + i}"]
        C.text_link(cell, text, target, size=C.T_BODY, bold=True)
        cell.alignment = C.align("left", "center", 1)


# =============================================================================== Haushaltsrechnung / Vermögensaufstellung
FORM_WIDTHS = {"A": 4.5, "B": 56, "C": 14, "D": 14, "E": 36, "F": 3}
HH_LABELS = {
    "B42": "Bewirtschaftung des kalkulierten Objekts (Jahr 1, inkl. Rücklage)",
}


def _form_common(ws, first, last, comment_rows):
    """Spaltenraster, eine Textkante (Einzug 1), Werte rechts mit Einzug 1, Kommentare 9 pt grau."""
    for k, w in FORM_WIDTHS.items():
        ws.column_dimensions[k].width = w
    ws.sheet_view.showGridLines = False
    bands = set()
    for r in range(first, last + 1):
        b = ws.cell(r, 2)
        fill = b.fill.fgColor.rgb[-6:] if b.fill is not None and b.fill.fill_type == "solid" else None
        if fill in (C.TINT, C.HEAD) and isinstance(b.value, str) and b.value.upper() == b.value:
            bands.add(r)
            continue
        if isinstance(b.value, str) and not C.is_formula(b.value):
            b.alignment = C.align("left", "center", 1)
        for cc in (3, 4):
            v = ws.cell(r, cc)
            if v.value is None:
                continue
            if v.alignment.horizontal == "right" or isinstance(v.value, (int, float)) or C.is_formula(v.value):
                v.alignment = C.align("right", "center", 1)
    for r in comment_rows:
        if r in bands:
            continue
        e = ws.cell(r, 5)
        if e.value is not None:
            e.font = C.font(C.T_SMALL, False, C.MUTED)
            e.alignment = C.align("left", "center", 1)
    return bands


def _subnav(ws, back, nxt):
    """Zell-Links als Rückfall zur Formen-Unterleiste (Z. 5): weiter (Z. 6) und zurück (Z. 7), rechtsbündig in E."""
    _link(ws, "E6", f"{nxt[0]}  ›", nxt[1], size=C.T_BODY)
    _link(ws, "E7", f"‹  {back[0]}", back[1], v="top")


def household(ws):
    for coord, text in HH_LABELS.items():
        C.set_text(ws[coord], text)
    _form_common(ws, 8, 47, range(18, 48))

    # Persönliche Angaben: alle Felder über C:E (volle Breite), Kinder als Zahl, „Beschäftigt seit“ als echtes Datum
    _unmerge(ws, 9, 14, "C", "E")
    kids = ws["C12"]
    if isinstance(kids.value, str) and kids.value.strip().isdigit():
        kids.value = int(kids.value.strip())
    since = ws["C14"]
    if isinstance(since.value, str) and re.match(r"^\d{2}\.\d{2}\.\d{4}$", since.value.strip()):
        d, m, y = (int(x) for x in since.value.strip().split("."))
        since.value = datetime.datetime(y, m, d)
    for r in range(9, 15):
        for cc in ("D", "E"):
            ws[f"{cc}{r}"].value = None
        C.safe_merge(ws, "C", r, "E", r)
        _input_box(ws, "C", r, "E")
    kids.number_format = "0"
    since.number_format = C.NUMFMT["date"]
    dv_k = DataValidation(type="whole", operator="between", formula1="0", formula2="20", allow_blank=True,
                          showInputMessage=True, showErrorMessage=True)
    dv_k.promptTitle, dv_k.prompt = "Kinder", "Anzahl unterhaltspflichtiger Kinder (0–20)."
    dv_k.errorTitle, dv_k.error = "Ungültige Anzahl", "Bitte eine ganze Zahl zwischen 0 und 20 eingeben."
    dv_d = DataValidation(type="date", operator="between", formula1="DATE(1950,1,1)", formula2="TODAY()",
                          allow_blank=True, showInputMessage=True, showErrorMessage=True)
    dv_d.promptTitle, dv_d.prompt = "Beschäftigt seit", "Datum im Format TT.MM.JJJJ, z. B. 01.01.2015."
    dv_d.errorTitle, dv_d.error = "Ungültiges Datum", "Bitte ein Datum zwischen 1950 und heute eingeben."
    for dv, ref in ((dv_k, "C12"), (dv_d, "C14")):
        ws.add_data_validation(dv)
        dv.add(ref)

    # Kommentar zum Überschuss kürzen, als Kommentar (9 pt grau) auch in der Ergebniszeile
    C.set_text(ws["E46"], "Banken erwarten > 0 nach neuem Kapitaldienst")
    ws["E46"].font = C.font(C.T_SMALL, False, C.MUTED)
    ws["E46"].alignment = C.align("left", "center", 1)
    ws["C47"].number_format = C.NUMFMT["pct1"]

    # Status des Überschusses zusätzlich knapp (Quote < 10 %) in Amber
    ws.conditional_formatting.add("C47", FormulaRule(formula=["AND(ISNUMBER($C$47),$C$47<0)"], stopIfTrue=True,
                                                     font=Font(color=C.RED, bold=True)))
    ws.conditional_formatting.add("C47", FormulaRule(formula=["AND(ISNUMBER($C$47),$C$47<0.1)"], stopIfTrue=True,
                                                     font=Font(color=C.AMBER, bold=True)))

    for ch in ws._charts:
        _anchor(ch, "G", 8, "M", 27)
    _subnav(ws, ("Bankgespräch", BANK), ("Weiter: Vermögensaufstellung", VA))


def assets(ws):
    _form_common(ws, 8, 34, range(10, 35))
    # Verpfändet? – Auswahlliste ja/nein, „ja“ in Amber (belastete Werte fallen auf)
    dv = DataValidation(type="list", formula1='"ja,nein"', allow_blank=True, showInputMessage=True,
                        showErrorMessage=True)
    dv.promptTitle, dv.prompt = "Verpfändet?", "„ja“, wenn der Vermögenswert bereits als Sicherheit dient."
    dv.errorTitle, dv.error = "Bitte auswählen", "Nur „ja“ oder „nein“ aus der Liste."
    ws.add_data_validation(dv)
    dv.add("D10:D18")
    for r in range(10, 19):
        ws.cell(r, 4).alignment = C.align("center", "center")
    ws.conditional_formatting.add("D10:D18", FormulaRule(formula=['LOWER(TRIM(D10))="ja"'], stopIfTrue=True,
                                                         font=Font(color=C.AMBER, bold=True)))
    ws["D9"].alignment = C.align("center", "center")
    for ch in ws._charts:
        _anchor(ch, "G", 8, "M", 19)
    _subnav(ws, ("Haushaltsrechnung", HH), ("Weiter: Bankgespräch", BANK))


def apply(wb):
    if BANK in wb.sheetnames:
        bank(wb[BANK])
    if HH in wb.sheetnames:
        household(wb[HH])
    if VA in wb.sheetnames:
        assets(wb[VA])
