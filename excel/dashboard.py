"""Dashboard-Blatt (Management-Übersicht auf einer Seite) für das Immobilien-Kalkulationstool Pro.

Das Blatt liest ausschließlich vorhandene Namen und Ergebniszellen der Vorlage
(Projektion, Cockpit-Hinweise, Ampel-Schwellen der Konfiguration) – es verändert
keine bestehende Berechnung. Farben kommen aus dem aktiven Thema von design_pro.
"""
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

EUR = '#,##0 "€";-#,##0 "€";"–"'
NUM = '#,##0;-#,##0;"–"'
PCT = '0.0 %'
PCT2 = '0.00 %'
DSCR = '0.00'


def build(wb, T):
    """T: Farb-Token (dict) des aktiven Themas."""
    P, P2, ACC, L, XL = T["TEAL"], T["TEAL_MID"], T["ACC"], T["TEAL_L"], T["TEAL_XL"]
    INK, INK2, MUTED, MUTED2, LINE, LINE2 = "1A1D21", "3A3F45", "5B6068", "8A9099", "E6E8EB", "D5D9DE"
    RED, AMB, GRN = "B42318", "B54708", "1F7A4D"
    RED_BG, AMB_BG, GRN_BG = "FDF1EF", "FEF4E8", "EDF7F1"
    SANS, DISP = "Aptos", "Aptos Display"

    ws = wb.create_sheet("Dashboard", index=1)
    ws.sheet_properties.tabColor = P
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100

    def f(size=10, bold=False, color=INK, name=SANS):
        return Font(name=name, sz=size, b=bold, color=color)

    def fill(c):
        return PatternFill("solid", start_color=c, end_color=c)

    def side(style, color):
        return Side(style=style, color=color)

    def put(ref, value=None, fmt=None, font=None, bg=None, al=None, border=None):
        c = ws[ref]
        if value is not None:
            c.value = value
        c.font = font or f()
        if fmt:
            c.number_format = fmt
        if bg:
            c.fill = fill(bg)
        if al:
            c.alignment = al
        if border:
            c.border = border
        return c

    def merge(rng, value=None, **kw):
        ws.merge_cells(rng)
        return put(rng.split(":")[0], value, **kw)

    def section(row, c1, c2, title, right=None):
        ws.row_dimensions[row].height = 22
        for col in range(c1, c2 + 1):
            ws.cell(row, col).border = Border(bottom=side("medium", P))
        put(f"{get_column_letter(c1)}{row}", title, font=f(12, True, P, DISP), al=Alignment(vertical="bottom"))
        if right:
            put(f"{get_column_letter(c2)}{row}", right, font=f(8.5, color=MUTED2),
                al=Alignment(horizontal="right", vertical="bottom"))

    LEFT = Alignment(horizontal="left", vertical="center")
    RIGHT = Alignment(horizontal="right", vertical="center")

    # ------------------------------------------------------------------ Raster
    ws.column_dimensions["A"].width = 2.4
    for i in range(2, 14):
        ws.column_dimensions[get_column_letter(i)].width = 12.9
    ws.column_dimensions["N"].width = 2.4
    for col in "OP":
        ws.column_dimensions[col].width = 14
    for r in range(4, 66):
        ws.row_dimensions[r].height = 16.5

    # ------------------------------------------------------------------ Kopfleiste
    for r, h in ((1, 6), (2, 27), (3, 2.25)):
        ws.row_dimensions[r].height = h
        for col in range(1, 15):
            ws.cell(r, col).fill = fill(ACC if r == 3 else P)
    merge("B2:D2", "MM HOLDING  ·  IMMOBILIEN-KALKULATION", font=f(8.5, True, T["ON_DARK_ACC"]),
          al=Alignment(horizontal="left", vertical="center", indent=1))
    merge("E2:H2", "Auswertung  ›  Dashboard", font=f(10, True, "FFFFFF"), al=Alignment(vertical="center", indent=1))
    for ref, text, target in (("L2", "‹  Start", "Start"), ("M2", "Cockpit  ›", "Cockpit")):
        c = put(ref, text, font=f(9, color=T["ON_DARK_2"]), al=Alignment(horizontal="right", vertical="center"))
        c.hyperlink = Hyperlink(ref=ref, location=f"'{target}'!A1", display=text)

    # ------------------------------------------------------------------ Titel
    ws.row_dimensions[5].height = 34
    merge("B5:G5", "Dashboard", font=f(26, True, P, DISP), al=Alignment(vertical="center"))
    merge("B6:G6", '=Obj_Name&"   ·   "&Obj_Adresse', font=f(10, color=MUTED), al=LEFT)
    meta = [("Kaufpreis", "=Kaufpreis", EUR), ("Kaufdatum", "=Kaufdatum", "DD.MM.YYYY"),
            ("Haltedauer", "=Haltedauer", '0 "Jahre"'), ("Rechtsform", '=IF(Rechtsform_Idx=1,"Privat","GmbH")', "@"),
            ("Bundesland", "=Bundesland", "@"), ("Stand", "=TODAY()", "DD.MM.YYYY")]
    for i, (lab, frm, fmt) in enumerate(meta):
        col = get_column_letter(8 + i)
        put(f"{col}5", lab, font=f(8, color=MUTED2), al=Alignment(horizontal="left", vertical="bottom"))
        put(f"{col}6", frm, fmt=fmt, font=f(9, True), al=Alignment(horizontal="left", vertical="center", shrink_to_fit=True))

    # ------------------------------------------------------------------ Kennzahlen-Check (Grundlage der Bewertung)
    check_top = 17
    section(check_top - 1, 2, 7, "Kennzahlen-Check", "Ist gegen Zielwert")
    heads = [("B", "C", "KENNZAHL", "left"), ("D", None, "IST", "right"), ("E", None, "ZIEL", "right"),
             ("F", None, "STATUS", "left"), ("G", None, "ERREICHUNG", "left")]
    for a, b, text, al in heads:
        rng = f"{a}{check_top}:{b}{check_top}" if b else f"{a}{check_top}"
        if b:
            ws.merge_cells(rng)
        put(f"{a}{check_top}", text, font=f(8, True, MUTED), bg="F3F4F5",
            al=Alignment(horizontal=al, vertical="center", indent=1 if al == "left" else 0),
            border=Border(bottom=side("thin", LINE2)))
    ws["C17"].fill = fill("F3F4F5")
    checks = [
        ("Bruttomietrendite", "=Bruttomietrendite", "=Ampel_BMR_gruen", "Ampel_BMR_gelb", PCT),
        ("Nettomietrendite", "=Nettomietrendite", "=Ampel_NMR_gruen", "Ampel_NMR_gelb", PCT),
        ("Kapitaldienstdeckung (DSCR)", "=DSCR_J1", "=Ampel_DSCR_gruen", "Ampel_DSCR_gelb", DSCR),
        ("Cashflow n. St. / Monat", "=CF_nSt_Monat_J1", 0, "-100", '#,##0 "€";-#,##0 "€";0 "€"'),
        ("EK-Rendite Jahr 1", "=EKR_Tile", "=Ampel_EKR_gruen", "Ampel_EKR_gelb", PCT),
        ("IRR nach Steuern", "=EK_IRR", "=Ampel_IRR_gruen", "Ampel_IRR_gelb", PCT),
    ]
    rows = {}
    for i, (lab, ist, ziel, gelb, fmt) in enumerate(checks):
        r = check_top + 1 + i
        rows[lab] = r
        ws.row_dimensions[r].height = 19
        merge(f"B{r}:C{r}", lab, font=f(10), al=Alignment(horizontal="left", vertical="center", indent=1))
        put(f"D{r}", ist, fmt=fmt, font=f(10, True), al=RIGHT)
        put(f"E{r}", ziel, fmt=fmt, font=f(10, color=MUTED2), al=RIGHT)
        put(f"F{r}", f'=IF(D{r}>=E{r},"erfüllt",IF(D{r}>={gelb},"prüfen","kritisch"))', fmt='"● "@',
            font=f(9, True), al=Alignment(horizontal="left", vertical="center", indent=1))
        put(f"G{r}", f'=REPT("█",ROUND(8*MIN(IF(E{r}>0,MAX(0,D{r}/E{r}),IF(D{r}>=E{r},1.25,0)),1.25),0))',
            font=f(7.5), al=LEFT)
        for col in "BCDEFG":
            ws[f"{col}{r}"].border = Border(bottom=side("thin", LINE))
        for word, colr in (("erfüllt", GRN), ("prüfen", AMB), ("kritisch", RED)):
            ws.conditional_formatting.add(f"F{r}:G{r}", FormulaRule(formula=[f'$F{r}="{word}"'], font=Font(color=colr, bold=True)))
    last = check_top + len(checks)
    merge(f"B{last + 1}:G{last + 1}", "Balken: 8 Blöcke = Zielwert erreicht · Schwellen grün/gelb: Blatt Konfiguration "
          "(Cashflow: gelb ab −100 € / Monat)", font=f(8, color=MUTED2), al=LEFT)
    rng_status = f"F{check_top + 1}:F{last}"
    r_dscr, r_cf = rows["Kapitaldienstdeckung (DSCR)"], rows["Cashflow n. St. / Monat"]

    # ------------------------------------------------------------------ Gesamtbewertung
    for r in (8, 9):
        ws.row_dimensions[r].height = 18
    merge("B8:C9", "GESAMTBEWERTUNG", font=f(8.5, True, MUTED), al=Alignment(horizontal="left", vertical="center", indent=1))
    merge("D8:F9", f'=IF(OR(F{r_dscr}="kritisch",F{r_cf}="kritisch"),"Liquidität kritisch",'
                   f'IF(COUNTIF({rng_status},"kritisch")>0,"Rendite kritisch",'
                   f'IF(COUNTIF({rng_status},"prüfen")>0,"Solide mit Prüfpunkten","Solide")))',
          fmt='"● "@', font=f(14, True, INK, DISP), al=LEFT)
    merge("G8:M9", '="Kapitaldienstdeckung "&FIXED(DSCR_J1,2)&IF(DSCR_J1<Ampel_DSCR_gruen," statt mindestens "," bei Ziel ")'
                   '&FIXED(Ampel_DSCR_gruen,2)&". Cashflow nach Steuern im Jahr 1: "&FIXED(CF_nSt_Monat_J1,0)'
                   '&" € / Monat, ab Jahr 2: "&FIXED(INDEX(Projektion!$D$37:$AQ$37,2),0)&" € / Monat. IRR nach Steuern "'
                   '&FIXED(EK_IRR*100,1)&" % über "&Haltedauer&" Jahre."',
          font=f(9.5, color=INK2), al=Alignment(horizontal="left", vertical="center", wrap_text=True))
    for word, fg, bg, ln in (("kritisch", RED, RED_BG, "E7B4AD"), ("Prüfpunkten", AMB, AMB_BG, "F1CFA5"), ('Solide"', GRN, GRN_BG, "B5DCC4")):
        cond = f'ISNUMBER(SEARCH("{word}",$D$8))' if word != 'Solide"' else '$D$8="Solide"'
        ws.conditional_formatting.add("B8:M9", FormulaRule(formula=[cond], fill=fill(bg), border=Border(top=side("thin", ln), bottom=side("thin", ln))))
        ws.conditional_formatting.add("D8:F9", FormulaRule(formula=[cond], font=Font(color=fg, bold=True)))

    # ------------------------------------------------------------------ KPI-Kacheln
    tiles = [
        ("GESAMTINVESTITION", "=Gesamtinvestition", EUR, '="Kaufpreis "&FIXED(Kaufpreis,0)&" € · NK "&FIXED(NK_Quote*100,1)&" %"'),
        ("EIGENKAPITALBEDARF", "=EK_Bedarf_gesamt", EUR, '="Quote "&FIXED(EK_Quote*100,1)&" % · Beleihung "&FIXED(Beleihung*100,1)&" %"'),
        ("CASHFLOW N. ST. / MONAT", "=CF_nSt_Monat_J1", EUR, '="Jahr 2: "&FIXED(INDEX(Projektion!$D$37:$AQ$37,2),0)&" €"'),
        ("BRUTTOMIETRENDITE", "=Bruttomietrendite", PCT, '="Faktor "&FIXED(Kaufpreisfaktor,1)&"× · netto "&FIXED(Nettomietrendite*100,1)&" %"'),
        ("KAPITALDIENSTDECKUNG", "=DSCR_J1", DSCR, '="Bankmaßstab ≥ "&FIXED(Ampel_DSCR_gruen,2)'),
        ("IRR NACH STEUERN", "=EK_IRR", PCT, '=Haltedauer&" Jahre · Multiple "&FIXED(EK_Multiple,2)&"×"'),
    ]
    ws.row_dimensions[11].height = 18
    ws.row_dimensions[12].height = 19
    ws.row_dimensions[13].height = 19
    ws.row_dimensions[14].height = 17
    gap = side("thick", "FFFFFF")
    for i, (lab, frm, fmt, sub) in enumerate(tiles):
        a, b = get_column_letter(2 + 2 * i), get_column_letter(3 + 2 * i)
        for r in range(11, 15):
            for col in (a, b):
                c = ws[f"{col}{r}"]
                c.fill = fill(XL)
                c.border = Border(top=side("thick", P) if r == 11 else None, right=gap if col == b and i < 5 else None)
        merge(f"{a}11:{b}11", lab, font=f(8, True, MUTED), al=Alignment(horizontal="left", vertical="bottom", indent=1))
        merge(f"{a}12:{b}13", frm, fmt=fmt, font=f(22, True, INK, DISP), al=Alignment(horizontal="left", vertical="center", indent=1, shrink_to_fit=True))
        merge(f"{a}14:{b}14", sub, font=f(8, color=MUTED), al=Alignment(horizontal="left", vertical="top", indent=1, shrink_to_fit=True))
    ws.conditional_formatting.add("F12", FormulaRule(formula=["CF_nSt_Monat_J1<0"], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add("F12", FormulaRule(formula=["CF_nSt_Monat_J1>=0"], font=Font(color=GRN, bold=True)))
    ws.conditional_formatting.add("J12", FormulaRule(formula=["DSCR_J1<Ampel_DSCR_gelb"], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add("J12", FormulaRule(formula=["DSCR_J1<Ampel_DSCR_gruen"], font=Font(color=AMB, bold=True)))
    ws.conditional_formatting.add("L12", FormulaRule(formula=["EK_IRR<Ampel_IRR_gelb"], font=Font(color=RED, bold=True)))

    # ------------------------------------------------------------------ Diagrammdaten (sichtbar, dezent)
    section(16, 15, 16, "Diagrammdaten")
    data = [("Miete", "=Miete_Ist_J1/12"), ("Bewirt.", "=-BWK_J1/12"), ("Zinsen", "=-Zins_J1/12"),
            ("Tilgung", "=-Tilgung_J1/12"), ("vor Steuern", "=CF_vSt_J1/12"), ("Steuer", "=-Steuer_J1/12"),
            ("nach Steuern", "=CF_nSt_J1/12")]
    for i, (lab, frm) in enumerate(data):
        r = 17 + i
        put(f"O{r}", lab, font=f(8, color=MUTED2), al=LEFT)
        put(f"P{r}", frm, fmt=NUM, font=f(8, color=MUTED2), al=RIGHT)

    # ------------------------------------------------------------------ Cashflow Jahr 1 (Säulen)
    section(16, 9, 13, "Cashflow Jahr 1", "€ je Monat")
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.legend = None
    bar.gapWidth = 60
    bar.add_data(Reference(ws, min_col=16, min_row=17, max_row=23), titles_from_data=False)
    bar.set_categories(Reference(ws, min_col=15, min_row=17, max_row=23))
    s = bar.series[0]
    s.graphicalProperties.solidFill = P
    s.graphicalProperties.line.noFill = True
    for idx, colr in enumerate([P, "B8C2CF", "B8C2CF", "B8C2CF", RED, ACC, P]):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = colr
        pt.graphicalProperties.line.noFill = True
        s.dPt.append(pt)
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    for attr in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
        setattr(s.dLbls, attr, False)
    bar.y_axis.numFmt = "#,##0"
    bar.y_axis.delete = False
    bar.x_axis.delete = False
    bar.x_axis.tickLblPos = "low"
    bar.width, bar.height = 12.6, 5.3
    ws.add_chart(bar, "I17")

    # ------------------------------------------------------------------ Vermögensentwicklung (Linien)
    section(27, 2, 7, "Vermögensentwicklung", "€ · 40 Jahre")
    line = LineChart()
    P_ = wb["Projektion"]
    for row, colr, width, title in ((41, P2, 28575, "Immobilienwert"), (42, "8A9099", 22225, "Restschuld"),
                                   (43, P, 38100, "Nettovermögen")):
        line.add_data(Reference(P_, min_col=4, max_col=43, min_row=row), from_rows=True, titles_from_data=False)
        ser = line.series[-1]
        ser.tx = SeriesLabel(v=title)
        ser.graphicalProperties.line.solidFill = colr
        ser.graphicalProperties.line.width = width
        ser.smooth = True
        if title == "Restschuld":
            ser.graphicalProperties.line.dashStyle = "dash"
    line.set_categories(Reference(P_, min_col=4, max_col=43, min_row=10))
    line.y_axis.numFmt = '#,##0,"T€"'
    line.y_axis.delete = False
    line.x_axis.delete = False
    line.x_axis.tickLblSkip = 5
    line.x_axis.tickMarkSkip = 5
    line.legend.position = "b"
    line.width, line.height = 15.4, 7.3
    ws.add_chart(line, "B28")

    # ------------------------------------------------------------------ Prüfhinweise
    section(27, 9, 13, "Prüfhinweise")
    put("M27", '=COUNTIF(I28:I39,"⚠*")&" Warnung(en)"', font=f(8.5, color=MUTED2), al=Alignment(horizontal="right", vertical="bottom"))
    for i in range(6):
        r = 28 + 2 * i
        merge(f"I{r}:M{r + 1}", f'=IF(Cockpit!B{50 + i}="","",Cockpit!B{50 + i})', font=f(9, color=INK2),
              al=Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1))
        for col in "IJKLM":
            ws[f"{col}{r + 1}"].border = Border(bottom=side("thin", LINE))
        ws.conditional_formatting.add(f"I{r}", FormulaRule(formula=[f'LEFT(I{r},1)="⚠"'], font=Font(color=RED, bold=True)))

    # ------------------------------------------------------------------ Kennzahlen im Zeitverlauf
    top = 42
    section(top, 2, 13, "Kennzahlen im Zeitverlauf", "€ p. a. · Jahr 1 = 12 Monate ab Kaufdatum")
    hr = top + 1
    ws.row_dimensions[hr].height = 19
    merge(f"B{hr}:C{hr}", "JAHR", font=f(8.5, True, "FFFFFF"), bg=P, al=Alignment(horizontal="left", vertical="center", indent=1))
    ws[f"C{hr}"].fill = fill(P)
    years = [1, 2, 3, 5, 10, 15, 20, 25, 30, 40]
    for j, y in enumerate(years):
        col = get_column_letter(4 + j)
        put(f"{col}{hr}", y, fmt="0", font=f(8.5, True, "FFFFFF"), bg=P, al=RIGHT)
    series = [("row", "Nettokaltmiete Ist", 15, 1), ("row", "Bewirtschaftungskosten", 26, -1),
              ("row", "Kapitaldienst", 32, -1), ("sum", "Cashflow vor Steuern", 33, 1),
              ("row", "Steuereffekt (+ Erstattung)", 35, -1), ("res", "Cashflow nach Steuern", 36, 1),
              ("row", "Kumulierter Cashflow n. St.", 38, 1), ("row", "Immobilienwert", 41, 1),
              ("row", "Restschuld", 42, 1), ("sum", "Nettovermögen", 43, 1)]
    for i, (kind, lab, prow, sign) in enumerate(series):
        r = hr + 1 + i
        ws.row_dimensions[r].height = 18
        bold = kind != "row"
        colr = P if kind == "res" else INK
        bg = L if kind == "res" else (XL if kind == "sum" else None)
        merge(f"B{r}:C{r}", lab, font=f(9.5, bold, colr), bg=bg, al=Alignment(horizontal="left", vertical="center", indent=1))
        if bg:
            ws[f"C{r}"].fill = fill(bg)
        for j in range(len(years)):
            col = get_column_letter(4 + j)
            frm = f"={'-' if sign < 0 else ''}INDEX(Projektion!$D${prow}:$AQ${prow},{col}${hr})"
            put(f"{col}{r}", frm, fmt=NUM, font=f(9.5, bold, colr), bg=bg, al=RIGHT)
        for col in range(2, 14):
            c = ws.cell(r, col)
            c.border = Border(top=side("thin", P) if kind != "row" else None, bottom=side("thin", LINE))
        if kind in ("sum", "res"):
            ws.conditional_formatting.add(f"D{r}:M{r}", FormulaRule(formula=[f"D{r}<0"], font=Font(color=RED, bold=True)))
    note_r = hr + len(series) + 1
    merge(f"B{note_r}:M{note_r}", "Quelle: Blatt Projektion · Vorzeichen: Einnahmen positiv, Ausgaben negativ · "
          "Alle Werte aktualisieren sich automatisch aus den Eingaben.", font=f(8, color=MUTED2), al=LEFT)

    # ------------------------------------------------------------------ Fußzeile
    fr = note_r + 2
    for col in range(2, 14):
        ws.cell(fr, col).border = Border(top=side("thin", LINE2))
    put(f"B{fr}", "Keine Gewähr. Ersetzt keine Rechts-, Steuer- oder Finanzberatung · Rechtsstand September 2026",
        font=f(8, color=MUTED2), al=LEFT)
    put(f"M{fr}", "MM Holding GmbH · Kornhausgasse 4 · 88250 Weingarten · HRB 750729", font=f(8, color=MUTED2), al=RIGHT)

    # ------------------------------------------------------------------ Druck & Schutz
    ws.print_area = f"A1:M{fr}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    ws.oddFooter.left.text = "MM Holding GmbH · Immobilien-Kalkulation · &A"
    ws.oddFooter.right.text = "Seite &P von &N"
    ws.freeze_panes = "A4"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    return ws
