"""Blatt „Dashboard“ – Gesamtbewertung der Investition auf einer Seite (reines Darstellungsblatt).

Das Blatt liest ausschließlich vorhandene Namen und Ergebniszellen der Vorlage (Projektion, Steuern,
Cockpit-Prüfhinweise, Ampel-Schwellen der Konfiguration) und verändert keine bestehende Berechnung.
Alle Tokens und Komponenten kommen aus core.py (einheitliche Sprache mit den übrigen Blättern).

Raster: A = Rand · sechs Spaltenpaare (B:C, E:F, H:I, K:L, N:O, Q:R) mit fünf Abstandsspalten
(D, G, J, M, P ≈ 12 px). Linkes Panel B:I, rechtes Panel K:R, volle Breite B:R.
Hilfsdaten (Wasserfall, Sortierung der Prüfhinweise) liegen in ausgeblendeten Spalten T:AJ.
"""
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break

import core as C

SHEET = "Dashboard"

# ------------------------------------------------------------------------------------------ Raster
PAIR_W, GAP_W, EDGE_W = 12.86, 1.7, 4.5          # 90 px · 12 px · 31 px
PAIRS = [("B", "C"), ("E", "F"), ("H", "I"), ("K", "L"), ("N", "O"), ("Q", "R")]
GAPS = ["D", "G", "J", "M", "P"]
YEAR_COLS = ["E", "F", "H", "I", "K", "L", "N", "O", "Q", "R"]
YEARS = [1, 2, 3, 5, 10, 15, 20, 25, 30, 40]
YEAR_INDENT = {c: (2 if i % 2 == 0 else 1) for i, c in enumerate(YEAR_COLS)}   # Rechtskanten im Abstand ≈ 96 px
FIRST_HELP, LAST_HELP = "T", "AK"                  # ausgeblendete Hilfsspalten
AXIS_COL, AXIS_ROW = "AK", 20                      # Jahresachse Vermögensentwicklung (jedes 5. Jahr beschriftet)
CHROME_COLS = 60                                   # Kopfleiste breit genug für alle Reiter (≥ 1480 px)

# ------------------------------------------------------------------------------------------ Zeilen
R_EYEBROW, R_H1, R_SUB = 5, 6, 7
R_LINKS = 9
R_VERDICT = 11                                     # 11 Label · 12 Status
R_TILE = 14                                        # 14 Label · 15 Wert · 16 Unterzeile
R_BAND1 = 18                                       # Kennzahlen-Check | Cashflow Jahr 1
R_CHECK_HEAD = 19
R_CHECK = 20                                       # 20–25
R_CHECK_NOTE = 26
R_BAND2 = 28                                       # Vermögensentwicklung (volle Breite)
R_CHART2 = (29, 40)
R_BAND3 = 42                                       # Kennzahlen im Zeitverlauf

# Wasserfall-Hilfstabelle (ausgeblendet): Zeilen 20–26
W_ROW = 20
W_COLS = dict(cat="T", val="U", start="V", end="W", base_p="X", base_n="Y", vis_p="Z", vis_n="AH", label="AI")
W_CARRIERS = ["AA", "AB", "AC", "AD", "AE", "AF", "AG"]
W_HEAD = "U29"                                     # Höhe der Beschriftungsträger (16 % der Spannweite)
HINT_KEY_COL, HINT_KEY_ROW = "AJ", 20              # Sortierschlüssel der 25 Prüfhinweise (⚠ zuerst)
N_HINTS = 10

DEDUCT = "B8C2CF"                                  # Abzüge im Wasserfall


def build(wb, T=None):
    """T (Farb-Tokens des Themas) wird nicht mehr benötigt – alle Farben kommen aus core.py."""
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET, index=1)
    ws.sheet_properties.tabColor = C.NAVY
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100

    _grid(ws)
    _chrome(ws)
    _header(ws)
    _link_row(ws)
    rows = _check_table(ws)
    _verdict(ws, rows)
    _tiles(ws)
    _waterfall(ws)
    _wealth_chart(ws, wb)
    last = _timeline(ws)
    last = _hints(ws, last + 2)
    _footer(ws, last + 2)
    C.hide_cols(ws, FIRST_HELP, LAST_HELP)
    _page(ws, last + 3)
    return ws


# ========================================================================================== Raster / Kopf
def _grid(ws):
    ws.column_dimensions["A"].width = EDGE_W
    for a, b in PAIRS:
        ws.column_dimensions[a].width = PAIR_W
        ws.column_dimensions[b].width = PAIR_W
    for g in GAPS:
        ws.column_dimensions[g].width = GAP_W
    ws.column_dimensions["S"].width = EDGE_W
    for r in range(4, 90):
        ws.row_dimensions[r].height = 15
    for r, h in ((4, 15), (8, 8), (10, 14), (13, 14), (17, 20), (27, 20), (41, 20)):
        C.set_height(ws, r, h)


def _chrome(ws):
    """Kopfleiste wie auf allen Blättern (Reiter setzt navigation.py als Formen darüber)."""
    for r, h in ((1, 6), (2, 33), (3, 3)):
        C.set_height(ws, r, h)
        for cc in range(1, CHROME_COLS + 1):
            c = ws.cell(r, cc)
            c.fill = C.fill(C.ACCENT if r == 3 else C.NAVY)
            c.border = Border()


def _header(ws):
    """Seitenkopf nach core.page_header: Brotkrume · H1 · Untertitel (Metadaten) · Objekt rechts."""
    kaufdatum = ('IFERROR(TEXT(DAY(Kaufdatum),"00")&"."&TEXT(MONTH(Kaufdatum),"00")&"."&YEAR(Kaufdatum),"–")')
    subtitle = ('="Kaufpreis "&FIXED(Kaufpreis,0)&" €  ·  Kauf am "&' + kaufdatum +
                '&"  ·  Haltedauer "&Haltedauer&" Jahre  ·  "&Bundesland')
    C.page_header(ws, "B", "R", "Auswertung", "Dashboard", subtitle,
                  context=("=Obj_Name", "=Obj_Adresse"), context_col="N")
    C.safe_merge(ws, "B", R_H1, "L", R_H1)
    C.safe_merge(ws, "B", R_SUB, "L", R_SUB)
    C.safe_merge(ws, "N", R_H1, "R", R_H1)
    C.safe_merge(ws, "N", R_SUB, "R", R_SUB)
    ws.cell(R_H1, 14).alignment = C.align("right", "bottom", 1)


def _link_row(ws):
    """Rücksprung und Vertiefung im Kachelraster; rechts die Rechtsform (Link zur Auswahl auf „Start“)."""
    C.button(ws, "B", R_LINKS, "C", "‹  Schritt 12: Ergebnis", "S12 Ergebnis", kind="secondary",
             tooltip="Zurück zum letzten Schritt des Leitfadens")
    for (a, b), (text, target, tip) in zip(PAIRS[1:5], [
            ("Cockpit  ›", "Cockpit", "Detailkennzahlen und Prüfhinweise"),
            ("Diagramme  ›", "Diagramme", "Alle Auswertungen als Diagramm"),
            ("Sensitivität  ›", "Sensitivität", "Was-wäre-wenn: Zins, Miete, Kaufpreis"),
            ("Bankgespräch  ›", "Bankgespräch", "Unterlagen für das Finanzierungsgespräch")]):
        C.button(ws, a, R_LINKS, b, text, target, kind="link", tooltip=tip)
    a, b = PAIRS[5]
    C.safe_merge(ws, a, R_LINKS, b, R_LINKS)
    c = ws[f"{a}{R_LINKS}"]
    c.value = f'="Rechtsform: "&{C.RECHTSFORM_SHORT}&"  ›"'
    C.style_range(ws, a, R_LINKS, b, R_LINKS, fil=C.fill(C.TINT), brd=Border())
    c.font = C.font(C.T_BODY, True, C.NAVY)
    c.alignment = C.align("center", "center")
    c.hyperlink = Hyperlink(ref=c.coordinate, location="'Start'!D23", display="Rechtsform",
                            tooltip="Rechtsform auf der Startseite ändern")


# ========================================================================================== Gesamtbewertung
def _verdict(ws, rows):
    r1, r2 = R_VERDICT, R_VERDICT + 1
    C.set_height(ws, r1, 16)
    C.set_height(ws, r2, 26)
    for c in C.iter_cells(ws, "B", r1, "R", r2):
        c.fill = C.fill(C.TINT_XL)
        c.border = Border()
    C.safe_merge(ws, "B", r1, "I", r1)
    lab = ws[f"B{r1}"]
    C.set_text(lab, "GESAMTBEWERTUNG")
    lab.font = C.font(C.T_MICRO, True, C.MUTED)
    lab.alignment = C.align("left", "bottom", 1)
    C.safe_merge(ws, "B", r2, "I", r2)
    st = ws[f"B{r2}"]
    status_rng = f"$H${R_CHECK}:$H${R_CHECK + 5}"
    st.value = (f'=IF(OR($H${rows["DSCR"]}="kritisch",$H${rows["CF"]}="kritisch"),"Liquidität kritisch",'
                f'IF(COUNTIF({status_rng},"kritisch")>0,"Rendite kritisch",'
                f'IF(COUNTIF({status_rng},"prüfen")>0,"Solide mit Prüfpunkten","Solide")))')
    st.number_format = '"●  "@'
    st.font = C.font(C.T_H2, True, C.INK, C.DISPLAY)
    st.alignment = C.align("left", "center", 1)
    C.safe_merge(ws, "K", r1, "R", r2)
    ex = ws[f"K{r1}"]
    ex.value = ('="DSCR (Jahr 1) "&FIXED(DSCR_J1,2)&IF(DSCR_J1<Ampel_DSCR_gruen," statt mindestens "," bei Ziel ")'
                '&FIXED(Ampel_DSCR_gruen,2)&"  ·  IRR n. St. "&FIXED(EK_IRR*100,1)&" % über "&Haltedauer&" Jahre"'
                '&CHAR(10)&"Cashflow n. St. "&FIXED(CF_nSt_Monat_J1,0)&" € / Monat im Jahr 1, ab Jahr 2 "'
                f'&FIXED({C.CF_YEAR2},0)&" € / Monat"')
    ex.font = C.font(C.T_SMALL, False, C.INK2)
    ex.alignment = C.align("left", "center", 1, wrap=True)
    b = f"$B${r2}"
    rules = ((f'ISNUMBER(SEARCH("kritisch",{b}))', C.RED, C.RED_BG, "E7B4AD"),
             (f'ISNUMBER(SEARCH("Prüfpunkten",{b}))', C.AMBER, C.AMBER_BG, "F1CFA5"),
             (f'{b}="Solide"', C.GREEN, C.GREEN_BG, "B5DCC4"))
    for cond, fg, bg, ln in rules:
        ws.conditional_formatting.add(f"B{r1}:R{r2}", FormulaRule(formula=[cond], fill=C.fill(bg), stopIfTrue=True))
    for cond, fg, bg, ln in rules:
        ws.conditional_formatting.add(f"B{r2}:I{r2}", FormulaRule(formula=[cond], font=Font(color=fg, bold=True),
                                                                  stopIfTrue=True))


# ========================================================================================== Kacheln
def _tiles(ws):
    tiles = [
        ("GI", "=Gesamtinvestition", '="Kaufpreis "&FIXED(Kaufpreis,0)&" € · NK "&FIXED(NK_Quote*100,1)&" %"'),
        ("EK", "=EK_Bedarf_gesamt", '="Quote "&FIXED(EK_Quote*100,1)&" % · Beleihung "&FIXED(Beleihung*100,1)&" %"'),
        ("CF", "=CF_nSt_Monat_J1", f'="ab Jahr 2: "&FIXED({C.CF_YEAR2},0)&" € / Monat"'),
        ("BMR", "=Bruttomietrendite", '="Faktor "&FIXED(Kaufpreisfaktor,1)&"× · netto "&FIXED(Nettomietrendite*100,1)&" %"'),
        ("DSCR", "=DSCR_J1", '="Bankmaßstab ≥ "&FIXED(Ampel_DSCR_gruen,2)'),
        ("IRR", "=EK_IRR", '="Verkauf nach "&Haltedauer&" Jahren · Multiple "&FIXED(EK_Multiple,2)&"×"'),
    ]
    for (a, b), (key, value, sub) in zip(PAIRS, tiles):
        C.kpi_tile(ws, a, b, R_TILE, R_TILE + 1, R_TILE + 2, value=value, sub=sub, kpi=key, gap_right=False)
    a, _ = PAIRS[2]
    ws.conditional_formatting.add(f"{a}{R_TILE + 2}:{PAIRS[2][1]}{R_TILE + 2}", FormulaRule(
        formula=[f"{C.CF_YEAR2}<0"], font=Font(color=C.RED)))


# ========================================================================================== Kennzahlen-Check
def _check_table(ws):
    C.band_l1(ws, R_BAND1, "B", "I", "Kennzahlen-Check", "Ist gegen Zielwert")
    C.table_head(ws, R_CHECK_HEAD, "B", "I", labels={
        "B": ("KENNZAHL", "left"), "E": ("IST", "right"), "F": ("ZIEL", "right"),
        "H": ("STATUS", "left"), "I": ("ERREICHUNG", "left")})
    ws[f"I{R_CHECK_HEAD}"].alignment = C.align("left", "center", 0)
    cf_status = f'IF(E{{r}}<0,"kritisch",IF({C.CF_YEAR2}<0,"prüfen","erfüllt"))'
    checks = [
        ("BMR", "Bruttomietrendite", "=Bruttomietrendite", "=Ampel_BMR_gruen", "Ampel_BMR_gelb", C.NUMFMT["pct1"]),
        ("NMR", "Nettomietrendite", "=Nettomietrendite", "=Ampel_NMR_gruen", "Ampel_NMR_gelb", C.NUMFMT["pct1"]),
        ("DSCR", "DSCR (Jahr 1)", "=DSCR_J1", "=Ampel_DSCR_gruen", "Ampel_DSCR_gelb", C.NUMFMT["dscr"]),
        ("CF", "Cashflow n. St. / Monat", "=CF_nSt_Monat_J1", 0, None, C.NUMFMT["eur_zero"]),
        ("EKR", "EK-Rendite Jahr 1", "=EKR_Tile", "=Ampel_EKR_gruen", "Ampel_EKR_gelb", C.NUMFMT["pct1"]),
        ("IRR", "IRR n. St.", "=EK_IRR", "=Ampel_IRR_gruen", "Ampel_IRR_gelb", C.NUMFMT["pct1"]),
    ]
    rows = {}
    for i, (key, label, ist, ziel, gelb, fmt) in enumerate(checks):
        r = R_CHECK + i
        rows[key] = r
        C.set_height(ws, r, 22)
        C.safe_merge(ws, "B", r, "C", r)
        lab = ws[f"B{r}"]
        C.set_text(lab, label)
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", 1)
        v = ws[f"E{r}"]
        v.value, v.number_format = ist, fmt
        v.font = C.font(C.T_BODY, True, C.INK)
        v.alignment = C.align("right", "center", 1)
        z = ws[f"F{r}"]
        z.value, z.number_format = ziel, fmt
        z.font = C.font(C.T_BODY, False, C.MUTED)
        z.alignment = C.align("right", "center", 1)
        s = ws[f"H{r}"]
        s.value = ("=" + cf_status.format(r=r)) if key == "CF" else \
            f'=IF(E{r}>=F{r},"erfüllt",IF(E{r}>={gelb},"prüfen","kritisch"))'
        s.number_format = '"●  "@'
        s.font = C.font(C.T_SMALL, True, C.INK)
        s.alignment = C.align("left", "center", 1)
        q = ws[f"I{r}"]
        q.value = (f"=IF(E{r}>=0,1,0)" if key == "CF" else
                   f"=IF(F{r}>0,MIN(MAX(0,E{r}/F{r}),1.25),IF(E{r}>=F{r},1,0))")
        q.number_format = ";;;"
        q.alignment = C.align("left", "center", 0)
        C.hairline(ws, r, "B", "I")
    last = R_CHECK + len(checks) - 1
    for word, colr in (("erfüllt", C.GREEN), ("prüfen", C.AMBER), ("kritisch", C.RED)):
        ws.conditional_formatting.add(f"H{R_CHECK}:H{last}", FormulaRule(
            formula=[f'$H{R_CHECK}="{word}"'], font=Font(color=colr, bold=True)))
    ws.conditional_formatting.add(f"I{R_CHECK}:I{last}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=1.25, color=C.BLUE,
        showValue=False, minLength=0, maxLength=100))
    # Fußnote + Link zu den Schwellen
    C.set_height(ws, R_CHECK_NOTE, 20)
    n = ws[f"B{R_CHECK_NOTE}"]
    C.set_text(n, "Balken = Erreichung des Zielwerts (voll = 125 %)")
    n.font = C.font(C.T_MICRO, False, C.MUTED)
    n.alignment = C.align("left", "center", 1)
    C.safe_merge(ws, "H", R_CHECK_NOTE, "I", R_CHECK_NOTE)
    lk = ws[f"H{R_CHECK_NOTE}"]
    C.text_link(lk, "Schwellen: Konfiguration  ›", "Konfiguration", "B74", size=C.T_MICRO, bold=False,
                tooltip="Ampel-Schwellwerte anpassen")
    lk.alignment = C.align("right", "center", 1)
    return rows


# ========================================================================================== Diagramme
def _txpr(size=C.T_MICRO, color=C.MUTED, bold=False, rot=None):
    cp = CharacterProperties(sz=int(size * 100), b=bold, solidFill=color)
    cp.latin = None
    body = RichTextProperties(rot=rot, vert="horz") if rot is not None else RichTextProperties()
    return RichText(bodyPr=body, p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _no_fill(gp=None):
    gp = gp or GraphicalProperties()
    gp.noFill = True
    gp.line = LineProperties()
    gp.line.noFill = True
    return gp


def _anchor(chart, c1, r1, c2, r2):
    """Diagramm füllt genau sein Panel: from = erste Panelzelle, to = Zelle nach dem Panel (Offsets 0)."""
    chart.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=C.col(c1) - 1, row=r1 - 1, colOff=0, rowOff=0),
        to=AnchorMarker(col=C.col(c2), row=r2, colOff=0, rowOff=0))


def _waterfall(ws):
    """Cashflow Jahr 1 als Wasserfall aus gestapelten Säulen (finish_pro macht daraus 3D).

    Excel stapelt positive und negative Werte getrennt. Jede Säule [unten, oben] wird deshalb in einen
    positiven Teil (unsichtbare Basis + sichtbares Stück) und einen negativen Teil (unsichtbare Basis +
    sichtbares Stück) zerlegt; so bleibt sie auch beim Kreuzen der Nulllinie ein zusammenhängender Balken
    auf der echten Werteachse. Direkt über jeder Säule liegt ein unsichtbarer Träger, dessen Beschriftung den
    Betrag zeigt (Reihenname = Zellbezug, gebietsschema-sicher per FIXED).
    """
    C.band_l1(ws, R_BAND1, "K", "R", "Cashflow Jahr 1", "€ je Monat")
    cats = [("Miete", "=Miete_Ist_J1/12", False, C.NAVY),
            ("Bewirt-\nschaftung", "=-BWK_J1/12", True, DEDUCT),
            ("Zinsen", "=-Zins_J1/12", True, DEDUCT),
            ("Tilgung", "=-Tilgung_J1/12", True, DEDUCT),
            ("= vor Steuern", "=CF_vSt_J1/12", False, C.BLUE),
            ("Steuer", "=-Steuer_J1/12", True, C.ACCENT),
            ("= nach Steuern", "=CF_nSt_J1/12", False, C.NAVY)]
    K = W_COLS
    n = len(cats)
    r_first, r_last = W_ROW, W_ROW + n - 1
    se = f"${K['start']}${r_first}:${K['end']}${r_last}"
    h = _abs_ref(W_HEAD)
    for i, (cat, frm, chain, _) in enumerate(cats):
        r = W_ROW + i
        C.set_text(ws[f"{K['cat']}{r}"], cat)
        ws[f"{K['val']}{r}"] = frm
        ws[f"{K['start']}{r}"] = f"={K['end']}{r - 1}" if chain else 0
        ws[f"{K['end']}{r}"] = f"={K['start']}{r}+{K['val']}{r}"
        lo = f"MIN({K['start']}{r},{K['end']}{r})"
        hi = f"MAX({K['start']}{r},{K['end']}{r})"
        ws[f"{K['base_p']}{r}"] = f"=IF({hi}>=0,MAX({lo},0),0)"
        ws[f"{K['base_n']}{r}"] = f"=IF({hi}>=0,0,IF({hi}+{h}>=0,{hi},{hi}+{h}))"
        ws[f"{K['vis_p']}{r}"] = f"=IF({hi}>=0,{hi}-MAX({lo},0),0)"
        ws[f"{K['vis_n']}{r}"] = f"=IF({hi}>=0,MIN({lo},0),{lo}-{hi})"
        for j, cc in enumerate(W_CARRIERS):
            ws[f"{cc}{r}"] = f"=IF({hi}+{h}>=0,{h},-{h})" if i == j else 0
        ws[f"{K['label']}{r}"] = (f'=IF({K["val"]}{r}>=0.5,"+",IF({K["val"]}{r}<=-0.5,"−",""))'
                                  f'&FIXED(ABS({K["val"]}{r}),0)')
    ws[W_HEAD] = f"=MAX(1,0.16*(MAX({se},0)-MIN({se},0)))"

    bar = BarChart()
    bar.type = "col"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.gapWidth = 60
    bar.legend = None
    bar.visible_cells_only = False               # Daten liegen in ausgeblendeten Spalten
    order = ["base_p", "base_n", "vis_p"] + ["car"] * n + ["vis_n"]
    cols = [K["base_p"], K["base_n"], K["vis_p"]] + W_CARRIERS + [K["vis_n"]]
    for cc in cols:
        bar.add_data(Reference(ws, min_col=C.col(cc), min_row=r_first, max_row=r_last), titles_from_data=False)
    bar.set_categories(Reference(ws, min_col=C.col(K["cat"]), min_row=r_first, max_row=r_last))
    names = {"base_p": "Basis +", "base_n": "Basis −", "vis_p": "Betrag +", "vis_n": "Betrag −"}
    car = 0
    for s, role in zip(bar.series, order):
        if role in ("base_p", "base_n"):
            s.tx = SeriesLabel(v=names[role])
            s.graphicalProperties = _no_fill()
        elif role in ("vis_p", "vis_n"):
            s.tx = SeriesLabel(v=names[role])
            s.graphicalProperties = _solid(C.NAVY)
            for idx, (_, _, _, colr) in enumerate(cats):
                pt = DataPoint(idx=idx)
                pt.graphicalProperties = _solid(colr)
                s.dPt.append(pt)
        else:
            r = W_ROW + car
            s.tx = SeriesLabel(strRef=StrRef(f=f"'{SHEET}'!${K['label']}${r}"))
            s.graphicalProperties = _no_fill()
            dl = DataLabel(idx=car, showLegendKey=False, showVal=False, showCatName=False, showSerName=True,
                           showPercent=False, showBubbleSize=False)
            dl.txPr = _txpr(C.T_MICRO, C.INK2, True)
            s.dLbls = DataLabelList(dLbl=[dl], showLegendKey=False, showVal=False, showCatName=False,
                                    showSerName=False, showPercent=False, showBubbleSize=False)
            car += 1
    bar.y_axis.delete = False
    bar.y_axis.numFmt = '#,##0" €"'
    bar.y_axis.majorTickMark = "none"
    bar.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=C.LINE, w=6350)))
    bar.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    bar.y_axis.txPr = _txpr(C.T_MICRO, C.MUTED)
    bar.x_axis.delete = False
    bar.x_axis.tickLblPos = "low"
    bar.x_axis.majorTickMark = "none"
    bar.x_axis.tickLblSkip = 1
    bar.x_axis.txPr = _txpr(C.T_MICRO, C.MUTED, False, rot=0)
    bar.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=C.MUTED2, w=9525))
    _anchor(bar, "K", R_CHECK_HEAD, "R", R_CHECK_NOTE)
    ws.add_chart(bar)


def _solid(colr):
    gp = GraphicalProperties(solidFill=colr)
    gp.line = LineProperties()
    gp.line.noFill = True
    return gp


def _abs_ref(ref):
    import re
    return re.sub(r"^([A-Z]+)(\d+)$", r"$\1$\2", ref)


def _wealth_chart(ws, wb):
    C.band_l1(ws, R_BAND2, "B", "R", "Vermögensentwicklung", "€ · jeweils Jahresende · 40 Jahre")
    for r in range(R_CHART2[0], R_CHART2[1] + 1):
        C.set_height(ws, r, 20)
    P = wb["Projektion"]
    line = LineChart()
    spec = [(41, "Immobilienwert", C.ACCENT, 22225, None),
            (42, "Restschuld", C.MUTED2, 19050, "dash"),
            (43, "Nettovermögen", C.NAVY, 34925, None)]
    for row, title, colr, w, dash in spec:
        line.add_data(Reference(P, min_col=4, max_col=43, min_row=row), from_rows=True, titles_from_data=False)
        s = line.series[-1]
        s.tx = SeriesLabel(v=title)
        s.smooth = False
        s.marker = Marker(symbol="none")
        s.graphicalProperties = GraphicalProperties()
        s.graphicalProperties.line = LineProperties(solidFill=colr, w=w, prstDash=dash)
    for k in range(40):
        ref = f"Projektion!{C.L(4 + k)}$10"
        ws[f"{AXIS_COL}{AXIS_ROW + k}"] = f'=IF(MOD({k},5)=0,{ref}&"","")'
    line.set_categories(Reference(ws, min_col=C.col(AXIS_COL), min_row=AXIS_ROW, max_row=AXIS_ROW + 39))
    line.y_axis.delete = False
    line.y_axis.numFmt = '#,##0," T€"'
    line.y_axis.majorTickMark = "none"
    line.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=C.LINE, w=6350)))
    line.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    line.y_axis.txPr = _txpr(C.T_MICRO, C.MUTED)
    line.x_axis.delete = False
    line.x_axis.tickLblSkip = 5
    line.x_axis.tickMarkSkip = 5
    line.x_axis.majorTickMark = "out"
    line.x_axis.txPr = _txpr(C.T_MICRO, C.MUTED, rot=0)
    line.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=C.LINE2, w=6350))
    line.legend.position = "t"
    line.legend.txPr = _txpr(C.T_MICRO, C.MUTED)
    line.legend.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.62, y=0.0, w=0.38, h=0.09))
    line.plot_area.layout = Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge",
                                                             x=0.055, y=0.11, w=0.93, h=0.76))
    _anchor(line, "B", R_CHART2[0], "R", R_CHART2[1])
    ws.add_chart(line)


# ========================================================================================== Zeitverlauf
def _timeline(ws):
    C.band_l1(ws, R_BAND3, "B", "R", "Kennzahlen im Zeitverlauf", "€ p. a.  ·  Jahr 1 = 12 Monate ab Kaufdatum")
    C.set_height(ws, R_BAND3 + 1, 6)
    hr = R_BAND3 + 2
    C.table_head(ws, hr, "B", "R", labels={"B": ("JAHR", "left")})
    for y, cc in zip(YEARS, YEAR_COLS):
        c = ws[f"{cc}{hr}"]
        c.value = y
        c.number_format = "0"
        c.alignment = C.align("right", "center", YEAR_INDENT[cc])
    groups = [
        ("CASHFLOW", [
            ("row", "Nettokaltmiete Ist", "Projektion", 15, 1, None, False),
            ("row", "Bewirtschaftungskosten", "Projektion", 26, -1, None, False),
            ("row", "Kapitaldienst", "Projektion", 32, -1, None, False),
            ("davon", "davon Zinsen", "Projektion", 30, -1, None, False),
            ("davon", "davon Tilgung", "Projektion", 31, -1, None, False),
            ("sum1", "= Cashflow vor Steuern", "Projektion", 33, 1, None, True),
            ("row", "± Steuerwirkung", "Projektion", 35, -1, '+#,##0;-#,##0;"–"', False),
            ("sum2", "= Cashflow nach Steuern", "Projektion", 36, 1, None, True),
            ("row", "Kumulierter Cashflow n. St.", "Projektion", 38, 1, None, True)]),
        ("STEUER", [
            ("row", "Abschreibungen (AfA)", "Steuern", 56, 1, None, False),
            ("row", "Steuerliches Ergebnis", "Steuern", 67, 1, None, True)]),
        ("VERMÖGEN", [
            ("row", "Immobilienwert (Jahresende)", "Projektion", 41, 1, None, False),
            ("row", "Restschuld", "Projektion", 42, -1, None, False),
            ("sum2", "= Nettovermögen", "Projektion", 43, 1, None, True),
            ("pct", "EK-Rendite", "Projektion", 45, 1, C.NUMFMT["pct1"], True)]),
    ]
    r = hr
    for title, items in groups:
        r += 1
        C.subhead_l2(ws, r, "B", "R", title, height=22)
        ws.cell(r, 2).alignment = C.align("left", "bottom", 1)
        for kind, label, sheet, prow, sign, fmt, neg_red in items:
            r += 1
            C.set_height(ws, r, 18 if kind != "davon" else 16)
            C.safe_merge(ws, "B", r, "D", r)
            lab = ws[f"B{r}"]
            C.set_text(lab, label)
            for cc in YEAR_COLS:
                c = ws[f"{cc}{r}"]
                c.value = (f"={'-' if sign < 0 else ''}INDEX({_q(sheet)}!$D${prow}:$AQ${prow},{cc}${hr})")
                c.number_format = fmt or C.NUMFMT["num"]
                c.alignment = C.align("right", "center", YEAR_INDENT[cc])
            C.hairline(ws, r, "B", "R")
            if kind == "davon":
                C.style_range(ws, "B", r, "R", r, fnt=C.font(C.T_SMALL, False, C.MUTED))
                lab.alignment = C.align("left", "center", 2)
            else:
                C.style_range(ws, "B", r, "R", r, fnt=C.font(C.T_BODY, False, C.INK))
                lab.alignment = C.align("left", "center", 1)
            if kind == "sum1":
                C.total(ws, r, "B", "R", level=1)
            elif kind == "sum2":
                C.total(ws, r, "B", "R", level=2)
            if neg_red:
                ws.conditional_formatting.add(f"E{r}:R{r}", FormulaRule(
                    formula=[f"AND(ISNUMBER(E{r}),E{r}<0)"], font=Font(color=C.RED)))
    r += 1
    C.set_height(ws, r, 20)
    note = ws[f"B{r}"]
    C.set_text(note, "Vorzeichen: Einnahmen und Vermögen +, Ausgaben und Schulden –  ·  Steuerwirkung: + Erstattung, "
                     "– Zahlung  ·  EK-Rendite = Vermögenszuwachs / Eigenkapital")
    note.font = C.font(C.T_MICRO, False, C.MUTED)
    note.alignment = C.align("left", "center", 1)
    C.safe_merge(ws, "N", r, "R", r)
    lk = ws[f"N{r}"]
    C.text_link(lk, "Alle 40 Jahre: Projektion  ›", "Projektion", size=C.T_MICRO, bold=False,
                tooltip="Vollständige Jahresrechnung")
    lk.alignment = C.align("right", "center", 1)
    return r


def _q(sheet):
    return f"'{sheet}'" if " " in sheet or "-" in sheet else sheet


# ========================================================================================== Prüfhinweise
def _hints(ws, band_row):
    """Bis zu 10 aktive Hinweise aus Cockpit!AB50:AB74, Warnungen (⚠) zuerst – über eine eigene Sortierspalte."""
    key_rng = f"${HINT_KEY_COL}${HINT_KEY_ROW}:${HINT_KEY_COL}${HINT_KEY_ROW + 24}"
    for i in range(25):
        src = f"Cockpit!$AB${50 + i}"
        ws[f"{HINT_KEY_COL}{HINT_KEY_ROW + i}"] = f'=IF({src}="","",IF(LEFT({src},1)="⚠",0,100)+{i + 1})'
    warn = 'COUNTIF(Cockpit!$AB$50:$AB$74,"⚠*")'
    total = f"COUNT({key_rng})"
    C.band_l1(ws, band_row, "B", "R", "Prüfhinweise & steuerliche Einordnung", "")
    head = ws[f"R{band_row}"]
    head.value = (f'=IF({warn}=0,"Keine Warnungen",IF({warn}=1,"1 Warnung",{warn}&" Warnungen"))&"  ·  "&'
                  f'IF({total}-{warn}=0,"keine Hinweise",IF({total}-{warn}=1,"1 Hinweis",({total}-{warn})&" Hinweise"))')
    C.set_height(ws, band_row + 1, 6)
    first = band_row + 2
    for k in range(N_HINTS):
        r = first + k
        C.set_height(ws, r, 26)
        C.safe_merge(ws, "B", r, "R", r)
        c = ws[f"B{r}"]
        pick = f'IFERROR(INDEX(Cockpit!$AB$50:$AB$74,MATCH(SMALL({key_rng},{k + 1}),{key_rng},0)),"")'
        if k == 0:
            c.value = f'=IF({total}=0,"✓  Keine Prüfhinweise – alle Plausibilitätsprüfungen ohne Befund.",{pick})'
        else:
            c.value = "=" + pick
        c.font = C.font(C.T_SMALL, False, C.INK2)
        c.alignment = C.align("left", "center", 1, wrap=True)
    last = first + N_HINTS - 1
    rng = f"B{first}:R{last}"
    b = f"$B{first}"
    edge = Border(bottom=C.side("hair", C.LINE2))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'LEFT({b},1)="⚠"'], font=Font(color=C.RED), fill=C.fill(C.RED_BG), border=edge, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{b}<>""'], fill=C.fill(C.TINT_XL), border=edge, stopIfTrue=True))
    return last


# ========================================================================================== Fuß / Druck
def _footer(ws, row):
    C.set_height(ws, row - 1, 16)
    C.footer(ws, row, "B", "R")
    for r in (row, row + 1):
        C.safe_merge(ws, "B", r, "L", r)
    st = ws[f"R{row}"]
    st.value = "=TODAY()"
    st.number_format = '"Stand "DD.MM.YYYY'
    st.font = C.font(C.T_MICRO, False, C.MUTED)
    st.alignment = C.align("right", "center", 1)


def _page(ws, last_row):
    ws.print_area = f"A1:S{last_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.row_breaks.append(Break(id=R_BAND3 - 1))
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    ws.oddHeader.left.text = None
    ws.oddFooter.left.text = "MM Holding GmbH · Immobilien-Kalkulation · &A"
    ws.oddFooter.right.text = "Seite &P von &N"
    for part in (ws.oddFooter.left, ws.oddFooter.right):
        part.size = 8
        part.font = "Calibri,Regular"
        part.color = C.MUTED
    ws.freeze_panes = "A4"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
