"""Blatt „Dashboard“ – Gesamtbewertung der Investition auf einer Seite (reines Darstellungsblatt).

Das Blatt liest ausschließlich vorhandene Namen und Ergebniszellen der Vorlage (Projektion, Steuern,
Cockpit-Prüfhinweise, Ampel-Schwellen der Konfiguration) und verändert keine bestehende Berechnung.
Alle Komponenten kommen aus core.py (Runde 2: EINE Komponentensprache):
  Seitenkopf C.page_header · Link-Reihe C.btn(soft/chip) · Kacheln C.tile (Wert neutral, Status-Chip) ·
  Abschnittsköpfe C.section · Statusspalte C.status_cf · Summenzeilen C.sum_row · Rot nur C.neg_red.

Raster: A = Rand · sechs Spaltenpaare (B:C, E:F, H:I, K:L, N:O, Q:R) mit fünf Abstandsspalten
(D, G, J, M, P = 12 px). Jedes Paar ist 180 px breit, aufgeteilt 84 | 96 px: So liegen die rechten Kanten
der zehn Jahresspalten der Zeitverlaufstabelle in exakt gleichen Abständen (96 px, P3-01), während
Kacheln, Buttons und Panels das 180-px-Raster behalten. Linkes Panel B:I, rechtes Panel K:R.
Hilfsdaten (Wasserfall, Sortierung der Prüfhinweise, Verkaufsmarke) liegen in ausgeblendeten Spalten T:AK.
"""
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import LegendEntry
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.pagebreak import Break, RowBreak

import core as C

SHEET = "Dashboard"

# ------------------------------------------------------------------------------------------ Raster
W_A, W_B, GAP_W, EDGE_W = 12.0, 13.71, 1.7, 4.5   # 84 px · 96 px · 12 px · 31 px (Paar = 180 px)
PAIRS = [("B", "C"), ("E", "F"), ("H", "I"), ("K", "L"), ("N", "O"), ("Q", "R")]
GAPS = ["D", "G", "J", "M", "P"]
YEAR_COLS = ["E", "F", "H", "I", "K", "L", "N", "O", "Q", "R"]
YEARS = [1, 2, 3, 5, 10, 15, 20, 25, 30, 40]
FIRST_HELP, LAST_HELP = "T", "AL"                  # ausgeblendete Hilfsspalten
CHROME_COLS = 19                                   # Kopfleiste A:S – endet mit der Randspalte S an der Inhaltskante (P01)

# ------------------------------------------------------------------------------------------ Zeilen
R_EYEBROW, R_H1, R_SUB = 5, 6, 7
R_LINKS = 9
R_VERDICT = 11                                     # 11 Label · 12 Status (cockpit.py liest B{R_VERDICT+1})
R_TILE = 14                                        # 14 Label · 15 Wert · 16 Kontextzeile
R_BAND1 = 18                                       # Kennzahlen-Check | Cashflow Jahr 1
R_CHECK_HEAD = 19
R_CHECK = 20                                       # 20–25
R_CHECK_NOTE = 26
R_BAND2 = 28                                       # Vermögensentwicklung (volle Breite)
R_CHART2 = (29, 40)
R_BAND3 = 42                                       # Kennzahlen im Zeitverlauf
R_YEAR, R_CAL = 43, 44                             # Kopf: Projektjahr · Kalenderjahr

# Wasserfall-Hilfstabelle (ausgeblendet): Zeilen 20–26
W_ROW = 20
W_COLS = dict(cat="T", val="U", start="V", end="W", base_p="X", base_n="Y", vis_p="Z", vis_n="AH", label="AI")
W_CARRIERS = ["AA", "AB", "AC", "AD", "AE", "AF", "AG"]
W_HEAD = "U29"                                     # Höhe der Beschriftungsträger (16 % der Spannweite)
W_ALT_P, W_ALT_N = "U31", "U32"                    # Steuer als Abfluss: Zeilen 31/32 (sichtbare Teile, Spalte V:AB)
HINT_KEY_COL, HINT_KEY_ROW = "AJ", 20              # Sortierschlüssel der 25 Prüfhinweise (Warnungen zuerst)
MARK_COL, MARK_ROW = "AK", 20                      # Verkaufsmarke (Säulenhöhe im Verkaufsjahr, sonst #NV)
DOT_COL = "AL"                                     # Punkt auf der Nettovermögenslinie (Zeilen 20–59, #NV)
MARK_LABEL = "AK62"                                # Beschriftung der Verkaufsmarke (Reihenname)
N_HINTS = 6

INFLOW, OUTFLOW, TOTAL = C.ACCENT, "B8C3D1", C.NAVY   # Wasserfall-Semantik (P2-10)
BAR_NEUTRAL = "B8C3D1"                             # Erreichungsbalken: neutral (Status trägt der Chip)
MARK_FILL = "CEDDF0"                               # Verkaufsmarke: 9CBBE2 zu 50 % auf Weiß
WARN_SYM, INFO_SYM = "▲", "•"                      # monochrome Kennzeichnung (ⓘ fehlt in Calibri)


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
    C.cf_close(ws)
    return ws


# ========================================================================================== Raster / Kopf
def _grid(ws):
    ws.column_dimensions["A"].width = EDGE_W
    for a, b in PAIRS:
        ws.column_dimensions[a].width = W_A
        ws.column_dimensions[b].width = W_B
    for g in GAPS:
        ws.column_dimensions[g].width = GAP_W
    ws.column_dimensions["S"].width = EDGE_W
    for r in range(4, 90):
        ws.row_dimensions[r].height = 15
    for r, h in ((4, 15), (8, 8), (10, C.H_GAP), (13, C.H_GAP), (17, 20), (27, 20), (41, 20)):
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
    """Seitenkopf nach der Vorlage core.page_header (P04): Z. 5 Dachzeile = Kategorie (P37, nie der Blattname),
    Z. 6 Titel links · Objekt rechts (10 pt fett), Z. 7 Untertitel links · Adresse und „Erstellt für …“ rechts
    (9 pt grau) – wortgleich mit den Schrittseiten. Alle rechten Elemente bündig an der Inhaltskante R."""
    kaufdatum = 'IFERROR(TEXT(DAY(Kaufdatum),"00")&"."&TEXT(MONTH(Kaufdatum),"00")&"."&YEAR(Kaufdatum),"–")'
    subtitle = ('="Kaufpreis "&FIXED(Kaufpreis,0)&" €  ·  Kauf am "&' + kaufdatum +
                '&"  ·  Haltedauer "&Haltedauer&" Jahre  ·  "&Bundesland')
    context = ("=Obj_Name",
               '=Obj_Adresse&IFERROR(IF(Erstellt_fuer="","","  ·  Erstellt für "&Erstellt_fuer),"")')
    C.page_header(ws, "B", "R", "Auswertung  ·  Gesamtbewertung auf einer Seite", "Dashboard", subtitle,
                  context=context, context_col="K")
    C.safe_merge(ws, "B", R_H1, "I", R_H1)
    C.safe_merge(ws, "B", R_SUB, "I", R_SUB)
    for r in (R_H1, R_SUB):
        C.safe_merge(ws, "K", r, "R", r)
    ws.cell(R_H1, 2).alignment = C.align("left", "bottom")
    ws.cell(R_SUB, 2).alignment = C.align("left", "top")
    ws.cell(R_H1, 11).alignment = C.align("right", "bottom")
    ws.cell(R_SUB, 11).alignment = C.align("right", "top")


def _link_row(ws):
    """Sprungleiste (P42, Button-Regel P13): Zurück = Ghost-Link ohne Fläche (5B6068), Vorwärts = einheitliche
    Sekundär-Buttons, Sprung zu einer Eingabe = gelber Eingabe-Chip „KAUF ALS · Privat ▾“ (Auswahl auf Start D23)."""
    C.btn(ws, "B", R_LINKS, "C", "‹  Schritt 12: Ergebnis", "S12 Ergebnis", kind="back",
          tooltip="Zurück zum letzten Schritt des Leitfadens", set_row=False)
    items = [("Cockpit  ›", "Cockpit", "Detailkennzahlen und Prüfhinweise"),
             ("Diagramme  ›", "Diagramme", "Alle Auswertungen als Diagramm"),
             ("Sensitivität  ›", "Sensitivität", "Was-wäre-wenn: Zins, Miete, Kaufpreis"),
             ("Bankgespräch  ›", "Bankgespräch", "Unterlagen für das Finanzierungsgespräch")]
    for (a, b), (text, target, tip) in zip(PAIRS[1:5], items):
        C.btn(ws, a, R_LINKS, b, text, target, kind="secondary", tooltip=tip, set_row=False)
    a, b = PAIRS[5]
    chip = C.btn(ws, a, R_LINKS, b, "KAUF ALS", "Start", kind="input", target_cell="D23",
                 tooltip="Kauf als Privatperson oder Gesellschaft – Auswahl auf der Startseite ändern", set_row=False)
    chip.value = f'="KAUF ALS  ·  "&{C.RECHTSFORM_SHORT}&"  ▾"'
    chip.data_type = "f"
    C.set_height(ws, R_LINKS, C.H_BTN)


# ========================================================================================== Gesamtbewertung
def _verdict(ws, rows):
    """Gesamturteil als stärkstes Element – ohne Signalfläche (P41, C.status_banner): Band F3F7FC, linke 3-px-Kante
    in Statusfarbe, EINE Pill „● kritisch · 1 von 6 erfüllt“. Das Urteil (20 pt fett) trägt wie ein Kachelwert die
    Statusfarbe (P11), die Begründung steht rechts als kurze „–“-Aufzählung (9 pt).
    cockpit.py liest das Urteil aus B{R_VERDICT + 1} (Wörter „kritisch“ / „Prüfpunkten“ / „Solide“ bleiben)."""
    r1, r2 = R_VERDICT, R_VERDICT + 1
    C.set_height(ws, r1, C.H_ROW)
    C.set_height(ws, r2, C.H_ROW3)
    b = f"$B${r2}"
    conds = [(f'ISNUMBER(SEARCH("kritisch",{b}))', "red"),
             (f'ISNUMBER(SEARCH("Prüfpunkten",{b}))', "amber"),
             (f'{b}="Solide"', "green")]
    C.safe_merge(ws, "B", r1, "C", r1)
    lab = ws[f"B{r1}"]
    C.set_text(lab, "GESAMTBEWERTUNG")
    C.safe_merge(ws, "E", r1, "I", r1)
    status_rng = f"$H${R_CHECK}:$H${R_CHECK + 5}"
    # Urteil in Statusfarbe (wie ein Kachelwert, P11) – Regel VOR der Kante anlegen und ohne stopIfTrue,
    # sonst stoppt die Kantenregel (B12) die Schriftfarbe
    for cond, lvl in conds:
        C.cf_rule(ws, f"B{r2}:I{r2}", cond, font_=Font(color=C.STATUS_COLORS[lvl][0], bold=True), stop=False)
    C.status_banner(ws, "B", r1, "R", r2, conditions=conds)
    # Status rechts in der Label-Zeile – wie die Kachel-Anatomie: links Beschriftung, rechts „● Wort“ (Statusfarbe)
    pill = ws[f"E{r1}"]
    C.status_pill(ws, pill, conditions=conds, style="chip",
                  suffix=f'COUNTIF({status_rng},"erfüllt")&" von 6 Kennzahlen erfüllt"')
    pill.font = C.font(C.T_SMALL, True, C.MUTED)
    pill.alignment = C.align("right", "bottom", 1)
    lab.font = C.font(C.T_LABEL, True, C.BLUE)
    lab.alignment = C.align("left", "bottom", 1)
    C.safe_merge(ws, "B", r2, "I", r2)
    st = ws[f"B{r2}"]
    st.value = (f'=IF(OR($H${rows["DSCR"]}="kritisch",$H${rows["CF"]}="kritisch"),"Liquidität kritisch",'
                f'IF(COUNTIF({status_rng},"kritisch")>0,"Rendite kritisch",'
                f'IF(COUNTIF({status_rng},"prüfen")>0,"Solide mit Prüfpunkten","Solide")))')
    st.number_format = "General"
    st.font = C.font(C.T_KPI, True, C.NAVY, C.DISPLAY)
    st.alignment = C.align("left", "center", 1)
    # Begründung rechts: feine Trennlinie links (LINE2) statt zweiter Fläche
    C.safe_merge(ws, "K", r1, "R", r2)
    ex = ws[f"K{r1}"]
    ex.value = ('="–  DSCR (Jahr 1) "&FIXED(DSCR_J1,2)&"×"&IF(DSCR_J1<Ampel_DSCR_gruen," statt mindestens "," bei Ziel ")'
                '&FIXED(Ampel_DSCR_gruen,2)&"×"'
                '&CHAR(10)&"–  IRR n. St. "&FIXED(EK_IRR*100,1)&" % über "&Haltedauer&" Jahre (Ziel "'
                '&FIXED(Ampel_IRR_gruen*100,1)&" %)"'
                '&CHAR(10)&"–  Cashflow n. St. "&FIXED(CF_nSt_Monat_J1,0)&" € / Monat im Jahr 1, ab Jahr 2 "'
                f'&FIXED({C.CF_YEAR2},0)&" € / Monat"')
    ex.font = C.font(C.T_SMALL, False, C.INK2)
    ex.alignment = C.align("left", "center", 1, wrap=True)
    for r in (r1, r2):
        ws.cell(r, C.col("K")).border = Border(left=C.side("thin", C.LINE2))


# ========================================================================================== Kacheln
def _tiles(ws):
    """Sechs Kacheln in kanonischer Reihenfolge und mit kanonischen Labels (P20: C.KPI_ORDER / C.kpi_label) –
    EINE Anatomie (P11, C.tile dark): Kopfstreifen · Wert 20 pt (Wertfarbe = Status) · Fußzeile links Kontext,
    rechts Status-Chip. Nebenkennzahlen (Faktor, Rate, Multiple) stehen in der Fußzeile, nie in der Wertzeile;
    die Zielwerte zeigt der Kennzahlen-Check direkt darunter (Spalte „Ziel“)."""
    t = R_TILE, R_TILE + 1, R_TILE + 2
    subs = {
        "GI": '="Kaufpreis "&FIXED(Kaufpreis,0)&" €  ·  NK "&FIXED(NK_Quote*100,1)&" %"',
        "EK": '="Quote "&FIXED(EK_Quote*100,1)&" %  ·  Beleihung "&FIXED(Beleihung*100,1)&" %"',
        "CF": f'="Jahr 2: "&FIXED({C.CF_YEAR2},0)&" €"',
        "BMR": '="Faktor "&FIXED(Kaufpreisfaktor,1)&"×"',
        "DSCR": '="Rate "&FIXED(Kapitaldienst_J1/12,0)&" €"',
        "IRR": '="Multiple "&FIXED(EK_Multiple,2)&"×"',
    }
    values = {"GI": "=Gesamtinvestition", "EK": "=EK_Bedarf_gesamt", "CF": "=CF_nSt_Monat_J1",
              "BMR": "=Bruttomietrendite", "DSCR": "=DSCR_J1", "IRR": "=EK_IRR"}
    for (a, b), key in zip(PAIRS, C.KPI_ORDER):
        label = C.kpi_label(key, caps=True, formula=True) if key == "IRR" else C.kpi_label(key, caps=True)
        ref = values[key][1:] if C.KPI[key].get("rule") else None
        C.tile(ws, a, b, *t, kpi=key, label=label, value=values[key], value_ref=ref, sub=subs[key],
               gap_right=False)
        lab = ws[f"{a}{R_TILE}"]
        if C.is_formula(label):                # Label als Anzeigeformel (Haltedauer)
            lab.value = label
            lab.data_type = "f"


# ========================================================================================== Kennzahlen-Check
def _check_table(ws):
    """Kennzahlen-Check: kanonische Reihenfolge (P20: Cashflow, Rendite, DSCR, IRR), Ist neutral, Ziel als
    kanonische Schwellenformulierung „≥ …“ (C.threshold_text – für DSCR wie für alle anderen, P42),
    Status als Chip, Erreichungsbalken neutral und bei 100 % gekappt („voll = Ziel erreicht“, P42)."""
    C.section(ws, R_BAND1, "B", "I", "Kennzahlen-Check", meta="Ist gegen Zielwert")
    C.section(ws, R_CHECK_HEAD, "B", "I", None, level=2, labels={
        "B": ("Kennzahl", "left"), "E": ("Ist", "right"), "F": ("Ziel", "right"),
        "H": ("Status", "left"), "I": ("Erreichung", "left")})
    ws[f"I{R_CHECK_HEAD}"].alignment = C.align("left", "center", 0)
    checks = [   # (KPI, Istwert, Zielname, Gelbname, Format)
        ("CF", "=CF_nSt_Monat_J1", None, None, "eur"),
        ("BMR", "=Bruttomietrendite", "Ampel_BMR_gruen", "Ampel_BMR_gelb", "pct1"),
        ("NMR", "=Nettomietrendite", "Ampel_NMR_gruen", "Ampel_NMR_gelb", "pct1"),
        ("DSCR", "=DSCR_J1", "Ampel_DSCR_gruen", "Ampel_DSCR_gelb", "dscr"),
        ("EKR", "=EKR_Tile", "Ampel_EKR_gruen", "Ampel_EKR_gelb", "pct1"),
        ("IRR", "=EK_IRR", "Ampel_IRR_gruen", "Ampel_IRR_gelb", "pct1"),
    ]
    rows = {}
    for i, (key, ist, gruen, gelb, fmt) in enumerate(checks):
        r = R_CHECK + i
        rows[key] = r
        C.set_height(ws, r, C.H_STEP_ROW)
        C.safe_merge(ws, "B", r, "C", r)
        lab = ws[f"B{r}"]
        if key == "IRR":
            lab.value = C.kpi_label("IRR", formula=True)
        else:   # „(Jahr 1)“ passt nicht in 180 px bei 10 pt – die Kachel darüber trägt das volle Label
            C.set_text(lab, "Cashflow n. St. / Monat" if key == "CF" else C.kpi_label(key))
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", 1)
        v = ws[f"E{r}"]
        v.value, v.number_format = ist, C.NUMFMT[fmt]
        v.font = C.font(C.T_BODY, True, C.NAVY)
        v.alignment = C.align("right", "center", 1)
        z = ws[f"F{r}"]
        z.value = "=" + C.threshold_text(key, prefix="≥ ")
        z.number_format = "General"
        z.font = C.font(C.T_BODY, False, C.MUTED)
        z.alignment = C.align("right", "center", 1)
        s = ws[f"H{r}"]
        if key == "CF":
            s.value = f'=IF(E{r}<0,"kritisch",IF({C.CF_YEAR2}<0,"prüfen","erfüllt"))'
        else:
            s.value = f'=IF(E{r}>={gruen},"erfüllt",IF(E{r}>={gelb},"prüfen","kritisch"))'
        s.number_format = C.NUMFMT["status_dot"]
        s.font = C.font(C.T_MICRO, True, C.MUTED)
        s.alignment = C.align("left", "center", 1)
        q = ws[f"I{r}"]
        q.value = (f"=IF(E{r}>=0,1,0)" if key == "CF" else
                   f"=IF({gruen}>0,MIN(MAX(0,E{r}/{gruen}),1),IF(E{r}>={gruen},1,0))")
        q.number_format = C.NUMFMT["hidden"]
        q.alignment = C.align("left", "center", 0)
        C.hairline(ws, r, "B", "I")
    last = R_CHECK + len(checks) - 1
    # Statusspalte als Chip (8 pt fett in Statusfarbe) – Istwerte neutral, negative Beträge rot (P15)
    C.status_cf(ws, f"H{R_CHECK}:H{last}", [(f'$H{R_CHECK}="kritisch"', "red"), (f'$H{R_CHECK}="prüfen"', "amber"),
                                            (f'$H{R_CHECK}="erfüllt"', "green")])
    C.neg_red(ws, f"E{rows['CF']}")
    # Erreichung: neutraler Balken (keine zweite Statusfarbe neben dem Chip), voll = Ziel erreicht,
    # 88 % der Zellbreite → rechts bleibt Luft zur Tabellenkante
    ws.conditional_formatting.add(f"I{R_CHECK}:I{last}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=1, color=BAR_NEUTRAL,
        showValue=False, minLength=0, maxLength=88))
    # Fußnote: Legende + Link zu den Schwellen
    C.set_height(ws, R_CHECK_NOTE, 20)
    C.safe_merge(ws, "B", R_CHECK_NOTE, "F", R_CHECK_NOTE)
    n = ws[f"B{R_CHECK_NOTE}"]
    n.value = C.status_legend()
    n.value.append(C.rich([("   ·   Balken voll = Ziel erreicht", C.T_MICRO, False, C.MUTED)])[0])
    n.alignment = C.align("left", "center", 1)
    C.safe_merge(ws, "H", R_CHECK_NOTE, "I", R_CHECK_NOTE)
    lk = ws[f"H{R_CHECK_NOTE}"]
    C.text_link(lk, "Schwellen  ›", "Konfiguration", "B74", size=C.T_MICRO, bold=False,
                tooltip="Ampel-Schwellenwerte in der Konfiguration anpassen")
    lk.alignment = C.align("right", "center", 1)
    return rows


# ========================================================================================== Diagramme
def _txpr(size=C.T_MICRO, color=C.MUTED, bold=False, rot=None, wrap=None):
    cp = CharacterProperties(sz=int(size * 100), b=bold, solidFill=color)
    cp.latin = None
    body = RichTextProperties(rot=rot, vert="horz") if rot is not None else RichTextProperties()
    if wrap is not None:
        body.wrap = wrap
    return RichText(bodyPr=body, p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _no_fill(gp=None):
    gp = gp or GraphicalProperties()
    gp.noFill = True
    gp.line = LineProperties()
    gp.line.noFill = True
    return gp


def _solid(colr):
    gp = GraphicalProperties(solidFill=colr)
    gp.line = LineProperties()
    gp.line.noFill = True
    return gp


def _anchor(chart, c1, r1, c2, r2):
    """Diagramm füllt genau sein Panel: from = erste Panelzelle, to = Zelle nach dem Panel (Offsets 0)."""
    chart.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=C.col(c1) - 1, row=r1 - 1, colOff=0, rowOff=0),
        to=AnchorMarker(col=C.col(c2), row=r2, colOff=0, rowOff=0))


def _abs_ref(ref):
    import re
    return re.sub(r"^([A-Z]+)(\d+)$", r"$\1$\2", ref)


def _legend_meta(ws, row, unit):
    """Mini-Legende „■ Zufluss  ■ Abfluss  ■ Summe · € je Monat“ rechts im Abschnittskopf (P2-10)."""
    c = ws[f"R{row}"]
    c.value = C.rich([("■ ", C.T_MICRO, False, INFLOW), ("Zufluss   ", C.T_MICRO, False, C.MUTED),
                      ("■ ", C.T_MICRO, False, OUTFLOW), ("Abfluss   ", C.T_MICRO, False, C.MUTED),
                      ("■ ", C.T_MICRO, False, TOTAL), ("Summe", C.T_MICRO, False, C.MUTED),
                      (f"   ·   {unit}", C.T_MICRO, False, C.BLUE)])
    c.alignment = C.align("right", "center", 1)


def _waterfall(ws):
    """Cashflow Jahr 1 als flacher Wasserfall aus gestapelten Säulen (2D laut Nutzerentscheidung).

    Farbsemantik (P2-10): Zufluss 4A86C8 · Abfluss B8C3D1 · Summen („= vor/nach Steuern“) 0B2A4A mit fettem Label.
    Die Steuer ist je nach Vorzeichen Zufluss (Erstattung) oder Abfluss (Zahlung): Ihr sichtbarer Teil liegt
    deshalb in zwei Reihen, von denen je nach Vorzeichen genau eine belegt ist.
    Excel stapelt positive und negative Werte getrennt. Jede Säule [unten, oben] wird in einen positiven Teil
    (unsichtbare Basis + sichtbares Stück) und einen negativen Teil zerlegt; so bleibt sie auch beim Kreuzen der
    Nulllinie ein zusammenhängender Balken. Direkt über jeder Säule liegt ein unsichtbarer Träger, dessen
    Beschriftung den Betrag zeigt (Reihenname = Zellbezug, gebietsschema-sicher per FIXED)."""
    C.section(ws, R_BAND1, "K", "R", "Cashflow Jahr 1")
    _legend_meta(ws, R_BAND1, "€ je Monat")
    # (Kategorie, Formel, verkettet, Farbe, Summe?, Vorzeichenfarbe?)
    cats = [("Miete", "=Miete_Ist_J1/12", False, INFLOW, False),
            ("Bewirt-\nschaftung", "=-BWK_J1/12", True, OUTFLOW, False),
            ("Zinsen", "=-Zins_J1/12", True, OUTFLOW, False),
            ("Tilgung", "=-Tilgung_J1/12", True, OUTFLOW, False),
            ("= vor Steuern", "=CF_vSt_J1/12", False, TOTAL, True),
            ("Steuer", "=-Steuer_J1/12", True, INFLOW, False),
            ("= nach Steuern", "=CF_nSt_J1/12", False, TOTAL, True)]
    tax = 5
    K = W_COLS
    n = len(cats)
    r_first, r_last = W_ROW, W_ROW + n - 1
    se = f"${K['start']}${r_first}:${K['end']}${r_last}"
    h = _abs_ref(W_HEAD)
    alt_row_p, alt_row_n = 31, 32                  # Steuer-Abfluss: sichtbare Teile je Kategorie (Spalten V:AB)
    alt_cols = ["V", "W", "X", "Y", "Z", "AA", "AB"]
    for i, (cat, frm, chain, _, _) in enumerate(cats):
        r = W_ROW + i
        C.set_text(ws[f"{K['cat']}{r}"], cat)
        ws[f"{K['val']}{r}"] = frm
        ws[f"{K['start']}{r}"] = f"={K['end']}{r - 1}" if chain else 0
        ws[f"{K['end']}{r}"] = f"={K['start']}{r}+{K['val']}{r}"
        lo = f"MIN({K['start']}{r},{K['end']}{r})"
        hi = f"MAX({K['start']}{r},{K['end']}{r})"
        ws[f"{K['base_p']}{r}"] = f"=IF({hi}>=0,MAX({lo},0),0)"
        ws[f"{K['base_n']}{r}"] = f"=IF({hi}>=0,0,IF({hi}+{h}>=0,{hi},{hi}+{h}))"
        vp = f"IF({hi}>=0,{hi}-MAX({lo},0),0)"
        vn = f"IF({hi}>=0,MIN({lo},0),{lo}-{hi})"
        if i == tax:   # Zahlung (Wert < 0) → Abfluss-Reihe, Erstattung → Zufluss-Reihe
            neg = f"{K['val']}{r}<0"
            ws[f"{K['vis_p']}{r}"] = f"=IF({neg},0,{vp})"
            ws[f"{K['vis_n']}{r}"] = f"=IF({neg},0,{vn})"
            ws[f"{alt_cols[i]}{alt_row_p}"] = f"=IF({neg},{vp},0)"
            ws[f"{alt_cols[i]}{alt_row_n}"] = f"=IF({neg},{vn},0)"
        else:
            ws[f"{K['vis_p']}{r}"] = f"={vp}"
            ws[f"{K['vis_n']}{r}"] = f"={vn}"
            ws[f"{alt_cols[i]}{alt_row_p}"] = 0
            ws[f"{alt_cols[i]}{alt_row_n}"] = 0
        for j, cc in enumerate(W_CARRIERS):
            ws[f"{cc}{r}"] = f"=IF({hi}+{h}>=0,{h},-{h})" if i == j else 0
        ws[f"{K['label']}{r}"] = (f'=IF({K["val"]}{r}>=0.5,"+",IF({K["val"]}{r}<=-0.5,"−",""))'
                                  f'&FIXED(ABS({K["val"]}{r}),0)')
    ws[W_HEAD] = f"=MAX(1,0.16*(MAX({se},0)-MIN({se},0)))"
    C.set_text(ws[f"U{alt_row_p}"], "Steuer als Abfluss +")
    C.set_text(ws[f"U{alt_row_n}"], "Steuer als Abfluss −")

    bar = BarChart()
    bar.type = "col"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.gapWidth = 60
    bar.legend = None
    bar.visible_cells_only = False               # Daten liegen in ausgeblendeten Spalten
    # Reihenfolge: positiver Stapel base_p → vis_p → alt_p → Träger; negativer Stapel base_n → Träger → vis_n → alt_n
    refs = [("base_p", K["base_p"]), ("base_n", K["base_n"]), ("vis_p", K["vis_p"]), ("alt_p", alt_row_p)]
    refs += [("car", cc) for cc in W_CARRIERS] + [("vis_n", K["vis_n"]), ("alt_n", alt_row_n)]
    for role, src in refs:
        if role in ("alt_p", "alt_n"):
            ref = Reference(ws, min_col=C.col("V"), max_col=C.col("AB"), min_row=src, max_row=src)
            bar.add_data(ref, from_rows=True, titles_from_data=False)
        else:
            bar.add_data(Reference(ws, min_col=C.col(src), min_row=r_first, max_row=r_last), titles_from_data=False)
    bar.set_categories(Reference(ws, min_col=C.col(K["cat"]), min_row=r_first, max_row=r_last))
    names = {"base_p": "Basis +", "base_n": "Basis −", "vis_p": "Betrag +", "vis_n": "Betrag −",
             "alt_p": "Abfluss +", "alt_n": "Abfluss −"}
    car = 0
    for s, (role, _) in zip(bar.series, refs):
        if role in ("base_p", "base_n"):
            s.tx = SeriesLabel(v=names[role])
            s.graphicalProperties = _no_fill()
        elif role in ("vis_p", "vis_n"):
            s.tx = SeriesLabel(v=names[role])
            s.graphicalProperties = _solid(TOTAL)
            for idx, (_, _, _, colr, _) in enumerate(cats):
                pt = DataPoint(idx=idx)
                pt.graphicalProperties = _solid(colr)
                s.dPt.append(pt)
        elif role in ("alt_p", "alt_n"):
            s.tx = SeriesLabel(v=names[role])
            s.graphicalProperties = _solid(OUTFLOW)
        else:
            r = W_ROW + car
            is_sum = cats[car][4]
            s.tx = SeriesLabel(strRef=StrRef(f=f"'{SHEET}'!${K['label']}${r}"))
            s.graphicalProperties = _no_fill()
            dl = DataLabel(idx=car, showLegendKey=False, showVal=False, showCatName=False, showSerName=True,
                           showPercent=False, showBubbleSize=False)
            dl.txPr = _txpr(C.T_MICRO, C.NAVY if is_sum else C.INK2, is_sum)
            s.dLbls = DataLabelList(dLbl=[dl], showLegendKey=False, showVal=False, showCatName=False,
                                    showSerName=False, showPercent=False, showBubbleSize=False)
            car += 1
    bar.y_axis.delete = False
    bar.y_axis.numFmt = C.typo_minus('#,##0" €";-#,##0" €";0" €"')
    bar.y_axis.majorUnit = 500                   # ruhiges Raster (P2-08); Grenzen automatisch (Werte variieren)
    bar.y_axis.majorTickMark = "none"
    bar.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=C.LINE, w=6350)))
    bar.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    bar.y_axis.txPr = _txpr(C.T_MICRO, C.MUTED)
    bar.x_axis.delete = False
    bar.x_axis.tickLblPos = "low"
    bar.x_axis.majorTickMark = "none"
    bar.x_axis.tickLblSkip = 1
    bar.x_axis.txPr = _txpr(C.T_MICRO, C.MUTED, False, rot=0)
    bar.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=C.MUTED2, w=12700))
    # Unsichtbare Linienreihe: macht das Diagramm zum Kombidiagramm, damit es flach (2D) bleibt –
    # Wasserfälle werden nach der Nutzerentscheidung nie dreidimensional dargestellt.
    flat = LineChart()
    flat.add_data(Reference(ws, min_col=C.col(K["end"]), min_row=r_first, max_row=r_last), titles_from_data=False)
    fs = flat.series[0]
    fs.tx = SeriesLabel(v="Stand")
    fs.smooth = False
    fs.marker = Marker(symbol="none")
    fs.graphicalProperties = GraphicalProperties()
    fs.graphicalProperties.line = LineProperties(noFill=True)
    bar += flat
    _anchor(bar, "K", R_CHECK_HEAD, "R", R_CHECK_NOTE)
    ws.add_chart(bar)


def _wealth_chart(ws, wb):
    """Vermögensentwicklung über 40 Jahre (flach, P2-09/P08): Y-Achse ab 0, nur horizontale Gitterlinien,
    Legende unten links. Verkaufsjahr als feine, blasse Marken-Säule vom Boden bis über die Linien; ihre Beschriftung
    „Verkauf 2037 / Nettovermögen / 197 T€“ steht damit frei über dem Kreuzungspunkt, ein navy Punkt markiert den Wert
    auf der Nettovermögenslinie. Nicht-Verkaufsjahre sind #NV (kein Punkt, keine Beschriftung)."""
    sale_year = "INDEX(Projektion!$D$10:$AQ$10,Haltedauer)"
    C.section(ws, R_BAND2, "B", "R", "Vermögensentwicklung",
              meta=f'="Verkauf nach "&Haltedauer&" J. ("&{sale_year}&")   ·   T€ · jeweils Jahresende · 40 Jahre"')
    ws[f"R{R_BAND2}"].data_type = "f"
    for r in range(R_CHART2[0], R_CHART2[1] + 1):
        C.set_height(ws, r, 20)
    C.set_text(ws[f"{MARK_COL}{MARK_ROW - 1}"], "Verkaufsmarke (Diagramm Vermögensentwicklung)")
    C.set_text(ws[f"{DOT_COL}{MARK_ROW - 1}"], "Punkt Nettovermögen im Verkaufsjahr")
    for i in range(40):
        y = i + 1
        top = f"MAX(INDEX(Projektion!$D$41:$AQ$41,{y}),INDEX(Projektion!$D$43:$AQ$43,{y}))"
        ws[f"{MARK_COL}{MARK_ROW + i}"] = f"=IF({y}=Haltedauer,{top}*1.12,NA())"
        ws[f"{DOT_COL}{MARK_ROW + i}"] = f"=IF({y}=Haltedauer,INDEX(Projektion!$D$43:$AQ$43,{y}),NA())"
    lbl = ws[MARK_LABEL]
    lbl.value = (f'="Verkauf "&{sale_year}&CHAR(10)&"Nettovermögen"&CHAR(10)'
                 '&FIXED(INDEX(Projektion!$D$43:$AQ$43,Haltedauer)/1000,0)&" T€"')
    line = LineChart()
    spec = [(41, "Immobilienwert", C.ACCENT, 19050, None),
            (42, "Restschuld", "8A94A6", 12700, "dash"),
            (43, "Nettovermögen", C.NAVY, 31750, None)]
    P = wb["Projektion"]
    for row, title, colr, w, dash in spec:
        line.add_data(Reference(P, min_col=4, max_col=43, min_row=row), from_rows=True, titles_from_data=False)
        s = line.series[-1]
        s.tx = SeriesLabel(v=title)
        s.smooth = False
        s.marker = Marker(symbol="none")
        s.graphicalProperties = GraphicalProperties()
        s.graphicalProperties.line = LineProperties(solidFill=colr, w=w, prstDash=dash)
    line.add_data(Reference(ws, min_col=C.col(DOT_COL), min_row=MARK_ROW, max_row=MARK_ROW + 39),
                  titles_from_data=False)
    dot = line.series[-1]
    dot.tx = SeriesLabel(v="Verkaufspunkt")
    dot.smooth = False
    dot.marker = Marker(symbol="circle", size=7)
    dot.marker.graphicalProperties = GraphicalProperties(solidFill=C.NAVY)
    dot.marker.graphicalProperties.line = LineProperties(solidFill=C.WHITE, w=12700)
    dot.graphicalProperties = GraphicalProperties()
    dot.graphicalProperties.line = LineProperties(noFill=True)
    mark = BarChart()
    mark.type = "col"
    mark.gapWidth = 500
    mark.visible_cells_only = False              # Hilfsreihen liegen in ausgeblendeten Spalten
    mark.add_data(Reference(ws, min_col=C.col(MARK_COL), min_row=MARK_ROW, max_row=MARK_ROW + 39),
                  titles_from_data=False)
    mk = mark.series[0]
    mk.tx = SeriesLabel(strRef=StrRef(f=f"'{SHEET}'!{_abs_ref(MARK_LABEL)}"))
    mk.graphicalProperties = _solid(MARK_FILL)
    mk.dLbls = DataLabelList(dLblPos="outEnd", showLegendKey=False, showVal=False, showCatName=False,
                             showSerName=True, showPercent=False, showBubbleSize=False)
    mk.dLbls.txPr = _txpr(C.T_MICRO, C.NAVY, True)      # dreizeilig: „Verkauf 2037“ / „Nettovermögen“ / „197 T€“
    mk.dLbls.spPr = GraphicalProperties(solidFill=C.WHITE)
    mk.dLbls.spPr.line = LineProperties(solidFill=C.TINT, w=6350)
    mark.set_categories(Reference(P, min_col=4, max_col=43, min_row=10))
    mark.legend.position = "b"
    mark.legend.txPr = _txpr(C.T_MICRO, C.MUTED)
    mark.legend.legendEntry = [LegendEntry(idx=0, delete=True), LegendEntry(idx=4, delete=True)]
    mark.legend.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.045, y=0.915, w=0.42, h=0.08))
    mark.plot_area.layout = Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge",
                                                             x=0.055, y=0.05, w=0.93, h=0.76))
    # Kategorien direkt aus Projektion (Zeile 10, Kalenderjahre). Beschriftet wird jedes 5. Jahr (in Excel wirksam).
    line.set_categories(Reference(P, min_col=4, max_col=43, min_row=10))
    ax = mark
    ax.y_axis.delete = False
    ax.y_axis.scaling.min = 0
    ax.y_axis.numFmt = C.NUMFMT["teur"]
    ax.y_axis.majorTickMark = "none"
    ax.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=C.LINE, w=6350)))
    ax.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ax.y_axis.txPr = _txpr(C.T_MICRO, C.MUTED)
    ax.x_axis.delete = False
    ax.x_axis.tickLblSkip = 5
    ax.x_axis.tickMarkSkip = 5
    ax.x_axis.majorTickMark = "out"
    ax.x_axis.numFmt = "0"
    ax.x_axis.txPr = _txpr(C.T_MICRO, C.MUTED, rot=0)
    ax.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=C.MUTED2, w=9525))
    mark += line                                   # Säule zuerst (liegt hinter den Linien), gemeinsame Achsen
    _anchor(mark, "B", R_CHART2[0], "R", R_CHART2[1])
    ws.add_chart(mark)


# ========================================================================================== Zeitverlauf
def _timeline(ws):
    """Kennzahlen im Zeitverlauf (P3-01, P2-01): Kopf zweizeilig „Jahr n“ + Kalenderjahr, gleiche Spaltenabstände,
    Summenstufen über C.sum_row, Rot nur in den Cashflow-Ergebniszeilen."""
    C.section(ws, R_BAND3, "B", "R", "Kennzahlen im Zeitverlauf", meta="€ p. a.  ·  Jahr 1 = 12 Monate ab Kaufdatum")
    hr = R_YEAR
    C.section(ws, hr, "B", "R", None, level=2, labels={"B": ("Projektjahr", "left")}, height=C.H_TILE_LABEL)
    C.style_range(ws, "B", R_CAL, "R", R_CAL, fil=C.fill(C.HEAD), brd=Border(bottom=C.side("thin", C.ACCENT)),
                  fnt=C.font(C.T_MICRO, False, C.MUTED))
    for c in C.iter_cells(ws, "B", hr, "R", hr):
        c.border = Border()
    C.set_height(ws, R_CAL, C.H_TILE_SUB)
    cal = ws[f"B{R_CAL}"]
    C.set_text(cal, "Kalenderjahr")
    cal.alignment = C.align("left", "top", 1)
    ws[f"B{hr}"].alignment = C.align("left", "bottom", 1)
    for y, cc in zip(YEARS, YEAR_COLS):
        c = ws[f"{cc}{hr}"]
        c.value = y
        c.number_format = '"Jahr "0'
        c.font = C.font(C.T_MICRO, True, C.BLUE)
        c.alignment = C.align("right", "bottom", 1)
        k = ws[f"{cc}{R_CAL}"]
        k.value = f"=INDEX(Projektion!$D$10:$AQ$10,{cc}${hr})"
        k.number_format = C.NUMFMT["year"]
        k.alignment = C.align("right", "top", 1)
    groups = [
        ("Cashflow", [
            ("row", "Nettokaltmiete Ist", "Projektion", 15, 1, None, False),
            ("row", "Bewirtschaftungskosten", "Projektion", 26, -1, None, False),
            ("row", "Kapitaldienst", "Projektion", 32, -1, None, False),
            ("davon", "davon Zinsen", "Projektion", 30, -1, None, False),
            ("davon", "davon Tilgung", "Projektion", 31, -1, None, False),
            ("sub", "= Cashflow vor Steuern", "Projektion", 33, 1, None, True),
            ("row", "± Steuerwirkung (Cash-Sicht)", "Projektion", 35, -1, C.NUMFMT["eur_plain_signed"], False),
            ("result", "= Cashflow nach Steuern", "Projektion", 36, 1, None, True),
            ("row", "Kumulierter Cashflow n. St.", "Projektion", 38, 1, None, True)]),
        ("Steuer", [
            ("row", "Abschreibungen (AfA)", "Steuern", 56, 1, None, False),
            ("row", "Steuerliches Ergebnis", "Steuern", 67, 1, None, True)]),
        ("Vermögen", [
            ("row", "Immobilienwert (Jahresende)", "Projektion", 41, 1, None, False),
            ("row", "Restschuld", "Projektion", 42, -1, None, False),
            ("result", "= Nettovermögen", "Projektion", 43, 1, None, False),
            ("pct", "EK-Rendite", "Projektion", 45, 1, C.NUMFMT["pct1"], False)]),
    ]
    r = R_CAL
    for title, items in groups:
        r += 1
        C.section(ws, r, "B", "R", title, level=2, variant="line", height=C.H_STEP_ROW)
        ws.cell(r, 2).alignment = C.align("left", "bottom", 1)
        for kind, label, sheet, prow, sign, fmt, neg_red in items:
            r += 1
            C.set_height(ws, r, C.H_ROW if kind != "davon" else C.H_TILE_SUB)
            C.safe_merge(ws, "B", r, "D", r)
            lab = ws[f"B{r}"]
            C.set_text(lab, label)
            for cc in YEAR_COLS:
                c = ws[f"{cc}{r}"]
                c.value = f"={'-' if sign < 0 else ''}INDEX({_q(sheet)}!$D${prow}:$AQ${prow},{cc}${hr})"
                c.number_format = fmt or C.NUMFMT["num"]
                c.alignment = C.align("right", "center", 1)
            C.hairline(ws, r, "B", "R")
            if kind == "davon":
                C.style_range(ws, "B", r, "R", r, fnt=C.font(C.T_SMALL, False, C.MUTED))
                lab.alignment = C.align("left", "center", 2)
            else:
                C.style_range(ws, "B", r, "R", r, fnt=C.font(C.T_BODY, False, C.INK))
                lab.alignment = C.align("left", "center", 1)
            if kind in ("sub", "result"):
                C.sum_row(ws, r, "B", "R", stage=kind, value_from="E")   # setzt Negativ-Rot selbst (P15)
            elif neg_red:
                C.neg_red(ws, f"E{r}:R{r}")
    r += 1
    C.set_height(ws, r, 20)
    C.safe_merge(ws, "B", r, "L", r)
    note = ws[f"B{r}"]
    note.value = C.rich([
        ("Cash-Sicht: + fließt zu, − fließt ab (Steuerwirkung: + Erstattung, − Zahlung)  ·  ", C.T_MICRO, False, C.MUTED),
        ("Rot", C.T_MICRO, True, C.RED), (" – negative Ergebnis- und Kumulwerte", C.T_MICRO, False, C.MUTED)])
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
    """Bis zu 6 aktive Hinweise aus Cockpit!AB50:AB74, Warnungen zuerst – über eine eigene Sortierspalte.
    Anzeige ohne Emoji (P2-14): „⚠“ → „▲“, „ℹ“ → „•“ (reine Anzeigeformel; die Quelltexte im Cockpit bleiben).
    Signalfarbe nie als Fläche (P41): alle Zeilen F3F7FC mit linker 3-px-Kante – Warnung rot (Kante + Schrift),
    Information Akzent. Gibt es mehr Hinweise als Plätze, verweist die letzte Zeile auf das Cockpit."""
    key_rng = f"${HINT_KEY_COL}${HINT_KEY_ROW}:${HINT_KEY_COL}${HINT_KEY_ROW + 24}"
    for i in range(25):
        src = f"Cockpit!$AB${50 + i}"
        ws[f"{HINT_KEY_COL}{HINT_KEY_ROW + i}"] = f'=IF({src}="","",IF(LEFT({src},1)="⚠",0,100)+{i + 1})'
    warn = 'COUNTIF(Cockpit!$AB$50:$AB$74,"⚠*")'
    total = f"COUNT({key_rng})"
    C.section(ws, band_row, "B", "R", "Prüfhinweise & steuerliche Einordnung", meta="")
    head = ws[f"R{band_row}"]
    head.value = (f'=IF({warn}=0,"Keine Warnungen",IF({warn}=1,"1 Warnung",{warn}&" Warnungen"))&"  ·  "&'
                  f'IF({total}-{warn}=0,"keine Hinweise",IF({total}-{warn}=1,"1 Hinweis",({total}-{warn})&" Hinweise"))')
    C.set_height(ws, band_row + 1, 6)
    first = band_row + 2
    for k in range(N_HINTS):
        r = first + k
        C.set_height(ws, r, C.H_STEP_ROW)
        C.safe_merge(ws, "B", r, "R", r)
        c = ws[f"B{r}"]
        pick = f'IFERROR(INDEX(Cockpit!$AB$50:$AB$74,MATCH(SMALL({key_rng},{k + 1}),{key_rng},0)),"")'
        mono = (f'IF(LEFT({pick},1)="⚠","{WARN_SYM}  "&TRIM(MID({pick},2,999)),'
                f'IF(LEFT({pick},1)="ℹ","{INFO_SYM}  "&TRIM(MID({pick},2,999)),{pick}))')
        if k == 0:
            c.value = f'=IF({total}=0,"Keine Prüfhinweise – alle Plausibilitätsprüfungen ohne Befund.",{mono})'
        elif k == N_HINTS - 1:
            c.value = (f'=IF({total}>{N_HINTS},"›  "&({total}-{N_HINTS - 1})&" weitere Hinweise im Cockpit",'
                       f'{mono})')
        else:
            c.value = "=" + mono
        c.font = C.font(C.T_SMALL, False, C.INK2)
        c.alignment = C.align("left", "center", 1, wrap=True)
    last = first + N_HINTS - 1
    rng = f"B{first}:R{last}"
    b = f"$B{first}"
    edge = C.side("hair", C.LINE2)
    ws.conditional_formatting.add(f"B{first}:B{last}", FormulaRule(
        formula=[f'LEFT({b},1)="{WARN_SYM}"'], font=Font(color=C.RED), fill=C.fill(C.TINT_XL),
        border=Border(left=C.side("thick", C.RED), bottom=edge), stopIfTrue=True))
    ws.conditional_formatting.add(f"B{first}:B{last}", FormulaRule(
        formula=[f'{b}<>""'], fill=C.fill(C.TINT_XL), border=Border(left=C.side("thick", C.ACCENT), bottom=edge),
        stopIfTrue=True))
    # Verbundene Zeile: Fläche/Schrift kommen aus der Ankerzelle B; die Trennlinie unten braucht jede Zelle
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{b}<>""'], border=Border(bottom=edge), stopIfTrue=True))
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
    """Druckwunsch (P1-02, umgesetzt zentral in global_rules.page_setup): Querformat, 1 Seite breit,
    Umbrüche vor „Vermögensentwicklung“ und „Kennzahlen im Zeitverlauf“."""
    ws.print_area = f"A1:R{last_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.row_breaks = RowBreak()
    for r in (R_BAND2, R_BAND3):
        ws.row_breaks.append(Break(id=r - 1))
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.45
    ws.freeze_panes = "A4"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
