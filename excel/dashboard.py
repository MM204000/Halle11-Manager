"""Blatt „Dashboard“ – Gesamtbewertung der Investition auf einer Seite (reines Darstellungsblatt).

Das Blatt liest ausschließlich vorhandene Namen und Ergebniszellen der Vorlage (Projektion, Steuern,
Cockpit-Prüfhinweise, Ampel-Schwellen der Konfiguration) und verändert keine bestehende Berechnung.

Runde 6 („Wow-Paket“):
  · Hero – nachtblaues Band über die Inhaltsbreite wie das Deckblatt eines Investment-Reports: Objekt, Adresse,
    „Erstellt für“, Kauf als, darunter das Gesamturteil groß in Weiß mit Status-Pill; rechts die goldene
    Linienzeichnung (icons.cover_art_path), unten eine feine Goldlinie.
  · Kennzahl-Tachos – drei flache Doughnut-Tachos (DSCR, Bruttomietrendite, IRR n. St.): 240°-Skala, Zonen
    rot/amber/grün aus den Ampel-Namen (erreichter Teil satt, Rest zart), Zeiger als schmales Tinten-Segment,
    Wert groß in der Öffnung, Ziel darunter. Hilfsdaten in ausgeblendeten Spalten (nur Anzeigeformeln).
  · Sparklines – Spalte „Verlauf 40 J.“ in der Zeitverlaufstabelle (sparklines.register auf Projektion/Steuern).
  · Icons (icons.py) in Abschnittsköpfen, Kachelköpfen und Tacho-Karten.
Alle Farben aus core-Tokens / chart_color.

Raster: A = Rand · sechs Spaltenpaare (B:C, E:F, H:I, K:L, N:O, Q:R) mit fünf Abstandsspalten
(D, G, J, M, P = 12 px). Jedes Paar ist 180 px breit, aufgeteilt 84 | 96 px. Linkes Panel B:I, rechtes Panel K:R,
Tacho-Karten B:F · H:L · N:R (je 372 px). Hilfsdaten in ausgeblendeten Spalten T:AP.
"""
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.pagebreak import Break, RowBreak

import core as C

try:                                               # Runde 6: Icon-System und Sparklines (Stufe 1, Agenten M/K)
    import icons as IC
except Exception:                                  # pragma: no cover – Blatt bleibt ohne Icons baubar
    IC = None
try:
    import sparklines as SP
except Exception:                                  # pragma: no cover
    SP = None

SHEET = "Dashboard"

# ------------------------------------------------------------------------------------------ Raster
W_A, W_B, GAP_W, EDGE_W = 12.0, 13.71, 1.7, 4.5   # 84 px · 96 px · 12 px · 31 px (Paar = 180 px)
PAIRS = [("B", "C"), ("E", "F"), ("H", "I"), ("K", "L"), ("N", "O"), ("Q", "R")]
GAPS = ["D", "G", "J", "M", "P"]
SPARK_COL = "E"                                    # „Verlauf 40 J.“ direkt hinter der Beschriftung
YEAR_COLS = ["F", "H", "I", "K", "L", "N", "O", "Q", "R"]
YEARS = [1, 2, 3, 5, 10, 15, 20, 30, 40]
FIRST_HELP, LAST_HELP = "T", "AP"                  # ausgeblendete Hilfsspalten
CHROME_COLS = 19                                   # Kopfleiste A:S – endet mit der Randspalte S an der Inhaltskante (P01)

# ------------------------------------------------------------------------------------------ Zeilen
R_HERO = 5                                         # Hero 5–14, Goldlinie 15
R_EYEBROW, R_OBJ, R_ADDR, R_META, R_RULE = 5, 6, 7, 8, 9
R_VERDICT = 10                                     # 10 Label · 11 Urteil (cockpit.py liest B{R_VERDICT+1}) · 12 Pill
R_PILL, R_REASONS, R_HERO_END, R_GOLD = 12, 13, 14, 15
R_LINKS = 17
R_TILE = 19                                        # 19 Label · 20 Wert · 21 Kontextzeile
R_GAUGE_BAND = 23                                  # Kennzahl-Tachos
R_GAUGE_HEAD = 25                                  # Kartenkopf · 26–33 Tacho-Rahmen · 34 Zonenlegende
R_GAUGE_CHART = (26, 33)                           # 26–29 Bogen · 30 Fuge (Mittellinie) · 31 Wert · 32 Ziel · 33 Abstand
GAUGE_VALUE_ROWS = (31, 32, 33)
R_GAUGE_LEGEND = 34
R_BAND1 = 36                                       # Kennzahlen-Check | Cashflow Jahr 1
R_CHECK_HEAD = 37
R_CHECK = 38                                       # 38–43
R_CHECK_NOTE = 44
R_BAND2 = 46                                       # Vermögensentwicklung (volle Breite)
R_CHART2 = (47, 58)
R_BAND3 = 60                                       # Kennzahlen im Zeitverlauf
R_YEAR, R_CAL = 61, 62                             # Kopf: Projektjahr · Kalenderjahr

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
GAUGE_COLS = ("AN", "AO", "AP")                    # Tacho-Hilfsdaten (je Tacho eine Spalte, Zeilen 19–31)
GAUGE_CAT_COL = "AM"                               # Segmentrollen (Kategorien) für finish_pro.tacho_style
GAUGE_ROW = 19
N_HINTS = 6

TOTAL = C.C_INK                                    # Summen („= vor/nach Steuern“) in Tinte (Runde 5)
# Wasserfall-Kategorien nach der mappenweiten Diagramm-Semantik (C.chart_color, Runde 5): Miete Blau · Bewirtschaftung
# Orange · Zinsen Violett · Tilgung Aqua · Steuer Gold · Summen Tinte – dieselbe Größe hat überall dieselbe Farbe.
WF_COLORS = {"Miete": C.chart_color("Miete"), "Bewirtschaftung": C.chart_color("Bewirtschaftung"),
             "Zinsen": C.chart_color("Zinsen"), "Tilgung": C.chart_color("Tilgung"), "Steuer": C.chart_color("Steuer")}
BAR_CHAR, BAR_STEPS = "█", 10                      # Erreichungsbalken als Textbalken in Statusfarbe (P1-14)
MARK_FILL = C.GOLD_LINE                            # Verkaufsmarke: zarte Goldsäule (Hervorhebung Verkaufsjahr, Runde 5)
WARN_SYM, INFO_SYM = "▲", "•"                      # monochrome Kennzeichnung (ⓘ fehlt in Calibri)

# Icons (icons.py): Abschnittsköpfe Navy 18 px, Kachelköpfe Gold 14 px auf Navy, Tacho-Karten Navy 18 px
SECTION_ICONS = {"Kennzahl-Tachos": "ziel", "Kennzahlen-Check": "schild", "Cashflow Jahr 1": "muenzen",
                 "Vermögensentwicklung": "trend", "Kennzahlen im Zeitverlauf": "kalender",
                 "Prüfhinweise & steuerliche Einordnung": "lupe"}
TILE_ICONS = {"GI": "haus", "EK": "schluessel", "CF": "muenzen", "BMR": "prozent", "DSCR": "bank", "IRR": "trend"}
ICON_INDENT = 4                                    # Titel-Einzug hinter dem Icon (≈ 36 px)

# Tachos: (KPI, Ist-Name, Label, Icon)
GAUGES = [("DSCR", "DSCR_J1", "DSCR (Jahr 1)", "bank"),
          ("BMR", "Bruttomietrendite", "Bruttomietrendite", "prozent"),
          ("IRR", "EK_IRR", None, "trend")]           # IRR: Label als Anzeigeformel mit Haltedauer
GAUGE_SCALE = 1.5                                  # Skalenende = 1,5 × Zielwert (grün)
GAUGE_NEEDLE = 0.012                               # Zeigerbreite als Anteil der Skala (Tintenkante macht ihn kräftig)
GAUGE_INSET = 82                                   # Tacho-Rahmen 82 px ab den Kartenkanten → 208 px breit, mittig
GAUGE_ZONE = (C.RED, C.AMBER, C.GREEN)
GAUGE_CATS = ("Zone rot", "Zone gelb", "Zone grün", "Zeiger", "Zone rot", "Zone gelb", "Zone grün", "unsichtbar")


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
    _link_row(ws)
    rows = _check_table(ws)
    _hero(ws, rows)
    _tiles(ws)
    _gauges(ws)
    _waterfall(ws)
    _wealth_chart(ws, wb)
    last = _timeline(ws)
    last = _hints(ws, last + 2)
    _footer(ws, last + 2)
    C.hide_cols(ws, FIRST_HELP, LAST_HELP)
    _page(ws, last + 3)
    C.cf_close(ws)
    _section_icons(ws)
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
    for r in range(4, 100):
        ws.row_dimensions[r].height = 15
    for r, h in ((4, C.H_GAP), (16, C.H_GAP), (18, C.H_GAP), (22, 20), (24, 6), (35, 20), (45, 20), (59, 20)):
        C.set_height(ws, r, h)


def _chrome(ws):
    """Kopfleiste wie auf allen Blättern (Reiter setzt navigation.py als Formen darüber)."""
    for r, h in ((1, 6), (2, 33), (3, 3)):
        C.set_height(ws, r, h)
        for cc in range(1, CHROME_COLS + 1):
            c = ws.cell(r, cc)
            c.fill = C.fill(C.GOLD if r == 3 else C.NAVY)   # Goldlinie unter der Kopfleiste (Runde 5)
            c.border = Border()


def _row_px(ws, r):
    h = ws.row_dimensions[r].height
    return round((15 if h is None else h) / 0.75)


def _hero(ws, rows):
    """Hero (Runde 6): nachtblaues Band B:R, Zeilen 5–14, darunter eine feine Goldlinie (Z. 15) – das Deckblatt
    eines Investment-Reports. Links die Textsäule (B:H, 468 px), rechts die goldene Linienzeichnung (Cover-Motiv,
    transparent über der Navy-Fläche, läuft nach links weich aus):
      Z. 5  Dachzeile „DASHBOARD · INVESTMENT-ÜBERSICHT“ (8,5 pt fett SKY; als Anzeigeformel, damit die mappenweite
            Brotkrumen-Regel sie nicht in Blau umfärbt) · rechts „‹  Start“ (SKY)
      Z. 6  Objekt 16 pt fett Weiß · Z. 7 Adresse · Bundesland · Z. 8 „Erstellt für …  ·  Kauf als …  ·  Kauf am …“
      Z. 9  kurze Goldregel (180 px) als Trenner zwischen Objekt und Urteil
      Z. 10 „GESAMTURTEIL“ · Z. 11 Urteil 30 pt fett Weiß · Z. 12 Status-Pill · Z. 13 Begründung (3 Zeilen, SKY)
    cockpit.py liest das Urteil aus B{R_VERDICT + 1} (Wörter „kritisch“ / „Prüfpunkten“ / „Solide“ bleiben)."""
    heights = {5: 24, 6: 24, 7: 18, 8: 18, 9: 12, 10: 20, 11: C.H_ROW3, 12: C.H_PILL, 13: C.text_row_height(3),
               14: C.H_GAP}
    for r, h in heights.items():
        C.set_height(ws, r, h)
    ws.row_dimensions[R_GOLD].height = 2.25
    navy = C.fill(C.NAVY)
    for r in range(R_HERO, R_HERO_END + 1):
        for c in C.iter_cells(ws, "B", r, "R", r):
            c.fill = navy
            c.border = Border()
    for c in C.iter_cells(ws, "B", R_GOLD, "R", R_GOLD):
        c.fill = C.fill(C.GOLD)
        c.border = Border()
    ind = 2

    def put(row, c2, value, size, bold, color, v="center", wrap=False):
        C.safe_merge(ws, "B", row, c2, row)
        cell = ws[f"B{row}"]
        cell.value = value
        cell.font = C.font(size, bold, color)
        cell.alignment = C.align("left", v, ind, wrap=wrap)
        return cell

    put(R_EYEBROW, "L", '="DASHBOARD  ·  INVESTMENT-ÜBERSICHT"', C.T_LABEL, True, C.SKY, "bottom")
    back = ws[f"R{R_EYEBROW}"]
    C.text_link(back, "‹  Start", "Start", size=C.T_LABEL, bold=False, tooltip="Zurück zur Startseite")
    back.font = C.font(C.T_LABEL, False, C.SKY)
    back.alignment = C.align("right", "bottom", 1)
    put(R_OBJ, "R", "=Obj_Name", C.T_H2, True, C.WHITE, "bottom")
    put(R_ADDR, "R", '=Obj_Adresse&"  ·  "&Bundesland', C.T_SMALL, False, C.SKY)
    kaufdatum = 'IFERROR(TEXT(DAY(Kaufdatum),"00")&"."&TEXT(MONTH(Kaufdatum),"00")&"."&YEAR(Kaufdatum),"–")'
    put(R_META, "R", C.minus_text(
        '=IFERROR(IF(Erstellt_fuer="","","Erstellt für "&Erstellt_fuer&"  ·  "),"")&"Kauf als "&'
        f'{C.RECHTSFORM_SHORT}&"  ·  Kaufpreis "&FIXED(Kaufpreis,0)&" €  ·  Kauf am "&{kaufdatum}'
        '&"  ·  Haltedauer "&Haltedauer&" Jahre"'), C.T_SMALL, False, C.SKY)
    # kurze Goldregel unter dem Objektblock (Einzug wie der Text: beginnt in B, 180 px)
    for c in C.iter_cells(ws, "B", R_RULE, "C", R_RULE):
        c.border = Border(bottom=C.side("medium", C.GOLD))

    # Gesamturteil
    b = f"$B${R_VERDICT + 1}"
    conds = [(f'ISNUMBER(SEARCH("kritisch",{b}))', "red"),
             (f'ISNUMBER(SEARCH("Prüfpunkten",{b}))', "amber"),
             (f'{b}="Solide"', "green")]
    status_rng = f"$H${R_CHECK}:$H${R_CHECK + 5}"
    lab = put(R_VERDICT, "H", None, C.T_LABEL, True, C.SKY, "bottom")
    C.set_text(lab, "GESAMTURTEIL")
    st = put(R_VERDICT + 1, "L", None, C.T_HERO, True, C.WHITE, "center")
    st.value = (f'=IF(OR($H${rows["DSCR"]}="kritisch",$H${rows["CF"]}="kritisch"),"Liquidität kritisch",'
                f'IF(COUNTIF({status_rng},"kritisch")>0,"Rendite kritisch",'
                f'IF(COUNTIF({status_rng},"prüfen")>0,"Solide mit Prüfpunkten","Solide")))')
    st.number_format = "General"
    # Status-Pill: zarter Status-Tint mit Statusschrift – hebt sich als heller Chip von der Navy-Fläche ab
    C.safe_merge(ws, "B", R_PILL, "E", R_PILL)
    pill = ws[f"B{R_PILL}"]
    C.status_pill(ws, pill, conditions=conds, style="pill",
                  suffix=f'COUNTIF({status_rng},"erfüllt")&" von 6 Kennzahlen erfüllt"')
    pill.font = C.font(C.T_SMALL, True, C.WHITE)          # Grundschrift (ohne Status) hell auf Navy
    pill.fill = C.fill(C.NAVY_2)
    pill.alignment = C.align("left", "center", 1)
    ex = put(R_REASONS, "L", None, C.T_SMALL, False, C.SKY, "center", wrap=True)
    ex.value = C.minus_text('="–  DSCR (Jahr 1) "&FIXED(DSCR_J1,2)&"×"&IF(DSCR_J1<Ampel_DSCR_gruen," statt mindestens "," bei Ziel ")'
                            '&FIXED(Ampel_DSCR_gruen,2)&"×"'
                            '&CHAR(10)&"–  IRR n. St. "&FIXED(EK_IRR*100,1)&" % über "&Haltedauer&" Jahre (Ziel "'
                            '&FIXED(Ampel_IRR_gruen*100,1)&" %)"'
                            '&CHAR(10)&"–  Cashflow n. St. "&FIXED(CF_nSt_Monat_J1,0)&" € / Monat im Jahr 1, ab Jahr 2 "'
                            f'&FIXED({C.CF_YEAR2},0)&" € / Monat"')
    ex.alignment = C.align("left", "center", ind + 1, wrap=True)
    # Cover-Motiv (Agent M): transparente Goldlinien über der Navy-Fläche, Motiv rechts
    if IC is not None:
        try:
            w = C.span_px(ws, "B", "R")
            h = sum(_row_px(ws, r) for r in range(R_HERO + 1, R_HERO_END + 1))
            IC.place_image(ws, f"B{R_HERO + 1}", IC.cover_art_path(w, h, background=False, focus="right",
                                                                    intensity=0.9), w, h)
        except Exception as exc:                                   # pragma: no cover
            print("Dashboard: Hero-Motiv fehlt:", exc)


def _link_row(ws):
    """Sprungleiste (P1-03): sechs gleich breite Plätze im Kachelraster. Zurück und die Sprünge sind Sekundär-Buttons
    (weiß, Rahmen 1D4F8A, 10 pt fett) – dieselbe Form, nur die Pfeilrichtung unterscheidet sie. Rechts der Kontext-Chip
    (P1-04, C.btn kind='chip': E7EEF7, Rahmen 4A86C8, 9 pt fett 1D4F8A) „Kauf als: Privat · ändern  ›“ – kein Eingabe-Gelb,
    kein ▾, denn die Auswahl selbst liegt auf Start D23."""
    C.btn(ws, "B", R_LINKS, "C", "‹  Schritt 12: Ergebnis", "S12 Ergebnis", kind="secondary",
          tooltip="Zurück zum letzten Schritt des Leitfadens", set_row=False)
    items = [("Cockpit  ›", "Cockpit", "Detailkennzahlen und Prüfhinweise"),
             ("Diagramme  ›", "Diagramme", "Alle Auswertungen als Diagramm"),
             ("Sensitivität  ›", "Sensitivität", "Was-wäre-wenn: Zins, Miete, Kaufpreis"),
             ("Bankgespräch  ›", "Bankgespräch", "Unterlagen für das Finanzierungsgespräch")]
    for (a, b), (text, target, tip) in zip(PAIRS[1:5], items):
        C.btn(ws, a, R_LINKS, b, text, target, kind="secondary", tooltip=tip, set_row=False)
    a, b = PAIRS[5]
    chip = C.btn(ws, a, R_LINKS, b, "Kauf als", "Start", kind="chip", target_cell="D23",
                 tooltip="Auswahl auf der Startseite ändern", set_row=False, size=C.T_SMALL)
    chip.value = f'="Kauf als: "&{C.RECHTSFORM_SHORT}&"  ·  ändern  ›"'
    chip.data_type = "f"
    C.set_height(ws, R_LINKS, C.H_BTN)


# ========================================================================================== Kacheln
def _tiles(ws):
    """Sechs Kacheln in kanonischer Reihenfolge und mit kanonischen Labels (P20: C.KPI_ORDER / C.kpi_label) –
    EINE Anatomie (P11, C.tile dark): Kopfstreifen · Wert 20 pt (Wertfarbe = Status) · Fußzeile links Kontext,
    rechts Status-Chip. Runde 6: rechts im Kopfstreifen ein feines Gold-Icon (14 px) je Kennzahl."""
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
        _icon(ws, f"{b}{R_TILE}", TILE_ICONS.get(key), C.GOLD, 14, align="right", dx=-8)


def _icon(ws, cell, name, color, px, align=None, dx=0, dy=0, valign="middle"):
    if IC is None or not name:
        return None
    try:
        return IC.place_icon(ws, cell, name, color, px=px, dx=dx, dy=dy, align=align, valign=valign)
    except Exception as exc:                                       # pragma: no cover
        print("Dashboard: Icon fehlt:", name, exc)
        return None


def _section_icons(ws):
    """Icons vor den Titeln der Abschnittsköpfe (Ebene 1): Navy 18 px, Titel mit Einzug dahinter."""
    for (r, c), cell in list(ws._cells.items()):
        if c != 2 or not isinstance(cell.value, str) or C.is_formula(cell.value):
            continue
        name = SECTION_ICONS.get(cell.value.strip())
        if not name or cell.font is None or (cell.font.sz or 0) < C.T_H3:
            continue
        cell.alignment = C.align("left", "center", ICON_INDENT)
        _icon(ws, cell.coordinate, name, C.NAVY, 18, dx=11)
    for a, rr in (("K", R_BAND1),):
        cell = ws[f"{a}{rr}"]
        name = SECTION_ICONS.get(str(cell.value or "").strip())
        if name:
            cell.alignment = C.align("left", "center", ICON_INDENT)
            _icon(ws, cell.coordinate, name, C.NAVY, 18, dx=11)


# ========================================================================================== Kennzahl-Tachos
def _gauges(ws):
    """Drei Kennzahl-Tachos (Runde 6) in Karten zu je 372 px (B:F · H:L · N:R).

    Karte (FBFAF7, Goldkante oben): Kopf (Icon Navy + Kennzahl 10 pt fett, rechts Status-Chip) · Tacho-Rahmen
    (Z. 26–33) · Zonenlegende „●  < 1,00×   ●  1,00 – 1,20×   ●  ≥ 1,20×“ in Statusfarben.
    Tacho (flacher DoughnutChart, 180°, untere Hälfte unsichtbar): Zonen rot/amber/grün aus den Ampel-Namen, Zeiger als
    schmales Tinten-Segment – im inneren Ring als Nadel bis in die Öffnung, im Zonenring als Marke. Stil (aktive Zone
    kräftig, übrige zurückgenommen, Lochgröße, Kreis oben im Rahmen) setzt finish_pro.tacho_style (Agent H) anhand der
    Segmentrollen (Kategorien „Zone rot/gelb/grün“, „Zeiger“, „unsichtbar“).
    In der durchsichtigen unteren Hälfte stehen in Zellen: Wert 20 pt fett Tinte (direkt unter der Mittellinie), Ziel
    (9 pt) und Abstand zum Ziel (8 pt). Skala 0 … 1,5 × Zielwert; Werte außerhalb stehen am Skalenende.
    Alle Hilfszellen sind Anzeigeformeln in den ausgeblendeten Spalten AM:AP (nur die Tachos lesen sie)."""
    C.section(ws, R_GAUGE_BAND, "B", "R", "Kennzahl-Tachos",
              meta="Zeiger = Istwert  ·  Zonen aus den Ampel-Schwellen der Konfiguration")
    C.set_height(ws, R_GAUGE_HEAD, C.H_STEP_ROW)
    for r in range(R_GAUGE_CHART[0], R_GAUGE_CHART[0] + 4):
        C.set_height(ws, r, 18)
    C.set_height(ws, R_GAUGE_CHART[0] + 4, 6)
    for r, h in zip(GAUGE_VALUE_ROWS, (C.H_TILE_VALUE, 20, 20)):
        C.set_height(ws, r, h)
    C.set_height(ws, R_GAUGE_LEGEND, 20)
    spans = [("B", "F"), ("H", "L"), ("N", "R")]
    for (c1, c2), hc, (kpi, name, label, icon) in zip(spans, GAUGE_COLS, GAUGES):
        _gauge_card(ws, c1, c2, hc, kpi, name, label, icon)


def _gauge_data(ws, hc, kpi, name):
    """Hilfsdaten eines Tachos in Spalte hc ab GAUGE_ROW (reine Anzeigeformeln). Rückgabe: {rolle: zeile}.
    Segmente (Kategorien in GAUGE_CAT_COL, von finish_pro.tacho_style als Rollen gelesen):
      Zone rot · Zone gelb · Zone grün (bis zum Zeiger) · Zeiger · Zone rot · Zone gelb · Zone grün (ab dem Zeiger) ·
      unsichtbar (= Skala, die untere Hälfte). Beide Ringe lesen denselben Block: im Zeigerring ist alles außer dem
      Zeiger unsichtbar, im Zonenring färbt tacho_style die Zone des Zeigers kräftig und die übrigen zurückgenommen."""
    spec = C.KPI[kpi]
    g, y = spec["green"], spec["yellow"]
    rows = {}
    k = [GAUGE_ROW]

    def put(role, formula, text=False):
        r = k[0]
        k[0] += 1
        rows[role] = r
        cell = ws[f"{hc}{r}"]
        cell.value = formula
        cell.number_format = "General" if text else "0.0000"
        return r

    C.set_text(ws[f"{hc}{GAUGE_ROW - 1}"], f"Tacho {C.kpi_label(kpi)} (Anzeige)")
    M = put("max", f"=MAX({GAUGE_SCALE}*{g},1.2*{y},0.0001)")
    W = put("w", f"={hc}{M}*{GAUGE_NEEDLE}")
    V = put("v", f"=IFERROR(MIN(MAX({name},0),{hc}{M}),0)")
    # Zeiger mittig auf dem Wert; liegt der Wert näher als eine halbe Zeigerbreite an einer Schwelle, rückt der Zeiger
    # ganz auf die Seite seines Status (4,0 % bei „gelb ab 4,0 %“ zeigt in die gelbe Zone, nicht auf die Fuge)
    v_, w2 = f"{hc}{V}", f"{hc}{W}/2"
    # (Abstand eine halbe Zeigerbreite zur Fuge: die Zone des Zeigers bleibt auf beiden Seiten eindeutig)
    snap = (f"IF(AND({v_}>={g},{v_}-{g}<{w2}),{g}+{w2},IF(AND({v_}>={y},{v_}-{y}<{w2}),{y}+{w2},"
            f"IF(AND({v_}<{y},{y}-{v_}<{w2}),{y}-3*{w2},IF(AND({v_}<{g},{g}-{v_}<{w2}),{g}-3*{w2},{v_}-{w2}))))")
    P0 = put("p0", f"=MIN(MAX({snap},0),{hc}{M}-{hc}{W})")
    P1 = put("p1", f"={hc}{P0}+{hc}{W}")
    m, w, p0, p1 = (f"{hc}{x}" for x in (M, W, P0, P1))
    seg = [f"=MIN({p0},{y})", f"=MAX(0,MIN({p0},{g})-{y})", f"=MAX(0,{p0}-{g})", f"={w}",
           f"=MAX(0,{y}-{p1})", f"=MAX(0,{g}-MAX({p1},{y}))", f"=MAX(0,{m}-MAX({p1},{g}))", f"={m}"]
    rows["seg"] = [put(f"s{i}", f) for i, f in enumerate(seg)]
    return rows


def _gauge_card(ws, c1, c2, hc, kpi, name, label, icon):
    rows = _gauge_data(ws, hc, kpi, name)
    spec = C.KPI[kpi]
    g, y = spec["green"], spec["yellow"]
    head, (ch1, ch2), leg = R_GAUGE_HEAD, R_GAUGE_CHART, R_GAUGE_LEGEND
    cols = [C.L(k) for k in range(C.col(c1), C.col(c2) + 1)]     # 5 Spalten: 84 · 96 · 12 · 84 · 96
    # Kartenfläche: FBFAF7, Goldkante oben, Haarlinie unter dem Kopf
    for r in range(head, leg + 1):
        for c in C.iter_cells(ws, c1, r, c2, r):
            c.fill = C.fill(C.TINT_XL)
            c.border = Border(top=C.side("medium", C.GOLD) if r == head else None,
                              bottom=C.side("thin", C.LINE) if r in (head, leg - 1) else None)
    # Kopf: Icon + Kennzahl (links) · Status-Chip (rechts)
    C.safe_merge(ws, cols[0], head, cols[2], head)
    lab = ws[f"{cols[0]}{head}"]
    if label is None:
        lab.value = C.kpi_label(kpi, formula=True)
        lab.data_type = "f"
    else:
        C.set_text(lab, label)
    lab.font = C.font(C.T_BODY, True, C.NAVY)
    lab.alignment = C.align("left", "center", 3)
    _icon(ws, lab.coordinate, icon, C.NAVY, 18, dx=8)
    C.safe_merge(ws, cols[3], head, cols[4], head)
    chip = ws[f"{cols[3]}{head}"]
    C.status_pill(ws, chip, kpi=kpi, value_ref=name, style="chip")
    chip.alignment = C.align("right", "center", 1)
    # Wert (20 pt) in der Öffnung des Bogens, darunter Ziel und Abstand – über die ganze Kartenbreite verbunden
    r_val, r_goal, r_gap = GAUGE_VALUE_ROWS
    fx_v = C._fmt_expr(kpi, name)
    for r, val, size, bold, colr in (
            (r_val, C.minus_text(f'=IFERROR({fx_v},"–")'), C.T_KPI, True, C.NAVY),
            (r_goal, "=" + C.threshold_text(kpi), C.T_SMALL, False, C.MUTED),
            (r_gap, C.minus_text(_gap_text(kpi, name, g)), C.T_MICRO, False, C.MUTED)):
        C.safe_merge(ws, c1, r, c2, r)
        cell = ws[f"{c1}{r}"]
        cell.value = val
        cell.number_format = "General"
        cell.font = C.font(size, bold, colr)
        cell.alignment = C.align("center", "top" if r == r_val else "center")
    # Zonenlegende in Statusfarben: links „< gelb“, Mitte „gelb – grün“, rechts „≥ grün“
    fy, fg = C._fmt_expr(kpi, y), C._fmt_expr(kpi, g)
    lo = ws[f"{cols[0]}{leg}"]
    lo.value = C.minus_text(f'="●  < "&{fy}')
    lo.font = C.font(C.T_MICRO, True, C.RED)
    lo.alignment = C.align("left", "center", 1)
    C.safe_merge(ws, cols[1], leg, cols[3], leg)
    mid = ws[f"{cols[1]}{leg}"]
    mid.value = C.minus_text(f'="●  "&{fy.replace(" %", "").replace("×", "")}&" – "&{fg}')
    mid.font = C.font(C.T_MICRO, True, C.AMBER)
    mid.alignment = C.align("center", "center")
    hi = ws[f"{cols[4]}{leg}"]
    hi.value = C.minus_text(f'="●  ≥ "&{fg}')
    hi.font = C.font(C.T_MICRO, True, C.GREEN)
    hi.alignment = C.align("right", "center", 1)

    # Doughnut: zwei Ringe auf denselben Daten (innen Zeiger, außen Zonen); Stil setzt finish_pro.tacho_style
    ch = DoughnutChart(holeSize=46, firstSliceAng=270)
    ch.varyColors = True
    seg = rows["seg"]
    data = Reference(ws, min_col=C.col(hc), min_row=seg[0], max_row=seg[-1])
    cats = Reference(ws, min_col=C.col(GAUGE_CAT_COL), min_row=seg[0], max_row=seg[-1])
    for i, role in enumerate(GAUGE_CATS):
        C.set_text(ws[f"{GAUGE_CAT_COL}{seg[0] + i}"], role)
    ch.add_data(data, titles_from_data=False)
    ch.add_data(data, titles_from_data=False)
    ch.set_categories(cats)
    s_in, s_out = ch.series
    short = C.kpi_label(kpi)
    s_in.tx = SeriesLabel(v=f"Tacho {short} Zeiger")
    s_out.tx = SeriesLabel(v=f"Tacho {short} Zonen")
    _slices(s_out, [GAUGE_ZONE[0], GAUGE_ZONE[1], GAUGE_ZONE[2], C.NAVY,
                    GAUGE_ZONE[0], GAUGE_ZONE[1], GAUGE_ZONE[2], None])
    _slices(s_in, [None, None, None, C.NAVY, None, None, None, None])
    ch.legend = None
    ch.title = None
    ch.visible_cells_only = False                  # Hilfsdaten liegen in ausgeblendeten Spalten
    ch.graphical_properties = _no_fill()
    ch.plot_area.graphicalProperties = _no_fill()
    # Rahmen mittig über der Karte, 208 px breit (Karte 372 px): tacho_style setzt den Kreis oben in den Rahmen
    # (Ø = Rahmenhöhe − 6 px ≈ 192 px); Wert, Ziel und Abstand stehen in der durchsichtigen unteren Hälfte.
    x0 = GAUGE_INSET
    x1 = C.span_px(ws, c1, c2) - GAUGE_INSET
    c_from, off_from = _px_to_col(ws, c1, x0)
    c_to, off_to = _px_to_col(ws, c1, x1)
    ch.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=C.col(c_from) - 1, row=ch1 - 1, colOff=int(off_from * 9525), rowOff=0),
        to=AnchorMarker(col=C.col(c_to) - 1, row=ch2, colOff=int(off_to * 9525), rowOff=0))
    ws.add_chart(ch)


def _px_to_col(ws, c1, x):
    """Pixelposition x ab der linken Kante von c1 → (Spalte, Versatz in px)."""
    k = C.col(c1)
    while True:
        w = C.col_px(ws, C.L(k))
        if x < w:
            return C.L(k), x
        x -= w
        k += 1


def _gap_text(kpi, name, g):
    """Abstand zum Ziel: „noch 0,59× bis zum Ziel“ bzw. „Ziel erreicht · +0,32×“ (Prozente in %-Punkten)."""
    if "%" in C.KPI[kpi]["fmt"]:
        d = f'FIXED(ABS({name}-{g})*100,1)&" %-Pkt."'
    else:
        d = f'FIXED(ABS({name}-{g}),2)&"×"'
    return f'=IFERROR(IF({name}>={g},"Ziel erreicht  ·  +"&{d},"noch "&{d}&" bis zum Ziel"),"")'


def _slices(ser, colors):
    for i, colr in enumerate(colors):
        pt = DataPoint(idx=i)
        if colr is None:
            pt.graphicalProperties = _no_fill()
        else:
            gp = GraphicalProperties(solidFill=colr)
            gp.line = LineProperties(solidFill=C.CHART_SEP, w=15875)
            pt.graphicalProperties = gp
        ser.dPt.append(pt)
    ser.graphicalProperties = _no_fill()


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
        # Erfüllungsgrad (reine Anzeige): voll = „erfüllt“. Cashflow: Jahr 1 ≥ 0, aber Jahr 2 negativ → 60 % („prüfen“)
        grad = (f"IF(E{r}<0,0,IF({C.CF_YEAR2}<0,0.6,1))" if key == "CF" else
                f"IF({gruen}>0,MIN(MAX(0,E{r}/{gruen}),1),IF(E{r}>={gruen},1,0))")
        q.value = f'=REPT("{BAR_CHAR}",ROUND({grad}*{BAR_STEPS},0))'
        q.number_format = "General"
        q.font = C.font(C.T_MICRO, False, C.MUTED2)      # 10 × █ bei 8 pt ≈ 78 px < 96 px Spalte
        q.alignment = C.align("left", "center", 0)
        C.hairline(ws, r, "B", "I")
    last = R_CHECK + len(checks) - 1
    # Statusspalte als Chip (8 pt fett in Statusfarbe) – Istwerte neutral, negative Beträge rot (P15)
    C.status_cf(ws, f"H{R_CHECK}:H{last}", [(f'$H{R_CHECK}="kritisch"', "red"), (f'$H{R_CHECK}="prüfen"', "amber"),
                                            (f'$H{R_CHECK}="erfüllt"', "green")])
    C.neg_red(ws, f"E{rows['CF']}")
    # Erreichung (P1-14): Textbalken in der Statusfarbe der Zeile – ein voller Balken steht nie neben „prüfen“
    C.status_cf(ws, f"I{R_CHECK}:I{last}", [(f'$H{R_CHECK}="kritisch"', "red"), (f'$H{R_CHECK}="prüfen"', "amber"),
                                            (f'$H{R_CHECK}="erfüllt"', "green")], bold=False)
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
    """Lesehilfe rechts im Abschnittskopf: Die Kategorien tragen ihre Semantikfarbe und sind an der Achse beschriftet,
    deshalb erklärt die Legende nur die Richtung (Zufluss über, Abfluss unter dem Balken) und die Summen (P2-10)."""
    c = ws[f"R{row}"]
    c.value = C.rich([("+ ", C.T_MICRO, True, C.INK), ("Zufluss   ", C.T_MICRO, False, C.MUTED),
                      (C.MINUS + " ", C.T_MICRO, True, C.INK), ("Abfluss   ", C.T_MICRO, False, C.MUTED),
                      ("■ ", C.T_MICRO, False, TOTAL), ("Summe", C.T_MICRO, False, C.MUTED),
                      (f"   ·   {unit}", C.T_MICRO, False, C.BLUE)])
    c.alignment = C.align("right", "center", 1)


def _line_legend(ws, row):
    """Linienlegende „— Immobilienwert  – – Restschuld  — Nettovermögen · T€ · Jahresende“ rechts im Abschnittskopf
    (P3-04) – gleiche Bauart wie die Wasserfall-Legende (8 pt, Symbol in Serienfarbe, Einheit in 1D4F8A)."""
    c = ws[f"R{row}"]
    c.value = C.rich([("━━  ", C.T_MICRO, True, C.chart_color("Immobilienwert")), ("Immobilienwert   ", C.T_MICRO, False, C.MUTED),
                      ("– – –  ", C.T_MICRO, True, C.chart_color("Restschuld")), ("Restschuld   ", C.T_MICRO, False, C.MUTED),
                      ("━━  ", C.T_MICRO, True, C.chart_color("Nettovermögen")), ("Nettovermögen", C.T_MICRO, False, C.MUTED),
                      ("   ·   T€ · jeweils Jahresende", C.T_MICRO, False, C.BLUE)])
    c.alignment = C.align("right", "center", 1)


def _waterfall(ws):
    """Cashflow Jahr 1 als flacher Wasserfall aus gestapelten Säulen (2D laut Nutzerentscheidung).

    Farbsemantik (Runde 5, C.chart_color): jede Kategorie in ihrer mappenweiten Farbe, Summen („= vor/nach Steuern“)
    in Tinte mit fettem Label; die Richtung zeigt die Lage des Balkens und das Vorzeichen der Beschriftung.
    Die Steuer ist je nach Vorzeichen Zufluss (Erstattung) oder Abfluss (Zahlung): Ihr sichtbarer Teil liegt
    deshalb in zwei Reihen, von denen je nach Vorzeichen genau eine belegt ist.
    Excel stapelt positive und negative Werte getrennt. Jede Säule [unten, oben] wird in einen positiven Teil
    (unsichtbare Basis + sichtbares Stück) und einen negativen Teil zerlegt; so bleibt sie auch beim Kreuzen der
    Nulllinie ein zusammenhängender Balken. An jeder Säule liegt ein unsichtbarer Träger (als letzte Reihen gestapelt),
    dessen Beschriftung den Betrag zeigt (Reihenname = Zellbezug, gebietsschema-sicher per FIXED): Zuflüsse und
    positive Summen ÜBER dem Balken, Abflüsse und negative Summen UNTER dem Balken (P2-12) – nie auf der Nulllinie.
    Liegt ein Abfluss ganz im Plus (bzw. ein Zufluss ganz im Minus), bleibt die Beschriftung an der freien Kante."""
    C.section(ws, R_BAND1, "K", "R", "Cashflow Jahr 1")
    _legend_meta(ws, R_BAND1, "€ je Monat")
    # (Kategorie, Formel, verkettet, Farbe, Summe?, Vorzeichenfarbe?)
    F = WF_COLORS
    cats = [("Miete", "=Miete_Ist_J1/12", False, F["Miete"], False),
            ("Bewirt-\nschaftung", "=-BWK_J1/12", True, F["Bewirtschaftung"], False),
            ("Zinsen", "=-Zins_J1/12", True, F["Zinsen"], False),
            ("Tilgung", "=-Tilgung_J1/12", True, F["Tilgung"], False),
            ("= vor Steuern", "=CF_vSt_J1/12", False, TOTAL, True),
            ("Steuer", "=-Steuer_J1/12", True, F["Steuer"], False),
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
        ws[f"{K['base_n']}{r}"] = f"=IF({hi}>=0,0,{hi})"
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
            # Träger oben (+h, positiver Stapel) oder unten (−h, negativer Stapel unter dem Balken)
            ws[f"{cc}{r}"] = (f"=IF({lo}>=0,{h},IF({hi}<0,-{h},IF({K['val']}{r}>=0,{h},-{h})))"
                              if i == j else 0)
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
    # Reihenfolge: positiver Stapel base_p → vis_p → alt_p → Träger (über dem Balken);
    # negativer Stapel base_n → vis_n → alt_n → Träger (unter dem Balken)
    refs = [("base_p", K["base_p"]), ("base_n", K["base_n"]), ("vis_p", K["vis_p"]), ("alt_p", alt_row_p),
            ("vis_n", K["vis_n"]), ("alt_n", alt_row_n)]
    refs += [("car", cc) for cc in W_CARRIERS]
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
            s.graphicalProperties = _solid(F["Steuer"])   # enthält nur die Steuer (als Zahlung)
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
    # Grenzen und Intervall automatisch: Die Träger reichen 16 % der Spannweite über/unter die Balken, damit
    # endet die Achse knapp über dem größten Wert (kein leeres Band, P2-12)
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
    C.section(ws, R_BAND2, "B", "R", "Vermögensentwicklung")
    _line_legend(ws, R_BAND2)
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
    spec = [(41, "Immobilienwert", C.chart_color("Immobilienwert"), 22225, None),
            (42, "Restschuld", C.chart_color("Restschuld"), 15875, "dash"),
            (43, "Nettovermögen", C.chart_color("Nettovermögen"), 31750, None)]
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
    dot.marker = Marker(symbol="circle", size=8)
    dot.marker.graphicalProperties = GraphicalProperties(solidFill=C.CHART_MARK)   # Gold-Marke (Verkaufspunkt)
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
    mk.dLbls.spPr.line = LineProperties(solidFill=C.GOLD_LINE, w=6350)
    mark.set_categories(Reference(P, min_col=4, max_col=43, min_row=10))
    mark.legend = None                             # Legende steht als Text im Abschnittskopf (P3-04, wie „Cashflow Jahr 1“)
    mark.plot_area.layout = Layout(manualLayout=ManualLayout(layoutTarget="inner", xMode="edge", yMode="edge",
                                                             x=0.055, y=0.05, w=0.93, h=0.84))
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
    ax.x_axis.majorTickMark = "none"               # Kategorieachsen ohne Teilstriche (P3-04)
    ax.x_axis.minorTickMark = "none"
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
    # Sparkline-Spalte „Verlauf 40 J.“ (Runde 6): Kopf zentriert, Zellen darunter bleiben leer (Sparkline-Ziel)
    sp = ws[f"{SPARK_COL}{hr}"]
    C.set_text(sp, "VERLAUF")
    sp.font = C.font(C.T_MICRO, True, C.BLUE)
    sp.alignment = C.align("center", "bottom")
    spk = ws[f"{SPARK_COL}{R_CAL}"]
    C.set_text(spk, "40 Jahre")
    spk.alignment = C.align("center", "top")
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
    # (Art, Beschriftung, Blatt, Zeile, Vorzeichen, Format, Negativ-Rot, Sparkline-Vorlage bzw. dict)
    groups = [
        ("Cashflow", [
            ("row", "Nettokaltmiete Ist", "Projektion", 15, 1, None, False, dict(style="rent", min_zero=True)),
            ("row", "Bewirtschaftungskosten", "Projektion", 26, -1, None, False, dict(style="costs", min_zero=True)),
            ("row", "Kapitaldienst", "Projektion", 32, -1, None, False, dict(style="costs", min_zero=True)),
            ("davon", "davon Zinsen", "Projektion", 30, -1, None, False, dict(kind="column", color=C.chart_color("Zinsen"), min_zero=True)),
            ("davon", "davon Tilgung", "Projektion", 31, -1, None, False, dict(kind="column", color=C.chart_color("Tilgung"), min_zero=True)),
            ("sub", "= Cashflow vor Steuern", "Projektion", 33, 1, None, True, "cashflow"),
            ("row", "± Steuerwirkung (Cash-Sicht)", "Projektion", 35, -1, C.NUMFMT["eur_plain_signed"], False, None),
            ("result", "= Cashflow nach Steuern", "Projektion", 36, 1, None, True, "cashflow"),
            ("row", "Kumulierter Cashflow n. St.", "Projektion", 38, 1, None, True, "cashflow")]),
        ("Steuer", [
            ("row", "Abschreibungen (AfA)", "Steuern", 56, 1, None, False, dict(kind="column", color=C.chart_color("Steuer"), min_zero=True)),
            ("row", "Steuerliches Ergebnis", "Steuern", 67, 1, None, True,
             dict(kind="column", color=C.chart_color("Steuerliches Ergebnis"), neg=C.C_NEG))]),
        ("Vermögen", [
            ("row", "Immobilienwert (Jahresende)", "Projektion", 41, 1, None, False, "value"),
            ("row", "Restschuld", "Projektion", 42, -1, None, False, "debt"),
            ("result", "= Nettovermögen", "Projektion", 43, 1, None, False, "wealth"),
            ("pct", "EK-Rendite", "Projektion", 45, 1, C.NUMFMT["pct1"], False, dict(style="trend", markers=True))]),
    ]
    r = R_CAL
    for title, items in groups:
        r += 1
        C.section(ws, r, "B", "R", title, level=2, variant="line", height=C.H_STEP_ROW)
        ws.cell(r, 2).alignment = C.align("left", "bottom", 1)
        for kind, label, sheet, prow, sign, fmt, neg_red, spark in items:
            r += 1
            C.set_height(ws, r, C.H_ROW)            # eine Höhe je Tabelle, davon-Zeilen nie höher (P3-16)
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
                C.sum_row(ws, r, "B", "R", stage=kind, value_from="F")   # setzt Negativ-Rot selbst (P15)
            elif neg_red:
                C.neg_red(ws, f"F{r}:R{r}")
            _spark(ws, f"{SPARK_COL}{r}", f"{_q(sheet)}!D{prow}:AQ{prow}", spark)
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


def _spark(ws, cell, data, spec):
    """Sparkline „Verlauf 40 J.“ (sparklines.register): Vorlage (str) oder Einzelargumente (dict)."""
    if SP is None or spec is None:
        return
    kw = {"style": spec} if isinstance(spec, str) else dict(spec)
    try:
        SP.register(ws, cell, data, **kw)
    except Exception as exc:                                       # pragma: no cover
        print("Dashboard: Sparkline fehlt:", cell, exc)


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
        # typografisches Minus vor Beträgen („(−18.397 €)“), Bindestriche in Wörtern („10-Jahres-Frist“) bleiben
        pick = f'SUBSTITUTE(SUBSTITUTE({pick},"(-","({C.MINUS}")," -"," {C.MINUS}")'
        mono = (f'IF(LEFT({pick},1)="⚠","{WARN_SYM}  "&TRIM(MID({pick},2,999)),'
                f'IF(LEFT({pick},1)="ℹ","{INFO_SYM}  "&TRIM(MID({pick},2,999)),{pick}))')
        if k == 0:
            c.value = f'=IF({total}=0,"Keine Prüfhinweise – alle Plausibilitätsprüfungen ohne Befund.",{mono})'
        elif k == N_HINTS - 1:
            c.value = (f'=IF({total}>{N_HINTS},"+ "&({total}-{N_HINTS - 1})&" weitere Hinweise – Eingaben im Leitfaden '
                       f'(S01–S12) prüfen",{mono})')
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
