"""Erzeugt die Excel-Mappe „Immobilien-Kalkulation“ (MM Holding) nach Entwurf v4.

Aufruf:  python excel/build_immobilien_kalkulation.py [ausgabe.xlsx]

Alle Kennzahlen sind Formeln; Eingaben stehen ausschließlich auf dem Blatt
„Eingaben“ (blau) bzw. Zielwerte auf „Konfiguration“.
"""
import datetime as dt
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

# --------------------------------------------------------------------------- Design-Tokens
FONT = "Calibri"
TEAL, TEAL_L, TEAL_XL, SUM_BG = "0E4A4F", "E6F0F0", "F2F6F6", "F1F7F7"
INK, MUTED, MUTED2 = "1A1D21", "5B6068", "7A8088"
LINE, LINE2, NAV_BG = "EDEEF0", "DADCDF", "F4F5F6"
RED, RED_BG = "B42318", "FDF1EF"
AMB, AMB_BG = "B54708", "FEF4E8"
GRN, GRN_BG = "1F7A4D", "EDF7F1"
BLUE, INPUT_BG, INPUT_LINE = "2F6FB0", "EAF1FA", "B9CFE8"
ACC, GREY_BAR, WHITE = "4FA3A5", "A9B8C6", "FFFFFF"
GROUP = {"ein": "C4C8CE", "inp": BLUE, "aus": TEAL, "ber": "6E9C9E", "bank": "7F9BB8", "anh": LINE2}

SHEETS = [  # Name, Gruppe, Kurzbeschreibung
    ("Start", "ein", "Deckblatt, Inhaltsverzeichnis und Legende"),
    ("Leitfaden", "ein", "Schritt-für-Schritt-Anleitung und Konventionen"),
    ("Cockpit", "aus", "Gesamtbewertung, Kennzahlen, Diagramme und Prüfhinweise"),
    ("Eingaben", "inp", "Alle Annahmen zu Objekt, Kauf, Finanzierung, Miete, Steuern, Haushalt"),
    ("Projektion", "ber", "Jahresprojektion über 30 Jahre, Exit-Rechnung, Renditekennzahlen"),
    ("Finanzierung", "ber", "Investitions- und Finanzierungsplan, Tilgungsplan"),
    ("Steuern", "ber", "AfA-Bemessung, 15-%-Grenze, Einkünfte aus V+V, Steuereffekt"),
    ("AfA-Vergleich", "ber", "Lineare vs. degressive Gebäude-AfA im Vergleich"),
    ("Sensitivität", "ber", "Cashflow und DSCR bei veränderten Zinsen und Mieten"),
    ("Bankgespräch", "bank", "Objekt- und Finanzierungsübersicht inkl. Haushaltsrechnung"),
    ("Hinweise", "anh", "Annahmen, Vereinfachungen, Rechtsgrundlagen, Haftung"),
    ("Konfiguration", "anh", "Zielwerte, Ampelschwellen, Bankparameter, Auswahllisten"),
]
NAV = [s[0] for s in SHEETS]

# --------------------------------------------------------------------------- Zahlenformate
EUR = '#,##0 "€";-#,##0 "€";"–"'
NUM = '#,##0;-#,##0;"–"'
PCT = '0.0 %'
PCT2 = '0.00 %'
DSCR_F = '0.00'
MULT = '0.0"×"'
DATE = 'DD.MM.YYYY'
YEAR_F = '0'
JA_NEIN = '"Ja";"Nein";"Nein"'

# --------------------------------------------------------------------------- Helfer
def font(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def side(color, style="thin"):
    return Side(style=style, color=color)


LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = Workbook()
# Standardschrift der Mappe (auch für leere Zellen)
wb._fonts = IndexedList([font(10)])
wb._named_styles["Normal"].font = font(10)


def name(nm, ws, ref):
    """Legt einen mappenweiten Namen auf eine Zelle/einen Bereich an."""
    parts = []
    for p in ref.split(":"):
        col = "".join(ch for ch in p if ch.isalpha())
        row = "".join(ch for ch in p if ch.isdigit())
        parts.append(f"${col}${row}")
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!{':'.join(parts)}")


def put(ws, ref, value=None, fmt=None, f=None, bg=None, al=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = f or font()
    if fmt:
        c.number_format = fmt
    if bg:
        c.fill = fill(bg)
    c.alignment = al or (RIGHT if isinstance(value, (int, float)) or (isinstance(value, str) and value.startswith("=")) else LEFT)
    if border:
        c.border = border
    return c


def merge(ws, rng, value=None, **kw):
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    return put(ws, first, value, **kw)


def style_range(ws, rng, bg=None, border=None, f=None):
    for row in ws[rng]:
        for c in row:
            if bg:
                c.fill = fill(bg)
            if border:
                c.border = border
            if f:
                c.font = f


def input_cell(ws, ref, value, fmt=None, al=None):
    c = put(ws, ref, value, fmt=fmt, f=font(10, True, BLUE), bg=INPUT_BG,
            al=al or (RIGHT if not isinstance(value, str) else LEFT),
            border=Border(left=side(INPUT_LINE), right=side(INPUT_LINE), top=side(INPUT_LINE), bottom=side(INPUT_LINE)))
    c.protection = Protection(locked=False)
    return c


def col_range(c1, c2):
    return [get_column_letter(i) for i in range(c1, c2 + 1)]


def section(ws, row, c1, c2, title, right=None):
    """Abschnittsüberschrift mit teal Unterstrich (wie im Entwurf)."""
    ws.row_dimensions[row].height = 21
    for col in col_range(c1, c2):
        ws[f"{col}{row}"].border = Border(bottom=side(TEAL, "medium"))
    a, b = get_column_letter(c1), get_column_letter(c2)
    put(ws, f"{a}{row}", title, f=font(11, True, TEAL), al=Alignment(vertical="bottom"))
    if right:
        put(ws, f"{b}{row}", right, f=font(8.5, color=MUTED2), al=Alignment(horizontal="right", vertical="bottom"))


def table_header(ws, row, cells, dark=False):
    """cells: Liste (Bereich, Text, Ausrichtung)."""
    for rng, text, al in cells:
        if ":" in rng:
            ws.merge_cells(rng)
        first = rng.split(":")[0]
        c = ws[first]
        c.value = text
        c.font = font(9, True, WHITE if dark else TEAL)
        c.alignment = Alignment(horizontal=al, vertical="center", wrap_text=True)
        style_range(ws, rng if ":" in rng else f"{rng}:{rng}", bg=TEAL if dark else TEAL_L)


def row_line(ws, row, c1, c2, color=LINE):
    for col in col_range(c1, c2):
        c = ws[f"{col}{row}"]
        c.border = Border(bottom=side(color), top=c.border.top)


def sum_row(ws, row, c1, c2, res=False):
    for col in col_range(c1, c2):
        c = ws[f"{col}{row}"]
        c.fill = fill(TEAL_L if res else SUM_BG)
        c.border = Border(top=side(TEAL, "medium"), bottom=side(TEAL, "medium") if res else side(LINE))
        c.font = font(10, True, TEAL if res else INK)


def note(ws, ref, text, size=8.5):
    return put(ws, ref, text, f=font(size, color=MUTED2, italic=False), al=LEFT)


def setup(ws, title, group, extra_cols=0, meta=False):
    ws.sheet_properties.tabColor = GROUP[group]
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.column_dimensions["A"].width = 2.4
    for i in range(2, 14 + extra_cols):
        ws.column_dimensions[get_column_letter(i)].width = 12.9
    ws.row_dimensions[1].height = 8
    # Banner
    ws.row_dimensions[2].height = 18
    style_range(ws, "B2:M2", bg=TEAL)
    put(ws, "B2", "MM Holding", f=font(9, True, WHITE), bg=TEAL)
    merge(ws, "H2:M2", "Immobilien-Kalkulation · Version 2.0", f=font(8.5, color="CFE3E3"), bg=TEAL, al=RIGHT)
    # Titel
    ws.row_dimensions[3].height = 19
    ws.row_dimensions[4].height = 19
    merge(ws, "B3:G4", title, f=font(22, True, TEAL), al=Alignment(vertical="center"))
    if meta:
        items = [("Objekt", "=Objekt", "@"), ("Kauf", "=Kaufdatum", DATE), ("Investor", "=Investor", "@"),
                 ("Lage", "=Lage", "@"), ("Haltedauer", "=Haltedauer", '0 "Jahre"'), ("Stand", "=Stand", DATE)]
        for i, (lab, frm, fmt) in enumerate(items):
            col = get_column_letter(8 + i)
            put(ws, f"{col}3", lab, f=font(8, color=MUTED2), al=Alignment(horizontal="left", vertical="bottom"))
            put(ws, f"{col}4", frm, fmt=fmt, f=font(9, True),
                al=Alignment(horizontal="left", vertical="top", shrink_to_fit=True))
    # Navigation
    ws.row_dimensions[5].height = 6
    ws.row_dimensions[6].height = 19
    for i, nm in enumerate(NAV):
        c = ws.cell(6, 2 + i, nm)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{nm}'!A1", display=nm)
        active = nm == ws.title
        c.font = font(9, active, WHITE if active else MUTED)
        c.fill = fill(TEAL if active else NAV_BG)
        c.alignment = CENTER
        c.border = Border(bottom=side(LINE2), right=side(WHITE, "medium"))


def footer(ws, row, c2=13):
    ws.row_dimensions[row].height = 18
    for col in col_range(2, c2):
        ws[f"{col}{row}"].border = Border(top=side(LINE2))
    put(ws, f"B{row}", "Keine Gewähr. Ersetzt keine Rechts-, Steuer- oder Finanzberatung · Rechtsstand September 2026",
        f=font(8, color=MUTED2), al=LEFT)
    put(ws, f"{get_column_letter(c2)}{row}", "MM Holding GmbH · Kornhausgasse 4 · 88250 Weingarten · HRB 750729",
        f=font(8, color=MUTED2), al=RIGHT)


def printing(ws, area, landscape=True, repeat_rows=None, one_page=False):
    ws.print_area = area
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if one_page else 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.left.text = "MM Holding · Immobilien-Kalkulation · &A"
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = "Seite &P von &N"
    ws.oddFooter.right.size = 8
    if repeat_rows:
        ws.print_title_rows = repeat_rows


def status_cf(ws, rng, anchor, bold=True):
    """Einfärbung nach Status-Text (erfüllt/prüfen/kritisch) in Zelle anchor."""
    for word, color in (("erfüllt", GRN), ("prüfen", AMB), ("kritisch", RED)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{anchor}="{word}"'], font=Font(color=color, bold=bold)))


# Blätter in Entwurfsreihenfolge anlegen
ws_map = {}
for i, (nm, grp, _) in enumerate(SHEETS):
    ws_map[nm] = wb.active if i == 0 else wb.create_sheet()
    ws_map[nm].title = nm
S = ws_map

# =========================================================================== KONFIGURATION
ws = S["Konfiguration"]
setup(ws, "Konfiguration", "anh")
note(ws, "B8", "Zielwerte und Ampelschwellen steuern die Bewertung im Cockpit. Blaue Felder sind änderbar.")
section(ws, 10, 2, 13, "Zielwerte Kennzahlen-Check", "Ist ≥ Ziel → erfüllt")
table_header(ws, 11, [("B11:D11", "Kennzahl", "left"), ("E11", "Ziel", "right"), ("F11:M11", "Erläuterung", "left")])
targets = [
    ("Bruttomietrendite", 0.05, PCT, "Ziel_Brutto", "Jahres-Sollmiete / Kaufpreis"),
    ("Nettomietrendite", 0.03, PCT, "Ziel_Netto", "Reinertrag Jahr 1 (Ist-Miete − Bewirtschaftung) / Gesamtinvestition"),
    ("Kapitaldienstdeckung (DSCR)", 1.2, DSCR_F, "Ziel_DSCR", "Reinertrag / Kapitaldienst; Banken erwarten meist 1,10 bis 1,20"),
    ("Cashflow n. St. / Monat", 0, EUR, "Ziel_CF", "Mindest-Cashflow nach Steuern je Monat im Jahr 1"),
    ("EK-Rendite Jahr 1", 0.10, PCT, "Ziel_EKR", "(Cashflow n. St. + Tilgung + Wertzuwachs) / Eigenkapital"),
    ("IRR nach Steuern", 0.06, PCT, "Ziel_IRR", "Interner Zinsfuß des Eigenkapitals über die Haltedauer inkl. Verkauf"),
]
for i, (lab, val, fmt, nm, txt) in enumerate(targets):
    r = 12 + i
    merge(ws, f"B{r}:D{r}", lab)
    input_cell(ws, f"E{r}", val, fmt)
    merge(ws, f"F{r}:M{r}", txt, f=font(9, color=MUTED))
    row_line(ws, r, 2, 13)
    name(nm, ws, f"E{r}")

section(ws, 19, 2, 13, "Ampelschwellen und Bankparameter")
params = [
    ("Schwelle „prüfen“ (Erreichung ab)", 0.75, PCT, "Schwelle_Pruefen",
     "Ist ≥ Ziel → erfüllt · Ist ≥ Ziel × Schwelle → prüfen · darunter → kritisch"),
    ("Bankansatz Mieteinnahmen", 0.75, PCT, "Bank_Mietansatz",
     "Anteil der Sollmiete, den Banken in der Haushaltsrechnung anrechnen (üblich 60–80 %)"),
    ("Mindestreserve in Monatsraten", 6, '0 "Monate"', "Min_Reserve",
     "Empfohlene liquide Reserve nach dem Kauf, ausgedrückt in Monatsraten"),
]
for i, (lab, val, fmt, nm, txt) in enumerate(params):
    r = 20 + i
    merge(ws, f"B{r}:D{r}", lab)
    input_cell(ws, f"E{r}", val, fmt)
    merge(ws, f"F{r}:M{r}", txt, f=font(9, color=MUTED))
    row_line(ws, r, 2, 13)
    name(nm, ws, f"E{r}")

section(ws, 24, 2, 13, "Auswahllisten", "Quelle für Dropdowns auf „Eingaben“")
table_header(ws, 25, [("B25:D25", "AfA-Satz Gebäude", "left"), ("E25:G25", "Solidaritätszuschlag", "left"),
                      ("H25:J25", "Kirchensteuer", "left")])
lists = {"B": ([0.02, 0.025, 0.03], "Liste_AfA"), "E": ([0.0, 0.055], "Liste_Soli"), "H": ([0.0, 0.08, 0.09], "Liste_KiSt")}
for col, (vals, nm) in lists.items():
    for j, v in enumerate(vals):
        put(ws, f"{col}{26 + j}", v, fmt=PCT, al=LEFT)
    name(nm, ws, f"{col}26:{col}{25 + len(vals)}")
note(ws, "B30", "AfA: 2,0 % (fertiggestellt 1925–2022), 2,5 % (vor 1925), 3,0 % (ab 2023) – § 7 Abs. 4 EStG · "
                "Kirchensteuer: 8 % in Bayern und Baden-Württemberg, sonst 9 %.")
footer(ws, 32)
printing(ws, "A1:M32")

# =========================================================================== EINGABEN
ws = S["Eingaben"]
setup(ws, "Eingaben", "inp")
ws.row_dimensions[8].height = 20
input_cell(ws, "B8", "Eingabe", al=CENTER)
merge(ws, "C8:M8", "Blau hinterlegte Felder sind Eingaben – alle übrigen Blätter rechnen automatisch. "
                   "Die Werte sind ein Beispiel (3-Zi.-ETW in Weingarten).", f=font(9, color=MUTED))
table_header(ws, 10, [("B10:D10", "Parameter", "left"), ("E10", "Wert", "right"), ("F10", "Einheit", "left"),
                      ("G10:M10", "Hinweis / Quelle", "left")])

INPUTS = [
    ("Objekt und Rahmendaten", [
        ("Objektbezeichnung", "3-Zi.-ETW, 75 m²", "@", "", "Objekt", "Erscheint in Kopfzeilen und im Bankgespräch", None),
        ("Lage", "Weingarten", "@", "", "Lage", "", None),
        ("Investor", "Privatperson", "@", "", "Investor", "Steuerlogik: Privatvermögen, Einkünfte aus Vermietung und Verpachtung (§ 21 EStG)", None),
        ("Wohnfläche", 75, '0.0', "m²", "Wohnflaeche", "Laut Teilungserklärung / Wohnflächenberechnung", ("dec", 1, 10000)),
        ("Kaufdatum (Übergang Nutzen und Lasten)", dt.date(2026, 12, 1), DATE, "Datum", "Kaufdatum", "Jahr 1 = 12 Monate ab diesem Datum", ("date",)),
        ("Stand der Kalkulation", dt.date(2026, 9, 24), DATE, "Datum", "Stand", "", ("date",)),
        ("Geplante Haltedauer", 12, '0', "Jahre", "Haltedauer", "Verkauf nach mehr als 10 Jahren steuerfrei (§ 23 EStG)", ("int", 1, 30)),
    ]),
    ("Kaufpreis und Nebenkosten", [
        ("Kaufpreis", 285000, EUR, "€", "Kaufpreis", "Laut Kaufvertrag", ("dec", 0, 1e9)),
        ("davon bewegl. Gegenstände (Küche, Inventar)", 8000, EUR, "€", "Inventar", "Grunderwerbsteuerfrei, eigene AfA – Wert angemessen belegen", ("dec", 0, 1e9)),
        ("Grunderwerbsteuer", 0.05, PCT2, "%", "GrESt_Satz", "Baden-Württemberg 5,0 %; Bemessung ohne bewegliche Gegenstände", ("dec", 0, 0.1)),
        ("Notar", 0.015, PCT2, "%", "Notar_Satz", "Richtwert ca. 1,5 % des Kaufpreises", ("dec", 0, 0.1)),
        ("Grundbuch", 0.005, PCT2, "%", "Grundbuch_Satz", "Richtwert ca. 0,5 % des Kaufpreises", ("dec", 0, 0.1)),
        ("Makler (inkl. USt.)", 0.0357, PCT2, "%", "Makler_Satz", "Käuferanteil laut Exposé", ("dec", 0, 0.1)),
        ("Renovierung nach Kauf (netto)", 14250, EUR, "€", "Renovierung", "Nettobetrag der ersten 3 Jahre – 15-%-Grenze prüft Blatt „Steuern“", ("dec", 0, 1e9)),
        ("Verteilung Erhaltungsaufwand", 1, '0', "Jahre", "Erhaltung_Jahre", "1 = sofort abziehen; 2–5 Jahre nach § 82b EStDV möglich", ("int", 1, 5)),
    ]),
    ("Finanzierung", [
        ("Darlehensbetrag", 245000, EUR, "€", "Darlehen", "Annuitätendarlehen, konstante Rate über die gesamte Laufzeit", ("dec", 0, 1e9)),
        ("Sollzins p. a.", 0.037, PCT2, "%", "Zins", "Laut Finanzierungsangebot", ("dec", 0, 0.2)),
        ("Anfängliche Tilgung p. a.", 0.02, PCT2, "%", "Tilgung", "", ("dec", 0, 0.2)),
        ("Zinsbindung", 10, '0', "Jahre", "Zinsbindung", "", ("int", 1, 30)),
        ("Anschlusszins (Annahme)", 0.045, PCT2, "%", "Anschlusszins", "Annahme für die Zeit nach Ende der Zinsbindung", ("dec", 0, 0.2)),
    ]),
    ("Miete und Bewirtschaftung", [
        ("Nettokaltmiete Soll", 950, EUR, "€/Monat", "Miete_Monat", "Laut Mietvertrag bzw. Mietspiegel", ("dec", 0, 1e7)),
        ("Mietausfallwagnis", 0.03, PCT, "%", "Mietausfall", "Leerstand und Zahlungsausfälle", ("dec", 0, 1)),
        ("Mietsteigerung p. a.", 0.02, PCT, "%", "Mietsteigerung", "", ("dec", -0.1, 0.2)),
        ("Hausgeld nicht umlagefähig", 1290, EUR, "€/Jahr", "Hausgeld_nu", "WEG-Verwaltung, Bankgebühren etc. – steuerlich abziehbar", ("dec", 0, 1e7)),
        ("Instandhaltungsrücklage", 900, EUR, "€/Jahr", "Ruecklage", "Liquiditätswirksam; steuerlich erst bei Verwendung abziehbar", ("dec", 0, 1e7)),
        ("Sondereigentumsverwaltung / Sonstiges", 300, EUR, "€/Jahr", "Verwaltung", "Steuerlich abziehbar", ("dec", 0, 1e7)),
        ("Kostensteigerung p. a.", 0.02, PCT, "%", "Kostensteigerung", "", ("dec", -0.1, 0.2)),
    ]),
    ("Wertentwicklung und Verkauf", [
        ("Marktwert bei Kauf", 297000, EUR, "€", "Marktwert", "Annahme laut Bewertung / Gutachten", ("dec", 0, 1e9)),
        ("Wertsteigerung p. a.", 0.02, PCT, "%", "Wertsteigerung", "", ("dec", -0.1, 0.2)),
        ("Verkaufskosten bei Exit", 0.03, PCT, "%", "Verkaufskosten", "Makler, Notar u. Ä. beim Verkauf", ("dec", 0, 0.2)),
    ]),
    ("Steuern", [
        ("Gebäudeanteil an den Anschaffungskosten", 0.8, PCT, "%", "Gebaeudeanteil", "Aufteilung Gebäude / Grund und Boden, z. B. BMF-Arbeitshilfe", ("dec", 0, 1)),
        ("AfA-Satz Gebäude", 0.02, PCT, "%", "AfA_Satz", "Auswahl: 2,0 % / 2,5 % / 3,0 % (§ 7 Abs. 4 EStG)", ("list", "=Liste_AfA")),
        ("Nutzungsdauer bewegliche Gegenstände", 10, '0', "Jahre", "ND_Inventar", "Einbauküche i. d. R. 10 Jahre", ("int", 1, 30)),
        ("Grenzsteuersatz Einkommensteuer", 0.42, PCT, "%", "Grenzsteuersatz", "Persönlicher Grenzsteuersatz", ("dec", 0, 0.45)),
        ("Solidaritätszuschlag", 0.055, PCT, "%", "Soli", "Auswahl: 0 % / 5,5 % der Einkommensteuer", ("list", "=Liste_Soli")),
        ("Kirchensteuer", 0.0, PCT, "%", "KiSt", "Auswahl: 0 % / 8 % / 9 % der Einkommensteuer", ("list", "=Liste_KiSt")),
    ]),
    ("Haushalt (für das Bankgespräch)", [
        ("Nettoeinkommen Haushalt", 5200, EUR, "€/Monat", "Netto_Einkommen", "Beispielwert – bitte eigene Angaben eintragen", ("dec", 0, 1e7)),
        ("Lebenshaltungskosten", 1900, EUR, "€/Monat", "Lebenshaltung", "Beispielwert", ("dec", 0, 1e7)),
        ("Eigene Wohnkosten (Miete oder Rate)", 1100, EUR, "€/Monat", "Wohnkosten", "Beispielwert", ("dec", 0, 1e7)),
        ("Bestehende Kreditraten", 0, EUR, "€/Monat", "Kreditraten", "Beispielwert", ("dec", 0, 1e7)),
        ("Liquide Reserve nach dem Kauf", 25000, EUR, "€", "Reserve", "Beispielwert", ("dec", 0, 1e9)),
    ]),
]
r = 11
for title, rows in INPUTS:
    r += 1
    section(ws, r, 2, 13, title)
    for lab, val, fmt, unit, nm, hint, val_rule in rows:
        r += 1
        merge(ws, f"B{r}:D{r}", lab)
        input_cell(ws, f"E{r}", val, fmt)
        put(ws, f"F{r}", unit, f=font(9, color=MUTED))
        merge(ws, f"G{r}:M{r}", hint, f=font(9, color=MUTED))
        row_line(ws, r, 2, 4)
        row_line(ws, r, 6, 13)
        name(nm, ws, f"E{r}")
        if val_rule:
            kind = val_rule[0]
            if kind == "list":
                dv = DataValidation(type="list", formula1=val_rule[1], allow_blank=False)
            elif kind == "date":
                dv = DataValidation(type="date", operator="greaterThan", formula1="36526")
            else:
                dv = DataValidation(type="whole" if kind == "int" else "decimal", operator="between",
                                    formula1=str(val_rule[1]), formula2=str(val_rule[2]))
            dv.error = "Bitte einen gültigen Wert eingeben."
            dv.errorTitle = "Ungültige Eingabe"
            dv.showErrorMessage = True
            ws.add_data_validation(dv)
            dv.add(f"E{r}")
    r += 1
footer(ws, r + 1)
ws.freeze_panes = "A11"
printing(ws, f"A1:M{r + 1}", landscape=False, repeat_rows="10:10")

# =========================================================================== FINANZIERUNG
ws = S["Finanzierung"]
setup(ws, "Finanzierung", "ber")
section(ws, 8, 2, 6, "Investitionsplan", "€ · % v. Kaufpreis")
inv = [
    ("Kaufpreis", "=Kaufpreis"),
    ("Grunderwerbsteuer (ohne bewegliche Gegenstände)", "=(Kaufpreis-Inventar)*GrESt_Satz"),
    ("Notar", "=Kaufpreis*Notar_Satz"),
    ("Grundbuch", "=Kaufpreis*Grundbuch_Satz"),
    ("Makler", "=Kaufpreis*Makler_Satz"),
    ("Summe Kaufnebenkosten", "=SUM(E10:E13)"),
    ("Renovierung / Erhaltungsaufwand", "=Renovierung"),
    ("Gesamtinvestition", "=E9+E14+E15"),
]
for i, (lab, frm) in enumerate(inv):
    rr = 9 + i
    merge(ws, f"B{rr}:D{rr}", lab)
    put(ws, f"E{rr}", frm, fmt=EUR)
    put(ws, f"F{rr}", f"=E{rr}/Kaufpreis", fmt=PCT, f=font(9, color=MUTED))
    row_line(ws, rr, 2, 6)
sum_row(ws, 14, 2, 6)
sum_row(ws, 16, 2, 6, res=True)
name("Kaufnebenkosten", ws, "E14")
name("Gesamtinvestition", ws, "E16")

section(ws, 8, 9, 13, "Finanzierungsplan")
fin = [
    ("Darlehensbetrag", "=Darlehen", EUR, None),
    ("Eigenkapitalbedarf", "=Gesamtinvestition-Darlehen", EUR, "Eigenkapital"),
    ("Eigenkapitalquote", "=Eigenkapital/Gesamtinvestition", PCT, "EK_Quote"),
    ("Beleihungsauslauf (Darlehen / Marktwert)", "=Darlehen/Marktwert", PCT, "LTV"),
    ("Annuität p. a. (Zins + Tilgung)", "=Darlehen*(Zins+Tilgung)", EUR, "Annuitaet"),
    ("Monatsrate", "=Annuitaet/12", EUR, "Monatsrate"),
    ("Restschuld Ende Zinsbindung", "=INDEX($H$20:$H$49,Zinsbindung)", EUR, "Restschuld_ZB"),
    ("Kaufpreisfaktor (Kaufpreis / Jahres-Sollmiete)", "=Kaufpreis/(Miete_Monat*12)", MULT, "Faktor"),
]
for i, (lab, frm, fmt, nm) in enumerate(fin):
    rr = 9 + i
    merge(ws, f"I{rr}:L{rr}", lab)
    put(ws, f"M{rr}", frm, fmt=fmt)
    row_line(ws, rr, 9, 13)
    if nm:
        name(nm, ws, f"M{rr}")
sum_row(ws, 10, 9, 13)

section(ws, 18, 2, 9, "Tilgungsplan", "€ p. a. · jährliche Betrachtung")
table_header(ws, 19, [("B19", "Jahr", "center"), ("C19", "Zinssatz", "right"), ("D19", "Restschuld Anfang", "right"),
                      ("E19", "Zinsen", "right"), ("F19", "Tilgung", "right"), ("G19", "Kapitaldienst", "right"),
                      ("H19", "Restschuld Ende", "right"), ("I19", "Phase", "left")], dark=True)
ws.row_dimensions[19].height = 28
for t in range(1, 31):
    rr = 19 + t
    put(ws, f"B{rr}", t, fmt=YEAR_F, al=CENTER)
    put(ws, f"C{rr}", f"=IF(B{rr}<=Zinsbindung,Zins,Anschlusszins)", fmt=PCT2)
    put(ws, f"D{rr}", "=Darlehen" if t == 1 else f"=H{rr - 1}", fmt=NUM)
    put(ws, f"E{rr}", f"=D{rr}*C{rr}", fmt=NUM)
    put(ws, f"F{rr}", f"=MAX(0,MIN(Annuitaet-E{rr},D{rr}))", fmt=NUM)
    put(ws, f"G{rr}", f"=E{rr}+F{rr}", fmt=NUM, f=font(10, True))
    put(ws, f"H{rr}", f"=D{rr}-F{rr}", fmt=NUM)
    put(ws, f"I{rr}", f'=IF(B{rr}<=Zinsbindung,"Zinsbindung","Anschluss")', f=font(9, color=MUTED), al=LEFT)
    row_line(ws, rr, 2, 9)
ws.conditional_formatting.add("B20:I49", FormulaRule(formula=["$B20=Zinsbindung"], border=Border(bottom=side(TEAL))))
note(ws, "B51", "Vereinfachung: Zinsen auf die Restschuld zu Jahresbeginn; bei monatlicher Zahlung fallen die Zinsen etwas geringer aus. "
                "Die Rate bleibt nach der Zinsbindung unverändert.")
footer(ws, 53)
ws.freeze_panes = "A20"
printing(ws, "A1:M53", repeat_rows="19:19")

# =========================================================================== STEUERN
ws = S["Steuern"]
setup(ws, "Steuern", "ber")
section(ws, 8, 2, 7, "AfA-Bemessungsgrundlage", "€")
afa = [
    ("Kaufpreis Immobilie (ohne bewegliche Gegenstände)", "=Kaufpreis-Inventar", None),
    ("Grunderwerbsteuer", "=Finanzierung!E10", None),
    ("Notar, Grundbuch, Makler (anteilig)", "=SUM(Finanzierung!E11:E13)*(Kaufpreis-Inventar)/Kaufpreis", None),
    ("Anschaffungskosten Immobilie", "=SUM(E9:E11)", "AK_Immobilie"),
    ("davon Grund und Boden (nicht abschreibbar)", "=AK_Immobilie*(1-Gebaeudeanteil)", None),
    ("davon Gebäude", "=AK_Immobilie*Gebaeudeanteil", "Gebaeude_AK"),
    ("zzgl. aktivierte Renovierung (bei Überschreiten der 15-%-Grenze)", "=IF(Regel15_ok=1,0,Renovierung)", None),
    ("AfA-Bemessungsgrundlage Gebäude", "=Gebaeude_AK+E15", "AfA_Basis"),
    ("AfA Gebäude p. a.", "=AfA_Basis*AfA_Satz", "AfA_Gebaeude"),
    ("Anschaffungskosten Inventar (inkl. anteiliger Nebenkosten)", "=Inventar+SUM(Finanzierung!E11:E13)*Inventar/Kaufpreis", "Inventar_AK"),
    ("AfA Inventar p. a.", "=Inventar_AK/ND_Inventar", "AfA_Inventar"),
]
for i, (lab, frm, nm) in enumerate(afa):
    rr = 9 + i
    merge(ws, f"B{rr}:D{rr}", lab)
    put(ws, f"E{rr}", frm, fmt=EUR)
    row_line(ws, rr, 2, 7)
    if nm:
        name(nm, ws, f"E{rr}")
sum_row(ws, 12, 2, 7)
sum_row(ws, 16, 2, 7)
sum_row(ws, 17, 2, 7, res=True)
sum_row(ws, 19, 2, 7, res=True)

section(ws, 8, 9, 13, "15-%-Grenze", "§ 6 Abs. 1 Nr. 1a EStG")
g15 = [
    ("Anschaffungskosten Gebäude", "=Gebaeude_AK", EUR, None),
    ("Grenze: 15 % der Gebäude-AK", "=Gebaeude_AK*0.15", EUR, "Grenze15"),
    ("Renovierung netto (ersten 3 Jahre)", "=Renovierung", EUR, None),
    ("Grenze eingehalten?", "=IF(Renovierung<=Grenze15,1,0)", JA_NEIN, "Regel15_ok"),
]
for i, (lab, frm, fmt, nm) in enumerate(g15):
    rr = 9 + i
    merge(ws, f"I{rr}:L{rr}", lab)
    put(ws, f"M{rr}", frm, fmt=fmt)
    row_line(ws, rr, 9, 13)
    if nm:
        name(nm, ws, f"M{rr}")
ws["M12"].font = font(10, True)
ws.conditional_formatting.add("M12", CellIsRule(operator="equal", formula=["1"], font=Font(color=GRN, bold=True)))
ws.conditional_formatting.add("M12", CellIsRule(operator="equal", formula=["0"], font=Font(color=RED, bold=True)))
merge(ws, "I13:M14", '=IF(Regel15_ok=1,"Sofort abziehbarer Erhaltungsaufwand (ggf. verteilt nach § 82b EStDV).",'
                     '"Anschaffungsnahe Herstellungskosten – nur über die Gebäude-AfA absetzbar.")',
      f=font(9, color=MUTED), al=Alignment(wrap_text=True, vertical="top"))

section(ws, 15, 9, 13, "Steuersatz")
tax = [
    ("Grenzsteuersatz Einkommensteuer", "=Grenzsteuersatz", PCT, None),
    ("Solidaritätszuschlag (auf ESt)", "=Soli", PCT, None),
    ("Kirchensteuer (auf ESt)", "=KiSt", PCT, None),
    ("Effektiver Steuersatz", "=Grenzsteuersatz*(1+Soli+KiSt)", PCT2, "Steuersatz_eff"),
]
for i, (lab, frm, fmt, nm) in enumerate(tax):
    rr = 16 + i
    merge(ws, f"I{rr}:L{rr}", lab)
    put(ws, f"M{rr}", frm, fmt=fmt)
    row_line(ws, rr, 9, 13)
    if nm:
        name(nm, ws, f"M{rr}")
sum_row(ws, 19, 9, 13, res=True)

section(ws, 21, 2, 12, "Einkünfte aus Vermietung und Verpachtung", "€ p. a. · Erstattung positiv")
table_header(ws, 22, [("B22", "Jahr", "center"), ("C22", "Mieteinnahmen", "right"), ("D22", "Bewirtschaftung abziehbar", "right"),
                      ("E22", "Schuldzinsen", "right"), ("F22", "AfA Gebäude", "right"), ("G22", "AfA Inventar", "right"),
                      ("H22", "Erhaltungsaufwand", "right"), ("I22", "Einkünfte V+V", "right"), ("J22", "Steuereffekt", "right"),
                      ("K22", "Restbuchwert Gebäude", "right"), ("L22", "Summe AfA", "right")], dark=True)
ws.row_dimensions[22].height = 30
for t in range(1, 31):
    rr = 22 + t
    pr = 10 + t  # Zeile in Projektion
    fr = 19 + t  # Zeile im Tilgungsplan
    put(ws, f"B{rr}", t, fmt=YEAR_F, al=CENTER)
    put(ws, f"C{rr}", f"=Projektion!E{pr}", fmt=NUM)
    put(ws, f"D{rr}", f"=-(Hausgeld_nu+Verwaltung)*(1+Kostensteigerung)^(B{rr}-1)", fmt=NUM)
    put(ws, f"E{rr}", f"=-Finanzierung!E{fr}", fmt=NUM)
    put(ws, f"F{rr}", f"=-MIN(AfA_Gebaeude,AfA_Basis+SUM(F$22:F{rr - 1}))", fmt=NUM)
    put(ws, f"G{rr}", f"=-IF(B{rr}<=ND_Inventar,AfA_Inventar,0)", fmt=NUM)
    put(ws, f"H{rr}", f"=-IF(AND(Regel15_ok=1,B{rr}<=Erhaltung_Jahre),Renovierung/Erhaltung_Jahre,0)", fmt=NUM)
    put(ws, f"I{rr}", f"=SUM(C{rr}:H{rr})", fmt=NUM, f=font(10, True))
    put(ws, f"J{rr}", f"=-I{rr}*Steuersatz_eff", fmt=NUM, f=font(10, True, TEAL))
    put(ws, f"K{rr}", f"=AfA_Basis+SUM(F$23:F{rr})", fmt=NUM)
    put(ws, f"L{rr}", f"=-(F{rr}+G{rr})", fmt=NUM)
    row_line(ws, rr, 2, 12)
note(ws, "B54", "Annahmen: Verluste werden vollständig mit anderen Einkünften verrechnet; Steuersatz vereinfacht als "
                "Grenzsteuersatz × (1 + Soli + Kirchensteuer). Instandhaltungsrücklage erst bei Verwendung abziehbar.")
footer(ws, 56)
ws.freeze_panes = "A23"
printing(ws, "A1:M56", repeat_rows="22:22")

# =========================================================================== PROJEKTION
ws = S["Projektion"]
setup(ws, "Projektion", "ber", extra_cols=7)
for col in col_range(3, 20):
    ws.column_dimensions[col].width = 14.5
section(ws, 8, 2, 20, "Jahresprojektion", "€ p. a. · Jahr 1 = 12 Monate ab Kaufdatum")
cols = [("B", "Jahr"), ("C", "Nettokaltmiete Soll"), ("D", "Mietausfall"), ("E", "Nettokaltmiete Ist"),
        ("F", "Bewirtschaftung"), ("G", "Zinsen"), ("H", "Tilgung"), ("I", "Kapitaldienst"), ("J", "Cashflow vor Steuern"),
        ("K", "Steuereffekt"), ("L", "Cashflow nach Steuern"), ("M", "CF n. St. kumuliert"), ("N", "Abschreibungen gesamt"),
        ("O", "Immobilienwert Jahresende"), ("P", "Restschuld Jahresende"), ("Q", "Nettovermögen"), ("R", "DSCR"),
        ("S", "Netto-Exit-Erlös"), ("T", "Zahlungsstrom Eigenkapital")]
table_header(ws, 9, [(f"{c}9", t, "center" if c == "B" else "right") for c, t in cols], dark=True)
ws.row_dimensions[9].height = 30
# Jahr 0 (Kauf)
put(ws, "B10", 0, fmt='0', al=CENTER, f=font(10, color=MUTED))
put(ws, "O10", "=Marktwert", fmt=NUM)
put(ws, "P10", "=Darlehen", fmt=NUM)
put(ws, "Q10", "=O10-P10", fmt=NUM)
put(ws, "T10", "=-Eigenkapital", fmt=NUM, f=font(10, True))
style_range(ws, "B10:T10", bg=NAV_BG)
for t in range(1, 31):
    rr = 10 + t
    fr = 19 + t
    sr = 22 + t
    put(ws, f"B{rr}", t, fmt=YEAR_F, al=CENTER)
    put(ws, f"C{rr}", f"=Miete_Monat*12*(1+Mietsteigerung)^(B{rr}-1)", fmt=NUM)
    put(ws, f"D{rr}", f"=-C{rr}*Mietausfall", fmt=NUM)
    put(ws, f"E{rr}", f"=C{rr}+D{rr}", fmt=NUM)
    put(ws, f"F{rr}", f"=-(Hausgeld_nu+Ruecklage+Verwaltung)*(1+Kostensteigerung)^(B{rr}-1)", fmt=NUM)
    put(ws, f"G{rr}", f"=-Finanzierung!E{fr}", fmt=NUM)
    put(ws, f"H{rr}", f"=-Finanzierung!F{fr}", fmt=NUM)
    put(ws, f"I{rr}", f"=G{rr}+H{rr}", fmt=NUM)
    put(ws, f"J{rr}", f"=E{rr}+F{rr}+I{rr}", fmt=NUM, f=font(10, True), bg=SUM_BG)
    put(ws, f"K{rr}", f"=Steuern!J{sr}", fmt=NUM)
    put(ws, f"L{rr}", f"=J{rr}+K{rr}", fmt=NUM, f=font(10, True, TEAL), bg=TEAL_L)
    put(ws, f"M{rr}", f"=SUM(L$11:L{rr})", fmt=NUM)
    put(ws, f"N{rr}", f"=Steuern!L{sr}", fmt=NUM)
    put(ws, f"O{rr}", f"=Marktwert*(1+Wertsteigerung)^B{rr}", fmt=NUM)
    put(ws, f"P{rr}", f"=Finanzierung!H{fr}", fmt=NUM)
    put(ws, f"Q{rr}", f"=O{rr}-P{rr}", fmt=NUM, f=font(10, True), bg=SUM_BG)
    put(ws, f"R{rr}", f'=IF(I{rr}=0,"",(E{rr}+F{rr})/-I{rr})', fmt=DSCR_F)
    put(ws, f"S{rr}", f"=IF(B{rr}=Haltedauer,Exit_Netto,0)", fmt=NUM)
    put(ws, f"T{rr}", f"=IF(B{rr}<=Haltedauer,L{rr}+S{rr},0)", fmt=NUM, f=font(10, True))
    row_line(ws, rr, 2, 20)
# Verkaufsjahr hervorheben, negative Cashflows rot
ws.conditional_formatting.add("B11:T40", FormulaRule(formula=["$B11=Haltedauer"], border=Border(top=side(TEAL), bottom=side(TEAL)),
                                                     fill=fill("FFF7E0")))
ws.conditional_formatting.add("R11:R40", FormulaRule(formula=['AND(ISNUMBER(R11),R11<Ziel_DSCR)'], font=Font(color=RED)))
ws.conditional_formatting.add("J11:J40", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED)))
ws.conditional_formatting.add("L11:L40", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))
note(ws, "B41", "Gelb markiert: Verkaufsjahr (Haltedauer). Vorzeichen: Einnahmen positiv, Ausgaben negativ, Steuererstattung positiv.")

section(ws, 43, 2, 7, "Exit-Rechnung", "Verkauf am Ende der Haltedauer")
ex = [
    ("Verkaufsjahr", "=Haltedauer", '0', None),
    ("Immobilienwert im Verkaufsjahr", "=INDEX(O11:O40,Haltedauer)", EUR, None),
    ("Verkaufskosten", "=-F45*Verkaufskosten", EUR, None),
    ("Ablösung Restschuld", "=-INDEX(P11:P40,Haltedauer)", EUR, None),
    ("Steuerlicher Restbuchwert", "=AK_Immobilie+Inventar_AK+IF(Regel15_ok=1,0,Renovierung)-SUMPRODUCT((B11:B40<=Haltedauer)*N11:N40)", EUR, None),
    ("Veräußerungsgewinn", "=F45+F46-F48", EUR, None),
    ("Steuer auf Veräußerungsgewinn (nur bis 10 Jahre)", "=-IF(Haltedauer<=10,MAX(0,F49)*Steuersatz_eff,0)", EUR, None),
    ("Netto-Exit-Erlös", "=F45+F46+F47+F50", EUR, "Exit_Netto"),
]
for i, (lab, frm, fmt, nm) in enumerate(ex):
    rr = 44 + i
    merge(ws, f"B{rr}:E{rr}", lab)
    put(ws, f"F{rr}", frm, fmt=fmt)
    row_line(ws, rr, 2, 7)
    if nm:
        name(nm, ws, f"F{rr}")
sum_row(ws, 51, 2, 7, res=True)
note(ws, "B52", "Nicht berücksichtigt: Vorfälligkeitsentschädigung bei Verkauf während der Zinsbindung.")

section(ws, 43, 9, 14, "Rendite-Kennzahlen")
kpis = [
    ("Bruttomietrendite (Sollmiete / Kaufpreis)", "=Miete_Monat*12/Kaufpreis", PCT, "Brutto_Rendite"),
    ("Nettomietrendite (Reinertrag J1 / Gesamtinvestition)", "=(E11+F11)/Gesamtinvestition", PCT, "Netto_Rendite"),
    ("Kapitaldienstdeckung Jahr 1", "=R11", DSCR_F, "DSCR_J1"),
    ("Cashflow n. St. / Monat – Jahr 1", "=L11/12", EUR, "CF_Monat_J1"),
    ("Cashflow n. St. / Monat – Jahr 2", "=L12/12", EUR, "CF_Monat_J2"),
    ("EK-Rendite Jahr 1 (inkl. Tilgung und Wertzuwachs)", "=(L11-H11+O11-O10)/Eigenkapital", PCT, "EK_Rendite_J1"),
    ("IRR nach Steuern", "=IRR(T10:T40,0.05)", PCT, "IRR_nSt"),
    ("Eigenkapital-Multiple", '=SUMIF(T10:T40,">0")/-SUMIF(T10:T40,"<0")', MULT, "Multiple"),
]
for i, (lab, frm, fmt, nm) in enumerate(kpis):
    rr = 44 + i
    merge(ws, f"I{rr}:M{rr}", lab)
    put(ws, f"N{rr}", frm, fmt=fmt, f=font(10, True))
    row_line(ws, rr, 9, 14)
    name(nm, ws, f"N{rr}")

section(ws, 54, 2, 7, "Diagrammdaten Cockpit", "Cashflow Jahr 1 · € je Monat")
chart_rows = [("Miete", "=E11/12"), ("Bewirt.", "=F11/12"), ("Zinsen", "=G11/12"), ("Tilgung", "=H11/12"),
              ("vor St.", "=J11/12"), ("Steuer", "=K11/12"), ("nach St.", "=L11/12")]
for i, (lab, frm) in enumerate(chart_rows):
    rr = 55 + i
    put(ws, f"B{rr}", lab)
    put(ws, f"C{rr}", frm, fmt=NUM)
    row_line(ws, rr, 2, 3)
footer(ws, 63, c2=20)
ws.freeze_panes = "C10"
printing(ws, "A1:T63", repeat_rows="9:9")

# =========================================================================== AFA-VERGLEICH
ws = S["AfA-Vergleich"]
setup(ws, "AfA-Vergleich", "ber")
section(ws, 8, 2, 11, "Vergleich der Abschreibungsmethoden", "Bemessungsgrundlage: Gebäude")
table_header(ws, 9, [("B9:C9", "", "left"), ("D9", "Linear 2,0 %", "right"), ("E9", "Linear 2,5 %", "right"),
                     ("F9", "Linear 3,0 %", "right"), ("G9", "Degressiv 5,0 %", "right"), ("H9:K9", "Hinweis", "left")], dark=True)
merge(ws, "B10:C10", "AfA-Satz")
for col, v in zip("DEFG", [0.02, 0.025, 0.03, 0.05]):
    put(ws, f"{col}10", v, fmt=PCT)
merge(ws, "H10:K10", "Gesetzliche Sätze nach § 7 Abs. 4 und 5a EStG", f=font(9, color=MUTED))
merge(ws, "B11:C11", "Voraussetzung")
for col, txt in zip("DEFG", ["Fertigst. 1925–2022", "Fertigst. vor 1925", "Fertigst. ab 2023", "Neubau, Baubeginn 10/2023–9/2029"]):
    put(ws, f"{col}11", txt, f=font(8.5, color=MUTED), al=Alignment(horizontal="right", vertical="center", wrap_text=True))
ws.row_dimensions[11].height = 26
merge(ws, "H11:K11", "Degressiv: jährlich 5 % vom Restbuchwert, Wechsel zu linear möglich (hier nicht modelliert)",
      f=font(9, color=MUTED), al=WRAP)
summ = [
    ("AfA Jahr 1", "={c}20", EUR),
    ("Summe AfA bis Verkauf", "=SUMPRODUCT(($B$20:$B$49<=Haltedauer)*{c}$20:{c}$49)", EUR),
    ("Steuerersparnis bis Verkauf", "={c}13*Steuersatz_eff", EUR),
    ("Restbuchwert bei Verkauf", "=AfA_Basis-{c}13", EUR),
    ("In „Eingaben“ gewählt", '=IF(ROUND(AfA_Satz,4)=ROUND({c}10,4),"✓ gewählt","")', "@"),
]
for i, (lab, frm, fmt) in enumerate(summ):
    rr = 12 + i
    merge(ws, f"B{rr}:C{rr}", lab)
    for col in "DEFG":
        put(ws, f"{col}{rr}", frm.format(c=col), fmt=fmt)
    row_line(ws, rr, 2, 11)
sum_row(ws, 14, 2, 7, res=True)
for col in "DEFG":
    ws[f"{col}16"].font = font(10, True, GRN)
note(ws, "H13", "Höhere AfA erhöht beim Verkauf innerhalb von 10 Jahren den steuerpflichtigen Gewinn.")

section(ws, 18, 2, 11, "Jahreswerte", "€ p. a.")
table_header(ws, 19, [("B19:C19", "Jahr", "center"), ("D19", "AfA 2,0 %", "right"), ("E19", "AfA 2,5 %", "right"),
                      ("F19", "AfA 3,0 %", "right"), ("G19", "AfA degr. 5 %", "right"), ("H19", "RBW 2,0 %", "right"),
                      ("I19", "RBW 2,5 %", "right"), ("J19", "RBW 3,0 %", "right"), ("K19", "RBW degr.", "right")], dark=True)
for t in range(1, 31):
    rr = 19 + t
    ws.merge_cells(f"B{rr}:C{rr}")
    put(ws, f"B{rr}", t, fmt=YEAR_F, al=CENTER)
    for col in "DEF":
        put(ws, f"{col}{rr}", f"=MIN({col}$10*AfA_Basis,AfA_Basis-SUM({col}$19:{col}{rr - 1}))", fmt=NUM)
    put(ws, f"G{rr}", f"=(AfA_Basis-SUM(G$19:G{rr - 1}))*G$10", fmt=NUM)
    for src, dst in zip("DEFG", "HIJK"):
        put(ws, f"{dst}{rr}", f"=AfA_Basis-SUM({src}$20:{src}{rr})", fmt=NUM, f=font(10, color=MUTED))
    row_line(ws, rr, 2, 11)
ws.conditional_formatting.add("B20:K49", FormulaRule(formula=["$B20=Haltedauer"], fill=fill("FFF7E0")))
note(ws, "B50", "RBW = Restbuchwert. Gelb markiert: Verkaufsjahr.")
footer(ws, 52)
ws.freeze_panes = "A20"
printing(ws, "A1:M52", repeat_rows="19:19")

# =========================================================================== SENSITIVITÄT
ws = S["Sensitivität"]
setup(ws, "Sensitivität", "ber")
section(ws, 8, 2, 7, "Basiswerte Jahr 1", "aus Eingaben und Steuern")
base = [
    ("Jahres-Sollmiete", "=Miete_Monat*12"),
    ("Bewirtschaftung gesamt (liquiditätswirksam)", "=Hausgeld_nu+Ruecklage+Verwaltung"),
    ("davon steuerlich abziehbar", "=Hausgeld_nu+Verwaltung"),
    ("AfA und Erhaltungsaufwand Jahr 1", "=-(Steuern!F23+Steuern!G23+Steuern!H23)"),
]
for i, (lab, frm) in enumerate(base):
    rr = 9 + i
    merge(ws, f"B{rr}:E{rr}", lab)
    put(ws, f"F{rr}", frm, fmt=EUR)
    row_line(ws, rr, 2, 7)
note(ws, "I9", "Blau: Variationsschritte – frei änderbar.")
note(ws, "I10", "Mittelfeld = Basisfall (entspricht Projektion Jahr 1).")

M_, BW, BA, AX = "$F$9", "$F$10", "$F$11", "$F$12"


def grid(top, title, right, kind):
    section(ws, top, 2, 9, title, right)
    merge(ws, f"D{top + 1}:H{top + 1}", "Änderung Nettokaltmiete", f=font(9, True, TEAL), bg=TEAL_L, al=CENTER)
    put(ws, f"B{top + 2}", "Sollzins", f=font(9, True, TEAL), bg=TEAL_L, al=CENTER)
    put(ws, f"C{top + 2}", "Δ Zins", f=font(9, True, TEAL), bg=TEAL_L, al=CENTER)
    for j, dm in enumerate([-0.10, -0.05, 0, 0.05, 0.10]):
        input_cell(ws, f"{get_column_letter(4 + j)}{top + 2}", dm, '+0 %;-0 %;0 %', al=CENTER)
    for i, dz in enumerate([-0.01, -0.005, 0, 0.005, 0.01]):
        rr = top + 3 + i
        put(ws, f"B{rr}", f"=Zins+C{rr}", fmt=PCT2, f=font(10, color=MUTED), al=CENTER)
        input_cell(ws, f"C{rr}", dz, '+0.00 %;-0.00 %;0.00 %', al=CENTER)
        for j in range(5):
            col = get_column_letter(4 + j)
            mi = f"({M_}*(1+{col}${top + 2})*(1-Mietausfall))"
            zi = f"(Darlehen*(Zins+$C{rr}))"
            ti = "(Darlehen*Tilgung)"
            if kind == "cfn":
                frm = f"=({mi}-{BW}-{zi}-{ti}-({mi}-{BA}-{zi}-{AX})*Steuersatz_eff)/12"
                fmt = EUR
            elif kind == "cfv":
                frm = f"=({mi}-{BW}-{zi}-{ti})/12"
                fmt = EUR
            else:
                frm = f"=({mi}-{BW})/({zi}+{ti})"
                fmt = DSCR_F
            c = put(ws, f"{col}{rr}", frm, fmt=fmt, al=CENTER,
                    border=Border(bottom=side(LINE), right=side(LINE)))
            if i == 2 and j == 2:
                c.font = font(10, True)
                c.border = Border(left=side(TEAL, "medium"), right=side(TEAL, "medium"),
                                  top=side(TEAL, "medium"), bottom=side(TEAL, "medium"))
    rng = f"D{top + 3}:H{top + 7}"
    if kind == "dscr":
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"D{top + 3}<Ziel_DSCR*Schwelle_Pruefen"], fill=fill(RED_BG), font=Font(color=RED)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"D{top + 3}<Ziel_DSCR"], fill=fill(AMB_BG), font=Font(color=AMB)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"D{top + 3}>=Ziel_DSCR"], fill=fill(GRN_BG), font=Font(color=GRN)))
    else:
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_BG), font=Font(color=RED)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=fill(GRN_BG), font=Font(color=GRN)))


grid(15, "Cashflow nach Steuern je Monat – Jahr 1", "€ / Monat", "cfn")
grid(25, "Cashflow vor Steuern je Monat – Jahr 1", "€ / Monat", "cfv")
grid(35, "Kapitaldienstdeckung (DSCR) – Jahr 1", "Ziel laut Konfiguration", "dscr")
note(ws, "B44", "Tilgungssatz bleibt konstant, die Rate ändert sich mit dem Zins. Steuereffekt mit effektivem Steuersatz laut Blatt „Steuern“.")
footer(ws, 46)
printing(ws, "A1:M46")

# =========================================================================== BANKGESPRÄCH
ws = S["Bankgespräch"]
setup(ws, "Bankgespräch", "bank", meta=True)


def block(top, title, rows, right=None):
    section(ws, top, 2, 13, title, right)
    for i, row in enumerate(rows):
        lab, frm, fmt = row[:3]
        hint = row[3] if len(row) > 3 else ""
        rr = top + 1 + i
        merge(ws, f"B{rr}:E{rr}", lab)
        put(ws, f"F{rr}", frm, fmt=fmt, f=font(10, True), al=LEFT if fmt == "@" else RIGHT)
        if hint:
            merge(ws, f"G{rr}:M{rr}", hint, f=font(9, color=MUTED), al=Alignment(horizontal="left", vertical="center", indent=1))
        row_line(ws, rr, 2, 13)
    return top + 1 + len(rows)


end = block(8, "Objekt", [
    ("Objekt", "=Objekt", "@"),
    ("Lage", "=Lage", "@"),
    ("Wohnfläche", "=Wohnflaeche", '0.0 "m²"'),
    ("Kaufpreis", "=Kaufpreis", EUR, '="Kaufpreis je m²: "&FIXED(Kaufpreis/Wohnflaeche,0)&" €"'),
    ("Marktwert (Annahme)", "=Marktwert", EUR),
    ("Nettokaltmiete Soll / Monat", "=Miete_Monat", EUR, '="Miete je m²: "&FIXED(Miete_Monat/Wohnflaeche,2)&" €"'),
])
end = block(end + 1, "Finanzierung", [
    ("Gesamtinvestition", "=Gesamtinvestition", EUR, '="davon Kaufnebenkosten "&FIXED(Kaufnebenkosten,0)&" €"'),
    ("Eigenkapital", "=Eigenkapital", EUR, '="Eigenkapitalquote "&FIXED(EK_Quote*100,1)&" %"'),
    ("Darlehen", "=Darlehen", EUR),
    ("Beleihungsauslauf (Darlehen / Marktwert)", "=LTV", PCT),
    ("Sollzins / anfängliche Tilgung", '=FIXED(Zins*100,2)&" % / "&FIXED(Tilgung*100,2)&" %"', "@"),
    ("Monatsrate", "=Monatsrate", EUR),
    ("Zinsbindung", "=Zinsbindung", '0 "Jahre"'),
    ("Restschuld Ende Zinsbindung", "=Restschuld_ZB", EUR, '="Anschlusszins angenommen: "&FIXED(Anschlusszins*100,2)&" %"'),
])
kd_top = end + 1
end = block(kd_top, "Kapitaldienstfähigkeit des Objekts", [
    ("Nettokaltmiete Ist p. a. (Jahr 1)", "=Projektion!E11", EUR),
    ("Bewirtschaftung p. a.", "=Projektion!F11", EUR),
    ("Reinertrag p. a.", "=Projektion!E11+Projektion!F11", EUR),
    ("Kapitaldienst p. a.", "=-Projektion!I11", EUR),
    ("Kapitaldienstdeckung (DSCR)", "=DSCR_J1", DSCR_F, '=IF(DSCR_J1>=Ziel_DSCR,"erfüllt – Bankmaßstab ≥ "&FIXED(Ziel_DSCR,2),"unter Bankmaßstab ≥ "&FIXED(Ziel_DSCR,2)&" – Objekt trägt sich nicht selbst")'),
])
sum_row(ws, kd_top + 3, 2, 13)
dscr_row = kd_top + 5
ws.conditional_formatting.add(f"F{dscr_row}:M{dscr_row}", FormulaRule(formula=[f"$F${dscr_row}<Ziel_DSCR"], font=Font(color=RED, bold=True)))
ws.conditional_formatting.add(f"F{dscr_row}:M{dscr_row}", FormulaRule(formula=[f"$F${dscr_row}>=Ziel_DSCR"], font=Font(color=GRN, bold=True)))
hh_top = end + 1
end = block(hh_top, "Haushaltsrechnung", [
    ("Nettoeinkommen Haushalt", "=Netto_Einkommen", EUR),
    ("+ Mieteinnahmen (Bankansatz)", "=Miete_Monat*Bank_Mietansatz", EUR, '="Bank rechnet "&FIXED(Bank_Mietansatz*100,0)&" % der Sollmiete an"'),
    ("− Lebenshaltungskosten", "=-Lebenshaltung", EUR),
    ("− Eigene Wohnkosten", "=-Wohnkosten", EUR),
    ("− Bestehende Kreditraten", "=-Kreditraten", EUR),
    ("− Bewirtschaftung Objekt", "=-(Hausgeld_nu+Ruecklage+Verwaltung)/12", EUR),
    ("− Rate neues Darlehen", "=-Monatsrate", EUR),
    ("Monatlicher Überschuss", f"=SUM(F{hh_top + 1}:F{hh_top + 7})", EUR,
     f'=IF(F{hh_top + 8}>=0,"tragfähig – ohne Berücksichtigung von Steuererstattungen","nicht tragfähig – Haushalt deckt die Belastung nicht")'),
    ("Liquide Reserve nach Kauf", "=Reserve", EUR, f'="entspricht "&FIXED(Reserve/Monatsrate,1)&" Monatsraten (Empfehlung ≥ "&Min_Reserve&")"'),
], right="€ / Monat")
res_row = hh_top + 8
sum_row(ws, res_row, 2, 13, res=True)
ws.conditional_formatting.add(f"F{res_row}:M{res_row}", FormulaRule(formula=[f"$F${res_row}<0"], font=Font(color=RED, bold=True)))
note(ws, f"B{end + 1}", "Haushaltswerte sind Beispielwerte (Blatt „Eingaben“, Abschnitt Haushalt). Banken rechnen mit eigenen Pauschalen.")
footer(ws, end + 3)
printing(ws, f"A1:M{end + 3}", landscape=False)

# =========================================================================== COCKPIT
ws = S["Cockpit"]
setup(ws, "Cockpit", "aus", meta=True)
for rr in range(7, 56):
    ws.row_dimensions[rr].height = 16.5
ws.row_dimensions[7].height = 8
# Gesamtbewertung
ws.row_dimensions[8].height = 17
ws.row_dimensions[9].height = 17
merge(ws, "B8:C9", "Gesamtbewertung", f=font(9, True, RED))
merge(ws, "D8:E9", '=IF(OR(F20="kritisch",F21="kritisch"),"Liquidität kritisch",IF(COUNTIF(F18:F23,"kritisch")>0,"Rendite kritisch",'
                   'IF(COUNTIF(F18:F23,"prüfen")>0,"Mit Einschränkungen","Solide")))',
      fmt='"● "@', f=font(12, True), al=LEFT)
merge(ws, "F8:M9", '="Kapitaldienstdeckung "&FIXED(DSCR_J1,2)&IF(DSCR_J1<Ziel_DSCR," statt mindestens "," bei Ziel ")&FIXED(Ziel_DSCR,2)'
                   '&". Cashflow nach Steuern im Jahr 1: "&FIXED(CF_Monat_J1,0)&" € / Monat, im Jahr 2: "&FIXED(CF_Monat_J2,0)&" € / Monat. '
                   'IRR nach Steuern "&FIXED(IRR_nSt*100,1)&" %."',
      f=font(9.5, color="3A3F45"), al=WRAP)
for word, fg, bg_, line in (("kritisch", RED, RED_BG, "E7B4AD"), ("Einschränkungen", AMB, AMB_BG, "F1CFA5"), ("Solide", GRN, GRN_BG, "B5DCC4")):
    ws.conditional_formatting.add("B8:M9", FormulaRule(formula=[f'ISNUMBER(SEARCH("{word}",$D$8))'], fill=fill(bg_),
                                                       border=Border(top=side(line), bottom=side(line))))
    ws.conditional_formatting.add("B8:E9", FormulaRule(formula=[f'ISNUMBER(SEARCH("{word}",$D$8))'], font=Font(color=fg, bold=True)))

# KPI-Kacheln
tiles = [
    ("Gesamtinvestition", "=Gesamtinvestition", EUR, '="Kaufpreis "&FIXED(Kaufpreis,0)&" €"'),
    ("Eigenkapitalbedarf", "=Eigenkapital", EUR, '="Quote "&FIXED(EK_Quote*100,1)&" % · LTV "&FIXED(LTV*100,1)&" %"'),
    ("Cashflow n. St. / Monat", "=CF_Monat_J1", EUR, '="Jahr 2: "&FIXED(CF_Monat_J2,0)&" €"'),
    ("Bruttomietrendite", "=Brutto_Rendite", PCT, '="Faktor "&FIXED(Faktor,1)&"× · netto "&FIXED(Netto_Rendite*100,1)&" %"'),
    ("Kapitaldienstdeckung", "=DSCR_J1", DSCR_F, '="Bankmaßstab ≥ "&FIXED(Ziel_DSCR,2)'),
    ("IRR nach Steuern", "=IRR_nSt", PCT, '=Haltedauer&" Jahre · Multiple "&FIXED(Multiple,1)&"×"'),
]
ws.row_dimensions[11].height = 18
ws.row_dimensions[12].height = 17
ws.row_dimensions[13].height = 17
ws.row_dimensions[14].height = 16
gap = side(WHITE, "thick")
for i, (lab, frm, fmt, sub) in enumerate(tiles):
    a, b = get_column_letter(2 + 2 * i), get_column_letter(3 + 2 * i)
    style_range(ws, f"{a}11:{b}14", bg=TEAL_XL)
    for rr in range(11, 15):
        for col in (a, b):
            ws[f"{col}{rr}"].border = Border(top=side(TEAL, "thick") if rr == 11 else None, right=gap if col == b else None)
    merge(ws, f"{a}11:{b}11", lab, f=font(8.5, True, "3A3F45"), bg=TEAL_XL, al=Alignment(horizontal="left", vertical="bottom", indent=1))
    merge(ws, f"{a}12:{b}13", frm, fmt=fmt, f=font(20, True), bg=TEAL_XL, al=Alignment(horizontal="left", vertical="center", indent=1))
    merge(ws, f"{a}14:{b}14", sub, f=font(8, color=MUTED), bg=TEAL_XL, al=Alignment(horizontal="left", vertical="top", indent=1, shrink_to_fit=True))
# Ampelfarbe für Kacheln (Kapitaldienstdeckung, Cashflow, IRR)
for anchor, cond in (("J12", "DSCR_J1<Ziel_DSCR"), ("F12", "CF_Monat_J1<Ziel_CF"), ("L12", "IRR_nSt<Ziel_IRR*Schwelle_Pruefen")):
    ws.conditional_formatting.add(anchor, FormulaRule(formula=[cond], font=Font(color=RED, bold=True)))

# Kennzahlen-Check
section(ws, 16, 2, 7, "Kennzahlen-Check", "Ist gegen Ziel")
table_header(ws, 17, [("B17:C17", "Kennzahl", "left"), ("D17", "Ist", "right"), ("E17", "Ziel", "right"),
                      ("F17", "Status", "left"), ("G17", "Erreichung", "left")])
checks = [
    ("Bruttomietrendite", "=Brutto_Rendite", "=Ziel_Brutto", PCT),
    ("Nettomietrendite", "=Netto_Rendite", "=Ziel_Netto", PCT),
    ("Kapitaldienstdeckung", "=DSCR_J1", "=Ziel_DSCR", DSCR_F),
    ("Cashflow n. St. / Monat", "=CF_Monat_J1", "=Ziel_CF", '#,##0 "€";-#,##0 "€";0 "€"'),
    ("EK-Rendite Jahr 1", "=EK_Rendite_J1", "=Ziel_EKR", PCT),
    ("IRR nach Steuern", "=IRR_nSt", "=Ziel_IRR", PCT),
]
for i, (lab, ist, ziel, fmt) in enumerate(checks):
    rr = 18 + i
    merge(ws, f"B{rr}:C{rr}", lab)
    put(ws, f"D{rr}", ist, fmt=fmt, f=font(10, True))
    put(ws, f"E{rr}", ziel, fmt=fmt, f=font(10, color=MUTED2))
    put(ws, f"F{rr}", f'=IF(D{rr}>=E{rr},"erfüllt",IF(AND(E{rr}>0,D{rr}>=E{rr}*Schwelle_Pruefen),"prüfen","kritisch"))',
        fmt='"● "@', f=font(9, True), al=Alignment(horizontal="left", vertical="center", indent=1))
    put(ws, f"G{rr}", f'=REPT("█",ROUND(8*MIN(IF(E{rr}>0,MAX(0,D{rr}/E{rr}),IF(D{rr}>=E{rr},1.25,0)),1.25),0))',
        f=font(7), al=LEFT)
    row_line(ws, rr, 2, 7)
    status_cf(ws, f"F{rr}:G{rr}", f"$F{rr}")
note(ws, "B24", "Balken: 8 Blöcke = Zielwert erreicht · Ziele und Ampelschwellen: Blatt Konfiguration")
ws["B24"].hyperlink = Hyperlink(ref="B24", location="'Konfiguration'!A1")

# Cashflow Jahr 1 (Diagramm)
section(ws, 16, 9, 13, "Cashflow Jahr 1", "€ je Monat")
bar = BarChart()
bar.type = "col"
bar.style = 10
bar.gapWidth = 45
bar.legend = None
bar.y_axis.numFmt = "#,##0"
bar.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=LINE))
bar.y_axis.delete = False
bar.x_axis.delete = False
bar.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
bar.x_axis.tickLblPos = "low"
data = Reference(S["Projektion"], min_col=3, min_row=55, max_row=61)
cats = Reference(S["Projektion"], min_col=2, min_row=55, max_row=61)
bar.add_data(data, titles_from_data=False)
bar.set_categories(cats)
ser = bar.series[0]
ser.graphicalProperties.solidFill = TEAL
ser.graphicalProperties.line.noFill = True
for idx, color in enumerate([TEAL, GREY_BAR, GREY_BAR, GREY_BAR, RED, ACC, TEAL]):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = color
    pt.graphicalProperties.line.noFill = True
    ser.dPt.append(pt)
ser.dLbls = DataLabelList()
ser.dLbls.showVal = True
ser.dLbls.showSerName = False
ser.dLbls.showCatName = False
ser.dLbls.showLegendKey = False
ser.dLbls.numFmt = "#,##0"
ser.dLbls.position = "outEnd"
bar.width = 12.4
bar.height = 6.0
ws.add_chart(bar, "I17")

# Vermögensentwicklung (Diagramm)
section(ws, 27, 2, 7, "Vermögensentwicklung", "€ · Jahre 1–30")
line = LineChart()
line.style = 12
P = S["Projektion"]
for col, color, width, dash, title in ((15, BLUE, 22000, None, "Immobilienwert"), (16, MUTED2, 16000, "dash", "Restschuld"),
                                        (17, TEAL, 32000, None, "Nettovermögen")):
    ref = Reference(P, min_col=col, min_row=11, max_row=40)
    line.add_data(ref, titles_from_data=False)
    s = line.series[-1]
    from openpyxl.chart.series import SeriesLabel
    s.tx = SeriesLabel(v=title)
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = width
    if dash:
        s.graphicalProperties.line.dashStyle = dash
    s.smooth = False
line.set_categories(Reference(P, min_col=2, min_row=11, max_row=40))
line.y_axis.numFmt = '#,##0,"T€"'
line.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=LINE))
line.y_axis.delete = False
line.x_axis.delete = False
line.x_axis.tickLblSkip = 5
line.x_axis.tickMarkSkip = 5
line.legend.position = "b"
line.width = 15.2
line.height = 6.4
ws.add_chart(line, "B28")

# Prüfhinweise
section(ws, 27, 9, 13, "Prüfhinweise")
put(ws, "M27", '=COUNTIF(I28:I37,"●*")&" Warnung(en)"', f=font(8.5, color=MUTED2), al=Alignment(horizontal="right", vertical="bottom"))
hints = [
    '=IF(DSCR_J1<Ziel_DSCR,"● DSCR im Jahr 1 nur "&FIXED(DSCR_J1,2)&" – Banken erwarten meist 1,10 bis 1,20.",'
    '"○ DSCR im Jahr 1 "&FIXED(DSCR_J1,2)&" – Bankmaßstab erfüllt.")',
    '=IF(Zinsbindung<Haltedauer,"● ","○ ")&"Zinsbindung endet nach "&Zinsbindung&" Jahren, Restschuld "&FIXED(Restschuld_ZB,0)'
    '&" € – Anschlusszins "&FIXED(Anschlusszins*100,2)&" % angenommen."',
    '="○ Steuerliches Ergebnis Jahr 1 "&FIXED(Steuern!I23,0)&" € – "&IF(Steuern!I23<0,"Verrechnung mit anderen Einkünften.","zu versteuern.")',
    '=IF(Regel15_ok=1,"○ Renovierung "&FIXED(Renovierung,0)&" € unter der 15-%-Grenze ("&FIXED(Grenze15,0)&" €) – sofort abziehbar.",'
    '"● Renovierung über der 15-%-Grenze ("&FIXED(Grenze15,0)&" €) – nur über die AfA absetzbar.")',
    '=IF(Haltedauer<=10,"● Verkauf innerhalb von 10 Jahren – Veräußerungsgewinn steuerpflichtig (§ 23 EStG).",'
    '"○ Bewegliche Gegenstände "&FIXED(Inventar,0)&" € – grunderwerbsteuerfrei, "&ND_Inventar&" Jahre AfA.")',
]
for i, frm in enumerate(hints):
    rr = 28 + 2 * i
    merge(ws, f"I{rr}:M{rr + 1}", frm, f=font(9), al=Alignment(wrap_text=True, vertical="center"))
    for col in "IJKLM":
        ws[f"{col}{rr + 1}"].border = Border(bottom=side(LINE))
    ws.conditional_formatting.add(f"I{rr}", FormulaRule(formula=[f'LEFT(I{rr},1)="●"'], font=Font(color=AMB)))

# Kennzahlen im Zeitverlauf
section(ws, 39, 2, 13, "Kennzahlen im Zeitverlauf", "€ p. a.")
years = [1, 2, 3, 5, 10, 15, 20, 25, 30]
ws.merge_cells("B40:D40")
put(ws, "B40", "Jahr", f=font(9, True, WHITE), bg=TEAL, al=LEFT)
style_range(ws, "B40:M40", bg=TEAL)
for j, y in enumerate(years):
    put(ws, f"{get_column_letter(5 + j)}40", y, fmt='0', f=font(9, True, WHITE), bg=TEAL, al=RIGHT)
zv = [
    ("row", "Nettokaltmiete Ist", "E"), ("row", "Bewirtschaftungskosten", "F"), ("row", "Kapitaldienst", "I"),
    ("sum", "Cashflow vor Steuern", "J"), ("row", "Steuereffekt", "K"), ("res", "Cashflow nach Steuern", "L"),
    ("row", "Kumulierter Cashflow nach Steuern", "M"), ("row", "Abschreibungen gesamt", "N"),
    ("row", "Immobilienwert Jahresende", "O"), ("row", "Restschuld Jahresende", "P"), ("sum", "Nettovermögen", "Q"),
]
for i, (kind, lab, src) in enumerate(zv):
    rr = 41 + i
    merge(ws, f"B{rr}:D{rr}", lab)
    for j in range(9):
        col = get_column_letter(5 + j)
        put(ws, f"{col}{rr}", f"=INDEX(Projektion!${src}$11:${src}$40,{col}$40)", fmt=NUM)
    row_line(ws, rr, 2, 13)
    if kind != "row":
        sum_row(ws, rr, 2, 13, res=(kind == "res"))
ws.conditional_formatting.add("E44:M44", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))
ws.conditional_formatting.add("E46:M46", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))
note(ws, "B52", "Jahr 1 = 12 Monate ab Kaufdatum · Steuereffekt: Erstattung positiv, Zahlung negativ · Details: Blatt Projektion")
footer(ws, 54)
printing(ws, "A1:M54", one_page=True)

# =========================================================================== START
ws = S["Start"]
setup(ws, "Immobilien-Kalkulation", "ein", meta=True)
section(ws, 8, 2, 7, "Objekt")
obj = [("Objekt", "=Objekt", "@"), ("Lage", "=Lage", "@"), ("Kaufpreis", "=Kaufpreis", EUR),
       ("Gesamtinvestition", "=Gesamtinvestition", EUR), ("Kaufdatum", "=Kaufdatum", DATE), ("Haltedauer", "=Haltedauer", '0 "Jahre"')]
for i, (lab, frm, fmt) in enumerate(obj):
    rr = 9 + i
    merge(ws, f"B{rr}:C{rr}", lab, f=font(10, color=MUTED))
    merge(ws, f"D{rr}:F{rr}", frm, fmt=fmt, f=font(10, True), al=LEFT)
    row_line(ws, rr, 2, 7)
section(ws, 8, 9, 13, "Ergebnis auf einen Blick")
res = [("Gesamtbewertung", "=Cockpit!D8", '"● "@'), ("Cashflow n. St. / Monat", "=CF_Monat_J1", EUR),
       ("Kapitaldienstdeckung", "=DSCR_J1", DSCR_F), ("IRR nach Steuern", "=IRR_nSt", PCT),
       ("Eigenkapitalbedarf", "=Eigenkapital", EUR), ("Nettovermögen Jahr 30", "=Projektion!Q40", EUR)]
for i, (lab, frm, fmt) in enumerate(res):
    rr = 9 + i
    merge(ws, f"I{rr}:J{rr}", lab, f=font(10, color=MUTED))
    merge(ws, f"K{rr}:M{rr}", frm, fmt=fmt, f=font(10, True), al=RIGHT)
    row_line(ws, rr, 9, 13)
for word, fg in (("kritisch", RED), ("Einschränkungen", AMB), ("Solide", GRN)):
    ws.conditional_formatting.add("K9", FormulaRule(formula=[f'ISNUMBER(SEARCH("{word}",$K$9))'], font=Font(color=fg, bold=True)))

section(ws, 16, 2, 13, "Inhalt", "Klick auf den Blattnamen öffnet das Blatt")
table_header(ws, 17, [("B17:C17", "Blatt", "left"), ("D17:E17", "Bereich", "left"), ("F17:M17", "Inhalt", "left")])
GROUP_LABEL = {"ein": "Einstieg", "inp": "Eingabe", "aus": "Ausgabe", "ber": "Berechnung", "bank": "Bank", "anh": "Anhang"}
for i, (nm, grp, desc) in enumerate(SHEETS):
    rr = 18 + i
    merge(ws, f"B{rr}:C{rr}", nm, f=font(10, True, TEAL, underline="single"))
    ws[f"B{rr}"].hyperlink = Hyperlink(ref=f"B{rr}", location=f"'{nm}'!A1", display=nm)
    merge(ws, f"D{rr}:E{rr}", GROUP_LABEL[grp], f=font(9, True, WHITE if grp in ("inp", "aus", "ber", "bank") else INK),
          bg=GROUP[grp], al=Alignment(horizontal="left", vertical="center", indent=1))
    merge(ws, f"F{rr}:M{rr}", desc, f=font(9.5, color=MUTED))
    row_line(ws, rr, 2, 13)

section(ws, 31, 2, 13, "Legende")
input_cell(ws, "B32", "12.345", al=CENTER)
merge(ws, "C32:G32", "Eingabe – nur diese Felder ändern (Blatt Eingaben, Konfiguration, Sensitivität)", f=font(9.5))
put(ws, "B33", 12345, fmt=NUM)
merge(ws, "C33:G33", "Berechnung – Formel, nicht überschreiben", f=font(9.5))
put(ws, "B34", 12345, fmt=NUM)
sum_row(ws, 34, 2, 2)
merge(ws, "C34:G34", "Zwischensumme", f=font(9.5))
put(ws, "B35", 12345, fmt=NUM)
sum_row(ws, 35, 2, 2, res=True)
merge(ws, "C35:G35", "Ergebnis", f=font(9.5))
put(ws, "I32", "erfüllt", fmt='"● "@', f=font(9.5, True, GRN))
merge(ws, "J32:M32", "Ist-Wert erreicht den Zielwert", f=font(9.5))
put(ws, "I33", "prüfen", fmt='"● "@', f=font(9.5, True, AMB))
merge(ws, "J33:M33", "Ist-Wert knapp unter Ziel (ab Schwelle laut Konfiguration)", f=font(9.5))
put(ws, "I34", "kritisch", fmt='"● "@', f=font(9.5, True, RED))
merge(ws, "J34:M34", "Ist-Wert deutlich unter Ziel", f=font(9.5))
footer(ws, 37)
printing(ws, "A1:M37")

# =========================================================================== LEITFADEN
ws = S["Leitfaden"]
setup(ws, "Leitfaden", "ein")
section(ws, 8, 2, 13, "In sechs Schritten zur Bewertung")
steps = [
    ("Eingaben", "Alle blau hinterlegten Felder auf dem Blatt „Eingaben“ prüfen und mit den Daten des Objekts überschreiben – "
                 "Kaufpreis, Nebenkosten, Finanzierungsangebot, Miete, Bewirtschaftung und persönlicher Steuersatz."),
    ("Cockpit", "Gesamtbewertung, Kennzahlen-Kacheln und den Kennzahlen-Check lesen. Die Ampel zeigt sofort, ob Liquidität "
                "und Rendite den Zielwerten entsprechen."),
    ("Prüfhinweise", "Die Hinweise im Cockpit erklären kritische Punkte: Kapitaldienstdeckung, Ende der Zinsbindung, "
                     "15-%-Grenze bei Renovierung und Spekulationsfrist."),
    ("Details", "Projektion, Finanzierung und Steuern zeigen jede Zahl Jahr für Jahr. Das Verkaufsjahr ist gelb markiert."),
    ("Szenarien", "Auf „Sensitivität“ wirken sich Zins- und Mietänderungen sofort auf Cashflow und DSCR aus. "
                  "Auf „AfA-Vergleich“ werden die Abschreibungsmethoden verglichen."),
    ("Bank", "Das Blatt „Bankgespräch“ ist druckfertig (A4 hoch) und fasst Objekt, Finanzierung, Kapitaldienstfähigkeit "
             "und Haushaltsrechnung zusammen."),
]
for i, (head, txt) in enumerate(steps):
    rr = 9 + 2 * i
    ws.row_dimensions[rr].height = 17
    ws.row_dimensions[rr + 1].height = 17
    merge(ws, f"B{rr}:B{rr + 1}", i + 1, f=font(20, True, TEAL), al=CENTER)
    put(ws, f"C{rr}", head, f=font(10, True, TEAL), al=Alignment(vertical="bottom"))
    merge(ws, f"D{rr}:M{rr + 1}", txt, f=font(9.5, color="3A3F45"), al=WRAP)
    row_line(ws, rr + 1, 2, 13)
section(ws, 22, 2, 13, "Konventionen")
conv = [
    "Jahr 1 umfasst die ersten 12 Monate ab Kaufdatum; alle Werte sind Jahreswerte, Cockpit-Kacheln teils je Monat.",
    "Vorzeichen: Einnahmen positiv, Ausgaben negativ. Steuereffekt: Erstattung positiv, Zahlung negativ.",
    "Zielwerte, Ampelschwellen und Bankparameter werden zentral auf „Konfiguration“ gepflegt.",
    "Wichtige Größen sind als Namen hinterlegt (z. B. =Gesamtinvestition, =Eigenkapital, =DSCR_J1) – Formeln bleiben lesbar.",
    "Schutz vor versehentlichem Überschreiben: Überprüfen → Blatt schützen. Eingabefelder sind bereits entsperrt.",
]
for i, txt in enumerate(conv):
    rr = 23 + i
    put(ws, f"B{rr}", "–", f=font(10, color=TEAL), al=CENTER)
    merge(ws, f"C{rr}:M{rr}", txt, f=font(9.5, color="3A3F45"))
footer(ws, 30)
printing(ws, "A1:M30")

# =========================================================================== HINWEISE
ws = S["Hinweise"]
setup(ws, "Hinweise", "anh")
blocks = [
    ("Annahmen und Vereinfachungen", [
        "Jährliche Betrachtung; Zinsen auf die Restschuld zu Jahresbeginn (monatliche Zahlweise leicht günstiger).",
        "Annuität bleibt nach Ende der Zinsbindung konstant; Anschlusszins ist eine Annahme.",
        "Steuereffekt = Einkünfte aus V+V × Grenzsteuersatz × (1 + Soli + Kirchensteuer); volle Verlustverrechnung unterstellt.",
        "Instandhaltungsrücklage mindert die Liquidität, ist steuerlich aber erst bei Verwendung abziehbar.",
        "Kaufnebenkosten werden anteilig auf Immobilie und bewegliche Gegenstände verteilt.",
        "Verkauf zum Marktwert am Ende der Haltedauer abzüglich Verkaufskosten; keine Vorfälligkeitsentschädigung.",
    ]),
    ("Rechtsgrundlagen", [
        "§ 7 Abs. 4 EStG – lineare Gebäude-AfA 2,0 % / 2,5 % / 3,0 %; § 7 Abs. 5a EStG – degressive AfA 5 % für Neubauten.",
        "§ 6 Abs. 1 Nr. 1a EStG – anschaffungsnahe Herstellungskosten (15-%-Grenze innerhalb von 3 Jahren, netto).",
        "§ 82b EStDV – Verteilung größeren Erhaltungsaufwands auf 2 bis 5 Jahre.",
        "§ 23 Abs. 1 Nr. 1 EStG – private Veräußerungsgeschäfte, Spekulationsfrist 10 Jahre.",
        "§ 1 und § 11 GrEStG, Landesrecht Baden-Württemberg – Grunderwerbsteuer 5,0 %; bewegliche Gegenstände nicht steuerbar.",
    ]),
    ("Haftungsausschluss", [
        "Diese Kalkulation ist ein Planungsinstrument. Sie ersetzt keine Rechts-, Steuer- oder Finanzberatung.",
        "Alle Angaben ohne Gewähr; Rechtsstand September 2026. Ergebnisse hängen vollständig von den Eingaben ab.",
    ]),
    ("Herausgeber", [
        "MM Holding GmbH · Kornhausgasse 4 · 88250 Weingarten · HRB 750729",
    ]),
]
rr = 7
for title, items in blocks:
    rr += 1
    section(ws, rr, 2, 13, title)
    for txt in items:
        rr += 1
        put(ws, f"B{rr}", "–", f=font(10, color=TEAL), al=CENTER)
        merge(ws, f"C{rr}:M{rr}", txt, f=font(9.5, color="3A3F45"))
        row_line(ws, rr, 3, 13)
    rr += 1
footer(ws, rr + 1)
printing(ws, f"A1:M{rr + 1}")

# =========================================================================== Abschluss
wb.active = NAV.index("Cockpit")
for w in wb.worksheets:
    w.sheet_view.tabSelected = w.title == "Cockpit"
wb.properties.title = "Immobilien-Kalkulation – Cockpit"
wb.properties.creator = "MM Holding GmbH"
wb.properties.subject = "Investitions-, Finanzierungs- und Steuerrechnung für eine vermietete Eigentumswohnung"
wb.calculation.fullCalcOnLoad = True

out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("Immobilien-Kalkulation.xlsx")
wb.save(out)
print(f"gespeichert: {out}")
