"""Blatt „Start“: Hero, Entscheidungskarte „① Kauf als“, vier Kennzahl-Kacheln, Ablauf ①–⑥, Legende mit
Tipp-Kasten, Blätterverzeichnis, Fuß – ausschließlich mit den zentralen Bausteinen aus core.py.

Raster (px): C 126 | D 112 | E 238 | F 238 | G 238  →  C:D = E = F = G (vier gleich breite Spalten).
  Hero:  Logo-Spalte C:D, Inhalt E:G; Kennzahlen und Buttons exakt in E | F | G (Drittelraster = Kachelspalten 2–4).
  Kacheln Z. 27–29: C:D | E | F | G, Zeile 29 = Kontextzeile mit Status („●  prüfen  ·  Ziel ≥ 5,0 %“).
Namensziele bleiben an ihrem Platz: Rechtsform = D23, ST_BMR = E28, ST_CF = C31, ST_DSCR = E31, ST_IRR = F31.
Die zweite Kachelreihe (Z. 30–32, Namensziele ST_CF/ST_DSCR/ST_IRR) wird ausgeblendet (P3-12: Start zeigt vier
Kacheln); ihre Formeln bleiben unverändert. Geändert werden nur Darstellung, statische Beschriftungen und nicht
referenzierte Anzeigeformeln.
"""
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Border, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C
from core import (ACCENT, AMBER, BLUE, GREEN, INK, INK2, MIST, MUTED, NAVY, RED, SKY, WHITE, align, fill, font,
                  side)

ui = C                                   # Altname (leitfaden.py importiert Helfer aus diesem Modul)
SHEET = "Start"
HERO_ROWS = range(5, 20)                 # B5:H19 Navy-Fläche, Z. 20 Weißraum
HERO_COLS = "BCDEFGH"
WIDTHS = {"C": 18, "D": 16, "E": 34, "F": 34, "G": 34}
EMU = 9525                               # EMU je Pixel
KAUF_ALS = 'CHOOSE(Rechtsform_Idx,"Privatperson","vv-GmbH","GmbH / Holding")'

# Ablauf „So gehen Sie vor“: (Titel, Zielblatt, Zielzelle, Beschreibung ≤ 95 Zeichen, einzeilig)
FLOW = [
    ("Kauf als wählen", "Start", "D23",
     "Privatperson oder Gesellschaft – die Auswahl oben steuert Steuern, Exit und Bankunterlagen."),
    ("Leitfaden: 12 Schritte", "Leitfaden", None,
     "Vom Objekt bis zum Ergebnis – jede Seite mit Eingaben, Zwischenergebnis und Einordnung."),
    ("Dashboard", "Dashboard", None,
     "Die Gesamtbewertung auf einer Seite: Kennzahlen mit Status, Cashflow und Vermögen."),
    ("Cockpit & Diagramme", "Cockpit", None,
     "Detailkennzahlen und Prüfhinweise im Cockpit, alle Auswertungen als Diagramm."),
    ("Sensitivität", "Sensitivität", None,
     "Break-even-Miete und -Zins, Miete × Zins und IRR-Matrix – wie robust ist die Rechnung?"),
    ("Bankunterlagen", "Bankgespräch", None,
     "Bankgespräch als A4-Übersicht, dazu Haushaltsrechnung und Vermögensaufstellung."),
]
CIRCLED = "①②③④⑤⑥"

# Blätterverzeichnis in Reiterreihenfolge; „{n}“ = Anzahl Diagramme
SHEETS = [
    ("Leitfaden", "Zwölf Schritte vom Objekt zum Ergebnis"),
    ("Dashboard", "Gesamtbewertung auf einer Seite"),
    ("Cockpit", "Detailkennzahlen und Prüfhinweise"),
    ("Diagramme", "{n} Diagramme zur Kalkulation"),
    ("Eingaben", "Alle Eingaben und Profi-Felder"),
    ("Projektion", "Miete, Kosten, Vermögen – 40 Jahre"),
    ("Steuern", "Steuerparameter, AfA, Ergebnis, Exit"),
    ("AfA-Vergleich", "Abschreibungsvarianten im Vergleich"),
    ("Finanzierung", "Tilgungspläne Darlehen I und II"),
    ("Sensitivität", "Break-even, Miete × Zins, IRR-Matrix"),
    ("Bankgespräch", "Investitionsübersicht für die Bank"),
    ("Haushaltsrechnung", "Selbstauskunft Einnahmen/Ausgaben"),
    ("Vermögensaufstellung", "Selbstauskunft Vermögen"),
    ("Hinweise", "Rechtsgrundlagen und Modellannahmen"),
    ("Konfiguration", "Stammdaten, Tarif 2026, Ampel-Schwellen"),
]


# ============================================================================ lokale Helfer (auch für leitfaden.py)
def rich(*parts):
    """Rich-Text aus (Text, Größe, fett, Farbe)-Teilen – dünne Hülle um core.rich."""
    return C.rich([tuple(p[:4]) for p in parts])


def unmerge_in(ws, c1, r1, c2, r2):
    c1, c2 = C.col(c1), C.col(c2)
    for mr in list(ws.merged_cells.ranges):
        if not (mr.max_row < r1 or mr.min_row > r2 or mr.max_col < c1 or mr.min_col > c2):
            ws.unmerge_cells(str(mr))


def drop_cf(ws, coords):
    """Bedingte Formate entfernen, deren Bereich vollständig in den genannten Zellen liegt."""
    targets = set(coords)
    new = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        cells = {C.L(c) + str(r) for cr in cf.sqref.ranges for (r, c) in cr.cells}
        if cells and cells <= targets:
            continue
        for rule in cf.rules:
            new.add(str(cf.sqref), rule)
    ws.conditional_formatting = new


def cells_of(c1, r1, c2, r2):
    return [f"{C.L(c)}{r}" for r in range(r1, r2 + 1) for c in range(C.col(c1), C.col(c2) + 1)]


def wipe(ws, c1, r1, c2, r2, formulas=False):
    """Bereich leeren: Verbünde lösen, statische Werte, Links und Stile entfernen.
    Formeln bleiben, außer formulas=True (nur für Bereiche ohne referenzierte Formeln)."""
    unmerge_in(ws, c1, r1, c2, r2)
    for c in C.iter_cells(ws, c1, r1, c2, r2):
        if isinstance(c, MergedCell):
            continue
        c.hyperlink = None
        if formulas or not C.is_formula(c.value):
            c.value = None
        c.fill = C.NOFILL
        c.border = Border()
        c.font = font()
        c.alignment = align("general", "center")
        c.number_format = "General"


def link(cell, text, sheet, target=None, size=C.T_BODY, bold=True, color=BLUE, tooltip=None, h="left", indent=1):
    """Textlink (nur in Listen/Fließtext – nie in einer Button-Reihe)."""
    C.text_link(cell, text, sheet, target_cell=target, size=size, bold=bold,
                tooltip=tooltip or f"Zum Blatt „{sheet}“")
    cell.hyperlink.location = C.link_loc(sheet, target)
    cell.font = font(size, bold, color)
    cell.alignment = align(h, "center", indent)


def set_widths(ws, widths):
    """Spaltenbreiten setzen; Bereichs-Dimensionen der Vorlage (z. B. D:G in einem <col>) vorher auftrennen."""
    from openpyxl.worksheet.dimensions import ColumnDimension
    for key, d in list(ws.column_dimensions.items()):
        if d.min and d.max and d.max > d.min:
            for idx in range(d.min + 1, min(d.max, 60) + 1):
                letter = C.L(idx)
                if letter not in ws.column_dimensions or ws.column_dimensions[letter].min != idx:
                    ws.column_dimensions[letter] = ColumnDimension(ws, index=letter, width=d.width,
                                                                   customWidth=d.customWidth, hidden=d.hidden)
            d.max = d.min
    for k, w in widths.items():
        dim = ws.column_dimensions[k]
        dim.width = w
        dim.min = dim.max = C.col(k)


def heights(ws, spec):
    for r, h in spec.items():
        C.set_height(ws, r, h)


def status_line(kpi, value_ref, tail, conditions=None):
    """Kontextzeile einer Kachel mit Status: ="●  prüfen  ·  "&<tail>  (Farbe per C.tile(status='dot')).
    tail: Excel-Ausdruck (ohne „=“) für den Kontext, z. B. '"Ziel ≥ "&FIXED(Ampel_BMR_gruen*100,1)&" %"'."""
    conds = C.status_conditions(kpi, value_ref, conditions)
    return C.status_formula(conds) + '&"  ·  "&' + tail


def goal_pct(name):
    return f'"Ziel ≥ "&FIXED({name}*100,1)&" %"'


def clear_rows(ws, r1, r2, c1="B", c2="H"):
    """Nur Stile (keine Formeln) in ausgeblendeten Zeilen neutralisieren."""
    unmerge_in(ws, c1, r1, c2, r2)
    for c in C.iter_cells(ws, c1, r1, c2, r2):
        if isinstance(c, MergedCell):
            continue
        c.fill = C.NOFILL
        c.border = Border()


# ============================================================================ „Kauf als“ (Logik bleibt)
def purchase_selector(wb):
    """Direktauswahl Privat / Kapitalgesellschaft auf der Startseite (steuert den Namen „Rechtsform“)."""
    ws, s09 = wb[SHEET], wb["S09 Steuern"]
    current = s09["D12"].value
    unmerge_in(ws, "D", 23, "H", 23)
    for c in "EFGH":                                   # D23:G23 wird ein Feld – Pfeil ▾ und Link in G23 entfallen
        ws[f"{c}23"].value = None
        ws[f"{c}23"].hyperlink = None
    ws.merge_cells("D23:G23")
    sel = ws["D23"]
    sel.value = current
    sel.protection = Protection(locked=False)
    dv = DataValidation(type="list", formula1="=L_Rechtsform", allow_blank=False, showDropDown=False,
                        showErrorMessage=True, errorStyle="stop", showInputMessage=True)
    dv.promptTitle = "Kauf als"
    dv.prompt = "Privatperson oder Kapitalgesellschaft (vermögensverwaltende bzw. gewerbliche GmbH) aus der Liste wählen."
    dv.errorTitle = "Bitte aus der Liste wählen"
    dv.error = "Privatperson, vermögensverwaltende GmbH oder gewerbliche GmbH/Holding über den Pfeil am Feldrand wählen."
    ws.add_data_validation(dv)
    dv.add("D23")
    wb.defined_names["Rechtsform"] = DefinedName("Rechtsform", attr_text="Start!$D$23")
    # Schritt 09 zeigt die Auswahl nur noch an (verknüpfter Wert mit Link zur Startseite)
    for dvs in list(s09.data_validations.dataValidation):
        if "D12" in str(dvs.sqref):
            s09.data_validations.dataValidation.remove(dvs)
    d12 = s09["D12"]
    d12.value = "=Rechtsform"
    C.input_style(d12, "linked")
    d12.protection = Protection(locked=True)
    d12.hyperlink = Hyperlink(ref="D12", location=C.link_loc("Start", "D23"), display="Kauf als – Startseite",
                              tooltip="Kauf als auf der Startseite ändern")
    s09["F12"].value = "Kauf als auf der Startseite ändern  ›"
    s09["F12"].data_type = "s"


# ============================================================================ Hero
def logo(ws, top_px=3, size_px=140):
    """Logo als OneCellAnchor mit fester Größe: links bündig in der Logo-Spalte C:D (Kante C + Einzug),
    vertikal über Eyebrow … Untertitel (Z. 6–11). Feste Größe → unabhängig von Zeilenhöhen, kein Verzerren."""
    for img in ws._images:
        a = getattr(img.anchor, "_from", None)
        if a is None or a.col > 3 or not (4 <= a.row <= 12):
            continue
        marker = AnchorMarker(col=2, colOff=12 * EMU, row=5, rowOff=top_px * EMU)
        img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(size_px * EMU, size_px * EMU))
        img.width = img.height = size_px


def hero(ws):
    unmerge_in(ws, "B", 5, "H", 21)
    for r in range(5, 22):
        for c in HERO_COLS:
            cell = ws[f"{c}{r}"]
            cell.fill = fill(NAVY) if r in HERO_ROWS else C.NOFILL
            cell.border = Border()
            cell.hyperlink = None
    texts = {                                              # Inhalte aus D (Vorlage/Runde 1) nach E übernehmen
        "eyebrow": ws["D6"].value, "title": ws["D9"].value, "sub": ws["D11"].value,
    }
    for r in range(5, 22):                                 # alles leeren – danach gezielt neu setzen
        for c in HERO_COLS:
            cell = ws[f"{c}{r}"]
            if cell.coordinate != "G7":                    # Eingabe „Erstellt für“ bleibt
                cell.value = None

    def put(cell, value, size, bold, color, h="left", v="center", merge_to=None, fmt=None):
        if merge_to:
            C.safe_merge(ws, cell[0], int(cell[1:]), merge_to, int(cell[1:]))
        c = ws[cell]
        if isinstance(value, str) and not value.startswith("="):
            C.set_text(c, value)
        else:
            c.value = value
        c.font = font(size, bold, color)
        c.alignment = align(h, v, 1 if h == "left" else 0)
        if fmt:
            c.number_format = fmt
        return c

    put("E6", texts["eyebrow"] or "MM HOLDING GMBH   ·   UNTERNEHMERISCHE BETEILIGUNGEN   ·   WEINGARTEN",
        C.T_MICRO, True, SKY, merge_to="F")
    put("E7", "=Obj_Name", C.T_BODY, False, MIST, merge_to="F")
    created_for(ws)
    put("E9", texts["title"] or "Immobilien-Kalkulation", C.T_HERO, True, WHITE, merge_to="G")
    put("E11", texts["sub"] or "Kauf · Finanzierung · Cashflow · Steuern · Exit", C.T_BODY, False, MIST, merge_to="G")
    put("E12", '="Stand "&TEXT(DAY(TODAY()),"00")&"."&TEXT(MONTH(TODAY()),"00")&"."&YEAR(TODAY())'
               f'&"   ·   Rechtsstand September 2026   ·   Kauf als: "&{KAUF_ALS}',
        C.T_MICRO, False, SKY, merge_to="G")

    # Kennzahlenband im Drittelraster E | F | G (= Kachelspalten 2–4)
    band = [("E", "GESAMTINVESTITION", "=Gesamtinvestition", C.NUMFMT["eur"]),
            ("F", "KAUFPREIS", "=Kaufpreis", C.NUMFMT["eur"]),
            ("G", "HALTEDAUER", "=Haltedauer", C.NUMFMT["years_n"])]
    for c, lab, val, fmt in band:
        put(f"{c}14", lab, C.T_MICRO, True, SKY, v="bottom")
        put(f"{c}15", val, C.T_H2, True, WHITE, fmt=fmt)
    for c in "EFG":
        ws[f"{c}14"].border = Border(top=side("hair", ACCENT))

    # Aktionszeile im selben Raster: Primär (weiß auf Navy) · Ghost · Ghost, 3-px-Fugen in Navy
    p = C.btn(ws, "E", 17, "E", "Leitfaden starten  ›", "Leitfaden", "primary",
              tooltip="Zur Übersicht der zwölf Schritte")
    p.font = font(C.T_BODY, True, NAVY)
    ws["E17"].fill = fill(WHITE)
    C.btn(ws, "F", 17, "F", "Dashboard  ›", "Dashboard", "ghost", tooltip="Gesamtbewertung auf einer Seite")
    C.btn(ws, "G", 17, "G", "Cockpit  ›", "Cockpit", "ghost", tooltip="Detailkennzahlen und Prüfhinweise")
    gap = side("thick", NAVY)
    edge = side("thin", ACCENT)
    ws["E17"].border = Border(top=side("thin", WHITE), bottom=side("thin", WHITE), left=side("thin", WHITE), right=gap)
    ws["F17"].border = Border(top=edge, bottom=edge, left=gap, right=gap)
    ws["G17"].border = Border(top=edge, bottom=edge, left=gap, right=edge)
    put("E18", "↓  Zuerst unten: Kauf als Privat oder GmbH wählen", C.T_MICRO, False, MIST, h="center")

    for c in HERO_COLS:                                                 # Abschluss: Akzentlinie
        ws[f"{c}19"].border = Border(bottom=side("thick", ACCENT))
    heights(ws, {4: 15, 5: 16, 6: 14, 7: 18, 8: 12, 9: 42, 10: 6, 11: 18, 12: 14, 13: 12, 14: 20, 15: 30,
                 16: 16, 17: C.H_BTN, 18: 22, 19: 10, 20: 12})
    logo(ws)


# ============================================================================ Erstellt für (Deckblatt-Feld)
def created_for(ws):
    """Eingabefeld „Erstellt für“ oben rechts im Hero + Name Erstellt_fuer (Start!$G$7).
    Gedämpft (P2-15): FFF5D6, Rahmen 1 pt E6CB77, 10 pt, Zeilenhöhe 18."""
    wb = ws.parent
    lab, inp = ws["G6"], ws["G7"]
    C.set_text(lab, "ERSTELLT FÜR")
    lab.font = font(C.T_MICRO, True, SKY)
    lab.alignment = align("left", "center", 1)
    if inp.value is None:
        C.set_text(inp, "Max Mustermann")
    C.input_style(inp, "required")
    inp.font = font(C.T_BODY, True, C.INPUT_FG)
    inp.alignment = align("left", "center", 1)
    dv = DataValidation(type="textLength", operator="lessThanOrEqual", formula1="60", allow_blank=True,
                        showErrorMessage=True, errorStyle="stop", showInputMessage=True)
    dv.promptTitle = "Erstellt für"
    dv.prompt = "Name oder Firma des Investors – erscheint auf Dashboard, Bankgespräch und den Schrittseiten."
    dv.errorTitle = "Zu lang"
    dv.error = "Bitte höchstens 60 Zeichen eingeben."
    ws.add_data_validation(dv)
    dv.add("G7")
    wb.defined_names["Erstellt_fuer"] = DefinedName("Erstellt_fuer", attr_text="Start!$G$7")


# ============================================================================ ① Kauf als – Entscheidungskarte
def selector(ws):
    unmerge_in(ws, "C", 21, "H", 22)
    C.section(ws, 21, "C", "G", "①  Kauf als – Privatperson oder Gesellschaft")
    lab = ws["C23"]
    lab.value = rich(("KAUF ALS", C.T_BODY, True, NAVY), ("\nPflichtauswahl", C.T_MICRO, False, BLUE))
    lab.font = font(C.T_BODY, True, NAVY)
    lab.alignment = align("left", "center", 1, wrap=True)
    lab.border = Border()
    lab.fill = C.NOFILL
    frame = side("medium", BLUE)
    for c in "DEFG":
        cell = ws[f"{c}23"]
        cell.fill = fill(C.INPUT_BG)
        cell.border = Border(top=frame, bottom=frame, left=side("thick", ACCENT) if c == "D" else None,
                             right=frame if c == "G" else None)
    for c in "H":
        ws[f"{c}23"].fill = C.NOFILL
        ws[f"{c}23"].border = Border()
    sel = ws["D23"]
    sel.font = font(C.T_BODY, True, C.INPUT_FG)
    sel.alignment = align("left", "center", 1)

    wipe(ws, "C", 24, "H", 24)
    C.safe_merge(ws, "D", 24, "G", 24)
    C.set_text(ws["D24"], "Wirkt auf: S09 Steuern  ·  S11 Exit  ·  Dashboard  ·  Bankgespräch")
    ws["D24"].font = font(C.T_MICRO, False, MUTED)
    ws["D24"].alignment = align("left", "center", 1)
    for c in HERO_COLS:
        for r in (20, 22):
            ws[f"{c}{r}"].fill = C.NOFILL
            ws[f"{c}{r}"].border = Border()
    heights(ws, {20: 12, 22: 8, 23: 36, 24: 18, 25: C.H_GAP})


# ============================================================================ Kacheln
def tiles(ws):
    """Vier Kacheln (C.tile dark), Reihenfolge wie Dashboard: Cashflow · Bruttomietrendite · DSCR · IRR.
    Wert neutral navy (negativer Cashflow rot), Status als „●  Wort“ in der Kontextzeile (Z. 29)."""
    unmerge_in(ws, "C", 26, "G", 26)
    C.section(ws, 26, "C", "G", "Aktuelle Kalkulation im Überblick", meta="Details im Dashboard  ›")
    m = ws["G26"]
    m.hyperlink = Hyperlink(ref="G26", location=C.link_loc("Dashboard"), display="Dashboard",
                            tooltip="Gesamtbewertung auf einer Seite")
    for r in range(27, 33):
        for c in "BCDEFGH":
            ws[f"{c}{r}"].border = Border()
            ws[f"{c}{r}"].fill = C.NOFILL
    drop_cf(ws, cells_of("C", 27, "G", 32))
    unmerge_in(ws, "C", 27, "G", 32)
    for c in "CDEFG":                                   # Kontextzeile (Z. 29) ist in der Vorlage leer
        ws[f"{c}29"].value = None
    ws["D27"].value = ws["D28"].value = None
    cf = C.CF_YEAR2
    C.tile(ws, "C", "D", 27, 28, 29, kpi="CF", label="Cashflow n. St. / Monat (Jahr 1)", value="=CF_nSt_Monat_J1",
           value_ref="CF_nSt_Monat_J1", status="dot", sub=status_line("CF", "CF_nSt_Monat_J1", f'"Jahr 2: "&FIXED({cf},0)&" € / Monat"'))
    C.tile(ws, "E", "E", 27, 28, 29, kpi="BMR", status="dot",
           sub=status_line("BMR", "Bruttomietrendite", goal_pct("Ampel_BMR_gruen")))
    C.tile(ws, "F", "F", 27, 28, 29, kpi="DSCR", label="DSCR (Jahr 1)", value="=DSCR_J1", value_ref="DSCR_J1", status="dot",
           sub=status_line("DSCR", "DSCR_J1", '"Ziel ≥ "&FIXED(Ampel_DSCR_gruen,2)&"×"'))
    C.tile(ws, "G", "G", 27, 28, 29, kpi="IRR", value="=EK_IRR", value_ref="EK_IRR", status="dot", gap_right=False,
           sub=status_line("IRR", "EK_IRR", goal_pct("Ampel_IRR_gruen")))
    g = ws["G27"]
    g.value = '="IRR N. ST. · VERKAUF NACH "&Haltedauer&" J."'
    g.data_type = "f"
    # zweite Reihe der Vorlage (Namensziele ST_CF/ST_DSCR/ST_IRR) ausblenden – Formeln bleiben unverändert
    clear_rows(ws, 30, 32)
    for c, fmt in (("C", C.NUMFMT["eur"]), ("E", C.NUMFMT["dscr"]), ("F", C.NUMFMT["pct1"])):
        ws[f"{c}31"].number_format = fmt
    C.hide_rows(ws, 30, 32)
    heights(ws, {33: C.H_GAP})


# ============================================================================ So gehen Sie vor
def flow(ws):
    wipe(ws, "C", 34, "H", 45)
    C.section(ws, 34, "C", "G", "So gehen Sie vor", meta="in sechs Schritten")
    for i, (title, sheet, target, desc) in enumerate(FLOW):
        r = 35 + i
        C.safe_merge(ws, "C", r, "D", r)
        a = ws[f"C{r}"]
        a.value = rich((CIRCLED[i], C.T_H3, True, ACCENT), ("   " + title + "  ›", C.T_BODY, True, BLUE))
        a.hyperlink = Hyperlink(ref=a.coordinate, location=C.link_loc(sheet, target), display=title,
                                tooltip="Zur Auswahl „Kauf als“" if sheet == SHEET else f"Zum Blatt „{sheet}“")
        a.font = font(C.T_BODY, True, BLUE)
        a.alignment = align("left", "center", 1)
        C.safe_merge(ws, "E", r, "G", r)
        d = ws[f"E{r}"]
        C.set_text(d, desc)
        d.font = font(C.T_SMALL, False, INK2)
        d.alignment = align("left", "center", 1)
        C.hairline(ws, r, "C", "G")
        C.set_height(ws, r, 24)
    heights(ws, {41: C.H_GAP, 42: 2, 43: 2, 44: 4, 45: 4})       # A44 = Sprungziel „≡ Alle Bereiche“


# ============================================================================ Legende + Tipp | Blätter
def legend_and_sheets(ws):
    """Links Legende (C Begriff | D:E Erklärung) + Kasten „Tipp & Support“, rechts Blätterverzeichnis
    (F Link | G Beschreibung). Beide Spalten gleich hoch (Z. 48–62)."""
    wipe(ws, "C", 46, "H", 63)
    C.section(ws, 46, "C", "E", "Legende")
    C.section(ws, 46, "F", "G", "Alle Blätter", meta="Schritte S01–S12: siehe Leitfaden")
    ws["E46"].border = Border(bottom=side("thin", ACCENT), right=side("thick", WHITE))
    C.set_height(ws, 47, 6)
    top = 48

    legend = [
        ("Eingabe", "Gelb hinterlegt – hier eingeben, Beispielwerte überschreiben"),
        ("Verknüpfung", "Blau – übernommen aus einer Eingabe an anderer Stelle"),
        ("Berechnung", "Schwarz – berechnet, bitte nicht ändern"),
        ("Ergebnis", "Hervorgehoben – Summe bzw. Blockergebnis"),
        ("–1.234 €", "Rot – nur negative Beträge in Ergebniszeilen"),
        ("Status", None),
        ("Prüfhinweis", None),
        ("Blattschutz", "ohne Passwort – bei Bedarf aufheben"),
        ("Navigation", "Reiter oben  ·  „Weiter ›“ führt zum nächsten Schritt"),
    ]
    for i, (term, text) in enumerate(legend):
        r = top + i
        c = ws[f"C{r}"]
        C.set_text(c, term)
        c.font = font(C.T_BODY, False, INK)
        c.alignment = align("left", "center", 1)
        C.safe_merge(ws, "D", r, "E", r)
        d = ws[f"D{r}"]
        if text:
            C.set_text(d, text)
        d.font = font(C.T_SMALL, False, MUTED)
        d.alignment = align("left", "center", 1)
        C.hairline(ws, r, "C", "E")
    inp = ws[f"C{top}"]
    C.input_style(inp, "required")
    inp.protection = Protection(locked=True)                    # nur Muster, keine Eingabezelle
    inp.alignment = align("left", "center", 1)
    ws[f"C{top + 1}"].font = font(C.T_BODY, False, BLUE)
    k = ws[f"C{top + 3}"]
    C.sum_row(ws, top + 3, "C", "C", "result")
    k.font = font(C.T_BODY, True, NAVY)
    ws[f"C{top + 4}"].font = font(C.T_BODY, True, RED)
    ws[f"D{top + 5}"].value = C.status_legend(("erfüllt", "prüfen", "kritisch"), size=C.T_SMALL)
    ws[f"D{top + 6}"].value = rich(("▲", C.T_SMALL, True, RED), ("  Warnung      ", C.T_SMALL, False, MUTED),
                                   ("ⓘ", C.T_SMALL, True, BLUE), ("  Information", C.T_SMALL, False, MUTED))
    for rr in range(top, top + len(legend)):
        C.set_height(ws, rr, C.H_ROW)

    # Tipp & Support (Einordnungs-Box aus core, neutral) unter der Legende
    r0 = top + len(legend) + 1                                   # 57 (56 = Luft)
    C.set_height(ws, r0 - 1, C.H_GAP)
    C.callout_box(ws, "C", r0, "E", r0 + 1, 62, title="Tipp & Support", pill=False, fit=None,
                  text="Beispielwerte in den gelben Feldern einfach überschreiben – alle Blätter rechnen sofort mit.\n"
                       "Jedes Blatt ist für den Druck auf A4 eingerichtet.\n"
                       "Blattwechsel per Tastatur: Strg + Bild ↓ / Bild ↑.\n"
                       "Version Pro  ·  Rechtsstand September 2026  ·  MM Holding GmbH, Weingarten")
    for rr in range(r0 + 1, 63):
        C.set_height(ws, rr, C.H_ROW)

    n_charts = len(ws.parent["Diagramme"]._charts) if "Diagramme" in ws.parent.sheetnames else 0
    r = top
    for name, desc in SHEETS:
        if name not in ws.parent.sheetnames and name != "Dashboard":      # Dashboard entsteht erst nach den Layouts
            continue
        link(ws[f"F{r}"], f"{name}  ›", name)
        g = ws[f"G{r}"]
        C.set_text(g, desc.format(n=n_charts or "Alle"))
        g.font = font(C.T_SMALL, False, MUTED)
        g.alignment = align("left", "center", 1)
        C.hairline(ws, r, "F", "G")
        C.set_height(ws, r, C.H_ROW)
        r += 1


def foot(ws):
    wipe(ws, "C", 63, "H", 66)
    C.set_height(ws, 63, C.H_GAP)
    C.footer(ws, 64, "C", "G")
    C.set_height(ws, 66, 15)


def layout(ws):
    set_widths(ws, WIDTHS)
    hero(ws)
    selector(ws)
    tiles(ws)
    flow(ws)
    legend_and_sheets(ws)
    foot(ws)
    ws.freeze_panes = "A4"
    ws.sheet_view.showRowColHeaders = False
    C.cf_close(ws)


def apply(wb):
    purchase_selector(wb)
    layout(wb[SHEET])
