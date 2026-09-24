"""Blatt „Start“: Hero, Rechtsform-Auswahl (Schritt ①), Kennzahl-Kacheln, Ablauf, Legende, Blätterliste, Fuß.

Raster: C 18 | D 16 | E, F, G je 34  →  C+D = E = F = G (vier gleich breite Kacheln).
Namensziele bleiben an ihrem Platz: Rechtsform = D23, ST_BMR = E28, ST_CF = C31, ST_DSCR = E31, ST_IRR = F31.
Geändert werden nur Darstellung, statische Beschriftungen und nicht referenzierte Anzeigeformeln.
"""
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Border, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

import core as ui
from core import (ACCENT, AMBER, BLUE, DISPLAY, GREEN, INK, INK2, MIST, MUTED, NAVY, RED, SANS, SKY, TINT, WHITE,
                  align, fill, font, side)

SHEET = "Start"
HERO_ROWS = range(5, 22)          # B5:H21 Navy-Fläche
HERO_COLS = "BCDEFGH"
WIDTHS = {"C": 18, "D": 16, "E": 34, "F": 34, "G": 34}

# Ablauf „So gehen Sie vor“: (Titel, Zielblatt, Zielzelle, Beschreibung)
FLOW = [
    ("Rechtsform wählen", "Start", "D23",
     "Privatperson oder Kapitalgesellschaft – im gelben Feld oben. Die Auswahl steuert die Steuerberechnung."),
    ("Leitfaden: Schritt 01–12", "Leitfaden", None,
     "Zwölf Seiten vom Objekt bis zum Ergebnis – jede mit Eingaben, Zwischenergebnis und Einordnung."),
    ("Dashboard: Ergebnis auf einen Blick", "Dashboard", None,
     "Die Gesamtbewertung auf einer Seite: Kennzahlen mit Ampel, Cashflow und Vermögensentwicklung."),
    ("Cockpit & Diagramme: Details", "Cockpit", None,
     "Detailkennzahlen und Prüfhinweise im Cockpit, alle Auswertungen als Diagramm."),
    ("Sensitivität: Was-wäre-wenn", "Sensitivität", None,
     "Break-even-Miete und -Zins, Miete × Zins und IRR-Matrix – wie robust ist die Rechnung?"),
    ("Bankunterlagen", "Bankgespräch", None,
     "Bankgespräch als A4-Übersicht, dazu Haushaltsrechnung und Vermögensaufstellung."),
]

# Blätterliste in Reiterreihenfolge (Dashboard direkt nach Leitfaden); „{n}“ = Anzahl Diagramme
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


# ============================================================================ lokale Helfer
def rich(*parts):
    """Rich-Text aus (Text, Größe, fett, Farbe[, Schriftart])-Teilen."""
    blocks = []
    for p in parts:
        text, size, bold, color = p[:4]
        name = p[4] if len(p) > 4 else SANS
        blocks.append(TextBlock(InlineFont(rFont=name, sz=size, b=bold, color=color), text))
    return CellRichText(*blocks)


def unmerge_in(ws, c1, r1, c2, r2):
    c1, c2 = ui.col(c1), ui.col(c2)
    for mr in list(ws.merged_cells.ranges):
        if not (mr.max_row < r1 or mr.min_row > r2 or mr.max_col < c1 or mr.min_col > c2):
            ws.unmerge_cells(str(mr))


def drop_cf(ws, coords):
    """Bedingte Formate entfernen, deren Bereich nur aus den genannten Zellen besteht (vor add_ampel)."""
    targets = set(coords)
    new = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        cells = {c for cr in cf.sqref.ranges for row in cr.cells for c in [ui.L(row[1]) + str(row[0])]}
        if cells and cells <= targets:
            continue
        for rule in cf.rules:
            new.add(str(cf.sqref), rule)
    ws.conditional_formatting = new


def wipe(ws, c1, r1, c2, r2, formulas=False):
    """Bereich leeren: Verbünde lösen, statische Werte, Links und Stile entfernen.
    Formeln bleiben, außer formulas=True (nur für Bereiche ohne Formeln der Vorlage)."""
    unmerge_in(ws, c1, r1, c2, r2)
    for c in ui.iter_cells(ws, c1, r1, c2, r2):
        if isinstance(c, MergedCell):
            continue
        c.hyperlink = None
        if formulas or not ui.is_formula(c.value):
            c.value = None
        c.fill = ui.NOFILL
        c.border = Border()
        c.font = font()
        c.alignment = align("general", "center")
        c.number_format = "General"


def link(cell, text, sheet, target=None, size=ui.T_BODY, bold=True, color=BLUE, tooltip=None, h="left", indent=1):
    ui.text_link(cell, text, sheet, target_cell=target, size=size, bold=bold,
                 tooltip=tooltip or f"Zum Blatt „{sheet}“")
    cell.font = font(size, bold, color)
    cell.alignment = align(h, "center", indent)


def set_widths(ws, widths):
    """Spaltenbreiten setzen; Bereichs-Dimensionen der Vorlage (z. B. D:G in einem <col>) vorher auftrennen."""
    from openpyxl.worksheet.dimensions import ColumnDimension
    for key, d in list(ws.column_dimensions.items()):
        if d.min and d.max and d.max > d.min:
            for idx in range(d.min + 1, min(d.max, 60) + 1):
                letter = ui.L(idx)
                if letter not in ws.column_dimensions or ws.column_dimensions[letter].min != idx:
                    ws.column_dimensions[letter] = ColumnDimension(ws, index=letter, width=d.width,
                                                                   customWidth=d.customWidth, hidden=d.hidden)
            d.max = d.min
    for k, w in widths.items():
        dim = ws.column_dimensions[k]
        dim.width = w
        dim.min = dim.max = ui.col(k)


def heights(ws, spec):
    for r, h in spec.items():
        ui.set_height(ws, r, h)


# ============================================================================ Rechtsform (Logik bleibt)
def purchase_selector(wb):
    """Direktauswahl Privat / Kapitalgesellschaft auf der Startseite (steuert den Namen „Rechtsform“)."""
    ws, s09 = wb[SHEET], wb["S09 Steuern"]
    current = s09["D12"].value
    unmerge_in(ws, "D", 23, "H", 23)
    ws.merge_cells("D23:G23")
    sel = ws["D23"]
    sel.value = current
    sel.protection = Protection(locked=False)
    dv = DataValidation(type="list", formula1="=L_Rechtsform", allow_blank=False, showDropDown=False,
                        showErrorMessage=True, errorStyle="stop", showInputMessage=True)
    dv.promptTitle = "Rechtsform"
    dv.prompt = "Privatperson oder Kapitalgesellschaft (vermögensverwaltende bzw. gewerbliche GmbH) aus der Liste wählen."
    dv.errorTitle = "Bitte aus der Liste wählen"
    dv.error = "Privatperson, vermögensverwaltende GmbH oder gewerbliche GmbH/Holding über ▾ auswählen."
    ws.add_data_validation(dv)
    dv.add("D23")
    wb.defined_names["Rechtsform"] = DefinedName("Rechtsform", attr_text="Start!$D$23")
    # Schritt 09 zeigt die Auswahl nur noch an (verknüpfter Wert mit Link zur Startseite)
    for dvs in list(s09.data_validations.dataValidation):
        if "D12" in str(dvs.sqref):
            s09.data_validations.dataValidation.remove(dvs)
    d12 = s09["D12"]
    d12.value = "=Rechtsform"
    ui.input_style(d12, "linked")
    d12.protection = Protection(locked=True)
    d12.hyperlink = Hyperlink(ref="D12", location="'Start'!D23", display="Rechtsform auf der Startseite",
                              tooltip="Rechtsform auf der Startseite ändern")
    s09["F12"].value = "Ändern auf der Startseite (Feld „Rechtsform“). " + str(s09["F12"].value or "")
    s09["F12"].data_type = "s"


# ============================================================================ Hero
def hero(ws):
    unmerge_in(ws, "D", 12, "G", 19)                       # alte Verbünde der Band-/Standzeilen
    for r in HERO_ROWS:
        for c in HERO_COLS:
            cell = ws[f"{c}{r}"]
            cell.fill = fill(NAVY)
            cell.border = Border()
            if not isinstance(cell, MergedCell):
                cell.hyperlink = None
    keep = {"D6", "D7", "D9", "D11", "D14", "F14", "G14", "D15", "F15", "G15", "D17", "D19", "F19", "G19"}
    for r in HERO_ROWS:                     # Eckzeichen, Feature-Liste, Impressum und sonstige Reste entfernen
        for c in HERO_COLS:
            cell = ws[f"{c}{r}"]
            if cell.coordinate not in keep and not isinstance(cell, MergedCell):
                cell.value = None
    for r in (6, 7, 9, 11, 17):
        ui.safe_merge(ws, "D", r, "G", r)

    e = ws["D6"]                                                         # Eyebrow: Firma
    e.font = font(ui.T_MICRO, True, SKY)
    e.alignment = align("left", "center")
    o = ws["D7"]                                                         # Objekt
    o.value = "=Obj_Name"
    o.font = font(12, False, MIST, DISPLAY)
    o.alignment = align("left", "center")
    t = ws["D9"]                                                         # Titel
    t.font = font(ui.T_HERO, True, WHITE, DISPLAY)
    t.alignment = align("left", "center")
    s = ws["D11"]                                                        # Unterzeile
    s.font = font(ui.T_BODY, False, MIST)
    s.alignment = align("left", "center")

    # Kennzahlenband: drei Spalten D:E | F | G (dieselben Spalten wie die Aktionszeile)
    ui.safe_merge(ws, "D", 14, "E", 14)
    ui.safe_merge(ws, "D", 15, "E", 15)
    band = [("D", "GESAMTINVESTITION", "=Gesamtinvestition", ui.NUMFMT["eur"], 0),
            ("F", "KAUFPREIS", "=Kaufpreis", ui.NUMFMT["eur"], 2),
            ("G", "HALTEDAUER", "=Haltedauer", ui.NUMFMT["years"], 2)]
    for c, lab, val, fmt, ind in band:
        lc, vc = ws[f"{c}14"], ws[f"{c}15"]
        ui.set_text(lc, lab)
        lc.font = font(ui.T_MICRO, True, SKY)
        lc.alignment = align("left", "bottom", ind)
        vc.value = val
        vc.font = font(ui.T_H2, True, WHITE, DISPLAY)
        vc.alignment = align("left", "top", ind)
        vc.number_format = fmt
    for c in "DEFG":
        ws[f"{c}14"].border = Border(top=side("hair", ACCENT))

    st = ws["D17"]                                                       # Standzeile ohne Koordinaten/Adresse
    st.value = ('="Stand "&TEXT(DAY(TODAY()),"00")&"."&TEXT(MONTH(TODAY()),"00")&"."&YEAR(TODAY())'
                '&"  ·  Rechtsstand September 2026"')
    st.font = font(ui.T_MICRO, False, MIST)
    st.alignment = align("left", "center")

    # Aktionszeile im Hero: primär (weiß auf Navy) + zwei Textlinks
    ui.safe_merge(ws, "D", 19, "E", 19)
    cta = ws["D19"]
    link(cta, "Leitfaden starten  ›", "Leitfaden", size=10.5, color=NAVY, h="center", indent=0,
         tooltip="Zur Übersicht der zwölf Schritte")
    for c in "DE":
        ws[f"{c}19"].fill = fill(WHITE)
    link(ws["F19"], "Dashboard  ›", "Dashboard", color=MIST, indent=2, tooltip="Gesamtbewertung auf einer Seite")
    link(ws["G19"], "Cockpit  ›", "Cockpit", color=MIST, indent=2, tooltip="Detailkennzahlen und Prüfhinweise")

    for c in HERO_COLS:                                                 # Abschluss: Akzentlinie
        ws[f"{c}21"].border = Border(bottom=side("thick", ACCENT))
    heights(ws, {4: 15, 5: 13.5, 6: 13.5, 7: 21.75, 8: 13.5, 9: 43.5, 10: 9.75, 11: 18, 12: 5, 13: 13.5,
                 14: 21.75, 15: 30, 16: 7.5, 17: 13.5, 18: 6, 19: 26, 20: 10, 21: 12})


# ============================================================================ Schritt ①: Rechtsform
def selector(ws):
    for c in HERO_COLS:
        ws[f"{c}22"].fill = ui.NOFILL
        ws[f"{c}22"].border = Border()
    lab = ws["C23"]
    lab.value = rich(("RECHTSFORM", 10, True, NAVY), ("\nBitte wählen ▾", 8, False, MUTED))
    lab.font = font(10, True, NAVY)
    lab.alignment = align("right", "center", 1, wrap=True)
    ln = side("medium", ui.INPUT_LINE)
    for c in "DEFGH":
        cell = ws[f"{c}23"]
        cell.fill = fill(ui.INPUT_BG)
        cell.border = Border(top=ln, bottom=ln, left=ln if c == "D" else None, right=ln if c == "H" else None)
    sel = ws["D23"]
    sel.font = font(11, True, ui.INPUT_FG)
    sel.alignment = align("left", "center", 1)
    arrow = ws["H23"]
    ui.set_text(arrow, "▾")
    arrow.font = font(11, True, ui.INPUT_FG)
    arrow.alignment = align("center", "center")
    arrow.hyperlink = Hyperlink(ref="H23", location="'Start'!D23", display="▾", tooltip="Rechtsform auswählen")

    # Zeile 24: Hilfetext statt Buttonzeile (Aktionen stehen im Hero)
    wipe(ws, "D", 24, "H", 24)
    ui.safe_merge(ws, "D", 24, "G", 24)
    ui.set_text(ws["D24"], "Privatperson oder Kapitalgesellschaft (vermögensverwaltende GmbH / gewerbliche "
                           "GmbH-Holding) – steuert die Steuerberechnung ab Schritt 09.")
    ws["D24"].font = font(8.5, False, MUTED)
    ws["D24"].alignment = align("left", "top", 1)
    heights(ws, {22: 14, 23: 36, 24: 18, 25: 12})


# ============================================================================ Kacheln
def tiles(ws):
    ui.band_l1(ws, 26, "C", "G", "Aktuelle Kalkulation im Überblick")
    for c in "CDEFG":
        for r in (27, 28, 29, 30, 31):
            ws[f"{c}{r}"].border = Border()
    ui.kpi_tile(ws, "C", "D", 27, 28, kpi="EK")
    ui.kpi_tile(ws, "E", "E", 27, 28, kpi="BMR")
    ui.kpi_tile(ws, "F", "F", 27, 28, kpi="FAKTOR")
    ui.kpi_tile(ws, "G", "G", 27, 28, kpi="RATE")
    ui.kpi_tile(ws, "C", "D", 30, 31, kpi="CF")
    ui.kpi_tile(ws, "E", "E", 30, 31, kpi="DSCR")
    ui.kpi_tile(ws, "F", "F", 30, 31, kpi="IRR")          # F30: Formel-Label bleibt
    ui.kpi_tile(ws, "G", "G", 30, 31, kpi="RF", value="=" + ui.RECHTSFORM_SHORT)
    for c in "BCDEFGH":                                     # Fugenzeile
        ws[f"{c}29"].fill = ui.NOFILL
        ws[f"{c}29"].border = Border()
    heights(ws, {29: 2.25, 32: 12, 33: 12})


# ============================================================================ So gehen Sie vor
def flow(ws):
    wipe(ws, "C", 34, "G", 45)
    ui.band_l1(ws, 34, "C", "G", "So gehen Sie vor")
    for i, (title, sheet, target, desc) in enumerate(FLOW):
        r = 35 + i
        num = ws.cell(r, 3)
        ui.set_text(num, f"{i + 1:02d}")
        num.font = font(14, True, ACCENT, DISPLAY)
        num.alignment = align("right", "center", 1)
        ui.safe_merge(ws, "D", r, "E", r)
        link(ws[f"D{r}"], f"{title}  ›", sheet, target, size=10.5,
             tooltip="Zum Auswahlfeld „Rechtsform“" if sheet == SHEET else None)
        ui.safe_merge(ws, "F", r, "G", r)
        d = ws[f"F{r}"]
        ui.set_text(d, desc)
        d.font = font(9.5, False, INK2)
        d.alignment = align("left", "center", 1, wrap=True)
        ui.hairline(ws, r, "C", "G")
        ui.set_height(ws, r, 30)
    heights(ws, {41: 10, 42: 8, 43: 2, 44: 2, 45: 2})


# ============================================================================ Legende + Blätter
def legend_and_sheets(ws):
    """Legende (C Begriff | D:E Erklärung) und Blätterliste (F Link | G Beschreibung), beide mit L2-Kopf;
    Zeile 47 ist eine schmale Luftzeile unter den Köpfen, damit Muster und Kopflinie sich nicht berühren."""
    wipe(ws, "C", 46, "G", 63)
    ui.subhead_l2(ws, 46, "C", "E", "Legende")
    ui.subhead_l2(ws, 46, "F", "G", "Blätter")
    ws["E46"].border = Border(bottom=side("thin", ACCENT), right=side("thick", WHITE))
    ui.set_height(ws, 47, 4)
    top = 48

    legend = [
        ("Eingabe", "Gelb hinterlegt – hier eingeben (Beispielwerte überschreiben)"),
        ("Verknüpfung", "Blau – übernommen aus dem Leitfaden"),
        ("Berechnung", "Schwarz – berechnet, bitte nicht ändern"),
        ("Kennzahl", "Hervorgehoben – Summe / Ergebnis"),
        ("Prüfhinweis", "⚠ Warnung  ·  ℹ Information"),
        ("Ampel", "grün / gelb / rot (Schwellen: Konfiguration)"),
        ("Blattschutz", "ohne Passwort – bei Bedarf aufheben"),
    ]
    for i, (term, text) in enumerate(legend):
        r = top + i
        c = ws[f"C{r}"]
        ui.set_text(c, term)
        c.font = font(ui.T_BODY, False, INK)
        c.alignment = align("left", "center", 1)
        ui.safe_merge(ws, "D", r, "E", r)
        d = ws[f"D{r}"]
        ui.set_text(d, text)
        d.font = font(9.5, False, MUTED)
        d.alignment = align("left", "center", 1)
    inp = ws[f"C{top}"]
    ui.input_style(inp, "required")
    inp.protection = Protection(locked=True)                    # nur Muster, keine Eingabezelle
    inp.alignment = align("left", "center", 1)
    ws[f"C{top + 1}"].font = font(ui.T_BODY, False, BLUE)
    k = ws[f"C{top + 3}"]
    k.fill = fill(ui.TINT_XL)
    k.font = font(ui.T_BODY, True, NAVY)
    k.border = Border(top=side("thin", NAVY))
    ws[f"C{top + 4}"].font = font(ui.T_BODY, True, RED)
    ws[f"C{top + 5}"].value = rich(("Ampel  ", ui.T_BODY, False, INK), ("●", ui.T_BODY, False, GREEN),
                                   ("●", ui.T_BODY, False, AMBER), ("●", ui.T_BODY, False, RED))

    n_charts = len(ws.parent["Diagramme"]._charts) if "Diagramme" in ws.parent.sheetnames else 0
    r = top
    for name, desc in SHEETS:
        if name not in ws.parent.sheetnames and name != "Dashboard":      # Dashboard entsteht erst nach den Layouts
            continue
        link(ws[f"F{r}"], f"{name}  ›", name)
        g = ws[f"G{r}"]
        ui.set_text(g, desc.format(n=n_charts or "Alle"))
        g.font = font(9.5, False, MUTED)
        g.alignment = align("left", "center", 1)
        ui.hairline(ws, r, "F", "G")
        r += 1
    for rr in range(top, max(r, top + len(legend))):
        ui.set_height(ws, rr, ui.H_ROW)


def foot(ws):
    wipe(ws, "C", 63, "H", 66)
    ui.set_height(ws, 63, 18)
    ui.footer(ws, 64, "C", "G")
    ui.set_height(ws, 66, 15)


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


def apply(wb):
    purchase_selector(wb)
    layout(wb[SHEET])
