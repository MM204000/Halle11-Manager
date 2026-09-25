"""Blatt „Start“: Hero, Entscheidungskarte „① Kauf als“, vier Kennzahl-Kacheln, Ablauf ①–⑥, Legende mit
Tipp-Kasten, Blätterverzeichnis, Fuß – ausschließlich mit den zentralen Bausteinen aus core.py.

Raster (px, Runde 4 / P1-08): C 224 | D 112 | E 336 | F 336 | G 336  →  C:D = E = F = G (vier gleich breite Spalten);
  Hero B:H von 31 bis 1 417 px = linke/rechte Inhaltskante der Schrittseiten und der festen Reiterleiste (NAV_END).
  Hero:  Logo-Spalte C:D (Logo zentriert), Inhalt E:G; Kennzahlen und Buttons exakt in E | F | G.
  Kacheln Z. 27–29: C:D | E | F | G – C.tile: Kopfstreifen · Wert in Statusfarbe · Fußzeile Kontext links / Status rechts.
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
from core import BLUE, GOLD, INK, INK2, LINE2, MIST, MUTED, NAVY, RED, SKY, WHITE, align, fill, font, side

ui = C                                   # Altname (leitfaden.py importiert Helfer aus diesem Modul)
SHEET = "Start"
HERO_ROWS = range(5, 20)                 # B5:H19 Navy-Fläche, Z. 20 Weißraum
HERO_COLS = "BCDEFGH"
HERO_RULE = C.mix(NAVY, SKY, 0.3)       # zarte Trennlinie AUF Navy (Kennzahlenband) – aus Tokens gemischt, kein Blau-Akzent
# Runde 4 (P1-08): rechte Kante des Heros = NAV_END der festen Reiterleiste (1 417 px = Kante der Schrittseiten)
WIDTHS = {"C": 32, "D": 16, "E": 48, "F": 48, "G": 48}   # C:D = E = F = G = 336 px, H 21 px → Hero bis 1 417 px
EMU = 9525                               # EMU je Pixel
KAUF_ALS = 'CHOOSE(Rechtsform_Idx,"Privatperson","vv-GmbH","GmbH / Holding")'

# Ablauf „So gehen Sie vor“: (Titel, Zielblatt, Zielzelle, Beschreibung ≤ 95 Zeichen, einzeilig)
FLOW = [
    ("Kauf als wählen", "Start", "D23",
     "Privatperson oder Gesellschaft – die Auswahl oben steuert Steuern, Exit und Bankunterlagen."),
    ("Leitfaden: Schritte 01–12", "Leitfaden", None,
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
ETAPPE = "ABCDEF"                        # Etappen A–F („Schritt“ bleibt S01–S12 vorbehalten, P37)

# Blätterverzeichnis in Reiterreihenfolge; „{n}“ = Anzahl Diagramme
SHEETS = [
    ("Leitfaden", "Zwölf Schritte vom Objekt zum Ergebnis"),
    ("Dashboard", "Gesamtbewertung auf einer Seite"),
    ("Cockpit", "Detailkennzahlen und Prüfhinweise"),
    ("Diagramme", "{n} Diagramme zur Kalkulation"),
    ("Eingaben", "Alle Eingaben und Profi-Felder"),
    ("Projektion", "Miete, Kosten, Vermögen – 40 Jahre"),
    ("Steuern", "Steuerberechnung je Jahr  (Eingaben: Schritt 09)"),
    ("AfA-Vergleich", "Abschreibungsvarianten im Vergleich"),
    ("Finanzierung", "Tilgungspläne I und II  (Eingaben: Schritt 07)"),
    ("Sensitivität", "Break-even, Miete × Zins, IRR-Matrix"),
    ("Bankgespräch", "Investitionsübersicht für die Bank"),
    ("Haushaltsrechnung", "Selbstauskunft Einnahmen/Ausgaben"),
    ("Vermögensaufstellung", "Selbstauskunft Vermögen"),
    ("Hinweise", "Rechtsgrundlagen und Modellannahmen"),
    ("Konfiguration", "Stammdaten, Tarif 2026, Ampel-Schwellen"),
]

# Anzeigename im Verzeichnis = Reiterbeschriftung (P2-05: ein Name pro Ziel); das Blatt selbst heißt unverändert
SHEET_LABEL = {"Steuern": "Steuer-Tabelle", "Finanzierung": "Tilgungsplan"}


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


fixed_m = C.fixed_m                      # FIXED mit typografischem Minus (zentral in core, Runde 4)


# Kanonische Kachel-Fußzeilen (P11/P20) – wortgleich mit dem Dashboard; Status setzt C.tile selbst (rechts).
TILE_SUB = {
    "GI": '="Kaufpreis "&FIXED(Kaufpreis,0)&" €  ·  NK "&FIXED(NK_Quote*100,1)&" %"',
    "EK": '="Quote "&FIXED(EK_Quote*100,1)&" %  ·  Beleihung "&FIXED(Beleihung*100,1)&" %"',
    "CF": f'="Jahr 2: "&{fixed_m(C.CF_YEAR2)}&" €"',
    "BMR": "=" + C.threshold_text("BMR"),
    "DSCR": '="Jahr 1  ·  "&' + C.threshold_text("DSCR"),
    "IRR": "=" + C.threshold_text("IRR"),
    "RATE": '="Volltilgung "&Volltilgung_Txt',
}


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
    for c in "EFGH":                                   # D23:F23 wird das Feld, G23 der Auswahl-Knopf (selector)
        ws[f"{c}23"].value = None
        ws[f"{c}23"].hyperlink = None
    ws.merge_cells("D23:F23")
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
def _logo_png(img):
    """Logo für Navy nachschärfen (P45, Runde 5): Buchstaben Weiß, Ring in Edel-Gold C.GOLD (wie das Original-Signet,
    nur kräftiger), Linien leicht verstärkt und voll deckend –
    das feine Original (Creme/Gold, 512 px) wirkt bei 150 px auf Navy sonst wie ein Wasserzeichen.
    Liefert einen Pfad auf eine temporäre PNG-Datei (openpyxl liest das Bild beim Speichern)."""
    import io
    import tempfile
    from PIL import Image, ImageFilter
    im = Image.open(io.BytesIO(img._data())).convert("RGBA")
    r, g, b, a = im.split()
    a = a.filter(ImageFilter.MaxFilter(5)).point(lambda v: min(255, int(v * 1.6)))
    ring = Image.new("L", im.size)
    px, rp = im.load(), ring.load()
    for y in range(im.size[1]):
        for x in range(im.size[0]):
            pr, pg, pb, pa = px[x, y]
            if pa and pr - pb > 45:                    # goldener Ring (R ≫ B) → GOLD, Rest (Buchstaben) → Weiß
                rp[x, y] = 255
    ring = ring.filter(ImageFilter.MaxFilter(5))
    white = Image.new("RGBA", im.size, (255, 255, 255, 255))
    gold = Image.new("RGBA", im.size, tuple(int(GOLD[i:i + 2], 16) for i in (0, 2, 4)) + (255,))
    out = Image.composite(gold, white, ring)
    out.putalpha(a)
    f = tempfile.NamedTemporaryFile(prefix="mm_logo_", suffix=".png", delete=False)
    out.save(f, format="PNG")
    f.close()
    return f.name


def logo(ws, size_px=150):
    """Logo als OneCellAnchor mit fester Größe (kein Verzerren), zentriert in der Logo-Spalte C:D und vertikal
    zentriert auf die Titelzone Eyebrow … Kennzahlen (Z. 6–15, P45)."""
    def row_px(r):
        return round((ws.row_dimensions[r].height or 15) / 0.75)
    zone = sum(row_px(r) for r in range(6, 16))
    top = max(0, (zone - size_px) // 2)
    row, off = 6, top
    while off >= row_px(row):
        off -= row_px(row)
        row += 1
    left = max(0, (C.span_px(ws, "C", "D") - size_px) // 2)
    col, coff = 3, left
    if coff >= C.col_px(ws, "C"):
        coff -= C.col_px(ws, "C")
        col = 4
    for img in ws._images:
        a = getattr(img.anchor, "_from", None)
        if a is None or a.col > 3 or not (4 <= a.row <= 12):
            continue
        try:
            img.ref = _logo_png(img)
        except Exception as exc:                        # Logo bleibt im Original, wenn PIL fehlt
            print(f"   Hinweis start.logo: Logo nicht umgefärbt ({exc})")
        marker = AnchorMarker(col=col - 1, colOff=coff * EMU, row=row - 1, rowOff=off * EMU)
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
        put(f"{c}15", val, C.T_H1, True, WHITE, fmt=fmt)        # P3-01: größte Zahl der Seite (22 pt > Kachel 20 pt)
    hair = side("thin", HERO_RULE)
    for c in "EFG":
        ws[f"{c}14"].border = Border(top=hair)

    # Aktionszeile im Drittelraster E | F | G (P3-01, Runde 5) – ausschließlich core-Buttons, gleiche Höhe (25,5 pt):
    #   E  primary  Nachtblau mit Goldrahmen – zuerst die Pflichtauswahl „Kauf als“ (Sprung auf D23)
    #   F  ghost    1B3553, Rahmen 3A5578    – Leitfaden starten
    #   G  ghost    1B3553, Rahmen 3A5578    – Dashboard
    # Auf der Navy-Fläche trägt der Primär-Button die Goldkante (EIN Goldelement je Komponente); die Fugen zwischen den
    # Buttons sind 3-px-Kanten in Heldenfarbe (wie die Kachelrinnen).
    C.btn(ws, "E", 17, "E", "Zuerst: Kauf als wählen  ↓", SHEET, "primary", target_cell="D23",
          tooltip="Pflichtauswahl: Privatperson oder Gesellschaft (Liste)")
    C.btn(ws, "F", 17, "F", "Leitfaden starten  ›", "Leitfaden", "ghost", tooltip="Zur Übersicht der zwölf Schritte")
    C.btn(ws, "G", 17, "G", "Dashboard  ›", "Dashboard", "ghost", tooltip="Gesamtbewertung auf einer Seite")
    gap = side("thick", NAVY)
    for c, left, right in (("E", False, True), ("F", True, True), ("G", True, False)):
        b = ws[f"{c}17"].border
        ws[f"{c}17"].border = Border(top=b.top, bottom=b.bottom, left=gap if left else b.left,
                                     right=gap if right else b.right)
    # Goldkante des Primär-Buttons auf Medium (2 px) – auf Nachtblau sonst zu fein; Ghost-Buttons bleiben 1 px
    g2 = side("medium", GOLD)
    ws["E17"].border = Border(top=g2, bottom=g2, left=g2, right=gap)

    for c in HERO_COLS:                                                 # Abschluss: Goldlinie (Hero-Linie, Runde 5)
        ws[f"{c}19"].border = Border(bottom=side("thick", GOLD))
    heights(ws, {4: C.H_HDR[4], 5: 16, 6: 14, 7: 18, 8: 12, 9: 42, 10: 6, 11: 18, 12: 14, 13: 12, 14: 20, 15: 34,
                 16: 16, 17: C.H_BTN, 18: 12, 19: 10, 20: 12})
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


# ============================================================================ Kauf als – Entscheidungskarte
SEL_LINE = GOLD                        # Edel-Gold C9A14A: Rahmen der Pflichtauswahl (Eingabefamilie, P06, Runde 5)


def selector(ws):
    """Pflichtauswahl als erkennbare Auswahlliste (P06, P3-01): Beschriftung „KAUF ALS / Pflichtfeld“ (ohne ▾),
    Feld D23:F23 (gelb, Datenüberprüfung Liste) + Hinweis G23 „Liste öffnen: Alt + ↓“ (weiß, linke Trennlinie E6CB77,
    Sprung auf D23) – beide in EINEM Rahmen der Eingabefamilie (medium C9A94A, links 3 px). Das ▾ steht nur noch
    einmal: in der Erklärzeile D24 unter dem Feld."""
    unmerge_in(ws, "C", 21, "H", 22)
    C.section(ws, 21, "C", "G", "Kauf als – Privatperson oder Gesellschaft", meta="Pflichtauswahl")
    lab = ws["C23"]
    lab.value = rich(("KAUF ALS", C.T_BODY, True, NAVY), ("\nPflichtfeld", C.T_MICRO, False, MUTED))
    lab.font = font(C.T_BODY, True, NAVY)
    lab.alignment = align("left", "center", 1, wrap=True)
    lab.border = Border()
    lab.fill = C.NOFILL
    lab.hyperlink = None
    unmerge_in(ws, "D", 23, "G", 23)
    for c in "EFG":
        ws[f"{c}23"].value = None
        ws[f"{c}23"].hyperlink = None
    C.safe_merge(ws, "D", 23, "F", 23)
    frame = side("medium", SEL_LINE)
    for c in "DEF":
        cell = ws[f"{c}23"]
        cell.fill = fill(C.INPUT_BG)
        cell.border = Border(top=frame, bottom=frame, left=side("thick", SEL_LINE) if c == "D" else None)
    sel = ws["D23"]
    sel.font = font(C.T_BODY, True, C.INPUT_FG)
    sel.alignment = align("left", "center", 1)
    # Hinweis im Feld (kein Knopf, keine Fläche): Tastenkürzel für die Liste; Klick markiert das Feld D23
    hint = ws["G23"]
    C.set_text(hint, "Liste öffnen:  Alt + ↓")
    hint.font = font(C.T_LABEL, False, BLUE)
    hint.alignment = align("center", "center")
    hint.fill = fill(WHITE)
    hint.border = Border(top=frame, bottom=frame, right=frame, left=side("thin", C.INPUT_LINE))
    hint.hyperlink = Hyperlink(ref="G23", location=C.link_loc(SHEET, "D23"), display="Kauf als",
                               tooltip="Feld „Kauf als“ markieren – dann Alt + ↓ oder den Pfeil am Feldrand")
    ws["H23"].fill = C.NOFILL
    ws["H23"].border = Border()

    wipe(ws, "C", 24, "H", 24)
    C.safe_merge(ws, "D", 24, "G", 24)
    C.set_text(ws["D24"], "▾  Aus der Liste wählen: Privatperson · vv-GmbH · gewerbliche GmbH/Holding"
                          "   ·   wirkt auf S09 Steuern, S11 Exit, Dashboard und Bankgespräch")
    ws["D24"].font = font(C.T_MICRO, False, MUTED)
    ws["D24"].alignment = align("left", "center", 1)
    for c in HERO_COLS:
        for r in (20, 22):
            ws[f"{c}{r}"].fill = C.NOFILL
            ws[f"{c}{r}"].border = Border()
    heights(ws, {20: 12, 22: 8, 23: 36, 24: 18, 25: C.H_GAP})


# ============================================================================ Kacheln
TILES = [  # (c1, c2, KPI, Wertformel) – Teilmenge der kanonischen Reihenfolge C.KPI_ORDER (P20)
    ("C", "D", "CF", "=CF_nSt_Monat_J1"),
    ("E", "E", "BMR", "=Bruttomietrendite"),
    ("F", "F", "DSCR", "=DSCR_J1"),
    ("G", "G", "IRR", "=EK_IRR"),
]


def tiles(ws):
    """Vier Kacheln mit C.tile (dark) – EINE Anatomie wie Dashboard/Leitfaden (P10/P11/P20):
    Kopfstreifen mit kanonischem Label · Wert in Statusfarbe · Fußzeile Kontext links, Status rechts
    (2-spaltig als Chip, 1-spaltig inline), Rinnen 3 px horizontal = vertikal (gap='auto')."""
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
    for c in "CDEFG":                                   # Kopf- und Fußzeile (Z. 27/29) sind in der Vorlage reine Texte
        ws[f"{c}29"].value = None
    ws["D27"].value = ws["D28"].value = None
    for c1, c2, key, val in TILES:
        vref = val[1:]
        C.tile(ws, c1, c2, 27, 28, 29, kpi=key, value=val, value_ref=vref, sub=TILE_SUB[key],
               label=C.kpi_label(key, caps=True, formula=True), gap_right=c2 != "G")
    # zweite Reihe der Vorlage (Namensziele ST_CF/ST_DSCR/ST_IRR) ausblenden – Formeln bleiben unverändert
    clear_rows(ws, 30, 32)
    for c, fmt in (("C", C.NUMFMT["eur"]), ("E", C.NUMFMT["dscr"]), ("F", C.NUMFMT["pct1"])):
        ws[f"{c}31"].number_format = fmt
    C.hide_rows(ws, 30, 32)
    heights(ws, {33: C.H_GAP})


# ============================================================================ So gehen Sie vor
def flow(ws):
    wipe(ws, "C", 34, "H", 45)
    C.section(ws, 34, "C", "G", "Ihr Weg durch das Tool", meta="6 Etappen · A–F")
    for i, (title, sheet, target, desc) in enumerate(FLOW):
        r = 35 + i
        C.safe_merge(ws, "C", r, "D", r)
        a = ws[f"C{r}"]
        a.value = rich((ETAPPE[i], C.T_H3, True, C.GOLD_INK), ("   " + title + "  ›", C.T_BODY, True, BLUE))
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
    """Links Legende (C Muster | D:E Erklärung) + Kasten „Tipp & Support“, rechts Blätterverzeichnis
    (F Link | G Beschreibung). Beide Spalten gleich hoch (Z. 48–62).
    Legende (P3-20): die vier Eingabezustände über C.input_legend (C.INPUT_STATES), danach Verknüpfung, Berechnung,
    Ergebnis, Negativ-Rot, Status und Prüfhinweis. Blattschutz und Navigation stehen im Tipp-Kasten."""
    wipe(ws, "C", 46, "H", 63)
    C.section(ws, 46, "C", "E", "Legende")
    C.section(ws, 46, "F", "G", "Alle Blätter", meta="Schritte S01–S12: siehe Leitfaden")
    ws["E46"].border = Border(bottom=side("thin", LINE2), right=side("thick", WHITE))   # Unterlinie wie section()
    C.set_height(ws, 47, 6)
    gutter = side("thick", WHITE)                 # Rinne Legende | Alle Blätter (P26) – wie die Kachelrinnen
    top = 48

    # Eingabezustände (Muster in C mit Kurzwort, Erklärung in D:E)
    states = {  # Kurzwort im Muster · Erklärung (Satzbau wie die übrigen Legendenzeilen: „Farbe – Bedeutung“)
        "required": ("Eingabe", "Gelb – hier eingeben, Beispielwerte überschreiben"),
        "optional": ("optional", "Gestrichelt – darf leer bleiben"),
        "override": ("überschreibbar", "Aus der Kalkulation vorgeschlagen – bei Bedarf überschreiben"),
        "inactive": ("inaktiv", "Grau – für die gewählte Methode ohne Wirkung"),
    }
    n_states = len(C.INPUT_STATES)
    for i in range(n_states):
        C.safe_merge(ws, "D", top + i, "E", top + i)
    C.input_legend(ws, top, [(f"C{top + i}", f"D{top + i}") for i in range(n_states)])
    for i, (state, word) in enumerate(C.INPUT_STATES):
        r = top + i
        sw, tx = ws[f"C{r}"], ws[f"D{r}"]
        short, expl = states.get(state, (word, word))
        C.set_text(sw, short)
        sw.alignment = align("left", "center", 1)
        C.set_text(tx, expl)
        tx.font = font(C.T_SMALL, False, MUTED)
        tx.alignment = align("left", "center", 1)
        C.hairline(ws, r, "D", "E")
        C.set_height(ws, r, C.H_ROW)

    legend = [
        ("Verknüpfung", "Blau – übernommen aus einer Eingabe an anderer Stelle"),
        ("Berechnung", "Schwarz – berechnet, bitte nicht ändern"),
        ("Ergebnis", "Hervorgehoben – Summe bzw. Blockergebnis"),
        (f"{C.MINUS}1.234 €", C.NEG_RULE_TEXT),
        ("Status", None),
        ("Prüfhinweis", None),
    ]
    r1 = top + n_states
    for i, (term, text) in enumerate(legend):
        r = r1 + i
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
        C.set_height(ws, r, C.H_ROW)
    ws[f"C{r1}"].font = font(C.T_BODY, False, BLUE)
    C.sum_row(ws, r1 + 2, "C", "C", "final")
    ws[f"C{r1 + 2}"].font = font(C.T_BODY, True, NAVY)
    ws[f"C{r1 + 3}"].font = font(C.T_BODY, True, RED)
    ws[f"D{r1 + 4}"].value = C.status_legend(("erfüllt", "prüfen", "kritisch"), size=C.T_SMALL)
    ws[f"D{r1 + 5}"].value = rich(("▲", C.T_SMALL, True, RED), ("  Warnung      ", C.T_SMALL, False, MUTED),
                                  ("ⓘ", C.T_SMALL, True, BLUE), ("  Information", C.T_SMALL, False, MUTED))
    last_leg = r1 + len(legend) - 1                                # 57

    # Tipp & Support (Einordnungs-Box aus core, neutral) unter der Legende: Kopf 59, Körper 60–62 (3 Zeilen)
    r0 = last_leg + 2
    C.set_height(ws, r0 - 1, C.H_ROW)                             # Luft; Zeile gehört rechts zur Blätterliste
    C.callout_box(ws, "C", r0, "E", r0 + 1, 62, title="Tipp & Support", pill=False, fit=None,
                  text="Hilfe zu jedem Eingabefeld: Zelle markieren – der Eingabehinweis erscheint.\n"
                       "Navigation: Reiter oben, „Weiter  ›“ führt zum nächsten Schritt · Strg + Bild ↓ / Bild ↑.\n"
                       "Blattschutz ohne Passwort · jedes Blatt ist für den Druck auf A4 eingerichtet.")
    for rr in range(r0, 63):
        C.set_height(ws, rr, C.H_ROW)

    n_charts = len(ws.parent["Diagramme"]._charts) if "Diagramme" in ws.parent.sheetnames else 0
    r = top
    for name, desc in SHEETS:
        if name not in ws.parent.sheetnames and name != "Dashboard":      # Dashboard entsteht erst nach den Layouts
            continue
        link(ws[f"F{r}"], f"{SHEET_LABEL.get(name, name)}  ›", name)
        g = ws[f"G{r}"]
        C.set_text(g, desc.format(n=n_charts or "Alle"))
        g.font = font(C.T_SMALL, False, MUTED)
        g.alignment = align("left", "center", 1)
        C.hairline(ws, r, "F", "G")
        C.set_height(ws, r, C.H_ROW)
        r += 1
    # Rinne zwischen den beiden Spalten: weiße 3-px-Kante rechts an E (unterbricht Haarlinien und Flächen)
    from copy import copy
    for rr in range(46, 63):
        cell = ws[f"E{rr}"]
        b = copy(cell.border)
        cell.border = Border(left=b.left, top=b.top, bottom=b.bottom, right=gutter)
        f_ = ws[f"F{rr}"]
        fb = copy(f_.border)
        if rr != 46:                                      # F46 behält die Akzentkante des Abschnittskopfs
            f_.border = Border(left=gutter, top=fb.top, bottom=fb.bottom, right=fb.right)


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
