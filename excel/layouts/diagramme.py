"""Blatt „Diagramme“ (Zellen, Raster, Anker) und Diagramm-INHALTE auf allen Blättern.

Läuft als letztes Layout-Modul. Aufgaben:
1. Inhalte aller Diagramme (alle Blätter außer Dashboard, das danach entsteht):
   einheitliche Horizonte (Cashflow 30 J., Bestände 35 J., AfA/Steuer 20 J.), Kategorien als Jahreszahlen (numRef statt
   multiLvlStrRef), „Daten in ausgeblendeten Zellen anzeigen“ (plotVisOnly=0), S08: keine Steuerreihen (Schritt 09 folgt).
2. Blatt „Diagramme“: Seitenkopf, L1-Bänder B:P, Raster B:P für die Diagramme (3er-, 2er- und schmal+breit-Reihen mit genau
   einer Spalte Lücke, einheitliche Höhe), Datentabellen im Tabellenstil, Hilfsreihe „Markierung Verkaufsjahr“ (Z. 202),
   Fuß in Z. 204/205.
Formeln und Werte bestehender Zellen bleiben unverändert; neu sind nur Anzeigezellen in leeren Zellen.
"""
import re

from openpyxl.chart.data_source import AxDataSource, NumRef
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Border

import core as K

SHEET = "Diagramme"
FIRST_COL, LAST_COL = "B", "P"          # Inhaltsraster (B = 34, C:P = 11)
YEAR_ROW, INDEX_ROW = 178, 177          # Kalenderjahre / Jahr-Index 1…40
MARK_ROW = 202                          # Anzeige-Hilfsreihe: Immobilienwert im Verkaufsjahr, sonst 0

# Horizont je Datenzeile auf „Diagramme“ (letzte Spalte): Bestände 35 J., Cashflow 30 J., AfA/Steuer 20 J.
HORIZON = {r: "AL" for r in (179, 180, 181, 182, 183, 189, 190, 191, MARK_ROW)}
HORIZON.update({r: "AG" for r in (184, 185, 186)})
HORIZON.update({r: "W" for r in range(192, 198)})

# Abschnitte: (Bandzeile, Titel, rechte Angabe)
SECTIONS = [
    (9, "Investition, Finanzierung und Kaufpreisaufteilung", "Zeitpunkt Kauf · Anteile"),
    (30, "Einnahmen, Ausgaben und Cashflow", "Jahr 1 · Jahre 1–30"),
    (51, "Entwicklung von Vermögen, Darlehen und Cashflow", "Jahre 1–35 · Jahresende"),
    (92, "Steuern und Abschreibung", "Jahre 1–20 · € p. a."),
    (113, "Exit – Verkauf nach der geplanten Haltedauer", "Verkaufsjahr aus Schritt 11"),
]
DATA_BAND = 138

# Diagramm-Raster: Art → (Spalte von, Spalte bis einschließlich, erste Zeile, letzte Zeile)
CHART_ROWS = 16
GRID = {
    "invest": ("B", "D", 11), "finanz": ("F", "J", 11), "kpa": ("L", "P", 11),
    "jahr1": ("B", "D", 32), "cashflow": ("F", "P", 32),
    "bestand": ("B", "G", 53), "restschuld": ("I", "P", 53),
    "kumzins": ("B", "G", 71), "cf_kum": ("I", "P", 71),
    "afa": ("B", "G", 94), "steuer": ("I", "P", 94),
    "exit": ("B", "D", 115), "ertrag": ("F", "P", 115),
}
# Zeilenhöhen der Diagrammzone (Abstände): Zeile → pt
SPACER_ROWS = {10: 8, 31: 8, 52: 8, 93: 8, 114: 8,
               27: 10, 28: 10, 29: 10, 48: 10, 49: 10, 50: 10, 69: 9, 70: 9,
               87: 6, 88: 6, 89: 6, 90: 6, 91: 6, 110: 10, 111: 10, 112: 10,
               131: 5, 132: 5, 133: 5, 134: 5, 135: 5, 136: 5, 137: 5}

# Datentabellen: Kopfzeile → (letzte Spalte, Titel, {Spalte: Kopftext})
SMALL_TABLES = {
    140: ("C", "Gesamtinvestition", {"C": "Betrag (€)"}),
    146: ("C", "Finanzierungsstruktur", {"C": "Betrag (€)"}),
    151: ("C", "Kaufpreisaufteilung", {"C": "Betrag (€)"}),
    157: ("C", "Verwendung des Verkaufserlöses", {"C": "Betrag (€)"}),
    163: ("C", "Gesamtertrag nach Steuern", {"C": "Betrag (€)"}),
    169: ("D", "Jahr 1 pro Monat (€)", {}),     # C169/D169 („Einnahmen“/„Ausgaben“) sind Diagrammkategorien
}
SERIES_ROWS = range(179, MARK_ROW + 1)


# ============================================================================ Diagramm-Inhalte (alle Blätter)
REF_RE = re.compile(r"^(?P<sheet>.+)!\$(?P<c1>[A-Z]+)\$(?P<r1>\d+):\$(?P<c2>[A-Z]+)\$(?P<r2>\d+)$")


def _parse(f):
    m = REF_RE.match(f or "")
    if not m:
        return None
    sh = m.group("sheet").strip("'").replace("''", "'")
    return sh, m.group("c1"), int(m.group("r1")), m.group("c2"), int(m.group("r2"))


def _val_ref(s):
    try:
        return s.val.numRef.f
    except AttributeError:
        return None


def _cat_ref(s):
    if s.cat is None:
        return None
    for k in ("numRef", "strRef", "multiLvlStrRef"):
        r = getattr(s.cat, k, None)
        if r is not None:
            return r.f
    return None


def _tx_ref(s):
    if s.tx is None:
        return ""
    if s.tx.strRef is not None:
        return s.tx.strRef.f or ""
    return s.tx.v or ""


def chart_kind(ws, ch):
    """Diagrammart aus den Datenbereichen (unabhängig von Nummer und Titel)."""
    rows, sheets = set(), set()
    for s in ch.series:
        p = _parse(_val_ref(s))
        if p:
            sheets.add(p[0])
            rows.add(p[2])
    if not rows:
        return None
    if sheets == {SHEET}:
        if rows & {179, 182, 183}:
            return "bestand"
        if rows <= {180, 181}:
            return "restschuld"
        if rows & {184} and rows & {185}:
            return "cashflow"
        if rows == {185}:
            return "cf_nach"
        if rows == {186}:
            return "cf_kum"
        if rows & {189, 190, 191}:
            return "kumzins"
        if rows & {192, 193, 194, 195}:
            return "afa"
        if rows & {196, 197}:
            return "steuer"
        if rows & set(range(170, 176)):
            return "jahr1"
        single = {141: "invest", 147: "finanz", 152: "kpa", 158: "exit", 164: "ertrag"}
        for r, k in single.items():
            if r in rows:
                return k
    return None


def _set_year_categories(s, sheet, c1, c2, row):
    s.cat = AxDataSource(numRef=NumRef(f=f"{sheet}!${c1}${row}:${c2}${row}"))


def _quote(sheet):
    return f"'{sheet}'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet


def chart_contents(wb):
    for ws in wb.worksheets:
        if ws.title == "Dashboard":
            continue
        for ch in ws._charts:
            ch.visible_cells_only = False          # Daten in ausgeblendeten Zeilen/Spalten weiter anzeigen
            kind = chart_kind(ws, ch)
            # S08: vor Steuern – die Steuer kommt erst in Schritt 09
            if ws.title.startswith("S08") and kind == "jahr1":
                ch.series[:] = [s for s in ch.series if not re.search(r"\$?B\$?17[15]$", _tx_ref(s))]
                for i, s in enumerate(ch.series):
                    s.idx, s.order = i, i
            for s in ch.series:
                p = _parse(_val_ref(s))
                if not p:
                    continue
                sh, c1, r1, c2, r2 = p
                if sh == SHEET and r1 == r2 and r1 in HORIZON and c1 == "D":
                    last = HORIZON[r1]
                    s.val.numRef.f = f"{SHEET}!$D${r1}:${last}${r1}"
                    _set_year_categories(s, SHEET, "D", last, YEAR_ROW)
                    continue
                cr = _parse(_cat_ref(s))
                if cr and s.cat is not None and s.cat.multiLvlStrRef is not None and cr[2] == cr[4]:
                    # einzeilige Kategorien (Jahre, Wertsteigerung) als Zahlenreihe statt Mehrebenen-Text
                    _set_year_categories(s, _quote(cr[0]), cr[1], cr[3], cr[2])


# ============================================================================ Blatt „Diagramme“
def _col(c):
    return K.col(c)


def _anchor(ch, c1, c2, r1, rows=CHART_ROWS):
    """Diagramm genau auf die Zellen c1..c2 × r1..r1+rows-1 legen (Offsets 0)."""
    frm = AnchorMarker(col=_col(c1) - 1, colOff=0, row=r1 - 1, rowOff=0)
    to = AnchorMarker(col=_col(c2), colOff=0, row=r1 - 1 + rows, rowOff=0)
    ch.anchor = TwoCellAnchor(_from=frm, to=to, editAs="twoCell")


def _clear_row_style(ws, row, c1, c2):
    for c in K.iter_cells(ws, c1, row, c2, row):
        c.fill = K.NOFILL
        c.border = Border()


def _unmerge_rows(ws, rows):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row in rows:
            ws.unmerge_cells(str(mr))


def layout_sheet(ws):
    # ---- Raster: B = 34, C:AQ = 11 (Zeitreihe ohne Rhythmuswechsel ab Jahr 17)
    ws.column_dimensions["A"].width = 4.5
    ws.column_dimensions["B"].width = 34
    for cc in range(_col("C"), _col("AQ") + 1):
        ws.column_dimensions[K.L(cc)].width = 11

    # ---- Seitenkopf Z. 5–7 und Zeile 8 für die Sprungleiste (Formen, navigation.py)
    for r in (5, 6, 7):
        for c in K.iter_cells(ws, "B", r, "P", r):
            if c.column > 2:
                c.value = None
    K.page_header(ws, "B", "P", "Auswertung", "Diagramme",
                  "Investition, Cashflow, Vermögen, Steuern und Exit – alle Diagramme aktualisieren sich automatisch "
                  "aus den Eingaben. Datenbasis: Projektion, Finanzierung und Steuern (Diagrammdaten am Blattende).")
    ws["B7"].alignment = K.align("left", "top")
    ws.row_dimensions[8].height = 30
    for c in ws[8]:
        c.value = None if c.column > 1 else c.value

    # ---- Abschnittsbänder B:P (statt B:S)
    _unmerge_rows(ws, [row for row, *_ in SECTIONS] + [DATA_BAND])
    for row, title, right in SECTIONS:
        _clear_row_style(ws, row, "Q", "U")
        K.band_l1(ws, row, FIRST_COL, LAST_COL, title, right)
    _clear_row_style(ws, DATA_BAND, "Q", "U")
    K.band_l1b(ws, DATA_BAND, FIRST_COL, LAST_COL, "DIAGRAMMDATEN", "Automatisch berechnet – bitte nicht ändern")

    # ---- Zeilen der Diagrammzone
    for r in range(10, DATA_BAND):
        if r in [s[0] for s in SECTIONS]:
            continue
        ws.row_dimensions[r].height = SPACER_ROWS.get(r, 15)

    # ---- Diagramme ins Raster
    for ch in ws._charts:
        kind = chart_kind(ws, ch)
        if kind in GRID:
            c1, c2, r1 = GRID[kind]
            _anchor(ch, c1, c2, r1)

    # ---- Datentabellen
    ws.row_dimensions[139].height = 10
    for head, (last, title, labels) in SMALL_TABLES.items():
        lab = {"B": (title, "left")}
        for cc, text in labels.items():
            lab[cc] = (text, "right")
        K.table_head(ws, head, "B", last, lab)
        for cc in range(_col("C"), _col(last) + 1):     # Köpfe über Zahlen rechtsbündig
            ws.cell(head, cc).alignment = K.align("right", "center", 1)
        r = head + 1
        while ws.cell(r, 2).value is not None:
            _body_row(ws, r, last)
            r += 1
        ws.row_dimensions[r].height = 12              # Luft zur nächsten Tabelle

    # Zeitreihen (Jahr 1–40)
    K.table_head(ws, INDEX_ROW, "B", "AQ", {"B": ("Zeitreihen (Jahr 1–40)", "left"), "C": ("Einheit", "left")})
    for cc in range(_col("D"), _col("AQ") + 1):
        c = ws.cell(INDEX_ROW, cc)
        c.number_format = '"Jahr "0'
        c.alignment = K.align("right", "center", 1)
    K.set_text(ws.cell(YEAR_ROW, 3), "Jahr")
    _mark_row(ws)
    for r in range(YEAR_ROW, MARK_ROW + 1):
        _body_row(ws, r, "AQ", unit=True)
    for cc in range(_col("D"), _col("AQ") + 1):
        ws.cell(YEAR_ROW, cc).number_format = "0"
    for cc in range(_col("D"), _col("AQ") + 1):
        ws.cell(201, cc).number_format = K.NUMFMT["pct1"]
    for r in (MARK_ROW,):
        K.memo(ws, r, "B", "AQ")
        for cc in range(_col("C"), _col("AQ") + 1):
            ws.cell(r, cc).alignment = K.align("right" if cc > 3 else "left", "center", 1)
        ws.cell(r, 2).alignment = K.align("left", "center", 1)

    # ---- Fuß: eine Leerzeile unter der Tabelle, Fuß in Z. 204/205
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 203:
            ws.unmerge_cells(str(mr))
    for r in (203, 204, 205):
        for c in ws[r]:
            if c.value is not None and not K.is_formula(c.value):
                c.value = None
            c.border = Border()
    ws.row_dimensions[203].height = 18
    K.footer(ws, 204, "B", "P")


def _body_row(ws, r, last, unit=False):
    """Datenzeile: Label 9 pt INK2, Werte 9 pt INK rechtsbündig, Haarlinie unten."""
    ws.row_dimensions[r].height = 17
    for c in K.iter_cells(ws, "B", r, last, r):
        c.fill = K.NOFILL
        c.border = Border(bottom=K.side("hair", K.LINE))
        if c.column == 2:
            c.font = K.font(K.T_SMALL, False, K.INK2)
            c.alignment = K.align("left", "center", 1)
        elif unit and c.column == 3:
            c.font = K.font(K.T_SMALL, False, K.MUTED)
            c.alignment = K.align("left", "center", 1)
        else:
            c.font = K.font(K.T_SMALL, False, K.INK)
            c.alignment = K.align("right", "center", 1)
            if K.is_formula(c.value):
                c.number_format = K.NUMFMT["num"]


def _mark_row(ws):
    """Anzeige-Hilfsreihe für die Markierung des Verkaufsjahres (Säule in den Bestandsdiagrammen)."""
    r = MARK_ROW
    if ws.cell(r, 2).value is None:
        K.set_text(ws.cell(r, 2), "Markierung Verkaufsjahr (Diagramm)")
    if ws.cell(r, 3).value is None:
        K.set_text(ws.cell(r, 3), "€")
    for cc in range(_col("D"), _col("AQ") + 1):
        c = ws.cell(r, cc)
        if c.value is None:
            L = K.L(cc)
            c.value = f"=IF({L}${INDEX_ROW}=Haltedauer,{L}179,0)"
        c.number_format = K.NUMFMT["num"]


def apply(wb):
    chart_contents(wb)
    if SHEET in wb.sheetnames:
        layout_sheet(wb[SHEET])
