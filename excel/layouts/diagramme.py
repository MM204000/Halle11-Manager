"""Blatt „Diagramme“ (Zellen, Raster, Anker) und Diagramm-INHALTE auf allen Blättern.

Läuft als letztes Layout-Modul. Aufgaben:
1. Inhalte aller Diagramme (alle Blätter außer Dashboard, das danach entsteht):
   einheitliche Horizonte (Cashflow 30 J., Bestände 35 J., im Cockpit alles 30 J., AfA/Steuer 20 J.), Kategorien als
   Jahreszahlen (numRef statt multiLvlStrRef), „Daten in ausgeblendeten Zellen anzeigen“ (plotVisOnly=0).
   Anzeige-Hilfsreihen (nur Darstellung, nichts verweist darauf außer Diagrammen):
   - „Einnahmen vs. Ausgaben“ (alle Blätter): flach gestapelt, Steuer als EINE Reihe mit dynamischem Namen
     (Erstattung links / Zahlung rechts – keine Nullreihe in der Legende), Summe über jeder Säule (Linienreihe ohne Linie).
     S08: vor Steuern (Schritt 09 folgt).
   - „Gesamtertrag“: echter flacher Wasserfall (unsichtbare Basis, Zu-/Abfluss/Ergebnis je als +/−-Teil, Beschriftung
     über Reihen-Namen aus Zellen – dynamisch in Excel und LibreOffice).
   - AfA-Vergleich: Balken nach Bedeutung (angewendet / anwendbar / nicht anwendbar) als drei Reihen mit #NV;
     Gutachten-Linie mit dynamischem Namen („liegt nicht vor“).
   - Restschuld/AfA: Nullreihen mit dynamischem Namen („Darlehen II (keines)“, „Sonder-AfA (keine)“).
   - Haushaltsrechnung/Vermögensaufstellung: Kurzlabels, nur Positionen > 0, absteigend sortiert (ohne Matrixformeln).
2. Blatt „Diagramme“: Seitenkopf, Abschnittsköpfe (core.section Ebene 1, „↑ Übersicht“), Raster B:P für die Diagramme,
   Datentabellen im Tabellenstil (Summenstufen über core.sum_row), Anhang „Diagrammdaten“ als eingeklappte Gliederung,
   Fuß direkt darunter.
Formeln und Werte bestehender Zellen bleiben unverändert; neu sind nur Anzeigezellen in leeren Zellen.
"""
import re

from openpyxl.chart import BarChart, LineChart, Series
from openpyxl.chart.data_source import AxDataSource, NumRef, StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import Border
from openpyxl.worksheet.properties import Outline

import core as K

SHEET = "Diagramme"
FIRST_COL, LAST_COL = "B", "P"          # Inhaltsraster (B = 34, C:P = 11)
YEAR_ROW, INDEX_ROW = 178, 177          # Kalenderjahre / Jahr-Index 1…40
MARK_ROW = 202                          # Anzeige-Hilfsreihe: Immobilienwert im Verkaufsjahr, sonst 0
SONDER_ROW = 203                        # Anzeige-Hilfsreihe: Sonder-AfA § 7b + § 7h/7i (eine Legendenposition)
GUT_ROW = 204                           # Anzeige-Hilfsreihe: AfA Gutachten kumuliert (Name dynamisch)
MODEL_ROW = 205                         # Anzeige-Hilfsreihe: AfA im Modell kumuliert (Jahre 1–10)
CF_POS, CF_NEG = 206, 207               # Anzeige-Hilfsreihen: Cashflow n. St. positiver / negativer Teil (Farbe je Vorzeichen)
FOOT_ROW = 209                          # Fuß (Haftung/Impressum) direkt unter dem eingeklappten Anhang

# Horizont je Datenzeile auf „Diagramme“ (letzte Spalte): Bestände 35 J., Cashflow 30 J., AfA/Steuer 20 J.
HORIZON = {r: "AL" for r in (179, 180, 181, 182, 183, 189, 190, 191, MARK_ROW)}
HORIZON.update({r: "AG" for r in (184, 185, 186, CF_POS, CF_NEG)})
HORIZON.update({r: "W" for r in list(range(192, 198)) + [SONDER_ROW]})
COCKPIT_LAST = "AG"                     # Cockpit: alle Zeitreihen 30 Jahre (P2-08)

# Abschnitte: (Bandzeile, Titel, Zusatz hinter dem Titel)
SECTIONS = [
    (9, "Investition, Finanzierung und Kaufpreisaufteilung", "· Zeitpunkt Kauf · Anteile"),
    (30, "Einnahmen, Ausgaben und Cashflow", "· Jahr 1 · Jahre 1–30"),
    (51, "Entwicklung von Vermögen, Darlehen und Cashflow", "· Jahre 1–35 · Jahresende"),
    (92, "Steuern und Abschreibung", "· Jahre 1–20 · € p. a."),
    (113, "Exit – Verkauf nach der geplanten Haltedauer", "· Verkaufsjahr aus Schritt 11"),
]
DATA_BAND = 138
GROUP_FIRST, GROUP_LAST = 139, CF_NEG    # eingeklappter Anhang

# Diagramm-Raster: Art → (Spalte von, Spalte bis einschließlich, erste Zeile)
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
SUM_ROWS = (161, 167)                     # „Nettoerlös (an Investor)“, „= Gesamtertrag“ → Summenstufe 1 (P1-14)

# ---- Anzeige-Hilfsblöcke (nur Diagramme; im eingeklappten Anhang rechts neben den kleinen Tabellen)
HH_HEAD, HH_N = 140, 12                   # Haushaltsrechnung: F140:N152
VA_HEAD, VA_N = 140, 9                    # Vermögensaufstellung: P140:V149
AFA_HEAD, AFA_N = 154, 8                  # AfA-Vergleich Summen: F154:I162
WF_HEAD = 163                             # Wasserfall Gesamtertrag: F163:U167
J1_HEAD = 169                             # Jahr 1 / Reihennamen: F169:H174
J1_TAX, J1_SUM, J1_SUM_VST, NAME_D2, NAME_GUT = 170, 171, 172, 173, 174

HH_SHORT = [(r"^wohnen", "Wohnen"), (r"^lebenshaltung", "Lebenshaltung"), (r"^mobilit", "Mobilität"),
            (r"^versicherung", "Versicherungen"), (r"vorsorge", "Altersvorsorge"), (r"^freizeit", "Freizeit"),
            (r"^unterhalt", "Unterhalt"), (r"kredite", "Kredite"), (r"^bewirtschaftung bestehend", "Bestandsobjekte"),
            (r"^kapitaldienst bestehend", "Bestandsdarlehen"), (r"^kapitaldienst des kalk", "Kapitaldienst neu"),
            (r"^bewirtschaftung des kalk", "Bewirtschaftung neu")]
VA_SHORT = [(r"^giro", "Girokonten"), (r"^tagesgeld", "Tages-/Festgeld"), (r"^wertpapier", "Wertpapiere"),
            (r"^lebens", "Versicherungen"), (r"^bauspar", "Bausparen"), (r"^immobilien", "Immobilien"),
            (r"^unternehmen", "Beteiligungen"), (r"^fahrzeug", "Fahrzeuge"), (r"^sonstig", "Sonstiges")]


# ============================================================================ Referenzen
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


def _literal(s):
    """Name als Literal-„Formel“ ("Name") → Text, sonst None."""
    if s.tx is not None and s.tx.strRef is not None:
        f = (s.tx.strRef.f or "").strip()
        if len(f) >= 2 and f[0] == '"' and f[-1] == '"':
            return f[1:-1].replace('""', '"')
    return None


def _tx_ref(s):
    if s.tx is None:
        return ""
    if s.tx.strRef is not None:
        return s.tx.strRef.f or ""
    return s.tx.v or ""


def _quote(sheet):
    return f"'{sheet}'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet


def _r(sheet, c1, r1, c2=None, r2=None):
    """Absolute Referenz „Blatt!$A$1:$B$2“."""
    a = f"{_quote(sheet)}!${c1}${r1}"
    return a if c2 is None else f"{a}:${c2}${r2 if r2 is not None else r1}"


def chart_kind(ws, ch):
    """Diagrammart aus den Datenbereichen (unabhängig von Nummer und Titel)."""
    rows, sheets, cols = set(), set(), set()
    for s in ch.series:
        p = _parse(_val_ref(s))
        if p:
            sheets.add(p[0])
            rows.add(p[2])
            cols.add(p[1])
    if not rows:
        return None
    if sheets == {SHEET}:
        if rows & {179, 182, 183}:
            return "bestand"
        if rows <= {180, 181}:
            return "restschuld"
        if rows & {184} and rows & {185}:
            return "cashflow"
        if rows == {185} or (rows and rows <= {CF_POS, CF_NEG}):
            return "cf_nach"
        if rows == {186}:
            return "cf_kum"
        if rows & {189, 190, 191}:
            return "kumzins"
        if rows & {192, 193, 194, 195}:
            return "afa"
        if rows & {196, 197}:
            return "steuer"
        if rows & set(range(170, 177)):
            return "jahr1"
        if cols <= {"C"}:
            single = {141: "invest", 147: "finanz", 152: "kpa", 158: "exit", 164: "ertrag"}
            for r, k in single.items():
                if r in rows:
                    return k
        if rows & set(range(164, 168)):
            return "ertrag"
    return None


def _set_year_categories(s, sheet, c1, c2, row):
    s.cat = AxDataSource(numRef=NumRef(f=f"{sheet}!${c1}${row}:${c2}${row}"))


def _label(text=None, ref=None):
    return SeriesLabel(strRef=StrRef(f=ref)) if ref else SeriesLabel(v=text)


def _series(ref, title=None, title_ref=None, cat_ref=None, cat_num=False):
    s = Series(ref)
    s.tx = _label(title, title_ref)
    if cat_ref:
        s.cat = AxDataSource(numRef=NumRef(f=cat_ref)) if cat_num else AxDataSource(strRef=StrRef(f=cat_ref))
    return s


def _replace(ws, idx, old, new):
    new.anchor = old.anchor
    ws._charts[idx] = new


# ============================================================================ Neuaufbau einzelner Diagramme
def _jahr1_chart(with_tax):
    """„Einnahmen vs. Ausgaben“ flach gestapelt: Miete · Bewirtschaftung · Zinsen · Tilgung (· Steuer) + Summe."""
    cats = _r(SHEET, "C", 169, "D", 169)
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 100
    for row, name in ((170, "Nettokaltmiete"), (172, "Bewirtschaftung"), (173, "Zinsen"), (174, "Tilgung")):
        bar.series.append(_series(_r(SHEET, "C", row, "D", row), name, cat_ref=cats))
    if with_tax:
        bar.series.append(_series(_r(SHEET, "G", J1_TAX, "H", J1_TAX), title_ref=_r(SHEET, "F", J1_TAX), cat_ref=cats))
    line = LineChart()
    row = J1_SUM if with_tax else J1_SUM_VST
    line.series.append(_series(_r(SHEET, "G", row, "H", row), "Summe", cat_ref=cats))
    line.y_axis.axId = bar.y_axis.axId
    line.x_axis = bar.x_axis
    bar += line
    return bar


def _ertrag_chart():
    """Gesamtertrag als echter Wasserfall: unsichtbare Basis + Zu-/Abfluss/Ergebnis (je +/−-Teil), Beschriftungen als
    vier Linienreihen (Wert nur an „ihrer“ Kategorie, Name = Beschriftungstext aus Zelle)."""
    cats = _r(SHEET, "F", WF_HEAD + 1, "F", WF_HEAD + 4)
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 60
    for c, name in (("I", "Basis"), ("J", "Zufluss"), ("K", "Zufluss (−)"), ("L", "Abfluss"), ("M", "Abfluss (−)"),
                    ("N", "Ergebnis"), ("O", "Ergebnis (−)")):
        bar.series.append(_series(_r(SHEET, c, WF_HEAD + 1, c, WF_HEAD + 4), name, cat_ref=cats))
    line = LineChart()
    for i, c in enumerate(("R", "S", "T", "U")):
        line.series.append(_series(_r(SHEET, c, WF_HEAD + 1, c, WF_HEAD + 4), title_ref=_r(SHEET, "Q", WF_HEAD + 1 + i),
                                   cat_ref=cats))
    line.y_axis.axId = bar.y_axis.axId
    line.x_axis = bar.x_axis
    bar += line
    return bar


def chart_contents(wb):
    for ws in wb.worksheets:
        if ws.title == "Dashboard":
            continue
        for idx, ch in enumerate(list(ws._charts)):
            ch.visible_cells_only = False          # Daten in ausgeblendeten Zeilen/Spalten weiter anzeigen
            for s in ch.series:                    # Literalnamen (<c:f>"Name"</c:f>) → <c:v>Name</c:v> (Excel-Reparatur)
                lit = _literal(s)
                if lit is not None:
                    s.tx = _label(lit)
            kind = chart_kind(ws, ch)
            if kind == "cf_nach":
                last = COCKPIT_LAST
                cats = f"{SHEET}!$D${YEAR_ROW}:${last}${YEAR_ROW}"
                ch.series[:] = [_series(_r(SHEET, "D", r, last, r), name, cat_ref=cats, cat_num=True)
                                for r, name in ((CF_POS, "Überschuss"), (CF_NEG, "Unterdeckung"))]
                ch.grouping, ch.overlap = "stacked", 100
                for i, s in enumerate(ch.series):
                    s.idx, s.order = i, i
                continue
            if kind == "jahr1":
                new = _jahr1_chart(with_tax=not ws.title.startswith("S08"))   # S08: vor Steuern
                new.visible_cells_only = False
                _replace(ws, idx, ch, new)
                continue
            if kind == "ertrag":
                new = _ertrag_chart()
                new.visible_cells_only = False
                _replace(ws, idx, ch, new)
                continue
            if ws.title == "AfA-Vergleich" and isinstance(ch, BarChart):
                _afa_sum_series(ch)
                continue
            if ws.title in ("Haushaltsrechnung", "Vermögensaufstellung") and isinstance(ch, BarChart):
                _ranked_series(ws, ch)
                continue
            if ws.title == "Sensitivität":
                for s in ch.series:                 # „Miete -10%“ → „Miete −10 %“ (Literalnamen im Diagramm)
                    if s.tx is not None and s.tx.v:
                        s.tx.v = re.sub(r"\s*-\s*(\d)", r" −\1", s.tx.v).replace("%", " %").replace("  ", " ")
            if kind == "afa":
                keep = []
                sonder_done = False
                for s in ch.series:
                    p = _parse(_val_ref(s))
                    if p and p[2] in (193, 194):
                        if not sonder_done:
                            s.val.numRef.f = _r(SHEET, "D", SONDER_ROW, "W", SONDER_ROW)
                            s.tx = _label(ref=_r(SHEET, "B", SONDER_ROW))
                            keep.append(s)
                            sonder_done = True
                        continue
                    keep.append(s)
                ch.series[:] = keep
            if ws.title == "AfA-Vergleich" and isinstance(ch, LineChart):
                first = _parse(_val_ref(ch.series[0])) if ch.series else None
                if first and not any("modell" in (x.tx.v or "").lower() for x in ch.series if x.tx is not None):
                    # angewendete Variante als eigene (kräftige) Linie aus der Summenzeile „Im Modell angewendet“
                    s_m = _series(_r(SHEET, first[1], MODEL_ROW, first[3], MODEL_ROW), "Im Modell angewendet")
                    s_m.cat = ch.series[0].cat
                    ch.series.append(s_m)
                    _PENDING["model"] = first
                for s in ch.series:
                    p = _parse(_val_ref(s))
                    name = (s.tx.v if s.tx is not None and s.tx.v else "") or ""
                    if p and "gutachten" in name.lower():
                        s.val.numRef.f = _r(SHEET, p[1], GUT_ROW, p[3], GUT_ROW)
                        s.tx = _label(ref=_r(SHEET, "F", NAME_GUT))
                        _gut_row(wb, p)
            for i, s in enumerate(ch.series):
                s.idx, s.order = i, i
                p = _parse(_val_ref(s))
                if not p:
                    continue
                sh, c1, r1, c2, r2 = p
                if sh == SHEET and r1 == r2 and r1 in HORIZON and c1 == "D":
                    last = HORIZON[r1]
                    if ws.title == "Cockpit":
                        last = COCKPIT_LAST
                    s.val.numRef.f = f"{SHEET}!$D${r1}:${last}${r1}"
                    _set_year_categories(s, SHEET, "D", last, YEAR_ROW)
                    if kind == "restschuld" and r1 == 181:
                        s.tx = _label(ref=_r(SHEET, "F", NAME_D2))
                    continue
                cr = _parse(_cat_ref(s))
                if cr and s.cat is not None and s.cat.multiLvlStrRef is not None and cr[2] == cr[4]:
                    # einzeilige Kategorien (Jahre, Wertsteigerung) als Zahlenreihe statt Mehrebenen-Text
                    _set_year_categories(s, _quote(cr[0]), cr[1], cr[3], cr[2])


# ---- Hilfsblöcke, die Diagramme auf fremden Blättern speisen (Zellen liegen auf „Diagramme“)
_PENDING = {}


def _afa_sum_series(ch):
    """AfA-Vergleich „Summe AfA je Variante“: drei Reihen nach Bedeutung, Kategorien mit Kurzlabel „Im Modell“."""
    s0 = ch.series[0]
    p, c = _parse(_val_ref(s0)), _parse(_cat_ref(s0))
    if not p or not c:
        return
    _PENDING["afa"] = (p, c)
    r1, r2 = AFA_HEAD + 1, AFA_HEAD + (p[4] - p[2]) + 1
    cats = _r(SHEET, "F", r1, "F", r2)
    ch.series[:] = [_series(_r(SHEET, col, r1, col, r2), name, cat_ref=cats)
                    for col, name in (("G", "Im Modell angewendet"), ("H", "Anwendbar"), ("I", "Nicht anwendbar"))]
    ch.grouping, ch.overlap, ch.gapWidth = "clustered", 100, 40
    for i, s in enumerate(ch.series):
        s.idx, s.order = i, i


def _ranked_series(ws, ch):
    """Haushalt/Vermögen: nur Positionen > 0, absteigend, Kurzlabels (Haushalt: bestehend dunkel, neues Objekt hell)."""
    s0 = ch.series[0]
    p, c = _parse(_val_ref(s0)), _parse(_cat_ref(s0))
    if not p or not c:
        return
    hh = ws.title == "Haushaltsrechnung"
    labels = [ws.cell(r, K.col(c[1])).value for r in range(c[2], c[4] + 1)]
    _PENDING["hh" if hh else "va"] = (p, c, labels)
    if hh:
        r1, r2 = HH_HEAD + 1, HH_HEAD + (p[4] - p[2]) + 1
        cats = _r(SHEET, "L", r1, "L", r2)
        ch.series[:] = [_series(_r(SHEET, "M", r1, "M", r2), "Bestehender Haushalt", cat_ref=cats),
                        _series(_r(SHEET, "N", r1, "N", r2), "Neues Objekt", cat_ref=cats)]
    else:
        r1, r2 = VA_HEAD + 1, VA_HEAD + (p[4] - p[2]) + 1
        cats = _r(SHEET, "U", r1, "U", r2)
        ch.series[:] = [_series(_r(SHEET, "V", r1, "V", r2), "Vermögenswerte", cat_ref=cats)]
    ch.grouping, ch.overlap, ch.gapWidth = "clustered", 100, 40
    for i, s in enumerate(ch.series):
        s.idx, s.order = i, i


def _gut_row(wb, p):
    _PENDING["gut"] = p


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


def _put(ws, coord, value, fmt=None, h="right", bold=False):
    """Anzeige-Hilfszelle: nur in leere Zellen (oder eigene frühere Formel); 8 pt grau."""
    c = ws[coord]
    if c.value is not None and not (K.is_formula(c.value) or isinstance(c.value, (int, float)) or c.value == value):
        raise ValueError(f"Diagramme!{coord} ist belegt: {c.value!r}")
    c.value = value
    c.font = K.font(K.T_MICRO, bold, K.MUTED)
    c.alignment = K.align(h, "center", 1)
    c.fill = K.NOFILL
    c.border = Border(bottom=K.side("hair", K.LINE))
    if fmt:
        c.number_format = fmt


def _helper_head(ws, row, c1, c2, labels):
    K.table_head(ws, row, c1, c2, {k: (v, "left" if i == 0 else "right") for i, (k, v) in enumerate(labels.items())})
    for c in K.iter_cells(ws, c1, row, c2, row):
        c.font = K.font(K.T_MICRO, True, K.BLUE)


def _short(text, table):
    t = (text or "").strip().lower()
    for pat, short in table:
        if re.search(pat, t):
            return short
    return re.sub(r"\s*\(.*?\)", "", text or "")[:18]


def helper_blocks(ws):
    """Alle Anzeige-Hilfsreihen der Diagramme (reine Darstellung; die Rechenzellen bleiben unberührt)."""
    num = K.NUMFMT["num"]
    # ---- Jahr 1: Steuer als eine Reihe, Summen, dynamische Reihennamen
    _helper_head(ws, J1_HEAD, "F", "H", {"F": "Hilfsreihen Jahr 1 / Namen", "G": "Einnahmen", "H": "Ausgaben"})
    _put(ws, f"F{J1_TAX}", '=IF(D175>0,"Steuerzahlung","Steuererstattung")', h="left")
    _put(ws, f"G{J1_TAX}", "=C171", num)
    _put(ws, f"H{J1_TAX}", "=D175", num)
    _put(ws, f"F{J1_SUM}", "Summe", h="left")
    _put(ws, f"G{J1_SUM}", "=SUM(C170:C175)", num)
    _put(ws, f"H{J1_SUM}", "=SUM(D170:D175)", num)
    _put(ws, f"F{J1_SUM_VST}", "Summe vor Steuern", h="left")
    _put(ws, f"G{J1_SUM_VST}", "=C170+C172+C173+C174", num)
    _put(ws, f"H{J1_SUM_VST}", "=D170+D172+D173+D174", num)
    _put(ws, f"F{NAME_D2}", '=IF(MAX(D181:AL181)>0,"Darlehen II","Darlehen II (keines)")', h="left")
    _put(ws, f"F{NAME_GUT}", f'=IF(MAX(D{GUT_ROW}:M{GUT_ROW})>0,"Gutachten (RND)","Gutachten (RND) – liegt nicht vor")',
         h="left")

    # ---- Wasserfall Gesamtertrag (C164:C167: Kum. Cashflow, Nettoerlös, − Eigenkapital, = Gesamtertrag)
    _helper_head(ws, WF_HEAD, "F", "U", {"F": "Wasserfall (Diagramm)", "G": "Start", "H": "Ende", "I": "Basis",
                                          "J": "Zufl. +", "K": "Zufl. −", "L": "Abfl. +", "M": "Abfl. −",
                                          "N": "Erg. +", "O": "Erg. −", "P": "Oberkante", "Q": "Beschriftung",
                                          "R": "Lbl 1", "S": "Lbl 2", "T": "Lbl 3", "U": "Lbl 4"})
    names = ["Cashflow kumuliert", "Nettoerlös Verkauf", "Eigenkapital", "Gesamtertrag"]
    for i in range(4):
        r = WF_HEAD + 1 + i
        total = i == 3
        _put(ws, f"F{r}", names[i], h="left")
        _put(ws, f"G{r}", "=0" if i == 0 or total else f"=H{r - 1}", num)
        _put(ws, f"H{r}", f"=C{r}" if total else f"=G{r}+C{r}", num)
        lo, hi = f"MIN(G{r},H{r})", f"MAX(G{r},H{r})"
        up = f"IF({lo}>=0,{hi}-{lo},MAX(0,{hi}))"
        dn = f"IF({hi}<=0,{lo}-{hi},MIN(0,{lo}))"
        _put(ws, f"I{r}", f"=IF({lo}>=0,{lo},IF({hi}<=0,{hi},0))", num)
        if total:
            for c in "JKLM":
                _put(ws, f"{c}{r}", "=0", num)
            _put(ws, f"N{r}", f"={up}", num)
            _put(ws, f"O{r}", f"={dn}", num)
        else:
            _put(ws, f"J{r}", f"=IF(C{r}>=0,{up},0)", num)
            _put(ws, f"K{r}", f"=IF(C{r}>=0,{dn},0)", num)
            _put(ws, f"L{r}", f"=IF(C{r}<0,{up},0)", num)
            _put(ws, f"M{r}", f"=IF(C{r}<0,{dn},0)", num)
            _put(ws, f"N{r}", "=0", num)
            _put(ws, f"O{r}", "=0", num)
        _put(ws, f"P{r}", f"=MAX(0,{hi})", num)
        sign = f'IF(C{r}<0,"−","")' if total else f'IF(C{r}<0,"−","+")'
        _put(ws, f"Q{r}", f'={sign}&FIXED(ABS(C{r})/1000,1)&" T€"')
        for j, c in enumerate("RSTU"):
            _put(ws, f"{c}{r}", f"=P{r}" if j == i else "=0", num)

    # ---- AfA-Vergleich: Summe je Variante nach Bedeutung
    if "afa" in _PENDING:
        (sh, vc, vr1, _, vr2), (csh, cc, cr1, _, _) = _PENDING["afa"]
        _helper_head(ws, AFA_HEAD, "F", "I", {"F": "AfA je Variante (Diagramm)", "G": "Modell", "H": "Anwendbar",
                                               "I": "Nicht anw."})
        stat = "C"          # Spalte „Anwendbar für dieses Objekt?“ (Ja …/Nein/Nur mit Gutachten)
        for i in range(vr2 - vr1 + 1):
            r, sr = AFA_HEAD + 1 + i, vr1 + i
            name, val, st = _r(csh, cc, cr1 + i), _r(sh, vc, sr), _r(sh, stat, sr)
            model = f'{name}="Im Modell angewendet"'
            _put(ws, f"F{r}", f'=IF({model},"Im Modell",{name})', h="left")
            _put(ws, f"G{r}", f"=IF({model},{val},0)", num)
            _put(ws, f"H{r}", f'=IF(AND(NOT({model}),LEFT({st},2)="Ja"),{val},0)', num)
            _put(ws, f"I{r}", f'=IF(AND(NOT({model}),LEFT({st},2)<>"Ja"),{val},0)', num)

    # ---- Haushaltsrechnung / Vermögensaufstellung: absteigend, nur > 0
    def ranked(key, head, c0, short_table):
        if key not in _PENDING:
            return
        (sh, vc, vr1, _, vr2), (csh, cc, cr1, _, _), labels = _PENDING[key]
        n = vr2 - vr1 + 1
        if key == "hh":      # Haushalt: zwei Reihen (bestehend / neues Objekt)
            cols = [K.L(_col(c0) + k) for k in range(9)]
            lab, val, key_, obj, kk, pos, lab_k, v1, v2 = cols
            names = ("Haushalt (Diagramm)", "Wert", "Sortier", "Objekt", "Rang", "Position", "Label", "Bestand",
                     "Objekt neu")
        else:                # Vermögen: eine Reihe
            cols = [K.L(_col(c0) + k) for k in range(7)]
            lab, val, key_, kk, pos, lab_k, v1 = cols
            obj = v2 = None
            names = ("Vermögen (Diagramm)", "Wert", "Sortier", "Rang", "Position", "Label", "Wert sortiert")
        heads = dict(zip(cols, names))
        _helper_head(ws, head, cols[0], cols[-1], heads)
        r1, r2 = head + 1, head + n
        rngv, rngk = f"${val}${r1}:${val}${r2}", f"${key_}${r1}:${key_}${r2}"
        for i in range(n):
            r = r1 + i
            _put(ws, f"{lab}{r}", _short(labels[i], short_table), h="left")
            _put(ws, f"{val}{r}", f"=N({_r(sh, vc, vr1 + i)})", num)
            _put(ws, f"{key_}{r}", f"=IF({val}{r}>0,{val}{r}+({n + 1}-{i + 1})/1000,0)", "0.000")
            if key == "hh":
                _put(ws, f"{obj}{r}", f'=IF(ISNUMBER(SEARCH("kalkulierten",{_r(csh, cc, cr1 + i)})),1,0)', num)
            _put(ws, f"{kk}{r}", i + 1, "0")
            _put(ws, f"{pos}{r}", f'=IF({kk}{r}<=COUNTIF({rngv},">0"),MATCH(LARGE({rngk},{kk}{r}),{rngk},0),"")', "0")
            _put(ws, f"{lab_k}{r}", f'=IF({pos}{r}="","",INDEX(${lab}${r1}:${lab}${r2},{pos}{r}))', h="left")
            if key == "hh":
                objk = f"INDEX(${obj}${r1}:${obj}${r2},{pos}{r})"
                valk = f"INDEX({rngv},{pos}{r})"
                _put(ws, f"{v1}{r}", f'=IF({pos}{r}="",0,IF({objk}=1,0,{valk}))', num)
                _put(ws, f"{v2}{r}", f'=IF({pos}{r}="",0,IF({objk}=1,{valk},0))', num)
            else:
                _put(ws, f"{v1}{r}", f'=IF({pos}{r}="",0,INDEX({rngv},{pos}{r}))', num)

    ranked("hh", HH_HEAD, "F", HH_SHORT)
    ranked("va", VA_HEAD, "P", VA_SHORT)


def time_helper_rows(ws):
    """Zeilen 203/204 der Zeitreihe: Sonder-AfA gesamt, AfA Gutachten (#NV ohne Gutachten)."""
    num = K.NUMFMT["num"]
    ws.cell(SONDER_ROW, 2).value = f'=IF(SUM(D{SONDER_ROW}:W{SONDER_ROW})>0,"Sonder-AfA (§ 7b, § 7h/7i)","Sonder-AfA (keine)")'
    K.set_text(ws.cell(SONDER_ROW, 3), "€")
    for cc in range(_col("D"), _col("AQ") + 1):
        L = K.L(cc)
        ws.cell(SONDER_ROW, cc).value = f"={L}193+{L}194"
        ws.cell(SONDER_ROW, cc).number_format = num
    gut = _PENDING.get("gut")
    K.set_text(ws.cell(GUT_ROW, 2), "AfA Gutachten kumuliert (Diagramm)")
    K.set_text(ws.cell(GUT_ROW, 3), "€")
    if gut:
        sh, c1, r1, c2, _ = gut
        src = [K.L(c) for c in range(_col(c1), _col(c2) + 1)]
        any_ref = _r(sh, c1, r1, c2, r1)
        for i, L in enumerate(src):
            tgt = ws.cell(GUT_ROW, _col("D") + i)
            tgt.value = f"={_r(sh, L, r1)}"
            tgt.number_format = num


def more_helper_rows(ws):
    """Zeilen 205–207: AfA im Modell kumuliert, Cashflow n. St. positiver/negativer Teil."""
    num = K.NUMFMT["num"]
    K.set_text(ws.cell(MODEL_ROW, 2), "AfA im Modell kumuliert (Diagramm)")
    K.set_text(ws.cell(MODEL_ROW, 3), "€")
    model = _PENDING.get("model")
    afa = _PENDING.get("afa")
    if model and afa:
        sh = model[0]
        mrow = afa[0][4]                               # letzte Zeile der Summentabelle = „Im Modell angewendet“
        c0 = _col(model[1])
        for i in range(_col(model[3]) - c0 + 1):
            L = K.L(c0 + i)
            src = _r(sh, L, mrow)
            ws.cell(MODEL_ROW, _col("D") + i).value = f"={src}" if i == 0 else f"={K.L(_col('D') + i - 1)}{MODEL_ROW}+{src}"
            ws.cell(MODEL_ROW, _col("D") + i).number_format = num
    for r, name, fn in ((CF_POS, "Cashflow n. St. – Überschuss (Diagramm)", "MAX"),
                        (CF_NEG, "Cashflow n. St. – Unterdeckung (Diagramm)", "MIN")):
        K.set_text(ws.cell(r, 2), name)
        K.set_text(ws.cell(r, 3), "€")
        for cc in range(_col("D"), _col("AQ") + 1):
            L = K.L(cc)
            ws.cell(r, cc).value = f"={fn}(0,{L}185)"
            ws.cell(r, cc).number_format = num


def layout_sheet(ws):
    # ---- Raster: B = 34, C:AQ = 11 (Zeitreihe ohne Rhythmuswechsel ab Jahr 17)
    ws.column_dimensions["A"].width = 4.5
    ws.column_dimensions["B"].width = 34
    for cc in range(_col("C"), _col("AQ") + 1):
        ws.column_dimensions[K.L(cc)].width = 11

    # ---- Seitenkopf Z. 5–7 (P1-18) und Zeile 8 für die Sprungleiste (Formen, navigation.py)
    for r in (5, 6, 7):
        for c in K.iter_cells(ws, "B", r, "P", r):
            if c.column > 2 and not K.is_formula(c.value):
                c.value = None
    K.page_header(ws, "B", "P", "Diagramme", "Diagramme",
                  "Alle Auswertungen als Grafik – aktualisieren sich automatisch aus den Eingaben "
                  "(Datenbasis: Projektion, Finanzierung, Steuern).",
                  context=("=Obj_Name", '=IF(Erstellt_fuer="","","Erstellt für "&Erstellt_fuer)'), context_col="P")
    ws["B7"].alignment = K.align("left", "top")
    ws["P6"].font = K.font(K.T_BODY, True, K.NAVY)
    ws["P7"].font = K.font(K.T_SMALL, False, K.MUTED)
    ws.row_dimensions[8].height = 30
    for c in ws[8]:
        c.value = None if c.column > 1 else c.value

    # ---- Abschnittsköpfe Ebene 1 (core.section) B:P mit Rücksprung „↑ Übersicht“ (P3-03)
    _unmerge_rows(ws, [row for row, *_ in SECTIONS] + [DATA_BAND])
    for row, title, extra in SECTIONS + [(DATA_BAND, "Diagrammdaten",
                                          "· Anhang eingeklappt – über [+] am linken Rand öffnen · bitte nicht ändern")]:
        _clear_row_style(ws, row, "Q", "U")
        for c in K.iter_cells(ws, "C", row, LAST_COL, row):
            if not K.is_formula(c.value):
                c.value = None
        K.section(ws, row, FIRST_COL, LAST_COL, title, level=1, meta="↑ Übersicht", extra=extra)
        top = ws.cell(row, _col(LAST_COL))
        K.text_link(top, "↑ Übersicht", SHEET, "A4", size=K.T_MICRO, bold=False, tooltip="Zum Seitenanfang")
        top.alignment = K.align("right", "center", 1)

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
    for r in SUM_ROWS:                                # Summenstufe 1 (P1-14)
        K.sum_row(ws, r, "B", "C", "sub")

    # Zeitreihen (Jahr 1–40): Kopf Jahr-Index, Kalenderjahr als zweite Kopfzeile (fett, E7EEF7)
    K.table_head(ws, INDEX_ROW, "B", "AQ", {"B": ("Zeitreihen (Jahr 1–40)", "left"), "C": ("Einheit", "left")})
    for cc in range(_col("D"), _col("AQ") + 1):
        c = ws.cell(INDEX_ROW, cc)
        c.number_format = '"Jahr "0'
        c.alignment = K.align("right", "center", 1)
    K.set_text(ws.cell(YEAR_ROW, 3), "Jahr")
    for mr in list(ws.merged_cells.ranges):          # alter Fuß (Z. 203 ff.) → Platz für Hilfszeilen und neuen Fuß
        if mr.max_row >= SONDER_ROW:
            ws.unmerge_cells(str(mr))
    for r in range(SONDER_ROW, FOOT_ROW + 2):
        for c in ws[r]:
            if c.value is not None and not K.is_formula(c.value):
                c.value = None
            c.border = Border()
    _mark_row(ws)
    time_helper_rows(ws)
    more_helper_rows(ws)
    for r in range(YEAR_ROW, CF_NEG + 1):
        _body_row(ws, r, "AQ", unit=True)
    for c in K.iter_cells(ws, "B", YEAR_ROW, "AQ", YEAR_ROW):
        c.fill = K.fill(K.TINT)
        c.font = K.font(K.T_SMALL, True, K.NAVY)
        c.border = Border(bottom=K.side("thin", K.ACCENT))
    for cc in range(_col("D"), _col("AQ") + 1):
        ws.cell(YEAR_ROW, cc).number_format = "0"
    for cc in range(_col("D"), _col("AQ") + 1):
        ws.cell(201, cc).number_format = K.NUMFMT["pct1"]
    for r in (MARK_ROW, SONDER_ROW, GUT_ROW, MODEL_ROW, CF_POS, CF_NEG):   # Hilfszeilen einheitlich: 9 pt kursiv grau
        K.memo(ws, r, "B", "AQ")
        for cc in range(_col("C"), _col("AQ") + 1):
            ws.cell(r, cc).alignment = K.align("right" if cc > 3 else "left", "center", 1)
        ws.cell(r, 2).alignment = K.align("left", "center", 1)

    # ---- Hilfsblöcke rechts neben den kleinen Tabellen
    helper_blocks(ws)

    # ---- Anhang einklappen (Gliederung wie Sensitivität), Fuß direkt darunter
    for r in range(GROUP_FIRST, GROUP_LAST + 1):
        ws.row_dimensions[r].outline_level = 1
        ws.row_dimensions[r].hidden = True
    if ws.sheet_properties.outlinePr is None:
        ws.sheet_properties.outlinePr = Outline(summaryBelow=True, summaryRight=True)
    else:
        ws.sheet_properties.outlinePr.summaryBelow = True
    ws.row_dimensions[CF_NEG + 1].height = 18
    K.footer(ws, FOOT_ROW, "B", "P")
    K.cf_close(ws)


def _body_row(ws, r, last, unit=False):
    """Datenzeile: Label 9 pt INK2, Werte 9 pt INK rechtsbündig, Haarlinie unten."""
    ws.row_dimensions[r].height = K.H_ROW
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
    _PENDING.clear()
    chart_contents(wb)
    if SHEET in wb.sheetnames:
        layout_sheet(wb[SHEET])
