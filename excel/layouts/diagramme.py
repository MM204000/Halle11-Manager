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
   - AfA-Vergleich: Balken nach Bedeutung (angewendet / anwendbar / nicht anwendbar) als drei Reihen, Kategorie ohne
     Wert mit „– n. v.“; Linien je Variante plus „Im Modell angewendet“ (Farbe nach Anwendbarkeit in finish_pro).
   - Restschuld/AfA: leere Reihen fallen in finish_pro aus der Legende; Fußnote unter dem Diagramm („Kein zweites
     Darlehen …“, „Keine Sonder-AfA …“).
   - „Cashflow vor und nach Steuern“: Kombidiagramm (Säule n. St. + Linie v. St.); Kreise: Legendentexte
     „Kurzname · 12 %“ für die Kleinstsegmente.
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
MARK_ROW = 202                          # Anzeige-Hilfsreihe: Immobilienwert im Verkaufsjahr, sonst 0 (Lot + Beschriftung)
SONDER_ROW = 203                        # Anzeige-Hilfsreihe: Sonder-AfA § 7b + § 7h/7i (eine Legendenposition)
ZB_ROW = 204                            # Anzeige-Hilfsreihe: Restschuld Darlehen I im Jahr der Zinsbindung, sonst 0
D1_ROW = 176                            # Anzeige-Hilfsreihe: Restschuld Darlehen I ohne das Zinsbindungsjahr
GAP_ROW = 175                           # Jahr 1: „Lücke“ (Ausgaben über Einnahmen) G:H nach / I:J vor Steuern
MODEL_ROW = 205                         # Anzeige-Hilfsreihe: AfA im Modell kumuliert (Jahre 1–10)
CF_POS, CF_NEG = 206, 207               # Anzeige-Hilfsreihen: Cashflow n. St. positiver / negativer Teil (Farbe je Vorzeichen)
FOOT_ROW = 209                          # Fuß (Haftung/Impressum) direkt unter dem eingeklappten Anhang

# Horizont je Datenzeile auf „Diagramme“ (letzte Spalte): Bestände 35 J., Cashflow 30 J., AfA/Steuer 20 J.
HORIZON = {r: "AL" for r in (179, 180, 181, 182, 183, 189, 190, 191, MARK_ROW, ZB_ROW, D1_ROW)}
HORIZON.update({r: "AG" for r in (184, 185)})
CF_NACH_LAST = "AL"                     # S12 „Cashflow nach Steuern“: Jahre 1–35 (Umschlag nach Volltilgung sichtbar, P1-07)
HORIZON[186] = "AL"                     # kumulierter Cashflow: 35 J. wie „Kumulierte Zinsen …“ daneben (P44)
HORIZON.update({r: "W" for r in list(range(192, 198)) + [SONDER_ROW]})
COCKPIT_LAST = "AG"                     # Cockpit: alle Zeitreihen 30 Jahre (P2-08)

# Abschnitte: (Bandzeile, Titel, Zusatz hinter dem Titel)
SECTIONS = [
    (9, "Investition, Finanzierung und Kaufpreisaufteilung", None),
    (30, "Einnahmen, Ausgaben und Cashflow", None),
    (51, "Entwicklung von Vermögen, Darlehen und Cashflow", None),
    (92, "Steuern und Abschreibung", None),
    (113, "Exit – Verkauf nach der geplanten Haltedauer", None),
]
DATA_BAND = 138
DATA_HINT = "Datengrundlage der Diagramme: eingeklappt ([+] links) – nur lesen"
GROUP_FIRST, GROUP_LAST = 139, CF_NEG    # eingeklappter Anhang

# Diagramm-Raster (P3-05): nur 3 × 1/3 (B–D | F–J | L–P) oder 1/3 + 2/3 (B–D | F–P) – jedes rechte Diagramm beginnt in F.
# Art → (Spalte von, Spalte bis einschließlich, erste Zeile); darüber je Diagramm ein Unterabschnitt (Titel + Meta).
CHART_ROWS = 16
GRID = {
    "invest": ("B", "D", 11), "finanz": ("F", "J", 11), "kpa": ("L", "P", 11),
    "jahr1": ("B", "D", 32), "cashflow": ("F", "P", 32),
    "bestand": ("B", "D", 53), "restschuld": ("F", "P", 53),
    "kumzins": ("B", "D", 71), "cf_kum": ("F", "P", 71),
    "afa": ("B", "D", 94), "steuer": ("F", "P", 94),
    "exit": ("B", "D", 115), "ertrag": ("F", "P", 115),
}
# Unterabschnitt je Diagramm (P2-10): Versalien 8,5 pt fett 1D4F8A mit Linie 4A86C8, Meta rechts (Horizont · Einheit)
SUBHEADS = {
    "invest": ("Gesamtinvestition", "Anteile in %"),
    "finanz": ("Finanzierungsstruktur", "Anteile in %"),
    "kpa": ("Kaufpreisaufteilung", "Anteile in %"),
    "jahr1": ("Einnahmen vs. Ausgaben", "Jahr 1 · € / Monat"),
    "cashflow": ("Cashflow vor und nach Steuern", "Jahre 1–30 · T€ p. a."),
    "bestand": ("Vermögen und Restschuld", "Jahre 1–35 · T€"),
    "restschuld": ("Restschuld am Jahresende", "Jahre 1–35 · T€ · Jahresende"),
    "kumzins": ("Zinsen, Tilgung und Steuer kumuliert", "Jahre 1–35 · T€"),
    "cf_kum": ("Kumulierter Cashflow nach Steuern", "Jahre 1–35 · T€"),
    "afa": ("Abschreibungen nach Komponenten", "Jahre 1–20 · T€ p. a."),
    "steuer": ("Steuerliches Ergebnis und Steuer", "Jahre 1–20 · T€ p. a."),
    "exit": ("Verwendung des Verkaufserlöses", "Anteile in %"),
    "ertrag": ("Gesamtertrag nach Steuern", "Brücke · T€"),
}
SUBHEAD_H = 24
# Zeilenhöhen der Diagrammzone (Abstände): Zeile → pt
SPACER_ROWS = {**{r: 13 for r in range(11, 27)},           # Kreisreihe etwas niedriger (weniger Leerraum um die Kreise)
               27: 10, 28: 10, 29: 10, 48: 10, 49: 10, 50: 10, 69: 9,
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
HH_HEAD, HH_N = 140, 12                   # Haushaltsrechnung: F140:O152 (O = Rang je Position)
VA_HEAD, VA_N = 140, 9                    # Vermögensaufstellung: P140:W149 (W = Rang je Position)
# Diagrammplätze (P2-09): die größten Positionen einzeln, ab dem letzten Platz „Übrige (n)“ – statt 12 bzw. 9 Plätzen
# mit leeren Zeilen unter den Balken
HH_SLOTS, VA_SLOTS = 8, 6
AFA_HEAD, AFA_N = 154, 8                  # AfA-Vergleich Summen: F154:I162
WF_HEAD = 163                             # Wasserfall Gesamtertrag: F163:U167
J1_HEAD = 169                             # Jahr 1 / Reihennamen: F169:H174
J1_TAX, J1_SUM, J1_SUM_VST, NAME_D2, NAME_GUT = 170, 171, 172, 173, 174

HH_SHORT = [(r"^wohnen", "Wohnen"), (r"^lebenshaltung", "Lebenshaltung"), (r"^mobilit", "Mobilität"),
            (r"^versicherung", "Versicherungen"), (r"vorsorge", "Altersvorsorge"), (r"^freizeit", "Freizeit"),
            (r"^unterhalt", "Unterhalt"), (r"kredite", "Kredite"), (r"^bewirtschaftung bestehend", "Bestandsobjekte"),
            (r"^kapitaldienst bestehend", "Bestandsdarlehen"), (r"^kapitaldienst des kalk", "Kapitaldienst neu"),
            (r"^bewirtschaftung des kalk", "Bewirtsch. neu")]
VA_SHORT = []                             # Vermögen: Kategorien wortgleich zu Spalte B (ohne Klammerzusatz, P19)

# ---- Kreise: Legendentexte „Name · 12 %“ (nur Anzeige; Beschriftung der Kleinstsegmente in der Legende, P17)
PIE_HEAD, PIE_COL = 140, "AI"
PIE_SHORT = {
    "invest": [(r"^kaufpreis", "Kaufpreis"), (r"^kaufneben", "Nebenkosten"), (r"^finanzierungsneben", "Finanzierungskosten"),
               (r"^ma(ß|ss)nahmen", "Maßnahmen")],
    "finanz": [(r"^darlehen ii\b", "Darlehen II"), (r"^darlehen i\b", "Darlehen I"), (r"^eigenkapital", "Eigenkapital")],
    "kpa": [(r"^geb", "Gebäude"), (r"^grund", "Boden"), (r"^beweg", "Inventar"), (r"r(ü|ue)cklage", "Rücklage")],
    "kanc": [(r"^grunderwerb", "Grunderwerbsteuer"), (r"^notar", "Notar"), (r"^grundbuch", "Grundbuch"),
             (r"^makler", "Makler"), (r"^sonst", "Sonstige")],
    "bewirt": [(r"^hausgeld", "Hausgeld"), (r"erhaltungsr", "Rücklage WEG"), (r"^verwaltung", "Verwaltung"),
               (r"instandhaltung", "Instandh."), (r"^mietausfall", "Mietausfall"), (r"werbungskosten", "Werbungsk.")],
    "exit": [(r"^verkaufskosten", "Verkaufskosten"), (r"^steuer", "Steuer"), (r"restschuld", "Restschuld"),
             (r"^nettoerl", "Nettoerlös")],
}


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


def _all_series(ch):
    """Reihen aller Teildiagramme (Kombidiagramme: Säule + Linie)."""
    out = list(ch.series)
    for sub in getattr(ch, "_charts", []):
        if sub is not ch:
            out += [s for s in sub.series if not any(s is o for o in out)]
    return out


def chart_kind(ws, ch):
    """Diagrammart aus den Datenbereichen (unabhängig von Nummer und Titel)."""
    rows, sheets, cols = set(), set(), set()
    for s in _all_series(ch):
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
        if rows & {180, 181, D1_ROW} and rows <= {180, 181, D1_ROW, ZB_ROW}:
            return "restschuld"
        if rows & {184} and rows & {185, CF_POS}:
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
    """„Einnahmen vs. Ausgaben“ flach gestapelt in Stapelfolge (P2-02): Einnahmen = Nettokaltmiete (· Steuer),
    Ausgaben = Bewirtschaftung · Zinsen · Tilgung; darüber die „Lücke“ (leerer, rot umrandeter Stapelteil auf der
    Einnahmen-Säule, nur wenn die Ausgaben höher sind) und die Summe über jeder Säule."""
    cats = _r(SHEET, "C", 169, "D", 169)
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 80
    bar.series.append(_series(_r(SHEET, "C", 170, "D", 170), "Nettokaltmiete", cat_ref=cats))
    if with_tax:
        bar.series.append(_series(_r(SHEET, "G", J1_TAX, "H", J1_TAX), title_ref=_r(SHEET, "F", J1_TAX), cat_ref=cats))
    for row, name in ((172, "Bewirtschaftung"), (173, "Zinsen"), (174, "Tilgung")):
        bar.series.append(_series(_r(SHEET, "C", row, "D", row), name, cat_ref=cats))
    g1, g2 = ("G", "H") if with_tax else ("I", "J")
    bar.series.append(_series(_r(SHEET, g1, GAP_ROW, g2, GAP_ROW), "Lücke", cat_ref=cats))
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
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 90
    for c, name in (("I", "Basis"), ("J", "Zufluss"), ("K", "Zufluss (−)"), ("L", "Abfluss"), ("M", "Abfluss (−)"),
                    ("N", "Ergebnis"), ("O", "Ergebnis (−)")):
        bar.series.append(_series(_r(SHEET, c, WF_HEAD + 1, c, WF_HEAD + 4), name, cat_ref=cats))
    line = LineChart()
    for i, c in enumerate(("R", "S", "T", "U")):
        line.series.append(_series(_r(SHEET, c, WF_HEAD + 1, c, WF_HEAD + 4), title_ref=_r(SHEET, "Q", WF_HEAD + 1 + i),
                                   cat_ref=cats))
    for c in ("V", "W", "X"):          # gestrichelte Verbindung Stufe k → k+1 (P2-12)
        line.series.append(_series(_r(SHEET, c, WF_HEAD + 1, c, WF_HEAD + 4), "Verbindung", cat_ref=cats))
    line.y_axis.axId = bar.y_axis.axId
    line.x_axis = bar.x_axis
    bar += line
    return bar


def _cashflow_chart(last):
    """„Cashflow vor und nach Steuern“ als Kombidiagramm (P32): Cashflow n. St. als Säule, v. St. als flache Linie –
    statt 60 schmaler gruppierter Säulen."""
    cats = f"{SHEET}!$D${YEAR_ROW}:${last}${YEAR_ROW}"
    bar = BarChart()
    # P1-07: Säulen positiv 1D4F8A, negativ B42318 – als zwei gestapelte Anzeige-Reihen (Vorzeichen), damit die Farbe in
    # jeder Excel-Version und in LibreOffice gilt (invertIfNegative ohne c14 wäre weiß)
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 80
    for r, name in ((CF_POS, "Überschuss n. St."), (CF_NEG, "Unterdeckung n. St.")):
        bar.series.append(_series(_r(SHEET, "D", r, last, r), name, cat_ref=cats, cat_num=True))
    line = LineChart()
    line.series.append(_series(_r(SHEET, "D", 184, last, 184), "Cashflow v. St.", cat_ref=cats, cat_num=True))
    line.y_axis.axId = bar.y_axis.axId
    line.x_axis = bar.x_axis
    bar += line
    for i, s in enumerate(bar.series + line.series):
        s.idx, s.order = i, i
    return bar


def _pie_kind(ws, ch):
    t = ws.title
    if t.startswith("S03"):
        return "kanc"
    if t.startswith("S06"):
        return "bewirt"
    return chart_kind(ws, ch)


def _pie_legend_cats(ws, ch, kind):
    """Kreis: Kategorien auf Anzeigezellen „Kurzname · 12 %“ umstellen (Legende der Kleinstsegmente)."""
    if kind not in PIE_SHORT or not ch.series:
        return
    s = ch.series[0]
    p, c = _parse(_val_ref(s)), _parse(_cat_ref(s))
    if not p or not c:
        return
    key = (kind, p)
    if key not in _PENDING.setdefault("pies", []):
        _PENDING["pies"].append(key)
        _PENDING.setdefault("pie_labels", {})[key] = [ws.parent[c[0]].cell(r, K.col(c[1])).value
                                                        for r in range(c[2], c[4] + 1)]
    r0 = PIE_HEAD + 1 + sum(k[1][4] - k[1][2] + 2 for k in _PENDING["pies"][:_PENDING["pies"].index(key)])
    r1 = r0 + (p[4] - p[2])
    s.cat = AxDataSource(strRef=StrRef(f=_r(SHEET, PIE_COL, r0, PIE_COL, r1)))


def _afa_line_series(ch):
    """AfA-Vergleich „Kumulierte AfA“: Varianten bleiben einzelne Reihen (Farbe nach Anwendbarkeit setzt finish_pro aus
    den Hilfszeilen F155:I162), dazu die kräftige Linie „Im Modell angewendet“ aus der Summenzeile (P33).
    P1-15: Die Modell-Linie steht ZWEIMAL – zuerst (Legende an erster Stelle: „Im Modell angewendet · Anwendbar · Nicht
    anwendbar“) und zuletzt (liegt über allen Varianten; Legendeneintrag entfällt in finish_pro)."""
    src = [s for s in ch.series if (p := _parse(_val_ref(s))) and p[0] != SHEET]
    if not src:
        return
    first = _parse(_val_ref(src[0]))
    _PENDING["model"] = first
    model = []
    for _ in range(2):
        s_m = _series(_r(SHEET, first[1], MODEL_ROW, first[3], MODEL_ROW), "Im Modell angewendet")
        s_m.cat = src[0].cat
        model.append(s_m)
    ch.series[:] = [model[0]] + src + [model[1]]


def _restschuld_series(ch):
    """Restschuld (S07, Diagramme): Darlehen I und II gestapelt; das Jahr der Zinsbindung als eigene Säule in 4A86C8
    (Beschriftung „Ende Zinsbindung“ setzt finish_pro) – dynamisch über die Anzeige-Hilfszeilen 176/204 (P3-03)."""
    keep, d1 = [], None
    for s in ch.series:
        p = _parse(_val_ref(s))
        if p and p[2] == 180 and d1 is None:
            d1 = s
            continue
        if p and p[2] in (D1_ROW, ZB_ROW):
            continue
        keep.append(s)
    if d1 is None:
        return
    last = _parse(_val_ref(d1))[3]
    cats = f"{SHEET}!$D${YEAR_ROW}:${last}${YEAR_ROW}"
    s1 = _series(_r(SHEET, "D", D1_ROW, last, D1_ROW), "Darlehen I", cat_ref=cats, cat_num=True)
    s2 = _series(_r(SHEET, "D", ZB_ROW, last, ZB_ROW), title_ref=_r(SHEET, "B", ZB_ROW), cat_ref=cats, cat_num=True)
    ch.series[:] = [s1, s2] + keep
    ch.grouping, ch.overlap = "stacked", 100


def chart_contents(wb):
    from openpyxl.chart import PieChart, PieChart3D
    for ws in wb.worksheets:
        if ws.title == "Dashboard":
            continue
        for idx, ch in enumerate(list(ws._charts)):
            ch.visible_cells_only = False          # Daten in ausgeblendeten Zeilen/Spalten weiter anzeigen
            for s in ch.series:                    # Literalnamen (<c:f>"Name"</c:f>) → <c:v>Name</c:v> (Excel-Reparatur)
                lit = _literal(s)
                if lit is not None:
                    s.tx = _label(lit)
            if isinstance(ch, (PieChart, PieChart3D)):
                _pie_legend_cats(ws, ch, _pie_kind(ws, ch))
                continue
            kind = chart_kind(ws, ch)
            if kind == "cashflow":
                new = _cashflow_chart(COCKPIT_LAST if ws.title == "Cockpit" else HORIZON[184])
                new.visible_cells_only = False
                _replace(ws, idx, ch, new)
                continue
            if kind == "cf_nach":
                last = COCKPIT_LAST if ws.title == "Cockpit" else CF_NACH_LAST
                cats = f"{SHEET}!$D${YEAR_ROW}:${last}${YEAR_ROW}"
                ch.series[:] = [_series(_r(SHEET, "D", r, last, r), name, cat_ref=cats, cat_num=True)
                                for r, name in ((CF_POS, "Überschuss nach Steuern"), (CF_NEG, "Unterdeckung nach Steuern"))]
                ch.grouping, ch.overlap, ch.gapWidth = "stacked", 100, 80
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
                for s in ch.series:                 # „Miete -10%“ → „−10 %“, „Miete Basis“ → „Basis“ (einzeilige Legende)
                    if s.tx is not None and s.tx.v:
                        t = re.sub(r"\s*-\s*(\d)", r" −\1", s.tx.v).replace("%", " %").replace("  ", " ")
                        s.tx.v = re.sub(r"^\s*Miete\s*", "", t).strip() or t
            if kind == "restschuld":
                _restschuld_series(ch)
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
                _afa_line_series(ch)
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
        r1, r2 = HH_HEAD + 1, HH_HEAD + min(HH_SLOTS, p[4] - p[2] + 1)
        cats = _r(SHEET, "L", r1, "L", r2)
        ch.series[:] = [_series(_r(SHEET, "M", r1, "M", r2), "Bestehender Haushalt", cat_ref=cats),
                        _series(_r(SHEET, "N", r1, "N", r2), "Neues Objekt", cat_ref=cats)]
    else:
        r1, r2 = VA_HEAD + 1, VA_HEAD + min(VA_SLOTS, p[4] - p[2] + 1)
        cats = _r(SHEET, "U", r1, "U", r2)
        ch.series[:] = [_series(_r(SHEET, "V", r1, "V", r2), "Vermögenswerte", cat_ref=cats)]
    ch.grouping, ch.overlap, ch.gapWidth = "clustered", 100, 60
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
    return re.sub(r"\s*\(.*?\)", "", text or "").strip()[:30]


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
    # „Lücke“: Ausgaben über Einnahmen (leerer, rot umrandeter Stapelteil auf der Einnahmen-Säule, P2-02)
    _put(ws, f"F{GAP_ROW}", "Lücke (n. St. G:H · v. St. I:J)", h="left")
    _put(ws, f"G{GAP_ROW}", f"=MAX(0,H{J1_SUM}-G{J1_SUM})", num)
    _put(ws, f"H{GAP_ROW}", "=0", num)
    _put(ws, f"I{GAP_ROW}", f"=MAX(0,H{J1_SUM_VST}-G{J1_SUM_VST})", num)
    _put(ws, f"J{GAP_ROW}", "=0", num)

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
        # Beschriftungshöhe: negative Stufe → Unterkante (Etikett unter dem Balken, P2-12), sonst Oberkante
        _put(ws, f"P{r}", f"=IF(C{r}<0,{lo},MAX(0,{hi}))", num)
        sign = f'IF(C{r}<0,"−","")' if total else f'IF(C{r}<0,"−","+")'
        _put(ws, f"Q{r}", f'={sign}&FIXED(ABS(C{r})/1000,1)&" T€"')
        for j, c in enumerate("RSTU"):
            _put(ws, f"{c}{r}", f"=P{r}" if j == i else "=0", num)
        for k, c in enumerate("VWX"):       # Verbindung Stufe k → k+1 auf Höhe „Ende“ der Stufe k
            if i in (k, k + 1):             # übrige Zellen bleiben LEER → Lücke im Diagramm (kein #NV)
                _put(ws, f"{c}{r}", f"=H{WF_HEAD + 1 + k}", num)
    for c, t in zip("VWX", ("Verb. 1", "Verb. 2", "Verb. 3")):
        _put(ws, f"{c}{WF_HEAD}", t)
        ws[f"{c}{WF_HEAD}"].font = K.font(K.T_MICRO, True, K.BLUE)

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
            _put(ws, f"F{r}", f'=IF({model},"Im Modell",{name}&IF(N({val})=0," – n. v.",""))', h="left")
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
            cols = [K.L(_col(c0) + k) for k in range(10)]
            lab, val, key_, obj, kk, pos, lab_k, v1, v2, rk = cols
            names = ("Haushalt (Diagramm)", "Wert", "Sortier", "Objekt", "Platz", "Position", "Label", "Bestand",
                     "Objekt neu", "Rang")
            slots = HH_SLOTS
        else:                # Vermögen: eine Reihe
            cols = [K.L(_col(c0) + k) for k in range(8)]
            lab, val, key_, kk, pos, lab_k, v1, rk = cols
            obj = v2 = None
            names = ("Vermögen (Diagramm)", "Wert", "Sortier", "Platz", "Position", "Label", "Wert sortiert", "Rang")
            slots = VA_SLOTS
        heads = dict(zip(cols, names))
        _helper_head(ws, head, cols[0], cols[-1], heads)
        r1, r2 = head + 1, head + n
        rngv, rngk = f"${val}${r1}:${val}${r2}", f"${key_}${r1}:${key_}${r2}"
        rngr = f"${rk}${r1}:${rk}${r2}"
        cnt = f'COUNTIF({rngv},">0")'
        for i in range(n):
            r = r1 + i
            _put(ws, f"{lab}{r}", _short(labels[i], short_table), h="left")
            _put(ws, f"{val}{r}", f"=N({_r(sh, vc, vr1 + i)})", num)
            _put(ws, f"{key_}{r}", f"=IF({val}{r}>0,{val}{r}+({n + 1}-{i + 1})/1000,0)", "0.000")
            if key == "hh":
                _put(ws, f"{obj}{r}", f'=IF(ISNUMBER(SEARCH("kalkulierten",{_r(csh, cc, cr1 + i)})),1,0)', num)
            _put(ws, f"{kk}{r}", i + 1, "0")
            _put(ws, f"{pos}{r}", f'=IF({kk}{r}<=COUNTIF({rngv},">0"),MATCH(LARGE({rngk},{kk}{r}),{rngk},0),"")', "0")
            _put(ws, f"{rk}{r}", f'=IF({key_}{r}>0,COUNTIF({rngk},">"&{key_}{r})+1,99)', "0")
            lab_f = f'IF({pos}{r}="","",INDEX(${lab}${r1}:${lab}${r2},{pos}{r}))'
            last = i + 1 == slots
            if last:        # letzter Diagrammplatz: bei mehr Positionen „Übrige (n)“ mit der Restsumme
                lab_f = f'IF({cnt}>{slots},"Übrige ("&({cnt}-{slots}+1)&")",{lab_f})'
            _put(ws, f"{lab_k}{r}", "=" + lab_f, h="left")
            if key == "hh":
                objk = f"INDEX(${obj}${r1}:${obj}${r2},{pos}{r})"
                valk = f"INDEX({rngv},{pos}{r})"
                f1 = f'IF({pos}{r}="",0,IF({objk}=1,0,{valk}))'
                f2 = f'IF({pos}{r}="",0,IF({objk}=1,{valk},0))'
                if last:
                    rest = f"({rngr}>={slots})*{rngv}"
                    f1 = f"IF({cnt}>{slots},SUMPRODUCT({rest}*(${obj}${r1}:${obj}${r2}=0)),{f1})"
                    f2 = f"IF({cnt}>{slots},SUMPRODUCT({rest}*(${obj}${r1}:${obj}${r2}=1)),{f2})"
                _put(ws, f"{v1}{r}", "=" + f1, num)
                _put(ws, f"{v2}{r}", "=" + f2, num)
            else:
                f1 = f'IF({pos}{r}="",0,INDEX({rngv},{pos}{r}))'
                if last:
                    f1 = f"IF({cnt}>{slots},SUMPRODUCT(({rngr}>={slots})*{rngv}),{f1})"
                _put(ws, f"{v1}{r}", "=" + f1, num)

    ranked("hh", HH_HEAD, "F", HH_SHORT)
    ranked("va", VA_HEAD, "P", VA_SHORT)

    # ---- Kreise: Legendentexte „Kurzname · 12 %“
    pies = _PENDING.get("pies", [])
    if pies:
        _helper_head(ws, PIE_HEAD, PIE_COL, PIE_COL, {PIE_COL: "Kreise: Legende"})
        r = PIE_HEAD + 1
        for key in pies:
            kind, (sh, vc, vr1, _vc2, vr2) = key
            labels = _PENDING["pie_labels"][key]
            tot = _r(sh, vc, vr1, vc, vr2)
            for i in range(vr2 - vr1 + 1):
                short = _short(labels[i] if i < len(labels) else "", PIE_SHORT[kind])
                v = _r(sh, vc, vr1 + i)
                _put(ws, f"{PIE_COL}{r}", f'="{short} · "&TEXT(IFERROR(MAX(0,{v})/SUM({tot}),0),"0 %")', h="left")
                r += 1
            r += 1


def time_helper_rows(ws):
    """Zeilen 203/204/176 der Zeitreihe: Sonder-AfA gesamt; Restschuld Darlehen I im Jahr der Zinsbindung (#NV sonst)
    bzw. ohne dieses Jahr (gestapelte Anzeige im Restschuld-Diagramm)."""
    num = K.NUMFMT["num"]
    ws.cell(SONDER_ROW, 2).value = f'=IF(SUM(D{SONDER_ROW}:W{SONDER_ROW})>0,"Sonder-AfA (§ 7b, § 7h/7i)","Sonder-AfA (keine)")'
    K.set_text(ws.cell(SONDER_ROW, 3), "€")
    for cc in range(_col("D"), _col("AQ") + 1):
        L = K.L(cc)
        ws.cell(SONDER_ROW, cc).value = f"={L}193+{L}194"
        ws.cell(SONDER_ROW, cc).number_format = num
    K.set_text(ws.cell(ZB_ROW, 2), "Ende Zinsbindung")
    K.set_text(ws.cell(ZB_ROW, 3), "€")
    K.set_text(ws.cell(D1_ROW, 2), "Restschuld Darlehen I ohne Zinsbindungsjahr (Diagramm)")
    K.set_text(ws.cell(D1_ROW, 3), "€")
    for cc in range(_col("D"), _col("AQ") + 1):
        L = K.L(cc)
        ws.cell(ZB_ROW, cc).value = f"=IF({L}${INDEX_ROW}=Zinsbindung_I,{L}180,0)"
        ws.cell(D1_ROW, cc).value = f"=IF({L}${INDEX_ROW}=Zinsbindung_I,0,{L}180)"
        ws.cell(ZB_ROW, cc).number_format = num
        ws.cell(D1_ROW, cc).number_format = num


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

    # ---- Abschnittsköpfe Ebene 1 (core.section) B:P; rechts nur „↑ Übersicht“ (Meta nie inline, P2-10)
    _unmerge_rows(ws, [row for row, *_ in SECTIONS] + [DATA_BAND])
    for row, title, extra in SECTIONS:
        _clear_row_style(ws, row, "Q", "U")
        for c in K.iter_cells(ws, "C", row, LAST_COL, row):
            if not K.is_formula(c.value):
                c.value = None
        K.section(ws, row, FIRST_COL, LAST_COL, title, level=1, meta="↑ Übersicht", extra=extra)
        top = ws.cell(row, _col(LAST_COL))
        K.text_link(top, "↑ Übersicht", SHEET, "A4", size=K.T_MICRO, bold=False, tooltip="Zum Seitenanfang")
        top.alignment = K.align("right", "center", 1)

    # ---- Anhang: kein Band mehr, sondern eine dezente Hinweiszeile (P3-05)
    for c in K.iter_cells(ws, "B", DATA_BAND, "U", DATA_BAND):
        if not K.is_formula(c.value):
            c.value = None
        c.fill = K.NOFILL
        c.border = Border(top=K.side("thin", K.LINE2)) if c.column <= _col(LAST_COL) else Border()
        c.hyperlink = None
    hint = ws.cell(DATA_BAND, 2)
    K.set_text(hint, DATA_HINT)
    hint.font = K.font(K.T_MICRO, False, K.MUTED)
    hint.alignment = K.align("left", "center", 1)
    ws.row_dimensions[DATA_BAND].height = K.H_ROW

    # ---- Zeilen der Diagrammzone
    sub_rows = {GRID[k][2] - 1 for k in GRID}
    for r in range(10, DATA_BAND):
        if r in [s[0] for s in SECTIONS]:
            continue
        ws.row_dimensions[r].height = SUBHEAD_H if r in sub_rows else SPACER_ROWS.get(r, 15)

    # ---- Diagramme ins Raster, darüber je ein Unterabschnitt (Titel links, Horizont · Einheit rechts)
    for ch in ws._charts:
        kind = chart_kind(ws, ch)
        if kind in GRID:
            c1, c2, r1 = GRID[kind]
            _anchor(ch, c1, c2, r1)
    for kind, (c1, c2, r1) in GRID.items():
        title, meta = SUBHEADS[kind]
        row = r1 - 1
        for c in K.iter_cells(ws, c1, row, c2, row):
            if not K.is_formula(c.value):
                c.value = None
        K.section(ws, row, c1, c2, title, level=2, variant="line", meta=meta, height=SUBHEAD_H)
        for c in K.iter_cells(ws, c1, row, c2, row):
            c.alignment = K.align(c.alignment.horizontal or "left", "bottom", 1 if c.column in (_col(c1), _col(c2)) else 0)
        ws.cell(row, _col(c2)).alignment = K.align("right", "bottom", 1)
        ws.cell(row, _col(c1)).alignment = K.align("left", "bottom", 1)

    # ---- Fußnoten unter Diagrammen mit leeren Reihen (statt Legendeneintrag „(keines)“, P19)
    for coord, text in ((f"{GRID['restschuld'][0]}{GRID['restschuld'][2] + CHART_ROWS}",
                         '=IF(MAX(D181:AL181)>0,"","Kein zweites Darlehen – dargestellt ist Darlehen I")'),
                        (f"{GRID['afa'][0]}{GRID['afa'][2] + CHART_ROWS}",
                         f'=IF(SUM(D{SONDER_ROW}:W{SONDER_ROW})>0,"","Keine Sonder-AfA (§ 7b, § 7h/7i) im Modell")')):
        c = ws[coord]
        if c.value is None or K.is_formula(c.value):
            c.value = text
            c.font = K.font(K.T_MICRO, False, K.MUTED)
            c.alignment = K.align("left", "center", 1)
            ws.row_dimensions[c.row].height = 12

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
    _body_row(ws, D1_ROW, "AQ", unit=True)
    for r in (MARK_ROW, SONDER_ROW, ZB_ROW, D1_ROW, MODEL_ROW, CF_POS, CF_NEG):   # Hilfszeilen einheitlich: 9 pt kursiv grau
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
    if ws.cell(r, 2).value is None or not K.is_formula(ws.cell(r, 2).value):
        # Reihenname = Beschriftung der Verkaufsmarke in den Bestandsdiagrammen („Verkaufspreis Jahr 12 · 377 T€“)
        ws.cell(r, 2).value = '="Verkaufspreis Jahr "&Haltedauer'

    if ws.cell(r, 3).value is None:
        K.set_text(ws.cell(r, 3), "€")
    for cc in range(_col("D"), _col("AQ") + 1):
        c = ws.cell(r, cc)
        if c.value is None or (K.is_formula(c.value) and "Haltedauer" in c.value):
            L = K.L(cc)
            c.value = f"=IF({L}${INDEX_ROW}=Haltedauer,{L}179,0)"
        c.number_format = K.NUMFMT["num"]


def apply(wb):
    _PENDING.clear()
    chart_contents(wb)
    if SHEET in wb.sheetnames:
        layout_sheet(wb[SHEET])
