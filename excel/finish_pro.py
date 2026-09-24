"""Nachlauf NACH der LibreOffice-Neuberechnung: Styling aller Diagramme und Blatt-XML-Korrekturen.

Aufruf:  python excel/finish_pro.py MAPPE.xlsx   (ändert die Datei an Ort und Stelle)

LibreOffice schreibt beim Roundtrip alle Diagramme neu (Schriften, Farben, Kategorien als multiLvlStrRef, Datenbeschriftungen
je Punkt). Deshalb wird das Diagramm-Styling hier – nach der Neuberechnung – vollständig und regelbasiert gesetzt:

* Diagrammart wird aus den Datenbereichen erkannt (nicht aus Nummer oder Titel), Farben nach Serien- bzw. Kategoriename.
* Titel: auf Schrittseiten und Dashboard entfernt (Panel-Kopf in der Zelle darüber), sonst links, 10 pt fett Navy.
* Achsen: Werteachse ohne Linie, ab 10.000 in „T€“, Hauptgitter 0,5 pt E6E8EB; Jahresachsen waagerecht im 5er-/2er-Takt,
  Beschriftung immer unten (tickLblPos low); Kategorien als einfache Zahlen-/Textreihe.
* Legende unten 8 pt (schmale Diagramme rechts), bei Kreisen und Einzelreihen entfernt; 3D nur für Kreise (ruhige
  Parameter, alle Segmente beschriftet), sonst flach. Reihenfarben mappenweit fest je Kennzahl (SERIES_RULES).
* Wasserfall, Summen über Säulen und dynamische Reihennamen kommen aus Anzeige-Hilfsreihen (layouts/diagramme.py).
Zellwerte werden nicht berührt, die berechneten Ergebnisse bleiben gültig.
"""
import html
import math
import os
import re
import sys
import zipfile
from copy import deepcopy

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import core as K  # noqa: E402

# ============================================================================ Namensräume / Schema-Reihenfolgen
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C14_NS = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
NS = {"c": C_NS, "a": A_NS}


def q(tag):
    ns, t = tag.split(":")
    return f"{{{NS[ns]}}}{t}"


def E(tag, val=None, **attrs):
    el = etree.Element(q(tag))
    if val is not None:
        el.set("val", str(val))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def SE(parent, tag, val=None, **attrs):
    el = E(tag, val, **attrs)
    parent.append(el)
    return el


ORDER = {
    "chartSpace": ["date1904", "lang", "roundedCorners", "AlternateContent", "style", "clrMapOvr", "pivotSource",
                   "protection", "chart", "spPr", "txPr", "externalData", "printSettings", "userShapes", "extLst"],
    "chart": ["title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall", "backWall", "plotArea", "legend",
              "plotVisOnly", "dispBlanksAs", "showDLblsOverMax", "extLst"],
    "title": ["tx", "layout", "overlay", "spPr", "txPr", "extLst"],
    "legend": ["legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst"],
    "catAx": ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt",
              "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "auto",
              "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst"],
    "dateAx": ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt",
               "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "auto",
               "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "extLst"],
    "valAx": ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt",
              "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt",
              "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst"],
    "serAx": ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt",
              "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt",
              "tickLblSkip", "tickMarkSkip", "extLst"],
    "scaling": ["logBase", "orientation", "max", "min", "extLst"],
    "barChart": ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "overlap", "serLines", "axId", "extLst"],
    "bar3DChart": ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "gapDepth", "shape", "axId", "extLst"],
    "lineChart": ["grouping", "varyColors", "ser", "dLbls", "dropLines", "hiLowLines", "upDownBars", "marker", "smooth",
                  "axId", "extLst"],
    "areaChart": ["grouping", "varyColors", "ser", "dLbls", "dropLines", "axId", "extLst"],
    "pieChart": ["varyColors", "ser", "dLbls", "firstSliceAng", "extLst"],
    "pie3DChart": ["varyColors", "ser", "dLbls", "extLst"],
    "bar_ser": ["idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "dPt", "dLbls", "trendline",
                "errBars", "cat", "val", "shape", "extLst"],
    "line_ser": ["idx", "order", "tx", "spPr", "marker", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "smooth",
                 "extLst"],
    "area_ser": ["idx", "order", "tx", "spPr", "pictureOptions", "dPt", "dLbls", "trendline", "errBars", "cat", "val",
                 "extLst"],
    "pie_ser": ["idx", "order", "tx", "spPr", "explosion", "dPt", "dLbls", "cat", "val", "extLst"],
    "view3D": ["rotX", "hPercent", "rotY", "depthPercent", "rAngAx", "perspective", "extLst"],
    "marker": ["symbol", "size", "spPr", "extLst"],
}
SER_ORDER = {"barChart": "bar_ser", "bar3DChart": "bar_ser", "lineChart": "line_ser", "line3DChart": "line_ser",
             "areaChart": "area_ser", "pieChart": "pie_ser", "pie3DChart": "pie_ser", "doughnutChart": "pie_ser"}
CHART_TAGS = ("barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart", "pieChart", "pie3DChart",
              "doughnutChart", "scatterChart", "radarChart")


def local(el):
    return etree.QName(el).localname


def put(parent, el, order):
    """Kind in schemagerechter Reihenfolge einsetzen (gleichnamiges vorhandenes Element wird ersetzt)."""
    name = local(el)
    old = parent.find(q("c:" + name)) if name in order else None
    if old is not None and name not in ("ser", "dPt", "dLbl", "legendEntry", "axId"):
        old.addprevious(el)
        parent.remove(old)
        return el
    pos = order.index(name)
    for i, ch in enumerate(list(parent)):
        n = local(ch)
        if n in order and order.index(n) > pos:
            ch.addprevious(el)
            return el
    parent.append(el)
    return el


def put_val(parent, name, val, order):
    return put(parent, E("c:" + name, val), order)


def drop(parent, *names):
    for n in names:
        for el in parent.findall(q("c:" + n)):
            parent.remove(el)


# ============================================================================ Zeichen-Bausteine
PT = 12700


def solid(color, alpha=None):
    sf = E("a:solidFill")
    clr = SE(sf, "a:srgbClr", color)
    if alpha is not None:
        SE(clr, "a:alpha", int(alpha * 1000))
    return sf


def line(w_pt=None, color=None, dash=None, nofill=False, cap=None):
    ln = E("a:ln")
    if w_pt is not None:
        ln.set("w", str(int(round(w_pt * PT))))
    if cap:
        ln.set("cap", cap)
    if nofill or color is None:
        SE(ln, "a:noFill")
    else:
        ln.append(solid(color))
        if dash:
            SE(ln, "a:prstDash", dash)
        SE(ln, "a:round")
    return ln


def sppr(fill=None, ln=None, nofill=False):
    sp = E("c:spPr")
    if nofill:
        SE(sp, "a:noFill")
    elif fill is not None:
        sp.append(fill)
    if ln is not None:
        sp.append(ln)
    return sp


def rpr(tag, size, color, bold=False):
    r = E(tag)
    r.set("sz", str(int(size * 100)))
    r.set("b", "1" if bold else "0")
    r.append(solid(color))
    SE(r, "a:latin", typeface=K.SANS)
    SE(r, "a:cs", typeface=K.SANS)
    return r


def txpr(size, color, bold=False, rot=None, wrap=None):
    tx = E("c:txPr")
    bp = SE(tx, "a:bodyPr")
    if rot is not None:
        bp.set("rot", str(rot))
        bp.set("vert", "horz")
    if wrap is not None:
        bp.set("wrap", wrap)
    SE(tx, "a:lstStyle")
    p = SE(tx, "a:p")
    ppr = SE(p, "a:pPr")
    ppr.append(rpr("a:defRPr", size, color, bold))
    SE(p, "a:endParaRPr", lang="de-DE")
    return tx


def manual_layout(target=None, x=None, y=None, w=None, h=None):
    lay = E("c:layout")
    ml = SE(lay, "c:manualLayout")
    if target:
        SE(ml, "c:layoutTarget", target)
    if x is not None:
        SE(ml, "c:xMode", "edge")
    if y is not None:
        SE(ml, "c:yMode", "edge")
    for k, v in (("x", x), ("y", y), ("w", w), ("h", h)):
        if v is not None:
            SE(ml, "c:" + k, v)
    return lay


def lbl_offset(x=0.0, y=0.0):
    """Versatz einer Datenbeschriftung relativ zur Standardlage (Anteil der Diagrammgröße, Modus „factor“)."""
    lay = E("c:layout")
    ml = SE(lay, "c:manualLayout")
    SE(ml, "c:x", x)
    SE(ml, "c:y", y)
    return lay


def num_fmt(code):
    """Zahlenformat (nicht quellverknüpft) – Negativ-Sektion immer mit typografischem Minus „−“ (P23)."""
    return E("c:numFmt", formatCode=K.typo_minus(code), sourceLinked="0")


def luminance(hex_):
    r, g, b = (int(hex_[i:i + 2], 16) / 255 for i in (0, 2, 4))
    lin = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in (r, g, b)]
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def text_on(fill):
    """Beschriftungsfarbe auf einer Fläche: Weiß auf dunklen, Navy auf hellen Tönen (Kontrast ≥ 4,5 : 1)."""
    return K.WHITE if luminance(fill or "FFFFFF") < 0.22 else K.NAVY



# ============================================================================ Design-Regeln
FMT_TEUR = '#,##0," T€"'
FMT_TEUR1 = '#,##0.0," T€"'
FMT_EUR = '#,##0" €"'
FMT_PCT = "0 %"
LBL_EUR = '#,##0" €";"−"#,##0" €"'
LBL_EUR_POS = '#,##0" €";;'
LBL_TEUR1 = '#,##0.0," T€";"−"#,##0.0," T€";'
# ---- Mappenweite Farbzuordnung (P2-02, Runde 4): alles in der Blau-Rampe, Status-Rot nur für negative Ergebnisse
INTEREST = K.BLUE      # Zinsen 1D4F8A
TILG = K.ACCENT        # Tilgung 4A86C8
BAR_GREY = K.SKY       # Bewirtschaftung 9CBBE2
TAX_IN = "7FA7D6"      # Steuererstattung (Zufluss) bzw. -zahlung (eine Reihe, dynamischer Name)
TAX_PAY = TAX_IN
KUM_TILG = "8DB3DE"    # kumulierte Tilgung (gestrichelt)
LINE_GREY = "8A94A6"   # Restschuld (gestrichelt)
CF_VST = "8FB3DE"      # Cashflow vor Steuern (Säulen)
NEG = K.RED            # negative Cashflow-Jahre (P1-07: B42318, leicht gedämpft)
NEG_ALPHA = 80
NEG_SOFT = NEG
OUTFLOW = K.MUTED2     # Abfluss in Brücken 8A9099 (P1-07)
PALE = "C9D6E8"        # nicht anwendbar
LEADER = "9AA5B4"      # Führungslinien Kreis
GRID_SENS = "D9DEE7"   # Hilfslinien im Sensitivitäts-Panel
WHITE_ON = {K.NAVY, K.BLUE, K.ACCENT, K.RED}   # weiße Beschriftung nur auf diesen Flächen

# Serienfarben nach Name: (Muster, Farbe, Linienstärke pt, Strichart) – mappenweit fest je Kennzahl (P2-07, P2-09)
SERIES_RULES = [
    (r"^immobilienwert", K.ACCENT, 1.5, None),
    (r"^nettoverm", K.NAVY, 2.5, None),
    (r"^restschuld", LINE_GREY, 1.0, "dash"),
    (r"^darlehen ii\b", K.SKY, None, None),
    (r"^darlehen i\b", K.BLUE, None, None),
    (r"^ende zinsbindung", K.ACCENT, None, None),
    (r"^nettokaltmiete", K.NAVY, None, None),
    (r"^steuererstattung", TAX_IN, None, None),
    (r"^steuerzahlung", TAX_PAY, None, None),
    (r"^kumulierte zinsen", K.BLUE, 2.0, None),
    (r"^zinsen", INTEREST, None, None),
    (r"^kumulierte tilgung", KUM_TILG, 2.0, "dash"),
    (r"^tilgung", TILG, None, None),
    (r"^bewirtschaftung", BAR_GREY, None, None),
    (r"^cashflow vor", K.NAVY, 2.0, None),           # Linie im Kombidiagramm (P1-07: 0B2A4A, 2 pt)
    (r"^cashflow nach", K.BLUE, None, None),
    (r"^kumulierter cashflow", K.NAVY, 2.25, None),
    (r"^kumulierte steuer", K.ACCENT, 1.5, None),
    (r"^afa regul", K.BLUE, None, None),
    (r"sonder-afa", K.ACCENT, None, None),
    (r"7h", K.ACCENT, None, None),
    (r"^bewegliche", K.SKY, None, None),
    (r"^steuerliches ergebnis", K.BLUE, None, None),
    (r"^steuer\b", K.SKY, None, None),                 # P1-15: Steuer-Reihe hell, klar vom Ergebnis getrennt
    (r"^(ü|ue)berschuss", K.BLUE, None, None),
    (r"^unterdeckung", NEG, None, None),
    # Wasserfall / AfA-Summe / Haushalt (Hilfsreihen aus layouts/diagramme.py)
    (r"^basis", None, None, None),
    (r"^zufluss", K.ACCENT, None, None),
    (r"^abfluss", OUTFLOW, None, None),
    (r"^ergebnis", K.NAVY, None, None),
    (r"^anwendbar", K.ACCENT, 1.5, None),
    (r"^nicht anwendbar", PALE, 1.0, None),
    (r"^bestehender haushalt", K.BLUE, None, None),
    (r"^neues objekt", K.ACCENT, None, None),
    (r"^verm(ö|oe)genswerte", K.BLUE, None, None),
    # AfA-Varianten (AfA-Vergleich, kumuliert): Farbe nach Anwendbarkeit, keine Strichvarianten (P33)
    (r"^im modell", K.NAVY, 2.5, None),
]
# Mietszenarien (Sensitivität, P1-15): geordnete Blau-Rampe in Richtung der Mietänderung, Basis kräftig
IRR_RULES = [(r"[-−]\s*10", "C9D6E6", 1.0), (r"[-−]\s*5", K.SKY, 1.0), (r"basis", K.NAVY, 2.25),
             (r"\+\s*5", K.ACCENT, 1.0), (r"\+\s*10", K.BLUE, 1.0)]
FALLBACK = [K.NAVY, K.ACCENT, K.SKY, K.BLUE, "6E7F96", K.MIST]

# Kreise (P1-01): Farbfolge fest nach Wert absteigend – größtes Segment 0B2A4A, dann 1D4F8A, 4A86C8, 9CBBE2, C9D6E6
PIE_RANK = ["0B2A4A", "1D4F8A", "4A86C8", "9CBBE2", "C9D6E6"]
PIE_RANK7 = ["0B2A4A", "1D4F8A", "2F6BAE", "4A86C8", "7FA7D6", "9CBBE2", "C9D6E6"]
PIE_INSIDE = 0.35        # Innen-Etikett nur ab 35 % (höchstens zwei je Kreis), sonst außen
PIE_MAX_INSIDE = 2
PIE_ALL_OUT = 4          # mehr als 4 Segmente: alle Etiketten außen mit Führungslinie
# Kurzlabels der Kategorien: Art → [(Muster, Kurzname)]
CAT_SHORT = {
    "invest": [(r"^kaufpreis", "Kaufpreis"), (r"^kaufneben", "Nebenkosten"), (r"^finanzierungsneben", "Finanzierungskosten"),
               (r"^ma(ß|ss)nahmen", "Maßnahmen")],
    "finanz": [(r"^darlehen ii\b", "Darlehen II"), (r"^darlehen i\b", "Darlehen I"), (r"^eigenkapital", "Eigenkapital")],
    "kpa": [(r"^geb", "Gebäude"), (r"^grund", "Boden"), (r"^beweg", "Inventar"), (r"r(ü|ue)cklage", "Rücklage")],
    "kanc": [(r"^grunderwerb", "Grunderwerbsteuer"), (r"^notar", "Notar"), (r"^grundbuch", "Grundbuch"),
             (r"^makler", "Makler"), (r"^sonst", "Sonstige")],
    "bewirt": [(r"^hausgeld", "Hausgeld"), (r"erhaltungsr", "Rücklage WEG"), (r"^verwaltung", "Verwaltung"),
               (r"instandh", "Instandh."), (r"^mietausfall", "Mietausfall"), (r"werbungsk", "Werbungsk."),
               (r"^r(ü|ue)cklage weg", "Rücklage WEG")],
    "exit": [(r"^verkaufskosten", "Verkaufskosten"), (r"^steuer", "Steuer"), (r"restschuld", "Restschuld"),
             (r"^nettoerl", "Nettoerlös")],
}
PIE_KINDS = {"invest", "finanz", "kpa", "kanc", "bewirt", "exit"}
DIRECT_LABELS = {"jahr1", "afa_summe", "hh", "va"}
HBAR_KINDS = {"afa_summe", "hh", "va"}
TIME_KINDS = {"bestand", "restschuld", "cashflow", "cf_nach", "cf_kum", "kumzins", "afa", "steuer"}

# Einheitliche Titel „Kennzahl (Zeitraum, Einheit)“ – {n} = Anzahl Jahre im Diagramm (Horizont steht im Titel)
TITLES = {
    "invest": "Gesamtinvestition",
    "finanz": "Finanzierungsstruktur",
    "kpa": "Kaufpreisaufteilung",
    "exit": "Verwendung des Verkaufserlöses",
    "ertrag": "Gesamtertrag nach Steuern (Brücke, T€)",
    "jahr1": "Einnahmen vs. Ausgaben (Jahr 1, € / Monat)",
    "bestand": "Vermögen und Restschuld (Jahre 1–{n}, {u})",
    "restschuld": "Restschuld am Jahresende bis zur Volltilgung (Jahre 1–{n}, {u})",
    "cashflow": "Cashflow vor und nach Steuern (Jahre 1–{n}, {u} p. a.)",
    "cf_nach": "Cashflow nach Steuern (Jahre 1–{n}, {u} p. a.)",
    "cf_kum": "Kumulierter Cashflow nach Steuern (Jahre 1–{n}, {u})",
    "kumzins": "Kumulierte Zinsen, Tilgung und Steuer (Jahre 1–{n}, {u})",
    "afa": "Abschreibungen nach Komponenten (Jahre 1–{n}, {u} p. a.)",
    "steuer": "Steuerliches Ergebnis und Steuer (Jahre 1–{n}, {u} p. a.)",
    "kanc": "Zusammensetzung der Kaufnebenkosten",
    "bewirt": "Bewirtschaftungskosten (Jahr 1, € p. a.)",
    "afa_summe": "Summe AfA je Variante (Jahre 1–10, T€)",
    "afa_kum": "Kumulierte AfA je Variante (Jahre 1–10, {u})",
    "irr": "IRR nach Steuern je Wertsteigerung (Linien: Mietänderung)",
    "hh": "Ausgaben je Monat (nur Positionen > 0)",
    "va": "Vermögenswerte (nur Positionen > 0)",
}
MARK_ROW = 202          # Diagramme: Immobilienwert im Verkaufsjahr (sonst 0), angelegt von layouts/diagramme.py
AFA_SUM_ROW = 155       # Diagramme F155:I162: AfA je Variante nach Bedeutung (G Modell · H anwendbar · I nicht anwendbar)
AFA_CLASS = {}          # Variante k → "Anwendbar" / "Nicht anwendbar" (aus den berechneten Hilfszeilen)
MARK_LAST = "AL"


# ============================================================================ Paket lesen: Diagramm → Blatt
def _rels(files, path):
    d, b = os.path.split(path)
    rp = f"{d}/_rels/{b}.rels"
    out = {}
    if rp in files:
        for m in re.finditer(rb"<Relationship\b[^>]*>", files[rp]):
            tag = m.group(0).decode()
            rid = re.search(r'Id="([^"]+)"', tag).group(1)
            tgt = re.search(r'Target="([^"]+)"', tag).group(1)
            out[rid] = os.path.normpath(os.path.join(d, tgt)).replace("\\", "/") if not tgt.startswith("/") else tgt[1:]
    return out


def chart_sheets(files):
    """{'xl/charts/chartN.xml': Blattname} sowie {Blattname: Blatt-XML-Pfad}."""
    wb = files["xl/workbook.xml"].decode("utf-8")
    wrels = _rels(files, "xl/workbook.xml")
    charts, sheets = {}, {}
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        tag = m.group(0)
        name = html.unescape(re.search(r'name="([^"]*)"', tag).group(1))
        rid = re.search(r'r:id="([^"]+)"', tag).group(1)
        sp = wrels.get(rid)
        if not sp:
            continue
        sheets[name] = sp
        for tgt in _rels(files, sp).values():
            if "/drawings/drawing" in tgt and tgt.endswith(".xml"):
                for ct in _rels(files, tgt).values():
                    if "/charts/chart" in ct:
                        charts[ct] = name
    return charts, sheets


def chart_sizes(files):
    """{'xl/charts/chartN.xml': (Breite, Höhe) in px} aus den Zeichnungsebenen (xfrm/ext)."""
    out = {}
    for name in files:
        if not (name.startswith("xl/drawings/drawing") and name.endswith(".xml")):
            continue
        rels = _rels(files, name)
        root = etree.fromstring(files[name])
        for gf in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}graphicFrame"):
            ch = gf.find(".//" + q("c:chart"))
            ext = gf.find(".//" + q("a:ext"))
            if ch is None or ext is None:
                continue
            tgt = rels.get(ch.get(f"{{{R_NS}}}id"))
            if tgt:
                out[tgt] = (int(ext.get("cx", "0")) / 9525, int(ext.get("cy", "0")) / 9525)
    return out


def chart_widths(files):
    return {k: v[0] for k, v in chart_sizes(files).items()}


def row_values(files, sheet_path, row):
    """Berechnete Werte einer Blattzeile {Spalte: Wert} (Zahlen)."""
    out = {}
    if sheet_path not in files:
        return out
    root = etree.fromstring(files[sheet_path])
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    for r in root.iterfind(".//m:sheetData/m:row", ns):
        if r.get("r") != str(row):
            continue
        for c in r.iterfind("m:c", ns):
            v = c.find("m:v", ns)
            ref = c.get("r")
            colm = re.match(r"([A-Z]+)", ref).group(1)
            try:
                out[colm] = float(v.text) if v is not None and c.get("t") in (None, "n") else None
            except (TypeError, ValueError):
                out[colm] = None
    return out



# ============================================================================ Diagramm-Analyse
def ser_name(ser):
    tx = ser.find(q("c:tx"))
    if tx is None:
        return ""
    v = tx.xpath(".//c:v/text()", namespaces=NS)
    if v:
        return v[0]
    f = tx.xpath(".//c:f/text()", namespaces=NS)
    return f[0].strip('"') if f else ""


def ser_ref(ser, part="val"):
    f = ser.xpath(f"./c:{part}//c:f/text()", namespaces=NS)
    return f[0] if f else ""


def cache_values(ser, part="val"):
    vals = []
    for pt in ser.xpath(f"./c:{part}//c:pt", namespaces=NS):
        v = pt.find(q("c:v"))
        try:
            vals.append(float(v.text))
        except (TypeError, ValueError, AttributeError):
            pass
    return vals


def categories(ser):
    """Kategorien als Liste (Texte) – auch aus LibreOffice-multiLvlStrCache."""
    cat = ser.find(q("c:cat"))
    if cat is None:
        return []
    ml = cat.find(q("c:multiLvlStrRef"))
    if ml is not None:
        lv = ml.findall(".//" + q("c:lvl"))
        vals = [("".join(l.xpath(".//c:v/text()", namespaces=NS))) for l in lv]
        cnt = ml.find(".//" + q("c:ptCount"))
        n = int(cnt.get("val")) if cnt is not None else 0
        if n <= 1 and len(vals) > 1:       # Zeilenbereich als „Ebenen“ gelesen → Reihenfolge umkehren
            return list(reversed(vals))
        return vals
    pts = cat.xpath(".//c:pt", namespaces=NS)
    res = [None] * len(pts)
    for i, pt in enumerate(pts):
        idx = int(pt.get("idx", i))
        if idx >= len(res):
            res.extend([None] * (idx + 1 - len(res)))
        res[idx] = "".join(pt.xpath("./c:v/text()", namespaces=NS))
    return res


def ref_rows(ref):
    m = re.match(r"^(?:'?(.+?)'?!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", ref or "")
    if not m:
        return None
    return (m.group(1) or "").replace("''", "'"), m.group(2), int(m.group(3)), m.group(4), int(m.group(5) or m.group(3))



def chart_kind(sheet, root):
    if sheet == "Dashboard":
        return "dash"
    # blattbezogene Diagramme zuerst (ihre Hilfsreihen liegen teils auf „Diagramme“)
    if sheet.startswith("S03"):
        return "kanc"
    if sheet.startswith("S06"):
        return "bewirt"
    if sheet == "AfA-Vergleich":
        return "afa_kum" if root.find(".//" + q("c:lineChart")) is not None else "afa_summe"
    if sheet == "Sensitivität":
        return "irr"
    if sheet == "Haushaltsrechnung":
        return "hh"
    if sheet == "Vermögensaufstellung":
        return "va"
    sers = root.xpath(".//c:plotArea/*/c:ser", namespaces=NS)
    rows, shs, cols = set(), set(), set()
    for s in sers:
        p = ref_rows(ser_ref(s))
        if p:
            shs.add(p[0])
            rows.add(p[2])
            cols.add(p[1])
    if shs == {"Diagramme"}:
        if rows & {179, 182, 183}:
            return "bestand"
        if rows & {176, 180, 181} and rows <= {176, 180, 181, 204}:
            return "restschuld"
        if 184 in rows and (185 in rows or 206 in rows):
            return "cashflow"
        if rows == {185} or (rows and rows <= {206, 207}):
            return "cf_nach"
        if rows == {186}:
            return "cf_kum"
        if rows & {189, 190, 191}:
            return "kumzins"
        if rows & {192, 193, 194, 195, 203}:
            return "afa"
        if rows & {196, 197}:
            return "steuer"
        if rows & set(range(170, 177)):
            return "jahr1"
        if cols <= {"C"}:
            for r, k in {141: "invest", 147: "finanz", 152: "kpa", 158: "exit", 164: "ertrag"}.items():
                if r in rows:
                    return k
        if rows & set(range(164, 168)):
            return "ertrag"
    return None


def match_series(name):
    n = (name or "").strip().lower()
    for pat, color, w, dash in SERIES_RULES:
        if re.search(pat, n):
            return color, w, dash
    return None


def short_cat(kind, name):
    name = (name or "").split(" · ")[0]          # Legendentext „Kurzname · 12 %“ (layouts/diagramme.py)
    for pat, short in CAT_SHORT.get(kind, []):
        if re.search(pat, name.strip().lower()):
            return short
    return re.sub(r"\s*\(.*?\)", "", name)[:22]


def pie_colors(kind, cats, vals):
    """Farbe je Segment (P1-01): fest nach Wert absteigend. Zwei Segmente: 0B2A4A + 4A86C8 (klarer Kontrast)."""
    n = len(cats)
    order = sorted(range(n), key=lambda i: -(vals[i] if i < len(vals) else 0))
    m = sum(1 for v in vals if v > 0) or 1
    ramp = PIE_RANK7 if m > len(PIE_RANK) else ([PIE_RANK[0], PIE_RANK[2]] if m == 2 else PIE_RANK)
    colors = [None] * n
    for rank, i in enumerate(order):
        colors[i] = ramp[min(rank, len(ramp) - 1)]
    return colors


def fix_literal_names(root):
    """Reihennamen als Literal-„Formel“ (<c:f>"Name"</c:f>, LibreOffice-Rest) → <c:v>Name</c:v> (sonst Excel-Reparatur)."""
    for tx in root.iter(q("c:tx")):
        if tx.getparent() is None or local(tx.getparent()) != "ser":
            continue
        sr = tx.find(q("c:strRef"))
        if sr is None:
            continue
        f = sr.find(q("c:f"))
        text = (f.text or "") if f is not None else ""
        if text.startswith('"') and text.endswith('"'):
            tx.remove(sr)
            SE(tx, "c:v").text = text[1:-1].replace('""', '"')


# ============================================================================ Kategorien vereinfachen
def simplify_categories(root):
    """multiLvlStrRef (LibreOffice-Rest bei einzeiligen Jahresbereichen) → numRef bzw. strRef mit flachem Cache."""
    for cat in root.iter(q("c:cat")):
        ml = cat.find(q("c:multiLvlStrRef"))
        if ml is None:
            continue
        f = ml.find(q("c:f")).text
        ser = cat.getparent()
        vals = categories(ser)
        numeric = all(re.fullmatch(r"-?\d+(\.\d+)?", v or "") for v in vals) and vals
        cat.remove(ml)
        if numeric:
            ref = SE(cat, "c:numRef")
            SE(ref, "c:f").text = f
            cache = SE(ref, "c:numCache")
            SE(cache, "c:formatCode").text = "General"
        else:
            ref = SE(cat, "c:strRef")
            SE(ref, "c:f").text = f
            cache = SE(ref, "c:strCache")
        SE(cache, "c:ptCount", len(vals))
        for i, v in enumerate(vals):
            pt = SE(cache, "c:pt", idx=i)
            SE(pt, "c:v").text = v


# ============================================================================ 3D (bleibt – ruhig parametriert)
def _walls(tag):
    el = E(tag)
    SE(el, "c:thickness", 0)
    el.append(sppr(nofill=True, ln=line(nofill=True)))
    return el


MAX_3D_CATS = 8          # Nutzerentscheidung: 3D nur für Kreise und einfache Säulen (≤ 8 Kategorien)
FLAT_KINDS = {"ertrag", "jahr1", "dash_wf"}   # Wasserfall/Brücken/Einnahmen-Ausgaben immer flach
KEEP_GAP = {"ertrag", "jahr1", "dash_wf", "afa_summe", "hh", "va", "cashflow"}   # Lücke/Überlappung aus dem Layout-Modul


def make_flat(root, n_cat, keep_gap=False):
    """3D-Säulen zurück in flache 2D-Säulen (Zeitreihen, Wasserfall, Balken): ruhige Standardgeometrie."""
    chart = root.find(q("c:chart"))
    plot = chart.find(q("c:plotArea"))
    for bc in plot.findall(q("c:bar3DChart")):
        bc.tag = q("c:barChart")
        drop(bc, "gapDepth", "shape")
    for tag in ("view3D", "floor", "sideWall", "backWall"):
        drop(chart, tag)
    for bc in ([] if keep_gap else plot.findall(q("c:barChart"))):
        grp = bc.find(q("c:grouping"))
        stacked = grp is not None and grp.get("val") in ("stacked", "percentStacked")
        n_ser = len(bc.findall(q("c:ser")))
        put_val(bc, "gapWidth", (60 if n_cat >= 20 else 80) if n_ser == 1 or stacked else (50 if n_cat >= 20 else 80), ORDER["barChart"])
        put_val(bc, "overlap", 100 if stacked else (0 if n_ser > 1 else 0), ORDER["barChart"])
    for sa in plot.findall(q("c:serAx")):
        ax_id = sa.find(q("c:axId")).get("val")
        plot.remove(sa)
        for ct in plot:
            for a in ct.findall(q("c:axId")):
                if a.get("val") == ax_id:
                    ct.remove(a)
    return None



def make_3d(root, n_cat, kind_name=None):
    """Nutzerentscheidung „Mischung“: 3D nur für Kreise und einfache Säulen (senkrecht, ≤ 8 Kategorien, kein Wasserfall);
    Zeitreihen, Wasserfall, waagerechte Balken und Liniendiagramme flach (2D). Kombidiagramme bleiben 2D."""
    chart = root.find(q("c:chart"))
    plot = chart.find(q("c:plotArea"))
    types = [local(e) for e in plot if local(e) in CHART_TAGS]
    if len(types) != 1:
        return None
    bar_dir = plot.find(".//" + q("c:barDir"))
    horizontal = bar_dir is not None and bar_dir.get("val") == "bar"
    if types[0] in ("barChart", "bar3DChart") and (n_cat > MAX_3D_CATS or horizontal or kind_name in FLAT_KINDS):
        return make_flat(root, n_cat, keep_gap=(kind_name in KEEP_GAP))
    kind = None
    for bc in plot.findall(q("c:barChart")):
        bc.tag = q("c:bar3DChart")
        drop(bc, "overlap", "serLines")
        kind = "bar"
    for pc in plot.findall(q("c:pieChart")):
        pc.tag = q("c:pie3DChart")
        drop(pc, "firstSliceAng")
        kind = "pie"
    if plot.find(q("c:bar3DChart")) is not None:
        kind = "bar"
    if plot.find(q("c:pie3DChart")) is not None:
        kind = "pie"
    if kind is None:
        return None
    for tag in ("view3D", "floor", "sideWall", "backWall"):
        drop(chart, tag)
    v = E("c:view3D")
    if kind == "pie":
        SE(v, "c:rotX", 40)
        SE(v, "c:rotY", pie_rotation(root))
        SE(v, "c:rAngAx", 0)
        SE(v, "c:perspective", 10)
    else:
        SE(v, "c:rotX", 10)
        SE(v, "c:rotY", 15)
        SE(v, "c:depthPercent", 40 if n_cat >= 20 else 50)
        SE(v, "c:rAngAx", 1)
    put(chart, v, ORDER["chart"])
    if kind == "bar":
        for tag in ("c:floor", "c:sideWall", "c:backWall"):
            put(chart, _walls(tag), ORDER["chart"])
        bc = plot.find(q("c:bar3DChart"))
        for pos in list(bc.iter(q("c:dLblPos"))):
            pos.getparent().remove(pos)
        put_val(bc, "gapWidth", 100 if n_cat >= 10 else 80, ORDER["bar3DChart"])
        put_val(bc, "gapDepth", 150, ORDER["bar3DChart"])
        put_val(bc, "shape", "box", ORDER["bar3DChart"])
    return kind


def pie_rotation(root, target=270):
    """Drehung so, dass das größte Segment links (9 Uhr) liegt und die kleinen Segmente rechts (3 Uhr) auslaufen –
    dort stapeln sich ihre Außen-Etiketten senkrecht statt nebeneinander (P1-01)."""
    ser = root.find(".//" + q("c:ser"))
    vals = [max(0.0, v) for v in cache_values(ser)] if ser is not None else []
    tot = sum(vals)
    if tot <= 0:
        return 0
    big = max(range(len(vals)), key=lambda i: vals[i])
    start = sum(vals[:big]) / tot * 360
    mid = start + vals[big] / tot * 180
    return int(round((target - mid) % 360))


# ============================================================================ Titel, Legende, Achsen
TITLE_PT = 8.5


def title_caps(text):
    """Diagrammtitel wie ein Unterabschnitt (P2-10): VERSALIEN, Zusatz „(Zeitraum, Einheit)“ als „· Zeitraum · Einheit“."""
    t = re.sub(r"\s*\((.*)\)\s*$", lambda m: " · " + m.group(1).replace(", ", " · "), text or "")
    t = t.upper()
    for a, b in (("AFA", "AfA"), ("P. A.", "p. a."), ("N. ST.", "n. St."), ("V. ST.", "v. St.")):
        t = t.replace(a, b)
    return t


def set_title(root, text, delete, width_px=None):
    chart = root.find(q("c:chart"))
    old = chart.find(q("c:title"))
    if delete:
        if old is not None:
            chart.remove(old)
        put_val(chart, "autoTitleDeleted", 1, ORDER["chart"])
        return
    if text is None and old is not None:
        text = "".join(old.xpath(".//a:t/text()", namespaces=NS))
    if not text:
        return
    text = title_caps(text)
    t = E("c:title")
    tx = SE(t, "c:tx")
    rich = SE(tx, "c:rich")
    SE(rich, "a:bodyPr", rot="0", vert="horz", wrap="square")
    SE(rich, "a:lstStyle")
    p = SE(rich, "a:p")
    ppr = SE(p, "a:pPr", algn="l")
    ppr.append(rpr("a:defRPr", TITLE_PT, K.BLUE, True))
    r = SE(p, "a:r")
    r.append(rpr("a:rPr", TITLE_PT, K.BLUE, True))
    r[-1].set("lang", "de-DE")
    SE(r, "a:t").text = text
    # linke Kante des Titels auf der Textkante der Bänder (Einzug 1 ≈ 12 px)
    x = round(min(0.05, max(0.004, 12 / width_px)), 4) if width_px else 0.012
    t.append(manual_layout(x=x, y=0.025))
    SE(t, "c:overlay", 0)
    t.append(sppr(nofill=True, ln=line(nofill=True)))
    put(chart, t, ORDER["chart"])
    put_val(chart, "autoTitleDeleted", 0, ORDER["chart"])


def set_legend(root, show, hidden_idx=(), pos="b"):
    chart = root.find(q("c:chart"))
    drop(chart, "legend")
    if not show:
        return
    lg = E("c:legend")
    SE(lg, "c:legendPos", pos)
    for i in sorted(set(hidden_idx)):
        le = SE(lg, "c:legendEntry")
        SE(le, "c:idx", i)
        SE(le, "c:delete", 1)
    SE(lg, "c:overlay", 0)
    lg.append(sppr(nofill=True, ln=line(nofill=True)))
    lg.append(txpr(8, K.MUTED))
    put(chart, lg, ORDER["chart"])


def nice_unit(span, max_ticks):
    """Hauptintervall 1/2/2,5/5 × 10^k, sodass höchstens max_ticks Intervalle entstehen."""
    if span <= 0 or max_ticks < 1:
        return None
    raw = span / max_ticks
    mag = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 5, 10):
        if raw <= m * mag + 1e-12:
            return m * mag
    return 10 * mag


def value_range(plot):
    """Wertebereich der Darstellung (gestapelte Säulen: Summen je Kategorie, positiv und negativ getrennt)."""
    lo = hi = 0.0
    for ct in plot:
        if local(ct) not in CHART_TAGS:
            continue
        grp = ct.find(q("c:grouping"))
        stacked = grp is not None and grp.get("val") in ("stacked", "percentStacked")
        sers = ct.findall(q("c:ser"))
        if stacked:
            pos, neg = {}, {}
            for s in sers:
                for pt in s.xpath("./c:val//c:pt", namespaces=NS):
                    try:
                        v = float(pt.find(q("c:v")).text)
                    except (TypeError, ValueError, AttributeError):
                        continue
                    i = int(pt.get("idx"))
                    (pos if v >= 0 else neg)[i] = (pos if v >= 0 else neg).get(i, 0) + v
            lo = min([lo] + list(neg.values()))
            hi = max([hi] + list(pos.values()))
        else:
            for s in sers:
                vals = cache_values(s)
                if vals:
                    lo, hi = min(lo, min(vals)), max(hi, max(vals))
    return lo, hi


def value_format(root, kind, n_cat, size=None):
    """Achsenformat und Einheit (P09): Jahresreihen immer in T€ (bei kleinen Intervallen mit einer Nachkommastelle),
    sonst ab 10.000 in T€; die Einheit im Titel folgt dem Achsenformat. Rückgabe (Format, Einheit, Intervall)."""
    plot = root.find(".//" + q("c:plotArea"))
    lo, hi = value_range(plot)
    big = max(abs(lo), abs(hi))
    h_px = size[1] if size else 300
    max_ticks = max(3, min(8, int(h_px * 0.72 / 26)))       # engere Teilung → kein leeres Band über den Daten
    unit = nice_unit(hi - lo, max_ticks) if kind not in DIRECT_LABELS else None
    if kind == "irr":
        return FMT_PCT, "%", unit
    if kind in TIME_KINDS or kind in ("cf_nach", "afa_kum", "afa_summe") or big >= 10000:
        return (FMT_TEUR1 if unit and unit < 1000 else FMT_TEUR), "T€", unit
    return FMT_EUR, "€", unit


def style_axes(root, kind, n_cat, size=None, horizontal=False):
    plot = root.find(".//" + q("c:plotArea"))
    lo, hi = value_range(plot)
    dash = kind in ("dash", "dash_line")
    vfmt, _u, unit = value_format(root, kind, n_cat, size)
    time_series = n_cat >= 10 and kind not in HBAR_KINDS
    for ax in plot.findall(q("c:valAx")):
        o = ORDER["valAx"]
        sc = ax.find(q("c:scaling"))
        if sc is not None and lo < 0:
            drop(sc, "min")
        elif sc is not None and kind not in HBAR_KINDS:
            # Wertachse ab 0 (P08) – keine abgeschnittene Achse, die den Verlauf übertreibt
            put(sc, E("c:min", 0), ORDER["scaling"])
        if kind in DIRECT_LABELS:
            # direkt beschriftete Diagramme: keine Werteachse, kein Gitter (Werte stehen an den Säulen)
            put_val(ax, "delete", 1, o)
            drop(ax, "majorGridlines", "minorGridlines")
        else:
            put_val(ax, "delete", 0, o)
            gl = E("c:majorGridlines")
            gl.append(sppr(ln=line(0.5, GRID_SENS if kind == "irr" else K.LINE)))
            put(ax, gl, o)
            drop(ax, "minorGridlines")
        put(ax, num_fmt(vfmt), o)
        put_val(ax, "majorTickMark", "none", o)
        put_val(ax, "minorTickMark", "none", o)
        put_val(ax, "tickLblPos", "nextTo", o)
        put(ax, sppr(nofill=True, ln=line(nofill=True)), o)
        put(ax, txpr(8, K.MUTED), o)
        if dash:
            pass                                   # Dashboard (Agent G): Intervall bleibt (Wasserfall majorUnit 500)
        elif unit:
            put_val(ax, "majorUnit", ("%g" % unit), o)
        else:
            drop(ax, "majorUnit")
        if kind in HBAR_KINDS:
            put_val(ax, "crosses", "max", o)
            drop(ax, "crossesAt")
    for tag in ("catAx", "dateAx"):
        for ax in plot.findall(q("c:" + tag)):
            o = ORDER[tag]
            put_val(ax, "delete", 0, o)
            drop(ax, "majorGridlines", "minorGridlines")
            if kind == "irr":
                put(ax, num_fmt("0 %"), o)
            elif time_series:
                put(ax, num_fmt("0"), o)
            put_val(ax, "majorTickMark", "none", o)          # keine Teilstriche auf Kategorieachsen (P3-04)
            put_val(ax, "minorTickMark", "none", o)
            put_val(ax, "tickLblPos", "low", o)
            # Kategorieachse = Nulllinie: bei negativen Werten 1 pt 8A9099, sonst 0,75 pt
            if kind in HBAR_KINDS:
                put(ax, sppr(ln=line(nofill=True)), o)
            elif horizontal:
                put(ax, sppr(ln=line(0.75, "C5CDD8")), o)
            else:
                # Nulllinie 0,75 pt 8A9099 (P1-07), bei negativen Werten 1 pt
                put(ax, sppr(ln=line(1.0 if lo < 0 else 0.75, K.MUTED2)), o)
            if kind in HBAR_KINDS:
                put(ax, txpr(9, K.INK2), o)
            else:
                # rot=60 (= 0,001°, in Excel waagerecht): verhindert in LibreOffice den Zeilenumbruch der Jahreszahlen
                put(ax, txpr(8, K.MUTED, rot=60 if (time_series and not horizontal and not dash) else None,
                             wrap="none" if time_series else None), o)
            if kind in HBAR_KINDS:
                sc = ax.find(q("c:scaling"))
                o_el = sc.find(q("c:orientation"))
                if o_el is None:
                    o_el = E("c:orientation")
                    put(sc, o_el, ORDER["scaling"])
                o_el.set("val", "maxMin")
            if tag == "catAx":
                if time_series:
                    skip = 5 if n_cat >= 30 else 2 if n_cat >= 15 else 1
                    put_val(ax, "tickLblSkip", skip, o)
                    put_val(ax, "tickMarkSkip", skip, o)
                else:
                    drop(ax, "tickLblSkip", "tickMarkSkip")
                put_val(ax, "noMultiLvlLbl", 1, o)
    for ax in plot.findall(q("c:serAx")):
        put_val(ax, "delete", 1, ORDER["serAx"])


# ============================================================================ Datenbeschriftungen
SHOW = ("showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize")


def _flags(parent, **on):
    for n in SHOW:
        SE(parent, "c:" + n, 1 if on.get(n) else 0)


def dlbls_off():
    d = E("c:dLbls")
    _flags(d)
    return d


def dlbls_val(fmt, size, color, bold=False, pos=None, bg=None, ser_name=False):
    d = E("c:dLbls")
    d.append(num_fmt(fmt))
    d.append(sppr(solid(bg), line(nofill=True)) if bg else sppr(nofill=True, ln=line(nofill=True)))
    d.append(txpr(size, color, bold, wrap="none"))
    if pos:
        SE(d, "c:dLblPos", pos)
    _flags(d, showVal=not ser_name, showSerName=ser_name)
    return d


def dlbl_delete(idx):
    lb = E("c:dLbl")
    SE(lb, "c:idx", idx)
    SE(lb, "c:delete", 1)
    return lb


def set_ser_dlbls(ser, dl, order):
    drop(ser, "dLbls")
    if dl is not None:
        put(ser, dl, order)


PIE_SMALL = []           # (Altbestand – Teillegenden entfallen seit Runde 4)
PIE_OUT = []             # Außen-Etiketten des zuletzt gestalteten Kreises: (Text, Mittelwinkel in Grad, von 12 Uhr im UZS)
PIE_LBL_OUT = 8.5
PIE_LBL_IN = 9


def pie_label_text(kind, name):
    return short_cat(kind, name).replace('"', "")


def pie_labels(ser, kind, cats, size=8, bold=False, vals=(), colors=(), rot=0):
    """Kreis – EINE Regel für alle Kreise der Mappe (P1-01): Innen-Etikett „Name · x %“ (weiß bzw. Navy je Fläche,
    fett) nur für Segmente ≥ 35 % – höchstens zwei je Kreis; alle anderen Segmente außen (outEnd) 8,5 pt 0B2A4A mit
    Führungslinie. Bei mehr als 4 Segmenten stehen alle Etiketten außen. Keine Teillegende. 0-%-Segmente ohne Etikett."""
    tot = sum(max(0.0, v) for v in vals) or 1.0
    shares = [(max(0.0, vals[i]) / tot) if i < len(vals) else 0.0 for i in range(len(cats))]
    n_pos = sum(1 for x in shares if x > 0)
    inside = set()
    if n_pos <= PIE_ALL_OUT:
        big = sorted((i for i in range(len(cats)) if shares[i] >= PIE_INSIDE), key=lambda i: -shares[i])
        inside = set(big[:PIE_MAX_INSIDE])
    del PIE_SMALL[:]
    del PIE_OUT[:]
    d = E("c:dLbls")
    acc = 0.0
    for i, name in enumerate(cats):
        short = pie_label_text(kind, name)
        share = shares[i]
        mid = (rot + (acc + share / 2) * 360) % 360
        acc += share
        if share <= 0:
            d.append(dlbl_delete(i))
            continue
        col = colors[i] if i < len(colors) else K.BLUE
        lb = SE(d, "c:dLbl")
        SE(lb, "c:idx", i)
        lb.append(num_fmt(f'[=0]"";"{short} · "0 %'))
        lb.append(sppr(nofill=True, ln=line(nofill=True)))
        if i in inside:
            lb.append(txpr(PIE_LBL_IN, text_on(col), True, wrap="none"))
        else:
            lb.append(txpr(PIE_LBL_OUT, K.NAVY, False, wrap="none"))
            PIE_OUT.append((f"{short} · {round(share * 100)} %", mid, i))
        SE(lb, "c:dLblPos", "ctr" if i in inside else "outEnd")
        _flags(lb, showPercent=True)
    d.append(num_fmt('[=0]"";0 %'))
    d.append(sppr(nofill=True, ln=line(nofill=True)))
    d.append(txpr(PIE_LBL_OUT, K.NAVY, False, wrap="none"))
    SE(d, "c:dLblPos", "outEnd")
    _flags(d, showPercent=True)
    SE(d, "c:showLeaderLines", 1)
    ll = SE(d, "c:leaderLines")
    ll.append(sppr(ln=line(0.75, LEADER)))
    put(ser, d, ORDER["pie_ser"])


# ============================================================================ Serien
def dpt(idx, fill, ln=None, invert=None, bar=False):
    p = E("c:dPt")
    SE(p, "c:idx", idx)
    if bar:
        SE(p, "c:invertIfNegative", 0 if invert is None else invert)
    else:
        SE(p, "c:bubble3D", 0)
    p.append(sppr(solid(fill), ln if ln is not None else line(nofill=True)))
    return p


def invert_ext(color):
    """c14:invertSolidFillFmt – Füllfarbe für negative Werte (Excel 2010+)."""
    ext = etree.Element(q("c:extLst"))
    e = etree.SubElement(ext, q("c:ext"), uri="{6F2FDCE9-48DA-4B69-8628-5D25D57E0C8C}")
    inv = etree.SubElement(e, f"{{{C14_NS}}}invertSolidFillFmt", nsmap={"c14": C14_NS})
    sp = etree.SubElement(inv, f"{{{C14_NS}}}spPr")
    sp.append(solid(color))
    return ext


def _line_style(ser, so, color, w, dash, hidden=False):
    if hidden or color is None:
        put(ser, sppr(ln=line(nofill=True)), so)
    else:
        put(ser, sppr(ln=line(w or 2.0, color, dash, cap="rnd")), so)
    mk = E("c:marker")
    SE(mk, "c:symbol", "none")
    put(ser, mk, so)
    drop(ser, "dPt")
    put_val(ser, "smooth", 0, so)


def _end_label(idx, fmt, color, bold, pos, bg=None, size=8):
    lb = E("c:dLbl")
    SE(lb, "c:idx", idx)
    lb.append(num_fmt(fmt))
    lb.append(sppr(solid(bg), line(nofill=True)) if bg else sppr(nofill=True, ln=line(nofill=True)))
    lb.append(txpr(size, color, bold, wrap="none"))
    SE(lb, "c:dLblPos", pos)
    _flags(lb, showVal=True)
    return lb


def _literal(text):
    """Zahlenformat, das nur einen festen Text zeigt (Endbeschriftung mit Variantennamen)."""
    return '"' + (text or "").replace('"', "'") + '";"' + (text or "").replace('"', "'") + '";"' + \
        (text or "").replace('"', "'") + '"'


def style_series(root, kind, n_cat, size=None):
    """Farben, Linien und Beschriftungen je Reihe. Rückgabe: Legendenindizes, die ausgeblendet werden."""
    plot = root.find(".//" + q("c:plotArea"))
    hidden = []
    k = 0
    bar_totals, bar_counts, gap = {}, {}, {}
    if kind == "jahr1":
        for s in plot.xpath("./c:barChart/c:ser|./c:bar3DChart/c:ser", namespaces=NS):
            is_gap = ser_name(s).strip().lower().startswith("lücke")
            for pt in s.xpath("./c:val//c:pt", namespaces=NS):
                try:
                    v = float(pt.find(q("c:v")).text)
                except (TypeError, ValueError, AttributeError):
                    continue
                i = int(pt.get("idx"))
                if is_gap:
                    gap[i] = gap.get(i, 0) + max(0.0, v)
                    continue
                bar_totals[i] = bar_totals.get(i, 0) + max(0.0, v)
                bar_counts[i] = bar_counts.get(i, 0) + (1 if v > 0 else 0)
    # Einnahmen-Säule mit nur EINEM Segment: Wert im Segment, wenn darüber die „Lücke“ steht – sonst als Summe darüber
    single_inside = {i for i, c in bar_counts.items() if c <= 1 and gap.get(i, 0) > 0}
    for ct in list(plot):
        ctn = local(ct)
        if ctn not in CHART_TAGS:
            continue
        so = ORDER[SER_ORDER.get(ctn, "bar_ser")]
        for ser in ct.findall(q("c:ser")):
            name = ser_name(ser)
            lname = name.strip().lower()
            idx = int(ser.find(q("c:idx")).get("val"))
            if ser.get("helper"):
                continue
            if ctn in ("pieChart", "pie3DChart"):
                cats = categories(ser)
                vals = cache_values(ser)
                drop(ser, "dPt", "explosion")
                if ser.find(q("c:spPr")) is not None:
                    ser.remove(ser.find(q("c:spPr")))
                colors = pie_colors(kind, cats, vals)
                for i, col in enumerate(colors):
                    put(ser, dpt(i, col, ln=line(1, K.WHITE)), so)
                rot = pie_rotation(root) if ctn == "pie3DChart" else 0
                pie_labels(ser, kind, cats, vals=vals, colors=colors, rot=rot)
                continue
            m = match_series(name)
            color, w, dash = m if m else (FALLBACK[k % len(FALLBACK)], None, None)
            if kind == "irr":
                for pat, col, wid in IRR_RULES:
                    if re.search(pat, lname):
                        color, w, dash = col, wid, None
                        break
            afa_cls = ser.get("afa_cls")
            if kind == "afa_kum" and afa_cls:
                # Farbe nach Anwendbarkeit (P33); Legende je Klasse ein Eintrag (Reihenname = Klasse)
                color, w, dash = match_series(afa_cls)
                tx = ser.find(q("c:tx"))
                if tx is not None:
                    for ch_ in list(tx):
                        tx.remove(ch_)
                    SE(tx, "c:v").text = afa_cls
                if not any(abs(v) > 1e-9 for v in cache_values(ser)):
                    color = None                       # Variante ohne Werte (z. B. kein RND-Gutachten): keine Linie
                elif ser.get("afa_halo"):
                    color, w = K.ACCENT, 5.0           # deckungsgleich mit dem Modell: 4A86C8-Rand um die Modell-Linie
            k += 1
            # ---- Linienreihen
            if ctn in ("lineChart", "line3DChart"):
                if kind == "jahr1":                       # Summe über jeder Säule (Linie unsichtbar)
                    _line_style(ser, so, None, None, None, hidden=True)
                    d = dlbls_val(LBL_EUR, 9, K.NAVY, True, pos="t")
                    k_ = 0
                    for i in sorted(single_inside):
                        d.insert(k_, dlbl_delete(i))
                        k_ += 1
                    set_ser_dlbls(ser, d, so)
                    hidden.append(idx)
                    continue
                if kind == "ertrag":
                    if lname.startswith("verbindung"):    # gestrichelte Verbindung zwischen den Stufen (P2-12)
                        _line_style(ser, so, K.MUTED2, 0.75, "dash")
                        drop(ser, "dLbls")
                        hidden.append(idx)
                        continue
                    # Beschriftungsreihen: Name = Betrag aus Zelle; negative Stufe unter dem Balken (P2-12)
                    _line_style(ser, so, None, None, None, hidden=True)
                    neg = name.strip().startswith(("−", "-"))
                    d = dlbls_val("General", 9, K.NAVY, True, pos="b" if neg else "t", ser_name=True)
                    p_ = ref_rows(ser_ref(ser))
                    own = "RSTU".find(p_[1]) if p_ else -1      # Reihe R/S/T/U beschriftet Kategorie 1/2/3/4
                    n_ = len(categories(ser)) or 4
                    for i in reversed(range(n_)):
                        if i != own:
                            d.insert(0, dlbl_delete(i))
                    set_ser_dlbls(ser, d, so)
                    hidden.append(idx)
                    continue
                _line_style(ser, so, color, w, dash)
                n = len(ser.xpath("./c:val//c:pt", namespaces=NS)) or len(cache_values(ser))
                pts = ser.xpath("./c:val//c:pt", namespaces=NS)
                last_idx = max([int(pt_.get("idx")) for pt_ in pts] + [max(n - 1, 0)])
                if kind == "irr" and re.search(r"basis", lname):
                    d = E("c:dLbls")
                    d.append(_end_label(last_idx, "0.0 %", K.NAVY, True, "r"))
                    _flags(d)
                    set_ser_dlbls(ser, d, so)
                elif kind == "afa_kum" and ser.get("afa_top"):
                    d = E("c:dLbls")
                    d.append(_end_label(last_idx, '"Im Modell · "#,##0.0," T€"', K.NAVY, True, "r", bg=K.WHITE))
                    _flags(d)
                    set_ser_dlbls(ser, d, so)
                    hidden.append(idx)
                elif kind == "afa_kum" and afa_cls and color and not ser.get("afa_halo"):
                    d = E("c:dLbls")                   # Endbeschriftung mit dem Variantennamen (P1-15)
                    d.append(_end_label(last_idx, _literal(ser.get("afa_name")), K.MUTED, False, "r"))
                    _flags(d)
                    set_ser_dlbls(ser, d, so)
                else:
                    drop(ser, "dLbls")
                continue
            # ---- Säulen/Balken
            if ctn in ("barChart", "bar3DChart"):
                drop(ser, "dPt", "extLst")
                put_val(ser, "invertIfNegative", 0, so)
                if kind == "ertrag":
                    if color is None:                     # unsichtbare Basis
                        put(ser, sppr(nofill=True, ln=line(nofill=True)), so)
                    else:
                        put(ser, sppr(solid(color), line(nofill=True)), so)
                    drop(ser, "dLbls")
                    if name.strip() not in ("Zufluss", "Abfluss", "Ergebnis"):
                        hidden.append(idx)
                    continue
                if kind == "jahr1":
                    vals = cache_values(ser)
                    if lname.startswith("lücke"):
                        # leerer, rot umrandeter Stapelteil „Lücke −453 €“ (P2-02)
                        put(ser, sppr(nofill=True, ln=line(1.0, K.RED, "dash")), so)
                        d = E("c:dLbls")
                        d.append(num_fmt('"Lücke −"#,##0" €";;'))
                        d.append(sppr(nofill=True, ln=line(nofill=True)))
                        d.append(txpr(9, K.RED, True, wrap="none"))
                        SE(d, "c:dLblPos", "ctr")
                        _flags(d, showVal=True)
                        set_ser_dlbls(ser, d, so)
                        hidden.append(idx)
                        continue
                    put(ser, sppr(solid(color), line(0.75, K.WHITE)), so)
                    d = E("c:dLbls")
                    for i, v in enumerate(vals):
                        tot = bar_totals.get(i, 0)
                        single = bar_counts.get(i, 0) <= 1 and i not in single_inside
                        if v <= 0 or (tot and v < 0.1 * tot) or single:
                            d.append(dlbl_delete(i))
                    d.append(num_fmt(LBL_EUR_POS))
                    d.append(sppr(nofill=True, ln=line(nofill=True)))
                    d.append(txpr(9, text_on(color), True, wrap="none"))
                    SE(d, "c:dLblPos", "ctr")
                    _flags(d, showVal=True)
                    set_ser_dlbls(ser, d, so)
                    continue
                if kind in ("afa_summe", "hh", "va"):
                    put(ser, sppr(solid(color or K.BLUE), line(nofill=True)), so)
                    fmt = LBL_TEUR1 if kind == "afa_summe" else LBL_EUR_POS
                    set_ser_dlbls(ser, dlbls_val(fmt, 8, K.INK2, pos="outEnd"), so)
                    continue
                if color == NEG:
                    put(ser, sppr(solid(color, NEG_ALPHA), line(nofill=True)), so)
                else:
                    put(ser, sppr(solid(color), line(nofill=True)), so)
                drop(ser, "dLbls")
                if kind == "restschuld" and lname.startswith("ende zinsbindung"):
                    hidden.append(idx)                  # Beschriftung am Balken statt Legendeneintrag
                if kind == "steuer" and lname.startswith("steuerliches ergebnis"):
                    vals = cache_values(ser)
                    if vals:
                        i_min = min(range(len(vals)), key=lambda i: vals[i])
                        rest = [abs(v) for i, v in enumerate(vals) if i != i_min] or [0]
                        if vals[i_min] < 0 and abs(vals[i_min]) > 2 * max(rest):
                            # Ausreißer (Sofortabzüge Jahr 1) als Wert an der Säule statt gekappter Achse (P1-15)
                            d = E("c:dLbls")
                            lb_ = _end_label(i_min, LBL_TEUR1, K.INK2, True, "outEnd")
                            w_ = size[0] if size else 900
                            lb_.insert(1, lbl_offset(round(34 / w_, 4), -0.045))
                            d.append(lb_)
                            _flags(d)
                            set_ser_dlbls(ser, d, so)
                continue
            if ctn == "areaChart":
                put(ser, sppr(solid(color, 10), line(nofill=True)), so)
                drop(ser, "dLbls")
    return hidden


def dashboard_series(root):
    """Dashboard (Agent G): Farben/Beschriftungen bleiben; nur Schrift vereinheitlichen."""
    for rp in list(root.iter(q("a:defRPr"))) + list(root.iter(q("a:rPr"))):
        lat = rp.find(q("a:latin"))
        if lat is None:
            lat = etree.SubElement(rp, q("a:latin"))
        lat.set("typeface", K.SANS)


# ============================================================================ Bestandsdiagramme: Fläche + Verkaufsjahr
def _numcache(vals, fmt="General"):
    cache = E("c:numCache")
    SE(cache, "c:formatCode").text = fmt
    SE(cache, "c:ptCount", len(vals))
    for i, v in enumerate(vals):
        if v is None:
            continue
        pt = SE(cache, "c:pt", idx=i)
        SE(pt, "c:v").text = repr(float(v)) if not float(v).is_integer() else str(int(v))
    return cache


def bestand_extras(root, mark):
    """Nettovermögen als Fläche (4A86C8, 25 %) unter den Linien – die Legende zeigt die Fläche; Verkaufsjahr als schmale
    Säule mit Beschriftung oberhalb der Linie auf weißem Grund."""
    plot = root.find(".//" + q("c:plotArea"))
    lc = plot.find(q("c:lineChart"))
    if lc is None or plot.find(q("c:areaChart")) is not None:
        return []
    axids = [a.get("val") for a in lc.findall(q("c:axId"))]
    sers = lc.findall(q("c:ser"))
    nv = next((s for s in sers if ser_name(s).lower().startswith("nettoverm")), None)
    for s_ in sers:
        for tag in ("c:idx", "c:order"):
            el = s_.find(q(tag))
            el.set("val", str(int(el.get("val")) + 2))
    nxt = 0
    hidden = []
    area = E("c:areaChart")
    SE(area, "c:grouping", "standard")
    SE(area, "c:varyColors", 0)
    if nv is not None:
        s = E("c:ser")
        SE(s, "c:idx", nxt)
        SE(s, "c:order", nxt)
        tx = deepcopy(nv.find(q("c:tx")))
        if tx is not None:
            s.append(tx)
        s.append(sppr(solid(K.ACCENT, 25), line(nofill=True)))
        s.append(deepcopy(nv.find(q("c:cat"))))
        s.append(deepcopy(nv.find(q("c:val"))))
        s.set("helper", "1")
        area.append(s)
        # Legende: die kräftige Navy-Linie steht in der Legende, die Fläche darunter nicht (P08)
        hidden.append(nxt)
        nxt += 1
    for a_ in axids:
        SE(area, "c:axId", a_)
    bar = None
    if mark and nv is not None:
        n = len(cache_values(nv))
        cols = [K.L(c) for c in range(K.col("D"), K.col("D") + n)]
        vals = [mark.get(c) or 0 for c in cols]
        at = next((i for i, v in enumerate(vals) if v > 0), None)
        if at is not None:
            # Verkaufsjahr (P3-03): schmale navy Säule bis „Immobilienwert“ (= Lot zur Achse, nur im Verkaufsjahr ≠ 0);
            # Beschriftung „Verkaufspreis Jahr 12 · 377 T€“ darüber auf weißem Grund.
            bar = E("c:barChart")
            SE(bar, "c:barDir", "col")
            SE(bar, "c:grouping", "clustered")
            SE(bar, "c:varyColors", 0)
            s = E("c:ser")
            SE(s, "c:idx", nxt)
            SE(s, "c:order", nxt)
            tx = SE(s, "c:tx")
            sr = SE(tx, "c:strRef")
            SE(sr, "c:f").text = f"Diagramme!$B${MARK_ROW}"
            sc = SE(sr, "c:strCache")
            SE(sc, "c:ptCount", 1)
            SE(SE(sc, "c:pt", idx=0), "c:v").text = f"Verkaufspreis Jahr {at + 1}"
            s.append(sppr(solid(K.NAVY, 75), line(nofill=True)))   # schmale navy Säule = Lot zur Achse
            SE(s, "c:invertIfNegative", 0)
            d = SE(s, "c:dLbls")

            def _mark_lbl(parent, on=True):
                parent.append(num_fmt('#,##0," T€";;'))
                parent.append(sppr(solid(K.WHITE), line(nofill=True)))
                parent.append(txpr(8, K.NAVY, True, wrap="none"))
                SE(parent, "c:dLblPos", "outEnd")
                _flags(parent, showVal=on, showSerName=on)
                SE(parent, "c:separator").text = " · "
            lb = SE(d, "c:dLbl")
            SE(lb, "c:idx", at)
            lb.append(lbl_offset(0.0, -0.05))
            _mark_lbl(lb)
            _mark_lbl(d, on=False)
            s.append(deepcopy(nv.find(q("c:cat"))))
            val = SE(s, "c:val")
            nr = SE(val, "c:numRef")
            SE(nr, "c:f").text = f"Diagramme!$D${MARK_ROW}:${cols[-1]}${MARK_ROW}"
            nr.append(_numcache(vals))
            s.set("helper", "1")
            bar.append(s)
            SE(bar, "c:gapWidth", 500)
            SE(bar, "c:overlap", 0)
            for a_ in axids:
                SE(bar, "c:axId", a_)
            hidden.append(nxt)
    lc.addprevious(area)
    if bar is not None:
        lc.addprevious(bar)
    return hidden


def restschuld_extras(root):
    """Restschuld: Beschriftung „Ende Zinsbindung“ über der hervorgehobenen Säule – als unsichtbare Linienreihe auf
    derselben Zeile (#NV außer im Zinsbindungsjahr), damit sie in Excel dynamisch mitwandert (P3-03)."""
    plot = root.find(".//" + q("c:plotArea"))
    bc = plot.find(q("c:barChart"))
    if bc is None or plot.find(q("c:lineChart")) is not None:
        return []
    src = next((x for x in bc.findall(q("c:ser")) if ser_name(x).lower().startswith("ende zinsbindung")), None)
    if src is None:
        return []
    vals = cache_values(src)
    pts = src.xpath("./c:val//c:pt", namespaces=NS)
    at = next((int(p_.get("idx")) for p_ in pts if float(p_.find(q("c:v")).text or 0) > 0), None)
    if at is None:
        return []
    lc = E("c:lineChart")
    SE(lc, "c:grouping", "standard")
    SE(lc, "c:varyColors", 0)
    s = deepcopy(src)
    for tag in ("c:idx", "c:order"):
        s.find(q(tag)).set("val", "91")
    drop(s, "spPr", "invertIfNegative", "dPt", "dLbls", "shape", "extLst")
    s.set("helper", "1")
    so = ORDER["line_ser"]
    put(s, sppr(ln=line(nofill=True)), so)
    mk = E("c:marker")
    SE(mk, "c:symbol", "none")
    put(s, mk, so)
    d = E("c:dLbls")
    lb = SE(d, "c:dLbl")
    SE(lb, "c:idx", at)
    lb.append(lbl_offset(0.0, -0.035))
    lb.append(sppr(solid(K.WHITE), line(nofill=True)))
    lb.append(txpr(8, K.BLUE, True, wrap="none"))
    SE(lb, "c:dLblPos", "r")
    _flags(lb, showSerName=True)
    _flags(d)
    put(s, d, so)
    put_val(s, "smooth", 0, so)
    lc.append(s)
    SE(lc, "c:marker", 1)
    for a_ in bc.findall(q("c:axId")):
        SE(lc, "c:axId", a_.get("val"))
    bc.addnext(lc)
    return [91]


def afa_prepare(root):
    """AfA-Vergleich „Kumulierte AfA“ (P1-15): Varianten nach Klasse ordnen – Legende „Im Modell angewendet ·
    Anwendbar · Nicht anwendbar“; deckungsgleiche anwendbare Variante wird zum 4A86C8-Rand der Modell-Linie; die
    zweite Modell-Reihe liegt zuoberst und trägt das Endetikett."""
    lc = root.find(".//" + q("c:lineChart"))
    if lc is None:
        return
    sers = lc.findall(q("c:ser"))
    models = [x for x in sers if ser_name(x).strip().lower().startswith("im modell")]
    variants = [x for x in sers if x not in models]
    mvals = cache_values(models[0]) if models else []
    for k_, x in enumerate(variants):
        cls = AFA_CLASS.get(k_, "Nicht anwendbar")
        x.set("afa_cls", cls)
        x.set("afa_name", ser_name(x))
        v = cache_values(x)
        if cls == "Anwendbar" and mvals and len(v) == len(mvals) and all(abs(a - b) < 0.5 for a, b in zip(v, mvals)):
            x.set("afa_halo", "1")
    order = models[:1] + [x for x in variants if x.get("afa_cls") == "Anwendbar"] + \
        [x for x in variants if x.get("afa_cls") != "Anwendbar"] + models[1:]
    if len(models) > 1:
        models[-1].set("afa_top", "1")
    anchor = sers[0].getprevious()
    for x in sers:
        lc.remove(x)
    for i, x in enumerate(order):
        x.find(q("c:idx")).set("val", str(i))
        x.find(q("c:order")).set("val", str(i))
    for x in reversed(order):
        if anchor is not None:
            anchor.addnext(x)
        else:
            lc.insert(0, x)


def pie_need(size, titled, outs):
    """Platzbedarf eines Kreises: (Durchmesser, linker Rand, rechter Rand, oben, unten) in px."""
    w, h = size if size else (400, 300)
    left = right = 10.0
    top = (34 if titled else 8) + 6.0
    bottom = 8.0
    for text, ang, *_ in outs:
        tw = K.text_width(text, PIE_LBL_OUT) + 10
        sx, cy = math.sin(math.radians(ang)), -math.cos(math.radians(ang))
        if sx > 0.35:
            right = max(right, tw + 8)
        elif sx < -0.35:
            left = max(left, tw + 8)
        else:
            left, right = max(left, tw / 2), max(right, tw / 2)
        if cy < -0.5:
            top = max(top, (34 if titled else 8) + 18)
        elif cy > 0.5:
            bottom = max(bottom, 20)
    d = max(70.0, min(260.0, w - left - right, (h - top - bottom) / 0.8))
    return d, left, right, top, bottom


def pie_layout(root, size=None, titled=True, legend=False, fixed_d=None):
    """Kreis (P1-01): Durchmesser und Lage aus dem Platz, den die Außen-Etiketten brauchen; die Kreisfläche liegt
    zwischen linkem und rechtem Etikettenrand zentriert, vertikal mittig. fixed_d: gleicher Kreis in einer Reihe."""
    plot = root.find(".//" + q("c:plotArea"))
    drop(plot, "layout")
    w, h = size if size else (400, 300)
    d, left, right, top, bottom = pie_need(size, titled, list(PIE_OUT))
    if fixed_d:
        d = min(d, fixed_d)
    LAST_PIE_D[0] = d
    free = w - left - right - d
    x = (left + max(0.0, free) / 2) / w
    # Kreis mittig im Rahmen, wenn die Etiketten es zulassen (ruhiger), sonst zwischen den Etikettenrändern
    xc = (1 - d / w) / 2
    if left <= xc * w and right <= xc * w:
        x = xc
    fw, fh = d / w, min(d * 0.8 / h, 1 - (top + bottom) / h)
    y = (top + (h - top - bottom - d * 0.8) / 2) / h
    plot.insert(0, manual_layout("inner", round(x, 4), round(max(0.0, y), 4), round(fw, 4), round(fh, 4)))
    spread_pie_labels(root, x * w + d / 2, y * h + 0.38 * d, d, w, h)


PIE_GAP = 15.0           # Mindestabstand zweier Außen-Etiketten (px, 8,5 pt)


def spread_pie_labels(root, cx, cy, d, w, h):
    """Außen-Etiketten kleiner Nachbarsegmente senkrecht entzerren (je Seite, Mindestabstand PIE_GAP) – als
    Versatz der Etiketten (manualLayout, Anteil der Diagrammhöhe); Führungslinien zeigen die Zuordnung."""
    outs = list(PIE_OUT)
    if len(outs) < 2:
        return
    rx, ry = d / 2 + 6, 0.30 * d + 6
    lbls = {}
    for dl in root.iter(q("c:dLbl")):
        i = dl.find(q("c:idx"))
        if i is not None:
            lbls[int(i.get("val"))] = dl
    for side in (1, -1):
        pts = []
        for text, ang, idx in outs:
            sx, cy_ = math.sin(math.radians(ang)), -math.cos(math.radians(ang))
            if (sx >= 0) != (side > 0):
                continue
            y0 = cy + ry * cy_ + (0.1 * d if cy_ > 0 else 0)
            pts.append([y0, y0, idx])
        if len(pts) < 2:
            continue
        pts.sort(key=lambda p_: p_[0])
        for k_ in range(1, len(pts)):
            pts[k_][1] = max(pts[k_][1], pts[k_ - 1][1] + PIE_GAP)
        shift = sum(p_[1] - p_[0] for p_ in pts) / len(pts)
        # zurückschieben, aber nie über das erste (oberste) Etikett hinaus nach oben
        shift = min(shift, max(0.0, pts[0][1] - 4))
        for p_ in pts:
            p_[1] -= shift
        for k_ in range(1, len(pts)):
            pts[k_][1] = max(pts[k_][1], pts[k_ - 1][1] + PIE_GAP)
        for y0, y1, idx in pts:
            dy = y1 - y0
            dl = lbls.get(idx)
            if dl is None or abs(dy) < 1:
                continue
            drop(dl, "layout")
            dl.insert(1, lbl_offset(round(side * 6 / w, 4), round(dy / h, 4)))


def set_pie_legend(root, n_cat, small, size=None):
    """Kreise tragen keine Legende mehr (P1-01) – alle Segmente sind direkt beschriftet."""
    drop(root.find(q("c:chart")), "legend")
    return False


def legend_hidden(plot, kind, hidden):
    """Legendeneinträge ohne sichtbares Element (P19): Reihen ohne Werte (nur 0/#NV) – z. B. „Darlehen II“ ohne
    zweites Darlehen, „Sonder-AfA“ ohne Sonder-AfA; AfA-Vergleich: je Klasse nur ein Eintrag."""
    out = list(hidden)
    seen = set()
    for s in plot.xpath("./*/c:ser", namespaces=NS):
        idx = int(s.find(q("c:idx")).get("val"))
        if idx in out:
            continue
        vals = cache_values(s)
        empty = not any(abs(v) > 1e-9 for v in vals)
        name = ser_name(s).strip().lower()
        if kind == "afa_kum":
            if empty or name in seen:
                out.append(idx)
            else:
                seen.add(name)
        elif empty and kind not in ("ertrag", "jahr1"):
            out.append(idx)
    return out


# ============================================================================ Hauptfunktion je Diagramm
LAST_PIE_D = [None]
TOP_ROW_PIES = {"invest", "finanz", "kpa"}      # obere Reihe „Diagramme“: gleicher Kreis (P1-01)


def style_chart(xml, sheet=None, mark=None, size=None, pie_d=None):
    root = etree.fromstring(xml)
    chart = root.find(q("c:chart"))
    plot = chart.find(q("c:plotArea"))
    kind = chart_kind(sheet or "", root)
    fix_literal_names(root)
    simplify_categories(root)
    sers = plot.xpath("./*/c:ser", namespaces=NS)
    n_cat = max([len(categories(s)) for s in sers] + [0])
    dashboard = kind == "dash"
    width_px = size[0] if size else None
    hidden = []
    del PIE_SMALL[:]
    if kind == "bestand":
        hidden = bestand_extras(root, mark)
    if kind == "afa_kum":
        afa_prepare(root)
    kind3d = None if dashboard else make_3d(root, n_cat, kind)
    if kind == "restschuld":
        hidden += restschuld_extras(root)
    horizontal = plot.find(".//" + q("c:barDir")) is not None and plot.find(".//" + q("c:barDir")).get("val") == "bar"
    is_pie = kind3d == "pie" or kind in PIE_KINDS

    # Titel: Schrittseiten und Dashboard tragen den Panel-Kopf in der Zelle darüber; Einheit = Achsenformat (P09)
    # „Diagramme“: Titel steht als Unterabschnitt in der Zelle über dem Diagramm (P2-10)
    delete_title = dashboard or bool(re.match(r"S\d\d ", sheet or "")) or sheet == "Diagramme"
    title = TITLES.get(kind)
    if title:
        title = title.format(n=n_cat, u=value_format(root, kind, n_cat, size)[1])
    set_title(root, title, delete_title, width_px)

    if dashboard:
        dashboard_series(root)
    else:
        hidden += style_series(root, kind, n_cat, size)
    for s in root.iter(q("c:ser")):
        if "helper" in s.attrib:
            del s.attrib["helper"]

    # Legende: immer unten (Reihenfolge = Stapelreihenfolge); Kreise rechts nur mit den Kleinstsegmenten
    has_pie_legend = False
    if is_pie:
        has_pie_legend = set_pie_legend(root, n_cat, list(PIE_SMALL), size)
    elif not dashboard:
        hidden = legend_hidden(plot, kind, hidden)
        visible = [s for s in plot.xpath("./*/c:ser", namespaces=NS)
                   if int(s.find(q("c:idx")).get("val")) not in hidden]
        # Balken (Haushalt/Vermögen): Legende direkt unter dem Titel (P2-09)
        set_legend(root, kind != "cf_nach" and len(visible) > 1, hidden, "t" if kind in ("hh", "va") else "b")
    else:
        lg = chart.find(q("c:legend"))
        if lg is not None:
            put(lg, txpr(8, K.MUTED), ORDER["legend"])

    ax_kind = kind
    if dashboard and plot.find(q("c:barChart")) is None:
        ax_kind = "dash_line"
    style_axes(root, ax_kind, n_cat, size, horizontal)
    if is_pie:
        pie_layout(root, size, titled=not delete_title, legend=has_pie_legend, fixed_d=pie_d)
    elif kind in ("hh", "va"):
        # Plotfläche über die volle Rahmenhöhe (P2-09): Titel + Legende oben, darunter die Balken
        w, h = size if size else (500, 300)
        top = (34 + (22 if chart.find(q("c:legend")) is not None else 0)) / h
        drop(plot, "layout")
        plot.insert(0, manual_layout("outer", 0.01, round(top, 4), 0.97, round(max(0.3, 1 - top - 0.03), 4)))
    elif kind == "afa_kum":
        # rechts ≈ 18 % frei für die Endbeschriftungen der Varianten (P1-15)
        w, h = size if size else (1000, 320)
        drop(plot, "layout")
        plot.insert(0, manual_layout("inner", round(52 / w, 4), round(40 / h, 4), 0.76, round(1 - 40 / h - 62 / h, 4)))

    if not dashboard:                  # Linien nie gestapelt (LibreOffice schreibt Kombi-Linien als „stacked“)
        for lc_ in plot.findall(q("c:lineChart")):
            g_ = lc_.find(q("c:grouping"))
            if g_ is not None:
                g_.set("val", "standard")
    for s_ in root.iter(q("c:ser")):
        for a_ in ("afa_cls", "afa_name", "afa_halo", "afa_top"):
            if a_ in s_.attrib:
                del s_.attrib[a_]
    put_val(chart, "plotVisOnly", 0, ORDER["chart"])
    put_val(chart, "dispBlanksAs", "gap", ORDER["chart"])
    # Plot- und Diagrammfläche ohne Füllung und Rahmen (weiß), Grundschrift 8 pt grau; Sensitivität im Panelstil F3F7FC
    psp = plot.find(q("c:spPr"))
    if psp is not None:
        plot.remove(psp)
    ext = plot.find(q("c:extLst"))
    (ext.addprevious if ext is not None else plot.append)(sppr(nofill=True, ln=line(nofill=True)))
    panel = kind == "irr"
    put(root, sppr(solid(K.TINT_XL), line(nofill=True)) if panel else sppr(nofill=True, ln=line(nofill=True)),
        ORDER["chartSpace"])
    put(root, txpr(8, K.MUTED), ORDER["chartSpace"])
    put_val(root, "roundedCorners", 0, ORDER["chartSpace"])
    for lat in root.iter(q("a:latin")):
        lat.set("typeface", K.SANS)
    for tag in ("a:cs", "a:ea"):
        for el in root.iter(q(tag)):
            el.set("typeface", K.SANS)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def finish(path):
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        files = {i.filename: zin.read(i.filename) for i in infos}
    charts, sheets = chart_sheets(files)
    mark = row_values(files, sheets.get("Diagramme", ""), MARK_ROW)
    AFA_CLASS.clear()
    for k in range(8):
        rv = row_values(files, sheets.get("Diagramme", ""), AFA_SUM_ROW + k)
        AFA_CLASS[k] = "Anwendbar" if (rv.get("H") or 0) > 0 or (rv.get("G") or 0) > 0 else "Nicht anwendbar"
    sizes = chart_sizes(files)
    orig, top_d = {}, {}
    for name in sorted(files):
        if name.startswith("xl/charts/chart") and name.endswith(".xml"):
            try:
                orig[name] = files[name]
                LAST_PIE_D[0] = None
                files[name] = style_chart(files[name], charts.get(name), mark, sizes.get(name))
                if charts.get(name) == "Diagramme" and LAST_PIE_D[0] and \
                        chart_kind("Diagramme", etree.fromstring(orig[name])) in TOP_ROW_PIES:
                    top_d[name] = LAST_PIE_D[0]
            except Exception as exc:  # ein Diagramm darf den Build nicht abbrechen
                print(f"WARNUNG Diagramm {name} ({charts.get(name)}): {exc!r}", file=sys.stderr)
    if len(top_d) > 1:                 # gleicher Kreisdurchmesser in der oberen Reihe von „Diagramme“
        d_min = min(top_d.values())
        for name in top_d:
            try:
                files[name] = style_chart(orig[name], charts.get(name), mark, sizes.get(name), pie_d=d_min)
            except Exception as exc:
                print(f"WARNUNG Diagramm {name} ({charts.get(name)}): {exc!r}", file=sys.stderr)
    try:
        import finish_sheets  # Blatt-XML-Korrekturen (Registerfarben, Datenüberprüfung, Ansicht)
        finish_sheets.fix(files)
    except ImportError:
        pass
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            zout.writestr(i, files[i.filename])
        for name in files:
            if name not in {i.filename for i in infos}:
                zout.writestr(name, files[name])
    print(f"Nachlauf abgeschlossen: {path} ({len(charts)} Diagramme)")


if __name__ == "__main__":
    finish(sys.argv[1])
