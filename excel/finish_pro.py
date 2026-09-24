"""Nachlauf NACH der LibreOffice-Neuberechnung: Styling aller Diagramme und Blatt-XML-Korrekturen.

Aufruf:  python excel/finish_pro.py MAPPE.xlsx   (ändert die Datei an Ort und Stelle)

LibreOffice schreibt beim Roundtrip alle Diagramme neu (Schriften, Farben, Kategorien als multiLvlStrRef, Datenbeschriftungen
je Punkt). Deshalb wird das Diagramm-Styling hier – nach der Neuberechnung – vollständig und regelbasiert gesetzt:

* Diagrammart wird aus den Datenbereichen erkannt (nicht aus Nummer oder Titel), Farben nach Serien- bzw. Kategoriename.
* Titel: auf Schrittseiten und Dashboard entfernt (Panel-Kopf in der Zelle darüber), sonst links, 10 pt fett Navy.
* Achsen: Werteachse ohne Linie, ab 10.000 in „T€“, Hauptgitter 0,5 pt E6E8EB; Jahresachsen waagerecht im 5er-/2er-Takt,
  Beschriftung immer unten (tickLblPos low); Kategorien als einfache Zahlen-/Textreihe.
* Legende unten 8 pt, bei Kreisen und Einzelreihen entfernt; 3D bleibt (ruhige Parameter), Kreise ohne Label-Kollision.
Zellwerte werden nicht berührt, die berechneten Ergebnisse bleiben gültig.
"""
import html
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


def num_fmt(code):
    return E("c:numFmt", formatCode=code, sourceLinked="0")


# ============================================================================ Design-Regeln
FMT_TEUR = '#,##0," T€"'
FMT_EUR = '#,##0" €"'
FMT_PCT = "0 %"
LBL_EUR = '#,##0" €";-#,##0" €"'
LBL_EUR_POS = '#,##0" €";;'
DARK = {"0B2A4A", "1D4F8A", "4A86C8", "6E7F96", "8A9099", "2D63A6"}

# Serienfarben nach Name: (Muster, Farbe, Linienstärke pt, Strichart)
SERIES_RULES = [
    (r"^immobilienwert", K.BLUE, 2.25, None),
    (r"^nettoverm", K.NAVY, 2.25, None),
    (r"^restschuld", K.MUTED2, 1.5, "dash"),
    (r"^darlehen ii\b", K.SKY, None, None),
    (r"^darlehen i\b", K.BLUE, None, None),
    (r"^nettokaltmiete", K.NAVY, None, None),
    (r"^steuererstattung", K.BLUE, None, None),
    (r"^steuerzahlung", "6E7F96", None, None),
    (r"^(kumulierte )?zinsen", K.ACCENT, 2.25, None),
    (r"^(kumulierte )?tilgung", K.SKY, 2.25, None),
    (r"^bewirtschaftung", K.MUTED2, None, None),
    (r"^cashflow vor", K.SKY, None, None),
    (r"^cashflow nach", K.NAVY, None, None),
    (r"^kumulierter cashflow", K.NAVY, 2.25, None),
    (r"^kumulierte steuer", "6E7F96", 2.25, None),
    (r"^afa regul", K.NAVY, None, None),
    (r"sonder-afa", K.ACCENT, None, None),
    (r"7h", K.BLUE, None, None),
    (r"^bewegliche", K.SKY, None, None),
    (r"^steuerliches ergebnis", K.BLUE, None, None),
    (r"^steuer\b", "6E7F96", None, None),
    # AfA-Varianten (AfA-Vergleich)
    (r"^linear 2,0", K.BLUE, 1.75, None),
    (r"^linear 2,5", K.ACCENT, 1.75, None),
    (r"^linear 3,0", K.SKY, 1.75, None),
    (r"^gutachten", "6E7F96", 1.75, "dash"),
    (r"^degressiv 5", K.MUTED2, 1.75, None),
    (r"^linear 3 ?% ?\+", K.ACCENT, 1.75, "dash"),
    (r"^degressiv ?\+", K.BLUE, 1.75, "dash"),
    (r"^im modell", K.NAVY, 2.5, None),
    # Mietszenarien (Sensitivität)
    (r"miete\s*-\s*10", K.SKY, 1.25, None),
    (r"miete\s*-\s*5", "6E9FD6", 1.25, None),
    (r"miete\s*basis", K.NAVY, 2.5, None),
    (r"miete\s*\+\s*5", K.ACCENT, 1.25, None),
    (r"miete\s*\+\s*10", K.BLUE, 1.25, None),
]
FALLBACK = [K.NAVY, K.ACCENT, K.SKY, K.BLUE, "6E7F96", K.MIST]

# Kategorie-Farben und Kurzlabels (Kreise, Einzelreihen): Art → [(Muster, Farbe, Kurzname)]
CAT_RULES = {
    "invest": [(r"^kaufpreis", K.NAVY, "Kaufpreis"), (r"^kaufneben", K.ACCENT, "Nebenkosten"),
               (r"^finanzierungsneben", "6E7F96", "Finanzierungskosten"), (r"^ma(ß|ss)nahmen", K.SKY, "Maßnahmen")],
    "finanz": [(r"^darlehen ii\b", K.SKY, "Darlehen II"), (r"^darlehen i\b", K.BLUE, "Darlehen I"),
               (r"^eigenkapital", K.NAVY, "Eigenkapital")],
    "kpa": [(r"^geb", K.NAVY, "Gebäude"), (r"^grund", K.SKY, "Boden"), (r"^beweg", K.ACCENT, "Inventar"),
            (r"r(ü|ue)cklage", K.MIST, "Rücklage")],
    "kanc": [(r"^grunderwerb", K.NAVY, "Grunderwerbsteuer"), (r"^notar", K.BLUE, "Notar"),
             (r"^grundbuch", K.ACCENT, "Grundbuch"), (r"^makler", K.SKY, "Makler"), (r"^sonst", K.MIST, "Sonstige")],
    "bewirt": [(r"^hausgeld", K.NAVY, "Hausgeld"), (r"erhaltungsr", K.BLUE, "Erhaltungsrücklage"),
               (r"^verwaltung", K.ACCENT, "Verwaltung"), (r"instandhaltung", K.SKY, "Instandhaltung"),
               (r"^mietausfall", K.MIST, "Mietausfall"), (r"werbungskosten", "6E7F96", "Werbungskosten")],
    "exit": [(r"^verkaufskosten", K.SKY, "Verkaufskosten"), (r"^steuer", "6E7F96", "Steuer"),
             (r"restschuld", K.ACCENT, "Restschuld"), (r"^nettoerl", K.NAVY, "Nettoerlös")],
    "ertrag": [(r"^kumulierter", K.BLUE, None), (r"^nettoerl", K.BLUE, None), (r"eigenkapital", K.MUTED2, None),
               (r"gesamtertrag", K.NAVY, None)],
    "afa_summe": [(r"^linear 2,0", K.BLUE, None), (r"^linear 2,5", K.ACCENT, None), (r"^linear 3,0", K.SKY, None),
                  (r"^gutachten", "6E7F96", "pattern"), (r"^degressiv 5", K.MUTED2, None),
                  (r"^linear 3 ?% ?\+", K.ACCENT, "pattern"), (r"^degressiv ?\+", K.BLUE, "pattern"),
                  (r"^im modell", K.NAVY, None)],
}
PIE_KINDS = {"invest", "finanz", "kpa", "kanc", "bewirt", "exit"}
DIRECT_LABELS = {"jahr1", "afa_summe", "hh", "va"}

# Einheitliche Titel „Kennzahl (Zeitraum, Einheit)“ – identisch für identische Diagramme
TITLES = {
    "invest": "Gesamtinvestition",
    "finanz": "Finanzierungsstruktur",
    "kpa": "Kaufpreisaufteilung",
    "exit": "Verwendung des Verkaufserlöses",
    "ertrag": "Gesamtertrag nach Steuern (Zusammensetzung, €)",
    "jahr1": "Einnahmen vs. Ausgaben (Jahr 1, € / Monat)",
    "bestand": "Vermögen und Restschuld (Jahre 1–35, €)",
    "restschuld": "Restschuld am Jahresende (Jahre 1–35, €)",
    "cashflow": "Cashflow vor und nach Steuern (Jahre 1–30, € p. a.)",
    "cf_nach": "Cashflow nach Steuern (Jahre 1–30, € p. a.)",
    "cf_kum": "Kumulierter Cashflow nach Steuern (Jahre 1–30, €)",
    "kumzins": "Kumulierte Zinsen, Tilgung und Steuer (Jahre 1–35, €)",
    "afa": "Abschreibungen nach Komponenten (Jahre 1–20, € p. a.)",
    "steuer": "Steuerliches Ergebnis und Steuer (Jahre 1–20, € p. a.)",
    "kanc": "Zusammensetzung der Kaufnebenkosten",
    "bewirt": "Bewirtschaftungskosten (Jahr 1, € p. a.)",
    "afa_summe": "Summe AfA je Variante (Jahre 1–10, €)",
    "afa_kum": "Kumulierte AfA je Variante (Jahre 1–10, €)",
    "irr": "IRR nach Steuern je Wertsteigerung (Linien: Mietszenarien)",
    "hh": "Ausgaben (€ / Monat)",
    "va": "Vermögenswerte (€)",
}
MARK_ROW = 202          # Diagramme: Immobilienwert im Verkaufsjahr (sonst 0), angelegt von layouts/diagramme.py
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


def chart_widths(files):
    """{'xl/charts/chartN.xml': Breite in px} aus den Zeichnungsebenen (xfrm/ext)."""
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
                out[tgt] = int(ext.get("cx", "0")) / 9525
    return out


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
    sers = root.xpath(".//c:plotArea/*/c:ser", namespaces=NS)
    rows, shs = set(), set()
    for s in sers:
        p = ref_rows(ser_ref(s))
        if p:
            shs.add(p[0])
            rows.add(p[2])
    if shs == {"Diagramme"}:
        if rows & {179, 182, 183}:
            return "bestand"
        if rows and rows <= {180, 181}:
            return "restschuld"
        if 184 in rows and 185 in rows:
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
        for r, k in {141: "invest", 147: "finanz", 152: "kpa", 158: "exit", 164: "ertrag"}.items():
            if r in rows:
                return k
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
    return None


def match_series(name):
    n = (name or "").strip().lower()
    for pat, color, w, dash in SERIES_RULES:
        if re.search(pat, n):
            return color, w, dash
    return None


def match_cat(kind, name):
    for pat, color, short in CAT_RULES.get(kind, []):
        if re.search(pat, (name or "").strip().lower()):
            return color, short
    return None, None


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


def make_3d(root, n_cat):
    """Säulen und Kreise in ruhige, schemakonforme 3D-Varianten überführen (Kombidiagramme bleiben 2D)."""
    chart = root.find(q("c:chart"))
    plot = chart.find(q("c:plotArea"))
    types = [local(e) for e in plot if local(e) in CHART_TAGS]
    if len(types) != 1:
        return None
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
        SE(v, "c:rotX", 25)
        SE(v, "c:rotY", 90)
        SE(v, "c:rAngAx", 0)
        SE(v, "c:perspective", 15)
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


# ============================================================================ Titel, Legende, Achsen
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
    t = E("c:title")
    tx = SE(t, "c:tx")
    rich = SE(tx, "c:rich")
    SE(rich, "a:bodyPr", rot="0", vert="horz", wrap="square")
    SE(rich, "a:lstStyle")
    p = SE(rich, "a:p")
    ppr = SE(p, "a:pPr", algn="l")
    ppr.append(rpr("a:defRPr", 10, K.NAVY, True))
    r = SE(p, "a:r")
    r.append(rpr("a:rPr", 10, K.NAVY, True))
    r[-1].set("lang", "de-DE")
    SE(r, "a:t").text = text
    # linke Kante des Titels auf der Textkante der Bänder (Einzug 1 ≈ 12 px)
    x = round(min(0.05, max(0.004, 12 / width_px)), 4) if width_px else 0.012
    t.append(manual_layout(x=x, y=0.025))
    SE(t, "c:overlay", 0)
    t.append(sppr(nofill=True, ln=line(nofill=True)))
    put(chart, t, ORDER["chart"])
    put_val(chart, "autoTitleDeleted", 0, ORDER["chart"])


def set_legend(root, show, hidden_idx=()):
    chart = root.find(q("c:chart"))
    drop(chart, "legend")
    if not show:
        return
    lg = E("c:legend")
    SE(lg, "c:legendPos", "b")
    for i in hidden_idx:
        le = SE(lg, "c:legendEntry")
        SE(le, "c:idx", i)
        SE(le, "c:delete", 1)
    SE(lg, "c:overlay", 0)
    lg.append(sppr(nofill=True, ln=line(nofill=True)))
    lg.append(txpr(8, K.MUTED))
    put(chart, lg, ORDER["chart"])


def style_axes(root, kind, n_cat, values, horizontal=False, dashboard=False):
    plot = root.find(".//" + q("c:plotArea"))
    lo = min(values) if values else 0
    hi = max(values) if values else 0
    big = max(abs(lo), abs(hi))
    if kind == "irr":
        vfmt = FMT_PCT
    elif big >= 10000:
        vfmt = FMT_TEUR
    else:
        vfmt = FMT_EUR
    time_series = n_cat >= 10 and kind not in ("afa_summe", "hh", "va")
    for ax in plot.findall(q("c:valAx")):
        o = ORDER["valAx"]
        sc = ax.find(q("c:scaling"))
        if sc is not None and lo < 0:
            drop(sc, "min")
        if kind in DIRECT_LABELS:
            # direkt beschriftete Diagramme: keine Werteachse, kein Gitter (Werte stehen an den Säulen)
            put_val(ax, "delete", 1, o)
            drop(ax, "majorGridlines", "minorGridlines")
        else:
            put_val(ax, "delete", 0, o)
            gl = E("c:majorGridlines")
            gl.append(sppr(ln=line(0.5, K.LINE)))
            put(ax, gl, o)
            drop(ax, "minorGridlines")
        put(ax, num_fmt(vfmt), o)
        put_val(ax, "majorTickMark", "none", o)
        put_val(ax, "minorTickMark", "none", o)
        put_val(ax, "tickLblPos", "nextTo", o)
        put(ax, sppr(nofill=True, ln=line(nofill=True)), o)
        put(ax, txpr(8, K.MUTED), o)
        if kind == "afa_summe":
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
            put_val(ax, "majorTickMark", "out" if time_series else "none", o)
            put_val(ax, "minorTickMark", "none", o)
            put_val(ax, "tickLblPos", "low", o)
            put(ax, sppr(ln=line(0.75, "C5CDD8" if horizontal else K.MUTED2)), o)
            put(ax, txpr(8, K.MUTED, rot=(60 if time_series else 0) if not horizontal else None, wrap="none" if time_series else None), o)
            if kind == "afa_summe":
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
def dlbls_off():
    d = E("c:dLbls")
    for n, v in (("showLegendKey", 0), ("showVal", 0), ("showCatName", 0), ("showSerName", 0), ("showPercent", 0),
                 ("showBubbleSize", 0)):
        SE(d, "c:" + n, v)
    return d


def dlbls_val(fmt, size, color, bold=False, pos=None, bg=None):
    d = E("c:dLbls")
    d.append(num_fmt(fmt))
    d.append(sppr(solid(bg), line(nofill=True)) if bg else sppr(nofill=True, ln=line(nofill=True)))
    d.append(txpr(size, color, bold, wrap="none"))
    if pos:
        SE(d, "c:dLblPos", pos)
    for n, v in (("showLegendKey", 0), ("showVal", 1), ("showCatName", 0), ("showSerName", 0), ("showPercent", 0),
                 ("showBubbleSize", 0)):
        SE(d, "c:" + n, v)
    return d


def set_ser_dlbls(ser, dl, order):
    drop(ser, "dLbls")
    if dl is not None:
        put(ser, dl, order)


def pie_labels(ser, kind, cats):
    """Kreis: je Segment „Kurzname · 12 %“, unter 3 % ohne Beschriftung, Führungslinien 0,5 pt."""
    d = E("c:dLbls")
    for i, name in enumerate(cats):
        _, short = match_cat(kind, name)
        short = (short or re.sub(r"\s*\(.*?\)", "", name or "")[:22]).replace('"', "")
        lb = SE(d, "c:dLbl")
        SE(lb, "c:idx", i)
        lb.append(num_fmt(f'[<0.03]"";"{short} · "0 %'))
        lb.append(sppr(nofill=True, ln=line(nofill=True)))
        lb.append(txpr(8, K.INK2, wrap="none"))
        SE(lb, "c:dLblPos", "outEnd")
        for n, v in (("showLegendKey", 0), ("showVal", 0), ("showCatName", 0), ("showSerName", 0),
                     ("showPercent", 1), ("showBubbleSize", 0)):
            SE(lb, "c:" + n, v)
    d.append(num_fmt('[<0.03]"";0 %'))
    d.append(sppr(nofill=True, ln=line(nofill=True)))
    d.append(txpr(8, K.INK2, wrap="none"))
    SE(d, "c:dLblPos", "outEnd")
    for n, v in (("showLegendKey", 0), ("showVal", 0), ("showCatName", 0), ("showSerName", 0), ("showPercent", 1),
                 ("showBubbleSize", 0)):
        SE(d, "c:" + n, v)
    SE(d, "c:showLeaderLines", 1)
    ll = SE(d, "c:leaderLines")
    ll.append(sppr(ln=line(0.5, "C5CDD8")))
    put(ser, d, ORDER["pie_ser"])


# ============================================================================ Serien
def dpt(idx, fill, ln=None, invert=None, pattern=None, bar=False):
    p = E("c:dPt")
    SE(p, "c:idx", idx)
    if bar:
        SE(p, "c:invertIfNegative", 0 if invert is None else invert)
    else:
        SE(p, "c:bubble3D", 0)
    if pattern:
        sp = E("c:spPr")
        pf = SE(sp, "a:pattFill", prst="wdUpDiag")
        SE(pf, "a:fgClr").append(E("a:srgbClr", fill))
        SE(pf, "a:bgClr").append(E("a:srgbClr", "FFFFFF"))
        sp.append(line(0.75, fill))
        p.append(sp)
    else:
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


def style_series(root, kind, n_cat):
    plot = root.find(".//" + q("c:plotArea"))
    k = 0
    for ct in list(plot):
        ctn = local(ct)
        if ctn not in CHART_TAGS:
            continue
        so = ORDER[SER_ORDER.get(ctn, "bar_ser")]
        for ser in ct.findall(q("c:ser")):
            name = ser_name(ser)
            if ser.get("helper"):
                continue
            if ctn in ("pieChart", "pie3DChart"):
                cats = categories(ser)
                drop(ser, "dPt", "explosion")
                ser.find(q("c:spPr")) is not None and ser.remove(ser.find(q("c:spPr")))
                for i, cname in enumerate(cats):
                    color, _ = match_cat(kind, cname)
                    color = color or FALLBACK[i % len(FALLBACK)]
                    put(ser, dpt(i, color, ln=line(1, K.WHITE)), so)
                pie_labels(ser, kind, cats)
                continue
            if kind in ("ertrag", "afa_summe", "hh", "va"):
                cats = categories(ser)
                drop(ser, "dPt", "extLst")
                put(ser, sppr(solid(K.BLUE), line(nofill=True)), so)
                # Gesamtertrag: Bausteine blau, negative Bausteine grau (dynamisch per invertIfNegative), Summe Navy
                put_val(ser, "invertIfNegative", 1 if kind == "ertrag" else 0, so)
                if kind == "ertrag":
                    put(ser, invert_ext(K.MUTED2), so)
                for i, cname in enumerate(cats):
                    if kind in ("hh", "va"):
                        color, pat = (K.ACCENT if "objekt" in (cname or "").lower() else K.BLUE), None
                    else:
                        color, pat = match_cat(kind, cname)
                        color = color or K.BLUE
                    if color == K.BLUE and not pat:
                        continue
                    put(ser, dpt(i, color, pattern=(pat == "pattern"), bar=True, invert=0), so)
                if kind == "ertrag":
                    set_ser_dlbls(ser, dlbls_val(LBL_EUR, 9, K.NAVY, True, bg=K.WHITE), so)
                elif kind == "afa_summe":
                    set_ser_dlbls(ser, dlbls_val('#,##0" €";;"–"', 8, K.INK2, bg=K.WHITE), so)
                else:
                    set_ser_dlbls(ser, dlbls_val(LBL_EUR_POS, 8, K.INK2, bg=K.WHITE), so)
                continue
            m = match_series(name)
            color, w, dash = m if m else (FALLBACK[k % len(FALLBACK)], None, None)
            k += 1
            if ctn in ("lineChart", "line3DChart"):
                put(ser, sppr(ln=line(w or 2.0, color, dash, cap="rnd")), so)
                mk = E("c:marker")
                SE(mk, "c:symbol", "none")
                put(ser, mk, so)
                drop(ser, "dPt")
                put_val(ser, "smooth", 0, so)
                if kind == "irr" and re.search(r"basis", name.lower()):
                    n = len(cache_values(ser))
                    d = E("c:dLbls")
                    lb = SE(d, "c:dLbl")
                    SE(lb, "c:idx", max(n - 1, 0))
                    lb.append(num_fmt("0.0 %"))
                    lb.append(sppr(nofill=True, ln=line(nofill=True)))
                    lb.append(txpr(8, K.NAVY, True, wrap="none"))
                    SE(lb, "c:dLblPos", "r")
                    for nn, v in (("showLegendKey", 0), ("showVal", 1), ("showCatName", 0), ("showSerName", 0),
                                  ("showPercent", 0), ("showBubbleSize", 0)):
                        SE(lb, "c:" + nn, v)
                    for nn, v in (("showLegendKey", 0), ("showVal", 0), ("showCatName", 0), ("showSerName", 0),
                                  ("showPercent", 0), ("showBubbleSize", 0)):
                        SE(d, "c:" + nn, v)
                    set_ser_dlbls(ser, d, so)
                else:
                    drop(ser, "dLbls")
            elif ctn in ("barChart", "bar3DChart"):
                drop(ser, "dPt")
                put(ser, sppr(solid(color), line(0.75, K.WHITE) if kind == "jahr1" else line(nofill=True)), so)
                if kind == "cf_nach":
                    put_val(ser, "invertIfNegative", 1, so)
                    drop(ser, "extLst")
                    put(ser, invert_ext(K.SKY), so)
                    put(ser, sppr(solid(K.NAVY), line(nofill=True)), so)
                else:
                    put_val(ser, "invertIfNegative", 0, so)
                if kind == "jahr1":
                    set_ser_dlbls(ser, dlbls_val(LBL_EUR_POS, 8, K.WHITE if color in DARK else K.NAVY, True), so)
                else:
                    drop(ser, "dLbls")
            elif ctn == "areaChart":
                put(ser, sppr(solid(color, 10), line(nofill=True)), so)
                drop(ser, "dLbls")


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
    """Nettovermögen zusätzlich als Fläche (10 %) unter den Linien; Verkaufsjahr als schmale Säule mit Beschriftung."""
    plot = root.find(".//" + q("c:plotArea"))
    lc = plot.find(q("c:lineChart"))
    if lc is None or plot.find(q("c:areaChart")) is not None:
        return []
    axids = [a.get("val") for a in lc.findall(q("c:axId"))]
    sers = lc.findall(q("c:ser"))
    nv = next((s for s in sers if ser_name(s).lower().startswith("nettoverm")), None)
    # Hilfsreihen erhalten idx/order 0 und 1 (Legendeneinträge 0/1 werden gelöscht – eindeutig in Excel und LibreOffice)
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
        s.append(sppr(solid(K.BLUE, 12), line(nofill=True)))
        s.append(deepcopy(nv.find(q("c:cat"))))
        s.append(deepcopy(nv.find(q("c:val"))))
        s.set("helper", "1")
        area.append(s)
        hidden.append(nxt)
        nxt += 1
    for a_ in axids:
        SE(area, "c:axId", a_)
    bar = None
    if mark and nv is not None:
        n = len(cache_values(nv))
        cols = [K.L(c) for c in range(K.col("D"), K.col("D") + n)]
        vals = [mark.get(c) or 0 for c in cols]
        if any(v > 0 for v in vals):
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
            SE(SE(sc, "c:pt", idx=0), "c:v").text = "Markierung Verkaufsjahr (Diagramm)"
            s.append(sppr(solid(K.ACCENT, 55), line(nofill=True)))
            SE(s, "c:invertIfNegative", 0)
            s.append(dlbls_val('"Verkauf · "#,##0," T€";;', 8, K.BLUE, True, pos="outEnd"))
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


def pie_layout(root):
    plot = root.find(".//" + q("c:plotArea"))
    drop(plot, "layout")
    plot.insert(0, manual_layout("inner", 0.16, 0.24, 0.68, 0.62))


# ============================================================================ Hauptfunktion je Diagramm
def style_chart(xml, sheet=None, mark=None, width_px=None):
    root = etree.fromstring(xml)
    chart = root.find(q("c:chart"))
    plot = chart.find(q("c:plotArea"))
    kind = chart_kind(sheet or "", root)
    simplify_categories(root)
    sers = plot.xpath("./*/c:ser", namespaces=NS)
    n_cat = max([len(categories(s)) for s in sers] + [0])
    values = [v for s in sers for v in cache_values(s)]
    dashboard = kind == "dash"
    hidden = []
    if kind == "bestand":
        hidden = bestand_extras(root, mark)
    kind3d = make_3d(root, n_cat)
    horizontal = plot.find(".//" + q("c:barDir")) is not None and plot.find(".//" + q("c:barDir")).get("val") == "bar"

    # Titel: Schrittseiten und Dashboard tragen den Panel-Kopf in der Zelle darüber
    delete_title = dashboard or bool(re.match(r"S\d\d ", sheet or ""))
    set_title(root, TITLES.get(kind), delete_title, width_px)

    if dashboard:
        dashboard_series(root)
    else:
        style_series(root, kind, n_cat)
    for s in root.iter(q("c:ser")):
        if "helper" in s.attrib:
            del s.attrib["helper"]

    # Legende: unten; entfällt bei Kreisen und bei nur einer sichtbaren Reihe
    visible = [s for s in plot.xpath("./*/c:ser", namespaces=NS)
               if int(s.find(q("c:idx")).get("val")) not in hidden]
    if not dashboard:
        set_legend(root, kind3d != "pie" and kind not in PIE_KINDS and len(visible) > 1, hidden)
    else:
        lg = chart.find(q("c:legend"))
        if lg is not None:
            put(lg, txpr(8, K.MUTED), ORDER["legend"])

    style_axes(root, kind, n_cat, values, horizontal, dashboard)
    if kind3d == "pie" or kind in PIE_KINDS:
        pie_layout(root)

    put_val(chart, "plotVisOnly", 0, ORDER["chart"])
    put_val(chart, "dispBlanksAs", "gap", ORDER["chart"])
    # Plot- und Diagrammfläche ohne Füllung und Rahmen, Grundschrift 8 pt grau
    psp = plot.find(q("c:spPr"))
    if psp is not None:
        plot.remove(psp)
    ext = plot.find(q("c:extLst"))
    (ext.addprevious if ext is not None else plot.append)(sppr(nofill=True, ln=line(nofill=True)))
    put(root, sppr(nofill=True, ln=line(nofill=True)), ORDER["chartSpace"])
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
    widths = chart_widths(files)
    for name in sorted(files):
        if name.startswith("xl/charts/chart") and name.endswith(".xml"):
            try:
                files[name] = style_chart(files[name], charts.get(name), mark, widths.get(name))
            except Exception as exc:  # ein Diagramm darf den Build nicht abbrechen
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
