"""Prüflauf (Lint) für die fertige Mappe: hält die Stilregeln des Premium-Redesigns dauerhaft fest.

Aufruf:  python excel/lint_pro.py MAPPE.xlsx [--json BERICHT.json] [--max N] [--sheet NAME] [--skip REGEL,...]
                                             [--no-warn] [--list-rules]

Läuft im Build nach finish_pro.py/navigation.py auf der neu berechneten Mappe (zwischengespeicherte Werte werden
für die Anzeige-Prüfungen gebraucht). Ausgabe: je Verstoß eine Zeile  Blatt | Zelle | Regel | Meldung,
am Ende eine Zusammenfassung je Regel und je Blatt.
Exit-Code 1, sobald mindestens ein Verstoß der Schwere „Fehler“ vorliegt (build_pro --strict bricht dann ab);
„Warnung“ markiert Heuristiken (Schriftmetrik, Raster), die geprüft, aber nicht erzwungen werden.

Schriftmetrik: Calibri-Laufweiten über Carlito (metrisch identisch) mit Pillow bei 96 dpi; fehlt die Schrift,
wird die Näherung core.text_px verwendet.

Bewusste Ausnahmen gehören in AUSNAHMEN (Blatt, Zelle oder Bereich, Regel, Begründung) – nicht in die Module.
"""
import argparse
import json
import os
import posixpath
import re
import sys
import warnings
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings("ignore", category=UserWarning)

from openpyxl import load_workbook  # noqa: E402
from openpyxl.cell.rich_text import CellRichText, TextBlock  # noqa: E402
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries  # noqa: E402

import core  # noqa: E402

FEHLER, WARNUNG = "Fehler", "Warnung"

# =============================================================================== Regeln
RULES = {
    "einzug":          (FEHLER, "Einzug > 0 bei horizontaler Ausrichtung general/center (Excel ignoriert ihn)"),
    "schrumpfen":      (FEHLER, "„An Zellgröße anpassen“ (shrink_to_fit) gesetzt"),
    "schrift_min":     (FEHLER, "Zellschrift unter 8 pt"),
    "form_schrift_min": (FEHLER, "Formschrift unter 7 pt"),
    "diagramm_schrift": (WARNUNG, "Diagrammschrift unter 7 pt"),
    "grau_klein":      (FEHLER, "Text ≤ 9 pt in 8A9099 (Sekundärtext bis 9 pt nur in 5B6068)"),
    "schriftart":      (FEHLER, "Schriftart außerhalb {Calibri, Calibri Light} (Zellen, dxf, Formen, Diagramme)"),
    "schriftart_leer": (WARNUNG, "leere Eingabe-/Tabellenzelle mit fremder Schriftart (wird beim Tippen sichtbar)"),
    "text_ueberlauf":  (FEHLER, "Text breiter als die Zelle und Nachbarzelle belegt (abgeschnitten)"),
    "text_knapp":      (WARNUNG, "Text breiter als 90 % der Zelle bei belegter Nachbarzelle (10 % Reserve unterschritten)"),
    "zahl_raute":      (FEHLER, "Zahl breiter als die Zelle (Excel zeigt ####)"),
    "zahl_knapp":      (WARNUNG, "Zahl breiter als 90 % der Zelle (10 % Reserve unterschritten)"),
    "umbruch_hoehe":   (FEHLER, "Umbrechender Text braucht mehr Zeilen als die Zeilenhöhe zeigt"),
    "umbruch_knapp":   (WARNUNG, "Umbrechender Text füllt die Zeilenhöhe nicht ganz aus (knapp)"),
    "zeile_zu_niedrig": (FEHLER, "Schriftgröße passt nicht in die Zeilenhöhe (Text vertikal abgeschnitten)"),
    "einheit_doppelt": (FEHLER, "„€“/„%“ in einer Einheitenspalte neben einem Zahlenformat mit derselben Einheit"),
    "fehlerwert":      (FEHLER, "sichtbarer Fehlerwert (#DIV/0!, #NV, #WERT! …) in einer Zelle"),
    "diagramm_verdeckt": (FEHLER, "Diagramm liegt über belegten Zellen"),
    "diagramm_rowoff": (WARNUNG, "Diagramm beginnt mitten in einer belegten Zeile (rowOff > 0)"),
    "diagramm_rand":   (FEHLER, "Diagramm ragt über die Panel-/Druckbereichskante hinaus"),
    "fixierlinie":     (FEHLER, "Form oder Diagramm kreuzt eine Fixierlinie"),
    "druck_papier":    (FEHLER, "Papierformat nicht A4 (paperSize ≠ 9)"),
    "druck_bereich":   (FEHLER, "Druckbereich fehlt"),
    "gueltigkeit":     (FEHLER, "Datenüberprüfung ohne Fehlermeldung (showErrorMessage = 0)"),
    "kpi_format":      (FEHLER, "KPI-Zahlenformat weicht von der KPI-Spezifikation (core.KPI) ab"),
    "kpi_ampel":       (WARNUNG, "KPI mit Ampelregel ohne bedingte Formatierung"),
    "link_ziel":       (FEHLER, "Link-Ziel existiert nicht (Blatt/Zelle/Name)"),
    "link_ziel_lage":  (WARNUNG, "Link-Ziel liegt in ausgeblendeter Zeile/Spalte"),
    "zeilenraster":    (WARNUNG, "Zeilenhöhe außerhalb des Rasters (P1-17)"),
    "theme_schrift":   (WARNUNG, "Designschriftart (Theme) nicht Calibri/Calibri Light"),
}

# Bewusste Ausnahmen: (Blatt, Zelle oder Bereich oder "*", Regel) → Begründung
AUSNAHMEN = {
    ("AfA-Vergleich", "D7", "grau_klein"): "Farblegende: „grau = nicht anwendbar“ zeigt die Graustufe der inaktiven Zeilen",
    ("Finanzierung", "D7", "grau_klein"): "Farblegende: „grau = entfällt“ zeigt die Graustufe der entfallenen Werte",
}

ALLOWED_FONTS = {"Calibri", "Calibri Light"}
ERROR_VALUES = ("#DIV/0!", "#N/A", "#NV", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#WERT!", "#BEZUG!",
                "#ZAHL!", "#NAME", "#GETTING_DATA", "Err:", "#SPILL!", "#CALC!")
MUTED2 = core.MUTED2.upper()
SEPARATORS = " \u00a0·|–—/•›‹"     # reine Trennzeichen-Läufe dürfen 8A9099 sein (Rich-Text-Trenner)

# Zeilenraster (pt): Tokens aus core + Umbruchhöhen fit_row_height (n × 13 + 8) und Hinweise (n × 13 + 10)
RASTER = {core.H_BAND, core.H_BAND_B, core.H_BAND_B + 2, core.H_HEAD, core.H_ROW, core.H_ROW2, core.H_ROW3,
          core.H_STEP_ROW, core.H_STEP_ROW2, core.H_TILE_LABEL, core.H_TILE_VALUE, core.H_TILE_SUB, core.H_BUTTON,
          14, 16, 22, 30}
RASTER |= {n * 13 + 8 for n in range(1, 12)} | {n * 13 + 10 for n in range(1, 12)}
SPACER_MAX = 12            # Abstands-/Fugenzeilen bis 12 pt sind frei
CHROME_ROWS = {1, 2, 3, 4, 5, 6, 7, 8}   # Kopfleiste, Seitenkopf, Schritt-Leiste: eigene Höhen je Modul

# KPI-Anzeigen: benannte Zellen → KPI-Schlüssel aus core.KPI
KPI_NAMES = {
    "GI": ["Gesamtinvestition"],
    "EK": ["EK_Bedarf_gesamt"],
    "RATE": ["Kapitaldienst_Monat_J1"],
    "CF": ["CF_nSt_Monat_J1", "ST_CF", "LF_CFN", "S08_CF", "S12_CF"],
    "CFV": ["LF_CFV"],
    "BMR": ["Bruttomietrendite", "BMR_Cockpit", "ST_BMR", "LF_BMR", "S02_BMR"],
    "NMR": ["Nettomietrendite", "S06_NMR"],
    "DSCR": ["DSCR_J1", "ST_DSCR", "LF_DSCR", "S08_DSCR"],
    "EKR": ["EKR_Tile", "S12_EKR"],
    "IRR": ["EK_IRR", "IRR_Cockpit", "IRR_Tile", "ST_IRR", "LF_IRR", "S11_IRR", "S12_IRR"],
    "FAKTOR": ["Kaufpreisfaktor"],
    "MULT": ["EK_Multiple"],
}
# KPI-Quellzellen in Rechenblättern: Format nur als Warnung (Rechenblatt darf genauer zeigen)
KPI_CALC_SHEETS = {"Eingaben", "Steuern", "Projektion", "Finanzierung", "Konfiguration"}

NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}
EMU = 9525
R_ID = "{%s}id" % NS["r"]


# =============================================================================== Schriftmetrik
class Metric:
    """Laufweite in px bei 96 dpi (Excel 100 %)."""
    DIRS = ["/usr/share/fonts/truetype/crosextra", "/usr/share/fonts/truetype/carlito", "/usr/share/fonts"]
    FILES = {(False, False): "Carlito-Regular.ttf", (True, False): "Carlito-Bold.ttf",
             (False, True): "Carlito-Italic.ttf", (True, True): "Carlito-BoldItalic.ttf"}

    def __init__(self):
        self.cache = {}
        self.paths = {}
        try:
            from PIL import ImageFont  # noqa: F401
            self.ok = True
        except Exception:
            self.ok = False
        for k, fn in self.FILES.items():
            for d in self.DIRS:
                for root, _, files in os.walk(d):
                    if fn in files:
                        self.paths[k] = os.path.join(root, fn)
                        break
                if k in self.paths:
                    break
        self.ok = self.ok and (False, False) in self.paths
        self.source = "Carlito (Calibri-metrisch)" if self.ok else "Näherung core.text_px"

    def font(self, size, bold, italic):
        key = (round(size * 4) / 4, bool(bold), bool(italic))
        if key not in self.cache:
            from PIL import ImageFont
            path = self.paths.get((key[1], key[2])) or self.paths.get((key[1], False)) or self.paths[(False, False)]
            self.cache[key] = ImageFont.truetype(path, max(1, int(round(size * 96 / 72 * 4)))) if path else None
        return self.cache[key]

    def width(self, text, size=10, bold=False, italic=False):
        if text is None or text == "":
            return 0.0
        s = str(text)
        if not self.ok:
            return core.text_px(s, size, bold)
        f = self.font(size, bold, italic)
        return f.getlength(s) / 4.0


METRIC = Metric()


# =============================================================================== Zahlenformat → Anzeigetext
def _split_sections(fmt):
    out, cur, q = [], "", False
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            q = not q
        if ch == "\\" and i + 1 < len(fmt):
            cur += fmt[i:i + 2]
            i += 2
            continue
        if ch == ";" and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
        i += 1
    out.append(cur)
    return out


def _cond_ok(cond, v):
    m = re.match(r"\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]", cond)
    if not m:
        return None
    op, n = m.group(1), float(m.group(2))
    return {"<": v < n, ">": v > n, "=": v == n, "<=": v <= n, ">=": v >= n, "<>": v != n}[op]


def format_number(v, fmt):
    """Anzeigetext einer Zahl in einem Excel-Format (deutsches Excel: gleiche Laufweite wie en-US)."""
    if isinstance(v, bool):
        return "WAHR" if v else "FALSCH"
    fmt = fmt or "General"
    if fmt in ("General", "Standard", "@"):
        if float(v).is_integer() and abs(v) < 1e11:
            return str(int(v))
        return f"{v:.10g}"
    secs = _split_sections(fmt)
    sec, neg_sign = None, v < 0
    conds = [s for s in secs if re.match(r"\s*\[(<|>|=|<=|>=|<>)-?[\d.]+\]", s)]
    if conds:
        for s in secs:
            ok = _cond_ok(re.match(r"\s*(\[[^\]]*\])?", s).group(1) or "", v)
            if ok is None or ok:
                sec = s
                neg_sign = False
                break
        sec = sec if sec is not None else secs[-1]
    elif v > 0 or len(secs) == 1:
        sec = secs[0]
    elif v < 0:
        sec = secs[1] if len(secs) > 1 and secs[1] else secs[0]
        neg_sign = not (len(secs) > 1 and secs[1])
    else:
        sec = secs[2] if len(secs) > 2 else secs[0]
    sec = re.sub(r"\[(Red|Black|Blue|Green|White|Yellow|Magenta|Cyan|Rot|Schwarz|Blau|Grün|Color\s*\d+|[<>=]+-?[\d.]+)\]", "",
                 sec, flags=re.I)
    sec = re.sub(r"\[\$([^\]-]*)(-[0-9A-Fa-f]+)?\]", lambda m: '"' + m.group(1) + '"', sec)
    # Datum/Zeit
    bare = re.sub(r'"[^"]*"|\\.', "", sec)
    if re.search(r"[DdMmYyJjTtHhSs]", bare) and not re.search(r"[0#?]", bare):
        return "31.12.2026" if re.search(r"[Yy]{2,}|[Jj]{2,}", bare) else "31.12."
    # Literale und Platzhalter trennen
    parts = []
    i = 0
    while i < len(sec):
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = len(sec) if j < 0 else j
            parts.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\" and i + 1 < len(sec):
            parts.append(("lit", sec[i + 1]))
            i += 2
        elif ch == "_" and i + 1 < len(sec):
            parts.append(("lit", " "))
            i += 2
        elif ch == "*" and i + 1 < len(sec):
            i += 2
        elif ch in "0#?.,":
            parts.append(("num", ch))
            i += 1
        elif ch == "%":
            parts.append(("pct", "%"))
            i += 1
        else:
            parts.append(("lit", ch))
            i += 1
    num = "".join(p[1] for p in parts if p[0] == "num")
    pct = sum(1 for p in parts if p[0] == "pct")
    x = abs(v) * (100 ** pct)
    if not num:
        body = ""
    else:
        dec = num.split(".", 1)[1] if "." in num else ""
        ndec = sum(1 for ch in dec if ch in "0#?")
        int_part = num.split(".", 1)[0]
        scale = len(int_part) - len(int_part.rstrip(","))
        x = x / (1000 ** scale)
        thousands = "," in int_part.rstrip(",")
        body = f"{x:,.{ndec}f}" if thousands else f"{x:.{ndec}f}"
        if not re.search(r"0", int_part) and body.startswith("0") and ndec:
            body = body[1:]
    out, placed = "", False
    for kind, t in parts:
        if kind == "num":
            if not placed:
                out += body
                placed = True
        else:
            out += t
    return ("-" if neg_sign and v != 0 else "") + out


def fmt_class(fmt):
    """Semantische Klasse eines Zahlenformats: (art, dezimalstellen) – art ∈ pct, eur, mult, num, text, other."""
    if not fmt:
        return ("other", None)
    sec = _split_sections(fmt)[0]
    lits = "".join(re.findall(r'"([^"]*)"', sec)) + "".join(re.findall(r"\\(.)", sec))
    bare = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", sec)
    if fmt.strip() == "@":
        return ("text", None)
    dec = bare.split(".", 1)[1] if "." in bare else ""
    nd = sum(1 for ch in dec if ch in "0#?")
    if "%" in bare:
        return ("pct", nd)
    if "€" in lits or "€" in bare or "EUR" in lits:
        return ("eur", nd)
    if "×" in lits or "x" == lits.strip():
        return ("mult", nd)
    if re.search(r"[0#]", bare):
        return ("num", nd)
    return ("other", None)


# =============================================================================== Hilfen
def rgb_of(color):
    try:
        if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
            return None
        return color.rgb[-6:].upper()
    except Exception:
        return None


def is_empty(v):
    if v is None:
        return True
    if isinstance(v, CellRichText):
        v = str(v)
    return isinstance(v, str) and v.strip() == ""


def a1(c, r):
    return f"{get_column_letter(c)}{r}"


def in_ref(ref, c, r):
    if ref == "*":
        return True
    try:
        c1, r1, c2, r2 = range_boundaries(ref)
    except Exception:
        return False
    return c1 <= c <= (c2 or c1) and r1 <= r <= (r2 or r1)


def excepted(sheet, ref, rule):
    """Befund durch AUSNAHMEN gedeckt? Schlüssel: Blatt, Zelle/Bereich/„*“, Regel."""
    first = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)", ref or "")
    for (s, r, ru) in AUSNAHMEN:
        if s != sheet or ru != rule:
            continue
        if r == "*" or r == ref:
            return True
        if first and in_ref(r, column_index_from_string(first.group(1)), int(first.group(2))):
            return True
    return False


class Geo:
    """Pixelgeometrie eines Blatts (Spalten/Zeilen, ausgeblendete = 0 px)."""

    def __init__(self, ws):
        self.ws = ws
        self.default_h = (ws.sheet_format.defaultRowHeight or 15)
        self._cx, self._ry = [0.0], [0.0]
        self.max_c = max(ws.max_column, 60)
        self.max_r = max(ws.max_row, 200)
        for c in range(1, self.max_c + 2):
            self._cx.append(self._cx[-1] + self.col_px(c))
        for r in range(1, self.max_r + 2):
            self._ry.append(self._ry[-1] + self.row_px(r))

    def col_hidden(self, c):
        d = self.ws.column_dimensions.get(get_column_letter(c))
        if d is not None and d.hidden:
            return True
        # openpyxl fasst gleichartige Spalten zu Gruppen (min..max) zusammen
        for d in self.ws.column_dimensions.values():
            if d.min and d.max and d.min <= c <= d.max and d.hidden:
                return True
        return False

    def row_hidden(self, r):
        d = self.ws.row_dimensions.get(r)
        return bool(d is not None and d.hidden)

    def col_px(self, c):
        if self.col_hidden(c):
            return 0
        d = self.ws.column_dimensions.get(get_column_letter(c))
        if d is None or not d.width:
            for dd in self.ws.column_dimensions.values():
                if dd.min and dd.max and dd.min <= c <= dd.max and dd.width:
                    d = dd
                    break
        return core.col_px(self.ws, c) if d is None or not d.width else int(((256 * d.width + int(128 / 7)) / 256) * 7)

    def row_pt(self, r):
        if self.row_hidden(r):
            return 0
        d = self.ws.row_dimensions.get(r)
        return d.height if (d is not None and d.height is not None) else self.default_h

    def row_px(self, r):
        return self.row_pt(r) * 96 / 72

    def x(self, c):      # linke Kante der Spalte c (1-basiert)
        c = max(1, c)
        while c >= len(self._cx):
            self._cx.append(self._cx[-1] + self.col_px(len(self._cx)))
        return self._cx[c - 1]

    def y(self, r):
        r = max(1, r)
        while r >= len(self._ry):
            self._ry.append(self._ry[-1] + self.row_px(len(self._ry)))
        return self._ry[r - 1]

    def col_at(self, px):
        c = 1
        while self.x(c + 1) <= px and c < 16384:
            c += 1
        return c

    def row_at(self, px):
        r = 1
        while self.y(r + 1) <= px and r < 1048576:
            r += 1
        return r


# =============================================================================== Paket (XML-Teile)
class Package:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())
        wbx = ET.fromstring(self.z.read("xl/workbook.xml"))
        rels = self.rels("xl/workbook.xml")
        self.sheet_part = {}
        for s in wbx.iter("{%s}sheet" % NS["m"]):
            rid = s.get(R_ID)
            self.sheet_part[s.get("name")] = rels.get(rid)

    def read(self, part):
        return self.z.read(part) if part in self.names else None

    def rels(self, part):
        d, b = posixpath.split(part)
        rp = posixpath.join(d, "_rels", b + ".rels")
        out = {}
        data = self.read(rp)
        if data is None:
            return out
        for rel in ET.fromstring(data):
            tgt = rel.get("Target")
            if rel.get("TargetMode") == "External" or tgt.startswith("#"):
                out[rel.get("Id")] = tgt
            else:
                out[rel.get("Id")] = posixpath.normpath(posixpath.join(d, tgt)) if not tgt.startswith("/") else tgt[1:]
        return out

    def drawing_of(self, sheet):
        part = self.sheet_part.get(sheet)
        if not part:
            return None
        for tgt in self.rels(part).values():
            if "/drawings/" in tgt and tgt.endswith(".xml") and "vmlDrawing" not in tgt:
                return tgt
        return None


def anchor_rect(el, geo):
    """(x0, y0, x1, y1) in px für twoCellAnchor / oneCellAnchor / absoluteAnchor."""
    def pt(node):
        # Versätze auf die Zellgröße begrenzen (Excel klemmt; LibreOffice schreibt mit eigener Spaltenmetrik
        # gelegentlich colOff > Spaltenbreite)
        c = int(node.find("xdr:col", NS).text)
        co = int(node.find("xdr:colOff", NS).text)
        r = int(node.find("xdr:row", NS).text)
        ro = int(node.find("xdr:rowOff", NS).text)
        dx = min(co / EMU, geo.col_px(c + 1))
        dy = min(ro / EMU, geo.row_px(r + 1))
        return geo.x(c + 1) + dx, geo.y(r + 1) + dy, (c, co, r, ro)

    tag = el.tag.split("}")[1]
    if tag == "twoCellAnchor":
        x0, y0, f = pt(el.find("xdr:from", NS))
        x1, y1, t = pt(el.find("xdr:to", NS))
        return (x0, y0, x1, y1), f, t
    if tag == "oneCellAnchor":
        x0, y0, f = pt(el.find("xdr:from", NS))
        ext = el.find("xdr:ext", NS)
        return (x0, y0, x0 + int(ext.get("cx")) / EMU, y0 + int(ext.get("cy")) / EMU), f, None
    pos, ext = el.find("xdr:pos", NS), el.find("xdr:ext", NS)
    x0, y0 = int(pos.get("x")) / EMU, int(pos.get("y")) / EMU
    return (x0, y0, x0 + int(ext.get("cx")) / EMU, y0 + int(ext.get("cy")) / EMU), None, None


# =============================================================================== Prüfer
class Linter:
    def __init__(self, path, only_sheet=None, skip=()):
        self.path = path
        self.wb = load_workbook(path, rich_text=True)
        self.wbv = load_workbook(path, data_only=True)
        self.pkg = Package(path)
        self.only = only_sheet
        self.skip = set(skip)
        self.found = []
        self.geo = {}
        self.merged = {}
        self.names = {}
        for n, d in self.wb.defined_names.items():
            self.names[n] = d.attr_text
        for ws in self.wb.worksheets:
            for n, d in ws.defined_names.items():
                self.names.setdefault(n, d.attr_text)

    # ------------------------------------------------------------------ Ausgabe
    def add(self, sheet, ref, rule, msg):
        if rule in self.skip or excepted(sheet, ref, rule):
            return
        self.found.append(dict(sheet=sheet, ref=ref, rule=rule, severity=RULES[rule][0], msg=msg))

    # ------------------------------------------------------------------ Blätter
    def sheets(self):
        for ws in self.wb.worksheets:
            if self.only and ws.title != self.only:
                continue
            yield ws, self.wbv[ws.title]

    def merges(self, ws):
        if ws.title not in self.merged:
            anchors, inner = {}, set()
            for mr in ws.merged_cells.ranges:
                anchors[(mr.min_col, mr.min_row)] = (mr.max_col, mr.max_row)
                for r in range(mr.min_row, mr.max_row + 1):
                    for c in range(mr.min_col, mr.max_col + 1):
                        if (c, r) != (mr.min_col, mr.min_row):
                            inner.add((c, r))
            self.merged[ws.title] = (anchors, inner)
        return self.merged[ws.title]

    def print_box(self, ws):
        """(c1, r1, c2, r2) des Druckbereichs oder None."""
        key = ("pa", ws.title)
        if key not in self.geo:
            box = None
            if ws.print_area:
                try:
                    box = range_boundaries(str(ws.print_area).split(",")[0].split("!")[-1].replace("$", ""))
                except Exception:
                    box = None
            self.geo[key] = box
        return self.geo[key]

    def in_print(self, ws, c, r):
        """Liegt die Zelle im Druckbereich? (c=1 prüft nur die Zeile)"""
        box = self.print_box(ws)
        if box is None:
            return True
        return box[1] <= r <= box[3] and (c == 1 or box[0] <= c <= box[2])

    def g(self, ws):
        if ws.title not in self.geo:
            self.geo[ws.title] = Geo(ws)
        return self.geo[ws.title]

    def run(self):
        for ws, wv in self.sheets():
            self.check_cells(ws, wv)
            self.check_rows(ws, wv)
            self.check_print(ws)
            self.check_validation(ws)
            self.check_links(ws)
            self.check_drawing(ws, wv)
        if not self.only:
            self.check_dxf()
            self.check_theme()
        self.check_kpi()
        return self.found

    # ------------------------------------------------------------------ Zellen
    def display_text(self, cell, vcell):
        """(Anzeigetext, Art, Läufe) – Läufe: [(text, pt, fett, kursiv, schrift, farbe)] (Rich Text je Lauf)."""
        f = cell.font
        base = (float(f.sz or 11), bool(f.b), bool(f.i), f.name, rgb_of(f.color))
        v = cell.value
        if cell.data_type == "f":
            v = vcell.value
        if v is None:
            return None, None, []
        if isinstance(v, CellRichText):
            runs = []
            for blk in v:
                if isinstance(blk, TextBlock):
                    ft = blk.font
                    runs.append((blk.text, float(ft.sz or base[0]), bool(ft.b) if ft.b is not None else base[1],
                                 bool(ft.i) if ft.i is not None else base[2], ft.rFont or base[3],
                                 rgb_of(ft.color) if ft.color is not None else base[4]))
                else:
                    runs.append((str(blk),) + base)
            return str(v), "text", runs
        if isinstance(v, str):
            return v, "text", [(v,) + base]
        if isinstance(v, (int, float)):
            try:
                t = format_number(float(v), cell.number_format)
            except Exception:
                t = str(v)
            return t, "num", [(t,) + base]
        t = "31.12.2026"   # Datum/Zeit
        return t, "num", [(t,) + base]

    def check_cells(self, ws, wv):
        geo = self.g(ws)
        anchors, inner = self.merges(ws)
        title = ws.title
        occupied = set()
        for row in ws.iter_rows():
            for c in row:
                if not is_empty(c.value):
                    occupied.add((c.column, c.row))
        for row in ws.iter_rows():
            for c in row:
                col, r = c.column, c.row
                if (col, r) in inner:
                    continue
                hidden = geo.row_hidden(r) or geo.col_hidden(col)
                f = c.font
                name = f.name if f is not None else None
                ref = c.coordinate
                if is_empty(c.value):
                    if name and name not in ALLOWED_FONTS and not hidden and self.in_print(ws, col, r) and (
                            c.fill is not None and c.fill.fill_type == "solid" or c.protection.locked is False):
                        self.add(title, ref, "schriftart_leer", f"„{name}“")
                    continue
                if hidden:
                    continue
                al = c.alignment
                h = al.horizontal or "general"
                if al.indent and h not in ("left", "right", "distributed"):
                    self.add(title, ref, "einzug", f"indent={al.indent:g}, horizontal={h}")
                if al.shrink_to_fit:
                    self.add(title, ref, "schrumpfen", "shrink_to_fit=1")
                text, kind, runs = self.display_text(c, wv[ref])
                if text is None or text == "":
                    continue
                fonts = {ru[4] for ru in runs if ru[0].strip() and ru[4]}
                for fn in sorted(fonts - ALLOWED_FONTS):
                    self.add(title, ref, "schriftart", f"„{fn}“")
                sizes = [ru[1] for ru in runs if ru[0].strip()]
                if sizes and min(sizes) < 8:
                    self.add(title, ref, "schrift_min", f"{min(sizes):g} pt")
                small_grey = [ru[1] for ru in runs if ru[0].strip(SEPARATORS) and ru[5] == MUTED2 and ru[1] <= 9]
                if small_grey:
                    self.add(title, ref, "grau_klein", f"{max(small_grey):g} pt in 8A9099")
                if kind == "text" and any(text.startswith(e) for e in ERROR_VALUES):
                    self.add(title, ref, "fehlerwert", text)
                    continue
                self.check_fit(ws, geo, c, text, kind, runs, h, al, anchors, occupied)
                self.check_unit(ws, wv, c, text)

    def check_fit(self, ws, geo, c, text, kind, runs, h, al, anchors, occupied):
        col, r = c.column, c.row
        c2, r2 = anchors.get((col, r), (col, r))
        avail = sum(geo.col_px(k) for k in range(col, c2 + 1))
        height_pt = sum(geo.row_pt(k) for k in range(r, r2 + 1))
        if avail <= 0 or height_pt <= 0:
            return
        indent_px = (al.indent or 0) * 9
        pad = 5 + indent_px
        wrap = bool(al.wrap_text) and kind == "text"
        paras = split_paragraphs(runs)
        if wrap:
            lines = []
            for para in paras:
                lines += wrap_runs(para, avail - pad - 2)
            need = sum(line_height(sz) for sz in lines)
            lh = line_height(max(lines))
            over = need - height_pt
            if over > 0.3 * lh:
                self.add(ws.title, c.coordinate, "umbruch_hoehe",
                         f"{len(lines)} Zeilen ≈ {need:.0f} pt, Höhe {height_pt:g} pt: „{short(text)}“")
            elif over > 1.5:
                self.add(ws.title, c.coordinate, "umbruch_knapp",
                         f"{len(lines)} Zeilen ≈ {need:.0f} pt, Höhe {height_pt:g} pt: „{short(text)}“")
            return
        # ohne Umbruch: Excel zeigt nur die erste Zeile vollständig; vertikal muss die größte Schrift passen
        first = paras[0] if paras else []
        sz = max([ru[1] for ru in first if ru[0].strip()] or [11])
        glyph = sz * 1.2
        if glyph > height_pt * 1.15:
            self.add(ws.title, c.coordinate, "zeile_zu_niedrig",
                     f"{sz:g} pt Schrift in {height_pt:g} pt Zeile: „{short(text)}“")
        w = sum(METRIC.width(ru[0], ru[1], ru[2], ru[3]) for ru in first) + pad
        shown = "".join(ru[0] for ru in first)
        if kind == "num":
            if w > avail:
                self.add(ws.title, c.coordinate, "zahl_raute",
                         f"„{shown}“ braucht {w:.0f} px, Zelle {avail} px")
            elif w > avail * 0.9 and w > avail - 8:
                self.add(ws.title, c.coordinate, "zahl_knapp", f"„{shown}“ {w:.0f}/{avail} px")
            return
        if w <= avail * 0.9:
            return
        # Überlauf-Richtung: links/general → rechts, rechts → links, zentriert → beide Seiten
        blocked, extra = [], w - avail
        dirs = {"right": [-1], "center": [1, -1], "centerContinuous": [1, -1]}.get(h, [1])
        if h == "fill" or h == "justify":
            dirs = [1]
        need_each = extra / len(dirs)
        _, inner = self.merges(ws)
        in_merge = (col, r) in anchors
        for d in dirs:
            got, k = 0, (c2 + 1 if d > 0 else col - 1)
            if in_merge:      # verbundene Zellen laufen in Excel nie über
                blocked.append(f"Verbund {a1(col, r)}:{a1(c2, r2)}")
                break
            while got < need_each and 1 <= k <= 16384:
                if (k, r) in occupied or (k, r) in inner or (k, r) in anchors:
                    blocked.append(a1(k, r))
                    break
                got += geo.col_px(k)
                k += d
        if not blocked:
            return
        if w > avail * 1.02:     # 2 % Toleranz für Innenabstand/Einzugsmetrik
            self.add(ws.title, c.coordinate, "text_ueberlauf",
                     f"{w:.0f} px Text in {avail} px, {blocked[0]} begrenzt: „{short(shown)}“")
        else:
            self.add(ws.title, c.coordinate, "text_knapp",
                     f"{w:.0f} px Text in {avail} px (< 10 % Reserve), {blocked[0]} begrenzt: „{short(shown)}“")

    UNITS = {"€": "€", "EUR": "€", "%": "%", "€/m²": "€/m²", "m²": "m²", "Jahre": "Jahre", "×": "×"}

    def check_unit(self, ws, wv, c, text):
        t = text.strip()
        unit = self.UNITS.get(t)
        if unit is None:
            return
        for k in (1, 2):
            if c.column - k < 1:
                break
            n = ws.cell(c.row, c.column - k)
            nv = wv.cell(c.row, c.column - k).value if n.data_type == "f" else n.value
            if is_empty(nv):
                continue
            if isinstance(nv, (int, float)) and not isinstance(nv, bool):
                fmt = n.number_format or ""
                lits = "".join(re.findall(r'"([^"]*)"', fmt))
                has = ("%" in re.sub(r'"[^"]*"', "", fmt) or "%" in lits) if unit == "%" else (unit in lits or unit in fmt)
                if has:
                    self.add(ws.title, c.coordinate, "einheit_doppelt",
                             f"„{t}“ neben {n.coordinate} (Format {fmt})")
            break

    # ------------------------------------------------------------------ Zeilenraster
    def check_rows(self, ws, wv):
        geo = self.g(ws)
        rows_with_content = set()
        for row in ws.iter_rows():
            for c in row:
                if not is_empty(c.value):
                    rows_with_content.add(c.row)
        anchors, _ = self.merges(ws)
        multi = {r for (c, r), (c2, r2) in anchors.items() if r2 > r}
        for r in sorted(rows_with_content):
            if r in CHROME_ROWS or geo.row_hidden(r) or r in multi or not self.in_print(ws, 1, r):
                continue
            d = ws.row_dimensions.get(r)
            h = d.height if d is not None else None
            if h is None or h <= SPACER_MAX:
                continue
            # LibreOffice rastet Höhen beim Speichern auf ganze Pixel ab (20 → 19,5 pt, 22 → 21,75 pt)
            if not any(-0.3 <= x - h <= 0.76 for x in RASTER):
                self.add(ws.title, f"{r}:{r}", "zeilenraster", f"Zeile {r}: {h:g} pt")

    # ------------------------------------------------------------------ Druck, Gültigkeit, Links
    def check_print(self, ws):
        if ws.sheet_state != "visible":
            return
        ps = ws.page_setup.paperSize
        if str(ps) != "9":
            self.add(ws.title, "-", "druck_papier", f"paperSize={ps}")
        if not ws.print_area:
            self.add(ws.title, "-", "druck_bereich", "kein _xlnm.Print_Area")

    def check_validation(self, ws):
        for dv in ws.data_validations.dataValidation:
            if (dv.type or "any") == "any":
                continue
            if not dv.showErrorMessage:
                self.add(ws.title, str(dv.sqref), "gueltigkeit", f"type={dv.type}, showErrorMessage=0")

    def resolve(self, loc):
        """Link-Ort → (blatt, spalte, zeile) oder Fehlertext."""
        loc = loc.lstrip("#").strip()
        m = re.match(r"^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::.*)?$", loc)
        if m:
            sh = (m.group(1) or m.group(2)).replace("''", "'")
            if sh not in self.wb.sheetnames:
                return f"Blatt „{sh}“ fehlt"
            return sh, column_index_from_string(m.group(3)), int(m.group(4))
        if loc in self.names:
            return self.resolve(self.names[loc])
        return f"„{loc}“ ist weder Zelle noch Name"

    def check_target(self, sheet, ref, loc, what):
        res = self.resolve(loc)
        if isinstance(res, str):
            self.add(sheet, ref, "link_ziel", f"{what} → {loc}: {res}")
            return
        sh, c, r = res
        geo = self.g(self.wb[sh])
        if geo.row_hidden(r) or geo.col_hidden(c):
            self.add(sheet, ref, "link_ziel_lage", f"{what} → {loc} liegt ausgeblendet")

    def check_links(self, ws):
        for row in ws.iter_rows():
            for c in row:
                hl = c.hyperlink
                if hl is None:
                    continue
                if hl.location:
                    self.check_target(ws.title, c.coordinate, hl.location, "Zell-Link")
                elif hl.target and str(hl.target).startswith("#"):
                    self.check_target(ws.title, c.coordinate, hl.target, "Zell-Link")

    # ------------------------------------------------------------------ Zeichnungen (Formen, Diagramme)
    def check_drawing(self, ws, wv):
        part = self.pkg.drawing_of(ws.title)
        if not part:
            return
        data = self.pkg.read(part)
        if not data:
            return
        root = ET.fromstring(data)
        rels = self.pkg.rels(part)
        geo = self.g(ws)
        title = ws.title
        # Fixierlinien
        xs = ys = 0
        if ws.freeze_panes:
            fc, fr, _, _ = range_boundaries(ws.freeze_panes + ":" + ws.freeze_panes)
            xs, ys = fc - 1, fr - 1
        bx, by = geo.x(xs + 1), geo.y(ys + 1)
        pa = self.print_box(ws)
        for el in root:
            tag = el.tag.split("}")[1]
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            (x0, y0, x1, y1), fr_, to_ = anchor_rect(el, geo)
            cnv = el.find(".//xdr:cNvPr", NS)
            name = cnv.get("name") if cnv is not None else "?"
            chart = el.find(".//c:chart", NS)
            kind = "Diagramm" if chart is not None else "Form"
            label = f"{kind} „{name}“"
            ref = f"{a1(geo.col_at(x0 + 1), geo.row_at(y0 + 1))}:{a1(geo.col_at(max(x0, x1 - 1)), geo.row_at(max(y0, y1 - 1)))}"
            if xs and x0 < bx - 1 and x1 > bx + 1:
                self.add(title, ref, "fixierlinie", f"{label} kreuzt die senkrechte Fixierlinie ({get_column_letter(xs)}|{get_column_letter(xs + 1)})")
            if ys and y0 < by - 1 and y1 > by + 1:
                self.add(title, ref, "fixierlinie", f"{label} kreuzt die waagerechte Fixierlinie (Zeile {ys}|{ys + 1})")
            # Formschriften (Text in Formen; +mn-lt/+mj-lt = Designschrift)
            for rpr in el.iter():
                t = rpr.tag.split("}")[1]
                if t in ("rPr", "defRPr", "endParaRPr"):
                    sz = rpr.get("sz")
                    if sz and int(sz) < 700 and chart is None:
                        self.add(title, ref, "form_schrift_min", f"{label}: {int(sz) / 100:g} pt")
                if t in ("latin", "ea", "cs") and t == "latin":
                    tf = rpr.get("typeface")
                    if tf and not tf.startswith("+") and tf not in ALLOWED_FONTS:
                        self.add(title, ref, "schriftart", f"{label}: „{tf}“")
            # Form-Links
            for hl in el.iter("{%s}hlinkClick" % NS["a"]):
                tgt = rels.get(hl.get(R_ID))
                if tgt and tgt.startswith("#"):
                    self.check_target(title, ref, tgt, f"Form-Link „{name}“")
                elif tgt is None and hl.get(R_ID):
                    self.add(title, ref, "link_ziel", f"Form-Link „{name}“: Beziehung {hl.get(R_ID)} fehlt")
            if chart is None:
                continue
            self.check_chart_part(title, ref, name, rels.get(chart.get(R_ID)))
            # Diagramm über belegten Zellen
            anchors, inner = self.merges(ws)
            covered = []
            c_lo, c_hi = geo.col_at(x0 + 3), geo.col_at(max(x0, x1 - 3))
            r_lo, r_hi = geo.row_at(y0 + 3), geo.row_at(max(y0, y1 - 3))
            for r in range(r_lo, r_hi + 1):
                if geo.row_hidden(r):
                    continue
                for cc in range(c_lo, c_hi + 1):
                    if geo.col_hidden(cc) or (cc, r) in inner:
                        continue
                    cell = ws.cell(r, cc)
                    if is_empty(cell.value):
                        continue
                    shown = cell.value
                    if cell.data_type == "f":
                        shown = wv.cell(r, cc).value
                    if is_empty(shown):
                        continue
                    covered.append(cell.coordinate)
            if covered:
                more = f" … (+{len(covered) - 4})" if len(covered) > 4 else ""
                self.add(title, ref, "diagramm_verdeckt", f"„{name}“ verdeckt {', '.join(covered[:4])}{more}")
            if fr_ is not None and fr_[3] > 0:
                r0 = fr_[2] + 1
                cols = range(geo.col_at(x0 + 1), geo.col_at(max(x0, x1 - 1)) + 1)
                if any(not is_empty(ws.cell(r0, k).value) for k in cols) and not covered:
                    self.add(title, ref, "diagramm_rowoff", f"„{name}“ beginnt {fr_[3] / EMU:.0f} px in Zeile {r0}")
            if pa is not None:
                pc1, pr1, pc2, pr2 = pa
                if x1 > geo.x(pc2 + 1) + 2 or y1 > geo.y(pr2 + 1) + 2 or x0 < geo.x(pc1) - 2:
                    self.add(title, ref, "diagramm_rand",
                             f"„{name}“ endet bei {x1:.0f}/{y1:.0f} px, Druckbereich bis {geo.x(pc2 + 1):.0f}/{geo.y(pr2 + 1):.0f} px")

    def check_chart_part(self, sheet, ref, name, part):
        if not part:
            self.add(sheet, ref, "link_ziel", f"Diagramm „{name}“: Diagrammteil fehlt")
            return
        data = self.pkg.read(part)
        if data is None:
            self.add(sheet, ref, "link_ziel", f"Diagramm „{name}“: {part} fehlt")
            return
        root = ET.fromstring(data)
        bad_fonts, small = set(), set()
        for el in root.iter():
            t = el.tag.split("}")[1]
            if t == "latin":
                tf = el.get("typeface")
                if tf and not tf.startswith("+") and tf not in ALLOWED_FONTS:
                    bad_fonts.add(tf)
            if t in ("rPr", "defRPr") and el.get("sz") and int(el.get("sz")) < 700:
                small.add(int(el.get("sz")) / 100)
        for tf in sorted(bad_fonts):
            self.add(sheet, ref, "schriftart", f"Diagramm „{name}“: „{tf}“")
        if small:
            self.add(sheet, ref, "diagramm_schrift", f"Diagramm „{name}“: {', '.join(f'{s:g}' for s in sorted(small))} pt")

    # ------------------------------------------------------------------ dxf, Theme
    def check_dxf(self):
        for i, dxf in enumerate(getattr(self.wb, "_differential_styles", None).dxf if getattr(self.wb, "_differential_styles", None) else []):
            f = dxf.font
            if f is not None and f.name and f.name not in ALLOWED_FONTS:
                self.add("(Mappe)", f"dxf {i}", "schriftart", f"bedingte Formatierung: „{f.name}“")
            if f is not None and f.sz and float(f.sz) < 8:
                self.add("(Mappe)", f"dxf {i}", "schrift_min", f"bedingte Formatierung: {float(f.sz):g} pt")

    def check_theme(self):
        data = self.pkg.read("xl/theme/theme1.xml")
        if not data:
            return
        root = ET.fromstring(data)
        for kind in ("majorFont", "minorFont"):
            el = root.find(f".//a:{kind}/a:latin", NS)
            if el is not None and el.get("typeface") not in ALLOWED_FONTS:
                self.add("(Mappe)", kind, "theme_schrift", f"{kind}: „{el.get('typeface')}“")

    # ------------------------------------------------------------------ KPI-Spezifikation
    def cf_ranges(self, ws):
        out = []
        for rng_ in ws.conditional_formatting:
            for cr in rng_.sqref.ranges:
                out.append((cr.min_col, cr.min_row, cr.max_col, cr.max_row))
        return out

    def check_kpi(self):
        targets = {}   # (blatt, spalte, zeile) → KPI
        for key, names in KPI_NAMES.items():
            for n in names:
                res = self.resolve(self.names[n]) if n in self.names else None
                if isinstance(res, tuple):
                    targets[res] = (key, n)
        # Anzeigezellen mit einfacher Formel =NAME oder =Blatt!Zelle auf eine KPI-Quelle
        name_of = {n: k for k, ns in KPI_NAMES.items() for n in ns}
        for ws in self.wb.worksheets:
            if self.only and ws.title != self.only:
                continue
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if c.data_type != "f" or not isinstance(v, str):
                        continue
                    body = v[1:].strip()
                    if body in name_of:
                        targets.setdefault((ws.title, c.column, c.row), (name_of[body], "=" + body))
                        continue
                    res = self.resolve(body) if "!" in body and re.match(r"^('[^']+'|[^!(),+\-*/&]+)!\$?[A-Z]+\$?\d+$", body) else None
                    if isinstance(res, tuple) and res in targets:
                        targets.setdefault((ws.title, c.column, c.row), (targets[res][0], "=" + body))
        cf_cache = {}
        for (sh, c, r), (key, via) in sorted(targets.items()):
            if self.only and sh != self.only:
                continue
            ws = self.wb[sh]
            anchors, inner = self.merges(ws)
            if (c, r) in inner:
                continue
            geo = self.g(ws)
            if geo.row_hidden(r) or geo.col_hidden(c):
                continue
            cell = ws.cell(r, c)
            spec = core.KPI[key]
            want, have = fmt_class(spec["fmt"]), fmt_class(cell.number_format)
            if have != want:
                rule = "kpi_format"
                if sh in KPI_CALC_SHEETS:
                    continue   # Rechenblätter dürfen genauer zeigen
                self.add(sh, cell.coordinate, rule,
                         f"{spec['label']} ({via}): Format {cell.number_format!r} statt {spec['fmt']!r}")
            if spec.get("rule") in ("ampel", "cf") and sh not in KPI_CALC_SHEETS:
                if sh not in cf_cache:
                    cf_cache[sh] = self.cf_ranges(ws)
                if not any(a <= c <= b and rr1 <= r <= rr2 for a, rr1, b, rr2 in cf_cache[sh]):
                    self.add(sh, cell.coordinate, "kpi_ampel", f"{spec['label']} ({via}) ohne Ampel")


# =============================================================================== Text-Hilfen
def line_height(sz):
    """Zeilenabstand in pt für Calibri (Excel: 11 pt → 15 pt, 10 pt → 12,75 pt, 9 pt → 12 pt)."""
    return max(sz * 1.25, sz + 2.5)


def split_paragraphs(runs):
    """Läufe an Zeilenumbrüchen trennen → [[lauf, …], …]."""
    paras, cur = [], []
    for ru in runs:
        pieces = ru[0].split("\n")
        for i, piece in enumerate(pieces):
            if i:
                paras.append(cur)
                cur = []
            if piece:
                cur.append((piece,) + tuple(ru[1:]))
    paras.append(cur)
    return paras


def wrap_runs(para, width):
    """Zeilenumbruch wie Excel (an Leerzeichen); liefert je Zeile die größte Schriftgröße."""
    if not para:
        return [11.0]
    tokens = []           # (wort, breite, leerzeichenbreite, pt)
    for ru in para:
        text, sz, b, i = ru[0], ru[1], ru[2], ru[3]
        for m in re.finditer(r"(\S+)(\s*)", text):
            tokens.append((METRIC.width(m.group(1), sz, b, i), METRIC.width(m.group(2), sz, b, i) if m.group(2) else 0.0, sz))
    lines, cur_w, cur_sz = [], 0.0, 0.0
    started = False
    for w, sp, sz in tokens:
        if started and cur_w + w > width:
            lines.append(cur_sz)
            cur_w, cur_sz = 0.0, 0.0
        cur_w += w + sp
        cur_sz = max(cur_sz, sz)
        started = True
    lines.append(cur_sz or para[0][1])
    return lines


def short(t, n=48):
    t = str(t).replace("\n", " ⏎ ")
    return t if len(t) <= n else t[:n - 1] + "…"


# =============================================================================== Hauptprogramm
def compress(lst):
    """Gleichartige Befunde in zusammenhängenden Zeilen derselben Spalte zu einem Bereich bündeln."""
    def key(f):
        m = re.match(r"^([A-Z]{1,3})(\d+)$", f["ref"]) or re.match(r"^()(\d+):\d+$", f["ref"])
        return (m.group(1), int(m.group(2))) if m else None

    order = {}
    for f in lst:
        order.setdefault(f["sheet"], len(order))
    lst = sorted(lst, key=lambda f: (order[f["sheet"]], key(f) or ("~", 0)))
    out = []
    for f in lst:
        k = key(f)
        if out and k and out[-1]["_k"] and out[-1]["sheet"] == f["sheet"] and out[-1]["_k"][0] == k[0] \
                and out[-1]["_last"] + 1 == k[1] and (f["rule"] != "zeilenraster" or out[-1]["msg"].split(": ")[-1] == f["msg"].split(": ")[-1]):
            out[-1]["_last"] = k[1]
            out[-1]["_n"] += 1
            continue
        g = dict(f, _k=k, _last=k[1] if k else None, _n=1)
        out.append(g)
    for g in out:
        if g["_n"] > 1:
            c = g["_k"][0]
            g["ref"] = f"{c}{g['_k'][1]}:{c}{g['_last']}" if c else f"{g['_k'][1]}:{g['_last']}"
            g["msg"] = f"{g['msg']}  (+{g['_n'] - 1} gleichartige Zeilen)"
    return out


def report(found, max_err, max_warn, show_warn):
    errs = [f for f in found if f["severity"] == FEHLER]
    warns = [f for f in found if f["severity"] == WARNUNG]
    groups = ([(WARNUNG, warns, max_warn)] if show_warn else []) + [(FEHLER, errs, max_err)]
    for sev, items, cap in groups:      # Fehler zuletzt: bleiben im gekürzten Build-Protokoll sichtbar
        by_rule = defaultdict(list)
        for f in items:
            by_rule[f["rule"]].append(f)
        for rule in RULES:
            if rule not in by_rule:
                continue
            lst = compress(by_rule[rule])
            print(f"\n[{sev}] {rule} – {RULES[rule][1]} ({len(by_rule[rule])})")
            for f in lst[:cap]:
                print(f"  {f['sheet']} | {f['ref']} | {rule} | {f['msg']}")
            if len(lst) > cap:
                print(f"  … {len(lst) - cap} weitere Gruppen (--max/--max-warn erhöhen oder --json)")
    # Zusammenfassung (am Ende, damit sie im Build-Protokoll sichtbar bleibt)
    print("\nZusammenfassung je Regel:")
    cnt = Counter(f["rule"] for f in found)
    for rule in RULES:
        if cnt.get(rule):
            print(f"  {RULES[rule][0]:<8} {rule:<18} {cnt[rule]:>4}")
    per_sheet = defaultdict(lambda: [0, 0])
    for f in found:
        per_sheet[f["sheet"]][0 if f["severity"] == FEHLER else 1] += 1
    if per_sheet:
        print("Je Blatt (Fehler / Warnungen):  " + " · ".join(
            f"{s} {e}/{w}" for s, (e, w) in per_sheet.items()))
    print(f"lint_pro: {len(errs)} Fehler, {len(warns)} Warnungen · Metrik: {METRIC.source}")
    return len(errs)


def selftest():
    """Baut eine kleine Mappe mit je einem Verstoß pro prüfbarer Regel und prüft, dass jede Regel anschlägt."""
    import tempfile
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Alignment, Font
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.hyperlink import Hyperlink

    wb = Workbook()
    ws = wb.active
    ws.title = "T"
    ws["A1"], ws["A1"].alignment = "Einzug", Alignment(indent=2)
    ws["A2"], ws["A2"].alignment = "Schrumpfen", Alignment(shrink_to_fit=True)
    ws["A3"], ws["A3"].font = "klein", Font(name="Calibri", sz=7)
    ws["A4"], ws["A4"].font = "grau", Font(name="Calibri", sz=9, color=core.MUTED2)
    ws["A5"], ws["A5"].font = "Arial", Font(name="Arial")
    ws["A6"], ws["B6"] = "ein sehr langer Text, der nicht in die Spalte passt", "x"
    ws["A7"], ws["A7"].number_format = 1234567890.5, '#,##0.00" €"'
    ws["A8"], ws["A8"].number_format, ws["B8"] = 5, '#,##0" €"', "€"
    ws["A9"] = "#DIV/0!"
    ws["A10"] = "Link"
    ws["A10"].hyperlink = Hyperlink(ref="A10", location="'Gibtsnicht'!A1")
    ws["D1"], ws["D1"].alignment = "Umbruch " * 12, Alignment(wrap_text=True)
    ws["D2"], ws["D2"].font = "Groß", Font(name="Calibri", sz=22)
    ws.row_dimensions[2].height = 15
    dv = DataValidation(type="list", formula1='"a,b"', showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("C1")
    for i in range(12, 20):
        ws.cell(i, 1, i)
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=1, min_row=12, max_row=19))
    ws.add_chart(ch, "A13")
    ws.freeze_panes = "B3"
    tmp = os.path.join(tempfile.mkdtemp(prefix="lint_selftest_"), "t.xlsx")
    wb.save(tmp)
    got = {f["rule"] for f in Linter(tmp).run()}
    want = {"einzug", "schrumpfen", "schrift_min", "grau_klein", "schriftart", "text_ueberlauf", "zahl_raute",
            "einheit_doppelt", "fehlerwert", "diagramm_verdeckt", "fixierlinie", "druck_papier", "druck_bereich",
            "gueltigkeit", "link_ziel", "umbruch_hoehe", "zeile_zu_niedrig"}
    missing = want - got
    print("Selbsttest:", "OK" if not missing else f"FEHLT {sorted(missing)}", f"({len(want)} Regeln)")
    return 0 if not missing else 1


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", nargs="?")
    ap.add_argument("--json", help="alle Befunde als JSON schreiben")
    ap.add_argument("--max", type=int, default=20, help="höchstens N Zeilen je Fehler-Regel ausgeben (Standard 20)")
    ap.add_argument("--max-warn", type=int, default=6, help="höchstens N Zeilen je Warn-Regel ausgeben (Standard 6)")
    ap.add_argument("--sheet", help="nur dieses Blatt prüfen")
    ap.add_argument("--skip", default="", help="Regeln auslassen, kommagetrennt")
    ap.add_argument("--no-warn", action="store_true", help="Warnungen nicht einzeln ausgeben")
    ap.add_argument("--list-rules", action="store_true")
    ap.add_argument("--selftest", action="store_true", help="Regeln an einer synthetischen Mappe prüfen")
    a = ap.parse_args()
    if a.selftest:
        sys.exit(selftest())
    if a.list_rules or not a.xlsx:
        for k, (sev, d) in RULES.items():
            print(f"{sev:<8} {k:<18} {d}")
        sys.exit(0)
    lin = Linter(a.xlsx, a.sheet, [s.strip() for s in a.skip.split(",") if s.strip()])
    found = lin.run()
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(found, fh, ensure_ascii=False, indent=1)
    n_err = report(found, a.max, a.max_warn, not a.no_warn)
    sys.exit(1 if n_err else 0)


if __name__ == "__main__":
    main()
