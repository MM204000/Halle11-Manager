"""Navigation als anklickbare Formen (DrawingML): Reiterleiste, Schritt-Leiste, Unterreiter, Sprungleisten.

Aufruf:  python excel/navigation.py MAPPE.xlsx

Wird NACH der Neuberechnung ausgeführt (LibreOffice würde Formen-Links sonst umschreiben).

- Reiterleiste (alle Blätter, Zeile 2): Marke + 12 gleich breite Reiter in vier Gruppen. Raster je Blatt (Runde 3):
  Marke bündig mit der linken, „Anhang“ bündig mit der rechten Inhaltskante (Geo.content_box); das Kopfband
  (Z. 1–3) endet an derselben Kante. Blätter mit fixierten Jahresspalten tragen dieselbe einzeilige Leiste in
  Standardbreite, an der Fixierlinie in zwei Formgruppen geteilt (keine Form kreuzt die Fixierlinie).
- Schritt-Leiste (S01–S12, Zeile 8): zwölf Schritt-Chips + „WEITER ›“ über die volle Inhaltsbreite; WEITER steht
  genau über dem unteren Weiter-Button (H:I).
- Unterreiter (Gruppenblätter, Zeile 5, rechtsbündig an der rechten Kante der Reiterleiste): Steuer-Tabelle | Bank | Anhang.
- Sprungleisten (Eingaben, Diagramme, Zeile 8).

Die Namenstabelle STEP_TABLE ist die einzige Quelle für alle Schritt-Beschriftungen (Chip, WEITER-Untertitel,
Eyebrow, Fuß-Buttons, Leitfaden-Liste); andere Module importieren sie (import navigation; navigation.step_title(3)).
Alle Formen werden nicht mitgedruckt. Rechenlogik und Zellinhalte bleiben unberührt.
"""
import os
import posixpath
import re
import sys
import zipfile
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core import (ACCENT, BLUE, CONTENT_EDGE, MIST, MUTED, NAVY, SKY, TINT, WHITE,  # noqa: E402
                  link_row, link_target, text_width)

EMU = 9525  # je Pixel (96 dpi)

# =============================================================================== Namenstabelle (einzige Quelle)
# (Blatt, Langname = Seitentitel C6, Chip-Name – darf nur verkürzen)
STEP_TABLE = [
    ("S01 Objekt", "Objekt", "Objekt"),
    ("S02 Kaufpreis & Miete", "Kaufpreis & Miete", "Kaufpreis"),
    ("S03 Kaufnebenkosten", "Kaufnebenkosten", "Nebenkosten"),
    ("S04 Kaufpreisaufteilung", "Kaufpreisaufteilung", "Aufteilung"),
    ("S05 Maßnahmen & Reserve", "Maßnahmen & Reserve", "Maßnahmen"),
    ("S06 Bewirtschaftung", "Bewirtschaftung", "Kosten"),
    ("S07 Finanzierung", "Finanzierung", "Finanzierung"),
    ("S08 Zwischenergebnis", "Zwischenergebnis", "Zwischenstand"),
    ("S09 Steuern", "Steuern", "Steuern"),
    ("S10 Abschreibung", "Abschreibung", "Abschreibung"),
    ("S11 Prognose & Exit", "Prognose & Exit", "Prognose"),
    ("S12 Ergebnis", "Ergebnis", "Ergebnis"),
]
STEPS = [chip for _, _, chip in STEP_TABLE]          # Chip-Namen (Kompatibilität)
STEP_OF_SHEET = {sheet: i for i, (sheet, _, _) in enumerate(STEP_TABLE, start=1)}
AFTER_LAST = ("Dashboard", "Gesamtbewertung")      # Ziel nach Schritt 12
BEFORE_FIRST = ("Leitfaden", "Leitfaden")           # Ziel vor Schritt 1


def step_no(i):
    """Schrittnummer immer zweistellig: „07“."""
    return f"{int(i):02d}"


def step_sheet(i):
    return STEP_TABLE[int(i) - 1][0]


def step_title(i):
    """Langname = Seitentitel C6 („Kaufpreis & Miete“)."""
    return STEP_TABLE[int(i) - 1][1]


def step_chip(i):
    return STEP_TABLE[int(i) - 1][2]


def step_eyebrow(i):
    """Eyebrow C5 der Schritt-Seiten: „SCHRITT 07 / 12“."""
    return f"SCHRITT {step_no(i)} / 12"


def step_caption(i):
    """„Schritt 08 · Zwischenergebnis“ (WEITER-Untertitel, Leitfaden-Liste, QuickInfo)."""
    return f"Schritt {step_no(i)} · {step_title(i)}"


def step_next(i):
    """(Zielblatt, Beschriftung) des nächsten Schritts; nach Schritt 12 das Dashboard."""
    i = int(i)
    return (step_sheet(i + 1), step_caption(i + 1)) if i < 12 else AFTER_LAST


def step_prev(i):
    i = int(i)
    return (step_sheet(i - 1), step_caption(i - 1)) if i > 1 else BEFORE_FIRST


# =============================================================================== Bereiche
# Reiterleiste in Ablaufreihenfolge, vier Gruppen: (Beschriftung, Zielblatt, QuickInfo)
NAV_GROUPS = [
    [("Start", "Start", "Startseite · Kaufstruktur wählen und Überblick"),
     ("Leitfaden", "Leitfaden", "Leitfaden · die Kalkulation in zwölf Schritten")],
    [("Dashboard", "Dashboard", "Dashboard · Gesamtbewertung auf einen Blick"),
     ("Cockpit", "Cockpit", "Cockpit · alle Kennzahlen im Detail"),
     ("Diagramme", "Diagramme", "Diagramme · Investition, Cashflow, Vermögen, Steuern, Exit")],
    [("Eingaben", "Eingaben", "Eingaben · alle Annahmen und Profi-Felder"),
     ("Projektion", "Projektion", "Projektion · Miete, Kosten, Cashflow und Vermögen über 40 Jahre"),
     ("Steuer-Tabelle", "Steuern", "Steuer-Tabelle · Parameter, AfA-Verlauf, steuerliches Ergebnis über 40 Jahre "
      "(Eingaben dazu: Schritt 09 Steuern)"),
     ("Tilgungsplan", "Finanzierung", "Tilgungsplan · Zins und Tilgung der Darlehen Jahr für Jahr "
      "(Eingaben dazu: Schritt 07 Finanzierung)"),
     ("Sensitivität", "Sensitivität", "Sensitivität · Break-even, Miete × Zins, IRR-Matrix")],
    [("Bank", "Bankgespräch", "Bank · Investitionsübersicht, Haushaltsrechnung, Vermögensaufstellung"),
     ("Anhang", "Hinweise", "Anhang · Hinweise und Konfiguration")],
]
NAV = [(label, target) for grp in NAV_GROUPS for label, target, _ in grp]
NAV_TIP = {label: tip for grp in NAV_GROUPS for label, _, tip in grp}
ACTIVE = {"Steuern": "Steuer-Tabelle", "AfA-Vergleich": "Steuer-Tabelle", "Finanzierung": "Tilgungsplan",
          "Haushaltsrechnung": "Bank", "Vermögensaufstellung": "Bank",
          "Bankgespräch": "Bank", "Hinweise": "Anhang", "Konfiguration": "Anhang"}

# Unterreiter der Gruppenblätter: (Beschriftung, Zielblatt)
SUBNAV = {
    "Steuer-Tabelle": [("Steuer-Tabelle", "Steuern"), ("AfA-Vergleich", "AfA-Vergleich")],
    "Bank": [("Bankgespräch", "Bankgespräch"), ("Haushaltsrechnung", "Haushaltsrechnung"),
             ("Vermögensaufstellung", "Vermögensaufstellung")],
    "Anhang": [("Hinweise", "Hinweise"), ("Konfiguration", "Konfiguration")],
}

# Sprungleisten (Zeile 8): (Beschriftung, Abschnittszeile, QuickInfo) – Ziel ist Spalte A der Zeile (P1-08)
JUMPS = {
    "Eingaben": [("1 Objekt", 9, "Objekt und Kaufpreis"), ("2 Kaufnebenkosten", 24, "Kaufnebenkosten"),
                 ("3 Finanzierungs-NK", 39, "Finanzierungsnebenkosten"), ("4 Aufteilung", 48, "Kaufpreisaufteilung"),
                 ("5 Maßnahmen", 61, "Maßnahmen und Reserve"), ("6 Bewirtschaftung", 75, "Bewirtschaftung"),
                 ("7 Finanzierung", 95, "Finanzierung"), ("8 Steuern", 119, "Steuern"),
                 ("9 Exit", 141, "Prognose und Exit")],
    "Diagramme": [("Investition", 9, "Investition und Finanzierung"),
                  ("Cashflow", 30, "Einnahmen, Ausgaben und Cashflow"),
                  ("Entwicklung", 51, "Entwicklung über die Laufzeit"), ("Steuern", 92, "Steuern"),
                  ("Exit", 113, "Exit"), ("Daten", 138, "Diagrammdaten")],
}
JUMP_LIMIT = {"Eingaben": "G"}      # Sprungleiste endet vor der Zell-Legende H8:L8


def area_of(sheet):
    """Bereich (Reiter) eines Blatts – auch für Brotkrumen („BANK  ›  HAUSHALTSRECHNUNG“)."""
    if re.match(r"S\d\d ", sheet) or sheet == "Leitfaden":
        return "Leitfaden"
    return ACTIVE.get(sheet, sheet)


# =============================================================================== Gestaltung
# Farben: Reiter auf Navy (Leerlauf 1A3F66 = Navy aufgehellt, nur hier), Schritt-Zustände (P1-07)
TAB_IDLE, TAB_TXT = "1A3F66", MIST
STEP_DONE_BG = "DCE7F4"                    # erledigt: Fläche, „✓ 01“ + Name in BLUE
STEP_NEXT_BG, STEP_NEXT_LINE, STEP_NEXT_TXT = WHITE, "D5DEEA", MUTED   # kommend: weiß, Rahmen 0,75 pt, 5B6068

# Raster je Blatt (R3-P01): Marke bündig mit der linken, „Anhang“ bündig mit der rechten Inhaltskante des Blatts.
BRAND_X = 45                               # Rückfall: linke Inhaltskante der Schritt-Seiten (Spalte C)
NAV_STD_W = 1302                           # Standardbreite der Leiste (Schritt-Seiten C…I) für Blätter mit Jahresspalten
TAB_MAX, TAB_GAP, GROUP_GAP, TAB_H = 92, 4, 8, 28
TAB_PAD = 4                                # Mindestluft links/rechts neben dem längsten Reitertext
BRAND_GAP = 16                             # Mindestluft zwischen Marke und erstem Reiter
STEP_GAP = 4                               # Schritt-Chips: gleicher Abstand auch vor WEITER
STEP_INS = 38100                           # Innenabstand der Schritt-Chips links/rechts (4 px, R3-P35)
WEITER_COL = "H"                           # WEITER steht genau über dem unteren Weiter-Button (Merge H:I)
PILL_H, PILL_SIZE, PILL_PAD, PILL_GAP = 22, 8, 10, 4   # Unterreiter und Sprungleisten (ein Pill-Stil)
INS = 25400                                # Innenabstand links/rechts (2 pt)
FONT = "Calibri"
# Leisten als an Zellen verankerte Gruppe (Standard). NAV_GROUPED=0: jede Form einzeln an Zellen verankert
# (in Excel identisch; nur Vorschau-Renderer mit abweichender Spaltenmetrik zeigen dann Verzerrungen).
GROUPED = os.environ.get("NAV_GROUPED", "1") != "0"

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_HYPER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"


def text_px(text, size, bold=False):
    """Laufweite (px) in Calibri – Carlito-Metrik aus core (identisch mit Excel)."""
    return text_width(str(text), size, bold)


def xattr(v):
    return escape(str(v), {'"': "&quot;"})


# =============================================================================== Blattgeometrie (Excel-Pixel)
class Styles:
    """Füllfarbe je Zellformat-Index (cellXfs → fills) aus xl/styles.xml."""

    def __init__(self, xml):
        fills = []
        fm = re.search(r"<fills[^>]*>(.*?)</fills>", xml, re.S)
        for f in re.findall(r"<fill>(.*?)</fill>|<fill/>", fm.group(1) if fm else "", re.S):
            pm = re.search(r'patternType="(\w+)"', f or "")
            cm = re.search(r'<fgColor[^>]*rgb="([0-9A-Fa-f]{6,8})"', f or "")
            fills.append(cm.group(1)[-6:].upper() if (pm and pm.group(1) == "solid" and cm) else None)
        fonts = []
        fo = re.search(r"<fonts[^>]*>(.*?)</fonts>", xml, re.S)
        for f in re.findall(r"<font>(.*?)</font>|<font/>", fo.group(1) if fo else "", re.S):
            sz = re.search(r'<sz val="([\d.]+)"', f or "")
            fonts.append((float(sz.group(1)) if sz else 11.0, bool(re.search(r"<b( val=\"(1|true)\")?/>", f or ""))))
        borders = []
        bo = re.search(r"<borders[^>]*>(.*?)</borders>", xml, re.S)
        for b in re.findall(r"<border\b[^>]*?(?:/>|>(.*?)</border>)", bo.group(1) if bo else "", re.S):
            borders.append(bool(re.search(r'\bstyle="(?!none)\w+"', b or "")))
        xm = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", xml, re.S)
        self.xf_fill, self.xf_font, self.xf_align, self.xf_border = [], [], [], []
        for x in re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", xm.group(1) if xm else "", re.S):
            fid = re.search(r'fillId="(\d+)"', x)
            k = int(fid.group(1)) if fid else 0
            self.xf_fill.append(fills[k] if k < len(fills) else None)
            fn = re.search(r'fontId="(\d+)"', x)
            k = int(fn.group(1)) if fn else 0
            self.xf_font.append(fonts[k] if k < len(fonts) else (11.0, False))
            al = re.search(r'<alignment [^>]*horizontal="(\w+)"', x)
            self.xf_align.append(al.group(1) if al else "general")
            bid = re.search(r'borderId="(\d+)"', x)
            k = int(bid.group(1)) if bid else 0
            self.xf_border.append(borders[k] if k < len(borders) else False)

    def fill(self, s):
        return self.xf_fill[s] if s is not None and s < len(self.xf_fill) else None

    def font(self, s):
        return self.xf_font[s] if s is not None and s < len(self.xf_font) else (11.0, False)

    def border(self, s):
        return self.xf_border[s] if s is not None and s < len(self.xf_border) else False

    def halign(self, s):
        return self.xf_align[s] if s is not None and s < len(self.xf_align) else "general"


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


class Geo:
    """Spalten-/Zeilenraster eines Blatts aus der Blatt-XML, in echten Excel-Pixeln (Ziffernbreite 7 px)."""

    def __init__(self, sheet_xml, sst=()):
        # Zellen der Kopfzone (Zeilen 1–8): {(zeile, spalte): (stil, hat_wert)}, Texte in self.text
        self.cells, self.text = {}, {}
        self.body = []                     # (zeile, spalte, stil, hat_wert) ab Zeile 4 – für content_box()
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', sheet_xml, re.S):
            r = int(cm.group(2))
            sm = re.search(r'\bs="(\d+)"', cm.group(3))
            inner = cm.group(5) or ""
            has_val = bool(re.search(r"<v>[^<]|<is>", inner))
            if r >= 4:
                self.body.append((r, col_index(cm.group(1)), int(sm.group(1)) if sm else None, has_val))
            if r > 8:
                continue
            key = (r, col_index(cm.group(1)))
            has_val = bool(re.search(r"<(v|is)\b", inner))
            self.cells[key] = (int(sm.group(1)) if sm else None, has_val)
            vm = re.search(r"<v>([^<]*)</v>", inner)
            tm = re.search(r't="(\w+)"', cm.group(3))
            if tm and tm.group(1) == "s" and vm and int(vm.group(1)) < len(sst):
                self.text[key] = sst[int(vm.group(1))]
            elif vm or "<is>" in inner:
                self.text[key] = vm.group(1) if vm else "".join(re.findall(r"<t[^>]*>([^<]*)</t>", inner))
        fm = re.search(r"<sheetFormatPr[^>]*>", sheet_xml)
        fa = dict(re.findall(r'(\w+)="([^"]*)"', fm.group(0))) if fm else {}
        self.def_row = float(fa.get("defaultRowHeight", 15))
        self.def_col = self.w2px(float(fa["defaultColWidth"])) if "defaultColWidth" in fa else 64
        self.cols = {}
        for cm in re.finditer(r"<col [^>]*/?>", sheet_xml):
            a = dict(re.findall(r'(\w+)="([^"]*)"', cm.group(0)))
            px = 0 if a.get("hidden") in ("1", "true") else (self.w2px(float(a["width"])) if "width" in a else self.def_col)
            for i in range(int(a["min"]), min(int(a["max"]), 400) + 1):
                self.cols[i] = px
        self.rows = {}
        for rm in re.finditer(r"<row [^>]*>", sheet_xml):
            a = dict(re.findall(r'(\w+)="([^"]*)"', rm.group(0)))
            r = int(a["r"])
            if a.get("hidden") in ("1", "true"):
                self.rows[r] = 0
            elif "ht" in a:
                self.rows[r] = round(float(a["ht"]) * 4 / 3)
        self.merges = []
        for mm in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', sheet_xml):
            self.merges.append((int(mm.group(2)), col_index(mm.group(1)), col_index(mm.group(3))))
        pm = re.search(r"<pane [^>]*/>", sheet_xml)
        pa = dict(re.findall(r'(\w+)="([^"]*)"', pm.group(0))) if pm else {}
        frozen = pa.get("state") in ("frozen", "frozenSplit")
        self.x_split = int(float(pa.get("xSplit", 0) or 0)) if frozen else 0
        self.y_split = int(float(pa.get("ySplit", 0) or 0)) if frozen else 0

    @staticmethod
    def w2px(w):
        return int(((256 * w + int(128 / 7)) / 256) * 7)

    def col_px(self, c):
        return self.cols.get(c, self.def_col)

    def row_px(self, r):
        return self.rows.get(r, round(self.def_row * 4 / 3))

    def x(self, c):
        """Linke Kante der Spalte c (1-basiert)."""
        return sum(self.col_px(i) for i in range(1, c))

    def y(self, r):
        return sum(self.row_px(i) for i in range(1, r))

    def col_at(self, x):
        """(Spaltenindex 0-basiert, Versatz px) für eine x-Position."""
        c, left = 1, 0
        while c < 16000:
            w = self.col_px(c)
            if x < left + w:
                return c - 1, x - left
            left += w
            c += 1
        return c - 1, 0

    def row_at(self, y):
        r, top = 1, 0
        while r < 1048576:
            h = self.row_px(r)
            if y < top + h:
                return r - 1, y - top
            top += h
            r += 1
        return r - 1, 0

    def split_px(self):
        return self.x(self.x_split + 1) if self.x_split else 0

    def content_box(self, styles, dxml=""):
        """(links, rechts) in px: äußere Kanten des sichtbaren Seiteninhalts ab Zeile 4 – Zellen mit Wert,
        Füllung oder Rahmen (auch verbundene Bereiche). Ausgeblendete Zeilen und Spalten zählen nicht
        (Hilfsspalten, eingeklappte Datenblöcke). Diagramme liegen im Zellraster (Panels); sie zählen nur als
        Rückfall, damit ein um wenige Pixel überstehender Diagrammrahmen die Kante nicht verschiebt."""
        cols, start = set(), {}
        for r, c, st, has_val in self.body:
            if not self.row_px(r) or not self.col_px(c):
                continue
            f = styles.fill(st)
            if has_val or (f and f != WHITE) or styles.border(st):
                cols.add(c)
                start[(r, c)] = True
        for r, c0, c1 in self.merges:
            if (r, c0) in start:
                cols.update(c for c in range(c0, c1 + 1) if self.col_px(c))
        right_px = [self.x(c + 1) for c in cols]
        left_px = [self.x(c) for c in cols]
        if not right_px:                   # Blatt ohne Zellinhalt: Diagramme/Bilder als Rückfall
            for a in ANCHOR_RE.finditer(dxml):
                body = a.group(0)
                if "<xdr:graphicFrame" not in body and "<xdr:pic" not in body:
                    continue
                fm = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff>", body)
                tm = re.search(r"<xdr:to><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff>", body)
                if fm and tm:
                    left_px.append(self.x(int(fm.group(1)) + 1) + int(fm.group(2)) // EMU)
                    right_px.append(self.x(int(tm.group(1)) + 1) + int(tm.group(2)) // EMU)
        if not right_px:
            return BRAND_X, BRAND_X + NAV_STD_W
        return min(left_px), max(right_px)


# =============================================================================== DrawingML-Bausteine
def run(text, size, color, bold=False, spc=0):
    sp = f' spc="{spc}"' if spc else ""
    return (f'<a:r><a:rPr lang="de-DE" sz="{int(round(size * 100))}" b="{1 if bold else 0}"{sp} dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/></a:rPr><a:t>{escape(text)}</a:t></a:r>')


def para(runs, align="ctr", line_pct=None):
    ls = f'<a:lnSpc><a:spcPct val="{line_pct}"/></a:lnSpc>' if line_pct else ""
    return f'<a:p><a:pPr algn="{align}">{ls}</a:pPr>{"".join(run(*r) for r in runs)}</a:p>'


class Canvas:
    """Sammelt Formen und Hyperlink-Beziehungen einer Zeichnungsebene.

    Leisten werden als Gruppe gesetzt: die Gruppe ist an Zellen verankert (twoCellAnchor, Kanten aus echten
    Excel-Pixeln), die Kinder liegen in einem 1:1-Koordinatenraum. Excel zeigt sie damit pixelgenau;
    Programme mit anderer Spaltenmetrik (LibreOffice) skalieren die Leiste gleichmäßig auf dieselben Spalten.
    """

    def __init__(self, geo, first_id, sheet=""):
        self.geo = geo
        self.sheet = sheet
        self.next_id = first_id
        self.shapes = []
        self.rels = []
        self._grp = None

    def rid(self, sheet, cell=None):
        rid = f"rIdNav{len(self.rels) + 1}"
        target = f"#'{sheet.replace(chr(39), chr(39) * 2)}'!{cell or link_target(sheet)}"
        # interne Ziele ohne TargetMode="External" – so speichert Excel Formen-Links selbst
        self.rels.append(f'<Relationship Id="{rid}" Type="{REL_HYPER}" Target="{xattr(target)}"/>')
        return rid

    def sid(self):
        self.next_id += 1
        return self.next_id

    def _cells(self, x, y, w, h):
        g = self.geo
        c0, dx0 = g.col_at(x)
        r0, dy0 = g.row_at(y)
        c1, dx1 = g.col_at(x + w)
        r1, dy1 = g.row_at(y + h)
        return (f"<xdr:from><xdr:col>{c0}</xdr:col><xdr:colOff>{dx0 * EMU}</xdr:colOff>"
                f"<xdr:row>{r0}</xdr:row><xdr:rowOff>{dy0 * EMU}</xdr:rowOff></xdr:from>"
                f"<xdr:to><xdr:col>{c1}</xdr:col><xdr:colOff>{dx1 * EMU}</xdr:colOff>"
                f"<xdr:row>{r1}</xdr:row><xdr:rowOff>{dy1 * EMU}</xdr:rowOff></xdr:to>")

    def _wrap(self, body, x, y, w, h, cells, edit_as=None):
        if not cells:
            return (f'<xdr:absoluteAnchor><xdr:pos x="{x * EMU}" y="{y * EMU}"/><xdr:ext cx="{w * EMU}" cy="{h * EMU}"/>'
                    f'{body}<xdr:clientData fPrintsWithSheet="0"/></xdr:absoluteAnchor>')
        ea = f' editAs="{edit_as}"' if edit_as else ""
        return (f"<xdr:twoCellAnchor{ea}>{self._cells(x, y, w, h)}{body}"
                f'<xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor>')

    def begin(self, name):
        """Gruppe beginnen – alle folgenden Formen bis end() gehören zu einer Leiste."""
        self._grp = (name, [], [])

    def end(self, edit_as=None):
        """Gruppe an Zellen verankern (edit_as=None: bewegt und skaliert mit den Zellen).

        Der Rahmen wird nach außen auf Spaltengrenzen gerundet (unsichtbare Rahmenform als erstes Kind):
        Anker mit colOff 0 bedeuten in jeder Spaltenmetrik dasselbe, die Leiste bleibt überall proportional."""
        name, items, boxes = self._grp
        self._grp = None
        if not items:
            return
        if not GROUPED:
            for sp, (x, y, w, h) in zip(items, boxes):
                self.shapes.append(self._wrap(sp, x, y, w, h, True, edit_as))
            return
        x0 = min(b[0] for b in boxes)
        y0 = min(b[1] for b in boxes)
        x1 = max(b[0] + b[2] for b in boxes)
        y1 = max(b[1] + b[3] for b in boxes)
        g = self.geo
        c0, d0 = g.col_at(x0)
        c1, d1 = g.col_at(x1)
        sx0 = x0 - d0
        sx1 = x1 if d1 == 0 else x1 - d1 + g.col_px(c1 + 1)
        if (sx0, sx1) != (x0, x1):
            frame = (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self.sid()}" name="{xattr(name)} Rahmen"/>'
                     f'<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="{sx0 * EMU}" y="{y0 * EMU}"/>'
                     f'<a:ext cx="{(sx1 - sx0) * EMU}" cy="{(y1 - y0) * EMU}"/></a:xfrm><a:prstGeom prst="rect">'
                     f'<a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:sp>')
            items.insert(0, frame)
            x0, x1 = sx0, sx1
        off = f'<a:off x="{x0 * EMU}" y="{y0 * EMU}"/><a:ext cx="{(x1 - x0) * EMU}" cy="{(y1 - y0) * EMU}"/>'
        ch = (f'<a:chOff x="{x0 * EMU}" y="{y0 * EMU}"/>'
              f'<a:chExt cx="{(x1 - x0) * EMU}" cy="{(y1 - y0) * EMU}"/>')
        body = (f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{self.sid()}" name="{xattr(name)}"/><xdr:cNvGrpSpPr/>'
                f'</xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm>{off}{ch}</a:xfrm></xdr:grpSpPr>{"".join(items)}</xdr:grpSp>')
        self.shapes.append(self._wrap(body, x0, y0, x1 - x0, y1 - y0, True, edit_as))

    def add(self, name, x, y, w, h, fill=None, paras=(), link=None, tooltip=None, radius=4, line=None,
            prst="roundRect", lins=INS, rins=INS, anchor="ctr", line_w=9525):
        """link: (Blatt, Zelle|None). radius in px. Außerhalb einer Gruppe: absolut verankert."""
        x, y, w, h = int(round(x)), int(round(y)), int(round(w)), int(round(h))
        hl = ""
        if link:
            tip = f' tooltip="{xattr(tooltip)}"' if tooltip else ""
            hl = f'<a:hlinkClick r:id="{self.rid(*link)}"{tip}/>'
        geom = (f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val '
                f'{min(50000, int(radius / max(1, min(w, h)) * 100000))}"/></a:avLst></a:prstGeom>'
                if prst == "roundRect" else f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>')
        fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
        ln = (f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
              else "<a:ln><a:noFill/></a:ln>")
        body = ""
        if paras:
            body = (f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" lIns="{lins}" '
                    f'tIns="0" rIns="{rins}" bIns="0" rtlCol="0" anchor="{anchor}"><a:noAutofit/></a:bodyPr>'
                    f'<a:lstStyle/>{"".join(paras)}</xdr:txBody>')
        sp = (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self.sid()}" name="{xattr(name)}">{hl}'
              f'</xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="{x * EMU}" y="{y * EMU}"/>'
              f'<a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>{geom}{fl}{ln}</xdr:spPr>{body}</xdr:sp>')
        if self._grp is not None:
            self._grp[1].append(sp)
            self._grp[2].append((x, y, w, h))
        else:
            self.shapes.append(self._wrap(sp, x, y, w, h, False))


def distribute(total, weights):
    """Ganzzahlige Breiten, die sich exakt zu total summieren."""
    s = sum(weights)
    raw = [total * w / s for w in weights]
    out = [int(v) for v in raw]
    rest = total - sum(out)
    for i in sorted(range(len(raw)), key=lambda k: raw[k] - out[k], reverse=True)[:rest]:
        out[i] += 1
    return out


# =============================================================================== Leisten
class Frame:
    """Raster der Navigation je Blatt (R3-P01/P03/P04).

    - box[blatt] = (links, rechts): Marke bündig mit der linken, letzter Reiter („Anhang“) bündig mit der rechten
      Inhaltskante (Geo.content_box). Kopfband (Z. 1–3) endet an derselben Kante.
    - Blätter mit fixierten Jahresspalten (Projektion, Steuern, AfA-Vergleich, Finanzierung): dieselbe einzeilige
      Leiste in Standardbreite (NAV_STD_W = Schritt-Seiten) ab der linken Inhaltskante; das Band läuft über die
      Jahresspalten weiter.
    - Unterreiter: je Blatt rechtsbündig an der rechten Kante der eigenen Leiste, mittig in Zeile 5."""

    def __init__(self, geos, styles, drawings):
        self.box, self.band = {}, {}
        std = None
        for name, g in geos.items():
            left, right = g.content_box(styles, drawings.get(name, ""))
            self.band[name] = right
            if name in STEP_OF_SHEET and std is None:
                std = right - left
            self.box[name] = [left, right]
        std = std or NAV_STD_W
        for name, g in geos.items():
            left, right = self.box[name]
            if g.x_split or right - left > std * 1.6:
                self.box[name][1] = left + std
                self.band[name] = max(self.band[name], left + std)

    def edges(self, name):
        return self.box.get(name, (BRAND_X, BRAND_X + NAV_STD_W))


def nav_frame(geos, styles=None, drawings=None):
    return Frame(geos, styles, drawings or {})


def tab_metrics(width):
    """Reiterbreite, Schriftgröße und Abstände für eine Leistenbreite: Reiter höchstens TAB_MAX breit; reicht die
    Breite nicht, zuerst engere Abstände, dann 8,5 bzw. 8 pt – nie abgeschnittene Beschriftungen."""
    labels = [lab for lab, _ in NAV]
    brand = int(text_px("Immobilien-Kalkulation", 8)) + 2
    for size, gap, ggap in ((9, TAB_GAP, GROUP_GAP), (9, 3, 6), (8.5, 3, 6), (8, 3, 6), (8, 2, 4)):
        need = max(text_px(lab, size, True) for lab in labels) + 2 * TAB_PAD
        gaps = 8 * gap + 3 * ggap
        w = min(TAB_MAX, int((width - brand - BRAND_GAP - gaps) / 12))
        if w >= need:
            return w, size, gap, ggap
    return w, size, gap, ggap


def tab_layout(left, right, split=0):
    """Positionen der Leiste: (Reiterbreite, Schrift, [(x, Beschriftung, Ziel, QuickInfo)], Markenbreite, Teilung).

    Ohne Fixierung: „Anhang“ endet bündig an right. Mit fixierten Spalten (split = x der Fixierlinie) steht
    zwischen zwei Reitern genau auf der Fixierlinie eine Fuge; die Leiste zerfällt dort in zwei Formgruppen, von
    denen keine die Fixierlinie kreuzt (Excel zeichnet Formen je Fensterausschnitt). Gewählt wird die Teilung,
    deren rechte Kante der Standardbreite am nächsten kommt und die der Marke genug Raum lässt."""
    tab_w, size, gap, ggap = tab_metrics(right - left)
    seq = []                                   # (Beschriftung, Ziel, QuickInfo, Abstand davor)
    for gi, grp in enumerate(NAV_GROUPS):
        for k, (label, target, tip) in enumerate(grp):
            seq.append((label, target, tip, 0 if not seq else (ggap if k == 0 else gap)))
    brand_min = int(text_px("Immobilien-Kalkulation", 8)) + 2 + BRAND_GAP

    def place(x0):
        out, x = [], x0
        for label, target, tip, before in seq:
            x += before
            out.append((x, label, target, tip))
            x += tab_w
        return out
    width = 12 * tab_w + sum(b for *_, b in seq)
    if not split or split <= left + brand_min:
        return tab_w, size, place(right - width), 0
    best = None
    for k in range(1, 12):                     # Fuge vor Reiter k liegt mittig auf der Fixierlinie
        before = seq[k][3]
        x_k = split + (before - before // 2)
        x0 = x_k - (k * tab_w + sum(b for *_, b in seq[1:k]) + before)
        if x0 - left < brand_min:
            continue
        pos = place(x0)
        score = abs(pos[-1][0] + tab_w - right)
        if best is None or score < best[0]:
            best = (score, pos, k)
    if best is None:
        return tab_w, size, place(right - width), 0
    return tab_w, size, best[1], best[2]


def tab_bar(cv, name, frame):
    """Reiterleiste (eine Komponente für alle 28 Blätter): Marke „MM HOLDING / Immobilien-Kalkulation“ bündig mit
    der linken Inhaltskante, 12 gleich breite Reiter in vier Gruppen, „Anhang“ bündig mit der rechten Kante.
    Blätter mit fixierten Spalten: dieselbe einzeilige Leiste, an der Fixierlinie in zwei Gruppen geteilt."""
    g = cv.geo
    top, h2 = g.row_px(1), g.row_px(2)
    ty = top + (h2 - TAB_H) / 2
    active = area_of(name)
    left, right = frame.edges(name)
    tab_w, size, pos, cut = tab_layout(left, right, g.split_px())
    cv.begin("Reiterleiste")
    cv.add("Marke", left, ty - 3, pos[0][0] - left - BRAND_GAP, TAB_H + 6, None,
           [para([("MM HOLDING", 10, WHITE, True, 60)], "l", 90000),
            para([("Immobilien-Kalkulation", 8, MIST, False)], "l", 90000)],
           link=("Start", None), tooltip="Zur Startseite", lins=0, rins=0)
    for k, (x, label, target, tip) in enumerate(pos):
        if cut and k == cut:
            cv.end("absolute")
            cv.begin("Reiterleiste rechts")
        on = label == active
        cv.add(f"Reiter {label}", x, ty, tab_w, TAB_H, WHITE if on else TAB_IDLE,
               [para([(label, size, NAVY if on else TAB_TXT, on)])],
               link=(target, None), tooltip=tip + (" (aktueller Bereich)" if on else ""),
               lins=0, rins=0)
    cv.end("absolute")
    return pos[-1][0] + tab_w


def step_bar(cv, step, frame):
    """Schritt-Leiste in Zeile 8 über die volle Inhaltsbreite (R3-P02): 12 gleich breite Chips + „WEITER ›“.
    WEITER ist genau so breit wie der untere Weiter-Button (H:I) und steht bündig darüber; alle Abstände 4 px.
    Zustände: besucht (DCE7F4, nur Nummer – ✓ bleibt dem Leitfaden-Status vorbehalten, R3-P25), aktiv (Navy +
    Akzentstrich), kommend (weiß, Rahmen D5DEEA, Schrift 5B6068). Innenabstand 4 px, keine Schrift unter 8 pt."""
    g = cv.geo
    left, right = frame.edges(cv.sheet)
    y = g.y(8) + 2
    h = max(24, g.row_px(8) - 8)
    x_weiter = g.x(col_index(WEITER_COL))
    if not left + 12 * 60 < x_weiter < right - 120:       # Rückfall bei abweichendem Raster
        x_weiter = right - 300
    pill_w = distribute(x_weiter - left - 12 * STEP_GAP, [1] * 12)
    x = left
    cv.begin("Schritt-Leiste")
    for i, w in enumerate(pill_w, start=1):
        done, cur = i < step, i == step
        line = None
        if cur:
            bg, c_no, c_lab = NAVY, SKY, WHITE
        elif done:
            bg, c_no, c_lab = STEP_DONE_BG, BLUE, BLUE
        else:
            bg, c_no, c_lab, line = STEP_NEXT_BG, STEP_NEXT_TXT, STEP_NEXT_TXT, STEP_NEXT_LINE
        lab = step_chip(i)
        spc = -10 if text_px(lab, 8) > w - 8 else 0
        tip = step_caption(i) + (" (aktuelle Seite)" if cur else " (bereits durchlaufen)" if done else "")
        cv.add(f"Schritt {step_no(i)}", x, y, w, h, bg,
               [para([(step_no(i), 8, c_no, True)], "ctr", 95000),
                para([(lab, 8, c_lab, False, spc)], "ctr", 95000)],
               link=(step_sheet(i), None), tooltip=tip, line=line, lins=STEP_INS, rins=STEP_INS)
        if cur:
            cv.add(f"Schritt {step_no(i)} aktiv", x, y + h + 2, w, 3, ACCENT, prst="rect")
        x += w + STEP_GAP
    nxt_sheet, nxt_caption = step_next(step)
    last = step >= 12
    cv.add("Weiter", x_weiter, y, right - x_weiter, h, BLUE,
           [para([("ZUM DASHBOARD  ›" if last else "WEITER  ›", 9, WHITE, True, 40)], "ctr", 90000),
            para([(nxt_caption, 8, TINT, False)], "ctr", 90000)],
           link=(nxt_sheet, None), lins=STEP_INS, rins=STEP_INS,
           tooltip=("Weiter zum Dashboard · Gesamtbewertung" if last else f"Weiter zu {nxt_caption}"))
    cv.end()


def pill_w(label, size=PILL_SIZE, pad=PILL_PAD):
    """Breite einer Pill – immer mit fetter Laufweite, damit aktive/inaktive Zustände gleich groß sind."""
    return int(round(text_px(label, size, True) + 2 * pad))


def pill_row(cv, items, y, h=PILL_H, x_left=None, x_right=None, size=PILL_SIZE, pad=PILL_PAD, gap=PILL_GAP,
             active=None, name="Reiter"):
    """Reihe kleiner Reiter (Unterreiter / Sprungleiste) im einheitlichen Pill-Stil: E7EEF7 mit 1D4F8A,
    aktiv 1D4F8A mit Weiß fett; links- oder rechtsbündig."""
    widths = [pill_w(lab, size, pad) for lab, *_ in items]
    total = sum(widths) + gap * (len(items) - 1)
    x = x_left if x_left is not None else x_right - total
    for (label, target, cell, tip), w in zip(items, widths):
        on = label == active
        cv.add(f"{name} {label}", x, y, w, h, BLUE if on else TINT,
               [para([(label, size, WHITE if on else BLUE, on)])],
               link=(target, cell), tooltip=tip, radius=3, lins=0, rins=0)
        x += w + gap
    return x


def sub_nav(cv, name, frame):
    """Unterreiter der Gruppenblätter (R3-P04): Zeile 5, rechtsbündig an der rechten Kante der Reiterleiste dieses
    Blatts (= rechte Inhaltskante), senkrecht mittig in Zeile 5. Links in derselben Zeile steht die Eyebrow."""
    area = area_of(name)
    grp = SUBNAV.get(area)
    if not grp:
        return
    g = cv.geo
    right = frame.edges(name)[1]
    y = g.y(5) + (g.row_px(5) - PILL_H) / 2
    active = next(lab for lab, t in grp if t == name)
    items = [(lab, t, None, f"{lab} öffnen" if t != name else f"{lab} (aktuelle Seite)") for lab, t in grp]
    cv.begin("Unterreiter")
    pill_row(cv, items, y, x_right=right, active=active, name="Unterreiter")
    cv.end()


# Rechte Kante der Kopfleisten-Fläche: Frame.band (Inhaltskante je Blatt; core.CONTENT_EDGE nur noch für chrome.py)
BAND_TO = CONTENT_EDGE


def trim_band(sxml, geo, styles, end_px):
    """Navy-/Akzentzellen der Kopfleiste (Z. 1–3) rechts von end_px auf das Standardformat zurücksetzen.

    Die Zellfläche endet an der letzten Spaltengrenze ≤ end_px; den Rest ergänzt band_extension als Form.
    Nur leere Zellen werden angefasst (Werte und Formeln bleiben unberührt)."""
    def fix_row(m):
        row = m.group(0)

        def fix_cell(cm):
            c = col_index(cm.group(2))
            sm = re.search(r'\bs="(\d+)"', cm.group(3))
            st = int(sm.group(1)) if sm else None
            if cm.group(4) != "/>" or styles.fill(st) not in (NAVY, ACCENT) or geo.x(c + 1) <= end_px + 2:
                return cm.group(0)
            r = int(cm.group(1)[len(cm.group(2)):])
            geo.cells[(r, c)] = (0, False)
            attrs = re.sub(r'\s*\bs="\d+"', "", cm.group(3))
            return f'<c r="{cm.group(1)}"{attrs} s="0"/>'
        return re.sub(r'<c r="(([A-Z]+)\d+)"([^>]*?)(/>|>)', fix_cell, row)
    return re.sub(r'<row r="[123]"[^>]*[^/]>.*?</row>', fix_row, sxml, flags=re.S)


def band_extension(cv, styles, end_px):
    """Sicherheitsnetz: reicht die Navy-Fläche der Kopfleiste (Zeile 2) nicht bis end_px, wird sie mit zwei
    Flächenformen (Navy Z. 1–2, Akzentlinie Z. 3) verlängert."""
    g = cv.geo
    navy_cols = [c for (r, c), (st, _) in g.cells.items() if r == 2 and styles.fill(st) == NAVY]
    if not navy_cols:
        return
    end = g.x(max(navy_cols) + 1)
    if end >= end_px or g.x_split:
        return
    start = max(0, end - 3)            # 3 px Überlappung: keine Haarfuge zwischen Zellfläche und Form
    h12 = g.row_px(1) + g.row_px(2)
    cv.begin("Kopfleiste Verlängerung")
    cv.add("Kopfleiste Fläche", start, 0, end_px - start, h12, NAVY, prst="rect")
    cv.add("Kopfleiste Linie", start, h12, end_px - start, g.row_px(3), ACCENT, prst="rect")
    cv.end("absolute")


def jump_bar(cv, name, frame):
    """Sprungleiste in Zeile 8 (Eingaben, Diagramme): Beschriftung „SPRINGEN ZU“ + Abschnitts-Pills im
    Pill-Stil der Unterreiter. Ziele: Abschnittsanker in Spalte A (core.link_row), damit Excel den Abschnitt
    an den Anfang des Scrollbereichs rollt."""
    jumps = JUMPS.get(name)
    if not jumps:
        return
    g = cv.geo
    y = g.y(8) + (g.row_px(8) - PILL_H) / 2
    x = frame.edges(name)[0]
    cap = "SPRINGEN ZU"
    cw = int(text_px(cap, 8, True) + 12)
    items = [(lab, name, link_row(name, row), f"Springen zu Abschnitt: {tip}") for lab, row, tip in jumps]
    # nie über den reservierten Bereich hinaus (Eingaben: Legende ab Spalte H, Wunsch E)
    limit = g.x(col_index(JUMP_LIMIT[name]) + 1) if name in JUMP_LIMIT else None
    pad = PILL_PAD
    while limit and pad > 6 and x + cw + sum(pill_w(lab, pad=pad) for lab, *_ in items) \
            + PILL_GAP * (len(items) - 1) > limit - 8:
        pad -= 1
    cv.begin("Sprungleiste")
    cv.add("Sprungleiste Titel", x, y, cw, PILL_H, None, [para([(cap, 8, MUTED, True, 40)], "l")],
           lins=0, rins=0)
    pill_row(cv, items, y, x_left=x + cw, name="Sprung", pad=pad)
    cv.end()


# =============================================================================== Paket-Hilfen
def shared_strings(files):
    xml = files.get("xl/sharedStrings.xml", b"").decode("utf-8", "ignore")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", xml, flags=re.S):
        out.append("".join(re.findall(r"<t[^>]*>([^<]*)</t>", si)).replace("&amp;", "&"))
    return out


def rels_path(part):
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def ensure_drawing(files, spath, used):
    """Zeichnungsebene des Blatts finden oder neu anlegen (falls ein Blatt keine Bilder/Diagramme hat)."""
    srels_p = rels_path(spath)
    srels = files.get(srels_p, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                f'<Relationships xmlns="{NS_REL}"></Relationships>'.encode()).decode()
    for rel in re.findall(r"<Relationship [^>]*/>", srels):
        tm = re.search(r'Type="([^"]+)"', rel)
        if tm and tm.group(1) == REL_DRAWING:  # nicht legacyDrawing/vmlDrawing (Kommentare)
            target = re.search(r'Target="([^"]+)"', rel).group(1)
            return posixpath.normpath(posixpath.join(posixpath.dirname(spath), target))
    n = 1
    while f"xl/drawings/drawing{n}.xml" in files or n in used:
        n += 1
    used.add(n)
    dpath = f"xl/drawings/drawing{n}.xml"
    files[dpath] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>').encode()
    rid = "rIdNavDrawing"
    srels = srels.replace("</Relationships>",
                          f'<Relationship Id="{rid}" Type="{REL_DRAWING}" Target="../drawings/drawing{n}.xml"/></Relationships>')
    files[srels_p] = srels.encode()
    sxml = files[spath].decode()
    for tag in ("<legacyDrawing", "<legacyDrawingHF", "<drawingHF", "<picture", "<oleObjects", "<controls",
                "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"):
        k = sxml.find(tag)
        if k >= 0:
            sxml = sxml[:k] + f'<drawing r:id="{rid}"/>' + sxml[k:]
            break
    if 'xmlns:r="' not in sxml[:1000]:
        sxml = sxml.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    files[spath] = sxml.encode()
    ct = files["[Content_Types].xml"].decode()
    if f'PartName="/{dpath}"' not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="/{dpath}" ContentType="{CT_DRAWING}"/></Types>')
        files["[Content_Types].xml"] = ct.encode()
    return dpath


ANCHOR_RE = re.compile(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", re.S)


def strip_monogram(dxml):
    """Monogramm-Bild der alten Kopfleiste (klein, Spalte A, Zeile 1–2) entfernen."""
    def keep(m):
        a = m.group(0)
        if "<xdr:pic>" not in a and "<xdr:pic " not in a:
            return a
        fm = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", a, re.S)
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"', a)
        if fm and int(fm.group(1)) == 0 and int(fm.group(2)) <= 1 and ext and int(ext.group(1)) < 480000:
            return ""
        return a
    return ANCHOR_RE.sub(keep, dxml)


def prune_rels(files, dpath, dxml):
    """Bildbeziehungen entfernen, die die Zeichnung nicht mehr verwendet."""
    p = rels_path(dpath)
    if p not in files:
        return
    drels = files[p].decode()

    def keep(m):
        rel = m.group(0)
        rid = re.search(r'Id="([^"]+)"', rel).group(1)
        if REL_IMAGE in rel and f'"{rid}"' not in dxml:
            return ""
        return rel
    files[p] = re.sub(r"<Relationship [^>]*/>", keep, drels).encode()


def prune_media(files):
    """Nicht mehr referenzierte Mediendateien entfernen."""
    used = set()
    for n, data in files.items():
        if n.endswith(".rels"):
            base = posixpath.dirname(posixpath.dirname(n))
            for t in re.findall(r'Target="([^"]+)"', data.decode("utf-8", "ignore")):
                if "media/" in t:
                    used.add(posixpath.normpath(posixpath.join(base, t)))
    for n in [n for n in files if n.startswith("xl/media/")]:
        if n not in used:
            del files[n]
            ct = files["[Content_Types].xml"].decode()
            ct = re.sub(rf'<Override PartName="/{re.escape(n)}"[^>]*/>', "", ct)
            files["[Content_Types].xml"] = ct.encode()


# =============================================================================== Einsetzen
def inject(path):
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        files = {i.filename: zin.read(i.filename) for i in infos}

    wb = files["xl/workbook.xml"].decode()
    wb_rels = files["xl/_rels/workbook.xml.rels"].decode()
    rel_target = {}
    for rm in re.finditer(r"<Relationship [^>]*/>", wb_rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', rm.group(0)))
        rel_target[a["Id"]] = a["Target"]
    sheets = []
    for sm in re.finditer(r"<sheet [^>]*/>", wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', sm.group(0)))
        name = a["name"].replace("&amp;", "&").replace("&apos;", "'").replace("&quot;", '"')
        t = rel_target[a["r:id"]]
        sheets.append((name, t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))))

    sst = shared_strings(files)
    styles = Styles(files.get("xl/styles.xml", b"").decode("utf-8", "ignore"))
    geos = {name: Geo(files[sp].decode(), sst) for name, sp in sheets}
    used_drawings = set()
    dpaths, drawings = {}, {}
    for name, spath in sheets:
        dpaths[name] = ensure_drawing(files, spath, used_drawings)
        drawings[name] = strip_monogram(files[dpaths[name]].decode())
    frame = nav_frame(geos, styles, drawings)
    n_links = 0

    for name, spath in sheets:
        dpath = dpaths[name]
        dxml = drawings[name]
        prune_rels(files, dpath, dxml)
        drels_p = rels_path(dpath)
        drels = files.get(drels_p, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                   f'<Relationships xmlns="{NS_REL}"></Relationships>'.encode()).decode()
        ids = [int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)] or [1]
        cv = Canvas(geos[name], max(ids) + 100, name)

        end_px = frame.band[name]
        files[spath] = trim_band(files[spath].decode(), geos[name], styles, end_px).encode()
        band_extension(cv, styles, end_px)
        frame.box[name][1] = tab_bar(cv, name, frame)      # tatsächliche rechte Kante (geteilte Leiste)
        if name in STEP_OF_SHEET:
            step_bar(cv, STEP_OF_SHEET[name], frame)
        sub_nav(cv, name, frame)
        jump_bar(cv, name, frame)

        if 'xmlns:r="' not in dxml[:800]:
            dxml = dxml.replace("<xdr:wsDr ", '<xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        if "</xdr:wsDr>" in dxml:
            dxml = dxml.replace("</xdr:wsDr>", "".join(cv.shapes) + "</xdr:wsDr>")
        else:  # leere Zeichnung als <xdr:wsDr .../>
            dxml = re.sub(r"<xdr:wsDr([^>]*)/>", lambda m: f"<xdr:wsDr{m.group(1)}>{''.join(cv.shapes)}</xdr:wsDr>", dxml)
        files[dpath] = dxml.encode()
        files[drels_p] = drels.replace("</Relationships>", "".join(cv.rels) + "</Relationships>").encode()
        n_links += len(cv.rels)

    prune_media(files)
    names = [i.filename for i in infos]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            if i.filename in files:
                zout.writestr(i, files[i.filename])
        for n, data in files.items():
            if n not in names:
                zout.writestr(n, data)
    print(f"Navigation eingesetzt: {path} ({len(sheets)} Blätter, {n_links} Formen-Links)")


if __name__ == "__main__":
    inject(sys.argv[1])
