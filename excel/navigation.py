"""Navigation als anklickbare Formen (DrawingML): Reiterleiste, Schritt-Leiste, Unterreiter, Sprungleisten.

Aufruf:  python excel/navigation.py MAPPE.xlsx

Wird NACH der Neuberechnung ausgeführt (LibreOffice würde Formen-Links sonst umschreiben).

- Reiterleiste (alle Blätter, Zeile 2): Marke + 12 Bereiche in vier Gruppen, bündig mit dem Raster der Schritt-Seiten.
  Blätter mit fixierten Spalten (pane xSplit) erhalten eine kompakte Leiste, die ganz im fixierten Bereich liegt
  (Regel: keine Form kreuzt eine Fixierlinie).
- Schritt-Leiste (S01–S12, Zeile 8): zwölf Schritt-Chips + „WEITER ›“, an Zellen verankert von Spalte C bis Ende I.
- Unterreiter (Gruppenblätter, Zeile 5 rechtsbündig): Steuern | Bank | Anhang.
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
from core import ACCENT, BLUE, MIST, MUTED, NAVY, SKY, TINT, WHITE, link_target  # noqa: E402

EMU = 9525  # je Pixel (96 dpi)

# =============================================================================== Namenstabelle (einzige Quelle)
# (Blatt, Langname = Seitentitel C6, Chip-Name – darf nur verkürzen)
STEP_TABLE = [
    ("S01 Objekt", "Objekt", "Objekt"),
    ("S02 Kaufpreis & Miete", "Kaufpreis & Miete", "Kaufpreis"),
    ("S03 Kaufnebenkosten", "Kaufnebenkosten", "Nebenkosten"),
    ("S04 Kaufpreisaufteilung", "Kaufpreisaufteilung", "Aufteilung"),
    ("S05 Maßnahmen & Reserve", "Maßnahmen & Reserve", "Maßnahmen"),
    ("S06 Bewirtschaftung", "Bewirtschaftung", "Bewirtschaftung"),
    ("S07 Finanzierung", "Finanzierung", "Finanzierung"),
    ("S08 Zwischenergebnis", "Zwischenergebnis", "Zwischenergebnis"),
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
     ("Steuern", "Steuern", "Steuern & AfA · Parameter, AfA-Verlauf, steuerliches Ergebnis"),
     ("Finanzierung", "Finanzierung", "Finanzierung · Tilgungspläne der Darlehen"),
     ("Sensitivität", "Sensitivität", "Sensitivität · Break-even, Miete × Zins, IRR-Matrix")],
    [("Bank", "Bankgespräch", "Bank · Investitionsübersicht, Haushaltsrechnung, Vermögensaufstellung"),
     ("Anhang", "Hinweise", "Anhang · Hinweise und Konfiguration")],
]
NAV = [(label, target) for grp in NAV_GROUPS for label, target, _ in grp]
NAV_TIP = {label: tip for grp in NAV_GROUPS for label, _, tip in grp}
ACTIVE = {"AfA-Vergleich": "Steuern", "Haushaltsrechnung": "Bank", "Vermögensaufstellung": "Bank",
          "Bankgespräch": "Bank", "Hinweise": "Anhang", "Konfiguration": "Anhang"}

# Unterreiter der Gruppenblätter: (Beschriftung, Zielblatt)
SUBNAV = {
    "Steuern": [("Steuern & AfA", "Steuern"), ("AfA-Vergleich", "AfA-Vergleich")],
    "Bank": [("Bankgespräch", "Bankgespräch"), ("Haushaltsrechnung", "Haushaltsrechnung"),
             ("Vermögensaufstellung", "Vermögensaufstellung")],
    "Anhang": [("Hinweise", "Hinweise"), ("Konfiguration", "Konfiguration")],
}
# Rechte Inhaltskante der nicht fixierten Gruppenblätter (Unterreiter rechtsbündig dazu)
SUBNAV_EDGE = {"Bankgespräch": "I", "Haushaltsrechnung": "E", "Vermögensaufstellung": "E",
               "Hinweise": "E", "Konfiguration": "G"}

# Sprungleisten (Zeile 8): (Beschriftung, Zielzelle, QuickInfo)
JUMPS = {
    "Eingaben": [("1 Objekt", "B9", "Objekt und Kaufpreis"), ("2 Kaufnebenkosten", "B24", "Kaufnebenkosten"),
                 ("3 Finanzierungs-NK", "B39", "Finanzierungsnebenkosten"), ("4 Aufteilung", "B48", "Kaufpreisaufteilung"),
                 ("5 Maßnahmen", "B61", "Maßnahmen und Reserve"), ("6 Bewirtschaftung", "B75", "Bewirtschaftung"),
                 ("7 Finanzierung", "B95", "Finanzierung"), ("8 Steuern", "B119", "Steuern"),
                 ("9 Exit", "B141", "Prognose und Exit")],
    "Diagramme": [("Investition", "B9", "Investition und Finanzierung"),
                  ("Cashflow", "B30", "Einnahmen, Ausgaben und Cashflow"),
                  ("Entwicklung", "B51", "Entwicklung über die Laufzeit"), ("Steuern", "B92", "Steuern"),
                  ("Exit", "B113", "Exit"), ("Daten", "B138", "Diagrammdaten")],
}


def area_of(sheet):
    """Bereich (Reiter) eines Blatts – auch für Brotkrumen („BANK  ›  HAUSHALTSRECHNUNG“)."""
    if re.match(r"S\d\d ", sheet) or sheet == "Leitfaden":
        return "Leitfaden"
    return ACTIVE.get(sheet, sheet)


# =============================================================================== Gestaltung
TAB_IDLE, TAB_TXT = "1A3F66", MIST        # Reiter auf Navy
STEP_NEXT_BG, STEP_NEXT_TXT = "F3F5F8", "5B6573"   # kommende Schritte (5,4 : 1)
TAB_W, TAB_GAP, GROUP_GAP, TAB_H = 88, 4, 12, 28
INS = 25400                                # Innenabstand links/rechts (2 pt)
FONT = "Aptos"
# Leisten als an Zellen verankerte Gruppe (Standard). NAV_GROUPED=0: jede Form einzeln an Zellen verankert
# (in Excel identisch; nur Vorschau-Renderer mit abweichender Spaltenmetrik zeigen dann Verzerrungen).
GROUPED = os.environ.get("NAV_GROUPED", "1") != "0"

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_HYPER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"


def text_px(text, size, bold=False):
    """Laufweite (px) in Aptos – Näherung wie core.text_px."""
    s = str(text)
    em = size * 96 / 72
    narrow = sum(1 for ch in s if ch in "iljtfrI.,:;!'|() ·")
    wide = sum(1 for ch in s if ch in "MWmw@%€ÄÖÜ≡")
    caps = sum(1 for ch in s if ch.isupper())
    units = len(s) * 0.52 - narrow * 0.22 + wide * 0.25 + caps * 0.08
    return units * em * (1.07 if bold else 1.0)


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
        xm = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", xml, re.S)
        self.xf_fill, self.xf_font, self.xf_align = [], [], []
        for x in re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", xm.group(1) if xm else "", re.S):
            fid = re.search(r'fillId="(\d+)"', x)
            k = int(fid.group(1)) if fid else 0
            self.xf_fill.append(fills[k] if k < len(fills) else None)
            fn = re.search(r'fontId="(\d+)"', x)
            k = int(fn.group(1)) if fn else 0
            self.xf_font.append(fonts[k] if k < len(fonts) else (11.0, False))
            al = re.search(r'<alignment [^>]*horizontal="(\w+)"', x)
            self.xf_align.append(al.group(1) if al else "general")

    def fill(self, s):
        return self.xf_fill[s] if s is not None and s < len(self.xf_fill) else None

    def font(self, s):
        return self.xf_font[s] if s is not None and s < len(self.xf_font) else (11.0, False)

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
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', sheet_xml, re.S):
            r = int(cm.group(2))
            if r > 8:
                continue
            key = (r, col_index(cm.group(1)))
            sm = re.search(r'\bs="(\d+)"', cm.group(3))
            inner = cm.group(5) or ""
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

    def __init__(self, geo, first_id):
        self.geo = geo
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
def nav_frame(geos):
    """Linke/rechte Kante der Reiterleiste: das Inhaltsraster der Schritt-Seiten (Spalte C bis Ende I)."""
    g = geos.get(step_sheet(1))
    left, right = (g.x(3), g.x(10)) if g else (45, 1354)
    return min(max(left, 30), 60), min(max(right, 1300), 1354)


def tab_bar(cv, name, frame, start_cell):
    """Reiterleiste: Marke + 12 Reiter (88 px, 4 px Abstand, 12 px zwischen den Gruppen), rechts bündig."""
    g = cv.geo
    top, h2 = g.row_px(1), g.row_px(2)
    ty = top + (h2 - TAB_H) / 2
    active = area_of(name)
    if g.x_split and g.split_px() < 1200:
        return compact_tab_bar(cv, name, ty, active, start_cell)
    left, right = frame
    n_gaps = sum(len(grp) - 1 for grp in NAV_GROUPS)
    width = 12 * TAB_W + n_gaps * TAB_GAP + (len(NAV_GROUPS) - 1) * GROUP_GAP
    x = right - width
    cv.begin("Reiterleiste")
    cv.add("Marke", left, ty - 3, x - left - 16, TAB_H + 6, None,
           [para([("MM HOLDING", 9.5, WHITE, True, 60)], "l", 90000),
            para([("Immobilien-Kalkulation", 7.5, MIST, False)], "l", 90000)],
           link=("Start", None), tooltip="Zur Startseite", lins=0, rins=0)
    for gi, grp in enumerate(NAV_GROUPS):
        if gi:
            x += GROUP_GAP - TAB_GAP
        for label, target, tip in grp:
            on = label == active
            cv.add(f"Reiter {label}", x, ty, TAB_W, TAB_H, WHITE if on else TAB_IDLE,
                   [para([(label, 9, NAVY if on else TAB_TXT, on)])],
                   link=(target, None), tooltip=tip + (" (aktueller Bereich)" if on else ""))
            x += TAB_W + TAB_GAP
    cv.end("absolute")


def compact_tab_bar(cv, name, ty, active, start_cell):
    """Kompakte Leiste für Blätter mit fixierten Spalten – liegt vollständig links der Fixierlinie."""
    g = cv.geo
    limit = g.split_px() - 14
    items = [("‹ Start", "Start", None, "Zur Startseite"),
             ("Leitfaden", "Leitfaden", None, NAV_TIP["Leitfaden"]),
             ("Dashboard", "Dashboard", None, NAV_TIP["Dashboard"])]
    act_target = next(t for grp in NAV_GROUPS for lab, t, _ in grp if lab == active)
    items.append((active, act_target, None, NAV_TIP[active] + " (aktueller Bereich)"))
    items.append(("≡ Alle Bereiche", "Start", start_cell, "Alle Blätter der Mappe (Übersicht auf der Startseite)"))
    x0 = 14
    cv.begin("Reiterleiste kompakt")
    cv.add("Marke", x0, ty + (TAB_H - 24) / 2, 24, 24, None, [para([("MM", 7.5, MIST, True)])],
           link=("Start", None), tooltip="Zur Startseite", radius=4, line=SKY, lins=0, rins=0)
    x = x0 + 24 + 10
    widths = [max(72, min(96, text_px(lab, 9, True) + 22)) for lab, *_ in items]
    avail = limit - x - TAB_GAP * (len(items) - 1)
    if sum(widths) > avail:
        widths = distribute(int(avail), widths)
    for (label, target, cell, tip), w in zip(items, widths):
        on = label == active
        cv.add(f"Reiter {label}", x, ty, w, TAB_H, WHITE if on else TAB_IDLE,
               [para([(label, 9, NAVY if on else TAB_TXT, on)])], link=(target, cell), tooltip=tip)
        x += w + TAB_GAP
    cv.end("absolute")


def step_bar(cv, step):
    """Schritt-Leiste in Zeile 8: 12 Chips (Breite nach Beschriftung, gleicher Innenabstand) + WEITER, C bis Ende I."""
    g = cv.geo
    left, right = g.x(3), g.x(10)
    y = g.y(8) + 2
    h = max(24, g.row_px(8) - 8)
    btn_w, gap, gap_btn = 190, 4, 12
    need = [text_px(step_chip(i), 7.5, True) for i in range(1, 13)]
    chips_total = int(right - left - btn_w - 11 * gap - gap_btn)
    pad = (chips_total - sum(need)) / 24
    widths = distribute(chips_total, [n + 2 * pad for n in need])
    x = left
    cv.begin("Schritt-Leiste")
    for i, w in enumerate(widths, start=1):
        done, cur = i < step, i == step
        if cur:
            bg, c_no, c_lab = NAVY, SKY, WHITE
        elif done:
            bg, c_no, c_lab = TINT, BLUE, BLUE
        else:
            bg, c_no, c_lab = STEP_NEXT_BG, STEP_NEXT_TXT, STEP_NEXT_TXT
        tip = step_caption(i) + (" (aktuelle Seite)" if cur else "")
        cv.add(f"Schritt {step_no(i)}", x, y, w, h, bg,
               [para([(step_no(i), 7, c_no, True, 20)], "ctr", 95000),
                para([(step_chip(i), 7.5, c_lab, cur)], "ctr", 95000)],
               link=(step_sheet(i), None), tooltip=tip)
        if cur:
            cv.add(f"Schritt {step_no(i)} aktiv", x, y + h + 2, w, 3, ACCENT, prst="rect")
        x += w + gap
    x += gap_btn - gap
    nxt_sheet, nxt_caption = step_next(step)
    last = step >= 12
    cv.add("Weiter", x, y, right - x, h, BLUE,
           [para([("ZUM DASHBOARD  ›" if last else "WEITER  ›", 9.5, WHITE, True, 40)], "ctr", 90000),
            para([(nxt_caption, 8, TINT, False)], "ctr", 90000)],
           link=(nxt_sheet, None),
           tooltip=("Weiter zum Dashboard · Gesamtbewertung" if last else f"Weiter zu {nxt_caption}"))
    cv.end()


def pill_row(cv, items, y, h, x_left=None, x_right=None, size=8, pad=11, gap=4, active=None, name="Reiter"):
    """Reihe kleiner Reiter (Unterreiter / Sprungleiste); links- oder rechtsbündig, an Zellen verankert."""
    widths = [int(round(text_px(lab, size, lab == active) + 2 * pad)) for lab, *_ in items]
    total = sum(widths) + gap * (len(items) - 1)
    x = x_left if x_left is not None else x_right - total
    for (label, target, cell, tip), w in zip(items, widths):
        on = label == active
        cv.add(f"{name} {label}", x, y, w, h, BLUE if on else TINT,
               [para([(label, size, WHITE if on else BLUE, on)])],
               link=(target, cell), tooltip=tip, radius=3)
        x += w + gap
    return x


def row_free(cv, styles, row, x0, x1):
    """True, wenn Zeile row zwischen x0 und x1 weder Werte, überlaufende Texte noch farbige Flächen trägt."""
    g = cv.geo
    for (r, c), (st, has_val) in g.cells.items():
        if r != row:
            continue
        cx0, cx1 = g.x(c), g.x(c + 1)
        text = g.text.get((r, c))
        if text and styles.halign(st) in ("general", "left", "fill"):
            sz, bold = styles.font(st)
            cx1 = max(cx1, cx0 + 8 + text_px(text, sz, bold) * 1.05)
        if cx1 <= x0 - 8 or cx0 >= x1:
            continue
        f = styles.fill(st)
        if has_val or (f and f != "FFFFFF"):
            return False
    return True


def sub_nav(cv, name, frame, styles):
    """Unterreiter der Gruppenblätter, rechtsbündig in der Eyebrow-Zeile 5 (auf fixierten Blättern im fixierten
    Bereich, auf sehr breiten Blättern bündig mit dem Ende der Reiterleiste). Ist Zeile 5 dort belegt
    (z. B. Titelband), weicht die Leiste in Zeile 4 aus."""
    grp = SUBNAV.get(area_of(name))
    if not grp:
        return
    g = cv.geo
    if g.x_split:
        right = g.split_px() - 14
    else:
        edge = SUBNAV_EDGE.get(name, "E")
        right = g.x(col_index(edge) + 1)
        if right > frame[1] + 40:
            right = frame[1]
    active = next(lab for lab, t in grp if t == name)
    items = [(lab, t, None, f"{lab} öffnen" if t != name else f"{lab} (aktuelle Seite)") for lab, t in grp]
    size, pad, gap = 8, 11, 4
    total = sum(int(round(text_px(lab, size, lab == active) + 2 * pad)) for lab, *_ in items) + gap * (len(items) - 1)
    row, h = 5, 22
    for cand in (5, 4):
        if g.row_px(cand) >= 18 and row_free(cv, styles, cand, right - total, right):
            row, h = cand, min(22, g.row_px(cand) - 2)
            break
    y = g.y(row) + (g.row_px(row) - h) / 2
    cv.begin("Unterreiter")
    pill_row(cv, items, y, h, x_right=right, active=active, name="Unterreiter", size=size, pad=pad, gap=gap)
    cv.end()


def band_extension(cv, styles, frame):
    """Sicherheitsnetz: reicht die Navy-Fläche der Kopfleiste (Zeile 2) nicht bis hinter die Reiterleiste,
    wird sie mit zwei Flächenformen (Navy Z. 1–2, Akzentlinie Z. 3) bis zur nächsten Spaltengrenze verlängert."""
    g = cv.geo
    navy_cols = [c for (r, c), (st, _) in g.cells.items() if r == 2 and styles.fill(st) == NAVY]
    if not navy_cols:
        return
    end = g.x(max(navy_cols) + 1)
    need = frame[1] + 14
    if end >= need or g.x_split:
        return
    c, off = g.col_at(need)
    stop = need if off == 0 else need - off + g.col_px(c + 1)
    h12 = g.row_px(1) + g.row_px(2)
    cv.begin("Kopfleiste Verlängerung")
    cv.add("Kopfleiste Fläche", end, 0, stop - end, h12, NAVY, prst="rect")
    cv.add("Kopfleiste Linie", end, h12, stop - end, g.row_px(3), ACCENT, prst="rect")
    cv.end("absolute")


def jump_bar(cv, name):
    """Sprungleiste in Zeile 8 (Eingaben, Diagramme): Beschriftung „Springen zu“ + Abschnitts-Reiter."""
    jumps = JUMPS.get(name)
    if not jumps:
        return
    g = cv.geo
    h = 22
    y = g.y(8) + (g.row_px(8) - h) / 2
    x = g.x(2)
    cap = "SPRINGEN ZU"
    cv.begin("Sprungleiste")
    cw = int(text_px(cap, 7.5, True) + 16)
    cv.add("Sprungleiste Titel", x, y, cw, h, None, [para([(cap, 7.5, MUTED, True, 40)], "l")],
           lins=0, rins=0)
    items = [(lab, name, cell, f"Springen zu Abschnitt: {tip}") for lab, cell, tip in jumps]
    pill_row(cv, items, y, h, x_left=x + cw, name="Sprung", pad=10)
    cv.end()


# =============================================================================== Paket-Hilfen
def shared_strings(files):
    xml = files.get("xl/sharedStrings.xml", b"").decode("utf-8", "ignore")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", xml, flags=re.S):
        out.append("".join(re.findall(r"<t[^>]*>([^<]*)</t>", si)).replace("&amp;", "&"))
    return out


def find_cell(sheet_xml, sst, texts):
    """Adresse der ersten Zelle, deren Text (getrimmt) in texts liegt."""
    for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', sheet_xml):
        idx = int(m.group(2))
        if idx < len(sst) and sst[idx].strip() in texts:
            return m.group(1)
    return None


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
    geos = {name: Geo(files[sp].decode(), sst) for name, sp in sheets}
    styles = Styles(files.get("xl/styles.xml", b"").decode("utf-8", "ignore"))
    frame = nav_frame(geos)
    start_sp = dict(sheets).get("Start")
    start_cell = None
    if start_sp:
        start_cell = find_cell(files[start_sp].decode(), sst, {"Blätter", "Alle Blätter", "Blätter der Mappe"})
    used_drawings = set()
    n_links = 0

    for name, spath in sheets:
        dpath = ensure_drawing(files, spath, used_drawings)
        dxml = strip_monogram(files[dpath].decode())
        prune_rels(files, dpath, dxml)
        drels_p = rels_path(dpath)
        drels = files.get(drels_p, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                   f'<Relationships xmlns="{NS_REL}"></Relationships>'.encode()).decode()
        ids = [int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)] or [1]
        cv = Canvas(geos[name], max(ids) + 100)

        band_extension(cv, styles, frame)
        tab_bar(cv, name, frame, start_cell)
        if name in STEP_OF_SHEET:
            step_bar(cv, STEP_OF_SHEET[name])
        sub_nav(cv, name, frame, styles)
        jump_bar(cv, name)

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
