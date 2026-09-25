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
from core import (ACCENT, BLUE, CONTENT_EDGE, GOLD, ICE, MIST, MUTED, NAVY, NAVY_2, SKY, WHITE,  # noqa: E402
                  link_row, link_target, text_width)

EMU = 9525  # je Pixel (96 dpi)

# =============================================================================== Namenstabelle (einzige Quelle)
# (Blatt, Langname = Seitentitel C6, Chip-Name – darf nur verkürzen; Runde 4 P2-05: ein Name pro Schritt, nur
#  Kaufpreis/Aufteilung/Maßnahmen bleiben als Anfangs- bzw. Endwort des Titels)
STEP_TABLE = [
    ("S01 Objekt", "Objekt", "Objekt"),
    ("S02 Kaufpreis & Miete", "Kaufpreis & Miete", "Kaufpreis"),
    ("S03 Kaufnebenkosten", "Kaufnebenkosten", "Kaufnebenkosten"),
    ("S04 Kaufpreisaufteilung", "Kaufpreisaufteilung", "Aufteilung"),
    ("S05 Maßnahmen & Reserve", "Maßnahmen & Reserve", "Maßnahmen"),
    ("S06 Bewirtschaftung", "Bewirtschaftung", "Bewirtschaftung"),
    ("S07 Finanzierung", "Finanzierung", "Finanzierung"),
    ("S08 Zwischenergebnis", "Zwischenergebnis", "Zwischenergebnis"),
    ("S09 Steuern", "Steuern", "Steuern"),
    ("S10 Abschreibung", "Abschreibung", "Abschreibung"),
    ("S11 Prognose & Exit", "Prognose & Exit", "Prognose & Exit"),
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
    [("Bank", "Bankgespräch", "Bank · Investitionsübersicht, Haushaltsrechnung, Vermögensaufstellung, Exposé"),
     ("Anhang", "Hinweise", "Anhang · Hinweise und Konfiguration")],
]
NAV = [(label, target) for grp in NAV_GROUPS for label, target, _ in grp]
NAV_TIP = {label: tip for grp in NAV_GROUPS for label, _, tip in grp}
ACTIVE = {"Steuern": "Steuer-Tabelle", "AfA-Vergleich": "Steuer-Tabelle", "Finanzierung": "Tilgungsplan",
          "Haushaltsrechnung": "Bank", "Vermögensaufstellung": "Bank",
          "Bankgespräch": "Bank", "Exposé": "Bank", "Hinweise": "Anhang", "Konfiguration": "Anhang"}

# Unterreiter der Gruppenblätter: (Beschriftung, Zielblatt)
SUBNAV = {
    "Steuer-Tabelle": [("Steuer-Tabelle", "Steuern"), ("AfA-Vergleich", "AfA-Vergleich")],
    "Bank": [("Bankgespräch", "Bankgespräch"), ("Haushaltsrechnung", "Haushaltsrechnung"),
             ("Vermögensaufstellung", "Vermögensaufstellung"), ("Exposé", "Exposé")],   # Exposé: Runde 6 (N)
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
# Farben (Runde 5 „Midnight & Gold“, nur core-Tokens): Kopfleiste NAVY mit GOLD-Linie (Z. 3); inaktive Reiter
# NAVY_2 mit MIST-Schrift (8,4:1), aktiver Reiter weiß, Schrift NAVY fett, mit GOLD-Kante (Unterstrich im Reiter).
TAB_IDLE, TAB_TXT = NAVY_2, MIST
TAB_ON, TAB_ON_TXT, TAB_MARK = WHITE, NAVY, GOLD
BAND_FILLS = (NAVY, GOLD, ACCENT)          # Zellfüllungen der Kopfleiste (ACCENT: Altstand vor Runde 5)
# Schritt-Zustände: erledigt „✓ 01“ ICE mit BLUE · aktiv NAVY mit GOLD-Unterlinie · offen weiß, Rahmen MIST, Schrift MUTED
STEP_DONE_BG = ICE
STEP_NEXT_BG, STEP_NEXT_LINE, STEP_NEXT_TXT = WHITE, MIST, MUTED
STEP_MARK = GOLD
# WEITER = Primär-Button-Look (core.BTN["primary"]: Nachtblau, Goldrahmen, Schrift weiß)
WEITER_BG, WEITER_LINE, WEITER_SUB = NAVY, GOLD, MIST
# Pills (Unterreiter, Sprungleisten) = Chip-Look (core.BTN["chip"]): ICE, Rahmen MIST, Schrift BLUE; aktiv NAVY/weiß
PILL_BG, PILL_LINE, PILL_TXT, PILL_ON = ICE, MIST, BLUE, NAVY

# Reiterleiste mit FESTER Geometrie auf allen 28 Blättern (Runde 4 P1-08, Excel-Pixel ab Blattursprung,
# xdr:absoluteAnchor – unabhängig von Spaltenbreiten, kein Einrasten auf Spaltengrenzen):
#   Marke ab NAV_X (= linke Inhaltskante aller Blätter, Spalte B bzw. C), 12 Reiter à TAB_W, Abstand konstant TAB_GAP,
#   x_n = TAB_X0 + n × (TAB_W + TAB_GAP); „Anhang“ endet bei NAV_END (= rechte Inhaltskante der Schritt-Seiten, I).
#   Die Fuge Dashboard | Cockpit liegt genau auf der Fixierlinie der Jahrestabellen (A:C = 535 px), damit dort
#   keine Form die Fixierlinie kreuzt. Das Navy-Band reicht bis NAV_END + NAV_X (rechts derselbe Innenabstand wie links).
NAV_X = 31
TAB_W, TAB_GAP, TAB_H = 94, 4, 28
NAV_END = 1417
TAB_X0 = NAV_END - 12 * TAB_W - 11 * TAB_GAP          # 245
BAND_END = NAV_END + NAV_X                            # 1448
BAND_WIDE = 1.25                                      # breitere Inhalte bis zu diesem Faktor: Band bis zur Inhaltskante
TAB_SIZE = 9
BRAND_X = NAV_X                            # Kompatibilität
NAV_STD_W = NAV_END - NAV_X
BRAND_GAP = 16                             # Luft zwischen Marke und erstem Reiter
STEP_GAP = 4                               # Schritt-Chips: gleicher Abstand auch vor WEITER
STEP_INS = 38100                           # Innenabstand der Schritt-Chips links/rechts (4 px, R3-P35)
# Fixierte Jahrestabellen (P2-11): Ortsmarke im fixierten Bereich A:C (bleibt beim waagerechten Scrollen stehen)
HERE_LABEL = {"Projektion": "Projektion", "Steuern": "Steuer-Tabelle", "AfA-Vergleich": "AfA-Vergleich",
              "Finanzierung": "Tilgungsplan"}
WEITER_COL = "H"                           # WEITER steht genau über dem unteren Weiter-Button (Merge H:I)
PILL_H, PILL_SIZE, PILL_PAD, PILL_GAP = 22, 8, 10, 4   # Unterreiter und Sprungleisten (ein Pill-Stil)
INS = 25400                                # Innenabstand links/rechts (2 pt)
FONT = "Calibri"
# Verankerung (P1-08): Reiterleiste als xdr:absoluteAnchor (feste EMU-Werte); Schritt-Leiste, Unterreiter und
# Sprungleisten als twoCellAnchor editAs="absolute" (Excel: „Von Zellposition und -größe unabhängig“ – die Formen
# behalten Lage und Größe, wenn Nutzer Spaltenbreiten oder Zeilenhöhen ändern). NAV_GROUPED=0: Einzelformen.
GROUPED = os.environ.get("NAV_GROUPED", "1") != "0"
ABS, FIXED = "abs", "absolute"
# Reiterleiste: Standard twoCellAnchor editAs="absolute" mit exakt aus der festen Pixelgeometrie berechneten Ankern.
# In Excel ist das pixelgleich zu einem absoluteAnchor (fixe Lage/Größe, bewegt sich nicht mit den Zellen); anders
# als ein absoluteAnchor bleibt die Leiste aber auch in Programmen mit abweichender Spaltenmetrik (LibreOffice,
# ≈ 6 % breitere Spalten) bündig zum Zellraster. NAV_ANCHOR=abs erzwingt xdr:absoluteAnchor.
TAB_ANCHOR = ABS if os.environ.get("NAV_ANCHOR") == "abs" else FIXED

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
            bm = re.search(r'<bgColor[^>]*rgb="([0-9A-Fa-f]{6,8})"', f or "")
            if pm and pm.group(1) == "solid" and cm:
                fills.append(cm.group(1)[-6:].upper())
            elif pm and pm.group(1) not in ("none", "solid") and (bm or cm):
                # Musterfüllung (LibreOffice schreibt manche Kopfleisten-Goldzellen als „mediumGray“ mit Gold-Hintergrund):
                # sichtbar ist im Wesentlichen die Hintergrundfarbe – so zählt die Zelle zur Kopfleiste (trim_band)
                fills.append((bm or cm).group(1)[-6:].upper())
            else:
                fills.append(None)
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

    def header_right(self, styles):
        """Rechte Kante des Kopf-Kontextblocks: rechtsbündige Zellen mit Wert in Zeile 6/7 (verbundene Bereiche bis
        zum Bereichsende). None, wenn es keinen rechten Kontextblock gibt."""
        edges = []
        for (r, c), (st, has_val) in self.cells.items():
            if r not in (6, 7) or not has_val or styles.halign(st) != "right":
                continue
            end = next((c1 for mr, c0, c1 in self.merges if mr == r and c0 == c), c)
            edges.append(self.x(end + 1))
        return max(edges) if edges else None

    def hero_box(self, styles, fills=(NAVY,)):
        """Deckblatt-Hero: (links, rechts, erste Zeile) der nachtblauen Zellfläche unter der Kopfleiste – erste Zeile
        ab Z. 5 (Z. 4 ist die von chrome.py vorgelegte Fuge) mit mindestens 800 px breiter Navy-Fläche; sonst None."""
        rows = {}
        for r, c, st, _hv in self.body:
            if 5 <= r <= 8 and self.row_px(r) and self.col_px(c) and styles.fill(st) in fills:
                rows.setdefault(r, set()).add(c)
        for r in sorted(rows):
            left, right = self.x(min(rows[r])), self.x(max(rows[r]) + 1)
            if right - left >= 800:
                return left, right, r
        return None

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
        self.media = {}                    # Paketpfad → Bytes (Icons der Reiterleiste, Runde 6)
        self._grp = None

    def rid(self, sheet, cell=None):
        rid = f"rIdNav{len(self.rels) + 1}"
        target = f"#'{sheet.replace(chr(39), chr(39) * 2)}'!{cell or link_target(sheet)}"
        # interne Ziele ohne TargetMode="External" – so speichert Excel Formen-Links selbst
        self.rels.append(f'<Relationship Id="{rid}" Type="{REL_HYPER}" Target="{xattr(target)}"/>')
        return rid

    def image_rid(self, path):
        """Bild (PNG) als Medium der Zeichnung anmelden → Beziehungs-ID. Gleiche Bilder teilen sich eine Mediendatei
        (Name aus der Prüfsumme), je Zeichnung genügt eine Beziehung."""
        import hashlib
        with open(path, "rb") as fh:
            data = fh.read()
        media = f"xl/media/nav_{hashlib.sha1(data).hexdigest()[:12]}.png"
        rid = f"rIdNavImg{hashlib.sha1(media.encode()).hexdigest()[:8]}"
        if media not in self.media:
            self.media[media] = data
            self.rels.append(f'<Relationship Id="{rid}" Type="{REL_IMAGE}" Target="../media/{posixpath.basename(media)}"/>')
        return rid

    def pic(self, name, x, y, w, h, path, link=None, tooltip=None):
        """Bild (Icon) als xdr:pic – in einer Gruppe wie eine Form; link wie bei add() (klickbar wie der Reiter)."""
        x, y, w, h = int(round(x)), int(round(y)), int(round(w)), int(round(h))
        hl = ""
        if link:
            tip = f' tooltip="{xattr(tooltip)}"' if tooltip else ""
            hl = f'<a:hlinkClick r:id="{self.rid(*link)}"{tip}/>'
        sp = (f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{self.sid()}" name="{xattr(name)}" descr="">{hl}</xdr:cNvPr>'
              f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
              f'<xdr:blipFill><a:blip r:embed="{self.image_rid(path)}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
              f'<xdr:spPr><a:xfrm><a:off x="{x * EMU}" y="{y * EMU}"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
              f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>')
        if self._grp is not None:
            self._grp[1].append(sp)
            self._grp[2].append((x, y, w, h))
        else:
            self.shapes.append(self._wrap(sp, x, y, w, h, False))

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

    def end(self, edit_as=FIXED):
        """Gruppe abschließen. edit_as=ABS: xdr:absoluteAnchor mit festen EMU-Werten (Reiterleiste, P1-08);
        FIXED („absolute“): twoCellAnchor editAs="absolute" – Excel bewegt/skaliert die Leiste nicht mit den Zellen;
        None: bewegt und skaliert mit den Zellen.

        Bei Zellankern wird der Rahmen nach außen auf Spaltengrenzen gerundet (unsichtbare Rahmenform als erstes
        Kind): Anker mit colOff 0 bedeuten in jeder Spaltenmetrik dasselbe, die Leiste bleibt überall proportional.
        Gruppen sind gegen Verschieben, Größenänderung und Auflösen gesperrt (a:grpSpLocks, P3-18)."""
        name, items, boxes = self._grp
        self._grp = None
        if not items:
            return
        if not GROUPED:
            for sp, (x, y, w, h) in zip(items, boxes):
                self.shapes.append(self._wrap(sp, x, y, w, h, edit_as != ABS, None if edit_as == ABS else edit_as))
            return
        x0 = min(b[0] for b in boxes)
        y0 = min(b[1] for b in boxes)
        x1 = max(b[0] + b[2] for b in boxes)
        y1 = max(b[1] + b[3] for b in boxes)
        g = self.geo
        if edit_as != ABS:
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
        body = (f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{self.sid()}" name="{xattr(name)}"/><xdr:cNvGrpSpPr>'
                f'<a:grpSpLocks noGrp="1" noUngrp="1" noRot="1" noMove="1" noResize="1"/></xdr:cNvGrpSpPr>'
                f'</xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm>{off}{ch}</a:xfrm></xdr:grpSpPr>{"".join(items)}</xdr:grpSp>')
        if edit_as == ABS:
            self.shapes.append(self._wrap(body, x0, y0, x1 - x0, y1 - y0, False))
        else:
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
    """Raster der Navigation je Blatt (Runde 4 P1-08/P2-03).

    - Reiterleiste: auf allen Blättern dieselbe feste Geometrie (NAV_X … NAV_END), unabhängig vom Blatt.
    - box[blatt] = (links, rechts): Inhaltskanten ab Zeile 4 (Schritt-Leiste, Sprungleiste).
    - band[blatt]: rechte Kante des Navy-Kopfbands = BAND_END; reicht der Inhalt wenig weiter (Cockpit,
      AfA-Vergleich), bis zur Inhaltskante; auf den breiten Jahrestabellen nur dann über die Jahresspalten, wenn
      das Band dort Seitenkontext trägt (Folgeseiten-Hinweise in Z. 2), sonst endet es bei BAND_END.
    - head[blatt]: rechte Kante des Kopf-Kontextblocks (Z. 6/7 rechtsbündig, „Beispiel: …“ / „Erstellt für …“) –
      daran schließen die Unterreiter in Zeile 5 rechtsbündig an (eine rechte Fluchtlinie im Kopf, P2-03)."""

    def __init__(self, geos, styles, drawings):
        self.box, self.band, self.head = {}, {}, {}
        self.sheets = set(geos)
        self.cover = {}                    # Deckblatt: {blatt: (links, rechts, erste Hero-Zeile)} – s. Geo.hero_box
        for name, g in geos.items():
            left, right = g.content_box(styles, drawings.get(name, ""))
            self.box[name] = [left, right]
            self.band[name] = right if BAND_END < right <= BAND_END * BAND_WIDE else BAND_END
            # Kopfband mit Seitenkontext gefüllt (calc.continuation: „Projektion · Jahre 11–20“ in W2/AG2/AQ2 an
            # jeder Druckseite): dann läuft das Band bis zum letzten Kontexttext (P1-08, Variante „füllen“)
            ctx = [g.x(c + 1) for (r, c), (_st, hv) in g.cells.items() if r == 2 and hv and g.x(c) >= BAND_END]
            if ctx:
                self.band[name] = max(self.band[name], max(ctx))
            self.head[name] = g.header_right(styles) or min(right, NAV_END)
            if name == COVER and styles is not None:
                hb = g.hero_box(styles)
                if hb:
                    self.cover[name] = hb
                    self.band[name] = hb[1]

    def edges(self, name):
        return self.box.get(name, (NAV_X, NAV_END))


def nav_frame(geos, styles=None, drawings=None):
    return Frame(geos, styles, drawings or {})


def tab_positions():
    """Feste Reiterpositionen [(x, Beschriftung, Ziel, QuickInfo)] – auf allen Blättern identisch (P1-08)."""
    seq = [(label, target, tip) for grp in NAV_GROUPS for label, target, tip in grp]
    return [(TAB_X0 + k * (TAB_W + TAB_GAP), label, target, tip) for k, (label, target, tip) in enumerate(seq)]


# Runde 6: feine Linien-Icons (icons.py) vor den Reiterbeschriftungen – ein Motiv je Bereich, 12 px, inaktiv MIST,
# aktiv GOLD. Die Icons sind Bilder in derselben Formgruppe und tragen denselben Link wie der Reiter.
TAB_ICONS = {"Start": "haus", "Leitfaden": "check", "Dashboard": "ziel", "Cockpit": "uhr", "Diagramme": "diagramm",
             "Eingaben": "stift", "Projektion": "trend", "Steuer-Tabelle": "paragraf", "Tilgungsplan": "kalender",
             "Sensitivität": "waage", "Bank": "bank", "Anhang": "dokument"}
NAV_ICONS = os.environ.get("NAV_ICONS", "1") != "0"
ICON_PX, ICON_GAP = 12, 4

# Deckblatt (Runde 6, Cover „Start“): Kopfleiste und Hero bilden EINE nachtblaue Fläche – das Kopfband wird exakt auf
# die Hero-Kanten zugeschnitten (chrome.py legt Z. 3/4 dort navy mit feiner Trennlinie vor), die Reiter stehen als
# „Ghost“-Reiter (ohne Fläche, Schrift MIST, aktiv Weiß fett mit Goldstrich) frei auf dem Nachtblau.
COVER = "Start"
COVER_INSET = 16                           # Marke rückt von der Hero-Kante ein (Innenabstand wie die Hero-Inhalte)


def _icon(name, color):
    """Pfad eines Reiter-Icons (3× Auflösung) oder None, wenn icons.py fehlt/fehlschlägt."""
    try:
        import icons
        return icons.icon_path(name, color, px=ICON_PX, res=3)
    except Exception as exc:                # Icons sind Schmuck – ohne sie bleibt die Leiste vollständig
        print(f"navigation: Icon {name} nicht verfügbar ({exc})")
        return None


def tab_bar(cv, name, frame):
    """Reiterleiste (eine Komponente für alle 28 Blätter, P1-08): Marke „MM HOLDING / Immobilien-Kalkulation“ ab
    NAV_X, zwölf gleich breite Reiter (TAB_W) mit konstantem Abstand, „Anhang“ endet bei NAV_END – als
    xdr:absoluteAnchor, damit die Leiste beim Blattwechsel pixelgenau stehen bleibt.

    Blätter mit fixierten Spalten (P2-11): die Leiste zerfällt an der Fixierlinie in zwei Gruppen (die Fuge
    Dashboard | Cockpit liegt auf der Fixierlinie). Im fixierten Teil steht neben der Marke eine Ortsmarke mit dem
    aktuellen Blatt; zusammen mit Start/Leitfaden/Dashboard bleibt das „Wo bin ich?“ beim Scrollen durch die
    Jahresspalten sichtbar."""
    g = cv.geo
    top, h2 = g.row_px(1), g.row_px(2)
    ty = top + (h2 - TAB_H) / 2
    active = area_of(name)
    pos = tab_positions()
    split = g.split_px()
    cover = name in frame.cover
    brand_x = NAV_X + (COVER_INSET if cover and frame.cover[name][0] >= NAV_X - 2 else 0)
    brand_w = TAB_X0 - BRAND_GAP - brand_x
    here = HERE_LABEL.get(name) if split else None
    if here:
        brand_w = int(text_px("Immobilien-Kalkulation", 8)) + 4
    cv.begin("Reiterleiste")
    cv.add("Marke", brand_x, ty - 3, brand_w, TAB_H + 6, None,
           [para([("MM HOLDING", 10, WHITE, True, 60)], "l", 90000),
            para([("Immobilien-Kalkulation", 8, MIST, False)], "l", 90000)],
           link=("Start", None), tooltip="Zur Startseite", lins=0, rins=0)
    if here:
        hw = TAB_X0 - BRAND_GAP - (NAV_X + brand_w + 12)
        hx = TAB_X0 - BRAND_GAP - hw
        cv.add("Ortsmarke", hx, ty + 3, hw, TAB_H - 6, TAB_IDLE,
               [para([(here, 8, WHITE, True)])], line=SKY, radius=11,
               link=(name, None), tooltip=f"Sie sind hier: {here} · zum Blattanfang", lins=0, rins=0)
    cut = next((k for k, (x, *_r) in enumerate(pos) if split and x + TAB_W > split + 8), None)
    for k, (x, label, target, tip) in enumerate(pos):
        if cut is not None and k == cut:
            cv.end(TAB_ANCHOR)
            cv.begin("Reiterleiste rechts")
        on = label == active
        w = TAB_W
        if cut is not None and k == cut - 1 and x + w > split:
            w = split - x                          # Fixierlinie wenige px im Reiter (AfA-Vergleich): Reiter endet dort
        tw = text_px(label, TAB_SIZE, on)
        ico = None
        if NAV_ICONS and label in TAB_ICONS and tw + ICON_PX + ICON_GAP <= w - 6:
            ico = _icon(TAB_ICONS[label], TAB_MARK if on else TAB_TXT)
        shift = ICON_PX + ICON_GAP if ico else 0
        if cover:
            bg, txt = None, WHITE if on else TAB_TXT
        else:
            bg, txt = TAB_ON if on else TAB_IDLE, TAB_ON_TXT if on else TAB_TXT
        link = (target, None)
        ltip = tip + (" (aktueller Bereich)" if on else "")
        cv.add(f"Reiter {label}", x, ty, w, TAB_H, bg, [para([(label, TAB_SIZE, txt, on)])],
               link=link, tooltip=ltip, lins=shift * EMU, rins=0)
        gx = x + (w - (tw + shift)) / 2               # linke Kante der Einheit Icon + Beschriftung
        if ico:
            cv.pic(f"Reiter {label} Icon", round(gx), round(ty + (TAB_H - ICON_PX) / 2), ICON_PX, ICON_PX, ico,
                   link=link, tooltip=ltip)
        if on:                                 # Goldkante des aktiven Reiters: Unterstrich unter Icon + Beschriftung
            mw = min(w - 12, int(tw + shift) + 12)
            cv.add(f"Reiter {label} aktiv", x + (w - mw) / 2, ty + TAB_H - 5, mw, 2, TAB_MARK, prst="rect")
    cv.end(TAB_ANCHOR)
    return NAV_END


def step_widths(total, labels, min_pad=5):
    """Chipbreiten der Schritt-Leiste: jeder Chip mindestens Name + 2 × min_pad; die übrige Breite hebt zuerst die
    kürzesten Chips auf ein gemeinsames Maß (Wasserstand) – volle Schrittnamen ohne Kürzel (P2-05), möglichst
    gleich breite Chips."""
    need = [int(text_px(lab, 8) + 2 * min_pad + 0.999) for lab in labels]
    if sum(need) >= total:
        return distribute(total, need)
    lo, hi = min(need), total
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if sum(max(n, mid) for n in need) <= total:
            lo = mid
        else:
            hi = mid
    w = [max(n, lo) for n in need]
    rest = total - sum(w)
    order = sorted(range(len(w)), key=lambda i: w[i])
    for i in order[:rest]:
        w[i] += 1
    return w


def step_bar(cv, step, frame):
    """Schritt-Leiste in Zeile 8 über die volle Inhaltsbreite: 12 Chips + „WEITER ›“. WEITER ist genau so breit wie
    der untere Weiter-Button (H:I) und steht bündig darüber; alle Abstände 4 px. Zustände (P3-10): erledigt „✓ 01“
    (ICE, BLUE), aktiv (Navy + Goldstrich), offen (weiß, Rahmen MIST, Schrift MUTED); WEITER im Primär-Button-Look. Chip-Namen = Blatttitel
    (P2-05). Verankerung editAs="absolute" (P1-08)."""
    g = cv.geo
    left, right = frame.edges(cv.sheet)
    y = g.y(8) + 2
    h = max(24, g.row_px(8) - 8)
    x_weiter = g.x(col_index(WEITER_COL))
    if not left + 12 * 60 < x_weiter < right - 120:       # Rückfall bei abweichendem Raster
        x_weiter = right - 300
    widths = step_widths(x_weiter - left - 12 * STEP_GAP, [step_chip(i) for i in range(1, 13)])
    x = left
    cv.begin("Schritt-Leiste")
    for i, w in enumerate(widths, start=1):
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
        tip = step_caption(i) + (" (aktuelle Seite)" if cur else " (erledigt)" if done else "")
        no = f"✓ {step_no(i)}" if done else step_no(i)
        cv.add(f"Schritt {step_no(i)}", x, y, w, h, bg,
               [para([(no, 8, c_no, True)], "ctr", 95000),
                para([(lab, 8, c_lab, False, spc)], "ctr", 95000)],
               link=(step_sheet(i), None), tooltip=tip, line=line, lins=STEP_INS, rins=STEP_INS)
        if cur:
            cv.add(f"Schritt {step_no(i)} aktiv", x, y + h + 2, w, 3, STEP_MARK, prst="rect")
        x += w + STEP_GAP
    nxt_sheet, nxt_caption = step_next(step)
    last = step >= 12
    cv.add("Weiter", x_weiter, y, right - x_weiter, h, WEITER_BG,
           [para([("ZUM DASHBOARD  ›" if last else "WEITER  ›", 9, WHITE, True, 40)], "ctr", 90000),
            para([(nxt_caption, 8, WEITER_SUB, False)], "ctr", 90000)],
           link=(nxt_sheet, None), lins=STEP_INS, rins=STEP_INS, line=WEITER_LINE, line_w=12700,
           tooltip=("Weiter zum Dashboard · Gesamtbewertung" if last else f"Weiter zu {nxt_caption}"))
    cv.end(FIXED)


def pill_w(label, size=PILL_SIZE, pad=PILL_PAD):
    """Breite einer Pill – immer mit fetter Laufweite, damit aktive/inaktive Zustände gleich groß sind."""
    return int(round(text_px(label, size, True) + 2 * pad))


def pill_row(cv, items, y, h=PILL_H, x_left=None, x_right=None, size=PILL_SIZE, pad=PILL_PAD, gap=PILL_GAP,
             active=None, name="Reiter"):
    """Reihe kleiner Reiter (Unterreiter / Sprungleiste) im einheitlichen Pill-Stil: E7EEF7 mit 1D4F8A,
    Chip-Look ICE/Rahmen MIST/Schrift BLUE, aktiv NAVY mit Weiß fett (Runde 5); links- oder rechtsbündig."""
    widths = [pill_w(lab, size, pad) for lab, *_ in items]
    total = sum(widths) + gap * (len(items) - 1)
    x = x_left if x_left is not None else x_right - total
    for (label, target, cell, tip), w in zip(items, widths):
        on = label == active
        cv.add(f"{name} {label}", x, y, w, h, PILL_ON if on else PILL_BG,
               [para([(label, size, WHITE if on else PILL_TXT, on)])], line=None if on else PILL_LINE,
               link=(target, cell), tooltip=tip, radius=3, lins=0, rins=0)
        x += w + gap
    return x


def sub_nav(cv, name, frame):
    """Unterreiter der Gruppenblätter (P2-03): Zeile 5, rechtsbündig an der rechten Kante des Kopf-Kontextblocks
    (Z. 6/7: „Beispiel: …“ / „Erstellt für …“; Bankgespräch: Spalte I), senkrecht mittig in Zeile 5. Links in
    derselben Zeile steht die Brotkrume. Verankerung editAs="absolute"."""
    area = area_of(name)
    grp = [(lab, t) for lab, t in SUBNAV.get(area, ()) if t in frame.sheets or not frame.sheets]
    if not grp or name not in (t for _lab, t in grp):
        return
    g = cv.geo
    right = frame.head.get(name) or frame.edges(name)[1]
    y = g.y(5) + (g.row_px(5) - PILL_H) / 2
    active = next(lab for lab, t in grp if t == name)
    items = [(lab, t, None, f"{lab} öffnen" if t != name else f"{lab} (aktuelle Seite)") for lab, t in grp]
    cv.begin("Unterreiter")
    pill_row(cv, items, y, x_right=right, active=active, name="Unterreiter")
    cv.end(FIXED)


# Rechte Kante der Kopfleisten-Fläche: Frame.band (Inhaltskante je Blatt; core.CONTENT_EDGE nur noch für chrome.py)
BAND_TO = CONTENT_EDGE


def trim_band(sxml, geo, styles, end_px, start_px=None, rows=3):
    """Navy-/Akzentzellen der Kopfleiste (Z. 1–3) rechts von end_px auf das Standardformat zurücksetzen.

    Die Zellfläche endet an der letzten Spaltengrenze ≤ end_px; den Rest ergänzt band_extension als Form.
    Deckblatt (Runde 6): zusätzlich links von start_px, über die Zeilen 1…rows (Kopfleiste + Fuge zum Hero).
    Nur leere Zellen werden angefasst (Werte und Formeln bleiben unberührt)."""
    def fix_row(m):
        row = m.group(0)

        def fix_cell(cm):
            c = col_index(cm.group(2))
            sm = re.search(r'\bs="(\d+)"', cm.group(3))
            st = int(sm.group(1)) if sm else None
            if cm.group(4) != "/>" or styles.fill(st) not in BAND_FILLS:
                return cm.group(0)
            if geo.x(c + 1) <= end_px + 2 and (start_px is None or geo.x(c) >= start_px - 2):
                return cm.group(0)
            r = int(cm.group(1)[len(cm.group(2)):])
            geo.cells[(r, c)] = (0, False)
            attrs = re.sub(r'\s*\bs="\d+"', "", cm.group(3))
            return f'<c r="{cm.group(1)}"{attrs} s="0"/>'
        return re.sub(r'<c r="(([A-Z]+)\d+)"([^>]*?)(/>|>)', fix_cell, row)
    rx = "|".join(str(r) for r in range(1, rows + 1))
    return re.sub(rf'<row r="(?:{rx})"[^>]*[^/]>.*?</row>', fix_row, sxml, flags=re.S)


def band_extension(cv, styles, end_px):
    """Reicht die Navy-Fläche der Kopfleiste (Zeile 2) nicht bis end_px, wird sie mit zwei Flächenformen (Navy Z. 1–2,
    Goldlinie Z. 3) bis end_px verlängert (absolut verankert; wird nicht gedruckt)."""
    g = cv.geo
    navy_cols = [c for (r, c), (st, _) in g.cells.items() if r == 2 and styles.fill(st) == NAVY]
    # Rückfall (Runde 6): ein nach chrome.py angelegtes Blatt ohne Navy-Zellen erhält das Band vollständig als Form
    end = g.x(max(navy_cols) + 1) if navy_cols else 0
    if end >= end_px:
        return
    start = max(0, end - 3) if navy_cols else 0   # 3 px Überlappung: keine Haarfuge zwischen Zelle und Form
    h12 = g.row_px(1) + g.row_px(2)
    cv.begin("Kopfleiste Verlängerung")
    cv.add("Kopfleiste Fläche", start, 0, end_px - start, h12, NAVY, prst="rect")
    cv.add("Kopfleiste Linie", start, h12, end_px - start, g.row_px(3), GOLD, prst="rect")
    cv.end(TAB_ANCHOR)


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
    cv.end(FIXED)


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
            if target.startswith("/"):     # absolutes Ziel (openpyxl-gespeicherte Mappen)
                return target.lstrip("/")
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
                    used.add(posixpath.normpath(posixpath.join(base, t)).lstrip("/"))
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
        if name in frame.cover:            # Deckblatt: Kopfband = Hero-Breite, Fuge Z. 3…Hero-Oberkante inklusive
            left, _right, top = frame.cover[name]
            files[spath] = trim_band(files[spath].decode(), geos[name], styles, end_px, left, max(4, top - 1)).encode()
        else:
            files[spath] = trim_band(files[spath].decode(), geos[name], styles, end_px).encode()
        band_extension(cv, styles, end_px)
        tab_bar(cv, name, frame)
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
        n_links += sum(1 for r in cv.rels if REL_HYPER in r)
        for media, data in cv.media.items():
            files[media] = data

    ct = files["[Content_Types].xml"].decode()
    if 'Extension="png"' not in ct:
        files["[Content_Types].xml"] = ct.replace(
            "<Override ", '<Default Extension="png" ContentType="image/png"/><Override ', 1).encode()
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
