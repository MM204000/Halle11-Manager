"""Gemeinsame Design-Bausteine des Immobilien-Kalkulationstools (Blau-Weiß).

Alle Layout-Module verwenden AUSSCHLIESSLICH diese Tokens und Komponenten, damit jedes Blatt
dieselbe visuelle Sprache spricht. Die Bausteine verändern nur Darstellung (Stil, Zahlenformat,
Zeilenhöhe, Verbund, statische Beschriftungen, reine Anzeigeformeln in leeren Zellen) – niemals
Formeln oder Werte bestehender Rechenzellen.

Komponentensprache (Runde 2, dokumentiert in scratchpad/CORE_API.md):
  section()      Abschnittskopf Ebene 1 / Ebene 2 (ein Stil)          – P1-12
  tile()         KPI-Kachel hell/dunkel, neutraler Wert + Status-Chip  – P1-11 / P1-10
  callout_box()  Einordnungs-Box mit Status-Pill und Statusbalken      – P1-04 / P1-10
  btn()/btn_row() Button-System primär/sekundär/soft/ghost/chip       – P1-13
  sum_row()      Summenstufen Abzug / Zwischensumme / Blockergebnis    – P1-14
  status_*()     Statuslogik (erfüllt/prüfen/kritisch) als Pill/Chip/Kante
  NUMFMT/numfmt  Zahlenformat-Katalog                                   – P2-12 / P3-08
  TYPE_SCALE     Typo-Skala 8/9/10/12,5/16/20/22/30, snap_size()        – P3-11
  link_target()  linke obere Zelle des Scrollbereichs, link_row()       – P1-08
Ältere Bausteine (kpi_tile, callout, button, total, band_l1b, band_form …) bleiben unverändert nutzbar.
"""
import math
import re

from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

# =============================================================================== Farben
NAVY = "0B2A4A"      # Primär: Kopfleiste, Titel, Kachel-Label, Jahreskopf, Schlüsselergebnis
BLUE = "1D4F8A"      # Sekundär: Primär-Button, Links, verknüpfte Werte, L2-Überschriften
ACCENT = "4A86C8"    # Akzent: Linien, Akzentkanten, Schrittnummern
SKY = "9CBBE2"
MIST = "C8D7EB"
TINT = "E7EEF7"      # L1-Band, Ergebniszeile
TINT_XL = "F3F7FC"   # Kachelfläche, Blockergebnis, Info-Zeilen
HEAD = "EEF3FA"      # Spaltenkopf
INK = "1A1D21"
INK2 = "3A3F45"
MUTED = "5B6068"     # Sekundärtext (bis 9 pt immer diese Farbe)
MUTED2 = "8A9099"    # nur Linien, inaktive Felder, Text ab 11 pt
LINE = "E6E8EB"      # Haarlinie Datenzeile
LINE2 = "D5D9DE"     # Trennlinie / inaktiver Rahmen
WHITE = "FFFFFF"
GREEN, AMBER, RED = "1F7A4D", "B54708", "B42318"
GREEN_BG, AMBER_BG, RED_BG = "EAF5EF", "FDF3E7", "FBECEB"
GREEN_LINE, AMBER_LINE, RED_LINE = "B5DCC4", "F1CFA5", "E7B4AD"   # Status-Kanten/-Rahmen auf hellem Grund
NEUTRAL_DASH = "9AA4B1"   # „–“ für entfallende Werte (lesbar, aber zurückgenommen)
INPUT_BG, INPUT_LINE, INPUT_FG = "FFF5D6", "E6CB77", BLUE
INACTIVE_BG, INACTIVE_FG = "F3F4F6", MUTED2

SANS, DISPLAY = "Calibri", "Calibri"

# =============================================================================== Typo-Skala und Zeilenraster
# Erlaubt sind nur 8 · 9 · 10 · 12,5 · 16 · 20 · 22 · 30 pt (P3-11). Keine Zellschrift unter 8 pt.
#   8   T_MICRO  nur Versalien-Labels (Kachel-Label, Spaltenkopf, Eyebrow), Fußzeilen, Kontextzeilen
#   9   T_SMALL  Sekundärtext, Einordnungstext, Ebene-2-Titel mit Fläche, Status-Pill
#   10  T_BODY   Fließtext, Tabellen, Buttons, Callout-Titel
#   12,5 T_H3    Abschnittskopf Ebene 1
#   16  T_H2     Zwischentitel groß (Dashboard-Urteil, Hero-Nebenzahlen)
#   20  T_KPI    Kachelwert
#   22  T_H1     Seitentitel
#   30  T_HERO   Hero (Start)
T_MICRO, T_SMALL, T_BODY, T_H3, T_H2, T_KPI, T_H1, T_HERO = 8, 9, 10, 12.5, 16, 20, 22, 30
TYPE_SCALE = (T_MICRO, T_SMALL, T_BODY, T_H3, T_H2, T_KPI, T_H1, T_HERO)
H_BAND, H_BAND_B, H_HEAD, H_ROW, H_ROW2, H_ROW3 = 24, 20, 20, 18, 30, 42
H_STEP_ROW, H_STEP_ROW2 = 22, 32            # Datenzeilen der Schritt-Seiten (1- / 2-zeilig)
H_TILE_LABEL, H_TILE_VALUE, H_TILE_SUB = 18, 30, 16
H_BUTTON = 25.5                             # Primär/Sekundär-Button (34 px) – P1-13
H_BTN = H_BUTTON
H_PILL = 22.5                               # Link-Pills / Statuschip in Link-Reihen (30 px)
H_CALLOUT_HEAD = 20                         # Kopfzeile der Einordnungs-Box
H_GAP = 12                                  # Leer-/Fugenzeile zwischen Komponenten
GRID_HEIGHTS = (H_ROW, H_ROW2, H_ROW3)      # Zeilenraster 1/2/3 Zeilen; ab 4 Zeilen n × 13 + 8


# Zulässige Zeilenhöhen (pt) – identisch mit lint_pro.RASTER; Abstands-/Fugenzeilen ≤ 12 pt sind frei.
ROW_RASTER = sorted({H_BAND, H_BAND_B, H_BAND_B + 2, H_HEAD, H_ROW, H_ROW2, H_ROW3, H_STEP_ROW, H_STEP_ROW2,
                     H_TILE_LABEL, H_TILE_VALUE, H_TILE_SUB, H_BUTTON, H_PILL, 14, 16, 22, 30}
                    | {n * 13 + 8 for n in range(1, 12)} | {n * 13 + 10 for n in range(1, 12)})


def px_pt(pt):
    """Höhe auf ganze Excel-Pixel runden (0,75 pt), aufwärts."""
    return math.ceil(round(pt / 0.75, 3)) * 0.75


def snap_size(pt):
    """Schriftgröße auf die Typo-Skala abbilden (9,5 → 10, 11 → 10, 12 → 12,5, 14 → 16; < 8 → 8).
    Regel: nächste Stufe; 11 wird bewusst zu 10 (Textschrift), 12/14 zur nächsthöheren Titelstufe."""
    if pt is None:
        return None
    if pt in TYPE_SCALE:
        return pt
    if pt < T_MICRO:
        return T_MICRO
    if pt <= 11:
        return T_SMALL if pt < 9.25 else T_BODY
    for s in TYPE_SCALE:
        if s >= pt:
            return s
    return T_HERO

# =============================================================================== Zahlenformate
# Katalog (P2-12 / P3-08). Regeln:
#   - Ergebnis-/Anzeigezellen zeigen 0 als „–“ (Schlüssel ohne Zusatz); EINGABEFELDER nie „–“ → Schlüssel *_in.
#   - Berechnete Quoten/Renditen immer 1 Nachkommastelle (pct1); 2 Stellen nur, wo fachlich nötig (Makler 3,57 %,
#     Steuersätze 46,26 %) → pct2 / pct2_in.
#   - DSCR, Faktoren, Multiples mit „×“ (dscr = 0,00×; mult1 = 0,0×).
#   - Jahre ganzzahlig mit Einheit: years_n → „1 Jahr“ / „12 Jahre“ (Label ohne „(Jahre)“).
#   - Steuerwirkung in Kennzahlblöcken: tax_effect („Zahlung 709 €“ / „Erstattung 709 €“) oder eur_signed („+709 €“).
NUMFMT = {
    "eur": '#,##0" €";-#,##0" €";"–"',
    "eur_zero": '#,##0" €";-#,##0" €";0" €"',
    "eur_in": '#,##0" €";-#,##0" €";0" €"',
    "eur2": '#,##0.00" €";-#,##0.00" €";"–"',
    "eur2_in": '#,##0.00" €";-#,##0.00" €";0.00" €"',
    "eur_signed": '+#,##0" €";-#,##0" €";"–"',
    "eur_signed_zero": '+#,##0" €";-#,##0" €";0" €"',
    "teur": '#,##0,," T€"',
    "num": '#,##0;-#,##0;"–"',
    "num_in": "#,##0",
    "num2": '#,##0.00;-#,##0.00;"–"',
    "num2_in": "#,##0.00",
    "int": "#,##0",
    "int0": '0;-0;"–"',
    "int_in": "0",
    "pct0": '0 %;-0 %;"–"',
    "pct1": '0.0 %;-0.0 %;"–"',
    "pct1_in": "0.0 %",
    "pct2": '0.00 %;-0.00 %;"–"',
    "pct2_in": "0.00 %",
    "dscr": '0.00"×";-0.00"×";"–"',
    "dscr_plain": "0.00",
    "mult1": '0.0"×"',
    "mult2": '0.00"×"',
    "year": "0",
    "years": '0" Jahre"',
    "years_n": '[=1]0" Jahr";0" Jahre"',
    "years_in": '[=1]0" Jahr";0" Jahre"',
    "date": "DD.MM.YYYY",
    "qm": '#,##0" m²"',
    "eur_qm": '#,##0.00" €/m²"',
    "tax_effect": '"Zahlung "#,##0" €";"Erstattung "#,##0" €";"–"',
    "yesno": '[=1]"ja";[=0]"nein";0',
    "status_dot": '"●  "@',
    "hidden": ";;;",
    "text": "@",
}


def numfmt(kind, entry=False):
    """Zahlenformat aus dem Katalog; entry=True liefert die Eingabefeld-Variante (0 sichtbar, nie „–“)."""
    if entry:
        for k in (f"{kind}_in", {"eur": "eur_in", "pct1": "pct1_in", "pct2": "pct2_in", "num": "num_in",
                                 "num2": "num2_in", "int0": "int_in", "years_n": "years_in", "eur2": "eur2_in"}.get(kind, "")):
            if k in NUMFMT:
                return NUMFMT[k]
    return NUMFMT[kind]


# =============================================================================== KPI-Spezifikation (einzige Quelle)
# rule: "ampel" = grün ab green, gelb ab yellow, sonst rot (Namen aus Konfiguration);
#       "cf"    = rot < 0, gelb wenn Jahr 1 ≥ 0 aber Jahr 2 < 0, sonst grün.
KPI = {
    "GI": dict(label="Gesamtinvestition", fmt=NUMFMT["eur"]),
    "EK": dict(label="Eigenkapitalbedarf inkl. Reserve", fmt=NUMFMT["eur"]),
    "RATE": dict(label="Rate an die Bank / Monat", fmt=NUMFMT["eur"]),
    "CF": dict(label="Cashflow n. St. / Monat (Jahr 1)", fmt=NUMFMT["eur"], rule="cf"),
    "CFV": dict(label="Cashflow v. St. / Monat (Jahr 1)", fmt=NUMFMT["eur"], rule="neg"),
    "BMR": dict(label="Bruttomietrendite", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_BMR_gruen", yellow="Ampel_BMR_gelb"),
    "NMR": dict(label="Nettomietrendite", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_NMR_gruen", yellow="Ampel_NMR_gelb"),
    "DSCR": dict(label="DSCR (Jahr 1)", fmt=NUMFMT["dscr"], rule="ampel", green="Ampel_DSCR_gruen", yellow="Ampel_DSCR_gelb"),
    "EKR": dict(label="EK-Rendite Jahr 1", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_EKR_gruen", yellow="Ampel_EKR_gelb"),
    "IRR": dict(label="IRR n. St.", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_IRR_gruen", yellow="Ampel_IRR_gelb"),
    "FAKTOR": dict(label="Kaufpreisfaktor", fmt=NUMFMT["mult1"]),
    "MULT": dict(label="Eigenkapital-Multiple", fmt=NUMFMT["mult2"]),
    "RF": dict(label="Rechtsform", fmt="@"),
}
CF_YEAR2 = "INDEX(Projektion!$D$37:$AQ$37,2)"
RECHTSFORM_SHORT = 'CHOOSE(Rechtsform_Idx,"Privat","vv-GmbH","GmbH")'

# Link-Ziele (P1-08): immer die linke obere Zelle des Scrollbereichs (unter bzw. rechts der Fixierung).
# Excel scrollt dann so, dass Kopfleiste (fixiert), Seitenkopf und Schritt-Leiste im Bild bleiben.
# Blätter ohne Eintrag: Fixierung A4 → Ziel A4 (gilt auch für S01–S12, Start, Leitfaden, Dashboard …).
SCROLL_ORIGIN = {"Projektion": ("D", 11), "Finanzierung": ("D", 11), "Steuern": ("D", 4), "AfA-Vergleich": ("D", 4)}
LINK_TARGET = {k: f"{c}{r}" for k, (c, r) in SCROLL_ORIGIN.items()}

# Letzte Inhaltsspalte von Blättern, deren Inhalt breiter ist als die Reiterleiste (≈ 1 368 px).
# Einzige Quelle für layouts/chrome.BAND_TO und navigation.BAND_TO (Kopfleistenbreite).
CONTENT_EDGE = {"Steuern": "AQ", "Projektion": "AQ", "Finanzierung": "AQ", "AfA-Vergleich": "Q",
                "Cockpit": "K", "Diagramme": "P", "Sensitivität": "P", "Konfiguration": "G", "Hinweise": "E"}
NAV_MIN_PX = 1368    # rechte Kante der Reiterleiste (Kopfleiste reicht mindestens bis hier)


def link_target(sheet):
    """Sprungziel für Blatt-Links: linke obere Zelle des Scrollbereichs (A4 bzw. D4/D11 bei Spaltenfixierung)."""
    return LINK_TARGET.get(sheet, "A4")


def link_row(sheet, row):
    """Abschnittsanker: Zeile `row` in der ersten Spalte des Scrollbereichs (A bzw. D). Zeilen oberhalb der
    Fixierung werden auf die erste scrollbare Zeile angehoben."""
    c, r0 = SCROLL_ORIGIN.get(sheet, ("A", 4))
    return f"{c}{max(int(row), r0)}"


def link_loc(sheet, cell=None):
    """Vollständiger Hyperlink-Ort „'Blatt'!Zelle“ (Apostrophe im Blattnamen verdoppelt)."""
    return f"'{sheet.replace(chr(39), chr(39) * 2)}'!{cell or link_target(sheet)}"


# =============================================================================== Primitive
def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


NOFILL = PatternFill(fill_type=None)


def side(style, color):
    return Side(style=style, color=color)


def font(size=T_BODY, bold=False, color=INK, name=SANS, italic=False, underline=None):
    return Font(name=name, sz=size, b=bold, i=italic, color=color, u=underline)


def align(h="left", v="center", indent=0, wrap=False):
    """Einzug > 0 nur mit horizontal left/right (Excel ignoriert ihn bei 'general')."""
    if indent and h not in ("left", "right", "distributed"):
        h = "left"
    return Alignment(horizontal=h, vertical=v, indent=indent, wrap_text=wrap, shrink_to_fit=False)


def col(c):
    return c if isinstance(c, int) else column_index_from_string(c)


def L(c):
    return c if isinstance(c, str) else get_column_letter(c)


def rng(c1, r1, c2=None, r2=None):
    c2 = c1 if c2 is None else c2
    r2 = r1 if r2 is None else r2
    return f"{L(c1)}{r1}:{L(c2)}{r2}"


def iter_cells(ws, c1, r1, c2, r2):
    for r in range(r1, r2 + 1):
        for cc in range(col(c1), col(c2) + 1):
            yield ws.cell(r, cc)


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def set_text(cell, text):
    """Statischen Text setzen (auch mit führendem '=' als Text, nie als Formel)."""
    cell.value = text
    if isinstance(text, str):
        cell.data_type = "s"


def rich(parts):
    """Rich-Text aus Teilen [(text, größe, fett, farbe[, kursiv])] – Calibri. Leere Teile entfallen."""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    blocks = []
    for p in parts:
        text, size, bold, color = p[:4]
        it = p[4] if len(p) > 4 else False
        if text:
            blocks.append(TextBlock(InlineFont(rFont=SANS, sz=size, b=bold or None, i=it or None, color=color), text))
    return CellRichText(blocks)


def safe_merge(ws, c1, r1, c2, r2):
    """Verbinden, sofern nicht bereits identisch verbunden; überlappende Verbünde vorher lösen."""
    target = rng(c1, r1, c2, r2)
    if col(c1) == col(c2) and r1 == r2:
        return
    for mr in list(ws.merged_cells.ranges):
        if str(mr) == target:
            return
        if not (mr.max_row < r1 or mr.min_row > r2 or mr.max_col < col(c1) or mr.min_col > col(c2)):
            ws.unmerge_cells(str(mr))
    ws.merge_cells(target)


def style_range(ws, c1, r1, c2, r2, fnt=None, fil=None, brd=None, aln=None, fmt=None):
    for c in iter_cells(ws, c1, r1, c2, r2):
        if fnt is not None:
            c.font = fnt
        if fil is not None:
            c.fill = fil
        if brd is not None:
            c.border = brd
        if aln is not None:
            c.alignment = aln
        if fmt is not None:
            c.number_format = fmt


def set_height(ws, row, pt):
    ws.row_dimensions[row].height = pt


# =============================================================================== Maße
def _dim_for(ws, idx):
    """ColumnDimension einer Spalte – auch wenn sie Teil eines Bereichs (min..max) ist."""
    exact = ws.column_dimensions.get(get_column_letter(idx))
    if exact is not None and (exact.customWidth or (exact.width and exact.width != 13) or exact.hidden):
        return exact
    for d in ws.column_dimensions.values():
        lo, hi = d.min or 0, d.max or 0
        if lo and hi and lo <= idx <= hi and lo != hi:
            return d
    return exact


def col_px(ws, c):
    """Excel-Pixel einer Spalte (Standardschrift Calibri 11 → Ziffernbreite 7 px).
    Bereichsfähig (gruppierte ColumnDimension min..max); ausgeblendete Spalten zählen 0 px."""
    d = _dim_for(ws, col(c))
    if d is not None and d.hidden:
        return 0
    if d is not None and d.width and (d.customWidth or d.width != 13):
        w = d.width
    else:
        w = (ws.sheet_format.defaultColWidth if ws.sheet_format is not None else None) or 8.43
    return int(((256 * w + int(128 / 7)) / 256) * 7)


def span_px(ws, c1, c2):
    return sum(col_px(ws, c) for c in range(col(c1), col(c2) + 1))


def text_px(text, size=T_BODY, bold=False):
    """Näherung der Laufweite von Calibri (px bei 96 dpi). Genauer: text_width() (Carlito-Metrik)."""
    if text is None:
        return 0
    s = str(text)
    em = size * 96 / 72
    narrow = sum(1 for ch in s if ch in "iljtfrI.,:;!'|() ·")
    wide = sum(1 for ch in s if ch in "MWmw@%€ÄÖÜ")
    caps = sum(1 for ch in s if ch.isupper())
    units = len(s) * 0.49 - narrow * 0.22 + wide * 0.25 + caps * 0.08
    return units * em * (1.07 if bold else 1.0)


class _Metric:
    """Calibri-Laufweite über Carlito (metrisch identisch); ohne Schrift/PIL Rückfall auf text_px × 0,93."""
    DIRS = ("/usr/share/fonts/truetype/crosextra", "/usr/share/fonts/truetype/carlito", "/usr/share/fonts")
    FILES = {(False, False): "Carlito-Regular.ttf", (True, False): "Carlito-Bold.ttf",
             (False, True): "Carlito-Italic.ttf", (True, True): "Carlito-BoldItalic.ttf"}

    def __init__(self):
        self.paths, self.cache, self.ok = {}, {}, None

    def _init(self):
        import os
        try:
            from PIL import ImageFont  # noqa: F401
            ok = True
        except Exception:
            ok = False
        for k, fn in self.FILES.items():
            for d in self.DIRS:
                for root, _, files in os.walk(d):
                    if fn in files:
                        self.paths[k] = os.path.join(root, fn)
                        break
                if k in self.paths:
                    break
        self.ok = ok and (False, False) in self.paths

    def width(self, text, size=T_BODY, bold=False, italic=False):
        if text is None or text == "":
            return 0.0
        if self.ok is None:
            self._init()
        if not self.ok:
            return text_px(text, size, bold) * 0.93
        key = (round(size * 4) / 4, bool(bold), bool(italic))
        if key not in self.cache:
            from PIL import ImageFont
            path = self.paths.get(key[1:]) or self.paths.get((key[1], False)) or self.paths[(False, False)]
            self.cache[key] = ImageFont.truetype(path, max(1, int(round(key[0] * 96 / 72 * 4))))
        return self.cache[key].getlength(str(text)) / 4.0


METRIC = _Metric()


def text_width(text, size=T_BODY, bold=False, italic=False):
    """Laufweite in px (Excel 100 %) mit Calibri-Metrik (Carlito) – Grundlage für Umbruch und Passung."""
    return METRIC.width(text, size, bold, italic)


def cell_inner_px(width_px, indent=0):
    """Nutzbare Textbreite einer Zelle: Einzug ≈ 9 px je Stufe, 5 px Innenabstand (wie lint_pro)."""
    return width_px - 9 * (indent or 0) - 5


def fits(text, width_px, size=T_BODY, bold=False, indent=0, reserve=0.95):
    """True, wenn der Text einzeilig mit Reserve (Standard 5 %) in die Breite passt."""
    return text_width(text, size, bold) <= cell_inner_px(width_px, indent) * reserve


def lines_needed_metric(text, width_px, size=T_BODY, bold=False, indent=0, italic=False):
    """Zeilenzahl beim Umbruch wie Excel: Umbruch an Leerzeichen und nach „-“ / „/“; Calibri-Metrik;
    width_px = Breite der Zelle bzw. des Verbunds; Einzug 9 px je Stufe, 5 px Innenabstand, 2 % Reserve."""
    if text is None or text == "":
        return 1
    avail = max(cell_inner_px(width_px, indent) * 0.98, 10)
    n = 0
    for para in str(text).split("\n"):
        tokens = re.findall(r"[^ \-/]*[\-/]|[^ \-/]+ ?| ", para)
        cur = ""
        n += 1
        for t in tokens:
            cand = cur + t
            if text_width(cand.rstrip(), size, bold, italic) <= avail or not cur.strip():
                cur = cand
            else:
                n += 1
                cur = t
    return n


def line_pt(size):
    """Excel-Zeilenhöhe je Textzeile (pt) für Calibri in dieser Größe."""
    return max(1.25 * size, size + 2.5)


def grid_height(pt, base=H_ROW):
    """Nächste Rasterhöhe ≥ pt: base (18) · 30 · 42 · danach n × 13 + 8."""
    for h in (base, H_ROW2, H_ROW3):
        if pt <= h + 0.01:
            return h
    n = 4
    while n * 13 + 8 < pt:
        n += 1
    return n * 13 + 8


def lines_needed(text, width_px, size=T_BODY, bold=False):
    if not text:
        return 1
    total = 0
    for para in str(text).split("\n"):
        w = text_px(para, size, bold)
        total += max(1, math.ceil(w / max(width_px - 10, 20)))
    return total


def _name_value(wb, name, depth=0):
    """Aktueller statischer Wert eines definierten Namens (oder None)."""
    if wb is None or depth > 3:
        return None
    try:
        dn = wb.defined_names.get(name)
        if dn is None:
            for k in list(wb.defined_names.keys()):
                if k.lower() == name.lower():
                    dn = wb.defined_names[k]
                    break
        if dn is None:
            return None
        for title, coord in dn.destinations:
            if title in wb.sheetnames and ":" not in coord:
                v = wb[title][coord.replace("$", "")].value
                return display_text(wb[title][coord.replace("$", "")], depth + 1) if is_formula(v) else v
    except Exception:
        return None
    return None


def display_text(cell, depth=0):
    """Angezeigter Text einer Zelle für Maß-Zwecke (Zeilenhöhe, Passung).
    Statisch → Wert; Rich-Text → Klartext; Formel → Auflösung:
      =Name / =Blatt!A1 / =A1        → aktueller statischer Wert der Quelle,
      =IF(…,"Text A","Text B")      → längstes Textliteral (sicher für alle Zweige),
      ="a"&X&"b"                     → alle Literale + 8 Zeichen je Bezug (obere Schranke)."""
    v = cell.value
    if v is None:
        return None
    if not isinstance(v, str):
        return str(v)
    if not is_formula(v):
        return v
    if depth > 3:
        return None
    ws = cell.parent
    wb = getattr(ws, "parent", None)
    f = v[1:].strip()
    m = re.fullmatch(r"([A-Za-z_][A-Za-z0-9_.]*)", f)
    if m and not re.fullmatch(r"[A-Z]{1,3}\d+", f):
        r = _name_value(wb, m.group(1), depth)
        return None if r is None else str(r)
    m = re.fullmatch(r"(?:'?([^'!]+)'?!)?\$?([A-Z]{1,3})\$?(\d+)", f)
    if m:
        try:
            src = (wb[m.group(1)] if m.group(1) else ws)[f"{m.group(2)}{m.group(3)}"]
            return display_text(src, depth + 1)
        except Exception:
            return None
    lits = [x.replace('""', '"') for x in re.findall(r'"((?:[^"]|"")*)"', f)]
    if not lits:
        return None
    if "&" in f:
        refs = len(re.findall(r"&", f)) + 1 - len(lits)
        return "".join(lits) + "0" * 8 * max(refs, 0)
    return max(lits, key=len)


def fit_row_height(ws, row, c1, c2, base=H_ROW, line_pt=13, pad=8, max_lines=None, metric=False, formulas=False,
                   hints=None):
    """Zeilenhöhe nach umbrechendem Inhalt der Zeile (verbundene Bereiche berücksichtigt).
    Standard wie bisher (Näherung, Formeln ignoriert). metric=True: Calibri-Metrik mit Einzug;
    formulas=True: Anzeigeformeln über display_text() auflösen; hints: {spalte: text} überschreibt."""
    merged = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_row == row == mr.max_row:
            merged[mr.min_col] = mr.max_col
    hints = {col(k): v for k, v in (hints or {}).items()}
    need = 1
    for cc in range(col(c1), col(c2) + 1):
        cell = ws.cell(row, cc)
        if cc in hints:
            text = hints[cc]
        elif cell.value is None or (is_formula(cell.value) and not formulas):
            continue
        else:
            text = display_text(cell) if formulas else cell.value
        if text is None:
            continue
        end = merged.get(cc, cc)
        if not (cell.alignment and cell.alignment.wrap_text) and cc not in hints:
            continue
        sz = cell.font.sz or T_BODY
        if metric:
            n = lines_needed_metric(text, span_px(ws, cc, end), sz, bool(cell.font.b),
                                    (cell.alignment.indent or 0) if cell.alignment else 0, bool(cell.font.i))
        else:
            n = lines_needed(text, span_px(ws, cc, end), sz, bool(cell.font.b))
        need = max(need, n)
    if max_lines:
        need = min(need, max_lines)
    ws.row_dimensions[row].height = base if need == 1 else need * line_pt + pad
    return need


def fit_row(ws, row, c1, c2, base=H_ROW, pad=6, grid=True, max_lines=None, hints=None, min_height=None):
    """Empfohlene Zeilenhöhen-Passung (Runde 2): Calibri-Metrik, Formel-Auflösung, Einzug, Zeilenhöhe je
    Schriftgröße (line_pt). grid=True rastet auf 18 · 30 · 42 · n×13+8 ein. Liefert die Zeilenzahl."""
    merged = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_row == row == mr.max_row:
            merged[mr.min_col] = mr.max_col
    hints = {col(k): v for k, v in (hints or {}).items()}
    need_pt, lines = base, 1
    for cc in range(col(c1), col(c2) + 1):
        cell = ws.cell(row, cc)
        text = hints.get(cc, None)
        if text is None:
            if cell.value is None or not (cell.alignment and cell.alignment.wrap_text):
                continue
            text = display_text(cell)
        if text is None:
            continue
        sz = cell.font.sz or T_BODY
        ind = (cell.alignment.indent or 0) if cell.alignment else 0
        n = lines_needed_metric(text, span_px(ws, cc, merged.get(cc, cc)), sz, bool(cell.font.b), ind,
                                bool(cell.font.i))
        if max_lines:
            n = min(n, max_lines)
        lines = max(lines, n)
        if n > 1:
            need_pt = max(need_pt, n * line_pt(sz) + pad)
    h = grid_height(need_pt, base) if grid else need_pt
    if min_height:
        h = max(h, min_height)
    ws.row_dimensions[row].height = h
    return lines


# =============================================================================== Überschriften
def band_l1(ws, row, c1, c2, title=None, right=None, height=H_BAND):
    """L1 Abschnitt: E7EEF7, Akzentkante links thick, Unterkante thin, 12,5 pt Display fett Navy."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.fill = fill(TINT)
        c.border = Border(left=side("thick", ACCENT) if k == 0 else None, bottom=side("thin", ACCENT))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_H3, True, NAVY, DISPLAY)
    first.alignment = align("left", "center", 1)
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(T_MICRO, False, BLUE)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def band_l1b(ws, row, c1, c2, title=None, right=None, height=H_BAND_B):
    """L1b Tabellenabschnitt: E7EEF7, 9 pt fett Versalien Navy, Akzentkante links, Unterkante Blau."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.fill = fill(TINT)
        c.border = Border(left=side("thick", ACCENT) if k == 0 else None, bottom=side("thin", BLUE))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_SMALL, True, NAVY)
    first.alignment = align("left", "center", 1)
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(T_MICRO, False, BLUE)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def band_form(ws, row, c1, c2, title=None, right=None, height=H_BAND_B + 2, extra=None):
    """Formular-Band (nur Eingaben/Konfiguration): Navy-Vollfläche, 10 pt fett weiß.
    extra: Zusatz hinter dem Titel als Rich-Text (9 pt 9CBBE2), z. B. „· Pflichtangaben“.
    Hinweis Runde 2 (P1-12): Navy-Vollflächen sind Kopfleiste und Kachel-Labels vorbehalten –
    für neue Abschnitte section(level=1, extra=…) verwenden."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = fill(NAVY)
        c.border = Border()
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_BODY, True, WHITE)
    first.alignment = align("left", "center", 1)
    if extra:
        base = title if title is not None else (first.value if isinstance(first.value, str) else "")
        if not is_formula(base):
            first.value = rich([(base, T_BODY, True, WHITE), (f"  {extra}", T_SMALL, False, SKY)])
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(T_MICRO, False, SKY)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def subhead_l2(ws, row, c1, c2, title=None, height=H_HEAD):
    """L2 Unterblock: 8,5→8/9 pt fett Versalien Blau, ohne Füllung, Unterkante thin Akzent."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = NOFILL
        c.border = Border(bottom=side("thin", ACCENT))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_SMALL, True, BLUE)
    first.alignment = align("left", "center", 1)
    set_height(ws, row, height)


def table_head(ws, row, c1, c2, labels=None, height=H_HEAD):
    """Spaltenkopf über volle Tabellenbreite: EEF3FA, 8 pt fett Versalien Blau, Unterkante thin Akzent.
    labels: {spalte: (text, 'left'|'right'|'center')} – Köpfe folgen der Ausrichtung ihrer Werte."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = fill(HEAD)
        c.border = Border(bottom=side("thin", ACCENT))
        c.font = font(T_MICRO, True, BLUE)
        if c.value is not None and c.alignment.horizontal not in ("right", "center"):
            c.alignment = align("left", "center", 1)
    for cc, spec in (labels or {}).items():
        text, h = spec if isinstance(spec, tuple) else (spec, "left")
        c = ws.cell(row, col(cc))
        set_text(c, text)
        c.alignment = align(h, "center", 1 if h in ("left", "right") else 0)
    set_height(ws, row, height)


def hairline(ws, row, c1, c2):
    for c in iter_cells(ws, c1, row, c2, row):
        b = c.border
        c.border = Border(left=None, right=None, top=b.top, bottom=side("hair", LINE))


def total(ws, row, c1, c2, level=2):
    """Summenstufen: 1 Zwischensumme (fett, hair oben), 2 Blockergebnis (F3F7FC, thin Navy oben),
    3 Schlüsselergebnis (E7EEF7, thin oben, double unten, fett Navy)."""
    for c in iter_cells(ws, c1, row, c2, row):
        f = c.font
        if level == 1:
            c.fill = NOFILL
            c.border = Border(top=side("hair", LINE2), bottom=side("hair", LINE))
            c.font = font(f.sz or T_BODY, True, INK)
        elif level == 2:
            c.fill = fill(TINT_XL)
            c.border = Border(top=side("thin", NAVY), bottom=side("hair", LINE))
            c.font = font(f.sz or T_BODY, True, NAVY)
        else:
            c.fill = fill(TINT)
            c.border = Border(top=side("thin", NAVY), bottom=side("double", NAVY))
            c.font = font(f.sz or T_BODY, True, NAVY)
    if level == 3:
        set_height(ws, row, max(ws.row_dimensions[row].height or 0, H_ROW))


def memo(ws, row, c1, c2):
    """Nachrichtliche Zeile: kursiv 9 pt grau, Einzug 1."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.font = font(T_SMALL, False, MUTED, italic=True)
        c.fill = NOFILL


# =============================================================================== Status (Ampel)
# Statusfarben-Disziplin (P1-10): Große Kennzahlen stehen neutral (Navy). Der Status erscheint als Pill/Chip
# („● erfüllt / ● prüfen / ● kritisch“) oder als Kante. Rot als Zahlfarbe nur für echte negative Beträge.
STATUS_WORDS = {"red": "kritisch", "amber": "prüfen", "green": "erfüllt"}
STATUS_COLORS = {"red": (RED, RED_BG, RED_LINE), "amber": (AMBER, AMBER_BG, AMBER_LINE),
                 "green": (GREEN, GREEN_BG, GREEN_LINE)}
STATUS_DOT = "●"


def _abs(v):
    """Zelladresse → absolute Adresse; Namen/Ausdrücke bleiben unverändert."""
    v = v.replace("$", "") if re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+", v or "") else v
    if re.fullmatch(r"[A-Z]{1,3}\d+", v):
        return re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", v)
    return v


def status_conditions(kpi=None, value_ref=None, conditions=None):
    """Statusbedingungen [(Excel-Bedingung, 'red'|'amber'|'green')] in Prüfreihenfolge (erste zutreffende gilt).
    kpi: Schlüssel aus KPI (rule ampel/cf/neg); value_ref: Zelle oder Name mit dem Wert; conditions: eigene Liste."""
    if conditions:
        return list(conditions)
    spec = KPI[kpi]
    rule = spec.get("rule")
    v = _abs(value_ref) if value_ref else None
    if not rule or not v:
        return []
    if rule == "ampel":
        return [(f"AND(ISNUMBER({v}),{v}<{spec['yellow']})", "red"),
                (f"AND(ISNUMBER({v}),{v}<{spec['green']})", "amber"),
                (f"ISNUMBER({v})", "green")]
    if rule == "cf":
        return [(f"{v}<0", "red"), (f"{CF_YEAR2}<0", "amber"), (f"ISNUMBER({v})", "green")]
    if rule == "neg":
        return [(f"{v}<0", "red"), (f"ISNUMBER({v})", "green")]
    return []


def status_formula(conditions, words=None, dot=True, empty=""):
    """Anzeigeformel für das Statuswort: =IF(b1,"●  kritisch",IF(b2,"●  prüfen",…,"")).
    words: {'red': …, 'amber': …, 'green': …} (Standard STATUS_WORDS)."""
    words = {**STATUS_WORDS, **(words or {})}
    expr = '"' + empty.replace('"', '""') + '"'
    for cond, lvl in reversed(list(conditions)):
        w = (f"{STATUS_DOT}  " if dot else "") + words[lvl]
        expr = f'IF({cond},"{w}",{expr})'
    return "=" + expr


def status_cf(ws, ref, conditions, font_color=True, bold=True, fill_bg=False, fill_white=False, border=None,
              border_style="thin", border_color="strong"):
    """Bedingte Formate je Status auf ref.
    font_color: Schrift in Statusfarbe · fill_bg: zarte Statusfläche (nur kleine Flächen!) · fill_white: weiß
    border: None | 'box' | 'left' | 'bottom' | 'top' · border_color: 'strong' (Statusfarbe) | 'line' (helle Linie)."""
    for cond, lvl in conditions:
        fg, bg, ln = STATUS_COLORS[lvl]
        kw = {"stopIfTrue": True}
        if font_color:
            kw["font"] = Font(color=fg, bold=bold)
        if fill_bg:
            kw["fill"] = fill(bg)
        elif fill_white:
            kw["fill"] = fill(WHITE)
        if border:
            sd = side(border_style, fg if border_color == "strong" else ln)
            kw["border"] = Border(**({"left": sd, "right": sd, "top": sd, "bottom": sd} if border == "box"
                                     else {border: sd}))
        ws.conditional_formatting.add(ref, FormulaRule(formula=[cond], **kw))


def status_pill(ws, cell, kpi=None, value_ref=None, conditions=None, formula=None, style="pill", words=None,
                h="right", indent=1):
    """Status als Wort mit Punkt in Statusfarbe (Anzeigeformel in einer LEEREN Zelle).
    style 'pill': 9 pt fett, bei Status weiße Fläche + 1-px-Rahmen in Statusfarbe (Einordnungs-Box, Link-Reihen);
    style 'chip': 8 pt fett, nur Schriftfarbe (Kachel-Kontextzeile, Tabellen-Statusspalte)."""
    conds = status_conditions(kpi, value_ref, conditions) if (kpi or conditions) else []
    if formula is not None:
        cell.value = formula
    elif conds:
        cell.value = status_formula(conds, words)
    cell.number_format = "General"
    cell.font = font(T_SMALL if style == "pill" else T_MICRO, True, MUTED)
    cell.alignment = align(h, "center", indent if h in ("left", "right") else 0)
    if conds:
        status_cf(ws, cell.coordinate, conds, font_color=True, fill_white=(style == "pill"),
                  border="box" if style == "pill" else None)
    return conds


def status_edge(ws, ref, kpi=None, value_ref=None, conditions=None, edge="left", style="thick"):
    """Statuskante: Rahmen einer Seite (Standard: linker 3-px-Balken) in Statusfarbe, per bedingter Formatierung.
    Hinweis: Excel zeigt bedingte Rahmen teils nur dünn – deshalb immer mit statischer Akzentkante kombinieren."""
    conds = status_conditions(kpi, value_ref, conditions) if (kpi or conditions) else []
    if conds:
        status_cf(ws, ref, conds, font_color=False, border=edge, border_style=style)
    return conds


def status_legend(labels=("Ziel erreicht", "knapp", "kritisch"), size=T_MICRO):
    """Rich-Text-Legende „● Ziel erreicht · ● knapp · ● kritisch“ (Punkte in Statusfarbe, Text 5B6068)."""
    parts = []
    for i, (lvl, lab) in enumerate(zip(("green", "amber", "red"), labels)):
        if i:
            parts.append(("   ·   ", size, False, MUTED))
        parts += [(STATUS_DOT + " ", size, True, STATUS_COLORS[lvl][0]), (lab, size, False, MUTED)]
    return rich(parts)


def neg_red(ws, ref, color=RED, bold=None):
    """Negative Beträge rot (nur Ergebnis-/Summenzeilen und Cashflow-Beträge), ≥ 0 bleibt neutral."""
    first = ref.split(":")[0].replace("$", "")
    kw = {"color": color}
    if bold is not None:
        kw["bold"] = bold
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f"AND(ISNUMBER({first}),{first}<0)"], font=Font(**kw),
                                                   stopIfTrue=False))


def cf_close(ws):
    """LibreOffice-Vorschau: Jeder Bereich mit bedingter Formatierung endet mit einer Immer-wahr-Regel
    „nicht durchgestrichen“ (in Excel wirkungslos). Sonst verlieren Zellen ohne zutreffende Regel den Einzug.
    Idempotent; am Ende eines Blatt-Moduls bzw. in global_rules.final aufrufen."""
    refs = []
    for cf in ws.conditional_formatting:
        if not any(r.formula == ["TRUE"] and r.dxf is not None and r.dxf.font is not None and r.dxf.font.strike is False
                   for r in cf.rules):
            refs.append(str(cf.sqref))
    for ref in refs:
        ws.conditional_formatting.add(ref, FormulaRule(formula=["TRUE"], font=Font(strike=False)))


def add_ampel(ws, ref, kpi_key, value_ref=None, with_fill=False):
    """Bedingte Schriftfarbe (und optional Tönung) nach KPI-Spezifikation.
    ref: Zielbereich (z. B. 'C12:D13'); value_ref: absolute Zelle/Name mit dem Wert (Default: erste Zelle von ref).
    Runde 2: für GROSSE Kennzahlen nicht mehr verwenden (Zahl neutral) – stattdessen status_pill/tile(status=…)."""
    spec = KPI[kpi_key]
    rule = spec.get("rule")
    if not rule:
        return
    v = value_ref or ref.split(":")[0].replace("$", "")
    conds = status_conditions(kpi_key, v)
    if rule == "neg":
        conds = conds[:1]
    status_cf(ws, ref, conds, font_color=True, bold=True, fill_bg=with_fill)


# =============================================================================== Komponenten
def kpi_tile(ws, c1, c2, label_row, value_row, sub_row=None, label=None, value=None, sub=None,
             kpi=None, fmt=None, value_rows=1, gap_right=True, value_size=T_KPI):
    """Kachel: Label-Streifen Navy (8 pt fett weiß, Satzschreibung), Wert 20 pt Display fett Navy auf F3F7FC,
    optionale Unterzeile 8 pt grau. value=None lässt einen vorhandenen Wert/Formel unverändert."""
    spec = KPI.get(kpi, {}) if kpi else {}
    last_value_row = value_row + value_rows - 1
    rows = [label_row] + list(range(value_row, last_value_row + 1)) + ([sub_row] if sub_row else [])
    for r in rows:
        for c in iter_cells(ws, c1, r, c2, r):
            c.fill = fill(NAVY if r == label_row else TINT_XL)
            c.border = Border(right=side("thick", WHITE) if (gap_right and col(c.column) == col(c2)) else None)
    safe_merge(ws, c1, label_row, c2, label_row)
    lab = ws.cell(label_row, col(c1))
    if label is not None or spec.get("label"):
        if not is_formula(lab.value) or label is not None:
            set_text(lab, label if label is not None else spec["label"])
    lab.font = font(T_MICRO, True, WHITE)
    lab.alignment = align("left", "center", 1)
    set_height(ws, label_row, H_TILE_LABEL)
    safe_merge(ws, c1, value_row, c2, last_value_row)
    val = ws.cell(value_row, col(c1))
    if value is not None:
        val.value = value
    val.font = font(value_size, True, NAVY, DISPLAY)
    val.alignment = align("left", "center", 1)
    f = fmt or spec.get("fmt")
    if f:
        val.number_format = f
    if value_rows == 1:
        set_height(ws, value_row, H_TILE_VALUE)
    if kpi and spec.get("rule"):
        add_ampel(ws, rng(c1, value_row, c2, last_value_row), kpi)
    if sub_row:
        safe_merge(ws, c1, sub_row, c2, sub_row)
        s = ws.cell(sub_row, col(c1))
        if sub is not None:
            s.value = sub
        s.font = font(T_MICRO, False, MUTED)
        s.alignment = align("left", "center", 1)
        set_height(ws, sub_row, H_TILE_SUB)


def button(ws, c1, row, c2, text, target_sheet=None, kind="primary", target_cell=None, tooltip=None):
    """Primär: 1D4F8A gefüllt, 10 pt fett weiß. Sekundär: weiß, Rahmen thin 1D4F8A, 10 pt fett Blau.
    Textlink: ohne Rahmen, 9,5→10 pt Blau. Jeder Link endet mit ' ›' oder beginnt mit '‹ '."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.border = Border()
        c.fill = NOFILL
    safe_merge(ws, c1, row, c2, row)
    c = ws.cell(row, col(c1))
    set_text(c, text)
    ln = side("thin", BLUE)
    if kind == "primary":
        style_range(ws, c1, row, c2, row, fil=fill(BLUE))
        c.font = font(T_BODY, True, WHITE)
        c.alignment = align("center", "center")
    elif kind == "secondary":
        for cc in iter_cells(ws, c1, row, c2, row):
            cc.border = Border(top=ln, bottom=ln, left=ln if cc.column == col(c1) else None,
                               right=ln if cc.column == col(c2) else None)
        c.font = font(T_BODY, True, BLUE)
        c.alignment = align("center", "center")
    else:
        c.font = font(T_BODY, False, BLUE)
        c.alignment = align("left", "center", 1)
    if target_sheet:
        loc = f"'{target_sheet}'!{target_cell or link_target(target_sheet)}"
        c.hyperlink = Hyperlink(ref=c.coordinate, location=loc, display=text, tooltip=tooltip)
    set_height(ws, row, H_BUTTON if kind != "link" else max(ws.row_dimensions[row].height or 0, H_ROW))


def text_link(cell, text, target_sheet, target_cell=None, size=T_BODY, bold=True, tooltip=None):
    set_text(cell, text)
    cell.font = font(size, bold, BLUE)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{target_sheet}'!{target_cell or link_target(target_sheet)}",
                               display=text, tooltip=tooltip)


def callout(ws, c1, head_row, c2, body_r1, body_r2, title="Einordnung", status=None, status_col=None):
    """Einordnung: ein Akzentbalken (thick 4A86C8) für Kopf und Text; Kopf 9 pt fett Navy,
    Text 9,5→10 pt INK2 auf F3F7FC, umbrechend, vertikal zentriert."""
    for r in range(head_row, body_r2 + 1):
        for c in iter_cells(ws, c1, r, c2, r):
            c.fill = fill(TINT_XL)
            c.border = Border(left=side("thick", ACCENT) if c.column == col(c1) else None,
                              bottom=side("thin", ACCENT) if r == head_row else None)
    h = ws.cell(head_row, col(c1))
    if title is not None:
        set_text(h, title)
    h.font = font(T_SMALL, True, NAVY)
    h.alignment = align("left", "center", 1)
    set_height(ws, head_row, H_HEAD)
    if status is not None:
        sc = ws.cell(head_row, col(status_col or c2))
        sc.value = status
        sc.font = font(T_MICRO, True, MUTED)
        sc.alignment = align("right", "center", 1)
    safe_merge(ws, c1, body_r1, c2, body_r2)
    b = ws.cell(body_r1, col(c1))
    b.font = font(T_SMALL, False, INK2)
    b.alignment = align("left", "center", 1, wrap=True)


def input_style(cell, state="required"):
    """Eingabezustände: required (gelb, Rahmen thin), optional (gelb, Rahmen dashed),
    inactive (grau), linked (weiß, dashed Akzent, Blau normal)."""
    if state in ("required", "optional"):
        ln = side("thin" if state == "required" else "dashed", INPUT_LINE)
        cell.fill = fill(INPUT_BG)
        cell.font = font(cell.font.sz or T_BODY, True, INPUT_FG)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)
        cell.protection = Protection(locked=False)
    elif state == "inactive":
        ln = side("thin", LINE2)
        cell.fill = fill(INACTIVE_BG)
        cell.font = font(cell.font.sz or T_BODY, False, INACTIVE_FG)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)
    elif state == "linked":
        ln = side("dashed", ACCENT)
        cell.fill = NOFILL
        cell.font = font(cell.font.sz or T_BODY, False, BLUE)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)


def inactive_when(ws, ref, condition):
    """Eingabefeld bedingt inaktiv darstellen (bleibt editierbar)."""
    ws.conditional_formatting.add(ref, FormulaRule(formula=[condition], stopIfTrue=True,
                                                   font=Font(color=INACTIVE_FG, bold=False),
                                                   fill=fill(INACTIVE_BG),
                                                   border=Border(left=side("thin", LINE2), right=side("thin", LINE2),
                                                                 top=side("thin", LINE2), bottom=side("thin", LINE2))))


def page_header(ws, c1, c2, eyebrow, title, subtitle=None, context=None, context_col=None):
    """Seitenkopf Z. 5–7: Eyebrow (8 pt fett Versalien Blau), H1 (22 pt Display fett Navy), Untertitel (10 pt grau).
    context: (formel_name, formel_adresse) rechts oben in context_col."""
    e, t, s = ws.cell(5, col(c1)), ws.cell(6, col(c1)), ws.cell(7, col(c1))
    set_text(e, eyebrow.upper())
    e.font = font(T_MICRO, True, BLUE)
    e.alignment = align("left", "bottom")
    if title is not None:
        if not is_formula(t.value) or not is_formula(title):
            t.value = title
    t.font = font(T_H1, True, NAVY, DISPLAY)
    t.alignment = align("left", "center")
    set_height(ws, 5, 16)
    set_height(ws, 6, 30)
    if subtitle is not None:
        s.value = subtitle
    if s.value is not None:
        s.font = font(T_BODY, False, MUTED)
        s.alignment = align("left", "top")
        set_height(ws, 7, 22)
    if context:
        cc = col(context_col or c2)
        n, a_ = ws.cell(6, cc), ws.cell(7, cc)
        n.value, a_.value = context
        n.font = font(T_BODY, True, NAVY)
        a_.font = font(T_SMALL, False, MUTED)
        n.alignment = align("right", "center", 1)
        a_.alignment = align("right", "top", 1)


FOOTER_1 = "Keine Gewähr für die Richtigkeit der Angaben · ersetzt keine Rechts-, Steuer- oder Finanzberatung · Rechtsstand September 2026"
FOOTER_2 = "MM Holding GmbH · Kornhausgasse 4 · 88250 Weingarten · Amtsgericht Ulm HRB 750729"


def footer(ws, row, c1, c2):
    """Fuß: Oberkante thin LINE2, zwei Zeilen 8 pt grau."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.border = Border(top=side("thin", LINE2))
        c.fill = NOFILL
    a1, a2 = ws.cell(row, col(c1)), ws.cell(row + 1, col(c1))
    set_text(a1, FOOTER_1)
    set_text(a2, FOOTER_2)
    for c in (a1, a2):
        c.font = font(T_MICRO, False, MUTED)
        c.alignment = align("left", "center")
    set_height(ws, row, 18)
    set_height(ws, row + 1, 14)


def hide_rows(ws, r1, r2):
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].hidden = True


def hide_cols(ws, c1, c2):
    for cc in range(col(c1), col(c2) + 1):
        ws.column_dimensions[L(cc)].hidden = True


# =============================================================================== Komponentensprache Runde 2
# Jede Komponente existiert genau einmal. Blatt-Module setzen nur noch Parameter, keine eigenen Varianten.

def _caps(text, caps):
    return text.upper() if (caps and isinstance(text, str) and not is_formula(text)) else text


def section(ws, row, c1, c2, title=None, level=1, meta=None, extra=None, variant="fill", labels=None, height=None,
            caps=True):
    """Abschnittskopf – EIN Stil für alle Blätter (P1-12).

    Ebene 1 (Abschnitt): Fläche E7EEF7, linke Akzentkante 3 px 4A86C8, Unterkante thin 4A86C8,
      Titel 12,5 pt fett 0B2A4A in Titelschreibung, Zeilenhöhe 24; meta rechts 8 pt 1D4F8A;
      extra: Zusatz hinter dem Titel als Rich-Text 9 pt 5B6068 (z. B. „· Pflichtangaben“).
    Ebene 2 (Unterabschnitt / Tabellenkopf): 8 pt fett VERSALIEN 1D4F8A, Zeilenhöhe 20,
      variant 'fill' = Fläche EEF3FA + Unterlinie 4A86C8 (Tabellenkopf; labels wie table_head),
      variant 'line' = ohne Fläche, nur Unterlinie 4A86C8 (Unterabschnitt in ruhigen Blättern).
    Navy-Vollflächen sind der Kopfleiste und den Kachel-Labels vorbehalten."""
    if level == 1:
        band_l1(ws, row, c1, c2, title, right=meta, height=height or H_BAND)
        first = ws.cell(row, col(c1))
        if extra and isinstance(first.value, str) and not is_formula(first.value):
            first.value = rich([(first.value, T_H3, True, NAVY), (f"  {extra}", T_SMALL, False, MUTED)])
        return first
    h = height or H_HEAD
    if variant == "fill":
        table_head(ws, row, c1, c2, labels, height=h)
    else:
        for c in iter_cells(ws, c1, row, c2, row):
            c.fill = NOFILL
            c.border = Border(bottom=side("thin", ACCENT))
            c.font = font(T_MICRO, True, BLUE)
        set_height(ws, row, h)
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, _caps(title, caps))
    elif caps and isinstance(first.value, str):
        first.value = _caps(first.value, caps)
    if caps:
        for cc in (labels or {}):
            c = ws.cell(row, col(cc))
            if isinstance(c.value, str):
                c.value = _caps(c.value, True)
    first.font = font(T_MICRO, True, BLUE)
    if first.alignment.horizontal not in ("right", "center"):
        first.alignment = align("left", "center", 1)
    if meta is not None:
        last = ws.cell(row, col(c2))
        set_text(last, meta)
        last.font = font(T_MICRO, False, BLUE)
        last.alignment = align("right", "center", 1)
    return first


def tile(ws, c1, c2, label_row, value_row, sub_row=None, label=None, value=None, sub=None, kpi=None, fmt=None,
         variant="dark", value_rows=1, value_span=None, sub_right=None, sub_right_fmt=None, status="auto",
         status_col=None, value_ref=None, conditions=None, words=None, neg=None, gap_right=True, gap_top=False,
         caps=True, value_size=T_KPI):
    """KPI-Kachel – EINE Komponente für die ganze Mappe (P1-11, P1-10).

    Aufbau (immer dreizeilig empfohlen): Label 18 pt · Wert 30 pt · Kontextzeile 16 pt.
      variant 'dark' : Label weiß 8 pt fett VERSALIEN auf 0B2A4A, Wert-/Kontextfläche F3F7FC (Fintech: Start,
                       Dashboard, Cockpit-Kopf, S08/S12)
      variant 'light': alles F3F7FC, Label 1D4F8A, Oberkante medium 4A86C8 (Private Banking: Rechenblätter,
                       AfA-Vergleich, Leitfaden-Nebenkacheln)
    Wert: 20 pt fett 0B2A4A, IMMER neutral. neg=True (Standard bei €-Beträgen und KPI-Regel cf/neg): Beträge < 0 rot.
    value_span: letzte Spalte des Wertfelds; die Spalten dahinter nehmen sub_right auf (8 pt 5B6068, rechtsbündig –
      Kontext im Wertfeld, z. B. „ab Jahr 2: –274 €“).
    status: 'auto' | 'chip' | 'dot' | 'edge' | None
      chip  – Statuswort „● prüfen“ 8 pt fett in Statusfarbe rechts in der Kontextzeile (status_col, Standard c2;
              Kontexttext dann über c1…status_col-1) – Standard bei Kacheln ≥ 2 Spalten
      dot   – Kontextzeile selbst in Statusfarbe (für Kontexttexte wie „unter Ziel 6,0 %“)
      edge  – linke Kante (3 px) von Wert + Kontext in Statusfarbe – Standard bei 1-spaltigen Kacheln
    Der Status gilt nur für KPI mit Regel (ampel/cf/neg) oder eigene conditions. value=None lässt eine
    vorhandene Formel stehen. gap_right: weiße 3-px-Rinne rechts; gap_top: weiße Kante oben (Reihenabstand)."""
    spec = KPI.get(kpi, {}) if kpi else {}
    c1i, c2i = col(c1), col(c2)
    last_v = value_row + value_rows - 1
    dark = variant != "light"
    rows = [label_row] + list(range(value_row, last_v + 1)) + ([sub_row] if sub_row else [])
    white3 = side("thick", WHITE)
    for r in rows:
        for c in iter_cells(ws, c1i, r, c2i, r):
            c.fill = fill(NAVY if (dark and r == label_row) else TINT_XL)
            top = None
            if r == label_row:
                top = white3 if (gap_top and dark) else (side("medium", ACCENT) if not dark else None)
            c.border = Border(right=white3 if (gap_right and c.column == c2i) else None, top=top)
    # Label
    safe_merge(ws, c1i, label_row, c2i, label_row)
    lab = ws.cell(label_row, c1i)
    if label is not None:
        set_text(lab, _caps(label, caps))
    elif spec.get("label") and not is_formula(lab.value):
        set_text(lab, _caps(spec["label"], caps))
    elif caps and isinstance(lab.value, str) and not is_formula(lab.value):
        lab.value = lab.value.upper()
    lab.font = font(T_MICRO, True, WHITE if dark else BLUE)
    lab.alignment = align("left", "center", 1)
    set_height(ws, label_row, H_TILE_LABEL)
    # Wert
    vs_end = col(value_span) if value_span else c2i
    safe_merge(ws, c1i, value_row, vs_end, last_v)
    val = ws.cell(value_row, c1i)
    if value is not None:
        val.value = value
    val.font = font(value_size, True, NAVY, DISPLAY)
    val.alignment = align("left", "center", 1)
    f = fmt or spec.get("fmt")
    if f:
        val.number_format = f
    if value_rows == 1:
        set_height(ws, value_row, H_TILE_VALUE)
    if vs_end < c2i:
        safe_merge(ws, vs_end + 1, value_row, c2i, last_v)
        sr = ws.cell(value_row, vs_end + 1)
        if sub_right is not None:
            sr.value = sub_right
        sr.font = font(T_MICRO, False, MUTED)
        sr.alignment = align("right", "center", 1, wrap=False)
        if sub_right_fmt:
            sr.number_format = sub_right_fmt
    # Status
    vref = value_ref or f"${get_column_letter(c1i)}${value_row}"
    conds = status_conditions(kpi, vref, conditions) if (conditions or spec.get("rule")) else []
    mode = status if conds else None
    if mode == "auto":
        mode = "chip" if (sub_row and (status_col or c2i > c1i)) else "edge"
    if mode == "chip" and not sub_row:
        mode = "edge"
    # Kontextzeile
    if sub_row:
        chip_c = (col(status_col) if status_col else c2i) if mode == "chip" else None
        end = chip_c - 1 if (chip_c and chip_c > c1i) else c2i
        safe_merge(ws, c1i, sub_row, end, sub_row)
        s = ws.cell(sub_row, c1i)
        if sub is not None:
            s.value = sub
        s.font = font(T_MICRO, False, MUTED)
        s.alignment = align("left", "center", 1)
        set_height(ws, sub_row, H_TILE_SUB)
        if chip_c and chip_c > c1i:
            if chip_c < c2i:
                safe_merge(ws, chip_c, sub_row, c2i, sub_row)
            status_pill(ws, ws.cell(sub_row, chip_c), conditions=conds, style="chip", words=words)
        if mode == "dot":
            status_cf(ws, s.coordinate, conds, font_color=True, bold=True)
    if mode == "edge":
        status_edge(ws, rng(c1i, value_row, c1i, sub_row or last_v), conditions=conds)
    # Negative Beträge
    if neg is None:
        neg = spec.get("rule") in ("cf", "neg") or ("€" in (f or "") and "%" not in (f or ""))
    if neg:
        neg_red(ws, rng(c1i, value_row, vs_end, last_v))
    return val


def callout_box(ws, c1, head_row, c2, body_r1, body_r2, title="Einordnung", text=None, kpi=None, value_ref=None,
                conditions=None, pill=None, pill_col=None, words=None, head=True, fit="auto", height_text=None):
    """Einordnungs-Box – EINE Komponente für S01–S12, Sensitivität, Bank … (P1-04).

    Kopf (20 pt): Fläche F3F7FC über c1…c2, Titel 10 pt fett 0B2A4A links (Einzug 1), Unterlinie thin C8D7EB,
      rechts Status-Pill (9 pt fett, bei Status weiß mit 1-px-Rahmen in Statusfarbe) in pill_col (Standard c2).
    Körper: immer F3F7FC, Text 9 pt 3A3F45, OBEN ausgerichtet, Einzug 1, umbrechend; Verbund c1…c2 × body_r1…body_r2.
    Status: nur der linke Balken (3 px) wechselt von 4A86C8 in die Statusfarbe – keine Flächentönung.
      kpi/value_ref (KPI-Regel) oder conditions [(Bedingung, 'red'|'amber'|'green')]; pill: eigene Formel für
      das Pill-Wort (Standard „● erfüllt/prüfen/kritisch“), pill=False unterdrückt die Pill.
    fit: Körperhöhe nach Textlänge (Calibri-Metrik; höchstens ≈ 12 px Luft unten).
      'auto' = einzeiliger Körper exakt, mehrzeiliger Körper: nur die letzte Zeile wachsen lassen · 'exact' ·
      None = Höhen nicht ändern. height_text: Text für die Bemessung (Standard: längster Formelzweig).
    Liefert die Zeilenzahl des Körpertexts. Abstand zum Folgeblock: genau eine Leerzeile H_GAP (12 pt)."""
    c1i, c2i = col(c1), col(c2)
    first = head_row if head else body_r1
    bar = side("thick", ACCENT)
    for r in range(first, body_r2 + 1):
        for c in iter_cells(ws, c1i, r, c2i, r):
            c.fill = fill(TINT_XL)
            c.border = Border(left=bar if c.column == c1i else None,
                              bottom=side("thin", MIST) if (head and r == head_row) else None)
    conds = status_conditions(kpi, value_ref, conditions) if (kpi or conditions) else []
    if head:
        pc = col(pill_col) if pill_col else c2i
        use_pill = pill is not False and (conds or pill) and pc > c1i
        t_end = pc - 1 if use_pill else c2i
        safe_merge(ws, c1i, head_row, t_end, head_row)
        h = ws.cell(head_row, c1i)
        if title is not None:
            set_text(h, title)
        h.font = font(T_BODY, True, NAVY)
        h.alignment = align("left", "center", 1)
        set_height(ws, head_row, H_CALLOUT_HEAD)
        if use_pill:
            if pc < c2i:
                safe_merge(ws, pc, head_row, c2i, head_row)
            status_pill(ws, ws.cell(head_row, pc), conditions=conds or None,
                        formula=pill if isinstance(pill, str) else None, style="pill", words=words)
    safe_merge(ws, c1i, body_r1, c2i, body_r2)
    b = ws.cell(body_r1, c1i)
    if text is not None:
        b.value = text
    b.font = font(T_SMALL, False, INK2)
    b.alignment = align("left", "top", 1, wrap=True)
    if conds:
        status_edge(ws, rng(c1i, first, c1i, body_r2), conditions=conds)
    n = 1
    if fit:
        t = height_text if height_text is not None else display_text(b)
        n = lines_needed_metric(t, span_px(ws, c1i, c2i), T_SMALL, False, 1) if t else 1
        need = px_pt(n * line_pt(T_SMALL) + 8)   # 3 px oben, ≤ 12 px unten
        if body_r1 == body_r2 or fit == "exact":
            others = sum((ws.row_dimensions[r].height or 15) for r in range(body_r1, body_r2))
            set_height(ws, body_r2, max(need - others, H_ROW if body_r1 == body_r2 else 6))
        else:
            cur = sum((ws.row_dimensions[r].height or 15) for r in range(body_r1, body_r2 + 1))
            if need > cur:
                set_height(ws, body_r2, (ws.row_dimensions[body_r2].height or 15) + need - cur)
    return n


BTN = {  # kind: (Fläche, Rahmen, Schrift, fett, Höhe)
    "primary": (BLUE, BLUE, WHITE, True, H_BTN),         # die eine Hauptaktion (Weiter ›, Schritt 01 starten ›)
    "secondary": (WHITE, BLUE, BLUE, True, H_BTN),       # Nebenaktion (‹ Zurück, Übersicht, Zum Dashboard)
    "soft": (TINT_XL, MIST, BLUE, False, H_PILL),        # Link-Pill in Link-Reihen (Dashboard Z. 9)
    "chip": (TINT, MIST, NAVY, True, H_PILL),            # Statuschip in Link-Reihen (Rechtsform: Privat ›)
    "ghost": (BLUE, ACCENT, WHITE, True, H_BTN),         # Sekundär auf Navy (Start-Hero)
}


def btn(ws, c1, row, c2, text, target_sheet=None, kind="primary", target_cell=None, tooltip=None, height=None,
        size=T_BODY, set_row=True):
    """Button-System (P1-13): primary | secondary | soft | chip | ghost | link.
    Alle Buttons: zentriert, 10 pt, Rahmen 1 px, Höhe 25,5 pt (primary/secondary/ghost) bzw. 22,5 pt (soft/chip).
    Text: Vorwärts endet mit „  ›“, Rückwärts beginnt mit „‹  “. Nackte Textlinks nie in einer Button-Reihe.
    target_sheet: Zellhyperlink auf link_target(Blatt) bzw. target_cell."""
    if kind == "link":
        button(ws, c1, row, c2, text, target_sheet, "link", target_cell, tooltip)
        return ws.cell(row, col(c1))
    bg, ln_c, fg, bold, h = BTN[kind]
    ln = side("thin", ln_c)
    safe_merge(ws, c1, row, c2, row)
    c1i, c2i = col(c1), col(c2)
    for cc in iter_cells(ws, c1i, row, c2i, row):
        cc.fill = fill(bg)
        cc.border = Border(top=ln, bottom=ln, left=ln if cc.column == c1i else None,
                           right=ln if cc.column == c2i else None)
    c = ws.cell(row, c1i)
    set_text(c, text)
    c.font = font(size, bold, fg)
    c.alignment = align("center", "center")
    if target_sheet:
        c.hyperlink = Hyperlink(ref=c.coordinate, location=link_loc(target_sheet, target_cell), display=text,
                                tooltip=tooltip)
    if set_row:
        set_height(ws, row, height or h)
    return c


def btn_row(ws, row, items, height=None):
    """Button-Reihe mit einheitlicher Höhe. items: [dict(c1=…, c2=…, text=…, target=…, kind=…, tooltip=…, cell=…)].
    Höhe = größte Höhe der verwendeten Arten (oder height). Liefert [(text, breite_px, passt)] zur Kontrolle –
    gleiche Breiten entstehen über gleiche Spaltenraster (z. B. C | E:F | H:I auf den Schritt-Seiten)."""
    hs, out = [], []
    for it in items:
        kind = it.get("kind", "secondary")
        btn(ws, it["c1"], row, it["c2"], it["text"], it.get("target"), kind, it.get("cell"), it.get("tooltip"),
            set_row=False)
        hs.append(BTN.get(kind, (None, None, None, None, H_BTN))[4])
        w = span_px(ws, it["c1"], it["c2"])
        out.append((it["text"], w, fits(it["text"], w, T_BODY, True, 0, 0.9)))
    set_height(ws, row, height or max(hs or [H_BTN]))
    return out


def sum_row(ws, row, c1, c2, stage="sub", value_from=None, neg=False, keep_size=True):
    """Summenstufen (P1-14) – Zeilenbeschriftung trägt das Rechenzeichen („– Zinsen“, „= Cashflow …“).
      'deduct' (0): Abzugs-/Teilzeile „–“: nur fett, keine Fläche, keine Zusatzlinie
      'sub'    (1): Zwischensumme „=“: fett 0B2A4A, Fläche F3F7FC, Oberlinie thin 1D4F8A
      'result' (2): das EINE Blockergebnis: fett 0B2A4A, Fläche E7EEF7, Oberlinie thin + Doppellinie unten 0B2A4A
    neg=True: negative Beträge in value_from…c2 (Standard: ab zweiter Spalte) rot – nur echte negative Ergebnisse."""
    stage = {0: "deduct", 1: "sub", 2: "result"}.get(stage, stage)
    for c in iter_cells(ws, c1, row, c2, row):
        sz = (c.font.sz if keep_size else None) or T_BODY
        if stage == "deduct":
            c.font = font(sz, True, c.font.color.rgb[-6:] if (c.font.color is not None and isinstance(c.font.color.rgb, str)) else INK)
            c.fill = NOFILL
        elif stage == "sub":
            c.fill = fill(TINT_XL)
            c.border = Border(top=side("thin", BLUE), bottom=side("hair", LINE))
            c.font = font(sz, True, NAVY)
        else:
            c.fill = fill(TINT)
            c.border = Border(top=side("thin", NAVY), bottom=side("double", NAVY))
            c.font = font(sz, True, NAVY)
    if stage == "result":
        set_height(ws, row, max(ws.row_dimensions[row].height or 0, H_ROW))
    if neg:
        vf = col(value_from) if value_from else col(c1) + 1
        if vf <= col(c2):
            neg_red(ws, rng(vf, row, c2, row))
