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

Komponentensprache (Runde 2, CORE_API.md): Raster = core.ROW_RASTER, Typo-Skala = core.TYPE_SCALE, Button-Arten =
core.BTN. Statusfarben-Disziplin (P1-10): große Kennzahlen neutral (kpi_statusfarbe), keine Statusflächen über
Matrizen (status_flaeche), keine Emoji (emoji); kpi_ampel akzeptiert Pill/Chip/Kante/Statusspalte im Umfeld der Zahl.
Callout-Körper (core.callout_box, Höhe nach Textlänge) sind vom Zeilenraster ausgenommen.
Runde 3 (review3): P11 „Wertfarbe = Status“ – eine per Regel gefärbte Kachelzahl ist nur dann ein Befund
(kpi_statusfarbe), wenn die Kachel kein Statuswort „● …“ zeigt (Chip/inline aus core.tile). Neu: nav_rueckweg (P05,
jedes sichtbare Blatt hat einen Zell-Link auf ein anderes Blatt), formel_text (P16, Formel + „@“), minus_typo (P23,
Bindestrich statt „−“ in Zellen, bedingten Formaten und Diagrammachsen; je Blatt und Format gebündelt), neg_rot (P15,
negative Beträge in core.sum_row-Zwischen-/Endsummen ohne Rot; Steuerwirkungszeilen „Zahlung/Erstattung“ ausgenommen).
Runde 4 (review4): eingabe_gelb (P1-04, Gelb FFF5D6/Eingaberahmen E6CB77 nie auf Formel- oder Link-Zellen; erlaubt sind
eine einzelne E6CB77-Kante und der Zustand „überschreibbar“ FFF9EA), dropdown_zeichen (▾ nur an Datenüberprüfung oder
als Hinweis direkt darunter), link_tooltip (P3-17, jeder Zell-Link hat einen ScreenTip), druck_umbruch (fitToPage und
manuelle Umbrüche schließen sich aus), minus_text (Bindestrich-Minus in Textverkettungen). button_hoehe akzeptiert in
gemischten Reihen die größte Soll-Höhe; kpi_statusfarbe erkennt „●“ und „▲“ als Statuswort.
Runde 5 (Farbdesign „Midnight & Gold“, DECISIONS „Nutzerentscheidung Runde 5“): altfarbe (kein int7-Hexwert aus
core.OLD_TO_NEW in Zellflächen/-schriften/-rahmen, bedingten Formaten, Datenbalken, Formen und Registerfarben; je Blatt,
Farbe und Ort gebündelt), gold_text (Gold nie als lesbarer Text unter 16 pt oder auf hellem Grund – Farbmuster „■“ und
Linien sind erlaubt, Text in Bronze C.GOLD_INK), kontrast (WCAG über C.contrast: Fließtext ≥ 4,5:1, ab 18 pt bzw.
14 pt fett ≥ 3:1; reine Symbol-/Legendenzeichen und der inaktive Zustand MUTED2 auf INACTIVE_BG ausgenommen; je Zeile
gebündelt), diagramm_altfarbe, diagramm_einfarbig (≥ 3 sichtbare Serien/Kreissegmente in einer Farbe; Hilfsreihen nach
core._CHART_HELPER zählen nicht) und diagramm_textfarbe (Werte-/Achsentexte nie in einer CHART_CAT-Serienfarbe).
Selbsttest (python excel/lint_pro.py --selftest) prüft jede Regel und als Gegenprobe, dass core.tile/callout_box/btn
keinen Befund auslösen.

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
    "grau_klein":      (FEHLER, f"Text ≤ 9 pt in {core.MUTED2} (Sekundärtext bis 9 pt nur in C.MUTED {core.MUTED})"),
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
    "kpi_ampel":       (WARNUNG, "KPI mit Ampelregel ohne Statusanzeige (Pill/Chip/Kante/Statusspalte in der Nähe)"),
    "link_ziel":       (FEHLER, "Link-Ziel existiert nicht (Blatt/Zelle/Name)"),
    "link_ziel_lage":  (WARNUNG, "Link-Ziel liegt in ausgeblendeter Zeile/Spalte"),
    "zeilenraster":    (WARNUNG, "Zeilenhöhe außerhalb des Rasters core.ROW_RASTER (P1-17)"),
    "theme_schrift":   (WARNUNG, "Designschriftart (Theme) nicht Calibri/Calibri Light"),
    # Komponentensprache Runde 2 (CORE_API.md): Typo-Skala, Statusfarben-Disziplin, Buttons
    "typo_skala":      (FEHLER, "Schriftgröße außerhalb der Typo-Skala core.TYPE_SCALE 8/8,5/9/10/12,5/16/20/22/30 pt (P3-11, P40)"),
    "emoji":           (WARNUNG, "Emoji/Farbsymbol (⚠ ℹ 🟢 …) statt monochromer Zeichen ● ▲ ■ › (Statusfarben-Disziplin)"),
    "kpi_statusfarbe": (WARNUNG, "Große Kennzahl (≥ 16 pt) in Statusfarbe ohne Statuswort daneben – Farbe als einziger Träger (P11: Wertfarbe = Status nur mit „● Wort“)"),
    "status_flaeche":  (WARNUNG, "Statusfarbe über eine ganze Fläche/Matrix (bedingte Füllung ≥ 3 × 2 Zellen oder grüne/amber Schrift > 12 Zellen)"),
    # Runde 3 (review3 P05, P16, P23, P15)
    "nav_rueckweg":    (FEHLER, "Sichtbares Blatt ohne ausgehenden Zell-Link auf ein anderes Blatt (Zell-Link-Rückfall der Navigation, P05)"),
    "formel_text":     (FEHLER, "Formelzelle mit Textformat „@“ (Excel zeigt nach Bearbeiten die Formel statt des Ergebnisses, P16)"),
    "minus_typo":      (WARNUNG, "Zahlenformat zeigt negatives Vorzeichen als Bindestrich „-“ statt typografischem Minus „−“ (P23)"),
    "neg_rot":         (WARNUNG, "Negative Summe/Endsumme (core.sum_row) ohne Negativ-Rot (P15)"),
    "button_hoehe":    (WARNUNG, "Button-Zeile weicht von der Button-Höhe core.BTN ab (Runde 4: alle Typen 25,5 pt; gemischte Reihe: größte Höhe – P1-03)"),
    # Runde 4 (review4 P1-04, P3-17, Themen „Farbsemantik“, „Druck“, „Formatbibliothek“)
    "eingabe_gelb":    (FEHLER, "Eingabe-Gelb FFF5D6 bzw. Eingaberahmen E6CB77 auf Formel- oder Link-Zelle (Gelb nur für echte Eingaben; Links/Chips C.ICE, überschreibbar = C.NOTE_BG – P1-04)"),
    "dropdown_zeichen": (FEHLER, "„▾“ an einer Zelle ohne Datenüberprüfung (▾ nur an Auswahllisten bzw. als Hinweis direkt darunter – P1-04)"),
    "link_tooltip":    (FEHLER, "Zell-Link ohne ScreenTip (hyperlink.tooltip leer; core.patch_tooltips nach der Neuberechnung – P3-17)"),
    "druck_umbruch":   (FEHLER, "„Anpassen an“ (fitToPage) zusammen mit manuellen Seitenumbrüchen – Excel ignoriert die Umbrüche"),
    "minus_text":      (WARNUNG, "Text/Textverkettung zeigt negative Zahl mit Bindestrich „-12“ statt „−12“ (C.minus_text / C.fixed_m)"),
    # Runde 5 (Farbdesign „Midnight & Gold“, DECISIONS „Nutzerentscheidung Runde 5“, CORE_API Runde 5)
    "altfarbe":        (FEHLER, "Alt-Farbe aus int7 (core.OLD_TO_NEW) in Zellstil, bedingtem Format, Form oder Registerfarbe – Token verwenden"),
    "gold_text":       (FEHLER, "Gold als Textfarbe unter 16 pt oder auf hellem Grund (Gold nur Linie/Kante/Fläche; Text in Bronze C.GOLD_INK)"),
    "kontrast":        (WARNUNG, "Textkontrast gegen die Zellfläche unter WCAG (Fließtext < 4,5:1, große Schrift < 3:1 – C.contrast)"),
    "diagramm_altfarbe": (WARNUNG, "Diagramm mit Alt-Farbe aus int7 (Blau-/Grauwerte) statt C.chart_color / C.CHART_CAT"),
    "diagramm_einfarbig": (WARNUNG, "Diagramm mit ≥ 3 sichtbaren Serien bzw. Kreissegmenten in nur einer Farbe (Runde 5: mehrfarbig nach Semantik)"),
    "diagramm_label_kontrast": (WARNUNG, "Datenbeschriftung innen (ctr/inEnd/inBase) mit Kontrast < 3:1 zur Balken-/Segmentfarbe (Weiß auf Gold/Aqua …)"),
    "diagramm_textfarbe": (WARNUNG, "Werte-/Achsen-/Legendentext in einer Serienfarbe (C.CHART_CAT) statt C.CHART_TEXT / CHART_TEXT2"),
}

# Bewusste Ausnahmen: (Blatt, Zelle oder Bereich oder "*", Regel) → Begründung
AUSNAHMEN = {
    ("AfA-Vergleich", "D7", "grau_klein"): "Farblegende: „grau = nicht anwendbar“ zeigt die Graustufe der inaktiven Zeilen",
    ("Finanzierung", "D7", "grau_klein"): "Farblegende: „grau = entfällt“ zeigt die Graustufe der entfallenen Werte",
    # Wunsch G (Runde 3): Kennzahlen-Check und Kachelwerte stehen bewusst neutral, Status im Chip H20:H25 bzw. L16/O16/R16
    ("Dashboard", "E20:E25", "kpi_ampel"): "Istwert neutral, Status im Chip H20:H25 derselben Zeile (P1-10)",
    ("Dashboard", "K15:R15", "kpi_ampel"): "Kachelwert neutral, Status in der Kontextzeile L16/O16/R16 (P1-10)",
}

ALLOWED_FONTS = {"Calibri", "Calibri Light"}

# Farbsemantik (P1-04): Eingabe-Gelb nur auf echten Eingabezellen (Konstanten ohne Formel und ohne Link)
INPUT_BG = getattr(core, "INPUT_BG", "FFF5D6").upper()
INPUT_LINE = getattr(core, "INPUT_LINE", "E6CB77").upper()
OVERRIDE_BG = getattr(core, "NOTE_BG", "FFF9EA").upper()      # „aus Kalkulation, überschreibbar“ (input_style override)
DROPDOWN = "▾"
# Bindestrich-Minus vor einer Zahl in Texten („Jahr 2: -274 €“); nicht nach Buchstabe/Ziffer („vv-GmbH“, „01-12“)
MINUS_TEXT_RE = re.compile(r"(?<![\w\d.,/)\]])-\s?\d")

# Statusfarben (Schrift) und Status-Flächen aus core
STATUS_FONT = {core.GREEN.upper(): "grün", core.AMBER.upper(): "amber", core.RED.upper(): "rot"}
STATUS_FILL = {x.upper() for x in (core.GREEN_BG, core.AMBER_BG, core.RED_BG, core.GREEN, core.AMBER, core.RED)}
KPI_BIG_PT = 16                       # ab dieser Größe gilt eine Zahl als „große Kennzahl“ (Kachelwert 20 pt)
# Emoji / farbige Symbolzeichen (nicht: ● ▲ ■ › ✓ ○ – monochrom und erlaubt)
EMOJI_RE = re.compile("[\U0001F000-\U0001FAFF\uFE0F\u2139\u26A0\u26A1\u26AA\u26AB\u26D4\u2705\u274C\u274E"
                      "\u2753-\u2757\u2B50\u2B55\u23F0-\u23FA\u231A\u231B\u2614\u2615]")
# Button-Arten aus core.BTN: (Fläche, Rahmen, Schrift) → (Art, Höhe)
BUTTON_KINDS = {}
for _k, _v in getattr(core, "BTN", {}).items():
    BUTTON_KINDS[(_v[0].upper(), _v[1].upper(), _v[2].upper())] = (_k, _v[4])
ERROR_VALUES = ("#DIV/0!", "#N/A", "#NV", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#WERT!", "#BEZUG!",
                "#ZAHL!", "#NAME", "#GETTING_DATA", "Err:", "#SPILL!", "#CALC!")
MUTED2 = core.MUTED2.upper()
# Runde 5: Farbdesign „Midnight & Gold“
OLD_COLORS = {k.upper(): v.upper() for k, v in getattr(core, "OLD_TO_NEW", {}).items()}
TOKEN_NAME = {}
for _n in ("NAVY", "NAVY_2", "BLUE", "ACCENT", "GOLD", "GOLD_BG", "GOLD_LINE", "GOLD_INK", "SKY", "MIST", "ICE", "TINT",
           "TINT_XL", "HEAD", "INK", "INK2", "MUTED", "MUTED2", "LINE", "LINE2", "LINE_SUB", "INACTIVE_BG",
           "NEUTRAL_DASH"):
    if isinstance(getattr(core, _n, None), str):
        TOKEN_NAME.setdefault(getattr(core, _n).upper(), _n)
GOLD_TEXT = {getattr(core, "GOLD", "C9A14A").upper(), "C9A94A", "C6A45C"}         # Edel-Gold (+ Alt-Gold)
GOLD_TEXT_OK_BG = {core.NAVY.upper(), getattr(core, "NAVY_2", core.NAVY).upper()}   # Gold-Text nur ≥ 16 pt auf Navy
SERIES_COLORS = {x.upper() for x in getattr(core, "CHART_CAT", ())}
WHITE = "FFFFFF"
INACTIVE_BG = getattr(core, "INACTIVE_BG", "F2F1ED").upper()
SYMBOL_RE = re.compile(r"[^\W_]")          # Buchstabe oder Ziffer: Lauf ist lesbarer Text (sonst Farbmuster/Symbol)
SEPARATORS = " \u00a0·|–—/•›‹"     # reine Trennzeichen-Läufe dürfen 8A9099 sein (Rich-Text-Trenner)

# Zeilenraster (pt): EINE Quelle – core.ROW_RASTER (Tokens inkl. H_PILL + Umbruchhöhen n × 13 + 8 / n × 13 + 10)
RASTER = set(getattr(core, "ROW_RASTER", ())) or (
    {core.H_BAND, core.H_BAND_B, core.H_BAND_B + 2, core.H_HEAD, core.H_ROW, core.H_ROW2, core.H_ROW3,
     core.H_STEP_ROW, core.H_STEP_ROW2, core.H_TILE_LABEL, core.H_TILE_VALUE, core.H_TILE_SUB, core.H_BUTTON, 14, 16, 22, 30}
    | {n * 13 + 8 for n in range(1, 12)} | {n * 13 + 10 for n in range(1, 12)})
TYPE_SCALE = set(getattr(core, "TYPE_SCALE", (8, 9, 10, 12.5, 16, 20, 22, 30)))
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


def hyphen_minus(fmt, value=None):
    """Zeigt das Format ein negatives Vorzeichen als Bindestrich „-“? (P23: typografisches Minus „−“)
    Mehrteilige Formate: Negativ-Sektion enthält „-“ (frei, \\- oder in Anführungszeichen). Einteilige Formate und
    Standard setzen „-“ automatisch – das zählt nur, wenn der Wert bekannt und negativ ist."""
    if not isinstance(fmt, str) or fmt.strip() == "@":
        return False
    secs = _split_sections(fmt)
    if any(re.match(r"\s*\[(<|>|=|<=|>=|<>)-?[\d.]+\]", x) for x in secs):
        return False                 # bedingte Formate: nicht eindeutig zuordenbar
    bare0 = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", secs[0])
    if fmt in ("General", "Standard") or len(secs) == 1:
        if fmt not in ("General", "Standard") and not re.search(r"[0#?]", bare0):
            return False             # Datum/Text
        return value is not None and not isinstance(value, bool) and isinstance(value, (int, float)) and value < 0
    neg = re.sub(r"\[[^\]]*\]", "", secs[1])
    return "-" in neg


def STATUS_WORD_FMT(numfmt):
    return isinstance(numfmt, str) and "●" in numfmt


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
        self.minus_cells = []     # (Blatt, Zelle, Format, Anzeige) – je Blatt und Format gebündelt gemeldet
        self.old_colors = defaultdict(list)   # (Blatt, Alt-Hex, Ort) → [Zellen] – gebündelt gemeldet (altfarbe)
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
            self.check_status_cf(ws, wv)
            self.check_print(ws)
            self.check_validation(ws)
            self.check_links(ws)
            self.check_input_look(ws, wv)
            self.check_drawing(ws, wv)
            self.check_nav(ws)
            self.check_neg(ws, wv)
            self.check_colors(ws, wv)
        self.report_minus()
        self.report_old_colors()
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
                if c.data_type == "f" and (c.number_format or "").strip() == "@":
                    self.add(title, ref, "formel_text", f"Formel mit Format „@“: {short(c.value, 40)}")
                text, kind, runs = self.display_text(c, wv[ref])
                if text is None or text == "":
                    continue
                fonts = {ru[4] for ru in runs if ru[0].strip() and ru[4]}
                for fn in sorted(fonts - ALLOWED_FONTS):
                    self.add(title, ref, "schriftart", f"„{fn}“")
                sizes = [ru[1] for ru in runs if ru[0].strip()]
                if sizes and min(sizes) < 8:
                    self.add(title, ref, "schrift_min", f"{min(sizes):g} pt")
                off = sorted({sz for sz in sizes if sz >= 8 and sz not in TYPE_SCALE})
                if off and self.in_print(ws, col, r):
                    self.add(title, ref, "typo_skala", f"{', '.join(f'{x:g}' for x in off)} pt: „{short(text, 32)}“")
                emo = EMOJI_RE.findall(text)
                if emo:
                    self.add(title, ref, "emoji", f"„{''.join(dict.fromkeys(emo))}“ in „{short(text, 40)}“")
                if kind == "num" and self.in_print(ws, col, r):
                    vv = wv[ref].value if c.data_type == "f" else c.value
                    if hyphen_minus(c.number_format, vv):
                        self.minus_cells.append((title, ref, c.number_format, text))
                if kind == "num" and sizes and max(sizes) >= KPI_BIG_PT:
                    colr = STATUS_FONT.get(rgb_of(f.color) if f is not None else None)
                    if colr:
                        self.add(title, ref, "kpi_statusfarbe", f"{max(sizes):g} pt „{short(text, 24)}“ fest in {colr}")
                small_grey = [ru[1] for ru in runs if ru[0].strip(SEPARATORS) and ru[5] == MUTED2 and ru[1] <= 9]
                if small_grey:
                    self.add(title, ref, "grau_klein", f"{max(small_grey):g} pt in 8A9099")
                if kind == "text" and MINUS_TEXT_RE.search(text) and self.in_print(ws, col, r):
                    m = MINUS_TEXT_RE.search(text)
                    self.add(title, ref, "minus_text",
                             f"„…{text[max(0, m.start() - 14):m.end() + 6]}…“ – Bindestrich statt „−“")
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
        free = self.callout_rows(ws, anchors) | self.button_rows(ws, anchors, geo)
        for r in sorted(rows_with_content):
            if r in CHROME_ROWS or geo.row_hidden(r) or r in multi or r in free or not self.in_print(ws, 1, r):
                continue
            d = ws.row_dimensions.get(r)
            h = d.height if d is not None else None
            if h is None or h <= SPACER_MAX:
                continue
            # LibreOffice rastet Höhen beim Speichern auf ganze Pixel ab (20 → 19,5 pt, 22 → 21,75 pt)
            if not any(-0.3 <= x - h <= 0.76 for x in RASTER):
                self.add(ws.title, f"{r}:{r}", "zeilenraster", f"Zeile {r}: {h:g} pt")

    @staticmethod
    def callout_rows(ws, anchors):
        """Letzte Zeile von Callout-Körpern (core.callout_box): verbundene F3F7FC-Bereiche mit oben ausgerichtetem
        9-pt-Umbruch. Ihre Höhe folgt der Textlänge (auf ganze Pixel) und liegt bewusst außerhalb des Rasters;
        ob der Text hineinpasst, prüft weiterhin umbruch_hoehe."""
        out = set()
        for (c, r), (c2, r2) in anchors.items():
            cell = ws.cell(r, c)
            al, f, fl = cell.alignment, cell.font, cell.fill
            if not (al.wrap_text and al.vertical == "top" and f is not None and float(f.sz or 11) == core.T_SMALL):
                continue
            if fl is None or fl.fill_type != "solid" or rgb_of(fl.fgColor) != core.TINT_XL.upper():
                continue
            out.add(r2)          # nur die letzte Körperzeile wächst mit dem Text (fit="auto")
        return out

    def button_rows(self, ws, anchors, geo):
        """Button-Zellen (core.btn) erkennen, Höhe gegen die Art prüfen; liefert die Zeilen (für zeilenraster frei)."""
        out = set()
        if not BUTTON_KINDS:
            return out
        _, inner = self.merges(ws)
        per_row = {}
        for row in ws.iter_rows():
            for c in row:
                if c.hyperlink is None or is_empty(c.value) or (c.column, c.row) in inner:
                    continue
                fl, f, bd = c.fill, c.font, c.border
                if fl is None or fl.fill_type != "solid" or bd is None or bd.top is None or not bd.top.style:
                    continue
                kind = BUTTON_KINDS.get((rgb_of(fl.fgColor), rgb_of(bd.top.color),
                                         rgb_of(f.color) if f is not None else None))
                if not kind or geo.row_hidden(c.row) or not self.in_print(ws, c.column, c.row):
                    continue
                out.add(c.row)
                per_row.setdefault(c.row, []).append((c, kind))
        for r, items in per_row.items():
            # gemischte Reihe (z. B. Buttons + Chip): eine Zeile hat nur eine Höhe → die größte Soll-Höhe gilt
            want = max(k[1] for _, k in items)
            h = geo.row_pt(r)
            if abs(h - want) > 0.76:
                c, kind = max(items, key=lambda x: x[1][1])
                self.add(ws.title, c.coordinate, "button_hoehe",
                         f"{kind[0]}-Button „{short(str(c.value), 28)}“ in {h:g}-pt-Zeile (soll {want:g} pt)")
        return out

    # ------------------------------------------------------------------ Statusfarben-Disziplin (bedingte Formate)
    def check_status_cf(self, ws, wv):
        """P1-10: keine Statusflächen über Matrizen, große Kennzahlen nicht per Regel grün/amber einfärben."""
        geo = self.g(ws)
        dxf_color = {}
        for cf in ws.conditional_formatting:
            cells = [(c, r) for rg in cf.sqref.ranges for r in range(rg.min_row, rg.max_row + 1)
                     for c in range(rg.min_col, rg.max_col + 1)]
            cells = [(c, r) for c, r in cells if not geo.row_hidden(r) and not geo.col_hidden(c)]
            if not cells:
                continue
            ref = str(cf.sqref).split(" ")[0]
            nrows = len({r for _, r in cells})
            ncols = len({c for c, _ in cells})
            fills, fonts = [], []
            for rule in cf.rules:
                d = rule.dxf
                if d is None:
                    continue
                fc = rgb_of(d.font.color) if d.font is not None and d.font.color is not None else None
                fl = None
                if d.fill is not None:
                    fl = rgb_of(d.fill.bgColor) or rgb_of(d.fill.fgColor)
                form = short((rule.formula or [""])[0], 36)
                if fl in STATUS_FILL and nrows >= 3 and ncols >= 2:
                    fills.append(form)
                elif fc in STATUS_FONT and STATUS_FONT[fc] != "rot" and len(cells) > 12:
                    fonts.append(STATUS_FONT[fc])
                if fc in STATUS_FONT and STATUS_FONT[fc] != "rot":
                    for c, r in cells:
                        dxf_color.setdefault((c, r), STATUS_FONT[fc])
            if fills:
                self.add(ws.title, ref, "status_flaeche",
                         f"bedingte Statusfüllung über {ncols} × {nrows} Zellen ({len(fills)} {'Regel' if len(fills) == 1 else 'Regeln'}, z. B. {fills[0]})")
            elif fonts:
                self.add(ws.title, ref, "status_flaeche",
                         f"Schrift {'/'.join(dict.fromkeys(fonts))} über {len(cells)} Zellen – Status als Pill/Spalte zeigen")
        _, inner = self.merges(ws)
        for (c, r), colr in sorted(dxf_color.items(), key=lambda x: (x[0][1], x[0][0])):
            if (c, r) in inner:
                continue
            cell = ws.cell(r, c)
            if is_empty(cell.value) or float(cell.font.sz or 11) < KPI_BIG_PT:
                continue
            v = wv.cell(r, c).value if cell.data_type == "f" else cell.value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and not self.status_word_near(ws, c, r):
                self.add(ws.title, cell.coordinate, "kpi_statusfarbe",
                         f"{float(cell.font.sz):g}-pt-Kennzahl per Regel {colr} gefärbt, aber kein „● Wort“ in der Kachel")

    def status_word_near(self, ws, c, r):
        """P11 „Wertfarbe = Status“: Die Farbe der Kachelzahl ist zulässig, wenn die Kachel den Status zusätzlich als Wort
        zeigt (core.tile: Chip „● prüfen“ rechts in der Fußzeile oder bedingtes Zahlenformat @* "● prüfen" inline).
        Gesucht wird in der Kachelbreite (Verbund der Label-/Wertzelle) von der Label- bis zur Fußzeile."""
        anchors, _ = self.merges(ws)
        c1, c2 = c, c
        for (ac, ar) in ((c, r - 1), (c, r)):
            if (ac, ar) in anchors:
                c2 = max(c2, anchors[(ac, ar)][0])
        rows = range(max(1, r - 1), r + 4)
        for rr in rows:
            for cc in range(c1, c2 + 1):
                cell = ws.cell(rr, cc)
                v = cell.value
                if isinstance(v, str) and any(g in v for g in "●▲") or STATUS_WORD_FMT(cell.number_format):
                    return True
        for cf in ws.conditional_formatting:
            hit = any(rg.min_col <= c2 and rg.max_col >= c1 and rg.min_row <= rows[-1] and rg.max_row >= rows[0]
                      for rg in cf.sqref.ranges)
            if hit and any(rule.dxf is not None and rule.dxf.numFmt is not None
                           and STATUS_WORD_FMT(rule.dxf.numFmt.formatCode) for rule in cf.rules):
                return True
        return False

    # ------------------------------------------------------------------ Runde 3: Navigation, Negativ-Rot, Minus
    def check_nav(self, ws):
        """P05: Jedes sichtbare Blatt hat mindestens einen Zell-Link auf ein anderes Blatt (Rückfall, falls Formen/
        Makros nicht greifen – z. B. Breadcrumb Z. 5 oder „‹ Start“)."""
        if ws.sheet_state != "visible":
            return
        for row in ws.iter_rows():
            for c in row:
                locs = []
                hl = c.hyperlink
                if hl is not None:
                    locs.append(hl.location or (str(hl.target) if hl.target and str(hl.target).startswith("#") else None))
                if c.data_type == "f" and isinstance(c.value, str) and "HYPERLINK(" in c.value.upper():
                    m = re.search(r'HYPERLINK\(\s*"#?([^"]+)"', c.value, re.I)
                    locs.append(m.group(1) if m else None)
                for loc in locs:
                    if not loc:
                        continue
                    res = self.resolve(loc)
                    if isinstance(res, tuple) and res[0] != ws.title:
                        return
        self.add(ws.title, "-", "nav_rueckweg", "kein Zell-Link auf ein anderes Blatt (Breadcrumb Z. 5 / „‹ Start“ fehlt)")

    def red_cells(self, ws):
        """Zellen mit bedingter roter Schrift (core.neg_red / status_cf rot)."""
        key = ("red", ws.title)
        if key not in self.geo:
            out = set()
            red = core.RED.upper()
            for cf in ws.conditional_formatting:
                if not any(rule.dxf is not None and rule.dxf.font is not None and rgb_of(rule.dxf.font.color) == red
                           for rule in cf.rules):
                    continue
                for rg in cf.sqref.ranges:
                    for rr in range(rg.min_row, rg.max_row + 1):
                        for cc in range(rg.min_col, rg.max_col + 1):
                            out.add((cc, rr))
            self.geo[key] = out
        return self.geo[key]

    def check_neg(self, ws, wv):
        """P15: negative Beträge in Zwischen-/Endsummen (core.sum_row 'sub'/'final') sind rot B42318 – per neg_red,
        [Red] im Format oder fester roter Schrift. Gemeldet wird nur, was im Build tatsächlich negativ ist."""
        geo = self.g(ws)
        navy, tint, line_sub = core.NAVY.upper(), core.TINT.upper(), getattr(core, "LINE_SUB", "D5DEEA").upper()
        red = core.RED.upper()
        _, inner = self.merges(ws)
        hits = defaultdict(list)
        for row in ws.iter_rows():
            if any(x.data_type == "s" and isinstance(x.value, str) and "Erstattung" in x.value for x in row):
                continue             # Steuerwirkung „+ Zahlung / – Erstattung“: negativ = Erstattung, bleibt neutral
            for c in row:
                if c.data_type not in ("f", "n") or (c.column, c.row) in inner or is_empty(c.value):
                    continue
                v = wv.cell(c.row, c.column).value if c.data_type == "f" else c.value
                if not isinstance(v, (int, float)) or isinstance(v, bool) or v >= 0 or round(v, 6) == 0:
                    continue
                if geo.row_hidden(c.row) or geo.col_hidden(c.column) or not self.in_print(ws, c.column, c.row):
                    continue
                f, fl, bd = c.font, c.fill, c.border
                if f is None or not f.b or rgb_of(f.color) != navy:
                    continue
                final = (fl is not None and fl.fill_type == "solid" and rgb_of(fl.fgColor) == tint
                         and bd is not None and bd.bottom is not None and bd.bottom.style == "double")
                sub = (bd is not None and bd.top is not None and bd.top.style == "thin"
                       and rgb_of(bd.top.color) == line_sub)
                if not (final or sub):
                    continue
                fmt = c.number_format or ""
                if "Erstattung" in fmt or "Zahlung" in fmt or fmt.strip() == "@" or fmt_class(fmt)[0] not in ("eur", "num"):
                    continue
                secs = _split_sections(fmt)
                if len(secs) > 1 and re.search(r"\[(Red|Rot)\]", secs[1], re.I):
                    continue
                if (c.column, c.row) in self.red_cells(ws):
                    continue
                hits[c.row].append((c.coordinate, "Endsumme" if final else "Zwischensumme", format_number(float(v), fmt)))
        for r, lst in hits.items():          # je Zeile eine Meldung
            ref = lst[0][0] if len(lst) == 1 else f"{lst[0][0]}:{lst[-1][0]}"
            self.add(ws.title, ref, "neg_rot", f"{lst[0][1]} Zeile {r}: {len(lst)} negative Beträge (z. B. {lst[0][2]}) nicht rot")

    def report_minus(self):
        """minus_typo gebündelt: je Blatt und Formatcode eine Meldung (sonst hunderte gleichartige Zeilen)."""
        groups = defaultdict(list)
        for sh, ref, fmt, text in self.minus_cells:
            groups[(sh, fmt)].append((ref, text))
        for (sh, fmt), cells in groups.items():
            ref, text = cells[0]
            more = f" (+{len(cells) - 1} Zellen)" if len(cells) > 1 else ""
            self.add(sh, ref, "minus_typo", f"Format {fmt!r}, z. B. „{short(text, 20)}“{more}")

    # ------------------------------------------------------------------ Druck, Gültigkeit, Links
    def check_print(self, ws):
        if ws.sheet_state != "visible":
            return
        ps = ws.page_setup.paperSize
        if str(ps) != "9":
            self.add(ws.title, "-", "druck_papier", f"paperSize={ps}")
        if not ws.print_area:
            self.add(ws.title, "-", "druck_bereich", "kein _xlnm.Print_Area")
        pr = ws.sheet_properties.pageSetUpPr
        rb = [b.id for b in ws.row_breaks.brk if b.man is not False]
        cb = [b.id for b in ws.col_breaks.brk if b.man is not False]
        if pr is not None and pr.fitToPage and (rb or cb):
            where = ", ".join([f"Zeile {x}" for x in rb] + [f"Spalte {get_column_letter(x)}" for x in cb])
            self.add(ws.title, "-", "druck_umbruch",
                     f"fitToPage aktiv, manuelle Umbrüche nach {where} werden ignoriert (feste Skalierung verwenden)")

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
                if not (getattr(hl, "tooltip", None) or "").strip():
                    self.add(ws.title, c.coordinate, "link_tooltip",
                             f"„{short(c.value if not isinstance(c.value, str) or not c.value.startswith('=') else hl.display, 32)}“"
                             f" → {hl.location or hl.target}")
                if hl.location:
                    self.check_target(ws.title, c.coordinate, hl.location, "Zell-Link")
                elif hl.target and str(hl.target).startswith("#"):
                    self.check_target(ws.title, c.coordinate, hl.target, "Zell-Link")

    def dv_cells(self, ws):
        """Zellen mit Datenüberprüfung (Auswahlliste o. ä.) – für dropdown_zeichen."""
        key = ("dv", ws.title)
        if key not in self.geo:
            out = set()
            for dv in ws.data_validations.dataValidation:
                for rg in dv.sqref.ranges:
                    for r in range(rg.min_row, min(rg.max_row, rg.min_row + 2000) + 1):
                        for cc in range(rg.min_col, min(rg.max_col, rg.min_col + 200) + 1):
                            out.add((cc, r))
            self.geo[key] = out
        return self.geo[key]

    def check_input_look(self, ws, wv):
        """P1-04 Farbsemantik: Gelb FFF5D6 und der Eingaberahmen E6CB77 kennzeichnen nur echte Eingaben (Konstanten).
        Zulässig bleiben: eine einzelne E6CB77-Kante (Hinweis-Callout, Trennlinie „Liste öffnen“ im Feld) und der
        Zustand „aus Kalkulation, überschreibbar“ (Fläche FFF9EA, Rahmen gestrichelt) auf Formelzellen.
        „▾“ nur an Zellen mit Datenüberprüfung oder als Hinweis direkt unter einer solchen (Start D24)."""
        geo = self.g(ws)
        anchors, inner = self.merges(ws)
        dvc = self.dv_cells(ws)
        for row in ws.iter_rows():
            for c in row:
                col, r = c.column, c.row
                if (col, r) in inner or geo.row_hidden(r) or geo.col_hidden(col):
                    continue
                is_f = c.data_type == "f"
                is_l = c.hyperlink is not None
                if is_f or is_l:
                    fl = rgb_of(c.fill.fgColor) if c.fill is not None and c.fill.fill_type == "solid" else None
                    bd = c.border
                    sides = [nm for nm, sd in (("links", bd.left), ("rechts", bd.right), ("oben", bd.top),
                                               ("unten", bd.bottom))
                             if sd is not None and sd.style and rgb_of(sd.color) == INPUT_LINE] if bd is not None else []
                    what = "Link" if is_l else "Formel"
                    if fl == INPUT_BG:
                        self.add(ws.title, c.coordinate, "eingabe_gelb",
                                 f"{what}-Zelle mit Eingabe-Gelb {INPUT_BG}: „{short(c.value, 32)}“"
                                 + ("" if is_l else " (überschreibbar → C.input_style(cell, \"override\"))"))
                    elif len(sides) >= 2 and (is_l or fl != OVERRIDE_BG):
                        self.add(ws.title, c.coordinate, "eingabe_gelb",
                                 f"{what}-Zelle mit Eingaberahmen {INPUT_LINE} ({'/'.join(sides)}): „{short(c.value, 32)}“")
                v = wv.cell(r, col).value if is_f else c.value
                if isinstance(v, CellRichText):
                    v = str(v)
                if not isinstance(v, str) or DROPDOWN not in v:
                    continue
                c2, _ = anchors.get((col, r), (col, r))
                span = range(col, c2 + 1)
                if any((k, r) in dvc for k in span) or any((k, r - 1) in dvc for k in span):
                    continue
                self.add(ws.title, c.coordinate, "dropdown_zeichen", f"„{short(v, 36)}“ ohne Datenüberprüfung")

    # ------------------------------------------------------------------ Runde 5: Farbdesign „Midnight & Gold“
    def note_old(self, sheet, hexv, where, ref):
        if hexv and hexv in OLD_COLORS:
            self.old_colors[(sheet, hexv, where)].append(ref)

    def check_colors(self, ws, wv):
        """Runde 5 (DECISIONS „Nutzerentscheidung Runde 5“): keine Alt-Hexwerte aus int7 (core.OLD_TO_NEW) in Zellstilen,
        bedingten Formaten und der Registerfarbe (altfarbe, je Blatt/Farbe/Ort gebündelt); Gold nie als Text unter 16 pt
        oder auf hellem Grund (gold_text); WCAG-Kontrast Text/Fläche (kontrast: < 4,5:1, ab 18 pt bzw. 14 pt fett < 3:1).
        Leere Zellen zählen für altfarbe mit (Flächen und Rahmen sind sichtbar), ausgeblendete Zeilen/Spalten nicht."""
        geo = self.g(ws)
        _, inner = self.merges(ws)
        title = ws.title
        low = defaultdict(list)      # kontrast je Zeile/Farbe gebündelt
        tab = ws.sheet_properties.tabColor if ws.sheet_properties is not None else None
        self.note_old(title, rgb_of(tab), "Registerfarbe", "-")
        for (r, col), c in sorted(ws._cells.items()):
            if geo.row_hidden(r) or geo.col_hidden(col):
                continue
            ref = c.coordinate
            fl = c.fill
            bg = WHITE
            if fl is not None and fl.fill_type == "solid":
                bg = rgb_of(fl.fgColor)
                self.note_old(title, bg, "Fläche", ref)
            elif fl is not None and fl.fill_type not in (None, "none"):
                bg = None                       # Muster/Verlauf: Kontrast nicht bestimmbar
            bd = c.border
            if bd is not None:
                for sd in (bd.left, bd.right, bd.top, bd.bottom):
                    if sd is not None and sd.style:
                        self.note_old(title, rgb_of(sd.color), "Rahmen", ref)
            f = c.font
            fcol = rgb_of(f.color) if f is not None and f.color is not None else None
            if (col, r) in inner or is_empty(c.value):
                continue
            v = wv.cell(r, col).value if c.data_type == "f" else c.value
            if is_empty(v):
                continue
            base_sz = float(f.sz or 11) if f is not None else 11.0
            base_b = bool(f.b) if f is not None else False
            runs = []           # (text, pt, fett, farbe | "theme")
            if isinstance(c.value, CellRichText) and c.data_type != "f":
                for blk in c.value:
                    if isinstance(blk, TextBlock):
                        ft = blk.font
                        runs.append((blk.text, float(ft.sz or base_sz),
                                     bool(ft.b) if ft.b is not None else base_b,
                                     (rgb_of(ft.color) or "theme") if ft.color is not None else (fcol or "000000")))
                    else:
                        runs.append((str(blk), base_sz, base_b, fcol or ("theme" if f is not None and f.color is not None else "000000")))
            else:
                runs.append((str(v), base_sz, base_b,
                             fcol or ("theme" if f is not None and f.color is not None else "000000")))
            for text, sz, bold, fg in runs:
                if not text.strip() or fg == "theme":
                    continue
                self.note_old(title, fg, "Schrift", ref)
                if not SYMBOL_RE.search(text):
                    continue         # reines Zeichen (■ ━ ● ◌ – –): Farbmuster/Legendenmarke, keine Schrift zum Lesen
                if fg in GOLD_TEXT and (sz < KPI_BIG_PT or bg not in GOLD_TEXT_OK_BG):
                    self.add(title, ref, "gold_text",
                             f"Gold {fg} als Schrift {sz:g} pt auf {bg or 'Muster'}: „{short(text, 28)}“ (Text → C.GOLD_INK)")
                    continue
                if bg is None or not self.in_print(ws, col, r) or (fg == MUTED2 and bg == INACTIVE_BG):
                    continue         # inaktiver Zustand (WCAG 1.4.3: deaktivierte Bedienelemente ausgenommen)
                cr = core.contrast(fg, bg)
                large = sz >= 18 or (bold and sz >= 14)
                need = 3.0 if large else 4.5
                if cr < need - 0.005:
                    shown = text if isinstance(v, str) or isinstance(c.value, CellRichText) else \
                        self.display_text(c, wv.cell(r, col))[0] or text
                    low[(r, fg, bg, sz, bold)].append((ref, shown, cr, need))
        for (r, fg, bg, sz, bold), items in sorted(low.items(), key=lambda x: (x[0][0], x[1][0][0])):
            ref0, shown, cr, need = items[0]
            rng = ref0 if len(items) == 1 else f"{ref0}:{items[-1][0]}"
            more = f" (+{len(items) - 1} Zellen)" if len(items) > 1 else ""
            self.add(title, rng, "kontrast",
                     f"{fg} auf {bg}: {cr:.2f}:1 (< {need:g}:1) bei {sz:g} pt{' fett' if bold else ''}: „{short(shown, 28)}“{more}")
        # bedingte Formate (dxf je Regel): Schrift, Fläche, Rahmen
        for cf in ws.conditional_formatting:
            ref = str(cf.sqref).split(" ")[0]
            for rule in cf.rules:
                d = rule.dxf
                if d is None:
                    continue
                if d.font is not None and d.font.color is not None:
                    self.note_old(title, rgb_of(d.font.color), "bedingte Schrift", ref)
                if d.fill is not None:
                    for x in (d.fill.bgColor, d.fill.fgColor):
                        self.note_old(title, rgb_of(x), "bedingte Fläche", ref)
                if d.border is not None:
                    for sd in (d.border.left, d.border.right, d.border.top, d.border.bottom):
                        if sd is not None and sd.style:
                            self.note_old(title, rgb_of(sd.color), "bedingter Rahmen", ref)
            for rule in cf.rules:          # Datenbalken/Farbskalen
                for obj in (getattr(rule, "dataBar", None), getattr(rule, "colorScale", None)):
                    if obj is None:
                        continue
                    for colr in (getattr(obj, "color", None) or []) if isinstance(getattr(obj, "color", None), list) \
                            else [getattr(obj, "color", None)]:
                        self.note_old(title, rgb_of(colr), "Datenbalken/Farbskala", ref)

    def report_old_colors(self):
        """altfarbe gebündelt: je Blatt, Alt-Hex und Ort eine Zeile mit Zellanzahl und Beispielen."""
        for (sheet, hexv, where), refs in sorted(self.old_colors.items()):
            new = OLD_COLORS.get(hexv, "?")
            tok = TOKEN_NAME.get(new, "Token")
            ex = ", ".join(refs[:3]) + (f" … (+{len(refs) - 3})" if len(refs) > 3 else "")
            self.add(sheet, refs[0], "altfarbe", f"{where} {hexv} → C.{tok} {new}: {len(refs)}× ({ex})")

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
            if chart is None:
                self.check_shape_colors(title, ref, name, el)
                txt = "".join(t.text or "" for t in el.iter("{%s}t" % NS["a"]))
                emo = EMOJI_RE.findall(txt)
                if emo:
                    self.add(title, ref, "emoji", f"{label}: „{''.join(dict.fromkeys(emo))}“ in „{short(txt, 40)}“")
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

    def check_shape_colors(self, sheet, ref, name, el):
        """Runde 5: Alt-Hexwerte in Formen (Füllung, Linie, Text) → altfarbe; Gold-Text in Formen nur ≥ 16 pt auf Navy."""
        a = "{%s}" % NS["a"]
        for clr in el.iter(a + "srgbClr"):
            self.note_old(sheet, (clr.get("val") or "").upper(), f"Form „{name}“", ref)
        for sp in el.iter():
            if sp.tag.split("}")[1] not in ("sp", "cxnSp"):
                continue
            sppr = sp.find("xdr:spPr", NS)
            fill = sppr.find("a:solidFill/a:srgbClr", NS) if sppr is not None else None
            bg = (fill.get("val") or "").upper() if fill is not None else None
            for rpr in sp.iter():
                if rpr.tag.split("}")[1] not in ("rPr", "defRPr", "endParaRPr"):
                    continue
                c = rpr.find("a:solidFill/a:srgbClr", NS)
                if c is None or (c.get("val") or "").upper() not in GOLD_TEXT:
                    continue
                sz = int(rpr.get("sz") or 1100) / 100
                if rpr.tag.endswith("endParaRPr"):
                    continue
                if sz < KPI_BIG_PT or bg not in GOLD_TEXT_OK_BG:
                    self.add(sheet, ref, "gold_text", f"Form „{name}“: Gold-Text {sz:g} pt auf {bg or 'ohne Fläche'}")
                    break

    @staticmethod
    def helper_series(ser):
        """Hilfsreihe (Wasserfall-Basis, Verbindungslinien …) nach core._CHART_HELPER – bleibt unsichtbar/neutral."""
        v = ser.find("c:tx//c:v", NS)
        nm = (v.text or "") if v is not None else ""
        rx, key = getattr(core, "_CHART_HELPER", None), getattr(core, "_chart_key", None)
        if not nm or rx is None or key is None:
            return False
        try:
            return bool(rx.search(key(nm)))
        except Exception:
            return False

    def check_chart_colors(self, sheet, ref, name, root):
        """Runde 5: Diagramme mehrfarbig nach C.CHART_SEMANTIC – keine int7-Altfarben, ≥ 3 sichtbare Serien bzw.
        Kreissegmente nicht einfarbig, Werte-/Achsentexte nicht in Serienfarbe."""
        c_, a_ = "{%s}" % NS["c"], "{%s}" % NS["a"]
        old = sorted({(x.get("val") or "").upper() for x in root.iter(a_ + "srgbClr")} & set(OLD_COLORS))
        if old:
            self.add(sheet, ref, "diagramm_altfarbe", f"„{name}“: {', '.join(old[:5])}{' …' if len(old) > 5 else ''}")

        def color_of(sppr, line=False):
            if sppr is None:
                return None
            if not line:
                if sppr.find("a:noFill", NS) is not None:
                    return "none"
                x = sppr.find("a:solidFill/a:srgbClr", NS)
                if x is not None:
                    return (x.get("val") or "").upper()
            ln = sppr.find("a:ln", NS)
            if ln is not None:
                if ln.find("a:noFill", NS) is not None:
                    return "none"
                x = ln.find("a:solidFill/a:srgbClr", NS)
                if x is not None:
                    return (x.get("val") or "").upper()
            return None

        for plot in root.iter(c_ + "plotArea"):
            for grp in plot:
                kind = grp.tag.split("}")[1]
                if not kind.endswith("Chart"):
                    continue
                line = kind in ("lineChart", "scatterChart", "radarChart")
                sers = grp.findall("c:ser", NS)
                pie = kind in ("pieChart", "pie3DChart", "doughnutChart", "ofPieChart")
                if pie:
                    for ser in sers:
                        cols = [color_of(dp.find("c:spPr", NS)) for dp in ser.findall("c:dPt", NS)]
                        cols = [x for x in cols if x and x != "none"]
                        if len(cols) >= 3 and len(set(cols)) == 1:
                            self.add(sheet, ref, "diagramm_einfarbig",
                                     f"„{name}“: {len(cols)} Kreissegmente alle {cols[0]} (C.chart_palette)")
                    continue
                cols = [color_of(ser.find("c:spPr", NS), line) for ser in sers if not self.helper_series(ser)]
                cols = [x for x in cols if x and x != "none"]
                if len(cols) >= 3 and len(set(cols)) == 1:
                    self.add(sheet, ref, "diagramm_einfarbig",
                             f"„{name}“ ({kind}): {len(cols)} Serien alle {cols[0]} (C.chart_color je Serie)")
        self.check_label_contrast(sheet, ref, name, root)
        bad = set()
        for tx in root.iter():
            if tx.tag.split("}")[1] not in ("txPr", "rich"):
                continue
            for rpr in tx.iter():
                if rpr.tag.split("}")[1] not in ("defRPr", "rPr"):
                    continue
                x = rpr.find("a:solidFill/a:srgbClr", NS)
                if x is not None and (x.get("val") or "").upper() in SERIES_COLORS:
                    bad.add((x.get("val") or "").upper())
        if bad:
            self.add(sheet, ref, "diagramm_textfarbe", f"„{name}“: Text in {', '.join(sorted(bad))} (C.CHART_TEXT/CHART_TEXT2)")

    INSIDE = {"ctr", "inEnd", "inBase"}

    def check_label_contrast(self, sheet, ref, name, root):
        """Innenliegende Datenbeschriftungen: Textfarbe gegen die Füllung der Säule/des Segments (≥ 3:1, Beschriftungen
        sind kurz und fett – WCAG-Grenze für große Schrift/Grafik). Einzel-Beschriftungen (c:dLbl idx) und
        Einzelpunkt-Füllungen (c:dPt idx) werden paarweise zugeordnet."""
        def fill(sppr):
            x = sppr.find("a:solidFill/a:srgbClr", NS) if sppr is not None else None
            return (x.get("val") or "").upper() if x is not None else None

        def txt_color(node):
            x = node.find("c:txPr//a:defRPr/a:solidFill/a:srgbClr", NS) if node is not None else None
            if x is None and node is not None:
                x = node.find("c:tx//a:rPr/a:solidFill/a:srgbClr", NS)
            return (x.get("val") or "").upper() if x is not None else None

        def pos(node):
            x = node.find("c:dLblPos", NS) if node is not None else None
            return x.get("val") if x is not None else None

        seen = set()
        for ser in root.iter("{%s}ser" % NS["c"]):
            base = fill(ser.find("c:spPr", NS))
            pts = {dp.find("c:idx", NS).get("val"): fill(dp.find("c:spPr", NS)) for dp in ser.findall("c:dPt", NS)
                   if dp.find("c:idx", NS) is not None}
            dl = ser.find("c:dLbls", NS)
            if dl is None:
                continue
            show = dl.find("c:showVal", NS)
            show_any = any(dl.find(f"c:{k}", NS) is not None and dl.find(f"c:{k}", NS).get("val") in ("1", "true")
                           for k in ("showVal", "showCatName", "showSerName", "showPercent"))
            d_col, d_pos = txt_color(dl), pos(dl)
            cases = []
            if show_any or show is None:
                cases.append((None, d_col, d_pos))
            for one in dl.findall("c:dLbl", NS):
                idx = one.find("c:idx", NS)
                if one.find("c:delete", NS) is not None and one.find("c:delete", NS).get("val") in ("1", "true"):
                    continue
                cases.append((idx.get("val") if idx is not None else None, txt_color(one) or d_col, pos(one) or d_pos))
            for idx, tc, ps in cases:
                bg = pts.get(idx, base) if idx is not None else base
                if not tc or not bg or ps not in self.INSIDE:
                    continue
                cr = core.contrast(tc, bg)
                if cr < 3.0 and (tc, bg) not in seen:
                    seen.add((tc, bg))
                    self.add(sheet, ref, "diagramm_label_kontrast",
                             f"„{name}“: Beschriftung {tc} auf {bg} {cr:.2f}:1 (C.CHART_TEXT bzw. Weiß nach C.contrast)")

    def check_chart_part(self, sheet, ref, name, part):
        if not part:
            self.add(sheet, ref, "link_ziel", f"Diagramm „{name}“: Diagrammteil fehlt")
            return
        data = self.pkg.read(part)
        if data is None:
            self.add(sheet, ref, "link_ziel", f"Diagramm „{name}“: {part} fehlt")
            return
        root = ET.fromstring(data)
        self.check_chart_colors(sheet, ref, name, root)
        bad_fonts, small, minus = set(), set(), set()
        for el in root.iter():
            t = el.tag.split("}")[1]
            if t == "numFmt" and el.get("sourceLinked") != "1" and hyphen_minus(el.get("formatCode")):
                minus.add(el.get("formatCode"))
            if t == "latin":
                tf = el.get("typeface")
                if tf and not tf.startswith("+") and tf not in ALLOWED_FONTS:
                    bad_fonts.add(tf)
            if t in ("rPr", "defRPr") and el.get("sz") and int(el.get("sz")) < 700:
                small.add(int(el.get("sz")) / 100)
        for tf in sorted(bad_fonts):
            self.add(sheet, ref, "schriftart", f"Diagramm „{name}“: „{tf}“")
        for fc in sorted(minus):
            self.add(sheet, ref, "minus_typo", f"Diagramm „{name}“: Achsen-/Beschriftungsformat {fc!r}")
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
            if dxf.numFmt is not None and hyphen_minus(dxf.numFmt.formatCode):
                self.add("(Mappe)", f"dxf {i}", "minus_typo", f"bedingtes Zahlenformat {dxf.numFmt.formatCode!r}")

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
                if not self.has_status(ws, c, r, cf_cache[sh]):
                    self.add(sh, cell.coordinate, "kpi_ampel", f"{spec['label']} ({via}) ohne Statusanzeige")

    STATUS_BOX = (-1, 4, -1, 2)      # Spalten c-1…c+4, Zeilen r-1…r+2: Kachel (Label/Wert/Kontext) bzw. Tabellenzeile

    def has_status(self, ws, c, r, ranges):
        """Status sichtbar? Neue Sprache (P1-10): die Zahl bleibt neutral, der Status sitzt als Pill/Chip in der
        Kontextzeile, als Kante (bedingter Rahmen) oder in einer Statusspalte derselben Zeile. Akzeptiert wird jede
        bedingte Formatierung oder Statusformel („●“) im Umfeld der Kennzahl."""
        dc1, dc2, dr1, dr2 = self.STATUS_BOX
        x1, x2, y1, y2 = c + dc1, c + dc2, r + dr1, r + dr2
        if any(a <= x2 and b >= x1 and rr1 <= y2 and rr2 >= y1 for a, rr1, b, rr2 in ranges):
            return True
        for rr in range(max(1, y1), y2 + 1):
            for cc in range(max(1, x1), x2 + 1):
                v = ws.cell(rr, cc).value
                if isinstance(v, str) and ("●" in v or "STATUS" in v.upper()):
                    return True
        return False


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
    from openpyxl.worksheet.pagebreak import Break
    from openpyxl.worksheet.properties import PageSetupProperties

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
    ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws["F1"], ws["F1"].font = "Elf", Font(name="Calibri", sz=11)                      # typo_skala
    ws["F2"] = "⚠ Achtung"                                                            # emoji
    ws["F3"], ws["F3"].font = 0.061, Font(name="Calibri", sz=20, color=core.GREEN)   # kpi_statusfarbe (fest)
    for rr in range(20, 24):                                                          # status_flaeche
        for cc in (8, 9):
            ws.cell(rr, cc, 1)
    core.status_cf(ws, "H20:I23", [("H20<0", "red"), ("ISNUMBER(H20)", "green")], font_color=False, fill_bg=True)
    core.btn(ws, "K", 30, "L", "Weiter  ›", "T", kind="primary")                    # button_hoehe
    ws.row_dimensions[30].height = 30
    ws["F32"], ws["F32"].number_format = "=1+1", "@"                                  # formel_text
    ws["F33"], ws["F33"].number_format = -5, '#,##0" €";-#,##0" €"'                  # minus_typo
    ws["B35"], ws["C35"] = "= Ergebnis", -1200                                        # neg_rot (Endsumme ohne Rot)
    ws["C35"].number_format = core.NUMFMT["eur"] if hasattr(core, "NUMFMT") else '#,##0" €"'
    core.sum_row(ws, 35, "B", "C", stage="final", neg=False)
    ws["H40"], ws["H40"].font = 0.043, Font(name="Calibri", sz=20, bold=True, color=core.NAVY)   # kpi_statusfarbe (Regel)
    core.status_cf(ws, "H40", [("H40<0.05", "amber")])
    # Runde 4: eingabe_gelb, dropdown_zeichen, link_tooltip (A10 ohne tooltip), druck_umbruch, minus_text
    ws["F42"] = "=A12*2"
    core.input_style(ws["F42"], "required")                                          # eingabe_gelb (Formel in Gelb)
    ws["F43"] = "Kauf als: Privat  ▾"                                                 # dropdown_zeichen (keine Liste)
    ws["F44"] = "Jahr 2: -274 €"                                                      # minus_text
    ws.sheet_properties.pageSetUpPr.fitToPage = True                                  # druck_umbruch
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.row_breaks.append(Break(id=30))
    # Runde 5: altfarbe, gold_text, kontrast, diagramm_altfarbe/einfarbig/textfarbe
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
    from openpyxl.styles import PatternFill
    old_hex = next(iter(core.OLD_TO_NEW))
    ws["F46"], ws["F46"].fill = "Altfarbe", PatternFill("solid", fgColor=old_hex)   # altfarbe (Fläche int7)
    ws["F46"].font = Font(name="Calibri", sz=10, color=core.WHITE)
    ws["F47"], ws["F47"].font = "Gold-Text", Font(name="Calibri", sz=10, color=core.GOLD)   # gold_text
    ws["F48"], ws["F48"].font = "blass", Font(name="Calibri", sz=10, color="C8C8C8")        # kontrast
    for i in range(12, 20):
        for k in (2, 3):
            ws.cell(i, k, i * k)
    ch2 = BarChart()
    ch2.add_data(Reference(ws, min_col=1, max_col=3, min_row=12, max_row=19))
    for ser in ch2.series:
        ser.graphicalProperties.solidFill = old_hex                                   # einfarbig + Altfarbe
    ch2.dataLabels = None
    ch2.x_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(
        defRPr=CharacterProperties(sz=900, solidFill=core.CHART_CAT[0])), endParaRPr=CharacterProperties())])
    from openpyxl.chart.label import DataLabelList
    ws.add_chart(ch2, "N40")                                                           # Achsentext in Serienfarbe
    ch3 = BarChart()
    ch3.add_data(Reference(ws, min_col=2, min_row=12, max_row=19))
    ch3.series[0].graphicalProperties.solidFill = core.C_GOLD
    ch3.series[0].dLbls = DataLabelList(showVal=True, dLblPos="ctr")                   # Weiß auf Gold: label_kontrast
    ch3.series[0].dLbls.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(
        defRPr=CharacterProperties(sz=900, b=True, solidFill=core.WHITE)), endParaRPr=CharacterProperties())])
    ws.add_chart(ch3, "N60")
    # Gegenprobe: neue Bausteine dürfen keine Befunde auslösen (Kachel mit Chip, Callout mit Texthöhe)
    ok = wb.create_sheet("OK")
    for c in "BCDEFGH":
        ok.column_dimensions[c].width = 14
    core.tile(ok, "B", "C", 10, 11, 12, kpi="BMR", value=0.043, sub="Ziel ≥ 5,0 %")
    ok["B11"].font = Font(name="Calibri", sz=20, bold=True, color=core.NAVY)
    core.callout_box(ok, "E", 10, "H", 11, 11, title="Einordnung",
                     text="Ein Satz, der die Kennzahl einordnet und über zwei Zeilen umbricht, damit die Höhe krumm wird.",
                     conditions=[("B11<0.05", "amber"), ("ISNUMBER(B11)", "green")])
    core.btn(ok, "B", 14, "C", "Weiter  ›", "OK", kind="primary")
    ok["B20"] = "Zurück"
    ok["B20"].hyperlink = Hyperlink(ref="B20", location="'T'!A1", tooltip="Zurück zu T")   # Rückweg mit ScreenTip
    core.btn(ok, "E", 14, "F", "Kauf als: Privat · ändern  ›", "T", kind="input")     # Alt-Art input → Chip, nie Gelb
    ok["B24"] = "Privatperson"                                                          # echte Eingabe mit Liste …
    core.input_style(ok["B24"], "required")
    dv_ok = DataValidation(type="list", formula1='"Privatperson,GmbH"', showErrorMessage=True)
    ok.add_data_validation(dv_ok)
    dv_ok.add("B24")
    ok["B25"] = "▾  Aus der Liste wählen"                                               # … und ▾-Hinweis direkt darunter
    ok["E24"] = "=C22*-1"
    core.input_style(ok["E24"], "override")                                             # überschreibbar: FFF9EA gestrichelt
    for x in ("B24", "B25", "E24"):
        ok[x].font = Font(name="Calibri", sz=10, color=core.BLUE)
    ok["E20"], ok["E20"].number_format = -5, core.NUMFMT["eur"]                       # typografisches Minus: kein Befund
    for x in ("B20", "E20", "B22", "C22"):
        ok[x].font = Font(name="Calibri", sz=10)
    ok["B22"], ok["C22"] = "= Cashflow", -300
    ok["C22"].number_format = core.NUMFMT["eur"]
    core.sum_row(ok, 22, "B", "C", stage="final")                                     # neg_red automatisch: kein Befund
    ok["H24"], ok["H24"].font = "■", Font(name="Calibri", sz=10, color=core.GOLD)   # Farbmuster (Legende): kein Befund
    ok["H25"], ok["H25"].font = "LEITFADEN › SCHRITT 01", Font(name="Calibri", sz=8.5, bold=True, color=core.GOLD_INK)
    ok["H26"], ok["H26"].font = "inaktiv", Font(name="Calibri", sz=10, color=core.MUTED2)
    ok["H26"].fill = PatternFill("solid", fgColor=core.INACTIVE_BG)                      # inaktiver Zustand: kein Befund
    ok.page_setup.paperSize = 9
    ok.print_area = "A1:M40"
    tmp = os.path.join(tempfile.mkdtemp(prefix="lint_selftest_"), "t.xlsx")
    wb.save(tmp)
    found = Linter(tmp).run()
    got = {f["rule"] for f in found if f["sheet"] == "T"}
    want = {"einzug", "schrumpfen", "schrift_min", "grau_klein", "schriftart", "text_ueberlauf", "zahl_raute",
            "einheit_doppelt", "fehlerwert", "diagramm_verdeckt", "fixierlinie", "druck_papier", "druck_bereich",
            "gueltigkeit", "link_ziel", "umbruch_hoehe", "zeile_zu_niedrig",
            "typo_skala", "emoji", "kpi_statusfarbe", "status_flaeche", "button_hoehe",
            "nav_rueckweg", "formel_text", "minus_typo", "neg_rot",
            "eingabe_gelb", "dropdown_zeichen", "link_tooltip", "druck_umbruch", "minus_text",
            "altfarbe", "gold_text", "kontrast", "diagramm_altfarbe", "diagramm_einfarbig", "diagramm_textfarbe",
            "diagramm_label_kontrast"}
    missing = want - got
    false_pos = [f"{f['ref']} {f['rule']}: {f['msg']}" for f in found if f["sheet"] == "OK"]
    print("Selbsttest:", "OK" if not missing else f"FEHLT {sorted(missing)}", f"({len(want)} Regeln)")
    print("Gegenprobe core-Bausteine:", "OK (0 Befunde)" if not false_pos else "FEHLALARM " + " · ".join(false_pos))
    return 0 if not missing and not false_pos else 1


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
