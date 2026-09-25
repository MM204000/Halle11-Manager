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
  TYPE_SCALE     Typo-Skala 8/8,5/9/10/12,5/16/20/22/30, snap_size()    – P3-11 / P40
  link_target()  linke obere Zelle des Scrollbereichs, link_row()       – P1-08
Ältere Bausteine (kpi_tile, callout, button, total, band_l1b, band_form …) bleiben unverändert nutzbar.

Runde 3 (Agent L, CORE_API.md „Runde 3 – Änderungen“): eine Kachel-Anatomie (Label 8,5 pt · Wert in Statusfarbe ·
Fußzeile Kontext links / Status rechts, Rinnen automatisch), EINE Pill (zarter Status-Tint, ohne Rahmen), drei
Summenstufen final/sub/memo, Button-Dreistufigkeit primary > secondary > tertiary, typografisches Minus „−“ in allen
Formaten, T_LABEL 8,5 pt, KPI_LABELS/KPI_ORDER, Hinweis-Callout note(), status_banner(), fit_row(metric=True),
page_header(right_legend=…).

Runde 4 (Agent L, CORE_API.md „Runde 4 – Änderungen“): EINE Kachel-Spezifikation mit festen Höhen 18/30/20 pt und genau
zwei Kopffarben (strong 0B2A4A / calm 1D4F8A je nach Blattrolle), Kachelwert immer 20 pt, Überhöhe der Wertzeile wandert in
die Abstandszeile (finalize_components) – P1-02; drei Button-Typen primary/secondary/tertiary, alle 25,5 pt, Link-Chips
E7EEF7 statt Gelb – P1-03; Callout-Höhe aus dem Text (n × 12,5 + 8 pt) mit Innenabstand rechts – P2-01; Kopfschablone
Z. 4/5/6/7 = 12/18/30/21,75 pt und kanonische Brotkrumen – P2-03/P2-04; Summenstufen final/sub/kpi/plain/memo – P2-06;
Formatbibliothek EUR/PCT1/PCT2/MULT2/EUR_M2/M2/YEARS mit Nullabschnitt „–“, fixed_m()/minus_text() für „−“ in
Textverkettungen – P2-07/P3-15; Band-Meta immer rechts – P2-10; ScreenTips für jeden Zell-Link – P3-17;
KPI_TILE_LABELS – P3-19; Eingabezustände inkl. „überschreibbar“ – P3-20.

Runde 5 (Agent L, CORE_API.md „Runde 5 – Farbdesign“): Farbdesign „Midnight & Gold“ – Nachtblau 0E2238 / 1E4E8C,
Edel-Gold C9A14A als Akzentkante (Abschnittskopf, Kachel-Kopflinie strong, Primär-Button-Rahmen, Formular-Band),
warme Neutralflächen F7F6F2 / FBFAF7 / F1EEE6, warme Linien E4E1D8 / D6D2C6, Bronze-Brotkrume 7A5C1E;
Diagramm-Palette CHART_CAT + Semantik CHART_SEMANTIC, chart_color() / chart_palette() / chart_dashed() / muted();
Umfärbe-Tabelle OLD_TO_NEW für finish_sheets (Agent K); contrast() für WCAG-Prüfungen.
"""
import math
import re

from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

# =============================================================================== Farben
# Runde 5 – Farbdesign „Midnight & Gold“ (DECISIONS „Nutzerentscheidung Runde 5“, CORE_API.md „Runde 5 – Farbdesign“).
# Die NAMEN der Tokens bleiben, die WERTE ändern sich. Kontraste (WCAG, gegen Weiß bzw. die jeweilige Fläche) im Kommentar.
NAVY = "0E2238"      # Tinte/Primär (Nachtblau): Kopfleiste, Hero, Titel, Kachelkopf strong, Primär-Button – 16,1:1 auf Weiß
NAVY_2 = "1B3553"    # zweite Nachtblau-Stufe: inaktive Reiter/Chips AUF Navy, Hero-Nebenfläche (Weiß darauf 12,4:1)
BLUE = "1E4E8C"      # Sekundär: Links, verknüpfte Werte, Kachelkopf calm, Sekundär-Button – 8,3:1 auf Weiß
ACCENT = "3B6A9E"    # Stahlblau: Tabellenkopf-/Unterlinien, gestrichelte Verknüpfung, Schrittnummern – 5,6:1 auf Weiß
GOLD = "C9A14A"      # Edel-Gold – NUR Linien/Kanten/Flächen (Abschnittskante, Kachel-Kopflinie, Primär-Button-Kante,
#                      aktiver Reiter, Hero-Linie); als Text nur ≥ 16 pt auf Navy (6,7:1), NIE Fließtext auf Hell
GOLD_BG = "F6EFDF"   # helles Gold: Hervorhebung (Verkaufsjahr, Basisfall, Markierung) – Navy darauf 14,1:1
GOLD_LINE = "E2CF9F" # zarte Goldlinie auf hellem Grund (Rahmen um GOLD_BG-Flächen)
GOLD_INK = "7A5C1E"  # Bronze als TEXT (Brotkrume/Eyebrow 8,5 pt fett) – 6,2:1 auf Weiß
SKY = "A9B8CC"       # Sekundärtext AUF Navy (8–9 pt) – 8,0:1 auf 0E2238
MIST = "CBD5E1"      # kühle helle Linie / ruhige Kante
ICE = "E8EDF4"       # kühle helle Fläche: Link-Chips, erledigte Schritte, Marker
TINT = "F7F6F2"      # Band (warm): L1-Abschnitt, Endsumme, Kennzahlenband – Tinte 15,6:1, MUTED 7,0:1
TINT_XL = "FBFAF7"   # sehr hell (warm): Kachelkörper, Einordnungs-Box, Info-Zeilen
HEAD = "F1EEE6"      # Spaltenkopf (warm, eine Stufe dunkler als das Band) – BLUE darauf 7,2:1
INK = "1A1D21"       # Text
INK2 = "3A3F45"      # Fließtext in Boxen – 10,7:1 auf Weiß
MUTED = "4E5561"     # Sekundärtext (bis 9 pt immer diese Farbe) – 7,5:1 auf Weiß, 7,0:1 auf TINT
MUTED2 = "8A9099"    # nur Linien, inaktive Felder, Text ab 11 pt
LINE = "E4E1D8"      # Haarlinie Datenzeile (warm)
LINE2 = "D6D2C6"     # Trennlinie / inaktiver Rahmen (warm)
WHITE = "FFFFFF"
GREEN, AMBER, RED = "1F7A4D", "B54708", "B42318"
GREEN_BG, AMBER_BG, RED_BG = "EAF5EF", "FDF3E7", "FBECEB"
GREEN_LINE, AMBER_LINE, RED_LINE = "B5DCC4", "F1CFA5", "E7B4AD"   # Status-Kanten/-Rahmen auf hellem Grund
LINE_SUB = "CFC9BA"   # Oberlinie der Zwischensumme (sum_row 'sub') – warme Trennlinie
OVERRIDE_BG = "FFF9EA"   # Eingabezustand „aus Kalkulation, überschreibbar“ (P3-20): zarte Fläche, gestrichelter Rahmen
NOTE_BG, NOTE_LINE = "FFF9EA", "E6CB77"   # Hinweis-Callout (note, P34): zarte Fläche, linke Kante wie Eingaberahmen
NEUTRAL_DASH = "9AA3AF"   # „–“ für entfallende Werte (lesbar, aber zurückgenommen)
INPUT_BG, INPUT_LINE, INPUT_FG = "FFF5D6", "E6CB77", BLUE
INACTIVE_BG, INACTIVE_FG = "F2F1ED", MUTED2

# ------------------------------------------------------------------------------- Diagramm-Palette (Runde 5)
# Kategorial, validiert (dataviz validate_palette, light): Helligkeitsband, Chroma, CVD ΔE ≥ 9,1, Normalsicht ΔE ≥ 22,9.
# FESTE Reihenfolge, nie zyklisch: eine 7. Kategorie wird Neutral (bzw. „Sonstige“). Kontrast < 3:1 bei Aqua/Gold/Magenta
# → Identität nie nur über Farbe: Legende + direkte Beschriftung; Werte-/Achsentexte in Tintenfarben (CHART_TEXT).
C_BLUE, C_ORANGE, C_AQUA, C_GOLD, C_VIOLET, C_MAGENTA = "2A78D6", "EB6834", "1BAF7A", "EDA100", "7A4FC9", "E87BA4"
C_NEUTRAL = "9AA3AF"          # Restschuld (gestrichelt), „nicht anwendbar“, Bestand
C_NEUTRAL_LIGHT = "CDD2D9"    # gedämpfte Serien (AfA-Varianten, die nicht gelten)
C_POS, C_NEG = "1BAF7A", "E34948"   # divergierend: Cashflow-Überschuss / Unterdeckung, Zufluss / Abfluss
C_INK = NAVY                  # Nettovermögen / Summen / Ergebnis
CHART_CAT = (C_BLUE, C_ORANGE, C_AQUA, C_GOLD, C_VIOLET, C_MAGENTA)
CHART_TEXT, CHART_TEXT2, CHART_GRID, CHART_AXIS = INK, MUTED, LINE, LINE2   # Beschriftung · Achse/Legende · Gitter · Achslinie
CHART_SEP = WHITE             # 2-px-Trennlinie zwischen Kreissegmenten/Stapeln
CHART_MARK = GOLD             # Markierung (Verkaufsjahr/-punkt) – Marke, keine Serie

# Semantik (mappenweit: gleiche Größe = gleiche Farbe). Reihenfolge = Prüfreihenfolge (erstes Muster gewinnt);
# Muster: regulärer Ausdruck, ohne Groß/Klein, gegen den Serien- bzw. Kategorienamen. Wert None = Hilfsreihe (unsichtbar
# lassen). dash: 'dash' für gestrichelte Linien (Restschuld).
CHART_SEMANTIC = (
    (r"^(basis|stand|summe|verbindung)\b|^basis [+−-]|^[+−-]?\d", None),     # Hilfsreihen (Wasserfall, Beschriftung)
    (r"nettoverm|nettoerl|gesamtertrag|^= ?nach|ergebnis \(|^ergebnis$", C_INK),
    (r"restschuld", C_NEUTRAL),
    (r"nicht anwendbar|\(keine|\(keines|n\. ?v\.", C_NEUTRAL_LIGHT),
    (r"verkaufs(preis|punkt|jahr)|^verkauf |ende zinsbindung", CHART_MARK),
    (r"^boden", C_NEUTRAL),                                                   # nicht abschreibbar
    (r"kumulierter cashflow", C_BLUE),
    (r"unterdeckung|lücke|abfluss|zuschuss|negativ", C_NEG),
    (r"überschuss|zufluss|cashflow kumuliert", C_POS),
    (r"zins|finanzierungskosten", C_VIOLET),
    (r"steuerliches ergebnis", C_BLUE),
    (r"steuer(?!n\b)|grunderwerb", C_GOLD),                                 # nicht „nach/vor Steuern“
    (r"tilgung|eigenkapital|immobilienwert|vermögen|anwendbar$", C_AQUA),
    (r"miete|einnahm|kaltmiete|im modell|afa regul|neues objekt|darlehen i\b(?! ?i)|kaufpreis|gebäude|cashflow", C_BLUE),
    (r"bewirtsch|kosten|hausgeld|verwaltung|instandh|mietausfall|werbungsk|ausgaben|kapitaldienst", C_ORANGE),
    (r"darlehen ii|sonder", C_VIOLET),
    (r"nebenkosten|sonstig|beweglich|inventar|rücklage|notar|makler|grundbuch", C_MAGENTA),
    (r"bestehend", C_NEUTRAL),
)
CHART_DASH = (r"restschuld", r"ende zinsbindung", r"^ziel|schwelle")   # gestrichelt darstellen
_CHART_RX = [(re.compile(p, re.I), c) for p, c in CHART_SEMANTIC]
_CHART_HELPER = _CHART_RX[0][0]


def _chart_key(name):
    t = re.sub(r"\s+", " ", str(name or "")).strip()
    return re.sub(r"\s*·\s*[\d.,]+\s*%$", "", t)        # Kreis-Kategorie „Makler · 34 %“ → „Makler“


def chart_color(name, default=None):
    """Farbe einer Diagrammserie/-kategorie nach der Semantik-Tabelle (Runde 5, für Agent H und G).
    'Nettokaltmiete' → 2A78D6 · 'Kumulierte Zinsen' → 7A4FC9 · 'Restschuld' → 9AA3AF · 'Nettovermögen' → 0E2238 ·
    'Überschuss n. St.' → 1BAF7A · 'Unterdeckung' → E34948. Hilfsreihen ('Basis', 'Summe', '+922') → None.
    Kein Treffer → default (None)."""
    t = _chart_key(name)
    for rx, c in _CHART_RX:
        if rx.search(t):
            return c
    return default


def chart_dashed(name):
    """True, wenn die Serie gestrichelt gezeichnet wird (Restschuld, Zinsbindungsende, Zielwerte)."""
    return any(re.search(p, str(name or ""), re.I) for p in CHART_DASH)


def chart_palette(names, semantic=True):
    """Farben für eine Serien-/Kategorienliste EINES Diagramms: zuerst Semantik (chart_color), jede Serienfarbe höchstens
    einmal; übrige Einträge erhalten die nächste freie Farbe aus CHART_CAT in fester Reihenfolge, danach C_NEUTRAL
    (nie zyklisch). Hilfsreihen ('Basis', 'Summe', '+922') → None. Beispiel Kreis Bewirtschaftung ['Hausgeld',
    'Rücklage WEG', 'Verwaltung', …] → Orange, Magenta, Blau, Aqua, Gold, Violett. Liefert eine Liste gleicher Länge."""
    helper = [bool(semantic and _CHART_HELPER.search(_chart_key(n))) for n in names]
    out, used = [], set()
    for n, h in zip(names, helper):
        c = None if (h or not semantic) else chart_color(n)
        if c in used and c not in (C_NEUTRAL, C_NEUTRAL_LIGHT, C_INK, CHART_MARK):
            c = None
        out.append(c)
        if c:
            used.add(c)
    for i, c in enumerate(out):
        if c is None and not helper[i]:
            out[i] = next((x for x in CHART_CAT if x not in used), C_NEUTRAL)
            used.add(out[i])
    return out


def mix(hex_a, hex_b=WHITE, t=0.5):
    """Farbmischung a→b (t = Anteil b): mix(C_BLUE, WHITE, 0.6) = gedämpfte Variante für nicht angewendete Serien."""
    a = [int(hex_a[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(hex_b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def muted(hex_c, t=0.55):
    """Gedämpfte Serienfarbe (Richtung Weiß) – z. B. AfA-Varianten, die nicht angewendet werden."""
    return mix(hex_c, WHITE, t)


def contrast(fg, bg=WHITE):
    """WCAG-Kontrastverhältnis zweier Hexfarben (Fließtext ≥ 4,5; große Schrift/Linien ≥ 3)."""
    def lum(h):
        r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4  # noqa: E731
        return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)
    la, lb = sorted((lum(fg), lum(bg)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


# Umfärbe-Tabelle für Agent K (finish_sheets: Styles, bedingte Formate, Zeichnungen – NICHT Diagramm-XML):
# alle bisher (int7) genutzten Blau-/Grau-/Alt-Gold-Hexwerte → Runde-5-Token. Status-, Eingabe- und Weiß-Werte bleiben.
OLD_TO_NEW = {
    # Primär/Sekundär/Akzent
    "0B2A4A": NAVY, "0F2D4F": NAVY, "173026": NAVY, "1A3F66": NAVY_2, "1A3F6B": NAVY_2,
    "1D4F8A": BLUE, "2F6BAE": ACCENT, "2F6FB0": ACCENT, "4570A5": ACCENT, "4A86C8": ACCENT,
    # helle Blautöne → kühle Linien/Flächen bzw. warme Bänder
    "7FA7D6": SKY, "8DB3DE": SKY, "8FB3DE": SKY, "8FA9C9": SKY, "9CBBE2": SKY, "7F9BB8": SKY,
    "B9CFE8": MIST, "C8D7EB": MIST, "C9D6E6": MIST, "C9D6E8": MIST, "C9DBEF": MIST, "C9D6E0": MIST,
    "CEDDF0": ICE, "D5E3F3": ICE, "EAF1FA": TINT_XL,
    "E7EEF7": TINT, "F3F7FC": TINT_XL, "F5F8FC": TINT_XL, "FBFCFE": TINT_XL, "EEF3FA": HEAD,
    # Grau/Linien → warm neutral
    "5B6068": MUTED, "E6E8EB": LINE, "D5D9DE": LINE2, "D5DFEB": LINE_SUB, "D5DEEA": LINE_SUB,
    "F3F4F6": INACTIVE_BG, "F4F5F6": INACTIVE_BG, "F3F4F5": INACTIVE_BG, "F7F8F9": TINT_XL,
    "98A2B3": NEUTRAL_DASH, "94A3B8": NEUTRAL_DASH, "9AA8B8": NEUTRAL_DASH, "9AA4B1": NEUTRAL_DASH,
    "9AA5B4": NEUTRAL_DASH, "64748B": MUTED,
    # Alt-Gold (Logo/Marke) → Edel-Gold
    "C9A94A": GOLD, "C6A45C": GOLD, "A08A55": GOLD_INK,
}

SANS, DISPLAY = "Calibri", "Calibri"

# =============================================================================== Typo-Skala und Zeilenraster
# Erlaubt sind nur 8 · 8,5 · 9 · 10 · 12,5 · 16 · 20 · 22 · 30 pt (P3-11, Runde 3 P40). Keine Zellschrift unter 8 pt.
#   8   T_MICRO  nur Seitenfuß/Rechtshinweise, Kachel-Kontextzeile und Status-Chip
#   8,5 T_LABEL  Versalien-Labels: Kachel-Label, Spaltenkopf (Ebene 2), Eyebrow, Matrixachsen, Band-Meta (P40)
#   9   T_SMALL  Sekundärtext, Einordnungstext, Ebene-2-Titel mit Fläche, Status-Pill
#   10  T_BODY   Fließtext, Tabellen, Buttons, Callout-Titel
#   12,5 T_H3    Abschnittskopf Ebene 1
#   16  T_H2     Zwischentitel groß (Dashboard-Urteil, Hero-Nebenzahlen)
#   20  T_KPI    Kachelwert
#   22  T_H1     Seitentitel
#   30  T_HERO   Hero (Start)
T_MICRO, T_SMALL, T_BODY, T_H3, T_H2, T_KPI, T_H1, T_HERO = 8, 9, 10, 12.5, 16, 20, 22, 30
T_LABEL = 8.5
TYPE_SCALE = (T_MICRO, T_LABEL, T_SMALL, T_BODY, T_H3, T_H2, T_KPI, T_H1, T_HERO)
H_BAND, H_BAND_B, H_HEAD, H_ROW, H_ROW2, H_ROW3 = 24, 20, 20, 18, 30, 42
H_STEP_ROW, H_STEP_ROW2 = 22, 32            # Datenzeilen der Schritt-Seiten (1- / 2-zeilig)
H_TILE_LABEL, H_TILE_VALUE, H_TILE_SUB = 18, 30, 20   # Runde 3 (P10): Fußzeile 16 → 20 pt, Luft unter dem Kontext
H_BUTTON = 25.5                             # Primär/Sekundär-Button (34 px) – P1-13
H_BTN = H_BUTTON
H_PILL = 22.5                               # Link-Pills / Statuschip in Link-Reihen (30 px)
H_CALLOUT_HEAD = 20                         # Kopfzeile der Einordnungs-Box
H_GAP = 12                                  # Leer-/Fugenzeile zwischen Komponenten
GRID_HEIGHTS = (H_ROW, H_ROW2, H_ROW3)      # Zeilenraster 1/2/3 Zeilen; ab 4 Zeilen n × 13 + 8
TEXT_PITCH, TEXT_PAD = 12.5, 8              # Fließtexttabellen (fit_row metric=True, P27): n × 12,5 + 8 pt


def text_row_height(n):
    """Zeilenhöhe einer Fließtextzeile mit n Textzeilen (P27): n × 12,5 + 8 pt, auf ganze Pixel → 21 · 33 · 45,75 · 58,5 …"""
    return math.ceil(round((max(1, n) * TEXT_PITCH + TEXT_PAD) / 0.75, 3)) * 0.75


# Zulässige Zeilenhöhen (pt) – identisch mit lint_pro.RASTER; Abstands-/Fugenzeilen ≤ 12 pt sind frei.
ROW_RASTER = sorted({H_BAND, H_BAND_B, H_BAND_B + 2, H_HEAD, H_ROW, H_ROW2, H_ROW3, H_STEP_ROW, H_STEP_ROW2,
                     H_TILE_LABEL, H_TILE_VALUE, H_TILE_SUB, H_BUTTON, H_PILL, 14, 16, 22, 30}
                    | {n * 13 + 8 for n in range(1, 12)} | {n * 13 + 10 for n in range(1, 12)}
                    | {text_row_height(n) for n in range(1, 12)})


def px_pt(pt):
    """Höhe auf ganze Excel-Pixel runden (0,75 pt), aufwärts."""
    return math.ceil(round(pt / 0.75, 3)) * 0.75


def snap_size(pt):
    """Schriftgröße auf die Typo-Skala abbilden (8,5 bleibt, 9,5 → 10, 11 → 10, 12 → 12,5, 14 → 16; < 8 → 8).
    Regel: nächste Stufe; 11 wird bewusst zu 10 (Textschrift), 12/14 zur nächsthöheren Titelstufe."""
    if pt is None:
        return None
    if pt in TYPE_SCALE:
        return pt
    if pt < T_MICRO:
        return T_MICRO
    if pt <= 11:
        if pt < 8.75:
            return T_LABEL if pt >= 8.25 else T_MICRO
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
MINUS = "\u2212"      # typografisches Minus „−“ (P23) – in allen Formaten statt Bindestrich
_M = '"\u2212"'
NUMFMT = {
    "eur": f'#,##0" €";{_M}#,##0" €";"–"',
    "eur_zero": f'#,##0" €";{_M}#,##0" €";0" €"',
    "eur_in": f'#,##0" €";{_M}#,##0" €";0" €"',
    "eur2": f'#,##0.00" €";{_M}#,##0.00" €";"–"',
    "eur2_in": f'#,##0.00" €";{_M}#,##0.00" €";0.00" €"',
    "eur_signed": f'"+"#,##0" €";{_M}#,##0" €";"–"',
    "eur_signed_zero": f'"+"#,##0" €";{_M}#,##0" €";0" €"',
    "eur_plain": f'#,##0;{_M}#,##0;"–"',
    "eur_plain_signed": f'"+"#,##0;{_M}#,##0;"–"',
    "teur": f'#,##0," T€";{_M}#,##0," T€";"–"',
    "num": f'#,##0;{_M}#,##0;"–"',
    "num_in": f"#,##0;{_M}#,##0",
    "num2": f'#,##0.00;{_M}#,##0.00;"–"',
    "num2_in": f"#,##0.00;{_M}#,##0.00",
    "int": f"#,##0;{_M}#,##0",
    "int0": f'0;{_M}0;"–"',
    "int_in": f"0;{_M}0",
    "pct0": f'0 %;{_M}0 %;"–"',
    "pct1": f'0.0 %;{_M}0.0 %;"–"',
    "pct1_in": f"0.0 %;{_M}0.0 %",
    "pct1_signed": f'"+"0.0 %;{_M}0.0 %;"±"0.0 %',
    "pct2": f'0.00 %;{_M}0.00 %;"–"',
    "pct2_in": f"0.00 %;{_M}0.00 %",
    "dscr": f'0.00"×";{_M}0.00"×";"–"',
    "dscr_plain": f"0.00;{_M}0.00",
    "mult1": f'0.0"×";{_M}0.0"×";"–"',
    "mult2": f'0.00"×";{_M}0.00"×";"–"',
    "year": "0",
    "years": '0" Jahre"',
    "years_n": '[=1]0" Jahr";0" Jahre"',
    "years_in": '[=1]0" Jahr";0" Jahre"',
    "years_calc": '[=1]0" Jahr";[=0]"–";0" Jahre"',      # berechnete/übernommene Jahre: 0 → „–“ (P2-07)
    "months_calc": '[=1]0" Monat";[=0]"–";0" Monate"',
    "children": '[=1]0" Kind";0" Kinder"',
    "date": "DD.MM.YYYY",
    "qm": f'#,##0" m²";{_M}#,##0" m²"',                 # Eingabe (0 sichtbar) – berechnet: "m2"
    "m2": f'#,##0" m²";{_M}#,##0" m²";"–"',
    "eur_qm": f'#,##0.00" €/m²";{_M}#,##0.00" €/m²";"–"',
    "eur_qm_in": f'#,##0.00" €/m²";{_M}#,##0.00" €/m²";0.00" €/m²"',
    "tax_effect": '"Zahlung "#,##0" €";"Erstattung "#,##0" €";"–"',
    "yesno": '[=1]"ja";[=0]"nein";0',
    "status_dot": '"●  "@',
    "hidden": ";;;",
    "text": "@",
}


# Formatbibliothek (Runde 4, P2-07) – EINE Quelle, benannt wie im Plan. Berechnete Werte zeigen 0 als „–“,
# Eingabefelder nutzen die *_in-Varianten aus NUMFMT (0 sichtbar). Negativ immer mit typografischem Minus „−“.
EUR = NUMFMT["eur"]            # 1.234 € · −1.234 € · –
EUR2 = NUMFMT["eur2"]
PCT1 = NUMFMT["pct1"]          # Renditen und Quoten: 5,3 %
PCT2 = NUMFMT["pct2"]          # Zinsen, Steuer- und AfA-Sätze: 3,57 %
MULT1 = NUMFMT["mult1"]        # Kaufpreisfaktor 25,0×
MULT2 = NUMFMT["mult2"]        # Eigenkapital-Multiplikator 1,98× · DSCR-artige Faktoren
DSCR = NUMFMT["dscr"]
EUR_M2 = NUMFMT["eur_qm"]      # 12,50 €/m²
M2 = NUMFMT["m2"]              # 68 m²
YEARS = NUMFMT["years_calc"]   # 1 Jahr · 12 Jahre · –
MONTHS = NUMFMT["months_calc"]
FORMATS = {"EUR": EUR, "EUR2": EUR2, "PCT1": PCT1, "PCT2": PCT2, "MULT1": MULT1, "MULT2": MULT2, "DSCR": DSCR,
           "EUR_M2": EUR_M2, "M2": M2, "YEARS": YEARS, "MONTHS": MONTHS}


def zero_dash(fmt):
    """Zahlenformat für BERECHNETE Werte: fehlender Nullabschnitt wird „–“ (P2-07); Negativ-Abschnitt mit „−“.
    Formate mit Bedingungen, Text, Datum und bereits dreiteilige Formate bleiben unverändert. Idempotent."""
    f = typo_minus(fmt)
    if not isinstance(f, str) or not f or "[" in f or "@" in f or f in ("General", ";;;"):
        return f
    low = f.lower()
    if any(k in low for k in ("yy", "dd", "mm", "hh")):
        return f
    parts = _split_fmt(f)
    if len(parts) == 2 and any(c in parts[0] for c in "0#"):
        return f + ';"–"'
    return f


def _split_fmt(fmt):
    secs, cur, q = [], "", False
    for ch in fmt:
        if ch == '"':
            q = not q
        if ch == ";" and not q:
            secs.append(cur)
            cur = ""
        else:
            cur += ch
    secs.append(cur)
    return secs


def fixed_m(expr, digits=0):
    """FIXED mit typografischem Minus (P2-07 / P3-15) für Anzeigeformeln: 'SUBSTITUTE(FIXED(x,0),"-","−")'.
    Beispiel: '="Jahr 2: "&' + fixed_m(CF_YEAR2) + '&" €"' → „Jahr 2: −274 €“."""
    return f'SUBSTITUTE(FIXED({expr},{digits}),"-","{MINUS}")'


def minus_text(formula):
    """Anzeigeformel umschreiben: jedes FIXED(…) / TEXT(…), das nicht schon in SUBSTITUTE steht, wird mit
    SUBSTITUTE(…,"-","−") umschlossen – typografisches Minus auch in Textverkettungen (P2-07). Nur für NEUE
    Anzeigeformeln (Kachel-Fußzeilen, Pill-Suffixe), nie für Formeln der Vorlage. Idempotent; Nicht-Formeln unverändert."""
    if not is_formula(formula) or not re.search(r"\b(FIXED|TEXT)\(", formula, re.I):
        return formula
    out, i, n = "", 0, len(formula)
    q = False
    while i < n:
        ch = formula[i]
        if ch == '"':
            q = not q
            out += ch
            i += 1
            continue
        m = None if q else re.match(r"(FIXED|TEXT)\(", formula[i:], re.I)
        prev = formula[i - 1] if i else ""
        if m and not (prev.isalnum() or prev in "._"):
            j, lvl, qq = i + len(m.group(0)), 1, False
            while j < n and lvl:
                c2 = formula[j]
                if c2 == '"':
                    qq = not qq
                elif not qq and c2 == "(":
                    lvl += 1
                elif not qq and c2 == ")":
                    lvl -= 1
                j += 1
            call = formula[i:j]
            datefmt = m.group(1).upper() == "TEXT" and re.search(r'"[^"]*[YyDdMmHh][^"]*"\s*\)$', call)
            if datefmt or re.search(r"SUBSTITUTE\($", out, re.I):
                out += call
            else:
                out += f'SUBSTITUTE({call},"-","{MINUS}")'
            i = j
            continue
        out += ch
        i += 1
    return out


def typo_minus(fmt):
    """Zahlenformat auf das typografische Minus umstellen (P23): In jeder Negativ-Sektion wird ein führendes
    „-“ / „\\-“ / „"-"“ zu „"−"“. Einsektionige Zahlenformate (z. B. #,##0" €") erhalten eine Negativ-Sektion.
    Text-, Datums-, Bedingungs- und Ausblendformate bleiben unverändert. Idempotent."""
    if not isinstance(fmt, str) or not fmt or fmt in ("General", "@", ";;;") or "[" in fmt:
        return fmt
    low = fmt.lower()
    if any(k in low for k in ("yy", "dd", "mm", "hh")) or "@" in fmt:
        return fmt
    secs, cur, q = [], "", False
    for ch in fmt:
        if ch == '"':
            q = not q
        if ch == ";" and not q:
            secs.append(cur)
            cur = ""
        else:
            cur += ch
    secs.append(cur)
    if not any(c in secs[0] for c in "0#"):
        return fmt
    if len(secs) == 1:
        return f"{secs[0]};{_M}{secs[0]}"
    s2 = secs[1]
    for old in ('"-"', "\\-", "-"):
        if s2.startswith(old):
            s2 = _M + s2[len(old):]
            break
    secs[1] = s2
    return ";".join(secs)


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
    "DSCR": dict(label="DSCR", fmt=NUMFMT["dscr"], rule="ampel", green="Ampel_DSCR_gruen", yellow="Ampel_DSCR_gelb"),
    "EKR": dict(label="EK-Rendite Jahr 1", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_EKR_gruen", yellow="Ampel_EKR_gelb"),
    "IRR": dict(label="IRR n. St.", fmt=NUMFMT["pct1"], rule="ampel", green="Ampel_IRR_gruen", yellow="Ampel_IRR_gelb"),
    "FAKTOR": dict(label="Kaufpreisfaktor", fmt=NUMFMT["mult1"]),
    "MULT": dict(label="Eigenkapital-Multiplikator", fmt=NUMFMT["mult2"]),   # Runde 4 (P2-07): einheitliches Label
    "RF": dict(label="Rechtsform", fmt="@"),
}
CF_YEAR2 = "INDEX(Projektion!$D$37:$AQ$37,2)"

# Kanonische KPI-Beschriftungen und -Reihenfolge (P20, Glossar) – wörtlich auf allen Blättern.
# Start/Leitfaden zeigen eine Teilmenge in dieser Reihenfolge. IRR trägt die Haltedauer (kpi_label('IRR', formula=True)).
KPI_ORDER = ("GI", "EK", "CF", "BMR", "DSCR", "IRR")
KPI_LABELS = {
    "GI": "Gesamtinvestition",
    "EK": "Eigenkapitalbedarf inkl. Reserve",
    "CF": "Cashflow n. St. / Monat (Jahr 1)",
    "BMR": "Bruttomietrendite",
    "DSCR": "DSCR",
    "IRR": "IRR n. St.",
    "RATE": "Rate an die Bank / Monat",
    "NMR": "Nettomietrendite",
    "EKR": "EK-Rendite Jahr 1",
}
KPI_LABEL_FORMULAS = {"IRR": '"IRR n. St. ("&Haltedauer&" J.)"'}
KPI_LABELS["MULT"] = "Eigenkapital-Multiplikator"
KPI_LABELS["FAKTOR"] = "Kaufpreisfaktor"
# Kachel-Labels (Runde 4, P1-02 / P3-19): In JEDER Kachel der Mappe steht dieselbe Kurzform, damit das Label bei
# 8,5 pt auch in schmalen Kacheln (Dashboard 180 px) mit Reserve passt. Der Zusatz „Jahr 1“ gehört in die Fußzeile.
# tile() setzt die Kurzform automatisch – auch wenn die Langform als Text (in beliebiger Schreibung) übergeben wird.
KPI_TILE_LABELS = {
    "EK": "EK-Bedarf inkl. Reserve",
    "CF": "Cashflow n. St. / Monat",
    "CFV": "Cashflow v. St. / Monat",
}
_TILE_SHORT = {KPI_LABELS.get(k, KPI[k]["label"]).upper(): v for k, v in KPI_TILE_LABELS.items()}
_TILE_SHORT.update({KPI[k]["label"].upper(): v for k, v in KPI_TILE_LABELS.items()})
_TILE_SHORT.update({"EIGENKAPITALBEDARF INKL. RESERVE": KPI_TILE_LABELS["EK"],
                    "CASHFLOW V. ST. / MONAT (JAHR 1)": KPI_TILE_LABELS["CFV"],
                    "CASHFLOW N. ST. / MONAT (JAHR 1)": KPI_TILE_LABELS["CF"],
                    "EIGENKAPITAL-MULTIPLE": KPI_LABELS["MULT"]})


def kpi_label(key, caps=False, formula=False, tile=False):
    """Kanonische Beschriftung (P20). formula=True liefert für IRR die Anzeigeformel mit Haltedauer
    (='IRR n. St. (12 J.)'), caps=True Versalien (bei Formeln über UPPER). tile=True: Kachel-Kurzform
    (KPI_TILE_LABELS, Runde 4 – „EK-Bedarf inkl. Reserve“, „Cashflow n. St. / Monat“)."""
    if formula and key in KPI_LABEL_FORMULAS:
        expr = KPI_LABEL_FORMULAS[key]
        return f"=UPPER({expr})" if caps else f"={expr}"
    t = (KPI_TILE_LABELS.get(key) if tile else None) or KPI_LABELS.get(key) or KPI[key]["label"]
    return t.upper() if caps else t
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


# Brotkrumen (Runde 4, P2-04): „REITER › BLATT“ – nur „›“ trennt Hierarchieebenen, „·“ nur innerhalb einer Ebene.
# page_header() setzt die kanonische Krume für diese Blätter automatisch (die übergebene Eyebrow wird ersetzt).
CRUMBS = {
    "Dashboard": ("Dashboard",), "Cockpit": ("Cockpit",), "Diagramme": ("Diagramme",),
    "Projektion": ("Berechnung", "Projektion"), "Steuern": ("Berechnung", "Steuer-Tabelle"),
    "Finanzierung": ("Berechnung", "Tilgungsplan"), "Sensitivität": ("Berechnung", "Sensitivität"),
    "AfA-Vergleich": ("Steuer-Tabelle", "AfA-Vergleich"),
    "Bankgespräch": ("Bank", "Bankgespräch"), "Haushaltsrechnung": ("Bank", "Haushaltsrechnung"),
    "Vermögensaufstellung": ("Bank", "Vermögensaufstellung"),
    "Hinweise": ("Anhang", "Hinweise"), "Konfiguration": ("Anhang", "Konfiguration"),
}
# Elternblatt je Reiter (Ziel des ersten, verlinkten Krumenteils und des Rücksprungs „‹ Elternblatt“ rechts in Z. 5).
CRUMB_PARENT = {"LEITFADEN": "Leitfaden", "BERECHNUNG": "Dashboard", "STEUER-TABELLE": "Steuern", "BANK": "Bankgespräch",
                "ANHANG": "Start", "DASHBOARD": "Start", "COCKPIT": "Start", "DIAGRAMME": "Start", "EINGABEN": "Start"}
PARENT = {"Dashboard": "Start", "Cockpit": "Start", "Diagramme": "Start", "Eingaben": "Start",
          "Projektion": "Dashboard", "Steuern": "Dashboard", "Finanzierung": "Dashboard", "Sensitivität": "Dashboard",
          "AfA-Vergleich": "Steuern", "Haushaltsrechnung": "Bankgespräch", "Vermögensaufstellung": "Bankgespräch",
          "Bankgespräch": "Start", "Hinweise": "Start", "Konfiguration": "Start", "Leitfaden": "Start"}
_STEP_TITLE = re.compile(r"^S(\d\d)\s+(.+)$")


def sheet_label(sheet):
    """Sprechender Name eines Blatts für ScreenTips: „S02 Kaufpreis & Miete“ → „Schritt 02 · Kaufpreis & Miete“,
    „Start“ → „Startseite“, „Steuern“ → „Steuer-Tabelle“, „Finanzierung“ → „Tilgungsplan“."""
    m = _STEP_TITLE.match(sheet or "")
    if m:
        return f"Schritt {m.group(1)} · {m.group(2)}"
    return {"Start": "Startseite", "Steuern": "Steuer-Tabelle", "Finanzierung": "Tilgungsplan"}.get(sheet, sheet)


def auto_tooltip(text=None, target_sheet=None, target_cell=None, source_sheet=None):
    """ScreenTip für einen Zell-Link (P3-17) – dieselbe Sprache wie die Formen-Tooltips:
    „Weiter zu Schritt 02 · Kaufpreis & Miete“, „Zurück zur Startseite“, „Zum Abschnittsanfang“, „Zum Dashboard“."""
    t = str(text or "").strip()
    if t.startswith("="):
        t = ""
    same = source_sheet is not None and target_sheet == source_sheet
    if "↑" in t:
        return "Zum Seitenanfang" if (not same or not target_cell or target_cell == link_target(target_sheet)) \
            else "Zum Abschnittsanfang"
    lab = sheet_label(target_sheet)
    clean = re.sub(r"^[‹›\s]+|[‹›\s]+$", "", t).strip()
    if same:
        down = "↓" in t
        clean = re.sub(r"[↓↑▾▼›‹]", "", clean).strip()
        if not clean:
            return "Zum Abschnittsanfang"
        return f"Zum Abschnitt „{clean}“" if down and ":" not in clean else f"Auf diesem Blatt: {clean}"
    if target_sheet == "Start" and target_cell and target_cell != link_target("Start"):
        return "Zur Eingabe auf der Startseite"
    art = " zur " if target_sheet == "Start" else (" zu " if _STEP_TITLE.match(target_sheet or "") else ": ")
    if t.startswith("‹"):
        return f"Zurück{art}{lab}"
    if t.endswith("›"):
        return f"Weiter{art}{lab}"
    return f"Öffnet{art if art != ': ' else ': '}{lab}".replace("Öffnet zur ", "Zur ").replace("Öffnet zu ", "Zu ")


def _hyperlink(cell, target_sheet, target_cell=None, text=None, tooltip=None):
    """Zell-Hyperlink mit IMMER gesetztem ScreenTip (P3-17)."""
    tip = tooltip or auto_tooltip(text, target_sheet, target_cell, cell.parent.title if cell.parent is not None else None)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=link_loc(target_sheet, target_cell), display=text,
                               tooltip=tip[:255])
    return cell.hyperlink


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
    try:
        est = _branch_text(f, wb, depth)
        if est is not None:
            return est
    except Exception:
        pass
    if "&" in f:
        refs = len(re.findall(r"&", f)) + 1 - len(lits)
        return "".join(lits) + "0" * 8 * max(refs, 0)
    return max(lits, key=len)


def _split_top(expr, sep):
    """Ausdruck an sep auf oberster Ebene teilen (Klammern und Anführungszeichen beachtet)."""
    parts, cur, lvl, q = [], "", 0, False
    for ch in expr:
        if ch == '"':
            q = not q
        elif not q and ch == "(":
            lvl += 1
        elif not q and ch == ")":
            lvl -= 1
        if ch == sep and lvl == 0 and not q:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    return parts


def _branch_text(expr, wb=None, depth=0):
    """Längster möglicher Anzeigetext einer Formel, zweigweise (Wunsch B, Runde 2): IF/IFERROR/CHOOSE werden je Zweig
    bewertet, &-Verkettungen stückweise (Literal = Text, Name = aktueller Wert, sonst 8 Ziffern)."""
    e = expr.strip()
    while e.startswith("(") and e.endswith(")") and len(_split_top(e[1:-1], ",")) == 1:
        e = e[1:-1].strip()
    m = re.match(r"(IF|IFERROR|CHOOSE)\((.*)\)$", e, re.S | re.I)
    if m and len(_split_top(m.group(2), ",")) >= 2 and _split_top(e, "&") == [e]:
        args = _split_top(m.group(2), ",")
        fn = m.group(1).upper()
        cands = args[1:] if fn in ("IF", "CHOOSE") else args
        texts = [_branch_text(a, wb, depth) or "" for a in cands]
        return max(texts, key=lambda t: text_width(t, T_SMALL)) if texts else ""
    parts = _split_top(e, "&")
    out = ""
    for p in parts:
        p = p.strip()
        if re.fullmatch(r'"((?:[^"]|"")*)"', p):
            out += p[1:-1].replace('""', '"')
        elif re.fullmatch(r"(IF|IFERROR|CHOOSE)\(.*\)", p, re.S | re.I):
            out += _branch_text(p, wb, depth) or ""
        elif re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", p) and not re.fullmatch(r"[A-Z]{1,3}\d+", p):
            v = _name_value(wb, p, depth + 1)
            out += str(v) if isinstance(v, str) else "0" * 8
        elif p:
            out += "0" * 8
    return out


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


def fit_row(ws, row, c1, c2, base=H_ROW, pad=6, grid=True, max_lines=None, hints=None, min_height=None,
            metric=False, valign_top=True):
    """Empfohlene Zeilenhöhen-Passung (Runde 2): Calibri-Metrik, Formel-Auflösung, Einzug, Zeilenhöhe je
    Schriftgröße (line_pt). grid=True rastet auf 18 · 30 · 42 · n×13+8 ein. Liefert die Zeilenzahl.
    metric=True (Runde 3, P27 – Fließtexttabellen wie Hinweise/Konfiguration): Höhe = Zeilen × 12,5 + 8 pt
    (1 Z. 21 · 2 Z. 33 · 3 Z. 45,75 · 4 Z. 58,5 pt), umbrechende Zellen oben ausgerichtet (valign_top)."""
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
        if metric and valign_top and cell.alignment is not None and cell.alignment.wrap_text:
            a = cell.alignment
            cell.alignment = Alignment(horizontal=a.horizontal, vertical="top", indent=a.indent, wrap_text=True,
                                       shrink_to_fit=False)
    if metric:
        h = text_row_height(lines)
    else:
        h = grid_height(need_pt, base) if grid else need_pt
    if min_height:
        h = max(h, min_height)
    ws.row_dimensions[row].height = h
    return lines


# =============================================================================== Überschriften
def band_l1(ws, row, c1, c2, title=None, right=None, height=H_BAND):
    """L1 Abschnitt (Runde 5): Band F7F6F2, GOLDKANTE links 3 px (C9A14A), Unterkante thin D6D2C6,
    12,5 pt fett Nachtblau."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.fill = fill(TINT)
        c.border = Border(left=side("thick", GOLD) if k == 0 else None, bottom=side("thin", LINE2))
    first = ws.cell(row, col(c1))
    if title is not None:
        set_text(first, title)
    first.font = font(T_H3, True, NAVY, DISPLAY)
    first.alignment = align("left", "center", 1)
    if right is not None:
        last = ws.cell(row, col(c2))
        set_text(last, right)
        last.font = font(T_LABEL, False, BLUE)
        last.alignment = align("right", "center", 1)
    set_height(ws, row, height)


def band_l1b(ws, row, c1, c2, title=None, right=None, height=H_BAND_B):
    """L1b Tabellenabschnitt (Runde 5): Band F7F6F2, 9 pt fett Versalien Nachtblau, Goldkante links, Unterkante D6D2C6."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.fill = fill(TINT)
        c.border = Border(left=side("thick", GOLD) if k == 0 else None, bottom=side("thin", LINE2))
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
    extra: Zusatz hinter dem Titel als Rich-Text (9 pt A9B8CC), z. B. „· Pflichtangaben“.
    Hinweis Runde 2 (P1-12): Navy-Vollflächen sind Kopfleiste und Kachel-Labels vorbehalten –
    für neue Abschnitte section(level=1, extra=…) verwenden."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = fill(NAVY)
        c.border = Border(bottom=side("medium", GOLD))      # Runde 5: Goldlinie unter dem Nachtblau-Band
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
    """Spaltenkopf über volle Tabellenbreite: F1EEE6, 8,5 pt fett Versalien Blau (P40), Unterkante thin Akzent.
    labels: {spalte: (text, 'left'|'right'|'center')} – Köpfe folgen der Ausrichtung ihrer Werte."""
    for c in iter_cells(ws, c1, row, c2, row):
        c.fill = fill(HEAD)
        c.border = Border(bottom=side("thin", ACCENT))
        c.font = font(T_LABEL, True, BLUE)
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
    """Alt-Summenstufen, seit Runde 3 auf die EINE Summenhierarchie abgebildet (P12):
    level 1/2 → sum_row('sub') (fett Navy, ohne Fläche, Oberlinie CFC9BA) · level 3 → sum_row('final')."""
    sum_row(ws, row, c1, c2, "final" if level >= 3 else "sub", neg=False)


def memo(ws, row, c1, c2, indent=2, clear_lines=False):
    """Nachrichtliche Zeile / Memo (P12): 9 pt kursiv 4E5561, ohne Fläche und Zusatzlinien, Beschriftung Einzug 2.
    clear_lines=True (sum_row('memo'), Runde 4 P1-10): auch ohne Rahmenlinien – die Doppellinie der Endsumme darüber
    bleibt der sichtbare Blockabschluss."""
    for k, c in enumerate(iter_cells(ws, c1, row, c2, row)):
        c.font = font(T_SMALL, False, MUTED, italic=True)
        c.fill = NOFILL
        if clear_lines:
            c.border = Border()
        if k == 0 and isinstance(c.value, str) and not (c.alignment and c.alignment.horizontal in ("right", "center")):
            c.alignment = align("left", "center", indent, wrap=bool(c.alignment and c.alignment.wrap_text))


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


def status_formula(conditions, words=None, dot=True, empty="", suffix=None):
    """Anzeigeformel für das Statuswort: =IF(b1,"●  kritisch",IF(b2,"●  prüfen",…,"")).
    words: {'red': …, 'amber': …, 'green': …} (Standard STATUS_WORDS – nur „erfüllt / prüfen / kritisch“, P21).
    suffix: Excel-Ausdruck für den Wert hinter dem Wort („● prüfen · 4,0 %“), z. B. 'FIXED(Bruttomietrendite*100,1)&" %"'."""
    words = {**STATUS_WORDS, **(words or {})}
    expr = '"' + empty.replace('"', '""') + '"'
    tail = f'&"  ·  "&{suffix}' if suffix else ""
    for cond, lvl in reversed(list(conditions)):
        w = (f"{STATUS_DOT}  " if dot else "") + words[lvl]
        expr = f'IF({cond},"{w}"{tail},{expr})'
    return "=" + expr


def cf_rule(ws, ref, formula, font_=None, fill_=None, border=None, numfmt=None, stop=True):
    """Bedingte Formatierung mit optionalem Zahlenformat (dxf numFmt) – Wunsch D: „entfällt“ → „–“,
    Status-Chip in 1-spaltigen Kacheln. formula ohne „=“; numfmt als Formatcode."""
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.styles.numbers import NumberFormat
    nf = None
    if numfmt:
        cache = ws.parent.__dict__.setdefault("_core_dxf_numfmt", {}) if ws.parent is not None else {}
        if numfmt not in cache:
            cache[numfmt] = 900 + len(cache)
        nf = NumberFormat(numFmtId=cache[numfmt], formatCode=numfmt)
    dxf = DifferentialStyle(font=font_, fill=fill_, border=border, numFmt=nf)
    ws.conditional_formatting.add(ref, Rule(type="expression", formula=[formula], dxf=dxf, stopIfTrue=stop))


def status_cf(ws, ref, conditions, font_color=True, bold=True, fill_bg=False, fill_white=False, border=None,
              border_style="thin", border_color="strong", numfmt=None):
    """Bedingte Formate je Status auf ref.
    font_color: Schrift in Statusfarbe · fill_bg: zarter Status-Tint (nur Pill/kleine Flächen!) · fill_white: weiß
    border: None | 'box' | 'left' | 'bottom' | 'top' · border_color: 'strong' (Statusfarbe) | 'line' (helle Linie)
    numfmt: {lvl: Formatcode} – bedingtes Zahlenformat je Status (z. B. Status-Chip rechts in der Fußzeile)."""
    for cond, lvl in conditions:
        fg, bg, ln = STATUS_COLORS[lvl]
        f_ = Font(color=fg, bold=bold) if font_color else None
        fl = fill(bg) if fill_bg else (fill(WHITE) if fill_white else None)
        bd = None
        if border:
            sd = side(border_style, fg if border_color == "strong" else ln)
            bd = Border(**({"left": sd, "right": sd, "top": sd, "bottom": sd} if border == "box" else {border: sd}))
        if numfmt and numfmt.get(lvl):
            cf_rule(ws, ref, cond, font_=f_, fill_=fl, border=bd, numfmt=numfmt[lvl])
            continue
        kw = {"stopIfTrue": True}
        if f_ is not None:
            kw["font"] = f_
        if fl is not None:
            kw["fill"] = fl
        if bd is not None:
            kw["border"] = bd
        ws.conditional_formatting.add(ref, FormulaRule(formula=[cond], **kw))


def status_pill(ws, cell, kpi=None, value_ref=None, conditions=None, formula=None, style="pill", words=None,
                h=None, indent=1, suffix=None):
    """Status als Wort mit Punkt in Statusfarbe (Anzeigeformel in einer LEEREN Zelle) – EINE Pill (P21).
    style 'pill': 9 pt fett, zentriert, bei Status zarter Tint der Statusfarbe (EAF5EF / FDF3E7 / FBECEB), KEIN Rahmen –
      liegt so innerhalb jeder Box-/Kachelfläche (Einordnungs-Box, Leitfaden-Kacheln, Link-Reihen, Banner).
      Breitere Pill: Zelle vorher über mehrere Spalten verbinden (callout_box pill_col=…).
    style 'chip': 8 pt fett, nur Schriftfarbe, rechtsbündig (Kachel-Fußzeile, Statusspalte in Tabellen).
    suffix: Wert hinter dem Wort („● prüfen · 4,0 %“), siehe status_formula."""
    conds = status_conditions(kpi, value_ref, conditions) if (kpi or conditions) else []
    if formula is not None:
        cell.value = formula
    elif conds:
        cell.value = status_formula(conds, words, suffix=suffix)
    cell.number_format = "General"
    pill = style == "pill"
    h = h or ("center" if pill else "right")
    cell.font = font(T_SMALL if pill else T_MICRO, True, MUTED)
    cell.alignment = align(h, "center", indent if h in ("left", "right") else 0)
    if conds:
        status_cf(ws, cell.coordinate, conds, font_color=True, fill_bg=pill)
    return conds


def status_edge(ws, ref, kpi=None, value_ref=None, conditions=None, edge="left", style="thick"):
    """Statuskante: Rahmen einer Seite (Standard: linker 3-px-Balken) in Statusfarbe, per bedingter Formatierung.
    Hinweis: Excel zeigt bedingte Rahmen teils nur dünn – deshalb immer mit statischer Akzentkante kombinieren."""
    conds = status_conditions(kpi, value_ref, conditions) if (kpi or conditions) else []
    if conds:
        status_cf(ws, ref, conds, font_color=False, border=edge, border_style=style)
    return conds


def status_banner(ws, c1, r1, c2, r2, kpi=None, value_ref=None, conditions=None, pill_cell=None, words=None,
                  suffix=None, pill_formula=None):
    """Signalfarben nicht als Fläche (P41): Banner/Hinweiszeile auf FBFAF7 mit statischer 3-px-Akzentkante links,
    die per bedingter Formatierung in die Statusfarbe wechselt; optional EINE Pill (pill_cell, leer) mit Status-Tint.
    Ersetzt rosa/grüne Vollflächen (Dashboard „Gesamtbewertung“, Cockpit-Prüfhinweise)."""
    c1i, c2i = col(c1), col(c2)
    for r in range(r1, r2 + 1):
        for c in iter_cells(ws, c1i, r, c2i, r):
            c.fill = fill(TINT_XL)
            c.border = Border(left=side("thick", ACCENT) if c.column == c1i else None)
    conds = status_conditions(kpi, value_ref, conditions) if (kpi or conditions) else []
    if conds:
        status_edge(ws, rng(c1i, r1, c1i, r2), conditions=conds)
        if pill_cell is not None:
            status_pill(ws, ws[pill_cell] if isinstance(pill_cell, str) else pill_cell, conditions=conds,
                        words=words, suffix=suffix, formula=pill_formula)
    return conds


def status_legend(labels=("erfüllt", "prüfen", "kritisch"), size=T_MICRO):
    """Rich-Text-Legende „● erfüllt · ● prüfen · ● kritisch“ (Punkte in Statusfarbe, Text 4E5561) – P21: nur diese drei Wörter."""
    parts = []
    for i, (lvl, lab) in enumerate(zip(("green", "amber", "red"), labels)):
        if i:
            parts.append(("   ·   ", size, False, MUTED))
        parts += [(STATUS_DOT + " ", size, True, STATUS_COLORS[lvl][0]), (lab, size, False, MUTED)]
    return rich(parts)


# Eine Schwellenformulierung (P21): „Ziel ≥ 1,20×“ bzw. „< 1,00× kritisch · 1,00–1,20× prüfen · ≥ 1,20× erfüllt“.
def _fmt_expr(kpi, name):
    f = KPI[kpi].get("fmt", "")
    if "%" in f:
        return f'FIXED({name}*100,1)&" %"'
    if "×" in f:
        return f'FIXED({name},2)&"×"'
    return f'FIXED({name},0)&" €"'


def threshold_text(kpi, prefix="Ziel ≥ "):
    """Excel-Ausdruck (ohne „=“) für die Zielschwelle: '"Ziel ≥ "&FIXED(Ampel_DSCR_gruen,2)&"×"'.
    KPI mit Regel cf/neg: '"Ziel ≥ 0 €"'. Für eine Anzeigeformel: '=' + threshold_text('DSCR')."""
    spec = KPI[kpi]
    if spec.get("rule") == "ampel":
        return f'"{prefix}"&{_fmt_expr(kpi, spec["green"])}'
    return f'"{prefix}0 €"'


def threshold_legend(kpi):
    """Anzeigeformel der Schwellen-Legende: ="< 1,00× kritisch · 1,00–1,20× prüfen · ≥ 1,20× erfüllt"."""
    spec = KPI[kpi]
    if spec.get("rule") != "ampel":
        return '="< 0 € kritisch  ·  ≥ 0 € erfüllt"'
    y, g = _fmt_expr(kpi, spec["yellow"]), _fmt_expr(kpi, spec["green"])
    return (f'="< "&{y}&" kritisch  ·  "&{y}&"–"&{g}&" prüfen  ·  ≥ "&{g}&" erfüllt"')


# Negativ-Rot-Regel (P15, mappenweit): Negative €-Ergebnis- und Kumulwerte in Herleitungen, Summen und Kennzahlen
# sind rot B42318 (neg_red); Eingaben und Bestandsgrößen bleiben neutral. sum_row('sub'/'final') setzt die Regel selbst.
NEG_RULE_TEXT = "Rot – negative Ergebnis- und Kumulwerte"


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
    Idempotent; am Ende eines Blatt-Moduls bzw. in global_rules.final aufrufen.
    Runde 4: ruft zuerst finalize_components(ws) auf (feste Kachelhöhen, ScreenTips) – dadurch greifen die
    Komponentenregeln auf jedem Blatt, auch wenn ein Modul Höhen nach tile() überschreibt."""
    try:
        finalize_components(ws)
    except Exception as exc:  # nie den Build brechen
        import sys
        print(f"WARNUNG core.finalize_components {ws.title}: {exc!r}", file=sys.stderr)
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
    """Alt-API – seit Runde 4 (P1-02) nur noch ein Aufruf von tile(): dieselbe Kachel-Spezifikation (18/30/20 pt,
    8,5-pt-Versal-Label, Wert 20 pt, Kopffarbe nach Blattrolle). Status wie bisher über die KPI-Regel (Wertfarbe)."""
    return tile(ws, c1, c2, label_row, value_row, sub_row, label=label, value=value, sub=sub, kpi=kpi, fmt=fmt,
                value_rows=value_rows, gap_right=gap_right, value_size=value_size, status="edge" if not sub_row else "auto")


def button(ws, c1, row, c2, text, target_sheet=None, kind="primary", target_cell=None, tooltip=None):
    """Alt-API. Seit Runde 4 (P1-03) identisch mit btn(): primary / secondary / tertiary (25,5 pt, 10 pt fett, zentriert).
    kind='link' bleibt der Textlink ohne Rahmen (10 pt Blau, links). Jeder Link trägt einen ScreenTip (P3-17)."""
    if kind != "link":
        btn(ws, c1, row, c2, text, target_sheet, kind if kind in BTN or kind in BTN_TEXT else "secondary",
            target_cell, tooltip)
        return
    for c in iter_cells(ws, c1, row, c2, row):
        c.border = Border()
        c.fill = NOFILL
    safe_merge(ws, c1, row, c2, row)
    c = ws.cell(row, col(c1))
    set_text(c, text)
    c.font = font(T_BODY, False, BLUE)
    c.alignment = align("left", "center", 1)
    if target_sheet:
        _hyperlink(c, target_sheet, target_cell, text, tooltip)
    set_height(ws, row, max(ws.row_dimensions[row].height or 0, H_ROW))


def text_link(cell, text, target_sheet, target_cell=None, size=T_BODY, bold=True, tooltip=None):
    """Textlink (Blau, ohne Rahmen). ScreenTip immer gesetzt (P3-17): tooltip oder auto_tooltip(text, Ziel)."""
    set_text(cell, text)
    cell.font = font(size, bold, BLUE)
    _hyperlink(cell, target_sheet, target_cell, text, tooltip)


def callout(ws, c1, head_row, c2, body_r1, body_r2, title="Einordnung", status=None, status_col=None):
    """Einordnung: ein Akzentbalken (thick 3B6A9E) für Kopf und Text; Kopf 9 pt fett Navy,
    Text 9,5→10 pt INK2 auf FBFAF7, umbrechend, vertikal zentriert."""
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


INPUT_STATES = (   # Legende der Eingabezustände (Runde 4, P3-20) – Reihenfolge und Wortlaut für Start/Leitfaden
    ("required", "Eingabe"),
    ("optional", "optional / leer"),
    ("override", "aus Kalkulation, überschreibbar"),
    ("inactive", "inaktiv für diese Methode"),
)
INACTIVE_PREFIX = "inaktiv – "   # einheitliches Präfix in der Hinweisspalte inaktiver Felder (P3-20)


def input_style(cell, state="required"):
    """Eingabezustände (P3-20): required (FFF5D6, Rahmen thin E6CB77) · optional (FFF5D6, Rahmen dashed) ·
    override (aus Kalkulation, überschreibbar: FFF9EA, Rahmen dashed E6CB77, Schrift 1E4E8C normal) ·
    inactive (F2F1ED, grau) · linked (weiß, dashed Akzent, Blau normal)."""
    if state in ("required", "optional"):
        ln = side("thin" if state == "required" else "dashed", INPUT_LINE)
        cell.fill = fill(INPUT_BG)
        cell.font = font(cell.font.sz or T_BODY, True, INPUT_FG)
        cell.border = Border(left=ln, right=ln, top=ln, bottom=ln)
        cell.protection = Protection(locked=False)
    elif state == "override":
        ln = side("dashed", INPUT_LINE)
        cell.fill = fill(OVERRIDE_BG)
        cell.font = font(cell.font.sz or T_BODY, False, INPUT_FG)
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


def input_legend(ws, row, cells, size=T_SMALL):
    """Legende der vier Eingabezustände (P3-20): cells = [(Musterzelle, Textzelle), …] in der Reihenfolge von
    INPUT_STATES. Die Musterzelle erhält den Zustand (ohne Text, Blattschutz bleibt), die Textzelle das Wort 9 pt 4E5561."""
    for (swatch, label), (state, word) in zip(cells, INPUT_STATES):
        sw = ws[swatch] if isinstance(swatch, str) else swatch
        tx = ws[label] if isinstance(label, str) else label
        input_style(sw, state)
        sw.protection = Protection(locked=True)
        set_text(tx, word)
        tx.font = font(size, False, MUTED)
        tx.alignment = align("left", "center", 1)


def inactive_when(ws, ref, condition):
    """Eingabefeld bedingt inaktiv darstellen (bleibt editierbar)."""
    ws.conditional_formatting.add(ref, FormulaRule(formula=[condition], stopIfTrue=True,
                                                   font=Font(color=INACTIVE_FG, bold=False),
                                                   fill=fill(INACTIVE_BG),
                                                   border=Border(left=side("thin", LINE2), right=side("thin", LINE2),
                                                                 top=side("thin", LINE2), bottom=side("thin", LINE2))))


H_HDR = {4: 12, 5: 18, 6: 30, 7: 21.75}   # Kopfschablone (Runde 4, P2-03): feste Zeilenhöhen auf ALLEN Blättern


def crumb(sheet, eyebrow=None):
    """Kanonische Brotkrume „REITER  ›  BLATT“ (P2-04). S01–S12: „LEITFADEN  ›  SCHRITT 03 / 12  ·  KAUFNEBENKOSTEN“.
    Unbekannte Blätter: die übergebene Eyebrow (Tupel → mit „  ›  “ verbunden)."""
    if sheet in CRUMBS:
        parts = CRUMBS[sheet]
    elif isinstance(eyebrow, (tuple, list)):
        parts = tuple(str(x) for x in eyebrow if x)
    else:
        parts = (str(eyebrow or ""),)
    if _STEP_TITLE.match(sheet or "") and parts and not str(parts[0]).upper().startswith("LEITFADEN"):
        parts = ("Leitfaden",) + tuple(parts)
    return "  ›  ".join(p for p in parts if p).upper()


def page_header(ws, c1, c2, eyebrow, title, subtitle=None, context=None, context_col=None, right_legend=None,
                legend_col=None, context_rows=(6, 7), back=None, back_col=None, canonical=True):
    """Seitenkopf-Vorlage (Runde 3 P04, Runde 4 P2-03/P2-04) – für alle Blätter gleich:
      Zeilenhöhen FEST: Z. 4 = 12 pt · Z. 5 = 18 pt · Z. 6 = 30 pt · Z. 7 = 21,75 pt (H_HDR).
      Z. 5 links Brotkrume „REITER  ›  BLATT“ (8,5 pt fett Versalien 1E4E8C). canonical=True setzt für die Blätter in
           CRUMBS (und S01–S12 mit Präfix „Leitfaden“) die kanonische Krume, egal was übergeben wird.
           Rechts: Unterreiter (Formen) ODER Rücksprung back=„Start“ → „‹  Start“ (8,5 pt 1E4E8C, Zell-Link, rechtsbündig in
           back_col bzw. c2); back=True nimmt das Elternblatt aus PARENT.
      Z. 6 links Titel 22 pt fett 0E2238; rechts context[0] (10 pt fett 0E2238, z. B. „Beispiel: …“).
      Z. 7 links Untertitel 10 pt 4E5561; rechts context[1] („Erstellt für …“, 9 pt 4E5561) ODER right_legend (8,5 pt).
    Alle rechten Elemente rechtsbündig OHNE Einzug in context_col/legend_col (Standard c2) – eine rechte Inhaltskante."""
    text = crumb(ws.title, eyebrow) if canonical else (
        "  ›  ".join(str(x) for x in eyebrow if x) if isinstance(eyebrow, (tuple, list)) else str(eyebrow)).upper()
    e, t, s = ws.cell(5, col(c1)), ws.cell(6, col(c1)), ws.cell(7, col(c1))
    set_text(e, text)
    e.font = font(T_LABEL, True, GOLD_INK)      # Runde 5: Bronze-Eyebrow (6,2:1)
    e.alignment = align("left", "bottom")
    if title is not None:
        if not is_formula(t.value) or not is_formula(title):
            t.value = title
    t.font = font(T_H1, True, NAVY, DISPLAY)
    t.alignment = align("left", "center")
    for r, h in H_HDR.items():
        set_height(ws, r, h)
    if subtitle is not None:
        s.value = subtitle
    if s.value is not None:
        s.font = font(T_BODY, False, MUTED)
        s.alignment = align("left", "top")
    r6, r7 = context_rows
    if context:
        cc = col(context_col or c2)
        n, a_ = ws.cell(r6, cc), ws.cell(r7, cc)
        if context[0] is not None:
            n.value = context[0]
        if len(context) > 1 and context[1] is not None and right_legend is None:
            a_.value = context[1]
        n.font = font(T_BODY, True, NAVY)
        a_.font = font(T_SMALL, False, MUTED)
        n.alignment = align("right", "center")
        a_.alignment = align("right", "top")
    if right_legend is not None:
        lc = ws.cell(r7, col(legend_col or context_col or c2))
        lc.value = right_legend
        lc.font = font(T_LABEL, False, MUTED)
        lc.alignment = align("right", "center")
    if back:
        target = PARENT.get(ws.title, "Start") if back is True else back
        bc = ws.cell(5, col(back_col or c2))
        if bc.column != e.column and (bc.value is None or (isinstance(bc.value, str) and bc.value.startswith("‹"))):
            text_link(bc, f"‹  {target}", target, size=T_LABEL, bold=False,
                      tooltip="Zurück zur Startseite" if target == "Start" else f"Zurück: {sheet_label(target)}")
            bc.alignment = align("right", "bottom")


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

_NO_CAPS = re.compile(r"^\s*(Jahr|Jahre|J\.|Monat|Mon\.)\s*\d|^\s*\d{4}\s*$|^\s*[↑↓›‹]")


def _caps(text, caps):
    """Versalien für statische Textlabels; Jahres-/Monatslabels („Jahr 1“, „2026“) und Links („↑ Übersicht“)
    bleiben unverändert (Wunsch D)."""
    if caps and isinstance(text, str) and not is_formula(text) and not _NO_CAPS.search(text):
        return text.upper()
    return text


def section(ws, row, c1, c2, title=None, level=1, meta=None, extra=None, variant="fill", labels=None, height=None,
            caps=True, keep_height=False):
    """Abschnittskopf – EIN Stil für alle Blätter (P1-12).

    Ebene 1 (Abschnitt): Fläche F7F6F2, linke GOLDKANTE 3 px C9A14A, Unterkante thin D6D2C6,
      Titel 12,5 pt fett 0E2238 in Titelschreibung, Zeilenhöhe 24; meta rechts 8 pt 1E4E8C;
      extra: Zusatz hinter dem Titel als Rich-Text 9 pt 4E5561 (z. B. „· Pflichtangaben“).
    Ebene 2 (Unterabschnitt / Tabellenkopf): 8 pt fett VERSALIEN 1E4E8C, Zeilenhöhe 20,
      variant 'fill' = Fläche F1EEE6 + Unterlinie 3B6A9E (Tabellenkopf; labels wie table_head),
      variant 'line' = ohne Fläche, nur Unterlinie 3B6A9E (Unterabschnitt in ruhigen Blättern).
    Navy-Vollflächen sind der Kopfleiste und den Kachel-Labels vorbehalten.
    Runde 3: Ebene-2-Titel und meta 8,5 pt (P40); keep_height=True lässt die Zeilenhöhe unverändert (Seitenleisten,
    die sich die Zeile mit Datenzeilen teilen – Wunsch E); Jahreslabels bleiben in Normalschreibung (Wunsch D)."""
    h0 = ws.row_dimensions[row].height
    # Runde 4 (P2-10): Das Band trägt nur Titel und Meta; Meta steht IMMER rechtsbündig (8 pt 1E4E8C), nie inline.
    # extra ohne meta wird deshalb zur rechten Meta, sofern die Bandzeile rechts frei ist (keine Spaltenköpfe/Jahre).
    if extra and meta is None and col(c2) > col(c1) and span_px(ws, c1, c2) <= 1700 and all(
            ws.cell(row, cc).value is None and type(ws.cell(row, cc)).__name__ != "MergedCell"
            for cc in range(col(c1) + 1, col(c2) + 1)):
        meta, extra = re.sub(r"^[·\s]+", "", str(extra)), None
    if level == 1:
        band_l1(ws, row, c1, c2, title, right=meta, height=height or H_BAND)
        if keep_height:
            ws.row_dimensions[row].height = h0
        first = ws.cell(row, col(c1))
        if meta is not None:
            ws.cell(row, col(c2)).font = font(T_MICRO, False, BLUE)
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
            c.font = font(T_LABEL, True, BLUE)
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
    first.font = font(T_LABEL, True, BLUE)
    if first.alignment.horizontal not in ("right", "center"):
        first.alignment = align("left", "center", 1)
    if meta is not None:
        last = ws.cell(row, col(c2))
        set_text(last, meta)
        last.font = font(T_MICRO, False, BLUE)
        last.alignment = align("right", "center", 1)
    if keep_height:
        ws.row_dimensions[row].height = h0
    return first


_TILE_FILLS = {NAVY, BLUE, TINT_XL}
# Kachel-Kopffarbe (Runde 4, P1-02): genau zwei benannte Varianten, festgelegt durch die ROLLE des Blatts (DECISIONS 3):
#   strong 0E2238 – Start, Leitfaden, Dashboard, Cockpit, S08, S12 (kräftig / Fintech)
#   calm   1E4E8C – alle Rechenschritte und ruhigen Blätter (S03–S11, AfA-Vergleich, Haushalt, Vermögen …)
TILE_STRONG_SHEETS = ("Start", "Leitfaden", "Dashboard", "Cockpit")
TILE_STRONG_PREFIX = ("S08", "S12")
TILE_HEAD = {"strong": NAVY, "calm": BLUE}
# Runde 5: strong-Kacheln tragen unter dem Nachtblau-Kopf eine 2-px-Goldlinie (edle Kopflinie); calm bleibt ruhig ohne.
TILE_HEAD_LINE = {"strong": side("medium", GOLD), "calm": None}


def tile_variant(ws, variant=None):
    """'strong' | 'calm' für ein Blatt. Die Rolle des Blatts gewinnt; variant ('dark'/'strong'/'light'/'calm') zählt nur
    auf Blättern ohne feste Rolle (Galerie/Tests)."""
    t = getattr(ws, "title", "") or ""
    if t in TILE_STRONG_SHEETS or t.startswith(TILE_STRONG_PREFIX):
        return "strong"
    if re.match(r"^S\d\d ", t) or t in ("AfA-Vergleich", "Haushaltsrechnung", "Vermögensaufstellung", "Bankgespräch",
                                        "Sensitivität", "Projektion", "Steuern", "Finanzierung", "Eingaben"):
        return "calm"
    return "calm" if variant in ("light", "calm") else "strong"


def _tile_registry(ws):
    return ws.__dict__.setdefault("_core_tiles", [])


def _fill_rgb(cell):
    try:
        fg = cell.fill.fgColor.rgb if (cell.fill is not None and cell.fill.fill_type == "solid") else None
        return fg[-6:].upper() if isinstance(fg, str) else None
    except Exception:
        return None


def tile(ws, c1, c2, label_row, value_row, sub_row=None, label=None, value=None, sub=None, kpi=None, fmt=None,
         variant="dark", value_rows=1, value_span=None, sub_right=None, sub_right_fmt=None, status="auto",
         status_col=None, value_ref=None, conditions=None, words=None, neg=None, gap_right=True, gap_top=False,
         caps=True, value_size=T_KPI, gap="auto", value_status=True, label_size=None):
    """KPI-Kachel – EINE Komponente und EINE Anatomie für die ganze Mappe (Runde 3: P10, P11, P40).

    Anatomie (immer dreizeilig, Runde 4 P1-02 FEST): Kopfstreifen 18 pt (H_TILE_LABEL) · Wert 30 pt (H_TILE_VALUE) ·
      Fußzeile 20 pt (H_TILE_SUB). Blätter vergeben keine eigenen Höhen: finalize_components() (läuft über cf_close in
      global_rules.final) setzt die Höhen zurück und legt Überhöhe der Wertzeile in die leere Abstandszeile unter der Kachel.
      Kopfstreifen: Label weiß 8,5 pt fett VERSALIEN (Kurzformen aus KPI_TILE_LABELS; nur im Notfall 8 pt) auf
        'strong' = 0E2238 (Start, Leitfaden, Dashboard, Cockpit, S08, S12) bzw. 'calm' = 1E4E8C (Rechenschritte, ruhige
        Blätter) – die Rolle des Blatts bestimmt die Variante (tile_variant); variant='dark'/'light' wirkt nur auf
        unbekannten Blättern.
      Wert: IMMER 20 pt fett (value_size > 20 wird auf 20 begrenzt) auf FBFAF7, links, Einzug 1.
      sub/sub_right-Formeln: FIXED/TEXT automatisch mit typografischem Minus (minus_text).
      Wertfarbe = Status (P11): mit Status übernimmt der Wert per
        bedingter Formatierung die Statusfarbe (grün/amber/rot), ohne Status 0E2238. value_status=False hält ihn neutral.
        neg (Standard bei €-Formaten und KPI-Regel cf/neg): Beträge < 0 rot (P15).
      Fußzeile: links Kontext/Ziel (8 pt 4E5561), rechts Status „● Wort“ (8 pt fett, Statusfarbe).
    status: 'auto' | 'chip' | 'inline' | 'dot' | 'edge' | None
      auto   – chip bei Kacheln ≥ 2 Spalten, inline bei 1-spaltigen Kacheln mit Fußzeile, sonst edge
      chip   – Status in eigener Zelle rechts (status_col, Standard c2; Kontext über c1…status_col-1)
      inline – 1-spaltige Kachel: Kontext links und „● Wort“ rechtsbündig in DERSELBEN Zelle (bedingtes Zahlenformat
               @* "● prüfen"); eine Zelle = eine Farbe: die Fußzeile steht dann in Statusfarbe (normal, nicht fett)
      dot    – (alt) Fußzeile selbst in Statusfarbe · edge – linke 3-px-Kante in Statusfarbe
    Rinnen (P10): gap='auto' setzt weiße 3-px-Kanten rechts (gap_right), links, wenn links eine Kachel angrenzt, und oben,
      wenn darüber eine Kachel endet (oder gap_top) – horizontal = vertikal. Mehr als 3 px kann Excel mit Rahmen nicht
      zeichnen; für ≥ 8 px eine schmale Rinnenspalte/-zeile (≈ 1,7 Zeichen / 6 pt) wie auf dem Dashboard verwenden.
    value_span: letzte Spalte des Wertfelds; dahinter sub_right (8 pt 4E5561, rechtsbündig im Wertfeld).
    value=None lässt eine vorhandene Formel stehen. Nicht-Anker-Zellen der Verbünde müssen leer sein."""
    spec = KPI.get(kpi, {}) if kpi else {}
    c1i, c2i = col(c1), col(c2)
    last_v = value_row + value_rows - 1
    head_bg = TILE_HEAD[tile_variant(ws, variant)]
    value_size = min(value_size or T_KPI, T_KPI)     # Runde 4 (P1-02): Kachelwert immer 20 pt, nie größer
    if isinstance(sub, str):
        sub = minus_text(sub)
    if isinstance(sub_right, str):
        sub_right = minus_text(sub_right)
    rows = [label_row] + list(range(value_row, last_v + 1)) + ([sub_row] if sub_row else [])
    white3 = side("thick", WHITE)
    head_line = TILE_HEAD_LINE.get(tile_variant(ws, variant))    # Runde 5: feine Goldlinie unter dem Kopf
    auto = gap == "auto" or gap is True
    left_gap = auto and c1i > 1 and _fill_rgb(ws.cell(label_row, c1i - 1)) in _TILE_FILLS
    top_gap = gap_top or (auto and label_row > 1 and _fill_rgb(ws.cell(label_row - 1, c1i)) in _TILE_FILLS)
    for r in rows:
        for c in iter_cells(ws, c1i, r, c2i, r):
            c.fill = fill(head_bg if r == label_row else TINT_XL)
            c.border = Border(right=white3 if (gap_right and c.column == c2i) else None,
                              left=white3 if (left_gap and c.column == c1i) else None,
                              top=white3 if (top_gap and r == label_row) else None,
                              bottom=head_line if (r == label_row and head_line is not None) else None)
    # Kopfstreifen
    safe_merge(ws, c1i, label_row, c2i, label_row)
    lab = ws.cell(label_row, c1i)
    if label is not None:
        if is_formula(label):
            lab.value = label
        else:
            short = KPI_TILE_LABELS.get(kpi) if (kpi and label.strip().upper() in _TILE_SHORT) else None
            set_text(lab, _caps(short or _TILE_SHORT.get(label.strip().upper(), label), caps))
    elif spec.get("label") and not is_formula(lab.value):
        set_text(lab, _caps(KPI_TILE_LABELS.get(kpi) or KPI_LABELS.get(kpi, spec["label"]), caps))
    elif isinstance(lab.value, str) and not is_formula(lab.value):
        v0 = _TILE_SHORT.get(lab.value.strip().upper(), lab.value)
        lab.value = v0.upper() if caps else v0
    # Kopfschrift 8,5 pt (P1-02). Nur wenn ein Label trotz Kurzform nicht passt, 8 pt als Notlösung.
    lsz = label_size or T_LABEL
    if label_size is None:
        t = display_text(lab)
        if t and not fits(t, span_px(ws, c1i, c2i), T_LABEL, True, 1, 0.97):
            lsz = T_MICRO
    lab.font = font(lsz, True, WHITE)
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
        mode = ("chip" if (status_col or c2i > c1i) else "inline") if sub_row else "edge"
    if mode in ("chip", "inline") and not sub_row:
        mode = "edge"
    if mode == "chip" and not (status_col or c2i > c1i):
        mode = "inline"
    # Fußzeile
    if sub_row:
        chip_c = (col(status_col) if status_col else c2i) if mode == "chip" else None
        end = chip_c - 1 if (chip_c and chip_c > c1i) else c2i
        safe_merge(ws, c1i, sub_row, end, sub_row)
        s_ = ws.cell(sub_row, c1i)
        if sub is not None:
            s_.value = sub
        s_.font = font(T_MICRO, False, MUTED)
        s_.alignment = align("left", "center", 1)
        set_height(ws, sub_row, H_TILE_SUB)
        if chip_c and chip_c > c1i:
            if chip_c < c2i:
                safe_merge(ws, chip_c, sub_row, c2i, sub_row)
            status_pill(ws, ws.cell(sub_row, chip_c), conditions=conds, style="chip", words=words)
        if mode == "inline":
            if s_.value is None:
                s_.value = '=""'
            wd = {**STATUS_WORDS, **(words or {})}
            nf = {lvl: f'@* "{STATUS_DOT}  {wd[lvl]}   "' for _, lvl in conds}   # 3 Leerzeichen ≈ Einzug 1 rechts
            status_cf(ws, s_.coordinate, conds, font_color=True, bold=False, numfmt=nf)
        if mode == "dot":
            status_cf(ws, s_.coordinate, conds, font_color=True, bold=True)
    if mode == "edge":
        status_edge(ws, rng(c1i, value_row, c1i, sub_row or last_v), conditions=conds)
    # Wertfarbe = Status (P11)
    if conds and value_status and mode is not None:
        status_cf(ws, rng(c1i, value_row, vs_end, last_v), conds, font_color=True, bold=True)
    # Negative Beträge (P15)
    if neg is None:
        neg = spec.get("rule") in ("cf", "neg") or ("€" in (f or "") and "%" not in (f or ""))
    if neg:
        neg_red(ws, rng(c1i, value_row, vs_end, last_v))
    reg = _tile_registry(ws)
    key = (c1i, label_row)
    reg[:] = [x for x in reg if (x["c1"], x["label_row"]) != key]
    reg.append(dict(c1=c1i, c2=c2i, label_row=label_row, value_row=value_row, value_rows=value_rows, sub_row=sub_row,
                    head=head_bg))
    return val


def note(ws, row, c1, c2, text=None, label="Hinweis", link_text=None, target_sheet=None, target_cell=None,
         link_col=None, row2=None, tooltip=None):
    """Hinweis-Callout – EINE Bauform (P34) für Leitfaden Z. 8, S01 C21, Start …:
    Fläche FFF9EA, linke Kante 3 px E6CB77 (Eingabe-Gelb = „betrifft Ihre Eingaben“), Label fett 0E2238 + Text 3A3F45
    (9 pt, Rich-Text bei statischem Text; der Text trägt KEINEN Link), Aktionslink „… ›“ in EIGENER Zelle rechts
    (link_col … c2, 9 pt fett 1E4E8C, rechtsbündig). row2: letzte Zeile bei mehrzeiligem Hinweis (Text oben)."""
    c1i, c2i = col(c1), col(c2)
    r2 = row2 or row
    lc = col(link_col) if (link_col and link_text) else None
    t_end = lc - 1 if lc else c2i
    for r in range(row, r2 + 1):
        for c in iter_cells(ws, c1i, r, c2i, r):
            c.fill = fill(NOTE_BG)
            c.border = Border(left=side("thick", NOTE_LINE) if c.column == c1i else None)
            if c.hyperlink is not None and (lc is None or c.column < lc):
                c.hyperlink = None
    safe_merge(ws, c1i, row, t_end, r2)
    tc = ws.cell(row, c1i)
    body = text if text is not None else tc.value
    if isinstance(body, str) and not is_formula(body) and label:
        tc.value = rich([(label + "   ", T_SMALL, True, NAVY), (body, T_SMALL, False, INK2)])
    elif body is not None:
        tc.value = body
    tc.font = font(T_SMALL, False, INK2)
    tc.alignment = align("left", "top" if r2 > row else "center", 1, wrap=r2 > row)
    if lc:
        if lc < c2i:
            safe_merge(ws, lc, row, c2i, row)
        lk = ws.cell(row, lc)
        set_text(lk, link_text)
        lk.font = font(T_SMALL, True, BLUE)
        lk.alignment = align("right", "top" if r2 > row else "center", 1)
        if target_sheet:
            lk.hyperlink = Hyperlink(ref=lk.coordinate, location=link_loc(target_sheet, target_cell),
                                     display=link_text, tooltip=tooltip)
    return tc


def callout_box(ws, c1, head_row, c2, body_r1, body_r2, title="Einordnung", text=None, kpi=None, value_ref=None,
                conditions=None, pill=None, pill_col=None, words=None, head=True, fit="auto", height_text=None,
                variant="info", link_text=None, target_sheet=None, target_cell=None, link_col=None, suffix=None,
                pad_right=True, rest_row=None):
    """Einordnungs-Box – EINE Komponente für S01–S12, Sensitivität, Bank … (P1-04).

    Runde 4 (P2-01): Körperhöhe = Textzeilen × 12,5 + 8 pt (callout_height, wie text_row_height) – nie an das Kachelraster
      daneben gebunden. rest_row: Zeile UNTER der Box (leere Abstandszeile), die eine zu große Resthöhe des letzten
      Körperzeile aufnimmt, statt die Box aufzublähen (fit='auto'/'exact').
      pad_right=True: Innenabstand rechts ≥ links – statischer Text wird mit festen Zeilenumbrüchen auf
      Boxbreite − 2 × Einzug umbrochen (Formeltexte bleiben unverändert; dort wirkt nur der linke Einzug).

    Kopf (20 pt): Fläche FBFAF7 über c1…c2, Titel 10 pt fett 0E2238 links (Einzug 1), Unterlinie thin CBD5E1,
      rechts Status-Pill (9 pt fett, bei Status weiß mit 1-px-Rahmen in Statusfarbe) in pill_col (Standard c2).
    Körper: immer FBFAF7, Text 9 pt 3A3F45, OBEN ausgerichtet, Einzug 1, umbrechend; Verbund c1…c2 × body_r1…body_r2.
    Status: nur der linke Balken (3 px) wechselt von 3B6A9E in die Statusfarbe – keine Flächentönung.
      kpi/value_ref (KPI-Regel) oder conditions [(Bedingung, 'red'|'amber'|'green')]; pill: eigene Formel für
      das Pill-Wort (Standard „● erfüllt/prüfen/kritisch“), pill=False unterdrückt die Pill.
    fit: Körperhöhe nach Textlänge (Calibri-Metrik; höchstens ≈ 12 px Luft unten).
      'auto' = einzeiliger Körper exakt, mehrzeiliger Körper: nur die letzte Zeile wachsen lassen · 'exact' ·
      None = Höhen nicht ändern. height_text: Text für die Bemessung (Standard: längster Formelzweig).
    Liefert die Zeilenzahl des Körpertexts. Abstand zum Folgeblock: genau eine Leerzeile H_GAP (12 pt).
    Runde 3: die Pill ist die EINE Pill (P21): zarter Status-Tint, ohne Rahmen, 9 pt fett, zentriert – sie liegt in der
      Box-Fläche. suffix: Wert hinter dem Statuswort („● prüfen · 4,0 %“).
    variant='note' (P34): Hinweis-Bauform – siehe note(); title = Label, text = Hinweis, Link in link_col … c2."""
    if variant == "note":
        note(ws, head_row if head else body_r1, c1, c2, text, title or "Hinweis", link_text, target_sheet, target_cell,
             link_col, row2=body_r2)
        return 1
    c1i, c2i = col(c1), col(c2)
    first = head_row if head else body_r1
    bar = side("thick", ACCENT)
    for r in range(first, body_r2 + 1):
        for c in iter_cells(ws, c1i, r, c2i, r):
            c.fill = fill(TINT_XL)
            c.border = Border(left=bar if c.column == c1i else None,
                              bottom=side("thin", LINE2) if (head and r == head_row) else None)
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
                        formula=pill if isinstance(pill, str) else None, style="pill", words=words, suffix=suffix)
    safe_merge(ws, c1i, body_r1, c2i, body_r2)
    b = ws.cell(body_r1, c1i)
    if text is not None:
        b.value = text
    b.font = font(T_SMALL, False, INK2)
    b.alignment = align("left", "top", 1, wrap=True)
    width = span_px(ws, c1i, c2i)
    if pad_right and isinstance(b.value, str) and not is_formula(b.value):
        wrapped = hard_wrap(b.value, width, T_SMALL, indent=1)
        # Bei fit=None bemisst der Aufrufer die Höhe: nur umbrechen, wenn keine zusätzliche Zeile entsteht.
        if fit or lines_needed_metric(wrapped, width, T_SMALL, False, 1) <= lines_needed_metric(b.value, width, T_SMALL,
                                                                                                 False, 1):
            b.value = wrapped
    if conds:
        status_edge(ws, rng(c1i, first, c1i, body_r2), conditions=conds)
    n = 1
    if fit:
        t = height_text if height_text is not None else display_text(b)
        n = lines_needed_metric(t, width, T_SMALL, False, 1) if t else 1
        need = callout_height(n)
        others = sum((ws.row_dimensions[r].height or 15) for r in range(body_r1, body_r2))
        last_h = ws.row_dimensions[body_r2].height or 15
        if body_r1 == body_r2 or fit == "exact":
            new_h = max(need - others, H_ROW if body_r1 == body_r2 else 6)
            if rest_row and new_h < last_h:
                set_height(ws, rest_row, (ws.row_dimensions[rest_row].height or 15) + last_h - new_h)
            set_height(ws, body_r2, new_h)
        else:
            cur = others + last_h
            if need > cur:
                set_height(ws, body_r2, last_h + need - cur)
            elif rest_row and cur - need > 6:
                cut = min(cur - need, last_h - 6)
                if cut > 0:
                    set_height(ws, body_r2, last_h - cut)
                    set_height(ws, rest_row, (ws.row_dimensions[rest_row].height or 15) + cut)
    return n


def callout_height(lines):
    """Körperhöhe einer Einordnungs-Box / eines Seitenpanels mit `lines` Textzeilen (9 pt): n × 12,5 + 8 pt,
    auf ganze Pixel (21 · 33 · 45,75 · 58,5 …) – höchstens ≈ 8 px Luft unter dem Text (P2-01)."""
    return text_row_height(lines)


def hard_wrap(text, width_px, size=T_SMALL, bold=False, indent=1, right_pad=None):
    """Statischen Text mit festen Umbrüchen („\n“) auf die Breite width_px − Einzug links − Innenabstand rechts
    umbrechen (P2-01: Innenabstand rechts ≥ links). right_pad Standard = 9 px je Einzugsstufe. Vorhandene Absätze
    bleiben; Excel bricht danach nicht mehr selbst um. Idempotent (bereits umbrochene Zeilen passen)."""
    if not isinstance(text, str) or not text or is_formula(text):
        return text
    pad = 9 * max(indent, 1) if right_pad is None else right_pad
    avail = max(cell_inner_px(width_px, indent) * 0.98 - pad, 20)
    out = []
    for para in text.split("\n"):
        tokens = re.findall(r"[^ \-/]*[\-/]|[^ \-/]+ ?| ", para)
        cur = ""
        lines = []
        for tk in tokens:
            cand = cur + tk
            if text_width(cand.rstrip(), size, bold) <= avail or not cur.strip():
                cur = cand
            else:
                lines.append(cur.rstrip())
                cur = tk.lstrip() if tk.strip() else ""
        lines.append(cur.rstrip())
        out.append("\n".join(lines))
    return "\n".join(out)


BTN = {  # kind: (Fläche, Rahmen, Schrift, fett, Höhe) – Runde 4 (P1-03): DREI Button-Typen, alle 25,5 pt, 10 pt fett
    # Runde 5 „Midnight & Gold“
    "primary": (NAVY, GOLD, WHITE, True, H_BTN),         # Hauptaktion „Weiter …  ›“ – Nachtblau, GOLDRAHMEN, Schrift weiß
    "secondary": (WHITE, BLUE, BLUE, True, H_BTN),       # „‹  Zurück …“ – immer links, weiß mit Rahmen 1E4E8C
    "tertiary": (TINT_XL, LINE2, BLUE, True, H_BTN),     # Mitte („Übersicht: Leitfaden“, „Dashboard“) – FBFAF7, Rahmen D6D2C6
    "chip": (ICE, MIST, BLUE, True, H_BTN),              # Link-/Kontext-Chip („Kauf als: Privat · ändern  ›“) – E8EDF4, nie Gelb
    "ghost": (NAVY_2, "3A5578", WHITE, True, H_BTN),     # Sekundär auf Navy (Start-Hero) – 1B3553, Rahmen 3A5578
}
# Alt-Arten (Runde 2/3) → Runde-4-Typen. Gelb FFF5D6 gehört nur echten Eingabezellen, deshalb wird 'input' zum Link-Chip.
BTN_ALIAS = {"soft": "tertiary", "back": "secondary", "input": "chip", "outline": "secondary"}
BTN_TEXT = {  # reine Textlink-Stufe ohne Rahmen/Fläche (nur noch explizit; nicht in Buttonzeilen verwenden)
    "text": (BLUE, True, H_BTN),
}
BTN_TIERS = ("primary", "secondary", "tertiary")        # Dreistufigkeit einer Buttonzeile (P13 / P1-03)


def btn_kind(kind):
    """Runde-4-Typ einer (auch alten) Button-Art: soft → tertiary, back → secondary, input → chip."""
    return BTN_ALIAS.get(kind, kind)


def btn(ws, c1, row, c2, text, target_sheet=None, kind="primary", target_cell=None, tooltip=None, height=None,
        size=T_BODY, set_row=True):
    """Button-System (Runde 4, P1-03) – drei Typen, alle 25,5 pt hoch, 10 pt fett, zentriert, Rahmen 1 px:
      primary   0E2238 gefüllt, Goldrahmen C9A14A, Schrift weiß     – „Weiter …  ›“, immer rechts (bündig an der Hauptspalte)
      secondary weiß, Rahmen 1E4E8C, Schrift 1E4E8C – „‹  Zurück …“, immer links, gleich breit wie Weiter
      tertiary  FBFAF7, Rahmen D6D2C6, Schrift 1E4E8C – Mitte („Übersicht: Leitfaden“)
    Dazu 'chip' (E8EDF4, Rahmen CBD5E1 – Link-/Kontext-Chip statt gelbem Eingabe-Look) und 'ghost' (auf Navy).
    Alte Arten werden abgebildet: soft → tertiary, back → secondary, input → chip. 'link' = Textlink (Alt-API).
    Jeder Button trägt einen ScreenTip (tooltip oder auto_tooltip, P3-17)."""
    if kind == "link":
        button(ws, c1, row, c2, text, target_sheet, "link", target_cell, tooltip)
        return ws.cell(row, col(c1))
    kind = btn_kind(kind)
    bg, ln_c, fg, bold, h = BTN[kind] if kind in BTN else (None, None) + BTN_TEXT.get(kind, BTN_TEXT["text"])
    ln = side("thin", ln_c) if ln_c else None
    safe_merge(ws, c1, row, c2, row)
    c1i, c2i = col(c1), col(c2)
    for cc in iter_cells(ws, c1i, row, c2i, row):
        cc.fill = fill(bg) if bg else NOFILL
        cc.border = Border(top=ln, bottom=ln, left=ln if cc.column == c1i else None,
                           right=ln if cc.column == c2i else None) if ln else Border()
    c = ws.cell(row, c1i)
    set_text(c, text)
    c.font = font(size, bold, fg)
    c.alignment = align("center", "center")
    if target_sheet:
        _hyperlink(c, target_sheet, target_cell, text, tooltip)
    if set_row:
        set_height(ws, row, height or h)
    return c


def _is_back(text):
    return isinstance(text, str) and text.lstrip().startswith("‹")


def _clear_btn_cell(cell):
    if not is_formula(cell.value):
        cell.value = None
    cell.hyperlink = None
    cell.fill = NOFILL
    cell.border = Border()


def btn_row(ws, row, items, height=None, tiers=True, trim=True):
    """Button-Reihe (Runde 4, P1-03): EINE Höhe 25,5 pt, Reihenfolge Zurück (links) · Mitte · Weiter (rechts).
    items: [dict(c1=…, c2=…, text=…, target=…, kind=…, tooltip=…, cell=…)]. Liefert [(text, breite_px, passt)].
    tiers=True: Enthält die Reihe einen Primär-Button, bleibt nur „‹ Zurück …“ sekundär; jeder weitere Sekundär-Button
      wird tertiär (FBFAF7-Fläche, Rahmen D6D2C6).
    trim=True: Stößt ein tertiärer Button/Chip ohne Zwischenspalte an einen Nachbarn, gibt er seine Randspalten ab
      (sofern der Text dann noch mit Reserve passt) – so bleibt zwischen allen Buttons eine weiße Fuge.
    Gleiche Breiten entstehen über gleiche Spaltenraster (Zurück C:D ≈ Weiter H:I auf den Schritt-Seiten)."""
    out = []
    has_primary = any(btn_kind(it.get("kind", "secondary")) == "primary" for it in items)
    specs = []
    for it in items:
        kind = btn_kind(it.get("kind", "secondary"))
        if tiers and has_primary and kind == "secondary" and not _is_back(it.get("text")):
            kind = "tertiary"
        specs.append([col(it["c1"]), col(it["c2"]), kind, it])
    if trim and len(specs) > 1:
        for i, sp in enumerate(specs):
            if sp[2] not in ("tertiary", "chip") or sp[1] <= sp[0]:   # Zurück/Weiter behalten ihre (gleiche) Breite
                continue
            others = [o for j, o in enumerate(specs) if j != i]
            need = text_width(sp[3]["text"], T_BODY, True) + 24
            while sp[1] > sp[0] and any(o[1] == sp[0] - 1 for o in others) and \
                    span_px(ws, sp[0] + 1, sp[1]) >= need:
                for cc in iter_cells(ws, sp[0], row, sp[0], row):
                    for mr in list(ws.merged_cells.ranges):
                        if mr.min_row <= row <= mr.max_row and mr.min_col <= sp[0] <= mr.max_col:
                            ws.unmerge_cells(str(mr))
                    _clear_btn_cell(cc)
                sp[0] += 1
            while sp[1] > sp[0] and any(o[0] == sp[1] + 1 for o in others) and \
                    span_px(ws, sp[0], sp[1] - 1) >= need:
                for mr in list(ws.merged_cells.ranges):
                    if mr.min_row <= row <= mr.max_row and mr.min_col <= sp[1] <= mr.max_col:
                        ws.unmerge_cells(str(mr))
                _clear_btn_cell(ws.cell(row, sp[1]))
                sp[1] -= 1
    for c1i, c2i, kind, it in specs:
        btn(ws, c1i, row, c2i, it["text"], it.get("target"), kind, it.get("cell"), it.get("tooltip"), set_row=False)
        w = span_px(ws, c1i, c2i)
        out.append((it["text"], w, fits(it["text"], w, T_BODY, True, 0, 0.9)))
    set_height(ws, row, height or H_BTN)
    return out


def sum_row(ws, row, c1, c2, stage="sub", value_from=None, neg=None, keep_size=True, top=True):
    """Summen- und Kennzahlenhierarchie (Runde 3 P12, Runde 4 P2-06/P1-10) – pro Block GENAU EINE Endsumme.
    Die Zeilenbeschriftung trägt das Rechenzeichen („– Zinsen“, „= Cashflow …“).
      'final'  (Alias 'result', 2): die EINE Endsumme: fett 0E2238, Fläche F7F6F2, Oberlinie thin 1E4E8C,
                    Doppellinie unten 0E2238, Mindesthöhe 18 pt
      'sub'    (1): Zwischensumme: fett 0E2238, OHNE Fläche, feine Oberlinie CFC9BA (top=False: ohne Linie, z. B. für
                    eine zweite Zwischensumme direkt darunter)
      'kpi'       : Kennzahlzeile: Beschriftung regulär 1A1D21, Wert(e) 0E2238 NICHT fett, keine Fläche/Zusatzlinie
      'kpi_band'  : Kennzahlenpaar im Band (Steuern IRR/Multiplikator): Fläche F7F6F2, Beschriftung regulär 0E2238,
                    Wert fett 1E4E8C, ohne Linien (die Doppellinie trägt nur die letzte Zeile des Paars: danach
                    ws.cell(...).border mit bottom=double oder die Endsumme 'final')
      'plain'     : Nebenzeile: regulär 10 pt 1A1D21, ohne Fläche/Zusatzlinien (Label ohne „=“)
      'deduct' (0): Abzugs-/Teilzeile „–“: nur fett, keine Fläche, keine Zusatzlinie
      'memo'   (3): nachrichtlich: 9 pt kursiv 4E5561, Beschriftung Einzug 2, OHNE Fläche und Rahmenlinien
    neg (Standard: an bei 'sub'/'final' – Negativ-Rot-Regel P15): negative Beträge ab value_from (Standard: zweite
    Spalte) rot B42318; neg=False schaltet ab (z. B. Steuerwirkung im Format tax_effect)."""
    stage = {0: "deduct", 1: "sub", 2: "final", 3: "memo", "result": "final", "total": "final"}.get(stage, stage)
    if stage == "memo":
        memo(ws, row, c1, c2, clear_lines=True)
        return
    vf0 = col(value_from) if value_from else col(c1) + 1
    for c in iter_cells(ws, c1, row, c2, row):
        sz = (c.font.sz if keep_size else None) or T_BODY
        is_val = c.column >= vf0
        if stage == "deduct":
            c.font = font(sz, True, c.font.color.rgb[-6:] if (c.font.color is not None and isinstance(c.font.color.rgb, str)) else INK)
            c.fill = NOFILL
        elif stage == "sub":
            c.fill = NOFILL
            c.border = Border(top=side("thin", LINE_SUB) if top else None, bottom=side("hair", LINE))
            c.font = font(sz, True, NAVY)
        elif stage == "kpi":
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE))
            c.font = font(sz, False, NAVY if is_val else INK)
        elif stage == "kpi_band":
            c.fill = fill(TINT)
            c.border = Border()
            c.font = font(sz, is_val, BLUE if is_val else NAVY)
        elif stage == "plain":
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE))
            c.font = font(T_BODY, False, INK)
        else:
            c.fill = fill(TINT)
            c.border = Border(top=side("thin", BLUE), bottom=side("double", NAVY))
            c.font = font(sz, True, NAVY)
    if stage == "final":
        set_height(ws, row, max(ws.row_dimensions[row].height or 0, H_ROW))
    if stage in ("kpi", "kpi_band", "plain") and neg is None:
        neg = False
    auto = neg is None
    if auto:
        neg = stage in ("sub", "final")
    if neg:
        vf = col(value_from) if value_from else col(c1) + 1
        if not auto:
            if vf <= col(c2):
                neg_red(ws, rng(vf, row, c2, row))
            return
        # Automatik: Zellen mit Wortformat (tax_effect „Zahlung/Erstattung“, Text) bleiben neutral
        run = None
        for cc in range(vf, col(c2) + 2):
            ok = cc <= col(c2) and not _neutral_fmt(ws.cell(row, cc).number_format)
            if ok and run is None:
                run = cc
            elif not ok and run is not None:
                neg_red(ws, rng(run, row, cc - 1, row))
                run = None


def _neutral_fmt(fmt):
    """Formate, deren Vorzeichen in Worten steckt (tax_effect) oder Text – nie rot."""
    return isinstance(fmt, str) and ("Erstattung" in fmt or "Zahlung" in fmt or fmt == "@")


# =============================================================================== Runde 4: Abschluss je Blatt
def _row_has_other_content(ws, row, spans, wrap_only=False):
    """True, wenn Zeile `row` außerhalb der Spaltenbereiche `spans` [(c1, c2)] Inhalt hat (bei wrap_only nur
    umbrechenden Inhalt oder Schrift > 20 pt) – dann darf die Zeilenhöhe nicht frei gesetzt werden."""
    for c in ws[row] if row <= ws.max_row else ():
        if c.value is None or any(a <= c.column <= b for a, b in spans):
            continue
        if not wrap_only:
            return True
        if (c.alignment is not None and c.alignment.wrap_text) or (c.font is not None and (c.font.sz or 0) > T_KPI):
            return True
    return False


def _multirow_text_merge(ws, row, spans):
    """True, wenn Zeile `row` in einem mehrzeiligen, umbrechenden Verbund außerhalb der Kacheln liegt."""
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.max_row > mr.min_row and \
                not any(a <= mr.min_col and mr.max_col <= b for a, b in spans):
            a = ws.cell(mr.min_row, mr.min_col)
            if a.value is not None and a.alignment is not None and a.alignment.wrap_text:
                return True
    return False


def _row_blank(ws, row):
    return all(c.value is None for c in ws[row]) if row <= ws.max_row else True


def enforce_tile_heights(ws):
    """P1-02: feste Kachelhöhen 18 / 30 / 20 pt für alle mit tile() gebauten Kacheln (Wertzeile einzeilig).
    Überhöhe der Wertzeile wandert in die leere Abstandszeile direkt unter der Fußzeile (bzw. über dem Kopf), damit
    daneben liegende Diagramme gleich hoch bleiben. Zeilen, die andere umbrechende Inhalte tragen, bleiben unberührt."""
    reg = [t for t in ws.__dict__.get("_core_tiles", [])
           if _fill_rgb(ws.cell(t["label_row"], t["c1"])) == t["head"]]
    if not reg:
        return
    spans_by_row = {}
    for t in reg:
        rows = [t["label_row"]] + list(range(t["value_row"], t["value_row"] + t["value_rows"])) + \
            ([t["sub_row"]] if t["sub_row"] else [])
        for r in rows:
            spans_by_row.setdefault(r, []).append((t["c1"], t["c2"]))
    done = set()
    for t in reg:
        plan = [(t["label_row"], H_TILE_LABEL)]
        if t["value_rows"] == 1:
            plan.append((t["value_row"], H_TILE_VALUE))
        if t["sub_row"]:
            plan.append((t["sub_row"], H_TILE_SUB))
        for r, h in plan:
            if r in done:
                continue
            cur = ws.row_dimensions[r].height or 15
            if abs(cur - h) < 0.01:
                done.add(r)
                continue
            spans = spans_by_row.get(r, [])
            if cur > h and (_row_has_other_content(ws, r, spans, wrap_only=True) or
                            (cur - h > 2.01 and _multirow_text_merge(ws, r, spans))):
                continue
            if cur < h and _multirow_text_merge(ws, r, spans) and h - cur > 2.01:
                continue
            excess = cur - h
            if r == t["value_row"] and excess > 2.01:
                spacer = None
                last = t["sub_row"] or t["value_row"]
                for cand in (last + 1, t["label_row"] - 1):
                    if cand >= 1 and cand not in spans_by_row and _row_blank(ws, cand) and \
                            not _multirow_text_merge(ws, cand, []):
                        spacer = cand
                        break
                if spacer is None:
                    continue
                set_height(ws, spacer, (ws.row_dimensions[spacer].height or 15) + excess)
            set_height(ws, r, h)
            done.add(r)


def ensure_tooltips(ws):
    """P3-17: Jeder Zell-Hyperlink bekommt einen ScreenTip (auto_tooltip aus Text und Ziel), falls keiner gesetzt ist."""
    for row in ws.iter_rows():
        for c in row:
            hl = c.hyperlink
            if hl is None or getattr(hl, "tooltip", None):
                continue
            loc = hl.location or ""
            m = re.match(r"^'?(.*?)'?!\$?([A-Z]{1,3})\$?(\d+)", loc)
            if not m:
                continue
            sheet = m.group(1).replace("''", "'")
            txt = c.value if isinstance(c.value, str) and not is_formula(c.value) else (hl.display or "")
            hl.tooltip = auto_tooltip(txt, sheet, f"{m.group(2)}{m.group(3)}", ws.title)[:255]


def finalize_components(ws):
    """Abschluss je Blatt (Runde 4): feste Kachelhöhen (P1-02) und ScreenTips für alle Zell-Links (P3-17).
    Idempotent; läuft automatisch in cf_close() (Blatt-Module und global_rules.final)."""
    enforce_tile_heights(ws)
    ensure_tooltips(ws)


def patch_tooltips(xlsx_path):
    """P3-17 nach der Neuberechnung: LibreOffice verwirft die ScreenTips der Zell-Hyperlinks. Diese Funktion schreibt
    in eine FERTIGE .xlsx (Zip, Blatt-XML) jedem <hyperlink> ohne tooltip einen ScreenTip aus auto_tooltip(display,
    Zielblatt, Zielzelle). Ändert nur das Attribut tooltip; idempotent. Aufruf z. B. am Ende von build_pro
    (nach navigation.py) oder in finish_sheets.fix. Liefert die Zahl der ergänzten ScreenTips."""
    import html
    import os
    import shutil
    import tempfile
    import zipfile
    with zipfile.ZipFile(xlsx_path) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
        infos = {n: z.getinfo(n) for n in names}
    wbx = data.get("xl/workbook.xml", b"").decode("utf-8")
    rels = data.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
    rid_target = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        i, t = re.search(r'Id="([^"]+)"', m.group(0)), re.search(r'Target="([^"]+)"', m.group(0))
        if i and t:
            rid_target[i.group(1)] = t.group(1)
    sheet_of = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wbx):
        tag = m.group(0)
        nm = re.search(r'name="([^"]*)"', tag)
        rid = re.search(r'r:id="([^"]*)"', tag)
        if nm and rid and rid.group(1) in rid_target:
            tgt = rid_target[rid.group(1)].lstrip("/")
            tgt = tgt if tgt.startswith("xl/") else "xl/" + tgt
            sheet_of[tgt] = html.unescape(nm.group(1))
    count = 0

    def fix(m, title):
        nonlocal count
        tag = m.group(0)
        if " tooltip=" in tag:
            return tag
        loc = re.search(r'location="([^"]*)"', tag)
        if not loc:
            return tag
        lm = re.match(r"^'?(.*?)'?!\$?([A-Z]{1,3})\$?(\d+)", html.unescape(loc.group(1)))
        if not lm:
            return tag
        disp = re.search(r'display="([^"]*)"', tag)
        tip = auto_tooltip(html.unescape(disp.group(1)) if disp else "", lm.group(1).replace("''", "'"),
                           f"{lm.group(2)}{lm.group(3)}", title)[:255]
        count += 1
        end = "/>" if tag.endswith("/>") else ">"
        return tag[: -len(end)].rstrip() + f' tooltip="{html.escape(tip, quote=True)}"' + end

    for n, title in sheet_of.items():
        if n in data:
            x = data[n].decode("utf-8")
            x2 = re.sub(r"<hyperlink\b[^>]*?/?>", lambda m: fix(m, title), x)
            if x2 != x:
                data[n] = x2.encode("utf-8")
    if not count:
        return 0
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(os.path.abspath(xlsx_path)))
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(infos[n], data[n])
    shutil.move(tmp, xlsx_path)
    return count
