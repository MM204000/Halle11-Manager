"""Formularblätter Eingaben, Konfiguration, Hinweise (Agent E) – ruhiger „Private Banking“-Stil, Runde 3.

Komponentensprache ausschließlich aus core.py:
  Seitenkopf core.page_header (P04: Eyebrow links / Unterreiter rechts · Titel + Objekt · Untertitel + „Erstellt für“) ·
  Abschnittskopf core.section (Ebene 1 E7EEF7 + Akzentkante, Ebene 2 EEF3FA-Tabellenkopf) ·
  Summenhierarchie core.sum_row (P12: 'sub' · genau ein 'final' je Abschnitt · 'memo' für nachrichtliche Kennzahlen) ·
  Zahlenformate core.NUMFMT/numfmt (P24/P38: Einheiten Jahre/Monate/m² im Format, „–“ für berechnete Nullen) ·
  Zeilenhöhen core.fit_row (Hinweise: metric=True, n × 12,5 + 8 pt, P27) · Rich-Text core.rich · Fuß core.footer.

Inhaltsbreite aller drei Blätter: Kopfleiste endet bei 1 368 px (core.NAV_MIN_PX), P3-07/Befund „Inhaltsbreite“:
  Eingaben    A 31 | B 300 | C:D 350 | E 88 | F 56 | G:L 543  → 1 368 px (Randspalte M liegt hinter der Kopfleiste)
  Konfiguration A 31 | B 301 | C 76 | D 70 | E 308 | F 40 (Rinne) | G 542 → 1 368 px
  Hinweise    A 31 | B 210 | C 550 | D 382 | E 195 → 1 368 px (Breiten nach minimaler Zeilenzahl ohne Grenzfälle)

Nur Darstellung: Formeln und Eingabewerte bleiben unverändert; verändert werden nur Stile, Zahlenformate,
statische Beschriftungen (auf die nichts verweist), Verbünde, Spaltenbreiten und Zeilenhöhen.
"""
import re
from copy import copy

from openpyxl.styles import Border
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C
from core import (ACCENT, AMBER, BLUE, INK, INK2, INPUT_LINE, LINE, LINE2, MUTED, NAVY, NOFILL, NUMFMT,
                  SANS, T_BODY, T_H3, T_MICRO, T_SMALL, align, font, side)

NBSP = " "
_CHOICES = {}


# ================================================================================================ Hilfen
def _ungroup_cols(ws):
    """Gruppierte Spaltenbreiten (<col min max>) in Einzelspalten auflösen (Sicherheitsnetz zu global_rules)."""
    dims = ws.column_dimensions
    for key, d in list(dims.items()):
        lo, hi = d.min, d.max
        if lo and hi and hi > lo:
            for ci in range(lo + 1, hi + 1):
                letter = get_column_letter(ci)
                nd = ColumnDimension(ws, index=letter, width=d.width, hidden=d.hidden, customWidth=d.customWidth,
                                     outlineLevel=d.outlineLevel, collapsed=d.collapsed, min=ci, max=ci)
                nd._style = copy(d._style)
                dims[letter] = nd
            d.max = lo


def _width(ws, letter, w):
    d = ws.column_dimensions[letter]
    ci = C.col(letter)
    d.min, d.max = ci, ci
    d.width = w


def _width_px(ws, letter, px):
    """Spaltenbreite so setzen, dass core.col_px genau `px` Pixel ergibt (Excel-Raster, Ziffernbreite 7 px)."""
    w = round((px + 0.3) / 7, 2)
    for _ in range(40):
        _width(ws, letter, w)
        got = C.col_px(ws, letter)
        if got == px:
            return
        w = round(w + (0.01 if got < px else -0.01), 2)


def _widths_px(ws, spec):
    for letter, px in spec:
        _width_px(ws, letter, px)


def _unmerge_row(ws, row, c1, c2):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= row <= mr.max_row and not (mr.max_col < C.col(c1) or mr.min_col > C.col(c2)):
            ws.unmerge_cells(str(mr))


def _clear_row(ws, row, c1, c2, values=True):
    _unmerge_row(ws, row, c1, c2)
    for c in C.iter_cells(ws, c1, row, c2, row):
        if values:
            c.value = None
            c.hyperlink = None
        c.fill = NOFILL
        c.border = Border()


def _typo(text):
    """Typografie für Fließtexte: geschütztes Leerzeichen nach „§/§§“ und in Abkürzungen (z. B., Abs. 1, Nr. 2 …),
    damit Paragraf und Nummer nie am Zeilenende getrennt werden (P3-07)."""
    if not isinstance(text, str) or C.is_formula(text):
        return text
    text = re.sub(r"(§§?) ", lambda m: m.group(1) + NBSP, text)
    text = re.sub(r"\b(Abs|Nr|S|Satz|Art|v|i|z|d|F|u|m|a|lt)\. (?=[\w\d(])", lambda m: m.group(1) + "." + NBSP, text)
    return text


def _plain(v):
    return str(v) if v is not None else ""


OBJ_CONTEXT = ("=Obj_Name", '=IFERROR(IF(Erstellt_fuer="","","Erstellt für "&Erstellt_fuer),"")')


def _header(ws, eyebrow, title, subtitle, last):
    """Seitenkopf-Vorlage Z. 4–7 (P04, core.page_header): Z. 5 Eyebrow links (rechts die Unterreiter-Formen),
    Z. 6 Titel links + Objekt rechts, Z. 7 Untertitel links + „Erstellt für …“ rechts – alle rechten Elemente
    bündig an der rechten Inhaltskante (Spalte `last`). Das Eyebrow ist zugleich Zell-Link zur Startseite
    (Rückfall ohne Formen-Links, Nutzerentscheidung 1)."""
    for r in (4, 5, 6, 7):
        _unmerge_row(ws, r, 2, last)
        for c in C.iter_cells(ws, 2, r, last, r):
            if c.coordinate not in ("B5", "B6", "B7"):
                c.value = None
            c.fill = NOFILL
            c.border = Border()
            c.hyperlink = None
    C.page_header(ws, "B", C.L(last), eyebrow, title, subtitle, context=OBJ_CONTEXT)
    for r, h in C.H_HDR.items():  # feste Kopfschablone 12 · 18 · 30 · 21,75 pt (P2-03) – wie auf allen Blättern
        ws.row_dimensions[r].height = h
    e = ws["B5"]
    e.hyperlink = Hyperlink(ref="B5", location=C.link_loc("Start"), display=e.value,
                            tooltip="Zur Startseite – alle Bereiche")


def _meta_cell(ws, row, c1, c2, title_px, text, stop=None):
    """Meta rechtsbündig im Band (P2-10: 8 pt 1D4F8A, nie inline) – in einem Verbund von der ersten Spalte hinter dem
    Titel bis `c2`, damit auch lange Quellenangaben einzeilig und ohne Überlauf stehen."""
    ci1, ci2 = C.col(c1), C.col(c2)
    start = ci1 + 1
    while start < ci2 and C.span_px(ws, c1, C.L(start - 1)) < title_px + 24:
        start += 1
    if start < ci2:
        C.safe_merge(ws, C.L(start), row, C.L(ci2), row)
    m = ws.cell(row, start)
    C.set_text(m, _typo(text))
    m.font = font(T_MICRO, False, BLUE)
    m.alignment = align("right", "center", 1)
    return m


def _section(ws, row, c1, c2, title, extra=None, num=None, up=False):
    """Abschnittskopf Ebene 1 (core.section) – Band trägt nur Titel und Meta (P2-10): Nummer in Akzentfarbe,
    Meta rechtsbündig 8 pt 1D4F8A; mit up=True steht rechts daneben „↑ Übersicht“ (Zell-Link, eigene Spalte)."""
    _clear_row(ws, row, c1, c2)
    C.section(ws, row, c1, c2, title, level=1, meta="↑ Übersicht" if up else None)
    parts = []
    if num:
        parts += [(num, T_H3, True, ACCENT), ("   ", T_H3, True, NAVY)]
    parts.append((title, T_H3, True, NAVY))
    first = ws.cell(row, C.col(c1))
    first.value = C.rich(parts)
    title_px = sum(C.text_width(t, T_H3, True) for t, *_ in parts) + 14
    if up:
        cell = ws.cell(row, C.col(c2))
        C.text_link(cell, "↑ Übersicht", ws.title, C.link_target(ws.title), size=C.T_LABEL, bold=False,
                    tooltip="Zum Seitenanfang mit der Sprungleiste")
        cell.alignment = align("right", "center", 1)
        if extra:
            _meta_cell(ws, row, c1, C.L(C.col(c2) - 1), title_px, extra + "   ·")
    elif extra:
        _meta_cell(ws, row, c1, c2, title_px, extra)
    return first


def _footer(ws, row, c1, c2, gap=C.H_GAP):
    for r in (row, row + 1):
        _clear_row(ws, r, c1, c2)
    C.footer(ws, row, c1, c2)
    ws.row_dimensions[row - 1].height = gap


def _row_height(ws, r, c1, c2, hints=None):
    """core.fit_row mit 4 pt Innenabstand: 1/2/3 Zeilen → 18/30/42 pt, danach n × 13 + 8."""
    return C.fit_row(ws, r, c1, c2, pad=4, hints=hints)


# ================================================================================================ Eingaben
# Abschnitte: Zeile → (Nummer, Titel, Zusatz)
EIN_BANDS = {
    9: ("1", "Objekt und Kaufpreis", "Stammdaten aus Schritt 01 und 02"),
    24: ("2", "Kaufnebenkosten", "Anschaffungsnebenkosten – anteilig mit dem Gebäude abgeschrieben"),
    39: ("3", "Finanzierungsnebenkosten", "sofort abziehbare Werbungskosten bzw. Betriebsausgaben"),
    48: ("4", "Kaufpreisaufteilung", "Grund und Boden / Gebäude · AfA-Bemessungsgrundlage"),
    61: ("5", "Maßnahmen nach Erwerb", "Renovierung, Sonderumlagen, Reserven"),
    75: ("6", "Miete und Bewirtschaftung", "Monatswerte, Stand Kauf"),
    95: ("7", "Finanzierung", "Darlehen II nur hier eingeben"),
    119: ("8", "Steuerliche Optionen", "Rechtsstand September 2026"),
    141: ("9", "Verkaufsszenario", "Exit am Ende der Haltedauer"),
}
EIN_LAST = 145
EIN_DAVON = (20, 22, 80, 81)
# Berechnete Beträge unter gleichnamigen Annahmen (P1-15): eigene Beschriftung, eingerückt
EIN_DERIVED = {33: "Notar – Betrag", 34: "Grundbuch – Betrag", 35: "Makler – Betrag",
               44: "Grundschuldbestellung – Betrag", 45: "Disagio – Betrag"}
# Haarlinie über der ersten berechneten Zeile eines Abschnitts (erst Annahmen, dann Ergebnisse, dann Summe)
EIN_RULE_ABOVE = (33, 44, 54, 90, 135)
# Summenhierarchie (Runde 3, P12): je Abschnitt GENAU EINE Endsumme ('final'), Zwischensummen 'sub'
# (ohne Fläche, Oberlinie D5DEEA), Kennzahlen nach der Endsumme nachrichtlich ('memo', 9 pt kursiv) –
# so steht die Doppellinie nie mitten im Abschnitt als scheinbares Ende.
EIN_SUB = (58, 112, 113)                    # P2-06: 56 (Gebäudeanteil) ist eine Kennzahl, keine Summe → regulär
EIN_SUB_NOTOP = (113,)                      # zweite Zwischensumme direkt unter der ersten
EIN_RESULT = (36, 46, 59, 71, 92, 114, 138)
EIN_MEMO = (37, 72, 73, 93, 115, 116, 117, 139)
# Zahlenformate (P2-12/P24): zweite Nachkommastelle nur bei Zinsen/Tilgung, Makler und Steuersätzen
EIN_PCT2 = {31, 98, 99, 101, 105, 106, 108, 111, 124, 136, 137, 138, 139}
EIN_INPUTS = range(104, 112)
# Einheiten im Zahlenformat (P38): Jahre / Monate / m² – Spalte E nur noch für Qualifier („pro Monat“, „je m²“)
_M = C.MINUS
FMT_YEARS = '[=1]0" Jahr";[=0]"–";0" Jahre"'        # berechnet/übernommen: 0 → „–“
FMT_MONTHS = '[=1]0" Monat";[=0]"–";0" Monate"'
FMT_QM = f'#,##0" m²";"{_M}"#,##0" m²";"–"'
EIN_YEARS = {21, 68, 100, 103, 110, 107, 129, 143}
EIN_MONTHS = {85}
EIN_QM = {14, 15}
EIN_UNIT_CLEAR = EIN_YEARS | EIN_MONTHS | EIN_QM | {91, 93}
STEP_PREFIX = re.compile(r"^Eingabe im Leitfaden, Schritt (\d\d)\s*(?:–\s*)?")

# Beschriftung kürzen (P1-04): Zeile → (neue Beschriftung, Zusatz vor dem Hinweis oder None)
EIN_LABELS = {
    15: ("Grundstücksfläche bzw. Miteigentumsanteil", None),
    20: ("davon bewegliche Gegenstände", "Einbauküche, Möbel – gesondert ausgewiesen"),
    26: ("Grunderwerbsteuersatz abweichend", "leer = automatisch nach Bundesland"),
    32: ("Sonstige Erwerbsnebenkosten", "Gutachten, Fahrtkosten, Übersetzungen"),
    42: ("Bearbeitungsgebühr / Wertermittlung / Vermittlung", "Bank bzw. Vermittler"),
    54: ("Kaufpreis Immobilie (ohne Inventar und Rücklage)", None),
    57: ("Anschaffungskosten Grund und Boden", "inkl. anteiliger Kaufnebenkosten; nicht abschreibbar"),
    58: ("Anschaffungskosten Gebäude = AfA-Basis", "inkl. anteiliger Kaufnebenkosten"),
    59: ("Anschaffungskosten gesamt", "Kaufpreis + Kaufnebenkosten"),
    68: ("Verteilung Erhaltungsaufwand (§ 82b EStDV)", None),
    70: ("Liquiditätsreserve einmalig", "nicht steuerrelevant, z. B. Mieterwechsel"),
    78: ("Umlagefähige Betriebskosten", "Vorauszahlung Mieter"),
    80: ("davon nicht umlagefähig", "Verwaltung WEG, Rücklage, Sonstiges"),
    81: ("davon Zuführung zur Erhaltungsrücklage", "in der nicht umlagefähigen Position enthalten"),
    85: ("Anfangsleerstand (Renovierung / Neuvermietung)", None),
    86: ("Sonstige Werbungskosten p. a.", "Steuerberatung, Kontoführung, Fahrten"),
    101: ("Darlehen I – Anschlusszins (Annahme)", "nach Ablauf der Zinsbindung"),
    113: ("Gesamtinvestition", "Kaufpreis + Kaufnebenkosten + Finanzierungsnebenkosten + Maßnahmen Jahr 1"),
    121: ("Rechtsform", None),
    123: ("Zu versteuerndes Einkommen ohne dieses Objekt", "nur Privatperson"),
    124: ("Grenzsteuersatz manuell", "leer = automatisch nach Tarif 2026"),
    125: ("Solidaritätszuschlag (nur Privatperson)", None),
    127: ("Gewerbesteuer-Hebesatz (nur gewerbliche GmbH)", "ohne erweiterte Kürzung"),
    129: ("Restnutzungsdauer lt. Gutachten", "nur Methode Gutachten"),
    130: ("Degressiv: Wechsel zur linearen AfA", "automatisch, sobald vorteilhaft"),
    131: ("Sonder-AfA § 7b EStG (Mietwohnungsneubau)", None),
    132: ("Erhöhte Absetzungen § 7h / § 7i EStG", "Sanierungsgebiet / Baudenkmal"),
    133: ("Begünstigte Sanierungskosten § 7h / § 7i", "im Kaufpreis enthalten"),
    138: ("Grenzbelastung Privat (ESt + Soli + KiSt)", None),
    145: ("Verkaufspreis manuell", "leer = Wertentwicklung lt. Prognose"),
}
# Redaktionell gekürzte Hinweise – möglichst einzeilig in G:L (gleichmäßiger 18-pt-Rhythmus, P27)
EIN_HINTS = {
    15: "ETW: Miteigentumsanteil × Grundstücksfläche (lt. Teilungserklärung); nur Bodenrichtwert-Methode",
    20: "Einbauküche, Möbel: grunderwerbsteuerfrei (§ 2 GrEStG), gesondert linear abschreibbar",
    22: "Gehört nicht zu den Gebäude-AK (nicht abschreibbar), unterliegt aber der GrESt (BFH II R 49/17)",
    50: "BMF-Arbeitshilfe; abweichend per Gutachten oder plausibler Kaufvertragsaufteilung (BFH IX R 26/19)",
    67: "Automatik: netto > 15 % der Gebäude-AK in 3 Jahren → Herstellungskosten (AfA statt Sofortabzug)",
    68: "1 Jahr = Sofortabzug; 2–5 Jahre = gleichmäßige Verteilung (nur Privatvermögen, Wohngebäude)",
    70: "Nicht steuerrelevant, z. B. Mieterwechsel – erhöht nur den Eigenkapitalbedarf",
    81: "Im nicht umlagefähigen Hausgeld enthalten – steuerlich erst bei Verausgabung (BFH IX R 19/24)",
    127: "Ohne erweiterte Kürzung, z. B. 400 %; mit erweiterter Kürzung entfällt die GewSt auf Mieterträge",
    129: "Nur Methode Gutachten: kürzere Nutzungsdauer per Gutachten (BFH IX R 25/19; BMF 22.2.2023)",
    130: "Automatisch, sobald vorteilhaft (§ 7 Abs. 5a S. 4 EStG): linear nach Restwert und Restnutzungsdauer",
    131: "5 % p. a. für 4 Jahre; Bauantrag 1.1.2023 – 30.9.2029, EH 40 + QNG, Baukosten ≤ 5.200 €/m²",
    132: "Sanierungsgebiet / Baudenkmal: 9 % p. a. (Jahre 1–8), 7 % p. a. (Jahre 9–12) auf bescheinigte Kosten",
}
EIN_HEAD = {"B": ("POSITION", "left"), "C": ("WERT", "right"),
            "F": ("QUELLE", "left"), "G": ("HINWEIS / STEUERLICHE EINORDNUNG", "left")}


def _step_sheets(wb):
    out = {}
    for name in wb.sheetnames:
        m = re.match(r"S(\d\d) ", name)
        if m:
            out[m.group(1)] = name
    return out


def _name_target(wb, formula):
    """„=Name“ → Zieladresse des Namens („'S01 Objekt'!D12“) oder None."""
    if not (isinstance(formula, str) and re.match(r"^=[A-Za-z_][A-Za-z0-9_.]*$", formula)):
        return None
    dn = wb.defined_names.get(formula[1:])
    if dn is None:
        return None
    t = str(dn.attr_text).replace("$", "")
    if "!" not in t or ":" in t:
        return None
    sh, addr = t.rsplit("!", 1)
    return f"'{sh.strip(chr(39))}'!{addr}"


def _choices(wb):
    """Name einer Leitfaden-Eingabe → mögliche Anzeigetexte (Auswahlliste ihrer Datenüberprüfung)."""
    rx = re.compile(r"^'?([^'!]+)'?!([A-Z]+)(\d+):([A-Z]+)(\d+)$")

    def values(ref):
        ref = ref.lstrip("=").replace("$", "")
        dn = wb.defined_names.get(ref)
        if dn is not None:
            ref = str(dn.attr_text).replace("$", "")
        m = rx.match(ref)
        if not m or m.group(1) not in wb.sheetnames:
            return None
        ws = wb[m.group(1)]
        return [str(c.value) for c in C.iter_cells(ws, m.group(2), int(m.group(3)), m.group(4), int(m.group(5)))
                if c.value is not None]

    by_cell = {}
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            if dv.type != "list" or not dv.formula1:
                continue
            vals = values(dv.formula1)
            if not vals:
                continue
            for cr in dv.sqref.ranges:
                for c in C.iter_cells(ws, cr.min_col, cr.min_row, cr.max_col, cr.max_row):
                    by_cell[f"{ws.title}!{c.coordinate}"] = vals
    out = {}
    for n, dn in wb.defined_names.items():
        t = str(dn.attr_text).replace("$", "").replace("'", "")
        if t in by_cell:
            out[n] = by_cell[t]
    return out


def _ein_fmt(r, fmt, entry):
    """Zahlenformat nach Katalog (P2-12/P38): Anzeige mit „–“ für 0, Eingabefelder (gelb) mit sichtbarer 0;
    Jahre / Monate / m² als Einheit im Format."""
    if r == 93:
        return NUMFMT["mult1"]
    if r == 16:
        return NUMFMT["year"]
    if r in EIN_YEARS:
        return NUMFMT["years_in"] if entry else FMT_YEARS
    if r in EIN_MONTHS:
        return FMT_MONTHS
    if r in EIN_QM:
        return FMT_QM
    if fmt is None or fmt == "General" or "yy" in fmt.lower():
        return fmt
    if "%" in fmt:
        if r == 127:
            return NUMFMT["pct0"]
        return C.numfmt("pct2" if r in EIN_PCT2 else "pct1", entry)
    if "€" in fmt:
        return C.numfmt("eur2" if ".00" in fmt else "eur", entry)
    if fmt.startswith("#,##0"):
        return C.numfmt("num", entry)
    return fmt


def eingaben(wb):
    ws = wb["Eingaben"]
    L_ = 12  # letzte Tabellenspalte L
    steps = _step_sheets(wb)
    global _CHOICES
    _CHOICES = _choices(wb)

    # ---- Spalten (P04/Befund „rechte Kante“): Inhalt A:L = 1 368 px = rechte Kante der Kopfleiste; die Randspalte M
    #      liegt dahinter (Kopfleiste endet an L). Wertspalte C:D 350 px, damit Auswahltexte links einzeilig stehen (P28).
    _ungroup_cols(ws)
    # P1-13 (Runde 4): Wertachse = rechte Kante von C (112 px, so breit wie ein Eingabefeld); D ist eine 8-px-Fuge,
    # E trägt die Einheit direkt hinter der Zahl. Auswahl-/Texte stehen linksbündig über C:E (438 px, 10 pt einzeilig).
    _widths_px(ws, (("B", 300), ("C", 112), ("D", 8), ("E", 318), ("F", 56),
                    ("G", 90), ("H", 90), ("I", 90), ("J", 90), ("K", 90), ("L", 93)))
    _width(ws, "M", 3)

    # ---- Seitenkopf (P04): Objekt / „Erstellt für“ rechts an der Inhaltskante
    _header(ws, "Eingaben", "Eingaben",
            "Alle Annahmen auf einen Blick – geändert wird im Leitfaden; direkt editierbar sind nur die "
            "Profi-Felder für Darlehen II.", L_)
    # Legende rechts neben der Sprungleiste (Zeile 8, Formen von navigation.py) – bündig an der Inhaltskante
    _clear_row(ws, 8, 2, L_)
    C.safe_merge(ws, "H", 8, "L", 8)
    lg = ws["H8"]
    lg.value = C.rich([("■  ", T_BODY, False, INPUT_LINE), ("Eingabe hier", T_SMALL, True, BLUE),
                       ("        ■  ", T_BODY, False, BLUE), ("aus dem Leitfaden – dort ändern", T_SMALL, False, BLUE),
                       ("        ■  ", T_BODY, False, INK), ("berechnet", T_SMALL, False, INK)])
    lg.font = font(T_SMALL, False, MUTED)
    lg.alignment = align("right", "center")
    ws.row_dimensions[8].height = 30

    # ---- Abschnitte: Ebene 1 (hell) + Tabellenkopf Ebene 2 + Datenzeilen
    band_rows = sorted(EIN_BANDS)
    for i, br in enumerate(band_rows):
        num, title, extra = EIN_BANDS[br]
        _section(ws, br, "B", "L", title, extra, num=num, up=True)
        head = br + 1
        end = (band_rows[i + 1] - 2) if i + 1 < len(band_rows) else EIN_LAST
        _clear_row(ws, head, 2, L_)
        C.section(ws, head, "B", "L", None, level=2, labels=EIN_HEAD)
        C.safe_merge(ws, "G", head, "L", head)
        for r in range(head + 1, end + 1):
            _data_row(ws, wb, steps, r, L_)
        gap = end + 1
        _clear_row(ws, gap, 2, L_, values=False)
        ws.row_dimensions[gap].height = C.H_GAP

    # ---- Haarlinie vor den berechneten Werten (P1-15)
    for r in EIN_RULE_ABOVE:
        for c in C.iter_cells(ws, 2, r - 1, L_, r - 1):
            c.border = Border(bottom=side("thin", LINE2))
        for c in C.iter_cells(ws, 2, r, L_, r):
            c.border = Border(top=side("thin", LINE2), bottom=side("hair", LINE))

    # ---- Summenhierarchie (P12): Zwischensumme · genau eine Endsumme je Abschnitt · Kennzahlen nachrichtlich
    for r in EIN_SUB + EIN_RESULT:
        stage = "sub" if r in EIN_SUB else "final"
        C.sum_row(ws, r, "B", "L", stage, top=r not in EIN_SUB_NOTOP)
        ws.cell(r, 2).alignment = align("left", "center", 1, wrap=True)
        for cc in (5, 7):
            ws.cell(r, cc).font = font(T_SMALL, False, MUTED)
    for r in EIN_MEMO:
        C.sum_row(ws, r, "B", "L", "memo")

    # ---- Profi-Felder (Darlehen II): Eingabeoptik NUR auf der Wertzelle C (112 px wie die Eingabefelder der Mappe);
    #      gemeinsame 1-px-Kanten E6CB77 trennen die acht Felder (P1-13), D/E bleiben weiß
    for r in EIN_INPUTS:
        c = ws.cell(r, 3)
        c.font = font(T_BODY, True, BLUE)
        C.input_style(c, "required")
        c.alignment = align("right", "center", 1)
        for cc in (4, 5):
            ws.cell(r, cc).fill = NOFILL

    # Datenüberprüfungen der Profi-Felder melden Fehler (Logik unverändert, nur Meldung aktiv)
    for dv in ws.data_validations.dataValidation:
        dv.showErrorMessage = True
        dv.errorStyle = "stop"
        dv.errorTitle = dv.errorTitle or "Ungültiger Wert"
        if dv.type == "decimal":
            dv.error = "Bitte einen Prozentsatz zwischen 0 % und 100 % eingeben (z. B. 2,5 %)."
        dv.showInputMessage = True
        dv.promptTitle = dv.promptTitle or "Profi-Feld Darlehen II"
        dv.prompt = dv.prompt or ("Satz in % eingeben (z. B. 2,5 %)." if dv.type == "decimal"
                                  else "Zinsbindung in Jahren (1–40).")

    _footer(ws, EIN_LAST + 2, 2, L_)


def _data_row(ws, wb, steps, r, last):
    b, cval, e, f = ws.cell(r, 2), ws.cell(r, 3), ws.cell(r, 5), ws.cell(r, 6)
    if b.value is None and cval.value is None:
        return
    # Hinweis aus F:L lösen → Quelle (F) + Hinweis (G:L)
    hint = f.value if not C.is_formula(f.value) else None
    _unmerge_row(ws, r, 6, last)
    for c in C.iter_cells(ws, 6, r, last, r):
        c.value = None
        c.hyperlink = None
    link = None
    if isinstance(hint, str):
        m = STEP_PREFIX.match(hint)
        if m:
            link = m.group(1)
            hint = hint[m.end():].strip() or None
    # Beschriftung kürzen, Zusatz in den Hinweis
    if r in EIN_LABELS:
        new, extra = EIN_LABELS[r]
        C.set_text(b, new)
        if extra:
            hint = f"{extra} – {hint}" if hint else extra
    if r in EIN_DERIVED:
        C.set_text(b, EIN_DERIVED[r])
    if r in EIN_HINTS:
        hint = EIN_HINTS[r]
    if hint:
        hint = hint.replace(" | ", " · ")
        hint = _typo(hint[0].upper() + hint[1:])
    # Einheit genau einmal: Jahre / Monate / m² stehen im Zahlenformat (P38)
    if r in EIN_UNIT_CLEAR or (isinstance(e.value, str) and e.value.strip() in ("€", "%")):
        e.value = None
    # Zeilenstil: keine senkrechten Linien, Haarlinie über die volle Breite
    for c in C.iter_cells(ws, 2, r, last, r):
        c.fill = NOFILL
        c.border = Border(bottom=side("hair", LINE))
    sub = r in EIN_DAVON or r in EIN_DERIVED
    b.font = font(T_BODY, False, INK2 if sub else INK)
    b.alignment = align("left", "center", 2 if sub else 1, wrap=True)
    col_c = cval.font.color.rgb[-6:] if (cval.font.color is not None and isinstance(cval.font.color.rgb, str)) else INK
    linked = col_c.upper() == BLUE
    entry = r in EIN_INPUTS
    fmt = cval.number_format
    is_text = (fmt in (None, "General")) and not isinstance(cval.value, (int, float)) and r != 16
    cval.number_format = _ein_fmt(r, fmt, entry)
    # P1-13: alle Werte 10 pt. Zahlen rechtsbündig an der Wertachse (rechte Kante C, Einzug 1), Einheit direkt dahinter;
    # Text-/Auswahlwerte linksbündig mit Einzug 1 über C:E – dort steht nie eine Einheit, also ragt kein Wert hinein.
    opts = _CHOICES.get(cval.value[1:], []) if (is_text and C.is_formula(cval.value)) else []
    longest = max(opts, key=lambda t: C.text_width(t)) if opts else None
    if is_text and ws.cell(r, 4).value is None and (e.value is None or r in EIN_UNIT_CLEAR):
        e.value = None
        C.safe_merge(ws, "C", r, "E", r)
    cval.font = font(T_BODY, False, BLUE if linked else INK)
    cval.alignment = align("left" if is_text else "right", "center", 1, wrap=is_text)
    # Einheit: 9 pt grau, ohne Einzug direkt hinter der Zahl (C-Einzug + 8-px-Fuge D)
    e.font = font(T_SMALL, False, MUTED)
    e.alignment = align("left", "center", 0)
    # Quelle-Link: feldgenau (P1-08), Beschriftung nach dem echten Zielblatt
    if link and link in steps:
        sheet = steps[link]
        target = _name_target(wb, cval.value) or C.link_loc(sheet)
        tsheet = target.split("!")[0].strip("'")
        if tsheet == sheet or tsheet not in wb.sheetnames:
            label, tip = f"S{link} ›", f"Im Leitfaden ändern: Schritt {int(link)} · {sheet[4:]}"
        elif tsheet == "Start":
            label, tip = "Start ›", "Auswahl auf der Startseite"
        else:
            label, tip = f"{tsheet} ›", f"Auswahl auf dem Blatt „{tsheet}“"
        C.set_text(f, label)
        f.font = font(T_SMALL, True, BLUE)
        f.hyperlink = Hyperlink(ref=f.coordinate, location=target, display=label, tooltip=tip)
        f.alignment = align("left", "center", 1)
    # Hinweis G:L
    g = ws.cell(r, 7)
    if hint:
        C.set_text(g, hint)
    g.font = font(T_SMALL, False, MUTED)
    g.alignment = align("left", "center", 1, wrap=True)
    C.safe_merge(ws, "G", r, "L", r)
    # Zeilenhöhe: Calibri-Metrik; Auswahlfelder nach der längsten Option (Excel passt die Höhe nicht nach)
    _row_height(ws, r, "B", "L", {"C": longest} if longest else None)


# ================================================================================================ Konfiguration
KON_BANDS = {  # Zeile → (Titel, Meta rechts: Rechtsgrundlage/Quelle) – Ebene 1 gruppiert, Ebene 2 benennt die Tabelle (P2-10)
    8: ("Grunderwerbsteuer und Einkommensteuer", "§ 11 GrEStG, Landesgesetze · § 32a EStG · Stand 2026"),
    43: ("Körperschaftsteuer und Konstanten",
         "§ 23 KStG: ab 2028 −1 %-Punkt p. a. bis 10 % (2032), zzgl. Soli 5,5 %"),
    74: ("Ampel-Schwellenwerte", "wirken auf Cockpit, Dashboard, Leitfaden und Sensitivität"),
}
# Tabellenköpfe Ebene 2 (EEF3FA, Unterlinie 4A86C8) – jede Tabelle hat genau einen, erste Spalte benennt die Zeilen
KON_HEADS = {
    9: {"B": ("Grunderwerbsteuer · Bundesland", "left"), "C": ("Steuersatz", "right"), "D": ("Gültig seit", "center")},
    27: {"B": ("Einkommensteuertarif 2026 · Position", "left"), "C": ("Wert", "right"), "E": ("Hinweis", "left")},
    44: {"B": ("Körperschaftsteuer · Kalenderjahr", "left"), "C": ("Steuersatz", "right")},
    56: {"B": ("Konstanten und Schwellenwerte · Position", "left"), "C": ("Wert", "right"), "E": ("Hinweis", "left")},
    75: {"B": ("Kennzahl", "left"), "C": ("Grün ab", "right"), "D": ("Gelb ab", "right"), "E": ("Hinweis", "left")},
}
KON_TABLES = [(10, 25), (28, 40), (45, 54), (57, 72), (76, 80)]
KON_GAPS = (26, 41, 55, 73)
KON_PANEL_HEADS = {9: "OBJEKTART", 15: "METHODE KAUFPREISAUFTEILUNG", 20: "STEUERLICHE BEHANDLUNG RENOVIERUNG",
                   25: "RECHTSFORM", 30: "VERANLAGUNG", 34: "JA / NEIN", 38: "KIRCHENSTEUER",
                   43: "AfA-METHODE GEBÄUDE", 51: "VERLUSTVERRECHNUNG"}
# Redaktionell gekürzte Erläuterungen (einzeilig in Spalte E, P1-04/P3-07)
KON_NOTES = {
    31: "69.879 – 277.825 €: 0,42·x – 11.135,63",
    37: "ab 277.826 €: 0,45·x – 19.470,38",
    38: "5,5 % der ESt/KSt; Freigrenze 2026: 20.350 € / 40.700 € ESt",
    58: "innerhalb von 3 Jahren nach Anschaffung (AK Gebäude)",
    59: "im Jahr der Anschaffung und in den 3 Folgejahren",
    61: "je m² Wohnfläche · bei Überschreitung entfällt § 7b ganz",
    62: "je m² Wohnfläche · darüber keine Sonder-AfA",
    66: "Verkauf nach > 10 Jahren steuerfrei (Privatvermögen)",
    67: "größerer Erhaltungsaufwand, Wohngebäude (Privat)",
    68: "bis 5 % sofort abziehbar bei ≥ 5 Jahren Zinsbindung",
    69: "BMF-Schreiben vom 20.10.2003",
    71: "zzgl. Soli = 26,375 % · Information, nicht im Cashflow",
}
KON_LABELS = {58: "Grenze anschaffungsnahe Herstellungskosten", 61: "Sonder-AfA § 7b: Baukostenobergrenze je m²", 62: "Sonder-AfA § 7b: max. Bemessungsgrundlage je m²",
              80: "EK-Rendite Jahr 1 (Vermögenszuwachs / EK)"}


def _kon_fmt(c, label):
    """Konfiguration: alle Werte sind Eingaben → Formate mit sichtbarer 0 (P2-12)."""
    f, v = c.number_format or "General", c.value
    if isinstance(label, str) and label.rstrip().endswith("(Jahre)"):
        return NUMFMT["years_in"]
    if c.coordinate in ("C78", "D78"):
        return NUMFMT["mult2"]
    if c.coordinate in ("C32", "C33", "C34", "C35"):
        return NUMFMT["num2_in"]
    if c.coordinate == "C54":
        return NUMFMT["year"]
    if "%" in f:
        if isinstance(v, (int, float)) and abs(round(v * 1000) - v * 1000) > 1e-6:
            return NUMFMT["pct2_in"]
        return NUMFMT["pct1_in"]
    if "€" in f:
        return NUMFMT["eur_in"]
    if isinstance(v, (int, float)) and v >= 1000:
        return NUMFMT["num_in"]
    return f


def konfiguration(wb):
    ws = wb["Konfiguration"]
    last = 7  # G
    _ungroup_cols(ws)
    # Breiten (P26): Rinne F 40 px (≥ 4 Zeichen) zwischen Tabellen und Auswahllisten; Summe A:G = 1 368 px
    _widths_px(ws, (("B", 301), ("C", 76), ("D", 70), ("E", 308), ("F", 40), ("G", 542)))

    _header(ws, "Anhang  ›  Konfiguration", "Konfiguration", None, last)
    w = ws["B7"]
    w.value = C.rich([("▲  Expertenbereich", T_BODY, True, AMBER),
                      ("   ·   Änderungen wirken auf alle Blätter – Reihenfolge der Auswahllisten nicht ändern",
                       T_BODY, False, MUTED)])
    w.font = font(T_BODY, False, MUTED)
    w.alignment = align("left", "center")

    # Quellen wandern in die Abschnittsköpfe (Zusatz), die Fußnotenzeilen werden Abstandszeilen
    for a in ("E9", "E10", "B26", "B41", "E45"):  # E45: KSt-Hinweis steht jetzt im Abschnittskopf (P29)
        ws[a].value = None

    # ---- Datenzeilen B:E
    for r1, r2 in KON_TABLES:
        for r in range(r1, r2 + 1):
            for c in C.iter_cells(ws, 2, r, 5, r):
                c.fill = NOFILL
                c.border = Border(bottom=side("hair", LINE))
            b = ws.cell(r, 2)
            if r in KON_LABELS:
                C.set_text(b, KON_LABELS[r])
            label = b.value
            b.font = font(T_BODY, False, INK)
            b.alignment = align("left", "center", 1, wrap=True)
            if isinstance(b.value, int):
                b.number_format = NUMFMT["year"]
            for cc in (3, 4):
                c = ws.cell(r, cc)
                if c.value is None:
                    continue
                if cc == 4 and isinstance(c.value, str):  # Datum als Text (gültig seit)
                    c.font = font(T_SMALL, False, MUTED)
                    c.alignment = align("center", "center")
                    continue
                c.number_format = _kon_fmt(c, label)
                c.font = font(T_BODY, False, INK)
                c.alignment = align("right", "center", 1)
                if c.protection.locked is False:
                    C.input_style(c, "required")
                    c.alignment = align("right", "center", 1)
            if isinstance(label, str) and label.rstrip().endswith("(Jahre)"):
                C.set_text(b, label.rstrip()[:-len("(Jahre)")].rstrip())
            e = ws.cell(r, 5)
            if r in KON_NOTES:
                C.set_text(e, KON_NOTES[r])
            if isinstance(e.value, str) and not C.is_formula(e.value):
                C.set_text(e, _typo(e.value))
            e.font = font(T_SMALL, False, MUTED)
            e.alignment = align("left", "center", 1, wrap=True)
            _row_height(ws, r, "B", "E")

    # ---- Panel G (Auswahllisten) – Listenwerte bleiben unverändert (Datenquellen der Dropdowns)
    for r in range(9, 54):
        g = ws.cell(r, 7)
        if r in KON_PANEL_HEADS:
            h = ws.row_dimensions[r].height
            C.section(ws, r, "G", "G", KON_PANEL_HEADS[r], level=2, caps=False)
            ws.row_dimensions[r].height = h
        elif g.value is not None:
            g.fill = NOFILL
            g.border = Border(bottom=side("hair", LINE))
            g.font = font(T_BODY, False, INK)
            g.alignment = align("left", "center", 1)
        else:
            g.fill = NOFILL
            g.border = Border()

    # ---- Abschnittsköpfe Ebene 1 (P1-12) und Tabellenköpfe Ebene 2 (P2-10: Spaltenköpfe nie im Band)
    for r, (title, extra) in KON_BANDS.items():
        _section(ws, r, "B", "E", title, extra)
    _section(ws, 8, "G", "G", "Auswahllisten für die Dropdown-Felder")
    for r, labels in KON_HEADS.items():
        _unmerge_row(ws, r, 2, 5)
        for c in C.iter_cells(ws, 2, r, 5, r):
            c.value = None
            c.hyperlink = None
        C.section(ws, r, "B", "E", None, level=2, labels=labels)
        ws.row_dimensions[r].height = C.H_HEAD
    ws["D9"].alignment = align("center", "center")

    # ---- Abstandszeilen: je eine Zeile (18 pt, weil rechts Listeneinträge darin stehen) vor jedem Abschnitt
    for r in KON_GAPS:
        for c in C.iter_cells(ws, 2, r, 5, r):
            c.fill = NOFILL
            c.border = Border()
        ws.row_dimensions[r].height = C.H_ROW
    # Z. 42: Fuge vor „AfA-Methode Gebäude“ rechts (P27) – 12 pt, weil links schon Z. 41 als Abstand dient
    ws.row_dimensions[42].hidden = False
    ws.row_dimensions[42].height = C.H_GAP
    for r in (26, 27, 41):  # Listeneinträge in Sonderzeilen mittig (Befund Auswahllisten)
        ws.cell(r, 7).alignment = align("left", "center", 1)
    # Nullkonvention (P38): Anzeigeformat mit „–“ für die Tarif-/Baukostengrenzen
    for coord in ("C28", "C29", "C30", "C31", "C61", "C62"):
        ws[coord].number_format = NUMFMT["eur"]

    _footer(ws, 82, 2, last, gap=C.H_ROW)


# ================================================================================================ Hinweise
# Redaktionell gestraffte Annahmetexte: klar ein- oder zweizeilig über C:D (keine Grenzfälle, P27)
HIN_TEXT = {
    31: "Konstanter Grenzsteuersatz (Privat) bzw. KSt-Staffel + Soli (+ GewSt) – keine vollständige "
        "Veranlagungsrechnung, keine Progressionswirkung des Objekts auf das übrige Einkommen.",
    33: "Zuführungen zur WEG-Erhaltungsrücklage und zur eigenen Instandhaltungsrücklage sind Liquiditätsabflüsse, "
        "aber nicht sofort steuerwirksam; Verausgabungen der WEG sind nicht separat modelliert.",
}
# Thema (B) · Vereinfachung (C:D) · Einordnung (E, Kurzkategorie – P3-12: vierspaltiges Raster läuft weiter)
HIN_TOPICS = {29: ("Zeitraster", "Vereinfachung"), 30: ("Finanzierung", "Vereinfachung"),
              31: ("Steuersatz", "Vereinfachung"), 32: ("AfA-Kombination", "konservativ"),
              33: ("Rücklagen", "Vereinfachung"), 34: ("GmbH", "nicht abgebildet"),
              35: ("Prognose", "Annahme"), 36: ("Haftung", "Rechtshinweis")}
HIN_WIDTHS = (("B", 210), ("C", 586), ("D", 313), ("E", 228))     # A:E = 1 368 px; E ≈ 32 Zeichen (P3-12), C/D so, dass keine Zeile an einer Umbruchgrenze liegt


def hinweise(wb):
    ws = wb["Hinweise"]
    last = 5
    _ungroup_cols(ws)
    _widths_px(ws, HIN_WIDTHS)
    _header(ws, "Anhang  ›  Hinweise", "Hinweise",
            "Steuerliche Regelungen und Modellannahmen · Rechtsstand September 2026 (Investitionssofortprogramm "
            "2025, JStG 2024) · keine Steuerberatung im Einzelfall", last)

    # ---- EIN Band, zwei Tabellen mit je einem Tabellenkopf (P2-10: Spaltenköpfe nie im Band)
    _section(ws, 8, "B", "E", "Steuerliche Regelungen und Modellannahmen", "17 Regelungen · 8 Annahmen · Kurzfassung")
    _clear_row(ws, 9, 2, last)
    C.section(ws, 9, "B", "E", None, level=2,
              labels={"B": ("THEMA", "left"), "C": ("REGELUNG (KURZFASSUNG)", "left"),
                      "D": ("UMSETZUNG IM TOOL", "left"), "E": ("FUNDSTELLE", "left")})
    # Thema 9 pt fett 0B2A4A · Text 10 pt · Fundstelle 8 pt grau (P3-12); Höhe = Zeilen × 12,5 + 8 pt (fit_row metric)
    spec = {2: (T_SMALL, True, NAVY), 3: (T_BODY, False, INK), 4: (T_BODY, False, INK), 5: (T_MICRO, False, MUTED)}
    for r in range(10, 27):
        for cc, (sz, bd, colr) in spec.items():
            c = ws.cell(r, cc)
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE2))
            if isinstance(c.value, str) and not C.is_formula(c.value):
                C.set_text(c, _typo(c.value))
            c.font = font(sz, bd, colr)
            c.alignment = align("left", "top", 1, wrap=True)
        C.fit_row(ws, r, "B", "E", metric=True)
    _clear_row(ws, 27, 2, last, values=False)
    ws.row_dimensions[27].height = C.H_ROW

    # ---- Modellannahmen: Kopfzeile wie Z. 9 (THEMA · VEREINFACHUNG · EINORDNUNG), Text nur über C:D
    _clear_row(ws, 28, 2, last)
    C.section(ws, 28, "B", "E", None, level=2,
              labels={"B": ("MODELLANNAHME", "left"), "C": ("VEREINFACHUNG IM MODELL", "left"),
                      "E": ("EINORDNUNG", "left")})
    for r in range(29, 37):
        text = re.sub(r"^\s*[•●]\s*", "", _plain(ws.cell(r, 2).value))
        topic, kind = HIN_TOPICS.get(r, ("", ""))
        if topic and text.lower().startswith(topic.lower() + ":"):  # Stichwort nicht wiederholen („GmbH: …“)
            text = text[len(topic) + 1:].lstrip()
            text = text[:1].upper() + text[1:]
        text = HIN_TEXT.get(r, text)
        _clear_row(ws, r, 2, last)
        C.set_text(ws.cell(r, 2), topic)
        ws.cell(r, 2).font = font(T_SMALL, True, NAVY)
        ws.cell(r, 2).alignment = align("left", "top", 1, wrap=True)
        C.safe_merge(ws, "C", r, "D", r)
        t = ws.cell(r, 3)
        C.set_text(t, _typo(text))
        t.font = font(T_BODY, False, INK)
        t.alignment = align("left", "top", 1, wrap=True)
        k = ws.cell(r, 5)
        C.set_text(k, kind)
        k.font = font(T_MICRO, False, MUTED)
        k.alignment = align("left", "top", 1, wrap=True)
        for c in C.iter_cells(ws, 2, r, last, r):
            c.border = Border(bottom=side("hair", LINE2))
        C.fit_row(ws, r, "B", "E", metric=True)
    _footer(ws, 38, 2, last)


# ================================================================================================
def _calibri(ws):
    """Alle Zellen (auch leere Flächen) in Calibri – keine Vorlagenschrift bleibt zurück."""
    for c in list(ws._cells.values()):
        f = c.font
        if f is not None and f.name != SANS:
            nf = copy(f)
            nf.name = SANS
            c.font = nf


def apply(wb):
    eingaben(wb)
    konfiguration(wb)
    hinweise(wb)
    for name in ("Eingaben", "Konfiguration", "Hinweise"):
        _calibri(wb[name])
        C.cf_close(wb[name])
