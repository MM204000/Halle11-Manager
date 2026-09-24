"""Formularblätter Eingaben, Konfiguration, Hinweise (Agent E) – ruhiger „Private Banking“-Stil, Runde 2.

Komponentensprache ausschließlich aus core.py (P1-12/P1-14/P1-18/P2-12):
  Seitenkopf core.page_header (Überzeile „<REITER> › <BLATT>“, Titel 22 pt, Untertitel 10 pt, Meta rechts 9 pt) ·
  Abschnittskopf core.section (Ebene 1 hell E7EEF7 + Akzentkante, Ebene 2 EEF3FA-Tabellenkopf) ·
  Summenstufen core.sum_row (Zwischensumme / Blockergebnis) · Zahlenformate core.NUMFMT/numfmt ·
  Zeilenhöhen core.fit_row (Calibri-Metrik, Raster 18/30/42/n×13+8) · Rich-Text core.rich · Fuß core.footer.

Inhaltsbreite aller drei Blätter: Kopfleiste endet bei 1 368 px (core.NAV_MIN_PX), P3-07/Befund „Inhaltsbreite“:
  Eingaben    A 31 | B 308 | C:D 322 | E 87 | F 56 | G:L 543  → 1 347 px (+ Randspalte M 21 = 1 368)
  Konfiguration A 31 | B 301 | C 84 | D 77 | E 308 | F 14 | G 553 → 1 368 px
  Hinweise    A 31 | B 210 | C 511 | D 385 | E 231 → 1 368 px (Breiten nach minimaler Gesamtzeilenzahl)

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


def _header(ws, eyebrow, title, subtitle, last, meta=None):
    """Seitenkopf Z. 4–7 nach Standard (P1-18) – Flächen/Linien rechts davon leeren."""
    for r in (4, 5, 6, 7):
        _unmerge_row(ws, r, 2, last)
        for c in C.iter_cells(ws, 2, r, last, r):
            if c.coordinate not in ("B5", "B6", "B7"):
                c.value = None
            c.fill = NOFILL
            c.border = Border()
            c.hyperlink = None
    C.page_header(ws, "B", C.L(last), eyebrow, title, subtitle)
    ws.row_dimensions[4].height = 10
    if meta:
        m = ws.cell(7, last)
        m.value = meta
        m.font = font(T_SMALL, False, MUTED)
        m.alignment = align("right", "top", 1)


def _section(ws, row, c1, c2, title, extra=None, num=None, up=False):
    """Abschnittskopf Ebene 1 (core.section) – Nummer in Akzentfarbe, Zusatz 9 pt grau, „↑ Übersicht“ rechts."""
    _clear_row(ws, row, c1, c2)
    C.section(ws, row, c1, c2, title, level=1, meta="↑ Übersicht" if up else None)
    parts = []
    if num:
        parts += [(num, T_H3, True, ACCENT), ("   ", T_H3, True, NAVY)]
    parts.append((title, T_H3, True, NAVY))
    if extra:
        parts.append(("     " + _typo(extra), T_SMALL, False, MUTED))
    first = ws.cell(row, C.col(c1))
    first.value = C.rich(parts)
    if up:
        cell = ws.cell(row, C.col(c2))
        C.text_link(cell, "↑ Übersicht", ws.title, C.link_target(ws.title), size=T_MICRO, bold=False,
                    tooltip="Zum Seitenanfang mit der Sprungleiste")
        cell.alignment = align("right", "center", 1)
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
# Summenstufen (P1-14): genau ein Blockergebnis je Abschnitt
EIN_SUB = (56, 112, 113)
EIN_RESULT = (36, 46, 58, 71, 92, 114, 138)
# Zahlenformate (P2-12): zweite Nachkommastelle nur bei Zinsen/Tilgung, Makler und Steuersätzen
EIN_PCT2 = {31, 98, 99, 101, 105, 106, 108, 111, 124, 136, 137, 138, 139}
EIN_INT = {21, 68, 85, 100, 103, 107, 110, 129, 143}
EIN_INPUTS = range(104, 112)
# Einheit genau einmal (P3-08); § 82b: „Jahr(e)“ (P2-12)
EIN_UNIT = {91: None, 93: None, 68: "Jahr(e)"}
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
# Redaktionell gekürzte Hinweise (einzeilig in G:L)
EIN_HINTS = {22: "Gehört nicht zu den Gebäude-AK (nicht abschreibbar), unterliegt aber der GrESt (BFH II R 49/17)"}
EIN_HEAD = {"B": ("POSITION", "left"), "C": ("WERT", "right"), "E": ("EINHEIT", "left"),
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
    """Zahlenformat nach Katalog (P2-12): Anzeige mit „–“ für 0, Eingabefelder mit sichtbarer 0."""
    if r == 93:
        return NUMFMT["mult1"]
    if r == 16:
        return NUMFMT["year"]
    if fmt is None or fmt == "General" or "yy" in fmt.lower():
        return fmt
    if "%" in fmt:
        if r == 127:
            return NUMFMT["pct0"]
        return C.numfmt("pct2" if r in EIN_PCT2 else "pct1", entry)
    if "€" in fmt:
        return C.numfmt("eur2" if ".00" in fmt else "eur", entry)
    if r in EIN_INT:
        return C.numfmt("int0", entry)
    if fmt.startswith("#,##0"):
        return C.numfmt("num", entry)
    return fmt


def eingaben(wb):
    ws = wb["Eingaben"]
    L_ = 12  # letzte Tabellenspalte L
    steps = _step_sheets(wb)
    global _CHOICES
    _CHOICES = _choices(wb)

    # ---- Spalten: Inhaltsbreite 1 347 px + Randspalte M → Kopfleiste endet wie überall bei 1 368 px
    _ungroup_cols(ws)
    for letter, w in (("B", 44), ("C", 23), ("D", 23), ("E", 12.5), ("F", 8)):
        _width(ws, letter, w)
    for ci in range(7, 12):
        _width(ws, get_column_letter(ci), 12.9)
    _width(ws, "L", 13.3)

    # ---- Seitenkopf (P1-18)
    _header(ws, "Eingaben  ›  Alle Annahmen", "Eingaben",
            "Alle Annahmen auf einen Blick – Werte stammen aus dem Leitfaden und werden dort geändert; "
            "direkt editierbar sind nur die Profi-Felder für Darlehen II.", L_,
            meta='=IFERROR(IF(Erstellt_fuer="","","Erstellt für "&Erstellt_fuer),"")')
    # Legende rechts neben der Sprungleiste (Zeile 8, Formen von navigation.py)
    _clear_row(ws, 8, 2, L_)
    C.safe_merge(ws, "H", 8, "L", 8)
    lg = ws["H8"]
    lg.value = C.rich([("■  ", T_BODY, False, INPUT_LINE), ("Eingabe hier", T_SMALL, True, BLUE),
                       ("        ■  ", T_BODY, False, BLUE), ("aus dem Leitfaden – dort ändern", T_SMALL, False, BLUE),
                       ("        ■  ", T_BODY, False, INK), ("berechnet", T_SMALL, False, INK)])
    lg.font = font(T_SMALL, False, MUTED)
    lg.alignment = align("right", "center", 1)
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
        C.safe_merge(ws, "C", head, "D", head)
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

    # ---- Summenstufen (P1-14): Zwischensumme F3F7FC · Blockergebnis E7EEF7 mit Doppellinie
    for r in EIN_SUB + EIN_RESULT:
        C.sum_row(ws, r, "B", "L", "sub" if r in EIN_SUB else "result")
        ws.cell(r, 2).alignment = align("left", "center", 1, wrap=True)
        for cc in (5, 7):
            ws.cell(r, cc).font = font(T_SMALL, False, MUTED)

    # ---- Profi-Felder (Darlehen II): Eingabezellen über C:D
    for r in EIN_INPUTS:
        for c in (ws.cell(r, 3), ws.cell(r, 4)):
            C.input_style(c, "required")
        ws.cell(r, 3).alignment = align("right", "center", 1)

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
    # Einheit genau einmal
    if r in EIN_UNIT:
        e.value = EIN_UNIT[r]
    elif isinstance(e.value, str) and e.value.strip() in ("€", "%"):
        e.value = None
    # Zeilenstil: keine senkrechten Linien, Haarlinie über die volle Breite
    for c in C.iter_cells(ws, 2, r, last, r):
        c.fill = NOFILL
        c.border = Border(bottom=side("hair", LINE))
    sub = r in EIN_DAVON or r in EIN_DERIVED
    b.font = font(T_BODY, False, INK2 if sub else INK)
    b.alignment = align("left", "center", 2 if sub else 1, wrap=True)
    # Wert (C:D verbunden) – eine rechte Kante für Zahlen UND Texte
    if ws.cell(r, 4).value is None:
        C.safe_merge(ws, "C", r, "D", r)
    col_c = cval.font.color.rgb[-6:] if (cval.font.color is not None and isinstance(cval.font.color.rgb, str)) else INK
    linked = col_c.upper() == BLUE
    entry = r in EIN_INPUTS
    cval.font = font(T_BODY, False, BLUE if linked else INK)
    fmt = cval.number_format
    is_text = (fmt in (None, "General")) and not isinstance(cval.value, (int, float)) and r != 16
    cval.number_format = _ein_fmt(r, fmt, entry)
    cval.alignment = align("right", "center", 1, wrap=is_text)
    # Einheit
    e.font = font(T_SMALL, False, MUTED)
    e.alignment = align("left", "center", 1)
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
    hints = None
    if is_text and C.is_formula(cval.value):
        opts = _CHOICES.get(cval.value[1:], [])
        if opts:
            hints = {"C": max(opts, key=lambda t: C.text_width(t))}
    _row_height(ws, r, "B", "L", hints)


# ================================================================================================ Konfiguration
KON_BANDS = {  # Zeile → (Titel, Zusatz mit Rechtsgrundlage/Quelle)
    8: ("Grunderwerbsteuer nach Bundesland", "§ 11 GrEStG · Landesgesetze, Übersicht z. B. finanz-tools.de (Stand 2026)"),
    27: ("Einkommensteuertarif 2026", "Quelle: gesetze-im-internet.de (§ 32a EStG i. d. F. ab VZ 2026)"),
    43: ("Körperschaftsteuersatz nach Kalenderjahr", "§ 23 KStG · Investitionssofortprogramm 2025"),
    56: ("Konstanten und Schwellenwerte", "steuerliche Parameter"),
    74: ("Ampel-Schwellenwerte", "Cockpit, Dashboard, Leitfaden und Sensitivität"),
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
    45: "ab 2028 −1 %-Punkt p. a. bis 10 % (2032); zzgl. Soli 5,5 %",
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
    for letter, w in (("B", 43), ("C", 12), ("D", 11), ("E", 44), ("F", 2), ("G", 79)):
        _width(ws, letter, w)

    _header(ws, "Anhang  ›  Konfiguration", "Konfiguration", None, last)
    w = ws["B7"]
    w.value = C.rich([("▲  Expertenbereich", T_BODY, True, AMBER),
                      ("   ·   Werte hier zentral pflegen – Änderungen wirken auf alle Blätter; Reihenfolge der "
                       "Auswahllisten nicht ändern · Rechtsstand September 2026", T_BODY, False, MUTED)])
    w.font = font(T_BODY, False, MUTED)
    w.alignment = align("left", "top")
    ws.row_dimensions[7].height = 22

    # Quellen wandern in die Abschnittsköpfe (Zusatz), die Fußnotenzeilen werden Abstandszeilen
    for a in ("E9", "E10", "B26", "B41"):
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

    # ---- Abschnittsköpfe Ebene 1 (P1-12) und Tabellenköpfe Ebene 2
    for r, (title, extra) in KON_BANDS.items():
        _section(ws, r, "B", "E", title, extra)
    _section(ws, 8, "G", "G", "Auswahllisten", "Dropdown-Quellen · Reihenfolge nicht ändern")
    ws.row_dimensions[9].height = C.H_HEAD
    for r, labels in ((9, {"B": ("BUNDESLAND", "left"), "C": ("STEUERSATZ", "right"), "D": ("GÜLTIG SEIT", "center")}),
                      (44, {"B": ("KALENDERJAHR", "left"), "C": ("STEUERSATZ", "right"), "E": ("HINWEIS", "left")}),
                      (75, {"B": ("KENNZAHL", "left"), "C": ("GRÜN AB", "right"), "D": ("GELB AB", "right"),
                            "E": ("HINWEIS", "left")})):
        for c in C.iter_cells(ws, 2, r, 5, r):
            c.value = None
        C.section(ws, r, "B", "E", None, level=2, labels=labels)
    ws["D9"].alignment = align("center", "center")

    # ---- Abstandszeilen: je eine Zeile (18 pt, weil rechts Listeneinträge darin stehen) vor jedem Abschnitt
    for r in KON_GAPS:
        for c in C.iter_cells(ws, 2, r, 5, r):
            c.fill = NOFILL
            c.border = Border()
        ws.row_dimensions[r].height = C.H_ROW
    ws.row_dimensions[42].hidden = True  # zweite Leerzeile vor dem KSt-Abschnitt (Abstand wie überall)
    for r in (26, 27, 41):  # Listeneinträge in Sonderzeilen mittig (Befund Auswahllisten)
        ws.cell(r, 7).alignment = align("left", "center", 1)

    _footer(ws, 82, 2, last, gap=C.H_ROW)


# ================================================================================================ Hinweise
HIN_TOPICS = {29: "Zeitraster", 30: "Finanzierung", 31: "Steuersatz", 32: "AfA-Kombination", 33: "Rücklagen",
              34: "GmbH", 35: "Prognose", 36: "Haftung"}


def hinweise(wb):
    ws = wb["Hinweise"]
    last = 5
    _ungroup_cols(ws)
    for letter, w in (("B", 30), ("C", 73), ("D", 55), ("E", 33)):
        _width(ws, letter, w)
    _header(ws, "Anhang  ›  Hinweise", "Hinweise",
            "Steuerliche Regelungen, Rechtsgrundlagen und Modellannahmen · Rechtsstand September 2026 "
            "(Wachstumsbooster / Investitionssofortprogramm 2025, JStG 2024, EStG i. d. F. 2026) · "
            "keine Steuerberatung im Einzelfall", last)

    _section(ws, 8, "B", "E", "Steuerliche Regelungen und ihre Umsetzung im Tool", "17 Themen · Kurzfassung")
    _clear_row(ws, 9, 2, last)
    C.section(ws, 9, "B", "E", None, level=2,
              labels={"B": ("THEMA", "left"), "C": ("REGELUNG (KURZFASSUNG)", "left"),
                      "D": ("UMSETZUNG IM TOOL", "left"), "E": ("FUNDSTELLE", "left")})
    spec = {2: (T_BODY, True, NAVY), 3: (T_BODY, False, INK), 4: (T_BODY, False, INK), 5: (T_SMALL, False, MUTED)}
    for r in range(10, 27):
        n = 1
        for cc, (sz, bd, colr) in spec.items():
            c = ws.cell(r, cc)
            c.fill = NOFILL
            c.border = Border(bottom=side("hair", LINE2))
            if isinstance(c.value, str) and not C.is_formula(c.value):
                C.set_text(c, _typo(c.value))
            c.font = font(sz, bd, colr)
            c.alignment = align("left", "top", 1, wrap=True)
            n = max(n, C.lines_needed_metric(_plain(c.value), C.col_px(ws, cc), sz, bd, 1))
        ws.row_dimensions[r].height = n * 13 + 8
    _clear_row(ws, 27, 2, last, values=False)
    ws.row_dimensions[27].height = C.H_GAP

    _section(ws, 28, "B", "E", "Modellannahmen und Vereinfachungen", "Grenzen des Modells")
    width = C.span_px(ws, 3, 4)
    for r in range(29, 37):
        text = re.sub(r"^\s*[•●]\s*", "", _plain(ws.cell(r, 2).value))
        _clear_row(ws, r, 2, last)
        C.set_text(ws.cell(r, 2), HIN_TOPICS.get(r, ""))
        ws.cell(r, 2).font = font(T_BODY, True, NAVY)
        ws.cell(r, 2).alignment = align("left", "top", 1, wrap=True)
        C.safe_merge(ws, "C", r, "D", r)
        t = ws.cell(r, 3)
        C.set_text(t, _typo(text))
        t.font = font(T_BODY, False, INK)
        t.alignment = align("left", "top", 1, wrap=True)
        for c in C.iter_cells(ws, 2, r, last, r):
            c.border = Border(bottom=side("hair", LINE2))
        n = C.lines_needed_metric(t.value, width, T_BODY, False, 1)
        ws.row_dimensions[r].height = n * 13 + 8
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
