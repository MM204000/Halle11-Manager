"""Formularblätter Eingaben, Konfiguration, Hinweise (Agent E) – ruhiger „Private Banking“-Stil.

Eingaben (P2-19, P1-03, P1-04, P2-03, P2-04, P2-06, P1-06, P1-17):
  Spalten B 54 | C:D 18+18 | E 13 | F 8 (QUELLE-Links „S01 ›“) | G:L 6 × 12,1 (Hinweis); Formular-Bänder mit
  kurzem Titel + Zusatz und „↑ Übersicht“; ein Tabellenstil über B:L; Einheit genau einmal; Zahlenformate mit
  Nullabschnitt; Zeilenraster 18/30/42 pt; Legende rechts neben der Sprungleiste (Formen: navigation.py).
Konfiguration (P2-20, P1-05, P1-06, P2-03, P2-06, P3-03): Warnleiste, Spaltenköpfe im Band, Quellen als Fußnote,
  Panel G 82 breit, einheitliche Formate/Ausrichtung, Fuß.
Hinweise (P2-21, P3-03): Spalten mit Einzug statt Linien, Zeilenhöhe nach Inhalt, Aufzählungspunkt in Akzentfarbe.

Nur Darstellung: Formeln und Eingabewerte bleiben unverändert; verändert werden nur Stile, Zahlenformate,
statische Beschriftungen (auf die nichts verweist), Verbünde, Spaltenbreiten und Zeilenhöhen.
"""
import re
from copy import copy

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Border
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.hyperlink import Hyperlink

import core as K
from core import (ACCENT, AMBER, BLUE, INK, INK2, INPUT_BG, INPUT_LINE, LINE, LINE2, MUTED, NAVY, NOFILL, NUMFMT,
                  SANS, SKY, T_BODY, T_MICRO, T_SMALL, align, fill, font, side)

ROW1, ROW2, ROW3 = K.H_ROW, K.H_ROW2, K.H_ROW3        # 18 / 30 / 42 pt
_CHOICES, _CACHED = {}, {}
PCT0 = '0 %;-0 %;"–"'


# ================================================================================================ Hilfen
def _ungroup_cols(ws):
    """Gruppierte Spaltenbreiten (<col min max>) in Einzelspalten auflösen."""
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
    ci = K.col(letter)
    d.min, d.max = ci, ci
    d.width = w


def _rich(parts):
    """parts: [(text, size, bold, color)] → CellRichText (Calibri)."""
    return CellRichText([TextBlock(InlineFont(rFont=SANS, sz=s, b=b, color=c), t) for t, s, b, c in parts])


def _unmerge_row(ws, row, c1, c2):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= row <= mr.max_row and not (mr.max_col < K.col(c1) or mr.min_col > K.col(c2)):
            ws.unmerge_cells(str(mr))


class _Metric:
    """Calibri-Laufweite über Carlito (metrisch identisch), sonst Näherung core.text_px."""

    def __init__(self):
        import os
        self.paths, self.cache = {}, {}
        for bold, fn in ((False, "Carlito-Regular.ttf"), (True, "Carlito-Bold.ttf")):
            for d in ("/usr/share/fonts/truetype/crosextra", "/usr/share/fonts/truetype/carlito", "/usr/share/fonts"):
                for root, _, files in os.walk(d):
                    if fn in files:
                        self.paths[bold] = os.path.join(root, fn)
                        break
                if bold in self.paths:
                    break
        try:
            from PIL import ImageFont  # noqa: F401
            self.ok = False in self.paths
        except Exception:
            self.ok = False

    def width(self, text, size, bold=False):
        if not text:
            return 0.0
        if not self.ok:
            return K.text_px(text, size, bold)
        key = (size, bold)
        if key not in self.cache:
            from PIL import ImageFont
            self.cache[key] = ImageFont.truetype(self.paths.get(bold, self.paths[False]), int(round(size * 96 / 72 * 4)))
        return self.cache[key].getlength(str(text)) / 4.0


_M = _Metric()


def _lines(text, width_px, size=T_BODY, bold=False, indent=1):
    """Zeilenzahl beim Umbruch wie Excel: Umbruch an Leerzeichen und nach „-“ / „/“;
    Einzug ≈ 9 px je Stufe, 5 px Innenabstand, 2 % Reserve."""
    if text is None or text == "":
        return 1
    avail = (width_px - 9 * indent - 5) * 0.98
    n = 0
    for para in str(text).split("\n"):
        tokens = re.findall(r"[^ \-/]*[\-/]|[^ \-/]+ ?| ", para)
        cur = ""
        n += 1
        for t in tokens:
            cand = cur + t
            if _M.width(cand.rstrip(), size, bold) <= avail or not cur.strip():
                cur = cand
            else:
                n += 1
                cur = t
    return n


def _grid_height(n):
    return {1: ROW1, 2: ROW2}.get(n, ROW3 + (n - 3) * 12)


def _fmt(f, pct=None):
    """Zahlenformat der Vorlage → NUMFMT (P2-06). pct: erzwungenes Prozentformat."""
    if f is None or f == "General":
        return f
    if "%" in f:
        if pct:
            return pct
        if "0.00" in f:
            return NUMFMT["pct2"]
        if "0.0" in f:
            return NUMFMT["pct1"]
        return PCT0
    if "€" in f:
        return NUMFMT["eur2"] if ".00" in f else NUMFMT["eur"]
    return f


def _plain(v):
    return str(v) if v is not None else ""


def _text_link(cell, text, location, tooltip, size=T_SMALL, color=BLUE, bold=True):
    K.set_text(cell, text)
    cell.font = font(size, bold, color)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=location, display=text, tooltip=tooltip)


def _band_title(title, extra=None):
    """Formular-Band: Titel 10 pt fett weiß, Zusatz 8,5 pt Sky direkt dahinter (eine Zeile)."""
    if not extra:
        return title
    return _rich([(title, T_BODY, True, "FFFFFF"), ("     " + extra, 8.5, False, SKY)])


def _header_rows(ws, eyebrow, title, subtitle, last):
    """Seitenkopf Z. 4–7 (P2-05) – Flächen/Linien rechts davon leeren."""
    for r in (4, 5, 6, 7):
        for c in K.iter_cells(ws, 2, r, last, r):
            c.value = None if c.coordinate not in ("B5", "B6", "B7") else c.value
            c.fill = NOFILL
            c.border = Border()
            c.hyperlink = None
    _unmerge_row(ws, 5, 2, last)
    _unmerge_row(ws, 6, 2, last)
    _unmerge_row(ws, 7, 2, last)
    K.page_header(ws, "B", last, eyebrow, title, subtitle)
    ws.row_dimensions[4].height = 10


def _footer(ws, row, c1, c2):
    _unmerge_row(ws, row, c1, c2)
    _unmerge_row(ws, row + 1, c1, c2)
    for r in (row, row + 1):
        for c in K.iter_cells(ws, c1, r, c2, r):
            c.value = None
            c.fill = NOFILL
            c.border = Border()
    K.footer(ws, row, c1, c2)
    ws.row_dimensions[row - 1].height = 14


# ================================================================================================ Eingaben
EIN_BANDS = {
    9: ("1  OBJEKT UND KAUFPREIS", None),
    24: ("2  KAUFNEBENKOSTEN", "Anschaffungsnebenkosten – werden anteilig mit dem Gebäude abgeschrieben"),
    39: ("3  FINANZIERUNGSNEBENKOSTEN", "sofort abziehbare Werbungskosten bzw. Betriebsausgaben"),
    48: ("4  KAUFPREISAUFTEILUNG", "Grund und Boden / Gebäude · AfA-Bemessungsgrundlage"),
    61: ("5  MASSNAHMEN NACH ERWERB", "Renovierung, Sonderumlagen, Reserven"),
    75: ("6  MIETE UND BEWIRTSCHAFTUNG", "Monatswerte, Stand Kauf"),
    95: ("7  FINANZIERUNG", "Darlehen II nur hier eingeben"),
    119: ("8  STEUERLICHE OPTIONEN", "Rechtsstand September 2026"),
    141: ("9  VERKAUFSSZENARIO", "Exit"),
}
EIN_LAST = 145
EIN_SUMS = (36, 46, 56, 58, 71, 92, 113, 114, 138)
EIN_DAVON = (20, 22, 80, 81)
STEP_PREFIX = re.compile(r"^Eingabe im Leitfaden, Schritt (\d\d)\s*(?:–\s*)?")

# Beschriftung kürzen (P1-04): Zeile → (neue Beschriftung, Zusatz vor dem Hinweis oder None)
EIN_LABELS = {
    15: ("Grundstücksfläche bzw. Miteigentumsanteil", None),
    20: ("davon bewegliche Gegenstände (gesondert ausgewiesen)", "Einbauküche, Möbel"),
    26: ("Grunderwerbsteuersatz abweichend", "leer = automatisch nach Bundesland"),
    32: ("Sonstige Erwerbsnebenkosten", "Gutachten, Fahrtkosten, Übersetzungen"),
    54: ("Kaufpreis Immobilie (ohne Inventar und Rücklage)", None),
    57: ("Anschaffungskosten Grund und Boden", "inkl. anteiliger Kaufnebenkosten; nicht abschreibbar"),
    58: ("Anschaffungskosten Gebäude = AfA-Basis", "inkl. anteiliger Kaufnebenkosten"),
    59: ("Anschaffungskosten gesamt", "Kaufpreis + Kaufnebenkosten"),
    70: ("Liquiditätsreserve einmalig (nicht steuerrelevant)", "z. B. Mieterwechsel"),
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
    130: ("Degressiv: automatischer Wechsel zur linearen AfA", "sobald vorteilhaft"),
    131: ("Sonder-AfA § 7b EStG (Mietwohnungsneubau)", None),
    132: ("Erhöhte Absetzungen § 7h / § 7i EStG", "Sanierungsgebiet / Baudenkmal"),
    133: ("Begünstigte Sanierungskosten § 7h / § 7i", "im Kaufpreis enthalten"),
    138: ("Grenzbelastung Privat (ESt + Soli + KiSt)", None),
    145: ("Verkaufspreis manuell", "leer = Wertentwicklung lt. Prognose"),
}
# Einheit genau einmal (P1-03): nur der Bezug bleibt stehen
UNIT_MAP = {"€": None, "%": None, "€/Monat": "pro Monat", "€ p. a.": "pro Jahr", "€/m²": "je m²",
            "€/m² p. a.": "je m² p. a.", "% Darlehen": "vom Darlehen", "% der Miete": "der Miete",
            "% Verkaufspreis": "vom Verkaufspreis", "fach": "Jahresmieten"}


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
    return t if "!" in t and ":" not in t else None


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
        return [str(c.value) for c in K.iter_cells(ws, m.group(2), int(m.group(3)), m.group(4), int(m.group(5)))
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
                for c in K.iter_cells(ws, cr.min_col, cr.min_row, cr.max_col, cr.max_row):
                    by_cell[f"{ws.title}!{c.coordinate}"] = vals
    out = {}
    for n, dn in wb.defined_names.items():
        t = str(dn.attr_text).replace("$", "").replace("'", "")
        if t in by_cell:
            out[n] = by_cell[t]
    return out


def _cached_values():
    """Angezeigte Werte der Vorlage (für Zeilenhöhen verknüpfter Textfelder)."""
    import os
    from openpyxl import load_workbook
    path = os.path.join(os.path.dirname(os.path.abspath(K.__file__)), "quelle", "Immobilien-Kalkulationstool_Pro_1.xlsx")
    try:
        ws = load_workbook(path, data_only=True)["Eingaben"]
        return {r: ws.cell(r, 3).value for r in range(11, EIN_LAST + 1)}
    except Exception:
        return {}


def eingaben(wb):
    ws = wb["Eingaben"]
    L_ = 12  # letzte Tabellenspalte L
    steps = _step_sheets(wb)
    global _CHOICES, _CACHED
    _CHOICES, _CACHED = _choices(wb), _cached_values()

    # ---- Spalten (Gesamtbreite ≈ unverändert)
    _ungroup_cols(ws)
    for letter, w in (("B", 54), ("C", 18), ("D", 18), ("E", 15), ("F", 8)):
        _width(ws, letter, w)
    for ci in range(7, 13):
        _width(ws, get_column_letter(ci), 11.8)

    # ---- Seitenkopf
    _header_rows(ws, "Alle Annahmen im Überblick", "Eingaben",
                 "Werte stammen aus dem Leitfaden und werden dort geändert – direkt editierbar sind nur die "
                 "Profi-Felder für Darlehen II.", L_)
    # Legende rechts neben der Sprungleiste (Zeile 8, Formen von navigation.py)
    for c in K.iter_cells(ws, 2, 8, L_, 8):
        c.value = None
        c.fill = NOFILL
        c.border = Border()
    K.safe_merge(ws, "H", 8, "L", 8)
    lg = ws["H8"]
    lg.value = _rich([("■ ", 10, False, INPUT_LINE), ("Eingabe hier", T_SMALL, True, BLUE),
                      ("      ■ ", 10, False, BLUE), ("aus Leitfaden – dort ändern", T_SMALL, False, BLUE),
                      ("      ■ ", 10, False, INK), ("berechnet", T_SMALL, False, INK)])
    lg.font = font(T_SMALL, False, MUTED)
    lg.alignment = align("right", "center", 1)
    ws.row_dimensions[8].height = 30

    # ---- Abschnitte
    band_rows = sorted(EIN_BANDS)
    for i, br in enumerate(band_rows):
        title, extra = EIN_BANDS[br]
        _unmerge_row(ws, br, 2, L_)
        for c in K.iter_cells(ws, 2, br, L_, br):
            c.value = None
        K.band_form(ws, br, 2, L_, None)
        ws.cell(br, 2).value = _band_title(title, extra)
        up = ws.cell(br, L_)
        _text_link(up, "↑ Übersicht", f"'{ws.title}'!{K.link_target(ws.title)}", "Zum Seitenanfang (Sprungleiste)",
                   size=T_MICRO, color=SKY, bold=False)
        up.alignment = align("right", "center", 1)
        ws.row_dimensions[br].height = 22
        head = br + 1
        end = (band_rows[i + 1] - 2) if i + 1 < len(band_rows) else EIN_LAST
        _table_head(ws, head, L_)
        for r in range(head + 1, end + 1):
            _data_row(ws, wb, steps, r, L_)
        # Leerzeile vor dem nächsten Band
        gap = end + 1
        for c in K.iter_cells(ws, 2, gap, L_, gap):
            c.fill = NOFILL
            c.border = Border()
        ws.row_dimensions[gap].height = 14 if gap != EIN_LAST + 1 else 14

    # ---- Summenzeilen über die volle Breite
    K.set_text(ws["E93"], "Jahresmieten")  # 25,0× Jahresmieten (P1-03)
    ws["E93"].font = font(T_SMALL, False, MUTED)
    ws["E93"].alignment = align("left", "center", 1)
    ws["C93"].number_format = NUMFMT["mult1"]
    K.total(ws, 112, 2, L_, level=1)
    ws["G112"].font = font(T_SMALL, False, MUTED)
    for r in EIN_SUMS:
        K.total(ws, r, 2, L_, level=2)
        for cc in (5,):
            ws.cell(r, cc).font = font(T_SMALL, False, MUTED)
        g = ws.cell(r, 7)
        if g.value is not None:
            g.font = font(T_SMALL, False, MUTED)

    # ---- Profi-Felder (Darlehen II): Eingabezellen über C:D
    for r in range(104, 112):
        for c in (ws.cell(r, 3), ws.cell(r, 4)):
            K.input_style(c, "required")
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


def _table_head(ws, row, last):
    _unmerge_row(ws, row, 2, last)
    for c in K.iter_cells(ws, 2, row, last, row):
        c.value = None
    K.table_head(ws, row, 2, last, {"B": ("POSITION", "left"), "C": ("WERT", "right"), "E": ("EINHEIT", "left"),
                                    "F": ("QUELLE", "left"), "G": ("HINWEIS / STEUERLICHE EINORDNUNG", "left")})
    K.safe_merge(ws, "C", row, "D", row)
    K.safe_merge(ws, "G", row, "L", row)
    ws.row_dimensions[row].height = K.H_HEAD


def _data_row(ws, wb, steps, r, last):
    b, cval, e, f = ws.cell(r, 2), ws.cell(r, 3), ws.cell(r, 5), ws.cell(r, 6)
    if b.value is None and cval.value is None:
        return
    # Hinweis aus F:L lösen → Quelle (F) + Hinweis (G:L)
    hint = f.value
    _unmerge_row(ws, r, 6, last)
    for c in K.iter_cells(ws, 6, r, last, r):
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
        K.set_text(b, new)
        if extra:
            hint = f"{extra} – {hint}" if hint else extra
    if hint:
        hint = hint.replace(" | ", " · ")
        hint = hint[0].upper() + hint[1:]
    # Einheit
    if isinstance(e.value, str):
        u = e.value.strip()
        if u in UNIT_MAP:
            e.value = UNIT_MAP[u]
    # Stile der Zeile (Tabellenstil: keine senkrechten Linien, Haarlinie über die volle Breite)
    for c in K.iter_cells(ws, 2, r, last, r):
        c.fill = NOFILL
        c.border = Border(bottom=side("hair", LINE))
    b.font = font(T_BODY, False, INK)
    b.alignment = align("left", "center", 2 if r in EIN_DAVON else 1, wrap=True)
    if r in EIN_DAVON:
        b.font = font(T_BODY, False, INK2)
    # Wert (C:D verbunden)
    if ws.cell(r, 4).value is None:
        K.safe_merge(ws, "C", r, "D", r)
    col_c = cval.font.color.rgb[-6:] if (cval.font.color is not None and isinstance(cval.font.color.rgb, str)) else INK
    linked = col_c.upper() == BLUE
    cval.font = font(T_BODY, False, BLUE if linked else INK)
    fmt = cval.number_format
    is_text = fmt == "General" and not (isinstance(cval.value, (int, float)))
    if r == 93:
        cval.number_format = NUMFMT["mult1"]
    else:
        cval.number_format = _fmt(fmt)
    cval.alignment = align("left", "center", 1, wrap=True) if is_text else align("right", "center", 1)
    # Einheit
    e.font = font(T_SMALL, False, MUTED)
    e.alignment = align("left", "center", 1)
    # Quelle-Link
    if link and link in steps:
        sheet = steps[link]
        target = _name_target(wb, cval.value) or f"'{sheet}'!{K.link_target(sheet)}"
        if "'" not in target.split("!")[0]:
            sh, addr = target.split("!")
            target = f"'{sh}'!{addr}"
        step_title = sheet[4:]
        _text_link(f, f"S{link} ›", target, f"Im Leitfaden ändern: Schritt {int(link)} · {step_title}")
        f.alignment = align("left", "center", 1)
    # Hinweis G:L
    g = ws.cell(r, 7)
    if hint:
        K.set_text(g, hint)
    g.font = font(T_SMALL, False, MUTED)
    g.alignment = align("left", "center", 1, wrap=True)
    K.safe_merge(ws, "G", r, "L", r)
    # Zeilenhöhe nach Zeilenzahl (18 / 30 / 42 pt)
    n = max(_lines(_plain(b.value), K.col_px(ws, 2), T_BODY, False, b.alignment.indent),
            _lines(_plain(g.value), K.span_px(ws, 7, last), T_SMALL))
    if is_text:
        texts = [cval.value] if not K.is_formula(cval.value) else []
        if K.is_formula(cval.value):
            texts = _CHOICES.get(cval.value[1:], []) or [_CACHED.get(r)]
        for t in texts:
            n = max(n, _lines(_plain(t), K.span_px(ws, 3, 4), T_BODY))
    ws.row_dimensions[r].height = _grid_height(n)


# ================================================================================================ Konfiguration
def konfiguration(wb):
    ws = wb["Konfiguration"]
    last = 7  # G
    _ungroup_cols(ws)
    _width(ws, "G", 82)

    _header_rows(ws, "Anhang  ›  Konfiguration  ·  Rechtsstand September 2026", "Konfiguration", None, last)
    # Warnhinweis als Untertitel-Zeile (ersetzt den Untertitel; ohne Fläche, damit Luft zum Band bleibt)
    w = ws["B7"]
    K.set_text(w, "⚠  Expertenbereich – Werte hier zentral pflegen; Änderungen wirken auf alle Blätter. "
                  "Reihenfolge der Auswahllisten nicht ändern.")
    w.font = font(T_SMALL, True, AMBER)
    w.alignment = align("left", "top")
    ws.row_dimensions[7].height = 22
    ws.row_dimensions[4].height = 10

    # ---- Formular-Bänder (B:E) und Panel-Band G8
    bands = {8: ("GRUNDERWERBSTEUER NACH BUNDESLAND", "§ 11 GrEStG · Landessätze"),
             27: ("EINKOMMENSTEUERTARIF 2026", "§ 32a Abs. 1 EStG"),
             43: ("KÖRPERSCHAFTSTEUERSATZ NACH KALENDERJAHR", "§ 23 KStG · Investitionssofortprogramm 2025"),
             56: ("KONSTANTEN UND SCHWELLENWERTE", "steuerlich"),
             74: ("AMPEL-SCHWELLENWERTE FÜR KENNZAHLEN", "Cockpit, Dashboard, Leitfaden, Sensitivität")}
    for r, (title, extra) in bands.items():
        for cc in (2, 3, 4, 5):
            ws.cell(r, cc).value = None
        K.band_form(ws, r, 2, 5, None)
        ws.cell(r, 2).value = _band_title(title, extra)
        ws.row_dimensions[r].height = 22
    # Spaltenköpfe im Band (Tabellen ohne eigene Kopfzeile)
    for r, heads in ((27, {3: ("WERT", "right"), 5: ("FORMEL / ERLÄUTERUNG", "left")}),
                     (56, {3: ("WERT", "right"), 5: ("ERLÄUTERUNG", "left")})):
        for cc, (t, h) in heads.items():
            c = ws.cell(r, cc)
            K.set_text(c, t)
            c.font = font(T_MICRO, True, SKY)
            c.alignment = align(h, "center", 1)
    K.band_form(ws, 8, 7, 7, None)
    ws["G8"].value = _band_title("AUSWAHLLISTEN", "Dropdown-Quellen – Reihenfolge nicht ändern")
    ws.row_dimensions[8].height = 22

    # ---- Tabellenköpfe
    K.table_head(ws, 9, 2, 5, {"B": ("BUNDESLAND", "left"), "C": ("STEUERSATZ", "right"),
                               "D": ("GÜLTIG SEIT", "center")})
    K.table_head(ws, 44, 2, 5, {"B": ("KALENDERJAHR", "left"), "C": ("STEUERSATZ", "right"),
                                "E": ("HINWEIS", "left")})
    K.table_head(ws, 75, 2, 5, {"B": ("KENNZAHL", "left"), "C": ("GRÜN AB", "right"), "D": ("GELB AB", "right"),
                                "E": ("HINWEIS", "left")})

    # Quelle der GrESt-Tabelle als Fußnote (Regel: jede Tabelle hat ihre Quelle darunter)
    src = ws["E10"].value
    ws["E10"].value = None
    if src:
        K.set_text(ws["B26"], f"Quelle: {src}")
    for a in ("B26", "B41"):
        c = ws[a]
        c.font = font(T_SMALL, False, MUTED)
        c.alignment = align("left", "top", 1)
        for cc in K.iter_cells(ws, 2, c.row, 5, c.row):
            cc.fill = NOFILL
            cc.border = Border()
        ws.row_dimensions[c.row].height = 22

    ws["E9"].value = None  # Quelle steht als Fußnote unter der Tabelle
    # Kürzungen (P1-04)
    K.set_text(ws["B61"], "Sonder-AfA § 7b: Baukostenobergrenze je m²")
    K.set_text(ws["E61"], "je m² Wohnfläche · Fallbeil: bei Überschreitung entfällt § 7b vollständig")
    K.set_text(ws["B62"], "Sonder-AfA § 7b: max. Bemessungsgrundlage je m²")
    K.set_text(ws["E62"], "je m² Wohnfläche · darüber hinaus keine Sonder-AfA")
    K.set_text(ws["E38"], "5,5 % der ESt/KSt; Freigrenze 2026: 20.350 € / 40.700 € ESt")
    K.set_text(ws["E45"], "ab 2028 −1 %-Punkt p. a. bis 10 % (2032); zzgl. Soli 5,5 %")

    # ---- Datenzeilen B:E
    tables = [(10, 25), (28, 40), (45, 54), (57, 72), (76, 80)]
    pct_rows = set()
    for r1, r2 in tables:
        for r in range(r1, r2 + 1):
            for c in K.iter_cells(ws, 2, r, 5, r):
                if c.protection.locked is not False:
                    c.fill = NOFILL
                c.border = Border(bottom=side("hair", LINE))
            b = ws.cell(r, 2)
            b.font = font(T_BODY, False, INK)
            b.alignment = align("left", "center", 1, wrap=True)
            if isinstance(b.value, int):
                b.number_format = NUMFMT["year"]
            for cc in (3, 4):
                c = ws.cell(r, cc)
                if c.value is None:
                    continue
                f = c.number_format
                if "%" in f:
                    c.number_format = NUMFMT["pct2"]
                    pct_rows.add(r)
                elif "€" in f:
                    c.number_format = NUMFMT["eur"]
                elif f == "0" and isinstance(c.value, (int, float)) and c.value >= 1000 and r not in (54,):
                    c.number_format = NUMFMT["int"]
                elif f == "0.00" and r in (32, 34):
                    c.number_format = "#,##0.00"
                if cc == 4 and isinstance(c.value, str):  # Datum als Text (gültig seit)
                    c.font = font(T_SMALL, False, MUTED)
                    c.alignment = align("center", "center")
                else:
                    c.alignment = align("right", "center", 1)
                if c.protection.locked is False:
                    K.input_style(c, "required")
                    c.alignment = align("right", "center", 1)
            e = ws.cell(r, 5)
            e.font = font(T_SMALL, False, MUTED)
            e.alignment = align("left", "center", 1, wrap=True)
            n = max(_lines(_plain(b.value), K.col_px(ws, 2), T_BODY),
                    _lines(_plain(e.value), K.col_px(ws, 5), T_SMALL))
            ws.row_dimensions[r].height = _grid_height(n)
    # Leerzeilen zwischen den Tabellen
    for r in (42, 55, 73):
        ws.row_dimensions[r].height = 14
        for c in K.iter_cells(ws, 2, r, 5, r):
            c.border = Border()

    # ---- Panel G (Auswahllisten)
    heads = [9, 15, 20, 25, 30, 34, 38, 43, 51]
    for r in range(9, 54):
        g = ws.cell(r, 7)
        if r in heads:
            g.fill = fill(K.HEAD)
            g.border = Border(bottom=side("thin", ACCENT))
            g.font = font(T_MICRO, True, BLUE)
            g.alignment = align("left", "center", 1)
            if isinstance(g.value, str):
                K.set_text(g, g.value.upper().replace("AFA", "AfA"))
            if r == 43:
                g.value = "AfA-METHODE GEBÄUDE"
            if r == 25:
                g.value = "RECHTSFORM"
        elif g.value is not None:
            g.fill = NOFILL
            g.border = Border(bottom=side("hair", LINE))
            g.font = font(T_BODY, False, INK)
            g.alignment = align("left", "center", 1)
        else:
            g.fill = NOFILL
            g.border = Border()
    # Panel-Einträge in den Band-/Fußnotenzeilen: Zeilenhöhe folgt der Tabelle links (P1-05)
    for r in (26, 27, 41, 43):
        if ws.row_dimensions[r].height is None or ws.row_dimensions[r].height < ROW1:
            ws.row_dimensions[r].height = ROW1

    _footer(ws, 82, 2, last)
    ws.row_dimensions[81].height = 14


# ================================================================================================ Hinweise
def hinweise(wb):
    ws = wb["Hinweise"]
    last = 5
    _ungroup_cols(ws)
    _width(ws, "C", 66)
    _width(ws, "D", 56)
    _header_rows(ws, "Anhang  ›  Hinweise", "Hinweise",
                 "Steuerliche Regelungen, Rechtsgrundlagen und Modellannahmen · Rechtsstand September 2026 "
                 "(Wachstumsbooster / Investitionssofortprogramm 2025, JStG 2024, EStG i. d. F. 2026) · "
                 "keine Steuerberatung im Einzelfall", last)

    K.band_l1b(ws, 8, 2, last, "STEUERLICHE REGELUNGEN UND IHRE UMSETZUNG IM TOOL")
    K.table_head(ws, 9, 2, last, {"B": ("THEMA", "left"), "C": ("REGELUNG (KURZFASSUNG)", "left"),
                                  "D": ("UMSETZUNG IM TOOL", "left"), "E": ("FUNDSTELLE", "left")})
    for r in range(10, 27):
        for c in K.iter_cells(ws, 2, r, last, r):
            c.fill = NOFILL
            c.border = Border(bottom=side("thin", LINE2))
        sizes = {2: (T_BODY, True, NAVY), 3: (T_BODY, False, INK), 4: (T_BODY, False, INK), 5: (T_SMALL, False, MUTED)}
        n = 1
        for cc, (sz, bd, colr) in sizes.items():
            c = ws.cell(r, cc)
            c.font = font(sz, bd, colr)
            c.alignment = align("left", "top", 1, wrap=True)
            n = max(n, _lines(_plain(c.value), K.col_px(ws, cc), sz, bd))
        ws.row_dimensions[r].height = n * 13 + 10
    ws.row_dimensions[27].height = 14
    for c in K.iter_cells(ws, 2, 27, last, 27):
        c.border = Border()

    K.band_l1b(ws, 28, 2, last, "MODELLANNAHMEN UND VEREINFACHUNGEN")
    width = K.span_px(ws, 2, last)
    for r in range(29, 37):
        c = ws.cell(r, 2)
        text = _plain(c.value)
        text = re.sub(r"^\s*•\s*", "", text)
        c.value = _rich([("•  ", T_BODY, True, ACCENT), (text, T_BODY, False, INK2)])
        c.font = font(T_BODY, False, INK2)
        c.alignment = align("left", "center", 1, wrap=True)
        for cc in K.iter_cells(ws, 2, r, last, r):
            cc.fill = NOFILL
            cc.border = Border(bottom=side("hair", LINE))
        n = _lines("•  " + text, width, T_BODY)
        ws.row_dimensions[r].height = 20 if n == 1 else 32 if n == 2 else 44
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
