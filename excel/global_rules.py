"""Mappenweite Regeln (Typo-Skala, Ausrichtung, Tabellenstil, Zahlenformate, Farbsemantik, Druck).

early(wb):      vor den Blatt-Layouts – Standards für ALLE Blätter (Blatt-Module dürfen sie danach überschreiben):
                Korrektorat, Zahlenformate, Typo-Skala, keine senkrechten Linien, Seitenkopf, Tabellenstil,
                Zeilenraster, Fuß, Farbsemantik, Blattschutz der Bank-Vorlagen.
final(wb):      nach den Blatt-Layouts – nur Sicherheits-Normalisierungen (kein „An Zellgröße anpassen“,
                Einzug nur mit links/rechts, Mindestschrift 8 pt, Sekundärtext ≤ 9 pt nie in 8A9099,
                nur Calibri, Normzitate in Versaltexten, dxf ohne Schriftnamen).
page_setup(wb): ganz am Ende – A4, Ränder, Kopf-/Fußzeile, Druckbereiche, Maßstab, Umbrüche, Drucktitel,
                Blattreihenfolge, Registerfarben, Ansicht, Objektschutz.

Nur Darstellung: keine Formel der Vorlage, kein Eingabewert und kein Name wird verändert.
"""
import re
from copy import copy

from openpyxl.styles import Alignment, Border, Font, Protection
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak
from openpyxl.worksheet.properties import PageSetupProperties

import core as C

# =============================================================================== Blattgruppen
STEP_RE = re.compile(r"S\d\d ")
FORM_SHEETS = ("Eingaben", "Konfiguration")
BANK_SHEETS = ("Bankgespräch", "Haushaltsrechnung", "Vermögensaufstellung")
CALC_SHEETS = ("Projektion", "Steuern", "Finanzierung", "AfA-Vergleich", "Sensitivität")
PRESENTATION = ("Start", "Leitfaden", "Dashboard", "Cockpit") + BANK_SHEETS   # + S01–S12 (ohne Zeilen-/Spaltenköpfe)

ORDER = ["Start", "Leitfaden"] + [None] * 12 + [
    "Dashboard", "Cockpit", "Diagramme", "Eingaben", "Projektion", "Steuern", "AfA-Vergleich", "Finanzierung",
    "Sensitivität", "Bankgespräch", "Haushaltsrechnung", "Vermögensaufstellung", "Hinweise", "Konfiguration"]

TAB_COLOR = {"Start": C.NAVY, "Leitfaden": C.NAVY, "Dashboard": C.NAVY,
             "Cockpit": C.ACCENT, "Diagramme": C.ACCENT, "Eingaben": C.ACCENT, "Projektion": C.ACCENT,
             "Steuern": C.ACCENT, "AfA-Vergleich": C.ACCENT, "Finanzierung": C.ACCENT, "Sensitivität": C.ACCENT,
             "Bankgespräch": "8FA9C9", "Haushaltsrechnung": "8FA9C9", "Vermögensaufstellung": "8FA9C9",
             "Hinweise": "A0A7B1", "Konfiguration": "A0A7B1"}


def tab_color(name):
    return C.BLUE if STEP_RE.match(name) else TAB_COLOR.get(name, C.ACCENT)


# =============================================================================== Seitenkopf-Standard (P2-05)
# Blatt: (Titelspalte, letzte Inhaltsspalte, Eyebrow, H1, Untertitel). Die Blatt-Module überschreiben ihn bei Bedarf.
HEADERS = {
    "Eingaben": ("B", "L", "Eingaben", "Eingaben",
                 "Vollständige Übersicht aller Eingaben und Profi-Felder  ·  Gelb hinterlegt = Eingabe, "
                 "blau = aus dem Leitfaden übernommen (dort ändern)"),
    "Sensitivität": ("C", "M", ("Berechnung", "Sensitivität"), "Sensitivität",
                     "Was passiert, wenn Miete, Zins oder Wertentwicklung anders laufen?  ·  Jahr-1-Größen als "
                     "lineare Näherung, die IRR-Matrix rechnet die volle Zahlungsreihe neu"),
    "Haushaltsrechnung": ("B", "E", ("Bank", "Haushaltsrechnung"), "Haushaltsrechnung",
                          "Selbstauskunft für die Bank  ·  Monatswerte in die gelb hinterlegten Felder, "
                          "Jahreswerte rechnen sich selbst"),
    "Vermögensaufstellung": ("B", "E", ("Bank", "Vermögensaufstellung"), "Vermögensaufstellung",
                             "Selbstauskunft für die Bank  ·  aktuelle Werte eingeben und belegen "
                             "(Gutachten, Depot- oder Kontoauszug)"),
    "Hinweise": ("B", "E", ("Anhang", "Hinweise"), "Hinweise",
                 "Steuerliche Regelungen, Rechtsgrundlagen und Modellannahmen  ·  Rechtsstand September 2026 "
                 "(Investitionssofortprogramm 2025, JStG 2024, EStG i. d. F. 2026)  ·  keine Steuerberatung im Einzelfall"),
    "Konfiguration": ("B", "G", ("Anhang", "Konfiguration"), "Konfiguration",
                      "Stammdaten und Rechtsstand September 2026  ·  Werte hier zentral pflegen, "
                      "alle Berechnungen greifen auf diese Zellen zu"),
}

# Tabellen (volle Breite für Kopf-, Summen- und Zeilenlinien, P2-03) und Inhaltsbreite für den Fuß
TABLES = {
    "Eingaben": [("B", "L")],
    "Haushaltsrechnung": [("B", "E")],
    "Vermögensaufstellung": [("B", "E")],
    "Hinweise": [("B", "E")],
    "Konfiguration": [("B", "E"), ("G", "G")],
}
CONTENT = {"Eingaben": ("B", "L"), "Haushaltsrechnung": ("B", "E"), "Vermögensaufstellung": ("B", "E"),
           "Hinweise": ("B", "E"), "Konfiguration": ("B", "G"), "Sensitivität": ("C", "M"),
           "Bankgespräch": ("B", "I")}

# =============================================================================== Zahlenformate (P2-12, Katalog core.NUMFMT)
NUMFMT_MAP = {
    '#,##0" €"': C.NUMFMT["eur"],
    '#,##0" €";\\-#,##0" €";\\–': C.NUMFMT["eur"],
    '#,##0" €";-#,##0" €";"–"': C.NUMFMT["eur"],
    '#,##0.00" €";\\-#,##0.00" €";\\–': C.NUMFMT["eur2"],
    '#,##0;\\-#,##0;\\–': C.NUMFMT["num"],
    "0.00%": C.NUMFMT["pct2"], "0.00 %": C.NUMFMT["pct2"],
    "0.0%": C.NUMFMT["pct1"], "0.0 %": C.NUMFMT["pct1"],
    "0.00\\x": C.NUMFMT["mult2"], '0.00"×"': C.NUMFMT["mult2"],
    '0.0"-fach"': C.NUMFMT["mult1"],
    "dd\\.mm\\.yyyy": C.NUMFMT["date"], "dd.mm.yyyy": C.NUMFMT["date"],
}
# Prozent-Regel (P2-12): Eingaben und übernommene Eingaben 0,0 % – zweite Stelle nur, wenn der Wert sie braucht
# (Makler 3,57 %); berechnete Quoten/Renditen immer 0,0 %. Fachlich zweistellig: Zins-/Tilgungssätze (Bankpraxis)
# und die angewendeten Steuersätze (46,26 %).
# Runde 3 (P24 „eine Größe, eine Genauigkeit“): Zins-, Tilgungs-, Steuer- und AfA-Sätze 0,00 %; Kosten-/GrESt-Sätze und
# Renditen 0,0 %. Gemischte Eingabespalten (S03 Notar/Grundbuch/Makler, S07 Finanzierung) einheitlich 0,00 % – und
# ihre Spiegel auf „Eingaben“ genauso.
PCT2_WORDS = ("zins", "tilgung")
PCT2_CELLS = {"S09 Steuern": ["I11:I14"], "Cockpit": ["G36"], "AfA-Vergleich": ["C12"],
              "Steuern": ["C18:C20", "C25:C26", "D71:AQ71", "D72:AQ72", "C87"],
              "Eingaben": ["C29:C31", "C41", "C43", "C124", "C136:C139"],
              "Konfiguration": ["C36:C40", "C45:C53"],   # ESt/Soli/KiSt/KSt wie auf „Steuern“ (5,50 %, 15,00 %)
              "S01 Objekt": ["I12"], "S03 Kaufnebenkosten": ["D12:D15"], "S07 Finanzierung": ["D13:D21"]}
# Ganzzahlige Prozente (Wunsch B: S05 „Wertsteigernder Anteil“ 100 %)
PCT0_CELLS = {"S05 Maßnahmen & Reserve": ["D18"], "Eingaben": ["C69"]}
# DSCR überall mit Faktorzeichen (0,00×); Dashboard F22 (Ziel „≥ 1,20×“) behält sein eigenes Format (Wunsch G)
DSCR_CELLS = {"Start": ["E31"], "Leitfaden": ["F13"], "S08 Zwischenergebnis": ["D23"], "Dashboard": ["N15", "E23"],
              "Cockpit": ["K24"], "Bankgespräch": ["C35"], "Sensitivität": ["D35:J41"], "Konfiguration": ["C78:D78"]}
# Formeln/Koeffizienten mit zwei Stellen einheitlich (Konfiguration C32:C35)
NUM2_CELLS = {"Konfiguration": ["C32:C35"]}
# P2-07: €/m² überall im Katalogformat EUR_M2 („12,50 €/m²“), Kinderzahl mit Einheit, Nullabschnitt „–“ für
# nicht zutreffende (inaktive) Eingabefelder, „0 €“ für Summen direkt unter Eingabespalten (gleiche Nulldarstellung
# wie die Eingaben darüber)
EUR_M2_CELLS = {"S02 Kaufpreis & Miete": ["I13"], "Eingaben": ["C91"], "Bankgespräch": ["I21"]}
CHILDREN_CELLS = {"Haushaltsrechnung": ["C12"]}
ZERO_DASH_INPUTS = {"S10 Abschreibung": {"D13": "years_calc", "D17": "eur"}}
ZERO_AS_INPUT = {"Vermögensaufstellung": ["C19", "C29"]}
DASH_TEXT_SUMS = {"Vermögensaufstellung": ["D29"]}   # fester Text „–“ in einer Summenzeile: wie ein Summenwert
LABEL_FIX = {"Eigenkapital-Multiple": "Eigenkapital-Multiplikator", "EK-Multiple": "Eigenkapital-Multiplikator"}
# Jahre ganzzahlig mit Einheit im Format – das Label verliert „(Jahre)“
YEARS_LABEL = re.compile(r"\s*\(Jahre\)\s*$")

# =============================================================================== Typo-Skala (P3-11)
# Erlaubt sind nur 8 · 9 · 10 · 12,5 · 16 · 20 · 22 · 30 pt (core.TYPE_SCALE); Zwischengrößen über core.snap_size.
SCALE = C.TYPE_SCALE
# Monochrome Kennzeichen statt Emoji (P2-14): ● ▲ ■ › ⓘ; farbige Emoji und die Emoji-Variante (U+FE0F) entfallen
SYMBOLS = set("✓○●▾›‹→↓↑▲■▸ⓘ•✗×–")
EMOJI_MAP = {"🟢": "●", "🟡": "●", "🔴": "●", "🟠": "●", "⚪": "○", "✅": "✓", "❌": "✗", "ℹ️": "ⓘ", "\ufe0f": "",
             "⚠": "▲", "ℹ": "ⓘ"}   # nur statische Texte – Formeltexte (LEFT(…)="⚠") bleiben unverändert

VERSAL_FIX = [(re.compile(r"§ 32A\b"), "§ 32a"), (re.compile(r"\bESTG\b"), "EStG"), (re.compile(r"\bGRESTG\b"), "GrEStG"),
              (re.compile(r"\bKSTG\b"), "KStG"), (re.compile(r"\bGEWSTG\b"), "GewStG"), (re.compile(r"\bESTDV\b"), "EStDV"),
              (re.compile(r"\bSOLZG\b"), "SolZG"), (re.compile(r"\bUSTG\b"), "UStG"), (re.compile(r"\bAFA\b"), "AfA"),
              (re.compile(r"\(AFA\)"), "(AfA)"), (re.compile(r"\bABS\. "), "Abs. "), (re.compile(r"\bNR\. "), "Nr. "),
              (re.compile(r"\bI\. D\. F\."), "i. d. F.")]

OLD_FONTS = {None, "Aptos", "Aptos Display", "Aptos Narrow", "Inter", "Fraunces", "IBM Plex Mono", "Arial"}
LINE_COLORS = {C.LINE, C.LINE2, "D6D0C2", "B9B4A6", "E2DCCE", "EBE6DB"}
# Systemregel „entfällt“ (bedingte Formatierung): lesbar zurückgenommen statt fast unsichtbar (D0D5DD ≈ 1,4:1)
FADED_OLD = {"D0D5DD", "DADCDF"}
FADED = "98A2B3"


# =============================================================================== Hilfen
def _rgb(color):
    if color is None or getattr(color, "type", None) != "rgb" or not isinstance(color.rgb, str):
        return None
    return color.rgb[-6:].upper()


def _fill_rgb(c):
    f = c.fill
    if f is None or f.fill_type != "solid":
        return None
    return _rgb(f.fgColor)


def _static(c):
    return isinstance(c.value, str) and c.data_type == "s" and not C.is_formula(c.value)


def _is_input(c):
    return _fill_rgb(c) == C.INPUT_BG or c.protection.locked is False


def _cells(ws, ref):
    if ":" not in ref:
        return [ws[ref]]
    return [c for row in ws[ref] for c in row]


def _merged_anchor_map(ws):
    m = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for cc in range(mr.min_col, mr.max_col + 1):
                m[(r, cc)] = (mr.min_row, mr.min_col, mr.max_col)
    return m


def _numeric_fmt(fmt):
    return fmt not in (None, "General", "@") and any(ch in fmt for ch in "0#")


def _map_size(c):
    """Zwischengrößen auf die Typo-Skala (core.snap_size: 9,5/11 → 10, 12 → 12,5, 14 → 16, < 8 → 8).
    8,5 pt ist seit Runde 3 eine eigene Stufe (T_LABEL: Versalien-Labels, Tabellenköpfe, Eyebrow) und bleibt."""
    sz = c.font.sz or 11
    if sz in SCALE:
        return sz
    return C.snap_size(sz)


def _set_font(c, **kw):
    f = c.font
    base = dict(name=f.name, sz=f.sz, b=f.b, i=f.i, u=f.u, color=f.color, strike=f.strike, vertAlign=f.vertAlign)
    base.update(kw)
    c.font = Font(**base)


def _set_align(c, **kw):
    a = c.alignment
    base = dict(horizontal=a.horizontal, vertical=a.vertical, indent=a.indent, wrap_text=a.wrap_text,
                shrink_to_fit=False, text_rotation=a.text_rotation)
    base.update(kw)
    c.alignment = Alignment(**base)


def _row_has_fill(ws, r, c1, c2, colors):
    return any(_fill_rgb(ws.cell(r, cc)) in colors for cc in range(C.col(c1), C.col(c2) + 1))


# =============================================================================== early
def ungroup_cols(ws):
    """Gruppierte Spaltenbreiten (min..max) in Einzelspalten zerlegen – gleiche Breiten, aber jede Spalte ist
    einzeln adressierbar (core.col_px und Breitenänderungen einzelner Spalten wirken damit korrekt)."""
    from copy import copy
    from openpyxl.utils import get_column_letter
    for key, dim in list(ws.column_dimensions.items()):
        lo, hi = dim.min or 0, dim.max or 0
        if not lo or hi <= lo:
            continue
        dim.max = lo
        for cc in range(lo + 1, hi + 1):
            letter = get_column_letter(cc)
            nd = copy(dim)
            nd.index = letter
            nd.min = nd.max = cc
            ws.column_dimensions[letter] = nd


def cell_rules(ws):
    """Korrektorat, Zahlenformate, Typo-Skala, keine senkrechten Linien (je Zelle)."""
    for c in list(ws._cells.values()):
        if c.number_format in NUMFMT_MAP:
            c.number_format = NUMFMT_MAP[c.number_format]
        if c.value is not None:
            sz = _map_size(c)
            if sz != c.font.sz:
                _set_font(c, sz=sz)
        b = c.border
        drop_l = b.left is not None and b.left.style in ("thin", "hair", "dotted") and _rgb(b.left.color) in LINE_COLORS
        drop_r = b.right is not None and b.right.style in ("thin", "hair", "dotted") and _rgb(b.right.color) in LINE_COLORS
        if (drop_l or drop_r) and not _is_input(c):
            c.border = Border(left=None if drop_l else b.left, right=None if drop_r else b.right, top=b.top,
                              bottom=b.bottom)
    numfmt_catalog(ws)


# =============================================================================== Zahlenformat-Katalog (P2-12)
_REF = re.compile(r"=\s*(?:'?([^'!]+)'?!)?\$?([A-Za-z_][A-Za-z0-9_.]*?)(\$?\d+)?\s*$")


def _pct_digits(fmt):
    """1 bzw. 2 bei einem schlichten Prozentformat (0.0 % / 0.00 % mit Varianten), sonst None."""
    if not fmt:
        return None
    head = fmt.split(";")[0].replace("\\", "").replace('"', "").replace(" ", "").lstrip("#")
    return {"0.0%": 1, "0.00%": 2}.get(head)   # auch „#0.00 %“ (Steuern) → Katalog pct2


def _ref_value(ws, c):
    """Wert einer schlichten Verknüpfung (=Name, =A1, =Blatt!A1) – None bei echten Berechnungen."""
    v = c.value
    if not C.is_formula(v):
        return None
    m = _REF.fullmatch(v.strip())
    if not m:
        return None
    sheet, name, row = m.groups()
    if row is None:
        return NAME_VALUES.get(name.lower(), None) if not sheet else None
    try:
        tgt = ws.parent[sheet] if sheet else ws
        val = tgt[f"{name.replace('$', '')}{row.replace('$', '')}"].value
    except Exception:
        return None
    return None if C.is_formula(val) else val


def _row_label(ws, c):
    for cc in range(c.column - 1, 1, -1):
        v = ws.cell(c.row, cc).value
        if isinstance(v, str) and not C.is_formula(v) and any(ch.isalpha() for ch in v):
            return v
    return ""


def _needs2(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v * 1000 - round(v * 1000)) > 1e-6


def _in_refs(coord_set, ws, key):
    if key not in coord_set:
        coord_set[key] = {c.coordinate for ref in key[1] for c in _cells(ws, ref)}
    return coord_set[key]


def numfmt_catalog(ws):
    """Katalog core.NUMFMT durchsetzen (P2-12): Prozent 0,0 % (zweite Stelle nur bei Bedarf bzw. Zins/Steuersatz),
    DSCR 0,00×, Jahre „1 Jahr / 12 Jahre“, Koeffizienten zweistellig. Eingabefelder zeigen 0 statt „–“."""
    t = ws.title
    cache = {}
    pct2 = _in_refs(cache, ws, ("p2", tuple(PCT2_CELLS.get(t, []))))
    for c in list(ws._cells.values()):
        if c.value is None:
            continue
        d = _pct_digits(c.number_format)
        if d is None:
            continue
        inp = _is_input(c)
        if inp and isinstance(c.value, (int, float)):
            v, kind = c.value, "entry"
        elif C.is_formula(c.value):
            v = _ref_value(ws, c)
            kind = "linked" if v is not None or _REF.fullmatch(c.value.strip()) else "calc"
        else:
            v, kind = c.value, "static"
        label = _row_label(ws, c).lower()
        if c.coordinate in pct2 or any(w in label for w in PCT2_WORDS):
            key = "pct2"
        elif kind == "calc":
            key = "pct1"
        else:
            key = "pct2" if _needs2(v) else "pct1"
        if kind == "static" and not inp:
            continue  # Kopf-/Achsenwerte (z. B. Sensitivität 0 % … 4 %) behalten ihr eigenes Format
        c.number_format = C.numfmt(key, entry=inp)
    for ref in DSCR_CELLS.get(t, []):
        for c in _cells(ws, ref):
            if c.value is not None and "%" not in (c.number_format or ""):
                c.number_format = C.numfmt("dscr", entry=_is_input(c))
    for ref in NUM2_CELLS.get(t, []):
        for c in _cells(ws, ref):
            c.number_format = C.numfmt("num2", entry=_is_input(c))
    for ref in PCT0_CELLS.get(t, []):
        for c in _cells(ws, ref):
            if c.value is not None and "%" in (c.number_format or ""):
                c.number_format = C.numfmt("pct0", entry=_is_input(c))
    # „(Jahre)“ im Label → Einheit im Zahlenformat
    for c in list(ws._cells.values()):
        if not (_static(c) and YEARS_LABEL.search(c.value)) or c.column > 3:
            continue
        for cc in range(c.column + 1, c.column + 3):
            v = ws.cell(c.row, cc)
            if v.value is None:
                continue
            if (v.number_format or "General") in ("General", "0", "0.0", "#,##0") and \
                    (isinstance(v.value, (int, float)) or C.is_formula(v.value)):
                v.number_format = C.NUMFMT["years_n"]
                C.set_text(c, YEARS_LABEL.sub("", c.value))
            break


def header(ws, spec):
    """Eyebrow (Z. 5) · H1 = Blattname (Z. 6) · Untertitel einzeilig (Z. 7)."""
    c1, c2, eyebrow, title, subtitle = spec
    for r in (5, 6, 7):
        for c in C.iter_cells(ws, 2, r, c2, r):
            if not C.is_formula(c.value):
                c.value = None
            c.fill = C.NOFILL
            c.border = Border()
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= 7 and mr.max_row >= 5 and mr.min_col >= 2:
            ws.unmerge_cells(str(mr))
    C.page_header(ws, c1, c2, eyebrow, title, subtitle)
    ws.row_dimensions[4].height = 12
    ws[f"{c1}7"].alignment = C.align("left", "top")


ACRONYMS = {"GMBH": "GmbH", "KFW": "KfW", "WEG": "WEG", "DSCR": "DSCR", "IRR": "IRR", "AFA": "AfA", "ESTG": "EStG",
            "ESTDV": "EStDV", "GRESTG": "GrEStG", "KSTG": "KStG", "GEWSTG": "GewStG", "USTG": "UStG", "SOLZG": "SolZG",
            "BGB": "BGB", "EK": "EK", "NK": "NK", "VK": "VK", "NOI": "NOI", "BMF": "BMF", "BFH": "BFH", "UST": "USt",
            "KAPEST": "KapESt", "HGB": "HGB", "JSTG": "JStG", "II": "II", "I": "I", "ABS.": "Abs.", "NR.": "Nr.",
            "S.": "S.", "§": "§", "I.": "i.", "D.": "d.", "F.": "F.", "P.": "p.", "A.": "a."}
LOWER_WORDS = {"und", "oder", "der", "die", "das", "des", "dem", "den", "im", "in", "mit", "nach", "vom", "von", "zu",
               "zur", "zum", "für", "auf", "aus", "bei", "ohne", "als", "über", "je", "pro", "inkl.", "bzw.", "ihre",
               "ihrer", "ihren", "am", "an", "bis", "ab", "lt.", "gem.", "sowie", "werden", "wird", "anteilig", "a",
               "einer", "eines", "ein", "eine", "gegen", "unter", "vor", "zzgl.", "ggf.", "sofort", "abziehbar",
               "abgeschrieben", "gesondert", "nur", "nicht", "ändern", "bitte", "hier", "je", "neu", "alle"}
ADJ_END = re.compile(r"(lich|liche|lichen|licher|liches|isch|ische|ischen|ischer|ige|igen|iger|bare|baren|barer|"
                     r"ierte|ierten)$")


def title_case(text):
    """Versalien-Band → Titelschreibung (Ebene 1, P1-12). Deutsche Heuristik: erstes Wort groß, Funktionswörter und
    Adjektive klein, Substantive groß, Abkürzungen/Normzitate korrekt. Texte in gemischter Schreibung bleiben."""
    if not isinstance(text, str) or C.is_formula(text) or text.upper() != text or not any(ch.isalpha() for ch in text):
        return text
    out, first = [], True
    for tok in re.split(r"(\s+|[–/()·,])", text):
        if not tok or tok.isspace() or tok in "–/()·,":
            out.append(tok)
            continue
        up = tok.upper()
        if up in ACRONYMS:
            out.append(ACRONYMS[up])
        elif not any(ch.isalpha() for ch in tok):
            out.append(tok)
            continue
        elif "-" in tok:
            out.append("-".join(ACRONYMS.get(p.upper(), p[:1] + p[1:].lower()) for p in tok.split("-")))
            first = False
            continue
        else:
            low = tok.lower()
            if first:
                out.append(low[:1].upper() + low[1:])
            elif low in LOWER_WORDS or ADJ_END.search(low):
                out.append(low)
            else:
                out.append(low[:1].upper() + low[1:])
        first = False
    return "".join(out)


def table_rules(ws, c1, c2, first=8, last=None):
    """Spaltenkopf, Summenzeile und Haarlinie über die volle Tabellenbreite; Köpfe folgen ihren Werten."""
    last = last or ws.max_row
    anchors = _merged_anchor_map(ws)
    band_colors = {C.NAVY, C.TINT}
    kinds = {}
    for r in range(first, last + 1):
        vals = [ws.cell(r, cc) for cc in range(C.col(c1), C.col(c2) + 1)]
        if not any(v.value is not None for v in vals) and not _row_has_fill(ws, r, c1, c2, band_colors | {C.HEAD}):
            continue
        if any(isinstance(v.value, str) and v.value.startswith("Keine Gewähr") for v in vals):
            break
        fills = {_fill_rgb(v) for v in vals}
        first_cell = ws.cell(r, C.col(c1))
        lb = first_cell.border.left
        if (_fill_rgb(first_cell) == C.NAVY or (_fill_rgb(first_cell) == C.TINT and lb is not None and lb.style == "thick")) \
                and not _is_input(first_cell):
            kinds[r] = "band"
        elif C.HEAD in fills:
            kinds[r] = "head"
        elif (fills & {C.TINT_XL, "F5F8FC", C.TINT}) and any(v.font.b for v in vals if v.value is not None):
            kinds[r] = "sum"
        elif any(v.value is not None for v in vals):
            kinds[r] = "row"
    for r, kind in kinds.items():
        if kind == "band":  # P1-12: ein Abschnittskopf (Ebene 1) auf allen Tabellenblättern
            first = ws.cell(r, C.col(c1))
            if _static(first) and not isinstance(first.value, type(None)):
                C.set_text(first, title_case(first.value))
            C.section(ws, r, c1, c2, level=1)
        elif kind == "head":
            below = next((k for k in range(r + 1, r + 4) if kinds.get(k) in ("row", "sum")), None)
            C.table_head(ws, r, c1, c2)
            for cc in range(C.col(c1), C.col(c2) + 1):
                h = ws.cell(r, cc)
                if h.value is None or below is None:
                    continue
                v = ws.cell(below, cc)
                if (r, cc) in anchors and anchors[(r, cc)][1] != cc:
                    continue
                num = isinstance(v.value, (int, float)) or (C.is_formula(v.value) and _numeric_fmt(v.number_format))
                h.alignment = C.align("right" if num else "left", "center", 1)
                if isinstance(h.value, str) and h.data_type == "s":
                    h.value = h.value.upper()
        elif kind == "sum":  # Zwischensumme (core.sum_row 'sub'); Eingabefelder behalten ihren Stil
            keep = {c.coordinate: (copy(c.font), copy(c.fill), copy(c.border)) for c in C.iter_cells(ws, c1, r, c2, r)
                    if _is_input(c)}
            C.sum_row(ws, r, c1, c2, "sub")
            for coord, (fnt, fil, brd) in keep.items():
                ws[coord].font, ws[coord].fill, ws[coord].border = fnt, fil, brd
        else:
            for c in C.iter_cells(ws, c1, r, c2, r):
                if _is_input(c):
                    continue
                b = c.border
                c.border = Border(top=b.top if (b.top is not None and b.top.style) else None,
                                  bottom=C.side("hair", C.LINE))
        if kind in ("row", "sum"):
            edges(ws, r, c1, c2, anchors)
    return kinds


def _is_num(c):
    return isinstance(c.value, (int, float)) and not isinstance(c.value, bool) or \
        (C.is_formula(c.value) and _numeric_fmt(c.number_format))


def edges(ws, r, c1, c2, anchors):
    """P1-06: Text links mit Einzug 1, Zahlen rechts mit Einzug 1 (zentrierte Zellen bleiben zentriert).
    Beschriftungen der ersten Spalte brechen um statt abgeschnitten zu werden."""
    for cc in range(C.col(c1), C.col(c2) + 1):
        c = ws.cell(r, cc)
        if c.value is None or ((r, cc) in anchors and anchors[(r, cc)][1] != cc):
            continue
        if c.alignment.horizontal == "center":
            continue
        if _is_num(c):
            _set_align(c, horizontal="right", indent=1)
        else:
            _set_align(c, horizontal="left", indent=1)
        if cc == C.col(c1) and _static(c):
            _set_align(c, wrap_text=True)


NAME_VALUES = {}


def resolve_names(wb):
    """Namen → aktueller Zellwert (für die Prozent-Regel übernommener Eingaben wie =Makler_Pct)."""
    NAME_VALUES.clear()
    for n, d in wb.defined_names.items():
        try:
            for title, coord in d.destinations:
                if title in wb.sheetnames and ":" not in coord:
                    v = wb[title][coord.replace("$", "")].value
                    if v is not None and not C.is_formula(v):
                        NAME_VALUES[n.lower()] = v
        except Exception:
            continue


def fit_row(ws, row, c1, c2, base=C.H_ROW, pad=6, max_lines=3):
    """Zeilenhöhe nach Inhalt – core.fit_row (Calibri-Metrik, Anzeigeformeln, Einzug, Raster 18/30/42)."""
    return C.fit_row(ws, row, c1, c2, base=base, pad=pad, max_lines=max_lines)


def row_grid(ws, kinds, c1, c2, pad=6, base=C.H_ROW, keep=None):
    """Zeilenraster: einzeilig 18 pt, zweizeilig 30 pt, dreizeilig 42 pt; einzeilig mittig, mehrzeilig oben.
    keep: Höhen einer bereits gerasterten Nachbartabelle derselben Zeilen (es gilt das Maximum)."""
    for r, kind in kinds.items():
        if kind == "head":
            h = C.H_HEAD
        elif kind == "band":
            h = C.H_BAND
        else:
            fit_row(ws, r, c1, c2, base=base, pad=pad)
            h = ws.row_dimensions[r].height
        if keep and r in keep:
            h = max(h, keep[r])
        ws.row_dimensions[r].height = h
        if kind in ("row", "sum"):
            for c in C.iter_cells(ws, c1, r, c2, r):
                if c.value is not None:
                    _set_align(c, vertical="center" if h <= base else "top")


def page_footer(ws, c1, c2, row=None):
    """Fuß: Leerzeile nach dem Inhalt, Oberkante, zwei Zeilen 8 pt grau (core.footer)."""
    if row is None:
        for c in ws._cells.values():
            if isinstance(c.value, str) and c.value.startswith("Keine Gewähr") and c.data_type == "s":
                row = c.row
                break
    if row is None:
        return
    for r in (row, row + 1):
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row <= r <= mr.max_row:
                ws.unmerge_cells(str(mr))
        for c in C.iter_cells(ws, c1, r, c2, r):
            c.value = None
    C.footer(ws, row, c1, c2)
    if row - 1 > 8:
        ws.row_dimensions[row - 1].height = 14


UNIT_TEXT = {"€": None, "%": None, "€/Monat": "pro Monat", "€ p. a.": "pro Jahr", "€/m²": "je m²",
             "€/m² p. a.": "je m² p. a.", "% Darlehen": "vom Darlehen", "% der Miete": "der Miete",
             "% Verkaufspreis": "vom VK-Preis", "fach": None}   # Faktor: Einheit „×“ steht im Zahlenformat


def forms_pre(ws):
    """Vor dem Tabellenstil: Einheiten, Kopftexte, Jahresspalte (Eingaben / Konfiguration)."""
    if ws.title == "Eingaben":
        for r in range(11, 146):
            e = ws.cell(r, 5)
            if _static(e) and e.value in UNIT_TEXT:
                if e.value == "fach":
                    ws.cell(r, 3).number_format = C.NUMFMT["mult1"]
                new = UNIT_TEXT[e.value]
                e.value = None if new is None else new
                if new is not None:
                    e.data_type = "s"
        for r in range(9, 146):
            c, d = ws.cell(r, 3), ws.cell(r, 4)
            if _fill_rgb(c) == C.HEAD and d.value is None and isinstance(c.value, str):
                C.safe_merge(ws, "C", r, "D", r)
                C.set_text(c, "WERT")
                c.alignment = C.align("right", "center", 1)
        for c in C.iter_cells(ws, "B", 71, "L", 71):  # „Summe Maßnahmen“ ist eine Summenzeile
            if c.value is not None:
                _set_font(c, b=True)
            if not _is_input(c):
                c.fill = C.fill(C.TINT)
    if ws.title == "Konfiguration":
        for r in range(45, 54):
            c = ws.cell(r, 2)
            c.number_format = C.NUMFMT["year"]
            _set_align(c, horizontal="left", indent=1)
        for r in range(10, 81):
            e = ws.cell(r, 5)
            if _static(e):
                _set_align(e, wrap_text=True)


def forms_defaults(ws):
    """Eingaben / Konfiguration: Kopf über Zahlen rechts, Datumsspalten zentriert, Jahre links (P1-06)."""
    if ws.title == "Eingaben":
        if ws["D10"].value is None:
            C.safe_merge(ws, "C", 10, "D", 10)
            C.set_text(ws["C10"], "WERT")
            ws["C10"].alignment = C.align("right", "center", 1)
        for r in range(11, 146):
            e = ws.cell(r, 5)
            if e.value is not None:
                _set_align(e, horizontal="left", indent=1)
    if ws.title == "Konfiguration":
        for r in range(10, 26):
            _set_align(ws.cell(r, 4), horizontal="center", indent=0)
        for coord in ("G27",):
            _set_align(ws[coord], horizontal="left", indent=0)
        g43 = ws["G43"]  # Listenkopf wie G9 / G15
        g43.fill = C.fill(C.HEAD)
        g43.font = C.font(C.T_MICRO, True, C.BLUE)
        g43.border = Border(bottom=C.side("thin", C.ACCENT))
        g43.alignment = C.align("left", "center", 1)
        if isinstance(g43.value, str):
            g43.value = g43.value.upper()
        for coord, text, h in (("C27", "WERT", "right"), ("E27", "FORMEL / ERLÄUTERUNG", "left"),
                               ("C56", "WERT", "right"), ("E56", "ERLÄUTERUNG", "left")):
            c = ws[coord]
            title = ws.cell(c.row, 2)
            if C.text_px(title.value, C.T_BODY, True) + 14 > C.span_px(ws, 2, c.column - 1):
                continue  # Bandtitel braucht den Platz
            if c.value is None:
                C.set_text(c, text)
                c.font = C.font(C.T_MICRO, True, C.SKY)
                c.alignment = C.align(h, "center", 1)
        if all(ws.cell(r, cc).value is None for r in (82, 83) for cc in range(2, 8)):
            page_footer(ws, "B", "G", 82)


NOTE_COLS = {"Eingaben": ("E", 11, 146), "Haushaltsrechnung": ("E", 18, 47), "Vermögensaufstellung": ("E", 10, 34),
             "Konfiguration": ("E", 10, 80)}


def note_cols(ws):
    """Einheiten- und Kommentarspalten: links mit Einzug 1, 9 pt grau (P1-06 (6))."""
    spec = NOTE_COLS.get(ws.title)
    if not spec:
        return
    letter, r1, r2 = spec
    for r in range(r1, r2 + 1):
        c = ws[f"{letter}{r}"]
        if c.value is None or _is_input(c):
            continue
        _set_align(c, horizontal="left", indent=1)
        if not c.font.b:
            _set_font(c, sz=C.T_SMALL, color=C.MUTED)


def bank_defaults(ws):
    """HH / VA: Ergebniszeile (Summenstufe „final“), Blattschutz mit entsperrten Eingaben (P2-08, P3-02).
    Die Statusfarben der Puffer setzt layouts/bank.py (P1-10, Wunsch F) – hier keine zweite Regel."""
    if ws.title == "Haushaltsrechnung":
        res = 46
    elif ws.title == "Vermögensaufstellung":
        res = 32
    else:
        return
    if ws.title == "Haushaltsrechnung" and isinstance(ws["E46"].value, str) and ws["E46"].data_type == "s":
        C.set_text(ws["E46"], "Banken erwarten > 0 nach neuem Kapitaldienst")
    # das EINE Blockergebnis; HH: Statusfarbe der Puffer setzt bank.py (C46:D46) – keine zweite Negativ-Regel (Wunsch F)
    C.sum_row(ws, res, "B", "E", "final", neg=False if ws.title == "Haushaltsrechnung" else None)
    for c in C.iter_cells(ws, "B", res, "E", res):
        if c.value is not None and c.column <= 4:
            _set_font(c, sz=C.T_H3, b=True, color=C.NAVY)
            _set_align(c, vertical="center")
        elif c.value is not None:
            _set_font(c, sz=C.T_SMALL, b=False, color=C.MUTED)
    ws.row_dimensions[res].height = C.H_BAND
    for r in range(9, res + 2):  # Kommentar-/Nachweisspalte neben Eingaben bleibt beschreibbar
        if ws.cell(r, 3).protection.locked is False:
            ws.cell(r, 5).protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.objects = True
    ws.protection.scenarios = True


def calc_colors(ws):
    """Formelergebnisse auf Rechenblättern in 1A1D21 – Blau nur für übernommene Eingaben (P2-08)."""
    for c in ws._cells.values():
        if c.row <= 7 or c.hyperlink is not None or not C.is_formula(c.value):
            continue
        if _rgb(c.font.color) == C.BLUE and not _is_input(c):
            _set_font(c, color=C.INK)


def early(wb):
    for ws in wb.worksheets:
        ungroup_cols(ws)
        cell_rules(ws)
        ws.sheet_properties.tabColor = tab_color(ws.title)
    for name, spec in HEADERS.items():
        if name in wb.sheetnames:
            header(wb[name], spec)
    resolve_names(wb)
    for name in FORM_SHEETS:
        if name in wb.sheetnames:
            forms_pre(wb[name])
    for name in NOTE_COLS:
        if name in wb.sheetnames:
            note_cols(wb[name])
    for name, spans in TABLES.items():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        done = {}
        for c1, c2 in spans:
            kinds = table_rules(ws, c1, c2, first=8)
            if name == "Hinweise":
                row_grid(ws, kinds, c1, c2, pad=10)
            else:
                row_grid(ws, kinds, c1, c2, keep=done)
            for r in kinds:
                done[r] = ws.row_dimensions[r].height or C.H_ROW
    for name, (c1, c2) in CONTENT.items():
        if name in wb.sheetnames and name != "Konfiguration":
            page_footer(wb[name], c1, c2)
    for name in FORM_SHEETS:
        if name in wb.sheetnames:
            forms_defaults(wb[name])
    for name in ("Haushaltsrechnung", "Vermögensaufstellung"):
        if name in wb.sheetnames:
            bank_defaults(wb[name])
    for name in CALC_SHEETS:
        if name in wb.sheetnames:
            calc_colors(wb[name])


# =============================================================================== final
def final(wb):
    for ws in wb.worksheets:
        finish_sheet(ws)


def finish_sheet(ws):
    """Sicherheits-Normalisierung eines Blatts (auch für das später gebaute Dashboard)."""
    normalise(ws)
    numfmt_catalog(ws)
    units_once(ws)
    units_convention(ws)
    zero_convention(ws)
    neg_red_net(ws)
    input_look(ws)
    breadcrumb_links(ws)
    cf_rules(ws)
    C.cf_close(ws)


# ------------------------------------------------------------------------------- Einheiten & Null (P38)
# Einheiten Jahre/Monate/m² stehen im Zahlenformat, die Einheitenspalte E nur für Qualifier („pro Monat“).
# Null: gelbe Eingabefelder zeigen 0 sichtbar, berechnete/übernommene Zellen „–“.
UNIT_FMT = {
    "years": ('[=1]0" Jahr";0" Jahre"', '[=1]0" Jahr";[=0]"–";0" Jahre"'),
    "months": ('[=1]0" Monat";0" Monate"', '[=1]0" Monat";[=0]"–";0" Monate"'),
    "qm": ('#,##0" m²"', f'#,##0" m²";{C._M}#,##0" m²";"–"'),
}
UNIT_KIND = {"Jahre": "years", "Jahr(e)": "years", "Jahr": "years", "Monate": "months", "Monat": "months",
             "m²": "qm"}
PLAIN_HEADS = {"0", "#,##0", "General"}


def _fmt_head(fmt):
    return (fmt or "General").split(";")[0].replace("\\", "").replace('"', "").strip()


def units_convention(ws):
    t = ws.title
    if STEP_RE.match(t):
        vcol = 4
    elif t == "Eingaben":
        vcol = 3
    else:
        return
    for r in range(9, ws.max_row + 1):
        v, e = ws.cell(r, vcol), ws.cell(r, 5)
        if v.value is None or not (isinstance(v.value, (int, float)) or C.is_formula(v.value)):
            continue
        if _fmt_head(v.number_format) not in PLAIN_HEADS:
            continue
        unit = e.value.strip() if _static(e) else None
        label = " ".join(str(ws.cell(r, cc).value or "") for cc in (2, 3) if _static(ws.cell(r, cc))).lower()
        kind = UNIT_KIND.get(unit) if unit else ("years" if re.search(r"jahre\b", label) else None)
        if kind is None:
            continue
        inp = _is_input(v)
        v.number_format = UNIT_FMT[kind][0 if inp else 1]
        if unit:
            e.value = None


UNIT_MARKS = ("€", "%", "×", "m²")


def _zero_part(part):
    return part.replace("\\", "").replace('"', "").strip()


def zero_convention(ws):
    """P2-07 – eine Nulldarstellung: berechnete/übernommene Zahlen zeigen 0 als „–“ (core.zero_dash), gelbe
    Eingabefelder zeigen 0 sichtbar (0 €, 0,00 %). Ausnahmen: nicht zutreffende Eingaben (ZERO_DASH_INPUTS) „–“,
    Summen direkt unter Eingabespalten (ZERO_AS_INPUT) wie die Eingaben „0 €“."""
    t = ws.title
    for c in ws._cells.values():
        if not C.is_formula(c.value) or _is_input(c):
            continue
        fmt = c.number_format or ""
        if not fmt or fmt in ("General", "@", ";;;") or not any(u in fmt for u in UNIT_MARKS):
            continue
        parts = C._split_fmt(fmt) if hasattr(C, "_split_fmt") else fmt.split(";")
        if len(parts) == 3 and "[" not in fmt and _zero_part(parts[2]) in ("0 €", "0.00 €", "0,00 €", "0 %",
                                                                             "0.0 %", "0.00 %"):
            c.number_format = ";".join(parts[:2] + ['"–"'])
        elif len(parts) == 2:
            c.number_format = C.zero_dash(fmt)
    for coord, key in ZERO_DASH_INPUTS.get(t, {}).items():
        c = ws[coord]
        if c.value is not None and not isinstance(c.value, str):
            c.number_format = C.NUMFMT[key]
    for coord in ZERO_AS_INPUT.get(t, []):
        c = ws[coord]
        if c.value is not None and "€" in (c.number_format or ""):
            c.number_format = C.NUMFMT["eur_in"]
    for coord in DASH_TEXT_SUMS.get(t, []):
        c = ws[coord]
        if isinstance(c.value, str) and c.value.strip() in ("–", "-"):
            ref = ws.cell(c.row, c.column - 1)
            c.font = copy(ref.font) if ref.value is not None else C.font(C.T_BODY, True, C.NAVY)
            _set_align(c, horizontal="right", indent=1, vertical=ref.alignment.vertical or "center")
    for key, refs in (("eur_qm", EUR_M2_CELLS), ("children", CHILDREN_CELLS)):
        for ref in refs.get(t, []):
            for c in _cells(ws, ref):
                if c.value is not None:
                    c.number_format = C.NUMFMT[key + "_in"] if (_is_input(c) and key + "_in" in C.NUMFMT) \
                        else C.NUMFMT[key]
    for c in ws._cells.values():
        if _static(c) and c.value.strip() in LABEL_FIX:
            C.set_text(c, LABEL_FIX[c.value.strip()])


# ------------------------------------------------------------------------------- Negativ-Rot (P15, Sicherheitsnetz)
NEG_RED_CELLS = {"S12 Ergebnis": ["D21", "D23"], "S09 Steuern": ["I15"], "Cockpit": ["G45"],
                 "Sensitivität": ["D55:H63"]}
NEG_RED_ROWS = {"Dashboard": ("Steuerliches Ergebnis", "E", "R")}


def _cf_covered(ws, coord):
    for cf in ws.conditional_formatting:
        if coord in cf.sqref:
            return True
    return False


def neg_red_net(ws):
    """Negative Ergebnis-/Kumulwerte rot – nur wo das Blatt-Modul noch keine Regel gesetzt hat."""
    refs = list(NEG_RED_CELLS.get(ws.title, []))
    spec = NEG_RED_ROWS.get(ws.title)
    if spec:
        row = _find_row(ws, spec[0], 3)
        if row:
            refs.append(f"{spec[1]}{row}:{spec[2]}{row}")
    for ref in refs:
        cells = [c for c in _cells(ws, ref) if c.value is not None and _numeric_fmt(c.number_format)
                 and "Zahlung" not in (c.number_format or "")]
        todo = [c for c in cells if not _cf_covered(ws, c.coordinate)]
        for c in todo:
            C.neg_red(ws, c.coordinate)


# ------------------------------------------------------------------------------- Eingabe-Optik (P31)
COMMENT_COLS = {"Haushaltsrechnung": ("E", 18, 47), "Vermögensaufstellung": ("E", 10, 34)}


def input_look(ws):
    """Jede entsperrte Zelle ist als Eingabe erkennbar. Kommentar-/Nachweisspalten der Bankformulare: zarte
    Eingabefläche, Unterkante im Eingabegelb, 9 pt kursiv grau (nur wo das Blatt-Modul nichts gesetzt hat)."""
    spec = COMMENT_COLS.get(ws.title)
    if not spec:
        return
    letter, r1, r2 = spec
    anchors = _merged_anchor_map(ws)
    for r in range(r1, r2 + 1):
        e = ws[f"{letter}{r}"]
        if e.protection.locked is not False or _fill_rgb(e) is not None:
            continue
        a = anchors.get((r, e.column))
        if a and (a[0], a[1]) != (r, e.column):
            continue
        if isinstance(e.value, str) and e.value.isupper():
            continue
        e.fill = C.fill(C.NOTE_BG)
        e.border = Border(bottom=C.side("thin", C.INPUT_LINE))
        e.font = Font(name=C.SANS, sz=C.T_SMALL, i=True, color=C.MUTED)
        e.alignment = C.align("left", "center", 1)


# ------------------------------------------------------------------------------- Zell-Link-Rückfall (P05)
# P2-04: erster Krumenteil → Elternblatt des Reiters (core.CRUMB_PARENT: BERECHNUNG → Dashboard, STEUER-TABELLE → Steuern,
# ANHANG → Start, LEITFADEN → Leitfaden …). Einteilige Krumen („DASHBOARD“, „COCKPIT“) sind reine Überzeilen ohne Link.
CRUMB_TARGET = {"BANK": "Bankgespräch", "STEUERN": "Steuern", "LEITFADEN": "Start", "EINSTIEG": "Start",
                **getattr(C, "CRUMB_PARENT", {})}
SUBNAV_SHEETS = ("Steuern", "AfA-Vergleich", "Bankgespräch", "Haushaltsrechnung", "Vermögensaufstellung",
                 "Hinweise", "Konfiguration")
BACK_EDGE = {"Leitfaden": "I", "Diagramme": "P", "Sensitivität": "M", "Eingaben": "L"}   # rechte Inhaltskante


def crumb_tooltip(target):
    """ScreenTip der Brotkrume – dieselbe Sprache wie die Formen („Zurück zur Startseite“, „Zurück: Dashboard“)."""
    if target == "Start":
        return "Zurück zur Startseite"
    lab = C.sheet_label(target) if hasattr(C, "sheet_label") else target
    return f"Zurück: {lab or target}"


def breadcrumb_links(ws):
    """Navigation ohne Formen: Brotkrume (Z. 5) „REITER › BLATT“ ist Zell-Link auf das Elternblatt des Reiters
    (S01–S12 → Leitfaden); einteilige Krumen bleiben ohne Link. Wo Z. 5 rechts frei ist (keine Unterreiter),
    zusätzlich „‹  Elternblatt“ an der Inhaltskante."""
    t = ws.title
    if t == "Start":
        return
    wbk = ws.parent
    eb = next((ws.cell(5, cc) for cc in (2, 3) if _static(ws.cell(5, cc)) and ws.cell(5, cc).value.strip()), None)
    if eb is None:
        return
    # Brotkrume einheitlich 8,5 pt fett Versalien 1D4F8A (core.page_header)
    eb.font = C.font(C.T_LABEL, True, C.BLUE)
    parts = [p.strip() for p in eb.value.split("›") if p.strip()]
    single = len(parts) < 2 and not STEP_RE.match(t)
    if STEP_RE.match(t):
        target = "Leitfaden"
    else:
        first = re.split(r"·", parts[0])[0].strip().upper()
        target = CRUMB_TARGET.get(first, "Start")
    # „Dashboard“ entsteht erst nach final() (dashboard.build) – als Ziel trotzdem gültig
    if target == t or (target not in wbk.sheetnames and target != "Dashboard"):
        target = "Start"
    if single:
        target = C.PARENT.get(t, "Start") if hasattr(C, "PARENT") else "Start"
    tip = crumb_tooltip(target)
    if single:
        if eb.hyperlink is not None:
            eb.hyperlink = None   # „DASHBOARD“ / „COCKPIT“ / „DIAGRAMME“: Überzeile, kein unerklärter Sprung
    elif eb.hyperlink is None or eb.hyperlink.location != C.link_loc(target):
        eb.hyperlink = Hyperlink(ref=eb.coordinate, location=C.link_loc(target), tooltip=tip)
    else:
        eb.hyperlink.tooltip = tip
    if t in SUBNAV_SHEETS or t in ("Dashboard", "Cockpit") or C.link_target(t) != "A4":
        return
    edge = "I" if STEP_RE.match(t) else BACK_EDGE.get(t)
    if not edge:
        return
    cell = ws[f"{edge}5"]
    anchors = _merged_anchor_map(ws)
    if cell.value is not None or (5, cell.column) in anchors or cell.column <= eb.column:
        return
    C.text_link(cell, f"‹  {target}", target, size=C.T_LABEL, bold=False, tooltip=tip)
    cell.alignment = C.align("right", "bottom")


def cf_rules(ws):
    """dxf ohne Schriftnamen; „entfällt“-Grau lesbar (98A2B3 kursiv statt D0D5DD); Warn-/Info-Zeilen ohne Emoji."""
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            d = rule.dxf
            if d is None or d.font is None:
                continue
            d.font.name = None
            if _rgb(d.font.color) in FADED_OLD and d.fill is None:
                d.font.color = FADED
                d.font.i = True


def units_once(ws):
    """Eine Einheit wird nur einmal genannt: steht sie schon in der Beschriftung oder im Zahlenformat (×),
    bleibt die Einheitenspalte leer (Jahre/Monate/m² wandern über units_convention ins Zahlenformat)."""
    if ws.title not in ("Eingaben", "Konfiguration"):
        return
    for r in range(9, ws.max_row + 1):
        b, cval, e = ws.cell(r, 2), ws.cell(r, 3), ws.cell(r, 5)
        if not (_static(e) and isinstance(b.value, str)) or _is_input(e):
            continue
        unit = e.value.strip()
        label = b.value.lower()
        if len(unit) <= 14 and unit.lower() in label:
            e.value = None
        elif "×" in (cval.number_format or "") and unit in ("Jahresmieten", "fach"):
            e.value = None


def _rich_runs(c):
    """Rich-Text: Typo-Skala, Sekundärtext ≤ 9 pt nie in 8A9099, nur Calibri, Mindestgröße 8 pt."""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    v = c.value
    if not isinstance(v, CellRichText):
        return
    for part in v:
        if not isinstance(part, TextBlock) or part.font is None:
            continue
        f = part.font
        col = _rgb(f.color) if f.color is not None else None
        if (f.sz or 11) <= 9 and col == C.MUTED2:
            f.color = C.MUTED
        if f.sz is not None and f.sz not in SCALE:
            f.sz = C.snap_size(f.sz)
        if f.rFont in OLD_FONTS - {None}:
            f.rFont = C.SANS


def _zero_visible(fmt):
    """Format ohne „–“-Nullabschnitt: editierbare Felder zeigen 0 (bzw. 0 €, 0 %) statt eines Strichs."""
    parts = fmt.split(";")
    if len(parts) == 3 and parts[2].strip('"\\ ') in ("–", "-"):
        return ";".join(parts[:2])
    return fmt


def _plain_symbols(text):
    """Farbige Emoji → monochrome Zeichen (P2-14)."""
    for k, v in EMOJI_MAP.items():
        text = text.replace(k, v)
    return text


def normalise(ws):
    for c in ws._cells.values():
        _rich_runs(c)
        fmt = c.number_format
        # P16: Formelzellen nie im Textformat „@“ (Excel zeigt sonst beim Bearbeiten die Formel als Text)
        if fmt == "@" and C.is_formula(c.value):
            c.number_format = fmt = "General"
        # P23: typografisches Minus „−“ in allen Zahlenformaten (idempotent; Text/Datum/Bedingungen bleiben)
        if fmt and fmt not in ("General", "@"):
            new = C.typo_minus(fmt)
            if new != fmt:
                c.number_format = fmt = new
        if fmt and ";" in fmt and _is_input(c):
            c.number_format = _zero_visible(fmt)
        al = c.alignment
        f = c.font
        # 1) kein „An Zellgröße anpassen“; Einzug nur mit links/rechts
        if al.shrink_to_fit or (al.indent and al.horizontal not in ("left", "right", "distributed")):
            h = al.horizontal
            if al.indent and h not in ("left", "right", "distributed"):
                num = isinstance(c.value, (int, float)) or (C.is_formula(c.value) and _numeric_fmt(c.number_format))
                h = "right" if num else "left"
            _set_align(c, horizontal=h, shrink_to_fit=False)
        if c.value is None:
            if f.name in OLD_FONTS:  # leere Zellen: beim Tippen erscheint sonst die Altschrift
                _set_font(c, name=C.SANS)
            continue
        # 2) Mindestschrift 8 pt und Typo-Skala (P3-11); 3) Sekundärtext ≤ 9 pt nie in 8A9099 (außer inaktive Felder)
        kw = {}
        if (f.sz or 11) not in SCALE:
            kw["sz"] = _map_size(c)
        if (kw.get("sz") or f.sz or 11) <= 9 and _rgb(f.color) == C.MUTED2 and _fill_rgb(c) != C.INACTIVE_BG:
            kw["color"] = C.MUTED
        # 4) nur Calibri (monochrome Symbolschrift bleibt erlaubt)
        if f.name in OLD_FONTS:
            kw["name"] = C.SANS
        if kw:
            _set_font(c, **kw)
        # 5) Normzitate in Versaltexten; farbige Emoji → monochrome Zeichen
        if _static(c):
            v = c.value
            if sum(ch.isupper() for ch in v) > 0.6 * max(1, sum(ch.isalpha() for ch in v)):
                for pat, rep in VERSAL_FIX:
                    v = pat.sub(rep, v)
            v = _plain_symbols(v)
            if v != c.value:
                C.set_text(c, v)


# =============================================================================== Druck (P1-23) und Register (P3-02)
MARGINS = dict(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.3, footer=0.3)


def _footer_row(ws, c1=1, c2=60):
    rows = [c.row for c in ws._cells.values()
            if isinstance(c.value, str) and c.value in (C.FOOTER_2,) and C.col(c1) <= c.column <= C.col(c2)]
    if rows:
        return max(rows)
    rows = [c.row for c in ws._cells.values() if isinstance(c.value, str) and c.value.startswith("MM HOLDING GMBH ·")]
    return max(rows) if rows else None


def _print(ws, area, orient="landscape", w=1, h=0, scale=None, rows=None, cols=None, breaks=(), col_breaks=(),
           over_then_down=False):
    """Druckeinrichtung. P1-06 (Runde 4): Excel ignoriert manuelle Umbrüche bei „Anpassen an“ – sind Zeilen- oder
    Spaltenumbrüche gesetzt, gilt deshalb IMMER ein fester Maßstab (fitToPage aus). Ohne Vorgabe ergibt er sich aus
    min(Seitenbreite, höchster Block zwischen zwei Umbrüchen) – so steht jeder Block geschlossen auf einer Seite.
    area: ein Bereich „A1:L76“ oder mehrere, durch Komma getrennt (jeder Teilbereich beginnt eine neue Seite)."""
    ws.print_area = area
    ws.page_setup.orientation = orient
    if scale is None and (breaks or col_breaks):
        scale = _fixed_scale(ws, area, orient, breaks, rows, cols, col_breaks)
    if scale:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
        ws.page_setup.scale = int(scale)
        ws.page_setup.fitToWidth = None
        ws.page_setup.fitToHeight = None
    else:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = w
        ws.page_setup.fitToHeight = h
        ws.page_setup.scale = None
    if rows:
        ws.print_title_rows = rows
    else:
        ws._print_rows = None
    if cols:
        ws.print_title_cols = cols
    else:
        ws._print_cols = None
    ws.row_breaks = RowBreak()
    for r in sorted(set(breaks)):
        ws.row_breaks.append(Break(id=r - 1))
    ws.col_breaks = ColBreak()
    for cc in col_breaks:
        ws.col_breaks.append(Break(id=C.col(cc)))
    ws.page_setup.pageOrder = "overThenDown" if over_then_down else None


def _areas(area):
    """„A1:M39,A40:AQ76“ → [(c1, r1, c2, r2), …]."""
    out = []
    for part in area.split(","):
        m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", part.strip().split("!")[-1])
        if m:
            out.append((m.group(1), int(m.group(2)), m.group(3), int(m.group(4))))
    return out


def _fixed_scale(ws, area, orient="landscape", breaks=(), rows=None, cols=None, col_breaks=()):
    """Größter ganzzahliger Maßstab, bei dem jede Seite zwischen zwei Umbrüchen (bzw. Teilbereichsgrenzen) in Breite
    und Höhe auf A4 passt – Drucktitel (Zeilen/Spalten) eingerechnet."""
    aw, ah = _avail(orient)
    t_h, t_r2, t_w, t_c2 = 0.0, 0, 0.0, 0
    if rows:
        a, t_r2 = (int(x) for x in rows.replace("$", "").split(":"))
        t_h = sum(_row_pt(ws, r) for r in range(a, t_r2 + 1))
    if cols:
        a, b = cols.replace("$", "").split(":")
        t_w, t_c2 = _width_pt(ws, a, b), C.col(b)
    need = [1.0]
    for c1, r1, c2, r2 in _areas(area):
        cuts = sorted({C.col(c) for c in col_breaks if C.col(c1) <= C.col(c) < C.col(c2)})
        edges = [C.col(c1) - 1] + cuts + [C.col(c2)]
        for k in range(len(edges) - 1):   # Drucktitel-Spalten kommen auf jede Seite rechts von ihnen dazu
            wpt = _width_pt(ws, edges[k] + 1, edges[k + 1]) + (t_w if edges[k] + 1 > t_c2 else 0.0)
            need.append(aw * RESERVE_W / max(wpt, 1))
        bounds = [r1] + sorted(b for b in set(breaks) if r1 < b <= r2) + [r2 + 1]
        for k in range(len(bounds) - 1):  # Drucktitel-Zeilen auf jeder Seite unterhalb von ihnen
            hpt = sum(_row_pt(ws, r) for r in range(bounds[k], bounds[k + 1]))
            need.append(ah * RESERVE / max(hpt + (t_h if bounds[k] > t_r2 else 0.0), 1))
    return max(10, int(min(need) * 100))


# ------------------------------------------------------------------------------- Seitenplanung
PAPER = {"landscape": (842.0, 595.0), "portrait": (595.0, 842.0)}   # A4 in pt
RESERVE = 0.985                                                      # Rundungs-/Druckerreserve
RESERVE_W = 0.92   # fester Maßstab (P1-06): Breitenreserve – Druckspalten fallen je nach Treiber/Schriftmetrik bis ~6 % breiter
                   # aus als die Bildschirmpixel („Anpassen an“ rechnete das bisher selbst)


def _avail(orient):
    pw, ph = PAPER[orient]
    return (pw - 72 * (MARGINS["left"] + MARGINS["right"]),
            ph - 72 * (MARGINS["top"] + MARGINS["bottom"]))


def _row_pt(ws, r):
    d = ws.row_dimensions.get(r) if hasattr(ws.row_dimensions, "get") else ws.row_dimensions[r]
    if d is not None and d.hidden:
        return 0.0
    return float(d.height) if (d is not None and d.height) else 15.0


def _width_pt(ws, c1, c2):
    return C.span_px(ws, c1, c2) * 0.75


def _chart_spans(ws):
    """Zeilenbereiche (1-basiert) aller Diagramme/Bilder – Seitenumbrüche nie mitten hindurch."""
    out = []
    for obj in list(getattr(ws, "_charts", [])) + list(getattr(ws, "_images", [])):
        a = obj.anchor
        fr, to = getattr(a, "_from", None), getattr(a, "to", None)
        if fr is not None and to is not None:
            out.append((fr.row + 1, to.row + 1))
    return out


def _merged_spans(ws):
    return [(mr.min_row, mr.max_row) for mr in ws.merged_cells.ranges if mr.max_row > mr.min_row]


def block_starts(ws, c1, c2, r1, r2):
    """Mögliche Umbruchstellen: Zeilen, vor denen eine Leer-/Fugenzeile steht, und Abschnittsköpfe."""
    def empty(r):
        return all(ws.cell(r, cc).value is None for cc in range(C.col(c1), C.col(c2) + 1)) and \
            not _row_has_fill(ws, r, c1, c2, {C.NAVY, C.TINT, C.HEAD, C.TINT_XL, C.BLUE})
    out = []
    for r in range(r1 + 1, r2 + 1):
        if _row_pt(ws, r) == 0:
            continue
        first = ws.cell(r, C.col(c1))
        lb = first.border.left
        head = _fill_rgb(first) in (C.NAVY,) or (_fill_rgb(first) == C.TINT and lb is not None and lb.style == "thick")
        prev = r - 1
        while prev > r1 and _row_pt(ws, prev) == 0:
            prev -= 1
        if head or (empty(prev) and _row_pt(ws, prev) >= 7.5 and not empty(r)):
            out.append(r)
    return out


def section_starts(ws, c1, c2, r1, r2):
    """Nur Abschnittsköpfe (Ebene 1 bzw. Navy-Band) – für die Aufteilung von Präsentationsblättern."""
    out = []
    for r in range(r1 + 1, r2 + 1):
        for cc in range(C.col(c1), C.col(c2) + 1):
            c = ws.cell(r, cc)
            lb = c.border.left
            if c.value is not None and _fill_rgb(c) == C.TINT and lb is not None and lb.style == "thick":
                out.append(r)
                break
            if isinstance(c.value, str) and not C.is_formula(c.value) and c.font is not None \
                    and (c.font.sz or 0) in (C.T_H3, C.T_H2) and c.font.b \
                    and _rgb(c.font.color) == C.NAVY and r > 8:
                out.append(r)
                break
    return out


def best_split(ws, c2, r2, orient="landscape", pages=2, candidates=()):
    """Aufteilung auf genau `pages` Seiten mit dem größten Maßstab (Umbrüche nur vor Abschnittsköpfen).
    Bei gleichem Maßstab (breitenbegrenzt) gewinnt die ausgewogenste Aufteilung (kleinster höchster Block)."""
    from itertools import combinations
    aw, ah = _avail(orient)
    w_scale = min(1.0, aw * RESERVE_W / max(_width_pt(ws, "A", c2), 1))
    heights = {r: _row_pt(ws, r) for r in range(1, r2 + 1)}
    best = (0.0, 0.0, ())
    for cut in combinations(sorted(set(candidates)), pages - 1):
        bounds = (1,) + cut + (r2 + 1,)
        segs = [sum(heights[r] for r in range(bounds[k], bounds[k + 1])) for k in range(pages)]
        sc = int(min([w_scale] + [ah * RESERVE / max(h, 1) for h in segs]) * 100)
        key = (sc, -max(segs))
        if key > best[:2]:
            best = (sc, -max(segs), cut)
    return best[0], list(best[2])


def paginate(ws, c1, c2, r1, r2, orient="landscape", forced=(), candidates=None, title_rows=None, min_scale=70,
             max_scale=100, one_page_segments=False):
    """Maßstab und Zeilenumbrüche so planen, dass jede Seite voll lesbar ist:
    Maßstab = Seitenbreite (max. 100 %); Umbrüche vor Abschnitten/Blöcken (nie durch Diagramme oder Verbünde).
    one_page_segments: jedes erzwungene Segment genau eine Seite (Maßstab sinkt dafür bis min_scale).
    Liefert (Maßstab in %, Umbruchzeilen)."""
    aw, ah = _avail(orient)
    scale = min(max_scale / 100, aw * RESERVE_W / max(_width_pt(ws, c1, c2), 1))   # = fester Druckmaßstab
    t_h = 0.0
    if title_rows:
        a, b = title_rows
        t_h = sum(_row_pt(ws, r) for r in range(a, b + 1))
    forced = sorted(r for r in set(forced) if r1 < r <= r2)
    if one_page_segments:
        bounds = [r1] + forced + [r2 + 1]
        segs = [sum(_row_pt(ws, r) for r in range(bounds[k], bounds[k + 1])) for k in range(len(bounds) - 1)]
        need = min(ah * RESERVE / max(h, 1) for h in segs)
        if need >= min_scale / 100:
            return int(min(scale, need) * 100), forced
    heads = set(section_starts(ws, c1 if c1 != "A" else "B", c2, r1, r2))
    cands = sorted(set(candidates if candidates is not None else block_starts(ws, c1, c2, r1, r2)) | set(forced) | heads)

    blocked = set()
    for a, b in _chart_spans(ws) + _merged_spans(ws):
        blocked.update(range(a + 1, b + 1))
    breaks, start, used = [], r1, 0.0
    limit = ah * 0.97 / scale
    r = r1
    while r <= r2:
        h = _row_pt(ws, r)
        room = limit - (t_h if breaks else 0.0)
        if r in forced and r > start:
            breaks.append(r)
            start, used = r, 0.0
        elif used + h > room and r > start:
            opts = [c for c in cands if start < c <= r and c not in blocked
                    and sum(_row_pt(ws, k) for k in range(start, c)) >= 0.3 * room]
            # keine Seite nur mit dem Fuß: Umbrüche kurz vor dem Blattende meiden
            full = [c for c in opts if sum(_row_pt(ws, k) for k in range(c, r2 + 1)) >= 0.15 * room]
            opts = full or opts
            cut = max(opts) if opts else r
            while cut in blocked and cut > start + 1:
                cut -= 1
            # Überschrift nie allein am Seitenende: Kopfzeilen direkt vor dem Umbruch wandern mit
            lead = [h for h in heads if cut - 3 <= h < cut and h > start]
            if lead:
                cut = min(lead)
            breaks.append(cut)
            start = cut
            used = sum(_row_pt(ws, k) for k in range(cut, r))
        used += h
        r += 1
    return int(scale * 100), breaks


def _chart_right_col(ws):
    cols = []
    for ch in getattr(ws, "_charts", []):
        to = getattr(ch.anchor, "to", None)
        if to is not None:
            cols.append(to.col + (1 if to.colOff else 0))
    return max(cols) if cols else 0


def _find_row(ws, text, c_max=20, default=None, contains=False, r_min=1):
    rows = [c.row for c in ws._cells.values()
            if c.column <= c_max and c.row >= r_min and isinstance(c.value, str) and not C.is_formula(c.value)
            and (text in c.value if contains else c.value.strip() == text)]
    return min(rows) if rows else default


def _smart(ws, c1, c2, r2, orient="landscape", forced=(), title_rows=None, one_page_segments=False, min_scale=70,
           **kw):
    """Druckbereich ab A1 (Kopfleiste und Akzentlinie immer im Druck) mit geplanten Umbrüchen.
    Mit Umbrüchen fester Maßstab (P1-06), ohne Umbruch „1 Seite breit“."""
    scale, breaks = paginate(ws, "A", c2, 1, r2, orient, forced=forced, title_rows=title_rows,
                             one_page_segments=one_page_segments, min_scale=min_scale)
    rows = f"{title_rows[0]}:{title_rows[1]}" if title_rows else None
    if not breaks and scale >= _width_scale(ws, c2, orient) - 1:
        _print(ws, f"A1:{c2}{r2}", orient, 1, 0, rows=rows, **kw)
    else:
        _print(ws, f"A1:{c2}{r2}", orient, rows=rows, breaks=breaks, **kw)
    return ws.page_setup.scale or scale, breaks


def _width_scale(ws, c2, orient="landscape"):
    aw, _ = _avail(orient)
    return int(min(1.0, aw * RESERVE / max(_width_pt(ws, "A", c2), 1)) * 100)


def print_setup(ws):
    """Druck (P1-02, P2-05, P2-06): alle Druckbereiche beginnen bei A1 (Kopfleiste und Akzentlinie im Druck),
    der Fuß mit Haftungsausschluss liegt immer im Druckbereich, jede Seite mindestens 70 % Maßstab."""
    t = ws.title
    fr = _footer_row(ws)
    if STEP_RE.match(t):  # ein Schritt = eine Seite
        _print(ws, f"A1:I{_footer_row(ws, 1, 9) or fr or 60}", "landscape", 1, 1)
    elif t == "Start":  # genau zwei Seiten, Umbruch vor dem Abschnitt mit dem größten Maßstab (Legende/Vorgehen)
        last = max(_footer_row(ws, 1, 8) or 0, 65)
        scale, breaks = best_split(ws, "H", last, pages=2, candidates=section_starts(ws, "B", "H", 20, last - 3))
        _print(ws, f"A1:H{last}", "landscape", breaks=breaks)   # fester Maßstab: der Umbruch wirkt auch in Excel
        if ws.freeze_panes is None:
            ws.freeze_panes = "A4"
    elif t == "Dashboard":  # quer, 3 Seiten (Wunsch G): Kopf/Urteil/Kacheln/Check | Vermögensentwicklung |
        # Kennzahlen im Zeitverlauf + Prüfhinweise – jede Seite vollständig, gemeinsamer Maßstab (≈ 75 %)
        last = _footer_row(ws, 1, 18) or ws.max_row
        cut = _find_row(ws, "Vermögensentwicklung", 18, 28)
        trend = _find_row(ws, "Kennzahlen im Zeitverlauf", 18, None, r_min=cut + 1)
        hints = _find_row(ws, "Prüfhinweise", 18, None, contains=True, r_min=cut + 1)
        forced = tuple(r for r in (cut, trend) if r)
        scale, breaks = _smart(ws, "A", "R", last, forced=forced, one_page_segments=True)
        if scale < 70 or len(breaks) > len(forced):  # zu lang: Prüfhinweise auf eine eigene Seite
            _smart(ws, "A", "R", last, forced=tuple(r for r in (cut, trend, hints) if r), one_page_segments=True)
    elif t == "Leitfaden":
        _print(ws, f"A1:I{_footer_row(ws, 1, 9) or 53}", "landscape", 1, 1)
    elif t == "Cockpit":  # P1-06 / Wunsch C: drei Blöcke, je eine Seite – Diagrammzeile (57–74) nie geteilt
        last = _footer_row(ws, 1, 12) or 76
        charts = _find_row(ws, "Diagramme", 12, None, r_min=40)
        forced = tuple(r for r in (32, charts or 57) if r <= last)
        _smart(ws, "A", "L", last, forced=forced)
    elif t == "Eingaben":
        _smart(ws, "A", "L", _footer_row(ws, 1, 12) or 148)
    elif t == "Diagramme":  # P1-06: Umbrüche nur an Abschnittsgrenzen (vor 30/51/92/113), fester Maßstab
        last = _footer_row(ws, 1, 42) or ws.max_row
        data = _find_row(ws, "Diagrammdaten", 3, contains=True, r_min=100) or 138
        visible_data = any(_row_pt(ws, r) > 0 and any(ws.cell(r, cc).value is not None for cc in range(2, 17))
                           for r in range(data + 1, max(data + 1, last - 2)))
        forced = [r for r in (30, 51, 92, 113) if r < data]
        if visible_data:
            forced.append(data)
        _smart(ws, "A", "P", last, forced=tuple(forced))
    elif t == "Steuern":  # P1-06: Parameter | Jahrestabelle (Jahr 1–40, Spalten A:C wiederholt) | Verkaufsszenario
        last = _footer_row(ws) or ws.max_row
        head = _find_row(ws, "Jahr der Kalkulation", 3, 40)
        exit_ = _find_row(ws, "Verkaufsszenario", 3, None, contains=True, r_min=head + 1) or 77
        area = f"A1:M{head - 1},A{head}:AQ{exit_ - 1},A{exit_}:M{last}"
        _print(ws, area, "landscape", cols="A:C", col_breaks=("M", "W", "AG"), over_then_down=True)
    elif t in ("Projektion", "Finanzierung"):  # Folgeseiten (Jahr 11–40) ab dem Tabellenkopf, ohne leeres Kopfband
        last = _footer_row(ws) or ws.max_row
        _, breaks = paginate(ws, "A", "M", 1, last, "landscape", title_rows=(8, 10), max_scale=75)
        area = f"A1:M{last},N8:AQ{last}"
        _print(ws, area, "landscape", rows="8:10", cols="A:C", breaks=breaks, col_breaks=("W", "AG"),
               over_then_down=True)
    elif t == "AfA-Vergleich":
        _smart(ws, "A", "Q", _footer_row(ws, 1, 17) or 75)
    elif t == "Sensitivität":  # Hilfsrechnung 65–115 ist eingeklappt; der Fuß (117/118) kommt mit auf die Seite
        _smart(ws, "A", "M", _footer_row(ws, 1, 13) or 118)
    elif t == "Bankgespräch":  # hoch, eine Seite
        _print(ws, f"A1:I{max(_footer_row(ws, 1, 9) or 0, 67)}", "portrait", 1, 1)
    elif t in ("Haushaltsrechnung", "Vermögensaufstellung"):
        from openpyxl.utils import get_column_letter
        last = _footer_row(ws, 1, 5) or (50 if t == "Haushaltsrechnung" else 37)
        right = _chart_right_col(ws)
        if right > 5:  # Diagramm rechts neben der Tabelle: quer, sonst hoch (A:E); Umbruch nur vor Abschnitten
            _smart(ws, "A", get_column_letter(right), last, "landscape")
        else:
            _smart(ws, "A", "E", last, "portrait")
    elif t == "Hinweise":
        _smart(ws, "A", "E", _footer_row(ws, 1, 5) or 39)
    elif t == "Konfiguration":
        _smart(ws, "A", "G", _footer_row(ws, 1, 7) or 83)


def page_label(title):
    """Druck-Kopfzeile: Reitername statt Blattname (core.sheet_label, z. B. Finanzierung → Tilgungsplan)."""
    try:
        lab = C.sheet_label(title)
    except Exception:
        lab = None
    return lab or title


def page_setup(wb):
    if "Dashboard" in wb.sheetnames:  # entsteht erst nach final(): dieselben Sicherheitsregeln
        finish_sheet(wb["Dashboard"])
    for ws in wb.worksheets:
        ps = ws.page_setup
        ps.paperSize = 9
        for k, v in MARGINS.items():
            setattr(ws.page_margins, k, v)
        ws.print_options.horizontalCentered = True
        # P2-04: Kopfzeile nennt das Blatt wie der Reiter („Steuer-Tabelle“, „Tilgungsplan“, „Schritt 02 · …“);
        # die Fußzeile behält &A (Blattname = Zuordnung der Seiten, auch für die Vorschau)
        ws.oddHeader.right.text = None if ws.title in BANK_SHEETS else page_label(ws.title).replace("&", "&&")
        ws.oddHeader.right.size = 8
        ws.oddHeader.right.color = C.MUTED
        ws.oddHeader.right.font = "Calibri,Regular"
        ws.oddFooter.left.text = "MM Holding GmbH · Immobilien-Kalkulation · &A"
        ws.oddFooter.right.text = "Seite &P von &N"
        for part in (ws.oddFooter.left, ws.oddFooter.right):
            part.size = 8
            part.font = "Calibri,Regular"
            part.color = C.MUTED
        try:
            print_setup(ws)
        except Exception as exc:  # ein Blatt darf den Druck-Setup der übrigen nicht verhindern
            print(f"WARNUNG Druck {ws.title}: {exc!r}")
        ws.sheet_properties.tabColor = tab_color(ws.title)
        if ws.title in PRESENTATION or STEP_RE.match(ws.title):
            ws.sheet_view.showRowColHeaders = False
        if ws.protection.sheet:
            ws.protection.objects = True
    reorder(wb)


def reorder(wb):
    """Blattreihenfolge wie der Ablauf (P3-02); Druckbereiche sind blattlokal und wandern mit."""
    steps = [n for n in wb.sheetnames if STEP_RE.match(n)]
    want = []
    for n in ORDER:
        if n is None:
            continue
        want.append(n)
        if n == "Leitfaden":
            want.extend(sorted(steps))
    order = [wb[n] for n in want if n in wb.sheetnames]
    order += [ws for ws in wb.worksheets if ws not in order]
    wb._sheets = order
    wb.active = 0
