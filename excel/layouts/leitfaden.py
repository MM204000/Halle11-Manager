"""Blatt „Leitfaden“: Seitenkopf mit Start-Aktion, Hinweiszeile, vier gleich breite Kacheln, Schrittliste
mit Pflichtangaben-Status, Diagramme exakt in der rechten Kachelspalte, Aktionszeile, Fuß.

Raster: C 8 | D 38 | E 46 | F 46 | G 0,25 | H 32 | I 14  →  Kacheln C:D | E | F | H:I je ≈ 330 px.
Kachel 4 bleibt in H:I (H10/H13 sind Namensziele LF_CFN/LF_IRR und dürfen nicht verbunden/verschoben werden);
G ist eine 3-px-Spalte und damit genau die Fuge wie die weißen Kanten zwischen den übrigen Kacheln.
Nicht ausblenden: LibreOffice rechnet Diagrammanker über ausgeblendete Spalten falsch um.
"""
import re

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font

import core as ui
from core import ACCENT, AMBER, BLUE, DISPLAY, GREEN, INK2, MUTED, WHITE, align, fill, font, side
from layouts.start import drop_cf, heights, link, rich, set_widths, unmerge_in, wipe

SHEET = "Leitfaden"
WIDTHS = {"C": 8, "D": 38, "E": 46, "F": 46, "G": 0.25, "H": 32, "I": 14}   # G = 3-px-Fuge vor Kachel 4

# Inhalt der Schrittliste (höchstens 90 Zeichen → zwei Zeilen à 30 pt)
CONTENT = [
    "Objektdaten, Baujahr, Kaufdatum und Bundesland – Basis für Grunderwerbsteuer und AfA.",
    "Kaufpreis und Nettokaltmiete ergeben die erste Kennzahl: die Bruttomietrendite.",
    "Grunderwerbsteuer, Notar, Grundbuch und Makler – sie erhöhen die AfA-Basis.",
    "Nur der Gebäudeanteil ist abschreibbar – Aufteilung in Grund und Boden und Gebäude.",
    "Renovierung, Sonderumlagen und Liquiditätsreserve – inklusive Prüfung der 15 %-Grenze.",
    "Hausgeld, Verwaltung, Rücklagen, Mietausfall – was von der Miete monatlich übrig bleibt.",
    "Darlehen, Zins, Tilgung und Zinsbindung – Rate, Eigenkapital und Restschuld.",
    "Cashflow vor Steuern: Miete abzüglich Bewirtschaftung, Zinsen und Tilgung.",
    "Rechtsform, Einkommen, Soli und Kirchensteuer – daraus folgt der angewendete Steuersatz.",
    "AfA-Methode, Sonder-AfA § 7b und erhöhte Absetzungen § 7h/7i – der steuerliche Hebel.",
    "Miet-, Kosten- und Wertsteigerung, Haltedauer und Verkauf – inklusive 10-Jahres-Frist.",
    "Cashflow nach Steuern, Eigenkapitalrendite, Vermögensaufbau und Rente nach Volltilgung.",
]
IRR_LABEL = '="IRR n. St. – Verkauf nach "&Haltedauer&" J."'   # wie Start F30 (Anzeigeformel, nicht referenziert)


def step_sheets(wb):
    """Schrittblätter in Schrittfolge; der Langname ist der Blattname ohne Präfix (= Titel C6 der Schrittseite)."""
    out = []
    for n in wb.sheetnames:
        m = re.match(r"S(\d\d) (.+)", n)
        if m:
            out.append((int(m.group(1)), n, m.group(2)))
    return sorted(out)


def header(ws):
    wipe(ws, "C", 5, "I", 8, formulas=True)              # die Vorlage hat hier keine Formeln
    ui.page_header(ws, "C", "I", "Leitfaden  ·  In zwölf Schritten zur Investitionsentscheidung", "Leitfaden",
                   "Schritt für Schritt durch Objekt, Kosten, Finanzierung, Steuern und Exit – "
                   "jede Seite mit Zwischenergebnis und Einordnung.")
    ui.button(ws, "H", 6, "I", "Schritt 01 starten  ›", "S01 Objekt", "primary",
              tooltip="Weiter zu Schritt 01 · Objekt")
    # Titelzeile bleibt 30 pt; weiße Ober-/Unterkante lässt den Button optisch 26 pt hoch erscheinen
    wht = side("thick", WHITE)
    for c in "HI":
        ws[f"{c}6"].border = Border(top=wht, bottom=wht)
    ui.set_height(ws, 6, 30)

    # Hinweiszeile: Beispielwerte + Rechtsform auf der Startseite
    ui.safe_merge(ws, "C", 8, "I", 8)
    h = ws["C8"]
    h.value = rich(("ℹ  Alle gelben Felder enthalten Beispielwerte – einfach überschreiben. "
                    "Die Rechtsform wählen Sie auf der ", 8.5, False, INK2), ("Startseite  ›", 8.5, True, BLUE))
    from openpyxl.worksheet.hyperlink import Hyperlink
    h.hyperlink = Hyperlink(ref="C8", location="'Start'!D23", display="Startseite",
                            tooltip="Rechtsform auf der Startseite wählen")
    h.font = font(8.5, False, INK2)
    h.alignment = align("left", "center", 1)
    for c in "CDEFGHI":
        cell = ws[f"{c}8"]
        cell.fill = fill(ui.INPUT_BG)
        cell.border = Border(left=side("thick", ui.INPUT_LINE) if c == "C" else None, bottom=wht)
    heights(ws, {4: 15, 5: 16, 7: 22, 8: 22})


def tiles(ws):
    for r in range(9, 14):
        for c in "CDEFGHI":
            ws[f"{c}{r}"].border = Border()
    drop_cf(ws, ["H10", "C13", "E13", "F13", "H13"])
    ui.kpi_tile(ws, "C", "D", 9, 10, kpi="EK")
    ui.kpi_tile(ws, "E", "E", 9, 10, kpi="RATE")
    ui.kpi_tile(ws, "F", "F", 9, 10, label="Volltilgung", fmt="@", gap_right=False)
    ui.kpi_tile(ws, "H", "I", 9, 10, kpi="CF")
    ui.kpi_tile(ws, "C", "D", 12, 13, kpi="CFV")
    ui.kpi_tile(ws, "E", "E", 12, 13, kpi="BMR")
    ui.kpi_tile(ws, "F", "F", 12, 13, kpi="DSCR", gap_right=False)
    ws["H12"].value = IRR_LABEL
    ui.kpi_tile(ws, "H", "I", 12, 13, kpi="IRR")
    for c in "BCDEFGHI":                                  # Fugenzeile
        ws[f"{c}11"].fill = ui.NOFILL
        ws[f"{c}11"].border = Border()
    heights(ws, {11: 2.25, 14: 12, 15: 12})


def steps(ws):
    wb = ws.parent
    unmerge_in(ws, "C", 16, "F", 31)
    ui.band_l1(ws, 16, "C", "F", "Die zwölf Schritte")
    lg = ws["F16"]
    lg.value = rich(("✓", ui.T_MICRO, True, GREEN), (" vorhanden   ", ui.T_MICRO, False, MUTED),
                    ("○", ui.T_MICRO, True, AMBER), (" fehlt", ui.T_MICRO, False, MUTED))
    lg.font = font(ui.T_MICRO, False, MUTED)
    lg.alignment = align("right", "center", 1)
    for c in "CDEF":
        ws[f"{c}17"].value = None
    ui.table_head(ws, 17, "C", "F", {"C": ("NR.", "left"), "D": ("SCHRITT", "left"), "E": ("INHALT", "left"),
                                      "F": ("PFLICHTANGABEN", "center")})
    for i, (num, name, long) in enumerate(step_sheets(wb)[:12]):
        r = 18 + i
        n = ws[f"C{r}"]
        ui.set_text(n, f"{num:02d}")
        n.font = font(14, True, ACCENT, DISPLAY)
        n.alignment = align("left", "center", 1)
        n.border = Border()
        d = ws[f"D{r}"]
        link(d, long, name, tooltip=f"Schritt {num:02d} · {long} öffnen")
        d.number_format = '@* "›  "'                       # Pfeile als bündige Spalte an der rechten Kante
        e = ws[f"E{r}"]
        ui.set_text(e, CONTENT[i] if i < len(CONTENT) else e.value)
        e.font = font(9.5, False, INK2)
        e.alignment = align("left", "center", 1, wrap=True)
        f = ws[f"F{r}"]
        f.font = font(11, True, MUTED)
        f.alignment = align("center", "center")
        f.fill = ui.NOFILL
        ui.hairline(ws, r, "C", "F")
        ui.set_height(ws, r, 30)
    drop_cf(ws, [f"F{r}" for r in range(18, 30)])
    ws.conditional_formatting.add("F18:F29", FormulaRule(formula=['F18="✓"'], font=Font(color=GREEN, bold=True),
                                                         stopIfTrue=True))
    ws.conditional_formatting.add("F18:F29", FormulaRule(formula=['F18="○"'], font=Font(color=AMBER, bold=True),
                                                         stopIfTrue=True))


def actions(ws):
    wipe(ws, "C", 30, "I", 31)
    ui.button(ws, "C", 31, "D", "Schritt 01 starten  ›", "S01 Objekt", "primary",
              tooltip="Weiter zu Schritt 01 · Objekt")
    ui.button(ws, "F", 31, "F", "Zum Dashboard  ›", "Dashboard", "secondary",
              tooltip="Gesamtbewertung auf einer Seite")
    heights(ws, {30: 12, 31: ui.H_BUTTON})


def charts(ws):
    """Beide Diagramme exakt in der rechten Kachelspalte H:I: Z. 16–23 und 24–31 (bündig mit der Aktionszeile)."""
    chs = sorted(ws._charts, key=lambda ch: ch.anchor._from.row)
    for ch, (r1, r2) in zip(chs, ((16, 23), (24, 31))):
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = 7, 0, r1 - 1, 0
        a.to.col, a.to.colOff, a.to.row, a.to.rowOff = 9, 0, r2, 0


def foot(ws):
    wipe(ws, "C", 32, "I", 60)
    ui.set_height(ws, 32, 18)
    ui.footer(ws, 33, "C", "I")
    for r in range(35, 61):
        ui.set_height(ws, r, 15)


def layout(ws):
    set_widths(ws, WIDTHS)
    ws.column_dimensions["G"].hidden = False
    header(ws)
    tiles(ws)
    steps(ws)
    actions(ws)
    charts(ws)
    foot(ws)
    ws.sheet_view.showRowColHeaders = False


def apply(wb):
    layout(wb[SHEET])
