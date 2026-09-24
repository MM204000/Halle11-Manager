"""Blatt „Leitfaden“: Seitenkopf mit Start-Aktion, Hinweiszeile, vier Kacheln (EK-Bedarf · Rate · IRR · Cashflow),
Schritt-Checkliste (Status ✓/○ · Schritt · Inhalt, einzeilig), zwei Diagramme in der rechten Kachelspalte,
Button-Reihe (Sekundär links · Primär rechts), Fuß – ausschließlich mit den zentralen Bausteinen aus core.py.

Raster (px): C 56 | D 266 | E 322 | F 322 | G 2 | H 224 | I 98  →  Kacheln C:D | E | F | H:I je 322 px.
Kachel 4 bleibt in H:I (H10 ist Namensziel LF_CFN). G ist eine 3-px-Spalte = die Fuge zwischen den Kacheln.
Die zweite Kachelreihe der Vorlage (Z. 12–13, Namensziele LF_CFV/LF_BMR/LF_DSCR/LF_IRR) wird ausgeblendet
(P3-12: vier Kacheln); ihre Formeln bleiben unverändert.
Nicht ausblenden (Spalten): LibreOffice rechnet Diagrammanker über ausgeblendete Spalten falsch um.
"""
import re

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C
from core import ACCENT, AMBER, BLUE, GREEN, INK2, MUTED, WHITE, align, fill, font, side
from layouts.start import (cells_of, clear_rows, drop_cf, goal_pct, heights, rich, set_widths, status_line,
                           unmerge_in, wipe)

SHEET = "Leitfaden"
WIDTHS = {"C": 8, "D": 38, "E": 46, "F": 46, "G": 0.25, "H": 32, "I": 14}   # G = 3-px-Fuge vor Kachel 4

# Inhalt der Schrittliste (einzeilig in E:F, ≤ 100 Zeichen)
CONTENT = [
    "Objektdaten, Baujahr, Kaufdatum und Bundesland – die Basis für Grunderwerbsteuer und AfA.",
    "Kaufpreis und Nettokaltmiete ergeben die erste Kennzahl: die Bruttomietrendite.",
    "Grunderwerbsteuer, Notar, Grundbuch und Makler – sie erhöhen die AfA-Basis.",
    "Nur der Gebäudeanteil ist abschreibbar – Aufteilung in Grund und Boden und Gebäude.",
    "Renovierung, Sonderumlagen und Liquiditätsreserve – inklusive Prüfung der 15 %-Grenze.",
    "Hausgeld, Verwaltung, Rücklagen und Mietausfall – was von der Miete monatlich übrig bleibt.",
    "Darlehen, Zins, Tilgung und Zinsbindung – Rate, Eigenkapital und Restschuld.",
    "Cashflow vor Steuern: Miete abzüglich Bewirtschaftung, Zinsen und Tilgung.",
    "Kauf als, Einkommen, Soli und Kirchensteuer – daraus folgt der angewendete Steuersatz.",
    "AfA-Methode, Sonder-AfA § 7b und erhöhte Absetzungen § 7h/7i – der steuerliche Hebel.",
    "Miet-, Kosten- und Wertsteigerung, Haltedauer und Verkauf – inklusive 10-Jahres-Frist.",
    "Cashflow nach Steuern, Eigenkapitalrendite, Vermögensaufbau und Rente nach Volltilgung.",
]
FIRST_ROW, N_STEPS = 18, 12


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
    C.page_header(ws, "C", "I", "Leitfaden  ·  In zwölf Schritten zur Investitionsentscheidung", "Leitfaden",
                  "Schritt für Schritt durch Objekt, Kosten, Finanzierung, Steuern und Exit – "
                  "jede Seite mit Zwischenergebnis und Einordnung.")
    # Leisten-Aktion oben rechts (entspricht „WEITER ›“ der Schrittseiten); Titelzeile bleibt 30 pt,
    # weiße Ober-/Unterkante lässt den Button optisch so hoch erscheinen wie alle anderen (≈ 34 px)
    C.btn(ws, "H", 6, "I", "Schritt 01 starten  ›", "S01 Objekt", "primary",
          tooltip="Weiter zu Schritt 01 · Objekt", set_row=False)
    wht = side("thick", WHITE)
    ws["H6"].border = Border(top=wht, bottom=wht)
    ws["I6"].border = Border(top=wht, bottom=wht)
    C.set_height(ws, 6, 30)

    # Hinweiszeile (Eingabe-Konvention, ohne Emoji): Beispielwerte + „Kauf als“ auf der Startseite
    C.safe_merge(ws, "C", 8, "G", 8)
    h = ws["C8"]
    h.value = rich(("Hinweis   ", C.T_SMALL, True, INK2),
                   ("Alle gelben Felder enthalten Beispielwerte – einfach überschreiben. "
                    "„Kauf als“ (Privatperson oder GmbH) wählen Sie auf der Startseite.", C.T_SMALL, False, INK2))
    h.font = font(C.T_SMALL, False, INK2)
    h.alignment = align("left", "center", 1)
    C.safe_merge(ws, "H", 8, "I", 8)
    lk = ws["H8"]
    C.text_link(lk, "Kauf als wählen  ›", "Start", "D23", size=C.T_SMALL, bold=True,
                tooltip="„Kauf als“ auf der Startseite wählen")
    lk.hyperlink.location = C.link_loc("Start", "D23")
    lk.alignment = align("right", "center", 1)
    for c in "CDEFGHI":
        cell = ws[f"{c}8"]
        cell.fill = fill(C.INPUT_BG)
        cell.border = Border(left=side("thick", C.INPUT_LINE) if c == "C" else None)
    heights(ws, {4: 15, 5: 16, 7: 22, 8: 22})


def tiles(ws):
    """Vier Kacheln (C.tile dark): EK-Bedarf · Rate · IRR · Cashflow n. St.; Werte neutral navy,
    Status als „●  Wort“ in der Kontextzeile (Z. 11)."""
    for r in range(9, 14):
        for c in "CDEFGHI":
            ws[f"{c}{r}"].border = Border()
            ws[f"{c}{r}"].fill = C.NOFILL
    drop_cf(ws, cells_of("C", 9, "I", 13))
    unmerge_in(ws, "C", 9, "I", 11)
    for c in "CDEFGHI":                                    # Kontextzeile Z. 11 ist in der Vorlage leer
        ws[f"{c}11"].value = None
    for a in ("D9", "D10", "I9", "I10"):
        ws[a].value = None
    C.tile(ws, "C", "D", 9, 10, 11, gap_top=True, kpi="EK", label="Eigenkapitalbedarf inkl. Reserve",
           sub='="Quote "&FIXED(EK_Quote*100,1)&" %  ·  Beleihung "&FIXED(Beleihung*100,1)&" %"')
    C.tile(ws, "E", "E", 9, 10, 11, gap_top=True, kpi="RATE", sub='="Volltilgung "&Volltilgung_Txt')
    C.tile(ws, "F", "F", 9, 10, 11, gap_top=True, kpi="IRR", value="=EK_IRR", value_ref="EK_IRR", status="dot", gap_right=False,
           sub=status_line("IRR", "EK_IRR", goal_pct("Ampel_IRR_gruen") + '&"  ·  Multiple "&FIXED(EK_Multiple,2)&"×"'))
    lab = ws["F9"]
    lab.value = '="IRR N. ST. · VERKAUF NACH "&Haltedauer&" J."'
    lab.data_type = "f"
    C.tile(ws, "H", "I", 9, 10, 11, gap_top=True, kpi="CF", label="Cashflow n. St. / Monat (Jahr 1)", value_ref="CF_nSt_Monat_J1", status="dot",
           gap_right=False, sub=status_line("CF", "CF_nSt_Monat_J1", f'"Jahr 2: "&FIXED({C.CF_YEAR2},0)&" € / Monat"'))
    # zweite Kachelreihe der Vorlage (Namensziele) ausblenden – Formeln bleiben unverändert
    clear_rows(ws, 12, 13, "B", "I")
    C.hide_rows(ws, 12, 13)
    heights(ws, {14: 6, 15: 6})


def steps(ws):
    """Checkliste: C Status (✓/○, Anzeigeformel aus der Vorlage) · D „01  Schritt ›“ · E:F Inhalt, einzeilig 24 pt."""
    wb = ws.parent
    unmerge_in(ws, "C", 16, "F", 31)
    C.section(ws, 16, "C", "F", "Die zwölf Schritte", meta="")
    lg = ws["F16"]
    lg.value = rich(("✓", C.T_MICRO, True, GREEN), ("  vollständig      ", C.T_MICRO, False, BLUE),
                    ("○", C.T_MICRO, True, AMBER), ("  Angaben fehlen", C.T_MICRO, False, BLUE))
    lg.alignment = align("right", "center", 1)
    for c in "CDEF":
        ws[f"{c}17"].value = None
    C.section(ws, 17, "C", "F", None, level=2, labels={"C": ("Status", "center"), "D": ("Schritt", "left"),
                                                        "E": ("Inhalt", "left")})
    ws["C17"].alignment = align("center", "center")
    for i, (num, name, long) in enumerate(step_sheets(wb)[:N_STEPS]):
        r = FIRST_ROW + i
        st = ws[f"C{r}"]                                   # Status: Formel der Vorlage (F) wandert nach C
        fsrc = ws[f"F{r}"].value
        if C.is_formula(fsrc):
            st.value = fsrc
        st.font = font(C.T_H3, True, MUTED)
        st.alignment = align("center", "center")
        st.border = Border()
        st.number_format = "General"
        d = ws[f"D{r}"]
        d.value = rich((f"{num:02d}", C.T_H3, True, ACCENT), (f"   {long}  ›", C.T_BODY, True, BLUE))
        d.hyperlink = Hyperlink(ref=d.coordinate, location=C.link_loc(name), display=long,
                                tooltip=f"Schritt {num:02d} · {long} öffnen")
        d.font = font(C.T_BODY, True, BLUE)
        d.alignment = align("left", "center", 1)
        d.number_format = "General"
        f = ws[f"F{r}"]
        f.value = None
        C.safe_merge(ws, "E", r, "F", r)
        e = ws[f"E{r}"]
        C.set_text(e, CONTENT[i] if i < len(CONTENT) else e.value)
        e.font = font(C.T_SMALL, False, INK2)
        e.alignment = align("left", "center", 1)
        e.fill = C.NOFILL
        C.hairline(ws, r, "C", "F")
        C.set_height(ws, r, C.H_BAND)
    last = FIRST_ROW + N_STEPS - 1
    drop_cf(ws, cells_of("C", FIRST_ROW, "F", last))
    rng = f"C{FIRST_ROW}:C{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'C{FIRST_ROW}="✓"'],
                                                   font=Font(color=GREEN, bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'C{FIRST_ROW}="○"'],
                                                   font=Font(color=AMBER, bold=True), stopIfTrue=True))


def actions(ws):
    """Button-Reihe (P1-13): Sekundär links (C:D) · Primär rechts (H:I), gleiche Breite und Höhe."""
    wipe(ws, "C", 30, "I", 31)
    C.btn_row(ws, 31, [
        dict(c1="C", c2="D", text="Zum Dashboard  ›", target="Dashboard", kind="secondary",
             tooltip="Gesamtbewertung auf einer Seite"),
        dict(c1="H", c2="I", text="Schritt 01 starten  ›", target="S01 Objekt", kind="primary",
             tooltip="Weiter zu Schritt 01 · Objekt"),
    ])
    heights(ws, {30: C.H_GAP})


def charts(ws):
    """Beide Diagramme exakt in der rechten Kachelspalte H:I, bündig mit der Schrittliste:
    Z. 16–22 (= Kopf + Schritte 01–05) und Z. 23–29 (Schritte 06–12), je 154 pt."""
    chs = sorted(ws._charts, key=lambda ch: ch.anchor._from.row)
    for ch, (r1, r2) in zip(chs, ((16, 22), (23, 29))):
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = 7, 0, r1 - 1, 0
        a.to.col, a.to.colOff, a.to.row, a.to.rowOff = 9, 0, r2, 0


def foot(ws):
    wipe(ws, "C", 32, "I", 60)
    C.set_height(ws, 32, C.H_GAP)
    C.footer(ws, 33, "C", "I")
    for r in range(35, 61):
        C.set_height(ws, r, 15)


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
    C.cf_close(ws)


def apply(wb):
    layout(wb[SHEET])
