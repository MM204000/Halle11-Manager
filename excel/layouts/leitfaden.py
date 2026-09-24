"""Blatt „Leitfaden“: Seitenkopf mit Start-Aktion, Hinweis (C.note), vier Kacheln in kanonischer Reihenfolge
(Gesamtinvestition · EK-Bedarf · Rate · Cashflow n. St. – P20),
Schritt-Checkliste (Status ✓/○ · Schritt · Inhalt, einzeilig), zwei Diagramme in der rechten Kachelspalte,
Button-Reihe (Sekundär links · Primär rechts), Fuß – ausschließlich mit den zentralen Bausteinen aus core.py.

Raster (px): A 18 | B 14 | C 56 | D 266 | E 322 | F 322 | G 2 | H 224 | I 98  →  Kacheln C:D | E | F | H:I je 322 px.
Linke Satzkante (P39): A + B = 4,5 Zeichen wie auf den B-Blättern → Titel/Inhalt beginnen bei x ≈ 31 px.
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
from layouts.start import TILE_SUB, cells_of, clear_rows, drop_cf, heights, rich, set_widths, unmerge_in, wipe

SHEET = "Leitfaden"
WIDTHS = {"A": 2.5, "B": 2, "C": 8, "D": 38, "E": 46, "F": 46, "G": 0.25, "H": 32, "I": 14}   # G = Fuge vor Kachel 4

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
    """Seitenkopf (P37: Dachzeile = Kategorie, nie der Blattname) + Hinweis als C.note in Z. 7 (P34).
    Z. 8 bleibt frei (H_GAP) – dieselbe Luft zwischen Hinweis und Kacheln wie zwischen Kacheln und Abschnittskopf."""
    wipe(ws, "C", 5, "I", 8, formulas=True)              # die Vorlage hat hier keine Formeln
    C.page_header(ws, "C", "I", ("Einstieg", "In zwölf Schritten zur Investitionsentscheidung"), "Leitfaden")
    # Leisten-Aktion oben rechts (entspricht „WEITER ›“ der Schrittseiten); Titelzeile bleibt 30 pt,
    # weiße Ober-/Unterkante lässt den Button optisch so hoch erscheinen wie alle anderen (≈ 34 px)
    C.btn(ws, "H", 6, "I", "Schritt 01 starten  ›", "S01 Objekt", "primary",
          tooltip="Weiter zu Schritt 01 · Objekt", set_row=False)
    wht = side("thick", WHITE)
    ws["H6"].border = Border(top=wht, bottom=wht)
    ws["I6"].border = Border(top=wht, bottom=wht)
    C.set_height(ws, 6, 30)
    # Hinweis: EINE Bauform (core.note) – Text ohne Link, Aktionslink in eigener Zelle rechts
    # Hinweis über den Kachelspalten 1–3 (C:G) – rechts darüber steht der Start-Button, darunter bleibt Weißraum
    C.note(ws, 7, "C", "G",
           "Gelbe Felder enthalten Beispielwerte – einfach überschreiben. „Kauf als“ wählen Sie auf der Startseite.",
           link_text="Kauf als wählen  ›", target_sheet="Start", target_cell="D23", link_col="F",
           tooltip="Pflichtauswahl „Kauf als“ auf der Startseite")
    heights(ws, {4: 15, 5: 16, 7: C.H_PILL, 8: C.H_GAP})


GI_TILES = [  # (c1, c2, KPI, Wertformel) – Teilmenge in der kanonischen Reihenfolge (P20); H10 = Namensziel LF_CFN
    ("C", "D", "GI", "=Gesamtinvestition"),
    ("E", "E", "EK", "=EK_Bedarf_gesamt"),
    ("F", "F", "RATE", "=Kapitaldienst_Monat_J1"),
    ("H", "I", "CF", None),
]


def tiles(ws):
    """Vier Kacheln (C.tile dark, eine Anatomie wie Start/Dashboard): Gesamtinvestition · EK-Bedarf · Rate · Cashflow.
    C10/E10/F10 sind reine Anzeigezellen der Vorlage (nicht referenziert); H10 (LF_CFN) bleibt unverändert."""
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
    for c1, c2, key, val in GI_TILES:
        C.tile(ws, c1, c2, 9, 10, 11, kpi=key, value=val, value_ref=(val or "=CF_nSt_Monat_J1")[1:],
               label=C.kpi_label(key, caps=True), sub=TILE_SUB[key], gap_right=c1 in ("C", "E"))
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
    # Häkchen-Semantik (P25): ✓ nur für echten Status. Schritt 05 hat keine Pflichtangaben („optional“),
    # Schritt 12 ist vollständig, wenn alle vorherigen Schritte vollständig sind (reine Anzeigeformeln).
    s05, s12 = ws[f"C{FIRST_ROW + 4}"], ws[f"C{last}"]
    C.set_text(s05, "optional")
    s05.font = font(C.T_MICRO, False, MUTED)
    s12.value = f'=IF(COUNTIF(C{FIRST_ROW}:C{last - 1},"○")>0,"○","✓")'
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
