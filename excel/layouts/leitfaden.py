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
from core import AMBER, BLUE, GREEN, INK2, MUTED, WHITE, align, font, side
from layouts.start import (TILE_ICON, TILE_SUB, cells_of, clear_rows, drop_cf, heights, icon, rich, section_icon,
                           set_widths, tile_icon, unmerge_in, wipe)

SHEET = "Leitfaden"
# Runde 4: rechte Inhaltskante = S01–S12 (Ende I ≈ 1 417 px) → Kacheln C:D | E | F | H:I je ≈ 346 px
WIDTHS = {"A": 2.5, "B": 2, "C": 8, "D": 41.5, "E": 49.5, "F": 49.5, "G": 0.25, "H": 35.5, "I": 14}   # G = Fuge

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
# Linien-Icon je Schritt (icons.py, Runde 6) – helle Badge GOLD_BG mit Nachtblau-Icon vor „01  Objekt ›“
STEP_ICON = ("haus", "euro", "dokument", "gebaeude", "werkzeug", "rechner", "bank", "diagramm", "paragraf",
             "kalender", "trend", "ziel")
STEP_BADGE = 24
OPT_MARK, OPT_GREY = "◌", C.NEUTRAL_DASH             # Status „optional“ (Schritt 05, P3-11) – Token 9AA3AF


def step_sheets(wb):
    """Schrittblätter in Schrittfolge; der Langname ist der Blattname ohne Präfix (= Titel C6 der Schrittseite)."""
    out = []
    for n in wb.sheetnames:
        m = re.match(r"S(\d\d) (.+)", n)
        if m:
            out.append((int(m.group(1)), n, m.group(2)))
    return sorted(out)


def header(ws):
    """Seitenkopf nach der Kopfschablone (P2-03/P2-04): Brotkrume „LEITFADEN“ (Link auf Start, „‹  Start“ rechts setzt
    global_rules), Titel, rechts Objekt und „Erstellt für“ wie auf S01–S12 (H6:I6 / H7:I7). Der frühere zweite
    „Schritt 01 starten“-Button oben rechts entfällt (P1-03) – die Hauptaktion steht einmal unten rechts (Z. 31).
    Z. 7 links: Hinweis als C.note (P34); Z. 8 bleibt frei (H_GAP)."""
    wipe(ws, "C", 5, "I", 8, formulas=True)              # die Vorlage hat hier keine Formeln
    C.safe_merge(ws, "H", 6, "I", 6)
    C.safe_merge(ws, "H", 7, "I", 7)
    C.page_header(ws, "C", "I", ("Leitfaden",), "Leitfaden", context_col="H",
                  context=("=Obj_Name",
                           '=IFERROR(IF(Erstellt_fuer="",Obj_Adresse,"Erstellt für "&Erstellt_fuer),Obj_Adresse)'))
    # Hinweis: EINE Bauform (core.note) – Text ohne Link, Aktionslink in eigener Zelle rechts (P3-11: Haken ehrlich)
    C.note(ws, 7, "C", "G",
           "Gelbe Felder enthalten Beispielwerte – bitte durch eigene ersetzen. "
           "Haken zeigen nur, ob Felder gefüllt sind.",
           link_text="Kauf als wählen  ›", target_sheet="Start", target_cell="D23", link_col="F",
           tooltip="Pflichtauswahl „Kauf als“ auf der Startseite")
    heights(ws, {8: C.H_GAP})


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
        tile_icon(ws, c1, 9, TILE_ICON[key])
    # zweite Kachelreihe der Vorlage (Namensziele) ausblenden – Formeln bleiben unverändert
    clear_rows(ws, 12, 13, "B", "I")
    C.hide_rows(ws, 12, 13)
    heights(ws, {14: 6, 15: 6})


def steps(ws):
    """Checkliste: C Status (✓/○, Anzeigeformel aus der Vorlage) · D „01  Schritt ›“ · E:F Inhalt, einzeilig 24 pt."""
    wb = ws.parent
    unmerge_in(ws, "C", 16, "F", 31)
    C.section(ws, 16, "C", "F", "Die zwölf Schritte")
    section_icon(ws, 16, "C", "check")
    # Legende der Statussymbole (P3-11: ehrlich – ✓ heißt nur „Angaben vorhanden“), rechtsbündig über E:F
    C.safe_merge(ws, "E", 16, "F", 16)
    lg = ws["E16"]
    lg.value = rich(("✓", C.T_MICRO, True, GREEN), ("  Angaben vorhanden (auch Beispielwerte)      ", C.T_MICRO, False, MUTED),
                    ("○", C.T_MICRO, True, AMBER), ("  Angaben fehlen      ", C.T_MICRO, False, MUTED),
                    (OPT_MARK, C.T_SMALL, True, OPT_GREY), ("  optional", C.T_MICRO, False, MUTED))
    lg.font = font(C.T_MICRO, False, MUTED)
    lg.alignment = align("right", "center", 1)
    for c in "CDEF":
        ws[f"{c}17"].value = None
    C.section(ws, 17, "C", "F", None, level=2, labels={"C": ("Angaben", "center"), "D": ("Schritt", "left"),
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
        d.value = rich((f"{num:02d}", C.T_H3, True, C.GOLD_INK), (f"   {long}  ›", C.T_BODY, True, BLUE))
        d.hyperlink = Hyperlink(ref=d.coordinate, location=C.link_loc(name), display=long,
                                tooltip=f"Schritt {num:02d} · {long} öffnen")
        d.font = font(C.T_BODY, True, BLUE)
        d.alignment = align("left", "center", 5)
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
        if i < len(STEP_ICON):
            icon(ws, f"D{r}", STEP_ICON[i], C.NAVY, STEP_BADGE, dx=8, valign="middle", bg=C.GOLD_BG)
    last = FIRST_ROW + N_STEPS - 1
    # Häkchen-Semantik (P25): ✓ nur für echten Status. Schritt 05 hat keine Pflichtangaben („optional“),
    # Schritt 12 ist vollständig, wenn alle vorherigen Schritte vollständig sind (reine Anzeigeformeln).
    s05, s12 = ws[f"C{FIRST_ROW + 4}"], ws[f"C{last}"]
    C.set_text(s05, OPT_MARK)                           # P3-11: eigenes Symbol statt des Worts „optional“
    s05.font = font(C.T_H3, True, OPT_GREY)
    s05.alignment = align("center", "center")
    s12.value = f'=IF(COUNTIF(C{FIRST_ROW}:C{last - 1},"○")>0,"○","✓")'
    drop_cf(ws, cells_of("C", FIRST_ROW, "F", last))
    rng = f"C{FIRST_ROW}:C{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'C{FIRST_ROW}="✓"'],
                                                   font=Font(color=GREEN, bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'C{FIRST_ROW}="○"'],
                                                   font=Font(color=AMBER, bold=True), stopIfTrue=True))


def actions(ws):
    """Button-Reihe wie auf S01–S12 (P1-03): C:D „‹  Zurück: Start“ (secondary, links) · F „Dashboard“ (tertiary,
    Mitte) · H:I „Schritt 01 starten  ›“ (primary, rechts) – Zurück und Weiter gleich breit, alle 25,5 pt."""
    wipe(ws, "C", 30, "I", 31)
    C.btn_row(ws, 31, [
        dict(c1="C", c2="D", text="‹  Zurück: Start", target="Start", kind="secondary",
             tooltip="Zurück zur Startseite"),
        dict(c1="F", c2="F", text="Dashboard", target="Dashboard", kind="tertiary",
             tooltip="Gesamtbewertung auf einer Seite"),
        dict(c1="H", c2="I", text="Schritt 01 starten  ›", target="S01 Objekt", kind="primary",
             tooltip="Weiter zu Schritt 01 · Objekt"),
    ])
    # Fuge zum Primär-Button: weiße 3-px-Kante rechts am Tertiär-Button (+ Fugenspalte G) – wie die Kachelrinnen.
    # (In E stünde der Tertiär-Button ohne Fuge direkt am Rahmen von „Zurück“ – dort gibt es keine Fugenspalte.)
    t = ws["F31"]
    b = t.border
    t.border = Border(left=b.left, top=b.top, bottom=b.bottom, right=side("thick", WHITE))
    heights(ws, {30: C.H_GAP})


def charts(ws):
    """P3-19: EIN Diagramm in voller Breite der rechten Spalte H:I (Vermögen und Restschuld, 3 Legendeneinträge),
    bündig mit Abschnittskopf (Z. 16) bis Schritt 08 (Z. 25). Das gedrängte Mini-Diagramm „Einnahmen vs. Ausgaben“
    entfällt (steht vollständig auf Dashboard/Diagramme). Darunter die Box „Fortschritt“ (Z. 27–29)."""
    keep = [ch for ch in ws._charts if getattr(ch, "tagname", "") != "barChart"]
    if not keep:                                           # Rückfall: das untere (zweite) Diagramm behalten
        keep = sorted(ws._charts, key=lambda ch: ch.anchor._from.row)[-1:]
    ws._charts = keep[:1]
    for ch in ws._charts:
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = 7, 0, 15, 0
        a.to.col, a.to.colOff, a.to.row, a.to.rowOff = 9, 0, 25, 0


def progress(ws):
    """Box „Fortschritt“ (C.callout_box, neutral) rechts neben Schritt 10–12: Anzahl der Schritte mit Angaben als Pill,
    darunter die nächste Aktion. Reine Anzeigeformeln; die Zeilenhöhen bleiben die der Schrittliste (24 pt)."""
    rng_ = f"$C${FIRST_ROW}:$C${FIRST_ROW + N_STEPS - 1}"
    C.callout_box(ws, "H", 27, "I", 28, 29, title="Fortschritt", fit=None, pill_col="I",
                  pill=f'=COUNTIF({rng_},"✓")&" von "&(COUNTA({rng_})-COUNTIF({rng_},"{OPT_MARK}"))',
                  text=f'=IF(COUNTIF({rng_},"○")=0,"Alle Schritte haben Angaben – Beispielwerte ersetzen, '
                       f'dann Ergebnis in Schritt 12 und im Dashboard prüfen.","Offen: "&COUNTIF({rng_},"○")'
                       f'&" Schritt(e) mit ○ – dort fehlen noch Angaben.")')
    for r in (27, 28, 29):
        C.set_height(ws, r, C.H_BAND)


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
    progress(ws)
    foot(ws)
    ws.sheet_view.showRowColHeaders = False
    C.cf_close(ws)


def apply(wb):
    layout(wb[SHEET])
