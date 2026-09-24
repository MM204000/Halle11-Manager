"""Blatt „Cockpit“: drei Spalten B:D | F:H | J:L, 2 × 3 Kacheln, Karten mit gemeinsamer Unterkante,
Prüfhinweise, Diagrammzeile und Seitenkopf/-fuß.

Umgesetzt: P1-13 (Raster B:L, Kacheln, Blöcke, Kopf, Statuspille), P2-18 (Prüfhinweise), P1-05/P1-06
(Ausrichtung, eine Wertkante je Block), P1-01 (Diagrammanker, Diagramm 4 entfernt), P1-17 (Zeilenraster),
P1-18/P1-19 (Kacheln, Formate, Ampel), P2-03 (Tabellenstil), P2-04 (L1/L2), P2-05 (Kopf/Fuß),
P2-06 (Zahlenformate, Datum, Trennzeichen), P2-12 (Vorzeichen/Beschriftungen).

Rechenlogik bleibt unberührt: Es ändern sich nur Stile, Zahlenformate, Zeilenhöhen, Verbünde,
statische Beschriftungen und nicht referenzierte Anzeigeformeln (B5/B6, B49, K15, G26, G27, G35).
Die Spalten N:X (alter Zeitverlauf, Kacheln EKR_Tile/IRR_Tile) bleiben erhalten und sind nur ausgeblendet.
"""
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.worksheet.hyperlink import Hyperlink

from openpyxl.utils import get_column_letter

from core import (ACCENT, CF_YEAR2, AMBER, AMBER_BG, BLUE, GREEN, GREEN_BG, H_BAND, H_ROW, H_TILE_LABEL, INK, INK2, KPI,
                  LINE, LINE2, MUTED, NAVY, NOFILL, NUMFMT, RECHTSFORM_SHORT, RED, RED_BG, T_BODY, T_MICRO,
                  T_SMALL, TINT_XL, WHITE, add_ampel, align, band_l1, col, fill, font, footer, is_formula,
                  iter_cells, kpi_tile, link_target, memo, page_header, safe_merge, set_height, set_text, side,
                  subhead_l2, table_head, total)

SHEET = "Cockpit"
BLOCKS = (("B", "C", "D"), ("F", "G", "H"), ("J", "K", "L"))   # (Beschriftung, Wert, Randspalte)
FIRST, LAST = "B", "L"
NBSP = " "
# Statuslinien wie das Dashboard-Band (nicht in core.py – siehe requests)
RED_LINE, AMBER_LINE, GREEN_LINE = "E7B4AD", "F1CFA5", "B5DCC4"

# --------------------------------------------------------------------------------------------- Inhalte
# Zeilentypen: "row" Datenzeile · "sub" Unterposition · "t1"/"t2"/"t3" Summenstufe · "memo" nachrichtlich
# fmt: Schlüssel aus core.NUMFMT oder ein Formatstring; label None = Beschriftung unverändert lassen.
UPPER = {
    "B": [
        (14, "row", "Objektart", None),
        (15, "row", "Wohn-/Nutzfläche", "qm"),
        (16, "row", "Baujahr", "year"),
        (17, "row", "Kaufpreis", "eur"),
        (18, "sub", "Kaufpreis je m² (Immobilie)", "eur"),
        (19, "row", "+ Kaufnebenkosten", "eur"),
        (20, "sub", "Kaufnebenkostenquote", "pct1"),
        (21, "row", "+ Finanzierungsnebenkosten", "eur"),
        (22, "row", "+ Maßnahmen Jahr 1", "eur"),
        (23, "t3", "= Gesamtinvestition", "eur"),
        (25, "l2", "KAUFPREISAUFTEILUNG (AFA-BEMESSUNGSGRUNDLAGE)", None),
        (26, "row", "Bodenanteil am Kaufpreis", "pct1"),
        (27, "row", "AK Grund und Boden (inkl. Nebenkosten)", "eur"),
        (28, "t2", "= AfA-Basis Gebäude (inkl. Nebenkosten)", "eur"),
        (29, "row", "Bewegliche Gegenstände (eigene AfA)", "eur"),
        (30, "row", "Anteil Erhaltungsrücklage (nicht abschreibbar)", "eur"),
    ],
    "F": [
        (14, "row", "Nettokaltmiete pro Monat (Soll)", "eur"),
        (15, "sub", "Miete je m² und Monat", "eur_qm"),
        (16, "row", "Jahresnettokaltmiete (Soll)", "eur"),
        (17, "row", "Nettokaltmiete Ist Jahr 1 (nach Ausfall)", "eur"),
        (18, "row", "– Bewirtschaftung Jahr 1 (inkl. Rücklagen)", "eur"),
        (19, "sub", "in % der Nettokaltmiete", "pct1"),
        (20, "t2", "= Einnahmenüberschuss (NOI) Jahr 1", "eur"),
        (22, "l2", "RENDITE UND ANNAHMEN", None),
        (23, "row", "Bruttomietrendite (Jahresmiete / Kaufpreis)", "pct1"),
        (24, "row", "Nettomietrendite (NOI / Gesamtinvestition)", "pct1"),
        (25, "row", "Kaufpreisfaktor (Kaufpreis / Jahresmiete)", "mult1"),
        (26, "row", "Mietsteigerung · Kosten · Wert p. a.", None),
        (27, "row", "Mietausfallwagnis · Leerstand Jahr 1", None),
    ],
    "J": [
        (14, "row", "Darlehen gesamt", "eur"),
        (15, "sub", "Darlehen I · Darlehen II", None),
        (16, "row", "Eigenkapital (ohne Reserve)", "eur"),
        (17, "row", "Eigenkapitalquote", "pct1"),
        (18, "row", "Beleihungsauslauf (Darlehen / Kaufpreis)", "pct1"),
        (19, "row", "Gewichteter Sollzins Jahr 1", "pct2"),
        (21, "l2", "KAPITALDIENST UND ZINSRISIKO", None),
        (22, "row", "Kapitaldienst pro Monat (Jahr 1)", "eur"),
        (23, "row", "Kapitaldienst pro Jahr (Jahr 1)", "eur"),
        (24, "row", "Kapitaldienstdeckung (DSCR)", "dscr"),
        (25, "row", "Restschuld nach Zinsbindung (Darlehen I)", "eur"),
        (26, "row", "Volltilgung (Restschuld = 0)", None),
        (27, "row", f"Anschlusszins +1{NBSP}%-Pkt. kostet p. a.", "eur"),
    ],
}
# Steuerwirkung im Cashflow: Erstattung (Wert < 0) erhöht den Cashflow → „+“, Zahlung → „−“
TAX_SIGNED = '"-"#,##0" €";"+"#,##0" €";"–"'
LOWER = {
    "B": [
        (36, "row", "Nettokaltmiete Ist", "eur"),
        (37, "row", "– Bewirtschaftungskosten (inkl. Rücklagen)", "eur"),
        (38, "t1", "= Einnahmenüberschuss (NOI)", "eur"),
        (39, "row", "– Zinsen", "eur"),
        (40, "row", "– Tilgung", "eur"),
        (41, "t2", "= Cashflow vor Steuern", "eur"),
        (42, "row", "± Steuerwirkung (+ Erstattung / – Zahlung)", TAX_SIGNED),
        (43, "t3", "= Cashflow nach Steuern", "eur"),
        (44, "memo", "nachrichtlich: Warmmiete", "eur"),
    ],
    "F": [
        (35, "row", "Rechtsform", None),
        (36, "row", "Steuersatz angewendet (Jahr 1)", "pct2"),
        (37, "row", "AfA-Methode angewendet", None),
        (38, "row", "Reguläre Gebäude-AfA Jahr 1", "eur"),
        (39, "row", f"+ Sonder-AfA §{NBSP}7b Jahr 1", "eur"),
        (40, "row", f"+ Erhöhte Absetzungen §{NBSP}7h/7i Jahr 1", "eur"),
        (41, "row", "+ AfA bewegliche Gegenstände", "eur"),
        (42, "t2", "= Abschreibungen gesamt Jahr 1", "eur"),
        (43, "row", "Erhaltungsaufwand sofort abziehbar Jahr 1", "eur"),
        (44, "row", "Aktivierte Herstellungskosten Jahr 1", "eur"),
        (45, "t1", "= Steuerliches Ergebnis Jahr 1", "eur"),
        (46, "t2", "= Steuerwirkung Jahr 1", "tax_effect"),
        (47, "memo", f"Steuereffekt gewählte AfA vs. Standard (10{NBSP}J.)", "eur"),
    ],
    "J": [
        (35, "row", "Verkauf nach", "years"),
        (36, "row", "Verkaufspreis (Prognose bzw. Eingabe)", "eur"),
        (37, "row", "– Verkaufskosten", "eur"),
        (38, "row", "– Restschuld (Ablösung)", "eur"),
        (39, "sub", "Veräußerungsgewinn (steuerlich)", "eur"),
        (40, "sub", "Steuerpflichtig?", None),
        (41, "row", "– Steuer auf Veräußerungsgewinn", "eur"),
        (42, "t2", "= Nettoerlös nach Steuern und Ablösung", "eur"),
        (43, "row", "+ Kumulierter Cashflow n. St. bis Verkauf", "eur"),
        (44, "row", "– Eingesetztes Eigenkapital", "eur"),
        (45, "t2", "= Gesamtertrag nach Steuern", "eur"),
        (46, "row", "Eigenkapital-Multiple", "mult2"),
        (47, "t3", "= Eigenkapitalrendite p. a. (IRR n. St.)", "pct1"),
    ],
}
# Anzeigeformeln (nicht referenziert): Trennzeichen „ · “ statt „|“, Rechtsform kurz (P1-19, P2-06)
DISPLAY_FORMULAS = {
    "K15": '=FIXED(Darlehen_I,0)&" € · "&IF(Darlehen_II=0,"–",FIXED(Darlehen_II,0)&" €")',
    "G26": '=FIXED(Mietsteigerung*100,1)&" % · "&FIXED(Kostensteigerung*100,1)&" % · "&FIXED(Wertsteigerung*100,1)&" %"',
    "G27": '=FIXED(Mietausfall_Pct*100,1)&" % · "&Leerstand_Monate&" Monate"',
    "G35": "=" + RECHTSFORM_SHORT,
}
# Kontextzeile rechts im Kachelwert (reine Anzeigeformeln in leeren Zellen, gebietsschema-sicher per FIXED)
TILE_SUB = {
    "GI": '="Kaufpreis "&FIXED(Kaufpreis,0)&" € · NK "&FIXED(NK_Quote*100,1)&" %"',
    "EK": '=IF(Reserve_Einmalig=0,"ohne Liquiditätsreserve","davon Reserve "&FIXED(Reserve_Einmalig,0)&" €")',
    "CF": '="ab Jahr 2: "&FIXED(' + CF_YEAR2 + ',0)&" € / Monat"',
    "BMR": '="Faktor "&FIXED(Kaufpreisfaktor,1)&"× · netto "&FIXED($G$24*100,1)&" %"',
    "EKR": "Vermögenszuwachs / Eigenkapital",
    "IRR": '=Haltedauer&" Jahre · Multiple "&FIXED(EK_Multiple,2)&"×"',
}
AMPEL = {"G23": "BMR", "G24": "NMR", "K24": "DSCR", "K47": "IRR"}
NEGATIVE_RED = ("C41:D41", "C43:D43", "K43")   # Ergebniszeilen: negativ rot (relativ je Zelle)
TAX_GREEN = ("C42:D42", "G46")                # Steuerwirkung: Erstattung grün, Zahlung Navy (P2-12)

HINT_ROWS = range(50, 56)
CHART_TOP, CHART_BOTTOM = 58, 73          # Diagramme füllen B58:D73 | F58:H73 | J58:L73
# Soll-Anker (0-basiert: from col/row → to col/row, Offsets 0) in Diagramm-Reihenfolge – für den Nachlauf
# nach LibreOffice (finish_pro), der die Anker in LibreOffice-Spaltenmaß zurückschreibt.
ANCHORS = {SHEET: [(1, CHART_TOP - 1, 4, CHART_BOTTOM), (5, CHART_TOP - 1, 8, CHART_BOTTOM),
                   (9, CHART_TOP - 1, 12, CHART_BOTTOM)]}
FOOTER_ROW = 75


# --------------------------------------------------------------------------------------------- Helfer
def _nf(key):
    return NUMFMT.get(key, key) if key else None


def _hide_cols(ws, first, last):
    """Spalten ausblenden, ohne überlappende <col>-Einträge zu erzeugen (Gruppen wie O:X bleiben eine Gruppe)."""
    lo, hi = col(first), col(last)
    covered = set()
    for key, dim in list(ws.column_dimensions.items()):
        mn, mx = dim.min or col(key), dim.max or col(key)
        if mn >= lo and mx <= hi:
            dim.hidden = True
            covered.update(range(mn, mx + 1))
    for c in range(lo, hi + 1):
        if c not in covered:
            from openpyxl.utils import get_column_letter
            d = ws.column_dimensions[get_column_letter(c)]
            d.min = d.max = c
            d.hidden = True


def _clear_cf(ws, predicate):
    """Bedingte Formate entfernen, deren Bereich predicate(sqref-String) erfüllt."""
    new = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        if predicate(str(cf.sqref)):
            continue
        for rule in cf.rules:
            new.add(str(cf.sqref), rule)
    ws.conditional_formatting = new


def _in_region(sqref):
    from openpyxl.utils.cell import range_boundaries
    for part in sqref.split():
        c1, r1, c2, r2 = range_boundaries(part)
        if r2 >= 5 and r1 <= 80 and c1 <= col("X") and c2 >= col("B"):
            return True
    return False


def _reset(ws, r1, r2, c1=FIRST, c2="M"):
    """Neutraler Grundzustand für den Seitenbereich (entfernt Stilreste der Vorlage wie Inter/F4F0E8)."""
    for c in iter_cells(ws, c1, r1, c2, r2):
        c.font = font(T_BODY, False, INK)
        c.fill = NOFILL
        c.border = Border()
        c.alignment = align("general", "center")
        if c.value is None:
            c.number_format = "General"


def _unmerge_rows(ws, r1, r2, lo="B", hi="L"):
    """Verbünde lösen, die im Seitenbereich B:L beginnen (die ausgeblendeten Kacheln N:X bleiben unberührt)."""
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= r1 and mr.max_row <= r2 and col(lo) <= mr.min_col <= col(hi):
            ws.unmerge_cells(str(mr))


def _link(cell, text, target_sheet, target_cell=None, tooltip=None, size=T_MICRO, h="right", formula=False):
    if formula:
        cell.value = text
    else:
        set_text(cell, text)
    cell.font = font(size, False, BLUE)
    cell.alignment = align(h, "center", 1)
    loc = f"'{target_sheet}'!{target_cell or link_target(target_sheet)}"
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=loc, tooltip=tooltip,
                               display=None if formula else text)


def _label(cell, text, kind):
    if text is not None:
        set_text(cell, text)
    if kind == "sub":
        cell.font = font(T_SMALL, False, MUTED)
        cell.alignment = align("left", "center", 2)
    else:
        cell.font = font(T_BODY, False, INK)
        cell.alignment = align("left", "center", 1)


def _value(ws, row, vcol, ecol, fmt, merge=True):
    if merge:
        safe_merge(ws, vcol, row, ecol, row)
    v = ws.cell(row, col(vcol))
    v.font = font(T_BODY, False, INK)
    v.alignment = align("right", "center", 1)
    if fmt:
        v.number_format = _nf(fmt)
    return v


def _verdict_ref():
    """Zelle der Gesamtbewertung auf dem Dashboard (dashboard.py: Label in Z. R_VERDICT, Status darunter in B)."""
    try:
        import dashboard
        r = int(getattr(dashboard, "R_VERDICT"))
        return f"Dashboard!$B${r + 1}"
    except Exception:
        return "Dashboard!$B$12"


# --------------------------------------------------------------------------------------------- Bereiche
def header(ws):
    """Z. 5 Eyebrow, Z. 6 H1 + Statuspille, Z. 7 Untertitel (P1-13, P2-05, P2-06)."""
    _unmerge_rows(ws, 5, 7)
    for c in iter_cells(ws, "C", 5, "L", 7):
        c.value = None
        c.hyperlink = None
    set_height(ws, 4, 15)
    date = 'TEXT(DAY(Kaufdatum),"00")&"."&TEXT(MONTH(Kaufdatum),"00")&"."&YEAR(Kaufdatum)'
    subtitle = (f'=Obj_Name&" · "&Obj_Adresse&" · Kauf "&{date}&" · Haltedauer "&Haltedauer&" Jahre · "'
                f'&{RECHTSFORM_SHORT}')
    page_header(ws, FIRST, LAST, "Auswertung", "Cockpit", subtitle=subtitle)
    # Statuspille: Gesamtbewertung des Dashboards, klickbar
    cap = ws["K5"]
    set_text(cap, "GESAMTBEWERTUNG")
    cap.font = font(T_MICRO, True, MUTED)
    cap.alignment = align("left", "bottom")
    safe_merge(ws, "K", 6, "L", 6)
    pill = ws["K6"]
    pill.value = f"={_verdict_ref()}"
    pill.number_format = '"●  "@'
    pill.font = font(T_BODY, True, NAVY)
    pill.alignment = align("center", "center")
    pill.hyperlink = Hyperlink(ref="K6", location="'Dashboard'!A4", tooltip="Zur Gesamtbewertung auf dem Dashboard")
    ln = side("thin", LINE2)
    for c in iter_cells(ws, "K", 6, "L", 6):
        c.fill = fill(TINT_XL)
        c.border = Border(top=ln, bottom=ln, left=ln if c.column == col("K") else None,
                          right=ln if c.column == col("L") else None)
    for word, fg, bg, line in (("kritisch", RED, RED_BG, RED_LINE), ("Prüfpunkten", AMBER, AMBER_BG, AMBER_LINE),
                               ("Solide", GREEN, GREEN_BG, GREEN_LINE)):
        cond = f'ISNUMBER(SEARCH("{word}",$K$6))' if word != "Solide" else '$K$6="Solide"'
        b = side("thin", line)
        ws.conditional_formatting.add("K6:L6", FormulaRule(
            formula=[cond], stopIfTrue=True, font=Font(color=fg, bold=True), fill=fill(bg),
            border=Border(left=b, right=b, top=b, bottom=b)))


def tiles(ws):
    """2 × 3 Kacheln im Raster B:D | F:H | J:L (P1-13, P1-18, P1-19)."""
    _unmerge_rows(ws, 8, 11)
    for c in iter_cells(ws, "B", 10, "L", 11):   # alte L1-Bänder aus Z. 11 (statische Texte) räumen
        if not is_formula(c.value):
            c.value = None
    spec = [
        ("B", "D", 8, 9, "GI", None),
        ("F", "H", 8, 9, "EK", None),
        ("J", "L", 8, 9, "CF", None),
        ("B", "D", 10, 11, "BMR", "=Bruttomietrendite"),
        ("F", "H", 10, 11, "EKR", "=EKR_Tile"),
        ("J", "L", 10, 11, "IRR", "=IRR_Tile"),
    ]
    for c1, c2, lr, vr, key, value in spec:
        ws.cell(lr, col(c1)).value = None
        k = KPI[key]
        kpi_tile(ws, c1, c2, lr, vr, label=k["label"], value=value, fmt=k["fmt"], gap_right=False)
        # Wert nur in der Beschriftungsspalte, rechts daneben die Kontextzeile (Fintech-Kachel: Zahl + Einordnung)
        mid = get_column_letter(col(c1) + 1)
        _unmerge_rows(ws, vr, vr, lo=c1, hi=c2)
        safe_merge(ws, mid, vr, c2, vr)
        for c in iter_cells(ws, c1, vr, c2, vr):     # Lösen/Verbinden setzt die Nicht-Anker-Zellen zurück
            c.fill = fill(TINT_XL)
            c.border = Border()
        sub = ws.cell(vr, col(mid))
        sub.value = TILE_SUB[key]
        sub.font = font(T_SMALL, False, MUTED)
        sub.alignment = align("right", "center", 1)
        if k.get("rule"):
            add_ampel(ws, f"{c1}{vr}", key)
    # Fuge zwischen den Kachelreihen: weiße Oberkante der zweiten Label-Zeile (keine Leerzeile verfügbar)
    for c1, c2 in (("B", "D"), ("F", "H"), ("J", "L")):
        for c in iter_cells(ws, c1, 10, c2, 10):
            c.border = Border(top=side("thick", WHITE))
    set_height(ws, 8, H_TILE_LABEL)
    set_height(ws, 10, H_TILE_LABEL)
    set_height(ws, 12, 15)


def _block(ws, rows, lab, val, edge, merge_values=True):
    kinds = {}
    for r, kind, text, fmt in rows:
        kinds[r] = kind
        cell = ws[f"{lab}{r}"]
        if kind == "l2":
            subhead_l2(ws, r, lab, edge, text, height=H_ROW)
            continue
        _label(cell, text, kind)
        single = merge_values and val != edge
        v = _value(ws, r, val, edge, fmt, merge=single and not (lab == "B" and r >= 35))
        if kind in ("t1", "t2", "t3"):
            total(ws, r, lab, edge, level=int(kind[1]))
            cell.alignment = align("left", "center", 1)
            v.alignment = align("right", "center", 1)
        elif kind == "memo":
            memo(ws, r, lab, edge)
            cell.alignment = align("left", "center", 1)
            for c in iter_cells(ws, val, r, edge, r):
                c.alignment = align("right", "center", 1)
        else:
            for c in iter_cells(ws, lab, r, edge, r):
                b = c.border
                c.border = Border(top=b.top, bottom=side("hair", LINE))
    # Die Linie über einer Summe gehört der Summe: keine Haarlinie darunter doppeln
    for r, kind in kinds.items():
        if kinds.get(r + 1) in ("t1", "t2", "t3", "l2"):
            for c in iter_cells(ws, lab, r, edge, r):
                c.border = Border(top=c.border.top)
    return kinds


def _card_bottom(ws, row):
    """Gemeinsame Kartenunterkante; die Doppellinie eines Schlüsselergebnisses bleibt erhalten."""
    for lab, _, edge in BLOCKS:
        for c in iter_cells(ws, lab, row, edge, row):
            b = c.border
            if b.bottom is not None and b.bottom.style == "double":
                continue
            c.border = Border(top=b.top, bottom=side("thin", LINE2))


def blocks(ws):
    """Obere und untere Kartenreihe (P1-13, P1-06, P2-03, P2-04, P2-12, P1-17)."""
    _unmerge_rows(ws, 13, 47)
    # obere Reihe: L1-Bänder in Z. 13 ersetzen die Navy-Köpfe
    band_l1(ws, 13, "B", "D", "Objekt & Investition")
    band_l1(ws, 13, "F", "H", "Miete, Bewirtschaftung & Rendite")
    band_l1(ws, 13, "J", "L", "Finanzierung", right="Zeitverlauf 40 Jahre ›")
    _link(ws["L13"], "Zeitverlauf 40 Jahre ›", "Projektion", tooltip="Alle Kennzahlen über 40 Jahre (Blatt Projektion)")
    for r in range(14, 31):
        set_height(ws, r, H_ROW)
    for lab, val, edge in BLOCKS:
        _block(ws, UPPER[lab], lab, val, edge)
    _card_bottom(ws, 30)
    set_height(ws, 31, 15)

    # untere Reihe
    for (lab, _, edge), title in zip(BLOCKS, ("Cashflow Jahr 1", "Steuern & AfA Jahr 1", "Exit-Szenario nach Haltedauer")):
        band_l1(ws, 32, lab, edge, title)
    for c in iter_cells(ws, "B", 33, "L", 34):      # redundante Navy-Köpfe B34/F34/J34 leeren
        if not is_formula(c.value):
            c.value = None
    ws.row_dimensions[33].hidden = True
    set_height(ws, 34, 6)
    for r in range(35, 48):
        set_height(ws, r, H_ROW)
    table_head(ws, 35, "B", "D", labels={"C": ("PRO MONAT", "right"), "D": ("PRO JAHR", "right")})
    set_height(ws, 35, H_ROW)
    _block(ws, LOWER["B"], "B", "C", "D", merge_values=False)
    for r, kind, _, fmt in LOWER["B"]:          # zweite Wertspalte (pro Jahr) der Ertragsrechnung
        d = ws[f"D{r}"]
        if fmt:
            d.number_format = _nf(fmt)
        d.alignment = align("right", "center", 1)
        if kind not in ("t1", "t2", "t3", "memo"):
            d.font = font(T_BODY, False, INK)
    _block(ws, LOWER["F"], "F", "G", "H")
    _block(ws, LOWER["J"], "J", "K", "L")
    _card_bottom(ws, 47)
    set_height(ws, 48, 15)

    for ref, formula in DISPLAY_FORMULAS.items():
        ws[ref].value = formula


def semantics(ws):
    """Ampeln und Vorzeichenfarben (P1-19, P2-12)."""
    for ref, key in AMPEL.items():
        add_ampel(ws, ref, key)
    for ref in NEGATIVE_RED:
        first = ref.split(":")[0]
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f"{first}<0"], stopIfTrue=True,
                                                       font=Font(color=RED, bold=True)))
    for ref in TAX_GREEN:
        first = ref.split(":")[0]
        ws.conditional_formatting.add(ref, FormulaRule(formula=[f"{first}<0"], stopIfTrue=True,
                                                       font=Font(color=GREEN)))


def hints(ws):
    """Prüfhinweise B49:L56 (P2-18)."""
    _unmerge_rows(ws, 49, 57)
    b49 = ws["B49"]
    set_text(b49, "Prüfhinweise & steuerliche Einordnung")
    band_l1(ws, 49, "B", "L")
    cnt_w = 'COUNTIF($AB$50:$AB$74,"⚠*")'
    cnt_i = 'COUNTIF($AB$50:$AB$74,"ℹ*")'
    safe_merge(ws, "K", 49, "L", 49)       # Zählung rechts im Band (Verbund: mehr Platz, stabil im LibreOffice-Export)
    l49 = ws["K49"]
    l49.value = (f'=IF(COUNT($AC$50:$AC$74)=0,"keine Hinweise",{cnt_w}&IF({cnt_w}=1," Warnung"," Warnungen")'
                 f'&" · "&{cnt_i}&IF({cnt_i}=1," Hinweis"," Hinweise"))')
    l49.font = font(T_MICRO, False, MUTED)
    l49.alignment = align("right", "center", 1)
    for r in HINT_ROWS:
        safe_merge(ws, "B", r, "L", r)
        c = ws[f"B{r}"]
        c.font = font(T_BODY, False, INK2)
        c.alignment = align("left", "center", 1)
        set_height(ws, r, H_ROW)
    ref = f"B{HINT_ROWS.start}:L{HINT_ROWS.stop - 1}"
    top = HINT_ROWS.start
    hl = side("thin", LINE)
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'LEFT($B{top},1)="⚠"'], stopIfTrue=True,
                                                   font=Font(color=RED), fill=fill(RED_BG),
                                                   border=Border(bottom=side("thin", RED_LINE))))
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'$B{top}<>""'], stopIfTrue=True,
                                                   fill=fill(TINT_XL), border=Border(bottom=hl)))
    # weitere Hinweise (Anzeige nur bei mehr als sechs aktiven Hinweisen)
    b56 = ws["B56"]
    _link(b56, '=IF(COUNT($AC$50:$AC$74)>6,"+ "&(COUNT($AC$50:$AC$74)-6)&" weitere Hinweise – Details im Dashboard ›","")',
          "Dashboard", tooltip="Alle aktiven Prüfhinweise auf dem Dashboard", size=T_SMALL, h="left", formula=True)
    set_height(ws, 56, 14)


def charts(ws):
    """Diagrammzeile: drei Diagramme exakt auf B:D | F:H | J:L, das vierte (AfA) nur noch auf „Diagramme“ (P1-01)."""
    _unmerge_rows(ws, 57, 57)
    band_l1(ws, 57, "B", "L", "Diagramme", right="Alle Diagramme ›")
    _link(ws["L57"], "Alle Diagramme ›", "Diagramme", tooltip="Alle Auswertungen auf dem Blatt Diagramme")
    keep = sorted((ch for ch in ws._charts if ch.anchor._from.col < col("N") - 1),
                  key=lambda ch: ch.anchor._from.col)
    ws._charts = [ch for ch in ws._charts if ch in keep]
    for ch, (c1, r1, c2, r2) in zip(keep, ANCHORS[SHEET]):
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = c1, 0, r1, 0
        a.to.col, a.to.colOff, a.to.row, a.to.rowOff = c2, 0, r2, 0
    for r in range(CHART_TOP, CHART_BOTTOM + 1):
        set_height(ws, r, 15)
    set_height(ws, CHART_BOTTOM + 1, 15)


def page_footer(ws):
    _unmerge_rows(ws, FOOTER_ROW, FOOTER_ROW + 1)
    footer(ws, FOOTER_ROW, FIRST, LAST)


def apply(wb):
    if SHEET not in wb.sheetnames:
        return
    ws = wb[SHEET]
    _clear_cf(ws, _in_region)
    _reset(ws, 4, FOOTER_ROW + 1)
    _hide_cols(ws, "N", "X")
    _hide_cols(ws, "Y", "AA")          # leere Randspalten (Hilfsspalten AB:AC sind bereits ausgeblendet)
    header(ws)
    tiles(ws)
    blocks(ws)
    semantics(ws)
    hints(ws)
    charts(ws)
    page_footer(ws)
