"""Blatt „Cockpit“: Seitenkopf mit Gesamtbewertung, sechs KPI-Kacheln (drei Paare „Betrag | Rendite“ im Raster
B | C:D · F | G:H · J | K:L), zwei Kartenreihen, Prüfhinweise, Diagrammzeile und Fuß.

Runde 2 – durchgängig auf die zentralen Bausteine aus core.py umgestellt (CORE_API.md):
  C.page_header (Seitenkopf P1-18) · C.tile (Kachel P1-11/P1-10, dreizeilig, Wert neutral, Status-Chip)
  C.section (Abschnittskopf Ebene 1/2, P1-12) · C.sum_row (Summenstufen P1-14) · C.status_cf / C.neg_red (P1-10)
  C.NUMFMT (Zahlenformat-Katalog P2-12/P3-08) · C.link_loc / C.link_row (Sprungziele P1-08) · C.cf_close.
Weitere Planpunkte: P2-14 (Prüfhinweise ohne Emoji, Semantik über Farbe/Kante), P3-02 (Unterzeilen 9 pt grau,
Abstände, Drill-down-Links je Blockkopf), P3-03 („↑ Übersicht“), P3-08 (Steuerwirkungs-Konvention).

Rechenlogik bleibt unberührt: Es ändern sich nur Stile, Zahlenformate, Zeilenhöhen, Verbünde, statische
Beschriftungen und nicht referenzierte Anzeigeformeln (B5/B6, B42, B49:B55, C9/G9/K9, K15, G26, G27, G35 …).
Die benannten Zellen (u. a. J9 = CF_nSt_Monat_J1) bleiben, wo sie sind. Die Spalten N:X (alter Zeitverlauf,
Kacheln EKR_Tile/IRR_Tile) bleiben erhalten und sind nur ausgeblendet.
"""
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C
from core import (ACCENT, BLUE, INK, INK2, LINE, LINE2, MIST, MUTED, NAVY, NOFILL, NUMFMT, RECHTSFORM_SHORT, RED,
                  RED_BG, RED_LINE, T_BODY, T_MICRO, T_SMALL, TINT_XL, WHITE, align, col, fill, font, is_formula,
                  iter_cells, safe_merge, set_height, set_text, side)

SHEET = "Cockpit"
BLOCKS = (("B", "C", "D"), ("F", "G", "H"), ("J", "K", "L"))   # (Beschriftung, Wert, Randspalte)
FIRST, LAST = "B", "L"
NBSP = " "

# --------------------------------------------------------------------------------------------- Inhalte
# Zeilentypen: "row" Datenzeile · "sub" Unterposition (Label und Wert 9 pt grau) · "sum" Zwischensumme (Stufe 1) ·
# "result" Blockergebnis (Stufe 2) · "memo" nachrichtlich · "l2" Unterabschnitt.
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
        (23, "result", "= Gesamtinvestition", "eur"),
        (25, "l2", "Kaufpreisaufteilung (AfA-Bemessungsgrundlage)", None),
        (26, "row", "Bodenanteil am Kaufpreis", "pct1"),
        (27, "row", "AK Grund und Boden (inkl. Nebenkosten)", "eur"),
        (28, "sum", "= AfA-Basis Gebäude (inkl. Nebenkosten)", "eur"),
        (29, "row", "Bewegliche Gegenstände (eigene AfA)", "eur"),
        (30, "row", "Erhaltungsrücklage (nicht abschreibbar)", "eur"),
    ],
    "F": [
        (14, "row", "Nettokaltmiete pro Monat (Soll)", "eur"),
        (15, "sub", "Miete je m² und Monat", "eur_qm"),
        (16, "row", "Jahresnettokaltmiete (Soll)", "eur"),
        (17, "row", "Nettokaltmiete Ist Jahr 1 (nach Ausfall)", "eur"),
        (18, "row", "– Bewirtschaftung Jahr 1 (inkl. Rücklagen)", "eur"),
        (19, "sub", "in % der Nettokaltmiete", "pct1"),
        (20, "result", "= Einnahmenüberschuss (NOI) Jahr 1", "eur"),
        (22, "l2", "Rendite und Annahmen", None),
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
        (21, "l2", "Kapitaldienst und Zinsrisiko", None),
        (22, "row", "Kapitaldienst pro Monat (Jahr 1)", "eur"),
        (23, "row", "Kapitaldienst pro Jahr (Jahr 1)", "eur"),
        (24, "row", "Kapitaldienstdeckung (DSCR)", "dscr"),
        (25, "row", "Restschuld nach Zinsbindung (Darlehen I)", "eur"),
        (26, "row", "Volltilgung (Restschuld = 0)", None),
        (27, "row", f"Anschlusszins +1{NBSP}%-Pkt. kostet p. a.", "eur"),
    ],
}
# Steuerzeile im Cashflow (P3-08): Das Vorzeichen steht wie bei „– Zinsen“ in der Beschriftung (dynamisch, wie
# S09 H16), die Beträge erscheinen ohne Vorzeichen. Erstattung (Wert < 0) erhöht den Cashflow → „+ Steuererstattung“.
TAX_ABS = '#,##0" €";#,##0" €";"–"'
TAX_LABEL = '=IF(D42<0,"+ Steuererstattung",IF(D42>0,"– Steuerzahlung","± Steuerwirkung"))'
LOWER = {
    "B": [
        (36, "row", "Nettokaltmiete Ist", "eur"),
        (37, "row", "– Bewirtschaftungskosten (inkl. Rücklagen)", "eur"),
        (38, "sum", "= Einnahmenüberschuss (NOI)", "eur"),
        (39, "row", "– Zinsen", "eur"),
        (40, "row", "– Tilgung", "eur"),
        (41, "sum", "= Cashflow vor Steuern", "eur"),
        (42, "row", TAX_LABEL, TAX_ABS),
        (43, "result", "= Cashflow nach Steuern", "eur"),
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
        (42, "sum", "= Abschreibungen gesamt Jahr 1", "eur"),
        (43, "row", "Erhaltungsaufwand sofort abziehbar Jahr 1", "eur"),
        (44, "row", "Aktivierte Herstellungskosten Jahr 1", "eur"),
        (45, "sum", "= Steuerliches Ergebnis Jahr 1", "eur"),
        (46, "result", "= Steuerwirkung Jahr 1", "tax_effect"),
        (47, "memo", f"Steuereffekt gewählte AfA vs. Standard (10{NBSP}J.)", "eur"),
    ],
    "J": [
        (35, "row", "Verkauf nach", "years_n"),
        (36, "row", "Verkaufspreis (Prognose bzw. Eingabe)", "eur"),
        (37, "row", "– Verkaufskosten", "eur"),
        (38, "row", "– Restschuld (Ablösung)", "eur"),
        (39, "sub", "Veräußerungsgewinn (steuerlich)", "eur"),
        (40, "sub", "Steuerpflichtig?", None),
        (41, "row", "– Steuer auf Veräußerungsgewinn", "eur"),
        (42, "sum", "= Nettoerlös nach Steuern und Ablösung", "eur"),
        (43, "row", "+ Kumulierter Cashflow n. St. bis Verkauf", "eur"),
        (44, "row", "– Eingesetztes Eigenkapital", "eur"),
        (45, "sum", "= Gesamtertrag nach Steuern", "eur"),
        (46, "row", "Eigenkapital-Multiple", "mult2"),
        (47, "result", "= Eigenkapitalrendite p. a. (IRR n. St.)", "pct1"),
    ],
}
# Anzeigeformeln (nicht referenziert): Trennzeichen „ · “ statt „|“, Rechtsform kurz (P1-19, P2-06)
DISPLAY_FORMULAS = {
    "K15": '=FIXED(Darlehen_I,0)&" € · "&IF(Darlehen_II=0,"–",FIXED(Darlehen_II,0)&" €")',
    "G26": '=FIXED(Mietsteigerung*100,1)&" % · "&FIXED(Kostensteigerung*100,1)&" % · "&FIXED(Wertsteigerung*100,1)&" %"',
    "G27": '=FIXED(Mietausfall_Pct*100,1)&" % · "&Leerstand_Monate&" Monate"',
    "G35": "=" + RECHTSFORM_SHORT,
}


def _goal(name):
    return f'="Ziel ≥ "&FIXED({name}*100,1)&" %"'


# Kacheln (P1-11): drei Paare „Betrag | Rendite“ – große Kachel in der Beschriftungsspalte (B/F/J, Wert = Formel der
# Vorlage in Z. 9, J9 ist der Name CF_nSt_Monat_J1), kleine Kachel über die beiden Wertspalten (Anzeigeformel).
# (c1, c2, kpi, value, Kontextzeile, Status)
TILES = [
    ("B", "B", "GI", None, '="Kaufpreis "&FIXED(Kaufpreis,0)&" € · Nebenkosten "&FIXED(NK_Quote*100,1)&" %"', None),
    ("C", "D", "BMR", "=Bruttomietrendite", _goal("Ampel_BMR_gruen"), "chip"),
    ("F", "F", "EK", None,
     '=IF(Reserve_Einmalig=0,"ohne Liquiditätsreserve","davon Liquiditätsreserve "&FIXED(Reserve_Einmalig,0)&" €")',
     None),
    ("G", "H", "EKR", "=EKR_Tile", _goal("Ampel_EKR_gruen"), "chip"),
    ("J", "J", "CF", None, '="●  ab Jahr 2: "&FIXED(' + C.CF_YEAR2 + ',0)&" € / Monat"', "dot"),
    ("K", "L", "IRR", "=IRR_Tile", _goal("Ampel_IRR_gruen"), "chip"),
]
TILE_LABEL = {"CF": "Cashflow n. St. / Monat (Jahr 1)", "EK": "Eigenkapitalbedarf inkl. Reserve",
              "EKR": "EK-Rendite Jahr 1", "IRR": "IRR n. St. (Haltedauer)"}
AMPEL = {"G23": "BMR", "G24": "NMR", "K24": "DSCR", "K47": "IRR"}   # Tabellenwerte 10 pt: Status als Schriftfarbe
NEGATIVE_RED = ("C41:D41", "C43:D43", "K43", "K45")           # echte negative Ergebnisse rot (P1-10)

HINT_ROWS = range(50, 56)
CHART_TOP, CHART_BOTTOM = 58, 73          # Diagramme füllen B58:D73 | F58:H73 | J58:L73
CHART_ROW = 16                            # Rasterhöhe der Diagrammzeilen (core.ROW_RASTER)
# Soll-Anker (0-basiert: from col/row → to col/row, Offsets 0) in Diagramm-Reihenfolge.
ANCHORS = {SHEET: [(1, CHART_TOP - 1, 4, CHART_BOTTOM), (5, CHART_TOP - 1, 8, CHART_BOTTOM),
                   (9, CHART_TOP - 1, 12, CHART_BOTTOM)]}
FOOTER_ROW = 75
GAP_SECTION = C.H_ROW                     # Weißraum zwischen den großen Abschnitten (Karten, Hinweise, Diagramme)


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
    """Zell-Link im Stil der Abschnitts-Meta (8 pt 1D4F8A, rechtsbündig) – Ziel über core.link_loc."""
    if formula:
        cell.value = text
    else:
        set_text(cell, text)
    cell.font = font(size, False, BLUE)
    cell.alignment = align(h, "center", 1)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=C.link_loc(target_sheet, target_cell), tooltip=tooltip,
                               display=None if formula else text)


def _up_link(cell):
    _link(cell, "↑ Übersicht", SHEET, C.link_target(SHEET), tooltip="Zum Seitenanfang (Kennzahlen)")


def _verdict_ref():
    """Zelle der Gesamtbewertung auf dem Dashboard (dashboard.py: Label in Z. R_VERDICT, Status darunter in B)."""
    try:
        import dashboard
        r = int(getattr(dashboard, "R_VERDICT"))
        return f"Dashboard!$B${r + 1}"
    except Exception:
        return "Dashboard!$B$12"


# --------------------------------------------------------------------------------------------- Seitenkopf
def header(ws):
    """Seitenkopf-Standard (P1-18): Z. 5 Überzeile, Z. 6 H1, Z. 7 Untertitel links; rechts die Gesamtbewertung
    (Label · Status-Pille mit Link auf die Prüfhinweise) und „Erstellt für …“ als Meta (9 pt)."""
    _unmerge_rows(ws, 5, 7)
    for c in iter_cells(ws, "C", 5, "L", 7):
        c.value = None
        c.hyperlink = None
    set_height(ws, 4, 15)
    date = 'TEXT(DAY(Kaufdatum),"00")&"."&TEXT(MONTH(Kaufdatum),"00")&"."&YEAR(Kaufdatum)'
    subtitle = (f'=Obj_Name&"  ·  "&Obj_Adresse&"  ·  Kauf am "&{date}&"  ·  Haltedauer "&Haltedauer&" Jahre  ·  "'
                f'&{RECHTSFORM_SHORT}')
    C.page_header(ws, FIRST, LAST, "Cockpit", "Cockpit", subtitle=subtitle)
    ws["B6"].alignment = align("left", "bottom")
    safe_merge(ws, "B", 7, "J", 7)
    # rechts: Gesamtbewertung (Label Z. 5, Pille Z. 6, Meta Z. 7) – Breite K:L = eine Wertspalte der Karten
    safe_merge(ws, "K", 5, "L", 5)
    cap = ws["K5"]
    set_text(cap, "GESAMTBEWERTUNG")
    cap.font = font(T_MICRO, True, BLUE)
    cap.alignment = align("left", "bottom")
    safe_merge(ws, "K", 6, "L", 6)
    pill = ws["K6"]
    pill.value = f"={_verdict_ref()}"
    pill.number_format = NUMFMT["status_dot"]
    pill.font = font(T_BODY, True, NAVY)
    pill.alignment = align("left", "center", 1)
    pill.hyperlink = Hyperlink(ref="K6", location=C.link_loc(SHEET, C.link_row(SHEET, 49)),
                               tooltip="Zu den Prüfhinweisen und der steuerlichen Einordnung")
    ln = side("thin", MIST)
    for c in iter_cells(ws, "K", 6, "L", 6):
        c.fill = fill(TINT_XL)
        c.border = Border(top=ln, bottom=ln, left=side("thick", ACCENT) if c.column == col("K") else None,
                          right=ln if c.column == col("L") else None)
    conds = [('ISNUMBER(SEARCH("kritisch",$K$6))', "red"), ('ISNUMBER(SEARCH("Prüfpunkten",$K$6))', "amber"),
             ('ISNUMBER(SEARCH("Solide",$K$6))', "green")]
    C.status_cf(ws, "K6:L6", conds, font_color=True, bold=True, fill_bg=True)
    C.status_edge(ws, "K6", conditions=conds)
    safe_merge(ws, "K", 7, "L", 7)
    meta = ws["K7"]
    meta.value = '=IFERROR(IF(Erstellt_fuer="","","Erstellt für "&Erstellt_fuer),"")'
    meta.font = font(T_SMALL, False, MUTED)
    meta.alignment = align("right", "top", 1)


# --------------------------------------------------------------------------------------------- Kacheln
def tiles(ws):
    """Sechs Kacheln in einer Reihe (P1-11): Label 18 · Wert 30 · Kontext 16 pt, Werte neutral (P1-10),
    Status als Chip (Renditen) bzw. als Kontextzeile in Statusfarbe (Cashflow). Paare „Betrag | Rendite“ sitzen
    exakt auf den Spalten der Karten darunter; zwischen den Paaren liegt die Rinnenspalte E/I."""
    _unmerge_rows(ws, 8, 12)
    keep = {"B9", "F9", "J9"}                        # Formeln der Vorlage (J9 = Name CF_nSt_Monat_J1)
    for c in iter_cells(ws, "B", 8, "L", 12):
        if c.coordinate not in keep:
            c.value = None
            c.hyperlink = None
    for c1, c2, key, value, sub, status in TILES:
        C.tile(ws, c1, c2, 8, 9, 10, kpi=key, label=TILE_LABEL.get(key), value=value, sub=sub,
               status=status, gap_right=(c1 == c2), value_ref=f"${c1}$9")
    set_height(ws, 11, 6)       # Fuge Kacheln → Karten: 2 × 6 pt = eine Leerzeile H_GAP
    set_height(ws, 12, 6)


# --------------------------------------------------------------------------------------------- Karten
def _label(cell, text, kind):
    if text is not None:
        if text.startswith("=IF("):          # dynamische Beschriftung (Anzeigeformel)
            cell.value = text
        else:                                # statischer Text, auch „= Summe …“
            set_text(cell, text)
    if kind == "sub":
        cell.font = font(T_SMALL, False, MUTED)
        cell.alignment = align("left", "center", 2)
    else:
        cell.font = font(T_BODY, False, INK)
        cell.alignment = align("left", "center", 1)


def _value(ws, row, vcol, ecol, fmt, kind, merge=True):
    if merge:
        safe_merge(ws, vcol, row, ecol, row)
    v = ws.cell(row, col(vcol))
    v.font = font(T_SMALL, False, MUTED) if kind == "sub" else font(T_BODY, False, INK)
    v.alignment = align("right", "center", 1)
    if fmt:
        v.number_format = _nf(fmt)
    return v


def _block(ws, rows, lab, val, edge, merge_values=True):
    kinds = {}
    for r, kind, text, fmt in rows:
        kinds[r] = kind
        cell = ws[f"{lab}{r}"]
        if kind == "l2":
            C.section(ws, r, lab, edge, text, level=2, variant="line", height=C.H_ROW)
            continue
        _label(cell, text, kind)
        v = _value(ws, r, val, edge, fmt, kind, merge=merge_values and val != edge)
        if kind in ("sum", "result"):
            C.sum_row(ws, r, lab, edge, "sub" if kind == "sum" else "result")
            cell.alignment = align("left", "center", 1)
            for c in iter_cells(ws, val, r, edge, r):
                c.alignment = align("right", "center", 1)
        elif kind == "memo":
            C.memo(ws, r, lab, edge)
            cell.alignment = align("left", "center", 1)
            for c in iter_cells(ws, val, r, edge, r):
                c.alignment = align("right", "center", 1)
        else:
            C.hairline(ws, r, lab, edge)
    # Die Linie über einer Summe gehört der Summe: keine Haarlinie darunter doppeln
    for r, kind in kinds.items():
        if kinds.get(r + 1) in ("sum", "result", "l2"):
            for c in iter_cells(ws, lab, r, edge, r):
                c.border = Border(top=c.border.top)
    return kinds


def _card_bottom(ws, row):
    """Gemeinsame Kartenunterkante; die Doppellinie eines Blockergebnisses bleibt erhalten."""
    for lab, _, edge in BLOCKS:
        for c in iter_cells(ws, lab, row, edge, row):
            b = c.border
            if b.bottom is not None and b.bottom.style == "double":
                continue
            c.border = Border(top=b.top, bottom=side("thin", LINE2))


def _block_head(ws, row, lab, edge, title, link):
    """Blockkopf Ebene 1 mit Drill-down-Link rechts (P3-02)."""
    C.section(ws, row, lab, edge, title, level=1)
    text, target, tip = link
    _link(ws[f"{edge}{row}"], text, target, tooltip=tip)


def blocks(ws):
    """Obere und untere Kartenreihe (P1-12 Köpfe, P1-14 Summenstufen, P3-02 Unterzeilen/Abstände/Drill-down)."""
    _unmerge_rows(ws, 13, 47)
    for c in iter_cells(ws, "B", 13, "L", 13):
        c.value = None
    heads_up = (("B", "D", "Objekt & Investition", ("Eingaben ›", "Eingaben", "Alle Annahmen auf einen Blick")),
                ("F", "H", "Miete, Bewirtschaftung & Rendite",
                 ("Projektion ›", "Projektion", "Miete, Kosten und Cashflow über die Jahre")),
                ("J", "L", "Finanzierung", ("Finanzierung ›", "Finanzierung", "Tilgungsplan beider Darlehen")))
    for lab, edge, title, link in heads_up:
        _block_head(ws, 13, lab, edge, title, link)
    for r in range(14, 31):
        set_height(ws, r, C.H_ROW)
    for lab, val, edge in BLOCKS:
        _block(ws, UPPER[lab], lab, val, edge)
    _card_bottom(ws, 30)
    set_height(ws, 31, GAP_SECTION)

    # untere Reihe – Kopf direkt über der ersten Zeile wie oben (Z. 33/34 ausgeblendet, P3-02)
    for c in iter_cells(ws, "B", 32, "L", 34):
        if not is_formula(c.value):
            c.value = None
    heads_low = (("B", "D", "Cashflow Jahr 1", ("Projektion ›", "Projektion", "Cashflow über die Haltedauer")),
                 ("F", "H", "Steuern & AfA Jahr 1", ("Steuern ›", "Steuern", "Steuerliche Ermittlung je Jahr")),
                 ("J", "L", "Exit-Szenario nach Haltedauer",
                  ("Sensitivität ›", "Sensitivität", "Was-wäre-wenn: Verkaufspreis, Zins, Miete")))
    for lab, edge, title, link in heads_low:
        _block_head(ws, 32, lab, edge, title, link)
    for r in (33, 34):
        ws.row_dimensions[r].hidden = True
        set_height(ws, r, 3)
    for r in range(35, 48):
        set_height(ws, r, C.H_ROW)
    C.section(ws, 35, "B", "D", None, level=2, labels={"C": ("Pro Monat", "right"), "D": ("Pro Jahr", "right")},
              height=C.H_ROW)
    ws["B35"].value = None
    _block(ws, LOWER["B"], "B", "C", "D", merge_values=False)
    for r, kind, _, fmt in LOWER["B"]:          # zweite Wertspalte (pro Jahr) der Ertragsrechnung
        d = ws[f"D{r}"]
        if fmt:
            d.number_format = _nf(fmt)
        d.alignment = align("right", "center", 1)
        if kind == "row":
            d.font = font(T_BODY, False, INK)
    _block(ws, LOWER["F"], "F", "G", "H")
    _block(ws, LOWER["J"], "J", "K", "L")
    _card_bottom(ws, 47)
    set_height(ws, 48, GAP_SECTION)

    for ref, formula in DISPLAY_FORMULAS.items():
        ws[ref].value = formula


def semantics(ws):
    """Status nur semantisch (P1-10): Zielkennzahlen in den Karten per Schriftfarbe, negative Ergebnisse rot."""
    for ref, key in AMPEL.items():
        C.add_ampel(ws, ref, key)
    for ref in NEGATIVE_RED:
        C.neg_red(ws, ref)


# --------------------------------------------------------------------------------------------- Prüfhinweise
# Anzeigeformeln B50:B55 (nicht referenziert): farbige Emoji der Hinweistexte → monochrome Zeichen (P2-14).
# ▲ = Warnung, ⓘ = Information; die Semantik tragen Fläche, Kante und Schriftfarbe (bedingte Formatierung).
HINT_FORMULA = ('=IFERROR(SUBSTITUTE(SUBSTITUTE(INDEX($AB$50:$AB$74,SMALL($AC$50:$AC$74,{k})),'
                '"⚠","▲",1),"ℹ","ⓘ",1),"")')


def hints(ws):
    """Prüfhinweise B49:L56 (P2-14, P3-03)."""
    _unmerge_rows(ws, 49, 57)
    for c in iter_cells(ws, "C", 49, "L", 49):
        c.value = None
    C.section(ws, 49, "B", "L", "Prüfhinweise & steuerliche Einordnung", level=1)
    cnt_w = 'COUNTIF($AB$50:$AB$74,"⚠*")'
    cnt_i = 'COUNTIF($AB$50:$AB$74,"ℹ*")'
    j49 = ws["J49"]
    j49.value = (f'=IF(COUNT($AC$50:$AC$74)=0,"keine Hinweise",{cnt_w}&IF({cnt_w}=1," Warnung"," Warnungen")'
                 f'&"  ·  "&{cnt_i}&IF({cnt_i}=1," Hinweis"," Hinweise"))')
    j49.font = font(T_MICRO, False, MUTED)
    j49.alignment = align("right", "center", 1)
    _up_link(ws["L49"])
    white = side("thin", WHITE)
    for k, r in enumerate(HINT_ROWS, start=1):
        ws[f"B{r}"].value = HINT_FORMULA.format(k=k)
        safe_merge(ws, "B", r, "L", r)
        c = ws[f"B{r}"]
        c.font = font(T_BODY, False, INK2)
        c.alignment = align("left", "center", 1)
        for cc in iter_cells(ws, "B", r, "L", r):
            cc.border = Border(bottom=white)
        set_height(ws, r, C.H_ROW)
    top, bot = HINT_ROWS.start, HINT_ROWS.stop - 1
    anchors = f"B{top}:B{bot}"
    warn, info = f'LEFT($B{top},1)="▲"', f'$B{top}<>""'
    # Fläche/Schrift auf der Ankerzelle (gilt für den ganzen Verbund), Kante links in Statusfarbe
    ws.conditional_formatting.add(anchors, FormulaRule(
        formula=[warn], stopIfTrue=True, font=Font(color=RED), fill=fill(RED_BG),
        border=Border(left=side("thick", RED), bottom=white)))
    ws.conditional_formatting.add(anchors, FormulaRule(
        formula=[info], stopIfTrue=True, font=Font(color=BLUE), fill=fill(TINT_XL),
        border=Border(left=side("thick", ACCENT), bottom=white)))
    # weitere Hinweise (Anzeige nur bei mehr als sechs aktiven Hinweisen)
    _link(ws["B56"], '=IF(COUNT($AC$50:$AC$74)>6,"+ "&(COUNT($AC$50:$AC$74)-6)&" weitere Hinweise – alle auf dem '
          'Dashboard ›","")', "Dashboard", tooltip="Alle aktiven Prüfhinweise auf dem Dashboard", size=T_SMALL,
          h="left", formula=True)
    set_height(ws, 56, GAP_SECTION)


# --------------------------------------------------------------------------------------------- Diagramme
def charts(ws):
    """Diagrammzeile: drei Diagramme exakt auf B:D | F:H | J:L, das vierte (AfA) nur noch auf „Diagramme“."""
    _unmerge_rows(ws, 57, 57)
    for c in iter_cells(ws, "B", 57, "L", 57):
        c.value = None
    C.section(ws, 57, "B", "L", "Diagramme", level=1)
    _link(ws["K57"], "Alle Diagramme ›", "Diagramme", tooltip="Alle Auswertungen auf dem Blatt Diagramme")
    _up_link(ws["L57"])
    keep = sorted((ch for ch in ws._charts if ch.anchor._from.col < col("N") - 1),
                  key=lambda ch: ch.anchor._from.col)
    ws._charts = [ch for ch in ws._charts if ch in keep]
    for ch, (c1, r1, c2, r2) in zip(keep, ANCHORS[SHEET]):
        a = ch.anchor
        a._from.col, a._from.colOff, a._from.row, a._from.rowOff = c1, 0, r1, 0
        a.to.col, a.to.colOff, a.to.row, a.to.rowOff = c2, 0, r2, 0
    for r in range(CHART_TOP, CHART_BOTTOM + 1):
        set_height(ws, r, CHART_ROW)
    set_height(ws, CHART_BOTTOM + 1, C.H_GAP)


def page_footer(ws):
    _unmerge_rows(ws, FOOTER_ROW, FOOTER_ROW + 1)
    C.footer(ws, FOOTER_ROW, FIRST, LAST)


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
    C.cf_close(ws)
