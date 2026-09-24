"""Leitfaden-Seiten S01–S12: Spaltenraster, Zeilenraster, Eingabetabellen, Einheiten, Ergebnis-Panel,
Einordnung, Eingabezustände, Kacheln (S08/S12), Button-Zeile, Zonen-Layout und Diagrammanker.

Nur Darstellung. Die Rechenformeln der Vorlage bleiben unverändert. Angepasst werden ausschließlich
nicht referenzierte Anzeigeformeln (Einordnungstexte, Textergebnisse, Kachel-Beschriftungen), statische
Beschriftungen, Stile, Zahlenformate, Zeilenhöhen/Spaltenbreiten, Verbünde, Links und Diagrammanker.

Zonen einer Schrittseite
  A  ab Z. 10: links die Eingabetabelle (C:F), rechts das Ergebnis dieses Schritts (H:I) mit der Einordnung
     direkt darunter.
  B  Diagrammband eine Leerzeile unter der Eingabetabelle (Kopf im L2-Stil, Diagramm genau im Raster).
  C  zwei Leerzeilen (12 pt) unter dem tiefsten Element die Button-Zeile, eine Leerzeile darunter der Fuß.
"""
import math
import re

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C

# ============================================================================ Raster und Namen
GRID = (("C", 33), ("D", 32), ("E", 13), ("F", 52), ("G", 3), ("H", 38), ("I", 16))
LONG = ["Objekt", "Kaufpreis & Miete", "Kaufnebenkosten", "Kaufpreisaufteilung", "Maßnahmen & Reserve",
        "Bewirtschaftung", "Finanzierung", "Zwischenergebnis", "Steuern", "Abschreibung", "Prognose & Exit",
        "Ergebnis"]
GAP = 12                 # Leerzeile zwischen Zonen
CHART_H = 210            # Diagrammhöhe (pt) im Diagrammband
ROW1, ROW2, ROW3 = C.H_STEP_ROW, C.H_STEP_ROW2, 44
LINE_PT = 13             # Zeilenhöhe Fließtext 9 pt in Einordnungskästen
NBSP = " "

UNITS = {"€": None, "%": None, "€/Monat": "pro Monat", "€ p. a.": "pro Jahr", "€/m²": "je m²",
         "€/m² p. a.": "je m² p. a.", "% Darlehen": "vom Darlehen", "% der Miete": "der Miete",
         "% Verkaufspreis": "vom Verkaufspreis"}

PX = {}                  # Spaltenbreiten in Excel-Pixeln (nach dem Setzen des Rasters)


def status_ampel(value, green, yellow, target):
    return (f'=IF({value}>={green},"Grün",IF({value}>={yellow},"Gelb","Rot"))&"  ·  Ziel ≥ "&{target}')


# Ampel-Status für die Einordnung: (KPI-Schlüssel für add_ampel, Wertausdruck, Statusformel)
STATUS = {
    "BMR": ("BMR", "Bruttomietrendite",
            status_ampel("Bruttomietrendite", "Ampel_BMR_gruen", "Ampel_BMR_gelb", 'FIXED(Ampel_BMR_gruen*100,1)&" %"')),
    "NMR": ("NMR", "S06_NMR",
            status_ampel("S06_NMR", "Ampel_NMR_gruen", "Ampel_NMR_gelb", 'FIXED(Ampel_NMR_gruen*100,1)&" %"')),
    "DSCR": ("DSCR", "DSCR_J1",
             status_ampel("DSCR_J1", "Ampel_DSCR_gruen", "Ampel_DSCR_gelb", "FIXED(Ampel_DSCR_gruen,2)")),
    "IRR": ("IRR", "EK_IRR",
            status_ampel("EK_IRR", "Ampel_IRR_gruen", "Ampel_IRR_gelb", 'FIXED(Ampel_IRR_gruen*100,1)&" %"')),
    "CFV": (None, "CF_vSt_J1", '=IF(CF_vSt_J1<0,"Rot","Grün")&"  ·  Ziel ≥ 0 €"'),
}

# ============================================================================ Blatt-Spezifikation
# results: Zeile → (Rolle, Beschriftung|None, Formel|None, Format|None)
#   Rollen: n normal · s Unterposition · i Zwischensumme · f Endergebnis · m nachrichtlich · t Textwert
SPEC = {
    1: dict(
        inputs=range(12, 20), text=(12, 13, 14, 19),
        results={11: ("n", None, None, None), 12: ("n", None, None, None), 13: ("n", None, None, None),
                 14: ("n", None, '=IF(Wohngebaeude=1,"Ja","Nein (Gewerbe)")', None)},
        callout=dict(head=16, src="H17", max_chars=235),
        info=(21, "Gelb hinterlegte Felder sind Eingaben mit Beispielwerten – einfach überschreiben. "
                  "Die Rechtsform (Privat oder GmbH) wählen Sie auf der Startseite."),
        nav=24, foot=26),
    2: dict(
        inputs=range(12, 18), indent2=(13, 15),
        labels={13: "davon bewegliche Gegenstände (Küche, Möbel)"},
        hints={13: "gesondert im Kaufvertrag ausgewiesen: grunderwerbsteuerfrei (§ 2 GrEStG) und linear "
                   "abschreibbar"},
        results={11: ("n", "Kaufpreis ohne Inventar/Rücklage", None, None),
                 12: ("n", None, None, None), 13: ("n", None, None, None),
                 14: ("n", "Kaufpreisfaktor", "=Kaufpreisfaktor", C.NUMFMT["mult1"]),
                 15: ("f", "= Bruttomietrendite", None, C.NUMFMT["pct1"]),
                 16: ("m", "nachrichtlich: Warmmiete pro Monat", "=Miete_Monat+NK_umlagefaehig", C.NUMFMT["eur"])},
        ampel={"I15": "BMR"},
        callout=dict(head=18, src="H19", max_chars=240, status="BMR"),
        nav=26, foot=28),
    3: dict(
        inputs=range(12, 17), optional=(12,),
        labels={12: "Grunderwerbsteuer abweichend", 16: "Sonstige Erwerbsnebenkosten"},
        hints={12: "optional · leer = Satz des Bundeslands (siehe Ergebnis)",
               16: "Gutachten, Fahrtkosten, Übersetzungen – Anschaffungsnebenkosten, sofern objektbezogen"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None), 13: ("n", None, None, None),
                 14: ("n", None, None, None), 15: ("n", None, None, None), 16: ("i", None, None, None),
                 17: ("s", None, None, C.NUMFMT["pct1"]), 18: ("f", None, None, None)},
        callout=dict(head=20, src="H21", max_chars=245),
        charts=[dict(side="L", title="Zusammensetzung der Kaufnebenkosten", unit="Anteile in %")],
        nav=44, foot=46, hide=(60, 66)),
    4: dict(
        inputs=range(12, 16), text=(12,),
        inactive={14: "KPA_Idx<>2", 15: "KPA_Idx<>3"},
        hints={14: "nur Methode „Bodenanteil in %“", 15: "nur Methode „Gebäudewert“ (ohne bewegliche Gegenstände)"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None),
                 13: ("n", None, None, C.NUMFMT["pct1"]),
                 14: ("n", "AK Grund und Boden (inkl. NK)", None, None),
                 15: ("f", "= AfA-Basis Gebäude (inkl. NK)", None, None),
                 16: ("m", "nachrichtlich: bewegliche Gegenstände", None, None)},
        callout=dict(head=18, src="H19", max_chars=305),
        charts=[dict(side="L", title="Kaufpreisaufteilung", unit="Anteile in %")],
        nav=43, foot=45),
    5: dict(
        inputs=range(12, 20), text=(16,),
        inactive={17: "OR(Rechtsform_Idx>=2,Wohngebaeude=0)"},
        labels={12: "Renovierung Jahr 1 (brutto)", 13: "Renovierung Jahr 2 (brutto)",
                14: "Renovierung Jahr 3 (brutto)", 15: f"Sonderumlagen der WEG (Jahr{NBSP}1)",
                17: "Verteilung Erhaltungsaufwand",
                19: "Einmalige Liquiditätsreserve"},
        hints={15: "Instandsetzungs-Umlagen: Erhaltungsaufwand, zählen zur 15 %-Grenze",
               17: f"§{NBSP}82b EStDV · 1 = Sofortabzug, 2–5 = Verteilung (nur Privat, Wohngebäude)",
               19: "z. B. Mieterwechsel · nicht steuerrelevant, erhöht nur das Eigenkapital"},
        units={17: "Jahr(e)"},
        results={11: ("n", "Maßnahmen Jahre 1–3 (netto)", None, None),
                 12: ("n", None, None, None),
                 13: ("t", "Steuerliche Behandlung",
                      '=IF(HK_Flag=1,"Aktiviert (AfA)",IF(Verteilung_eff>1,"Verteilt (§ 82b)","Sofortabzug"))', None),
                 14: ("n", "Maßnahmen + Sonderumlage Jahr 1", None, None),
                 15: ("i", None, None, None), 16: ("f", None, None, None)},
        callout=dict(head=18, src="H19", max_chars=300),
        charts=[dict(side="L", title="Zusammensetzung der Gesamtinvestition", unit="Anteile in %")],
        nav=44, foot=46),
    6: dict(
        inputs=range(12, 20), indent2=(13, 14),
        labels={13: "davon nicht umlagefähig",
                14: "davon Zuführung zur Erhaltungsrücklage", 19: "Sonstige Werbungskosten"},
        hints={13: "Verwaltung WEG, Rücklage, Sonstiges – trägt der Eigentümer",
               14: "in „nicht umlagefähig“ enthalten · steuerlich erst bei Verausgabung durch die WEG abziehbar "
                   "(BFH IX R 19/24)",
               19: "Steuerberatung, Kontoführung, Fahrten"},
        results={11: ("n", "Nettokaltmiete Ist (nach Ausfall)", None, None),
                 12: ("n", "– Bewirtschaftung inkl. Rücklagen", None, None),
                 13: ("s", "in % der Nettokaltmiete", None, C.NUMFMT["pct1"]),
                 14: ("i", "= Einnahmenüberschuss (NOI)", None, None),
                 15: ("f", "= Nettomietrendite", None, C.NUMFMT["pct1"])},
        ampel={"I15": "NMR"}, band_right="Jahr 1 · pro Monat",
        callout=dict(head=17, src="H18", max_chars=275, status="NMR"),
        charts=[dict(side="L", title="Bewirtschaftungskosten Jahr 1", unit="€ p. a. · Anteile")],
        nav=41, foot=43, hide=(60, 66)),
    7: dict(
        inputs=range(12, 22), note=22,
        labels={16: "Darlehen I – Anschlusszins nach Zinsbindung"},
        hints={16: "Annahme · Zinsänderungsrisiko; die Annuität bleibt konstant"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None), 13: ("i", None, None, None),
                 14: ("n", "Eigenkapitalquote", "=EK_Quote", C.NUMFMT["pct1"]),
                 15: ("n", "Beleihungsauslauf", "=Beleihung", C.NUMFMT["pct1"]),
                 16: ("n", "Restschuld nach Zinsbindung", None, None),
                 17: ("n", None, None, None),
                 18: ("f", "= Rate an die Bank / Monat", "=Kapitaldienst_Monat_J1", C.NUMFMT["eur"])},
        callout=dict(head=20, src="H20", body=21, max_chars=265, status="DSCR"),
        charts=[dict(side="L", title="Restschuld am Jahresende", unit="T€ · Darlehen I und II", chart_col=7),
                dict(side="R", title="Finanzierungsstruktur", unit="Anteile in %", chart_col=2)],
        nav=44, foot=46),
    9: dict(
        inputs=range(12, 19), text=(12, 13, 16, 17), linked=12, optional=(15,),
        inactive={13: "Rechtsform_Idx>=2", 14: "Rechtsform_Idx>=2", 15: "Rechtsform_Idx>=2",
                  16: "Rechtsform_Idx>=2", 17: "Rechtsform_Idx>=2", 18: "Rechtsform_Idx=1"},
        labels={13: "Veranlagung", 14: "Zu versteuerndes Einkommen ohne dieses Objekt",
                15: "Grenzsteuersatz manuell", 16: "Solidaritätszuschlag berücksichtigen",
                17: "Kirchensteuer", 18: "Gewerbesteuer-Hebesatz"},
        hints={12: "Rechtsform auf der Startseite ändern  ›", 13: "nur Privatperson",
               14: "nur Privatperson · Basis für den Grenzsteuersatz nach Tarif 2026",
               15: "nur Privatperson · optional, leer = automatisch nach Tarif 2026 (z. B. 42 % oder 45 %)",
               16: "nur Privatperson · „Nein“, wenn die ESt unter der Freigrenze bleibt (2026: 20.350 € / "
                   "40.700 €)",
               17: "nur Privatperson · der Sonderausgabenabzug der KiSt ist in der Belastung berücksichtigt",
               18: "nur GmbH ohne erweiterte Kürzung · z. B. 400 % → Eingabe 400 %"},
        results={11: ("n", None, None, None),
                 12: ("n", "Grenzbelastung inkl. Soli / KiSt", None, None),
                 13: ("n", "Steuersatz GmbH Jahr 1", None, None),
                 14: ("i", None, None, None), 15: ("n", None, None, None),
                 16: ("f", '="= Steuer"&IF(I16<0,"erstattung","zahlung")&" Jahr 1"', None,
                      '#,##0" €";#,##0" €";"–"')},
        band_right="Jahr 1",
        callout=dict(head=18, src="H19", max_chars=310),
        nav=27, foot=29),
    10: dict(
        inputs=range(12, 18), text=(12, 14, 15, 16),
        inactive={13: "AfA_Idx<>6", 14: "AfA_Idx<>5", 17: 'Denkmal<>"Ja"'},
        labels={13: "Restnutzungsdauer lt. Gutachten", 14: "Degressiv: Auto-Wechsel zur linearen AfA",
                15: f"Sonder-AfA §{NBSP}7b (Mietwohnungsneubau)",
                16: f"Erhöhte Absetzungen §{NBSP}7h / §{NBSP}7i EStG"},
        hints={13: "nur Methode Gutachten · kürzere Nutzungsdauer per Sachverständigengutachten (BFH IX R 25/19)",
               14: f"nur degressive AfA · Wechsel zur linearen Restwert-AfA, sobald vorteilhaft (§{NBSP}7 Abs. 5a "
                   "S. 4 EStG)",
               15: "5 % p. a. für 4 Jahre zusätzlich; Bauantrag 1.1.2023 – 30.9.2029, EH 40 + QNG, "
                   "AK ≤ 5.200 €/m²",
               16: "Sanierungsgebiet / Baudenkmal: 9 % p. a. (Jahre 1–8) und 7 % p. a. (Jahre 9–12) auf die "
                   "bescheinigten Kosten",
               17: f"nur bei §{NBSP}7h / §{NBSP}7i · lt. Bescheinigung der Denkmal- bzw. Gemeindebehörde"},
        results={11: ("n", "Angewendete AfA-Methode",
                      '=IF(Methode_eff=5,"Degressiv 5 %",IF(Methode_eff=6,"Gutachten (RND)","Linear "&'
                      'FIXED(CHOOSE(Methode_eff,AfA_Satz_auto,0.02,0.025,0.03)*100,1)&" %"))', None),
                 12: ("n", None, None, None), 13: ("n", f"+ Sonder-AfA §{NBSP}7b", None, None),
                 14: ("n", f"+ Erhöhte Absetzungen §{NBSP}7h/7i", None, None),
                 15: ("n", None, None, None), 16: ("f", None, None, None),
                 17: ("m", f"Steuereffekt ggü. Standard-AfA (10{NBSP}J.)", None, None)},
        band_right="Jahr 1",
        callout=dict(head=19, src="H20", max_chars=310),
        charts=[dict(side="L", title="Abschreibungen nach Komponenten", unit="Jahre 1–15 · € p. a.")],
        nav=45, foot=47),
    11: dict(
        inputs=range(12, 19), text=(18,), optional=(17,),
        inactive={18: "Rechtsform_Idx>=2"},
        labels={16: "Verkaufskosten", 17: "Verkaufspreis manuell"},
        units={16: None},
        hints={16: "vom Verkaufspreis · Makler, Notar, Löschung der Grundschuld", 17: "optional · leer = Wertentwicklung lt. Prognose"},
        results={11: ("n", "Verkaufspreis bei Exit", None, None),
                 12: ("n", None, None, None),
                 13: ("n", "Verkauf steuerpflichtig?", '=IF(Exit_steuerpflichtig=1,"Ja","Nein")', None),
                 14: ("n", None, None, None),
                 15: ("i", "= Nettoerlös nach Steuer und Ablösung", None, None),
                 16: ("i", None, None, None),
                 17: ("f", "= Eigenkapitalrendite (IRR n. St.)", None, C.NUMFMT["pct1"])},
        ampel={"I17": "IRR"}, band_right="Verkauf",
        callout=dict(head=19, src="H20", max_chars=240, status="IRR"),
        charts=[dict(side="L", title="Immobilienwert, Restschuld und Nettovermögen", unit="T€ · Jahresende")],
        nav=45, foot=47),
}


# ============================================================================ Messen
def nlines(text, width_px, size=C.T_BODY, bold=False, indent=1):
    """Zeilenzahl bei Wortumbruch (Calibri-Näherung aus core.text_px, 5 % Reserve)."""
    if text is None or text == "":
        return 1
    usable = width_px - 6 - 9 * indent
    total = 0
    for para in str(text).split("\n"):
        line, n = 0.0, 1
        for word in para.split(" "):
            w = C.text_px(word, size, bold) * 0.96  # text_px liegt gemessen 5–13 % über Calibri
            sp = C.text_px(" ", size, bold)
            if line and line + sp + w > usable:
                n += 1
                line = w
            else:
                line = line + (sp if line else 0) + w
        total += n
    return total


def px(*cols):
    return sum(PX[c] for c in cols)


def est_lines(chars, width_px=None, size=C.T_SMALL):
    """Zeilenzahl eines Fließtexts mit höchstens `chars` Zeichen (Calibri-Mittelwert aus core.text_px,
    dazu 4 % Reserve für den Wortumbruch)."""
    width_px = width_px or px("H", "I")
    avg = C.text_px("Die Maßnahmen bleiben unter der Grenze – Sofortabzug ist möglich.", size) / 66
    return max(1, math.ceil(chars * avg * 1.04 / (width_px - 6 - 9)))


def height_for(lines):
    return ROW1 if lines <= 1 else (ROW2 if lines == 2 else ROW3)


# ============================================================================ Hilfen
def blank(cell):
    cell.value = None
    cell.hyperlink = None
    cell.font = C.font()
    cell.fill = C.NOFILL
    cell.border = Border()
    cell.alignment = Alignment()
    cell.number_format = "General"


def unmerge_from(ws, first_row):
    for mr in list(ws.merged_cells.ranges):
        if mr.max_row >= first_row and mr.min_col >= 3:
            ws.unmerge_cells(str(mr))


def reset_styles(ws, r1, r2, c1="C", c2="L"):
    for c in C.iter_cells(ws, c1, r1, c2, r2):
        c.font = C.font()
        c.fill = C.NOFILL
        c.border = Border()
        c.alignment = Alignment(vertical="center")


def drop_cf(ws, coords):
    """Bedingte Formate entfernen, deren Bereich eine der Zellen enthält (werden neu gesetzt)."""
    coords = set(coords)
    new = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        cells = set()
        for rg in cf.sqref.ranges:
            for row in range(rg.min_row, rg.max_row + 1):
                for cc in range(rg.min_col, rg.max_col + 1):
                    cells.add(f"{C.L(cc)}{row}")
        if cells & coords:
            continue
        for rule in cf.rules:
            new.add(str(cf.sqref), rule)
    ws.conditional_formatting = new


def tint(ws, ref, key):
    """Ampel-Tönung (nur Füllung) für Einordnungskästen – dieselben Schwellen wie der Leitwert."""
    if key == "CFV":
        rules = [("CF_vSt_J1<0", C.RED_BG), ("ISNUMBER(CF_vSt_J1)", C.GREEN_BG)]
    else:
        kpi, v, _ = STATUS[key]
        spec = C.KPI[kpi]
        rules = [(f"AND(ISNUMBER({v}),{v}<{spec['yellow']})", C.RED_BG),
                 (f"AND(ISNUMBER({v}),{v}<{spec['green']})", C.AMBER_BG),
                 (f"ISNUMBER({v})", C.GREEN_BG)]
    for formula, bg in rules:
        ws.conditional_formatting.add(ref, FormulaRule(formula=[formula], fill=C.fill(bg), stopIfTrue=True))


def status_cell(ws, cell, key):
    """Statuszeile rechts im Kopf der Einordnung (ohne Wertung, in Ampelfarbe)."""
    kpi, v, formula = STATUS[key]
    cell.value = formula
    cell.font = C.font(C.T_MICRO, True, C.MUTED)
    cell.alignment = C.align("right", "center", 1)
    if kpi:  # mit Tönung, sonst überdeckt die Schriftregel (stopIfTrue) die Kasten-Tönung
        C.add_ampel(ws, cell.coordinate, kpi, value_ref=v, with_fill=True)
    else:
        ws.conditional_formatting.add(cell.coordinate, FormulaRule(formula=[f"{v}<0"], font=Font(color=C.RED, bold=True),
                                                                   fill=C.fill(C.RED_BG), stopIfTrue=True))
        ws.conditional_formatting.add(cell.coordinate, FormulaRule(formula=[f"ISNUMBER({v})"],
                                                                   font=Font(color=C.GREEN, bold=True),
                                                                   fill=C.fill(C.GREEN_BG), stopIfTrue=True))


def callout_box(ws, c1, head_row, c2, r1, r2, title, status=None, head=True):
    """Einordnung nach core.callout. head=False: nur der Körper (Kopf ist ein L1-Band, S08/S12)."""
    if head:
        C.callout(ws, c1, head_row, c2, r1, r2, title=title)
    else:
        for r in range(r1, r2 + 1):
            for c in C.iter_cells(ws, c1, r, c2, r):
                c.border = Border(left=C.side("thick", C.ACCENT) if c.column == C.col(c1) else None)
                c.fill = C.fill(C.TINT_XL)
        C.safe_merge(ws, c1, r1, c2, r2)
    b = ws.cell(r1, C.col(c1))
    b.font = C.font(C.T_SMALL, False, C.INK2)
    b.alignment = C.align("left", "center", 1, wrap=True)
    if status:
        status_cell(ws, ws.cell(head_row, C.col(c2)), status)
        tint(ws, C.rng(c1, head_row if head else r1, c2, r2), status)


def move_formula(ws, src, dst):
    """Nicht referenzierte Anzeigeformel an eine neue Zelle setzen (Formeltext bleibt gleich)."""
    if src == dst:
        return ws[dst].value
    v = ws[src].value
    ws[dst].value = v
    blank(ws[src])
    return v


def l2_head(ws, row, c1, c2, title, unit=None, height=C.H_HEAD):
    C.subhead_l2(ws, row, c1, c2, title, height=height)
    if unit:
        u = ws.cell(row, C.col(c2))
        C.set_text(u, unit)
        u.font = C.font(C.T_MICRO, False, C.BLUE)
        u.alignment = C.align("right", "center", 1)


def l2_head_inline(ws, row, c1, c2, title, unit, height=C.H_HEAD):
    """L2-Kopf für schmale Panels: Einheit direkt hinter dem Titel (grau), damit nichts abgeschnitten wird."""
    C.subhead_l2(ws, row, c1, c2, None, height=height)
    cell = ws.cell(row, C.col(c1))
    cell.value = CellRichText([TextBlock(InlineFont(rFont=C.SANS, sz=C.T_SMALL, b=True, color=C.BLUE), title),
                               TextBlock(InlineFont(rFont=C.SANS, sz=C.T_MICRO, color=C.MUTED), f"  ·  {unit}")])
    for c in C.iter_cells(ws, C.col(c1) + 1, row, c2, row):
        c.value = None


def set_anchor(chart, c1, r1, c2, r2):
    """Diagramm genau in C1:R1 … C2:R2 (Zellgrenzen, ohne Versatz)."""
    a = chart.anchor
    a._from.col, a._from.colOff, a._from.row, a._from.rowOff = C.col(c1) - 1, 0, r1 - 1, 0
    a.to.col, a.to.colOff, a.to.row, a.to.rowOff = C.col(c2), 0, r2, 0


def rich_link(cell, word, rest, target_sheet, target_cell=None, tooltip=None):
    """Listen-Link: Link-Wort fett Blau, Beschreibung grau, Pfeil am Ende."""
    blue = InlineFont(rFont=C.SANS, sz=C.T_BODY, b=True, color=C.BLUE)
    grey = InlineFont(rFont=C.SANS, sz=C.T_BODY, color=C.MUTED)
    cell.value = CellRichText([TextBlock(blue, word), TextBlock(grey, f" – {rest}"), TextBlock(blue, "  ›")])
    cell.font = C.font(C.T_BODY, True, C.BLUE)
    cell.alignment = C.align("left", "center", 1)
    loc = f"'{target_sheet}'!{target_cell or C.link_target(target_sheet)}"
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=loc, display=f"{word} – {rest}", tooltip=tooltip)


def row_line(ws, row, c1, c2):
    for c in C.iter_cells(ws, c1, row, c2, row):
        b = c.border
        c.border = Border(left=b.left, right=b.right, top=b.top, bottom=C.side("hair", C.LINE))


# ============================================================================ Bausteine der Seite
def grid(ws):
    for letter, w in GRID:
        ws.column_dimensions[letter].width = w
    for letter, _ in GRID:
        PX[letter] = C.col_px(ws, letter)


def header(ws, n):
    title = ws["C6"].value if isinstance(ws["C6"].value, str) else LONG[n - 1]
    C.page_header(ws, "C", "I", f"Schritt {n:02d} / 12  ·  {LONG[n - 1]}", title,
                  context=("=Obj_Name", '=Obj_Adresse&IFERROR(IF(Erstellt_fuer="","","  ·  Erstellt für "&Erstellt_fuer),"")'),
                  context_col="H")
    C.safe_merge(ws, "C", 7, "F", 7)
    C.safe_merge(ws, "H", 6, "I", 6)
    C.safe_merge(ws, "H", 7, "I", 7)


def nav_and_footer(ws, n, row, names):
    """Button-Zeile (Zurück sekundär · Übersicht als Textlink · Weiter primär) und Fuß."""
    set_h = ws.row_dimensions
    set_h[row - 2].height = GAP
    set_h[row - 1].height = GAP
    if n == 1:
        prev_sheet, prev_text, prev_tip = "Leitfaden", "‹  Zurück: Leitfaden", "Zurück zur Übersicht aller Schritte"
    else:
        prev_sheet = names[n - 1]
        prev_text = f"‹  Zurück: {LONG[n - 2]}"
        prev_tip = f"Zurück zu Schritt {n - 1:02d} · {LONG[n - 2]}"
    if n == 12:
        next_sheet, next_text, next_tip = "Dashboard", "Weiter: Dashboard  ›", "Weiter zum Dashboard · Gesamtbewertung"
    else:
        next_sheet = names[n + 1]
        next_text = f"Weiter: {LONG[n]}  ›"
        next_tip = f"Weiter zu Schritt {n + 1:02d} · {LONG[n]}"
    C.button(ws, "C", row, "C", prev_text, prev_sheet, kind="secondary", tooltip=prev_tip)
    if n != 1:
        C.button(ws, "E", row, "F", "Übersicht: Leitfaden  ›", "Leitfaden", kind="link",
                 tooltip="Alle zwölf Schritte im Überblick")
        lk = ws.cell(row, 5)
        lk.font = C.font(C.T_BODY, True, C.BLUE)
        lk.alignment = C.align("center", "center")
    C.button(ws, "H", row, "I", next_text, next_sheet, kind="primary", tooltip=next_tip)
    set_h[row].height = C.H_BUTTON
    set_h[row + 1].height = GAP
    C.footer(ws, row + 2, "C", "I")
    for c in C.iter_cells(ws, "C", row + 2, "I", row + 2):
        c.border = Border(top=C.side("hair", C.LINE))
    return row + 4


def clear_rows_after(ws, first, last):
    for r in range(first, last + 1):
        for c in C.iter_cells(ws, "C", r, "L", r):
            if c.data_type == "f":
                print(f"  steps: {ws.title}!{c.coordinate} enthält eine Formel unter dem Fuß – bleibt stehen")
                continue
            if c.value is not None or c.hyperlink is not None:
                blank(c)
        ws.row_dimensions[r].height = None


def style_input_table(ws, n, sp):
    """Kopfzeile Z. 11 und Eingabezeilen: Beschriftung, Eingabe, Einheit, Hinweis; gibt Zeilenbedarf zurück."""
    need = {}
    numeric_head = n != 1
    C.table_head(ws, 11, "C", "F", {"C": ("POSITION", "left"), "D": ("EINGABE", "right" if numeric_head else "left"),
                                    "E": ("EINHEIT", "left"), "F": ("HINWEIS", "left")})
    text_rows = set(sp.get("text", ()))
    for r in sp["inputs"]:
        lab, inp, unit, hint = (ws.cell(r, k) for k in (3, 4, 5, 6))
        if r in sp.get("labels", {}):
            C.set_text(lab, sp["labels"][r])
        if r in sp.get("hints", {}):
            C.set_text(hint, sp["hints"][r])
        if r in sp.get("units", {}):
            C.set_text(unit, sp["units"][r])
        elif isinstance(unit.value, str) and unit.value.strip() in UNITS:
            new = UNITS[unit.value.strip()]
            unit.value = new
        ind = 2 if r in sp.get("indent2", ()) else 1
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", ind, wrap=True)
        unit.font = C.font(C.T_SMALL, False, C.MUTED)
        unit.alignment = C.align("left", "center", 1)
        hint.font = C.font(C.T_SMALL, False, C.MUTED)
        hint.alignment = C.align("left", "center", 1, wrap=True)
        # Eingabezustand
        state = "linked" if r == sp.get("linked") else ("optional" if r in sp.get("optional", ()) else "required")
        is_text = r in text_rows
        span = ("D", "E") if (is_text and unit.value in (None, "")) else ("D",)
        C.input_style(inp, state)
        inp.font = C.font(C.T_BODY, state != "linked", C.INPUT_FG)
        if len(span) == 2:
            C.input_style(ws.cell(r, 5), state)
            C.safe_merge(ws, "D", r, "E", r)
        inp.alignment = C.align("left", "center", 1, wrap=True) if is_text else C.align("right", "center", 1)
        for c in (lab, hint) + ((unit,) if len(span) == 1 else ()):
            row_line(ws, r, C.L(c.column), C.L(c.column))
        # Zeilenbedarf
        lines = max(nlines(lab.value, PX["C"], indent=ind),
                    nlines(hint.value, PX["F"], C.T_SMALL),
                    nlines(inp.value, px(*span), bold=True) if (is_text and not C.is_formula(inp.value)) else 1)
        need[r] = lines
        if lines > 2:
            print(f"  steps: {ws.title} Z. {r} braucht {lines} Zeilen: C {nlines(lab.value, PX['C'], indent=ind)} "
                  f"F {nlines(hint.value, PX['F'], C.T_SMALL)} ({lab.value!r} | {hint.value!r} | {inp.value!r})")
    for r, cond in sp.get("inactive", {}).items():
        ref = f"D{r}:E{r}" if r in text_rows and ws.cell(r, 5).value in (None, "") else f"D{r}"
        C.inactive_when(ws, ref, cond)
    return need


RESULT_FONT = {"n": (C.T_BODY, False, C.INK), "t": (C.T_BODY, False, C.INK), "s": (9.5, False, C.MUTED)}


def style_results(ws, sp):
    need = {}
    for r, (role, label, formula, fmt) in sp["results"].items():
        lab, val = ws.cell(r, 8), ws.cell(r, 9)
        if label is not None:
            if label.startswith('="'):
                lab.value = label
            else:
                C.set_text(lab, label)
        if isinstance(lab.value, str) and lab.value.startswith("   "):
            C.set_text(lab, lab.value.strip())
        if formula is not None:
            val.value = formula
        if fmt is not None:
            val.number_format = fmt
        for c in (lab, val):
            c.fill = C.NOFILL
            c.border = Border(bottom=C.side("hair", C.LINE))
        if role in RESULT_FONT:
            sz, b, color = RESULT_FONT[role]
            lab.font = C.font(sz, b, color)
            val.font = C.font(sz, b, C.INK if role != "s" else C.MUTED)
        elif role == "i":
            C.total(ws, r, "H", "I", level=1)
        elif role == "f":
            C.total(ws, r, "H", "I", level=3)
        elif role == "m":
            C.memo(ws, r, "H", "I")
            for c in (lab, val):
                c.border = Border()
        lab.alignment = C.align("left", "center", 2 if role == "s" else 1)
        val.alignment = C.align("right", "center", 1, wrap=(role == "t"))
        size = 9 if role == "m" else (9.5 if role == "s" else C.T_BODY)
        lines = 1 if lab.data_type == "f" else nlines(lab.value, PX["H"], size, role in ("i", "f"))
        need[r] = lines
        if lines > 1:
            print(f"  steps: {ws.title} H{r} umbricht ({lab.value})")
    for coord, kpi in sp.get("ampel", {}).items():
        drop_cf(ws, [coord])
        C.add_ampel(ws, coord, kpi)
    return need


def left_chart(ws, sheet_charts, spec, head_row, c1="C", c2="F", heights=None, target=CHART_H):
    """Kopf im L2-Stil und Diagramm darunter; gibt die letzte Diagrammzeile zurück."""
    l2_head(ws, head_row, c1, c2, spec["title"], spec.get("unit"))
    heights[head_row] = max(heights.get(head_row, 0), C.H_HEAD)
    r, acc = head_row + 1, 0.0
    while acc < target - 4:
        h = heights.get(r)
        if h is None:
            h = 15
            heights[r] = h
        acc += h
        r += 1
    last = r - 1
    chart = sheet_charts.get(spec.get("chart_col", C.col(c1) - 1))
    if chart is not None:
        set_anchor(chart, c1, head_row + 1, c2, last)
    return last


# ============================================================================ Standardseite
def standard_page(ws, n, sp, names):
    heights = {}
    old_nav, old_foot = sp["nav"], sp["foot"]
    # alte Button- und Fußzeilen leeren (werden unten neu gesetzt)
    for coord in (f"C{old_nav}", f"E{old_nav}", f"H{old_nav}", f"C{old_foot}", f"C{old_foot + 1}"):
        blank(ws[coord])
    # Diagramme nach Spalte merken
    charts = {ch.anchor._from.col: ch for ch in ws._charts}

    # Bänder Z. 10
    C.band_l1(ws, 10, "C", "F", "Eingaben")
    C.band_l1(ws, 10, "H", "I", "Ergebnis dieses Schritts", right=sp.get("band_right"))

    need_l = style_input_table(ws, n, sp)
    need_r = style_results(ws, sp)
    if n == 9:  # Rechtsform: verknüpfte Anzeige, Auswahl auf der Startseite
        ws["D12"].value = '=CHOOSE(Rechtsform_Idx,"Privatperson","vv-GmbH","GmbH / Holding")'
        C.text_link(ws["F12"], "Rechtsform auf der Startseite ändern  ›", "Start", "D23", size=C.T_SMALL, bold=True,
                    tooltip="Privat oder GmbH wählen Sie auf der Startseite („Kauf als“)")
        ws["F12"].alignment = C.align("left", "center", 1)
        ws.conditional_formatting.add("I16", FormulaRule(formula=["$I$16<0"], font=Font(color=C.GREEN, bold=True),
                                                         stopIfTrue=True))
    inputs_end = max(sp["inputs"])
    results_end = max(sp["results"])
    heights[11] = C.H_HEAD if need_r.get(11, 1) <= 1 else ROW2
    for r in range(12, max(inputs_end, results_end) + 1):
        heights[r] = height_for(max(need_l.get(r, 1), need_r.get(r, 1)))

    # Hinweiszeile unter der Tabelle (S07: zweites Darlehen)
    left_end = inputs_end
    if sp.get("note"):
        r = sp["note"]
        cell = ws.cell(r, 3)
        blank(cell)
        link_row = next((rr for rr in range(95, 125) if str(ws.parent["Eingaben"].cell(rr, 2).value or "")
                         .startswith("Darlehen II")), 104)
        C.text_link(cell, "Darlehen II (z. B. KfW mit Tilgungszuschuss) auf dem Blatt „Eingaben“ erfassen  ›",
                    "Eingaben", f"C{link_row}", size=C.T_SMALL, bold=True,
                    tooltip="Darlehen II wird nur auf dem Blatt „Eingaben“ erfasst")
        cell.alignment = C.align("left", "center", 1)
        heights[r] = ROW1
        left_end = r

    # Hinweiszeile zu den Eingaben (S01: gelbe Felder = Beispielwerte)
    if sp.get("info"):
        r, text = sp["info"]
        heights[r - 1] = GAP
        for c in C.iter_cells(ws, "C", r, "F", r):
            c.fill = C.fill(C.INPUT_BG)
            c.border = Border(left=C.side("thick", C.INPUT_LINE) if c.column == 3 else None)
        C.safe_merge(ws, "C", r, "F", r)
        cell = ws.cell(r, 3)
        C.set_text(cell, text)
        cell.font = C.font(C.T_SMALL, False, C.INK2)
        cell.alignment = C.align("left", "center", 1, wrap=True)
        heights[r] = height_for(nlines(text, px("C", "D", "E", "F"), C.T_SMALL))
        left_end = r

    # Einordnung
    co = sp["callout"]
    head = co["head"]
    body = co.get("body", head + 1)
    if co["src"] != f"H{body}":
        move_formula(ws, co["src"], f"H{body}")
    for rr in range(results_end + 1, head):
        heights.setdefault(rr, GAP)
    lines = est_lines(co["max_chars"])
    need_pt = lines * LINE_PT + 10
    heights[head] = max(heights.get(head, 0), C.H_HEAD)
    r, acc = body, 0.0
    while acc < need_pt - 2:
        if r in heights:
            acc += heights[r]
        else:
            rest = need_pt - acc
            h = rest if rest <= 30 else 15
            heights[r] = max(h, 12)
            acc += heights[r]
        r += 1
    callout_end = r - 1
    for rr in range(results_end + 1, callout_end + 1):  # alte Köpfe/Reste im Einordnungsbereich
        for c in (ws.cell(rr, 8), ws.cell(rr, 9)):
            if c.coordinate == f"H{body}" or c.value is None:
                continue
            if c.data_type == "f":
                print(f"  steps: {ws.title}!{c.coordinate} Formel im Einordnungsbereich – bleibt stehen")
                continue
            blank(c)
    callout_box(ws, "H", head, "I", body, callout_end, "Einordnung", status=co.get("status"))
    ws.cell(body, 8).font = C.font(C.T_SMALL, False, C.INK2)

    # Diagrammband
    chart_end = 0
    ch_specs = sp.get("charts", [])
    if len(ch_specs) == 1:
        head_row = left_end + 2
        heights.setdefault(left_end + 1, GAP)
        chart_end = left_chart(ws, charts, ch_specs[0], head_row, heights=heights)
    elif len(ch_specs) == 2:
        head_row = max(left_end, callout_end) + 2
        heights.setdefault(head_row - 1, GAP)
        chart_end = left_chart(ws, charts, ch_specs[0], head_row, "C", "F", heights)
        l2_head(ws, head_row, "H", "I", ch_specs[1]["title"], ch_specs[1].get("unit"))
        right = charts.get(ch_specs[1]["chart_col"])
        if right is not None:
            set_anchor(right, "H", head_row + 1, "I", chart_end)

    content_end = max(left_end, results_end, callout_end, chart_end)
    for r, h in heights.items():
        ws.row_dimensions[r].height = h
    nav_row = content_end + 3
    end = nav_and_footer(ws, n, nav_row, names)
    return end


# ============================================================================ S08 Zwischenergebnis
def page_s08(ws, names):
    n = 8
    for coord in ("C57", "E57", "H57", "C59", "C60", "H27"):
        blank(ws[coord])
    charts = {ch.anchor._from.col: ch for ch in ws._charts}
    h = ws.row_dimensions

    C.band_l1(ws, 10, "C", "F", "Cashflow vor Steuern – Jahr 1")
    for c in C.iter_cells(ws, "C", 10, "F", 10):  # weiße Fuge statt Linie über den Kacheln
        c.border = Border(left=c.border.left, bottom=C.side("thick", C.WHITE))
    C.band_l1(ws, 10, "H", "I", "Einordnung")

    tiles = [("C", "D", 11, 12, "Nettokaltmiete Ist / Monat", None, C.NUMFMT["eur"]),
             ("E", "F", 11, 12, "Einnahmenüberschuss (NOI) / Monat", None, C.NUMFMT["eur"]),
             ("C", "D", 14, 15, C.KPI["RATE"]["label"], None, C.NUMFMT["eur"]),
             ("E", "F", 14, 15, C.KPI["CFV"]["label"], "CFV", C.NUMFMT["eur"])]
    drop_cf(ws, ["E15", "D23", "D22"])
    for c1, c2, lr, vr, label, kpi, fmt in tiles:
        C.kpi_tile(ws, c1, c2, lr, vr, label=label, kpi=kpi, fmt=fmt, gap_right=(c1 == "C"))
    h[13].height = 2.25

    # Einordnung rechts neben den Kacheln (Formel aus H28)
    move_formula(ws, "H28", "H11")
    callout_box(ws, "H", 10, "I", 11, 15, None, status="CFV", head=False)
    h[10].height = C.H_BAND

    # Herleitung, Diagramm und Kapitaldienstdeckung
    h[16].height = GAP
    l2_head(ws, 17, "C", "D", "Herleitung pro Monat")
    l2_head_inline(ws, 17, "E", "F", "Einnahmen vs. Ausgaben", "€ / Monat, vor Steuern")
    labels = {18: "Nettokaltmiete Ist", 19: "– Bewirtschaftung inkl. Rücklagen", 20: "– Zinsen", 21: "– Tilgung",
              22: "= Cashflow vor Steuern / Monat", 23: "Kapitaldienstdeckung (DSCR)"}
    for r, text in labels.items():
        lab, val = ws.cell(r, 3), ws.cell(r, 4)
        C.set_text(lab, text)
        lab.font, val.font = C.font(), C.font()
        lab.alignment = C.align("left", "center", 1)
        val.alignment = C.align("right", "center", 1)
        row_line(ws, r, "C", "D")
        h[r].height = ROW1
    C.total(ws, 22, "C", "D", level=3)
    C.add_ampel(ws, "D22", "CFV")
    ws["D23"].number_format = C.NUMFMT["dscr"]
    C.add_ampel(ws, "D23", "DSCR")
    for c in (ws["C23"], ws["D23"]):
        c.border = Border(top=C.side("hair", C.LINE2), bottom=C.side("hair", C.LINE))
        c.alignment = C.align(c.alignment.horizontal, "bottom", 1)
    h[23].height = 28
    chart = charts.get(7)
    if chart is not None:
        set_anchor(chart, "E", 18, "F", 23)

    move_formula(ws, "H34", "H18")
    dscr_lines = est_lines(215)
    body_end = 18
    acc = ROW1
    while acc < dscr_lines * LINE_PT + 10 - 2:
        body_end += 1
        acc += h[body_end].height or ROW1
    callout_box(ws, "H", 17, "I", 18, body_end, "Kapitaldienstdeckung", status="DSCR")
    h[17].height = C.H_HEAD
    return nav_and_footer(ws, n, 26, names)


# ============================================================================ S12 Ergebnis
EINORDNUNG_S12_ZUSATZ = ('"Ab "&IF(Darlehen_Summe>0,Volltilgung_Txt,"sofort")&" entfällt die Rate an die Bank – der '
                         'Cashflow nach Steuern steigt dann auf rund "&FIXED(INDEX(Projektion!$D$37:$AQ$37,MIN(40,'
                         'IFERROR(MATCH(0,Finanzierung!$D$38:$AQ$38,0),40)+1)),0)&" € pro Monat: die Zusatzrente aus '
                         'dem Objekt (Prognosewerte)."')

LINKS_S12 = [("Dashboard", "Gesamtbewertung, Ampel, Zeitverlauf", "Dashboard"),
             ("Cockpit", "alle Kennzahlen im Detail", "Cockpit"),
             ("Diagramme", "Investition, Cashflow, Vermögen", "Diagramme"),
             ("AfA-Vergleich", "Abschreibungsvarianten", "AfA-Vergleich"),
             ("Sensitivität", "Break-even und Was-wäre-wenn", "Sensitivität"),
             ("Bankgespräch", "druckfertige Übersicht", "Bankgespräch")]


def page_s12(ws, names):
    n = 12
    for coord in ("C61", "E61", "H61", "C63", "C64", "H28", "H43", "H44", "H45", "H46", "H47", "H48"):
        blank(ws[coord])
    for ch in list(ws._charts):  # „Nettovermögen und Restschuld“ wiederholt S11 → entfernt
        if ch.anchor._from.col == 2:
            ws._charts.remove(ch)
    charts = {ch.anchor._from.col: ch for ch in ws._charts}
    h = ws.row_dimensions

    C.band_l1(ws, 10, "C", "F", "Ergebnis der Kalkulation")
    for c in C.iter_cells(ws, "C", 10, "F", 10):
        c.border = Border(left=c.border.left, bottom=C.side("thick", C.WHITE))
    C.band_l1(ws, 10, "H", "I", "Einordnung")

    drop_cf(ws, ["C12", "C15", "E15", "E12", "D23", "D22"])
    tiles = [
        ("C", "D", 11, 12, 13, C.KPI["CF"]["label"], "CF", C.NUMFMT["eur"],
         '="ab Jahr 2: "&FIXED(' + C.CF_YEAR2 + ',0)&" € / Monat"'),
        ("E", "F", 11, 12, 13, "Steuerwirkung Jahr 1 / Monat", None, C.NUMFMT["tax_effect"],
         '="Steuersatz Jahr 1: "&FIXED(Steuersatz_J1*100,1)&" %"'),
        ("C", "D", 14, 15, 16, C.KPI["EKR"]["label"], "EKR", C.NUMFMT["pct1"],
         "Vermögenszuwachs im Jahr 1 / eingesetztes Eigenkapital"),
        ("E", "F", 14, 15, 16, C.KPI["IRR"]["label"], "IRR", C.NUMFMT["pct1"],
         '="Haltedauer "&Haltedauer&" Jahre  ·  Multiple "&FIXED(EK_Multiple,2)&"×"'),
        ("C", "D", 17, 18, 19, None, None, C.NUMFMT["eur"], "Immobilienwert abzüglich Restschuld"),
        ("E", "F", 17, 18, 19, '="Zusatzrente pro Monat ab "&IF(Darlehen_Summe>0,Volltilgung_Txt,"sofort")', None,
         C.NUMFMT["eur"], "Cashflow nach Steuern nach Volltilgung (Prognose)"),
    ]
    for c1, c2, lr, vr, sr, label, kpi, fmt, sub in tiles:
        lab = label
        if lab is not None and C.is_formula(lab):
            ws.cell(lr, C.col(c1)).value = lab
            lab = None
        C.kpi_tile(ws, c1, c2, lr, vr, sub_row=sr, label=lab, kpi=kpi, fmt=fmt, sub=sub, gap_right=(c1 == "C"))
        if lr > 11:
            for c in C.iter_cells(ws, c1, lr, c2, lr):
                c.border = Border(right=c.border.right, top=C.side("thick", C.WHITE))
    # Steuerwirkung: Erstattung grün, Zahlung Navy
    ws.conditional_formatting.add("E12:F12", FormulaRule(formula=["$E$12<0"], font=Font(color=C.GREEN, bold=True),
                                                         stopIfTrue=True))

    # Einordnung rechts neben den Kacheln: Text aus H29 und H37 in einem Kasten
    first = ws["H29"].value
    ws["H11"].value = first + "&CHAR(10)&CHAR(10)&" + EINORDNUNG_S12_ZUSATZ
    blank(ws["H29"])
    blank(ws["H37"])
    callout_box(ws, "H", 10, "I", 11, 19, None, status="IRR", head=False)
    h[10].height = C.H_BAND

    # Zeile 20: Herleitung · Diagramm · Weiter zur Auswertung
    l2_head(ws, 20, "C", "D", "Herleitung pro Monat", height=28)
    l2_head_inline(ws, 20, "E", "F", "Cashflow nach Steuern", "Jahre 1–30, € p. a.", height=28)
    l2_head(ws, 20, "H", "I", "Weiter zur Auswertung", height=28)
    for c in C.iter_cells(ws, "C", 20, "I", 20):
        if c.value is not None:
            c.alignment = C.align(c.alignment.horizontal or "left", "bottom", 1)
    labels = {21: "Cashflow vor Steuern / Monat", 22: "± Steuerwirkung / Monat", 23: "= Cashflow nach Steuern / Monat",
              24: C.KPI["EK"]["label"], 25: "Gesamtertrag n. St. bis Verkauf", 26: C.KPI["MULT"]["label"]}
    for r, text in labels.items():
        lab, val = ws.cell(r, 3), ws.cell(r, 4)
        C.set_text(lab, text)
        lab.font, val.font = C.font(), C.font()
        lab.alignment = C.align("left", "center", 1)
        val.alignment = C.align("right", "center", 1)
        row_line(ws, r, "C", "D")
        h[r].height = ROW1
    ws["D22"].number_format = C.NUMFMT["tax_effect"]
    ws.conditional_formatting.add("D22", FormulaRule(formula=["$D$22<0"], font=Font(color=C.GREEN), stopIfTrue=True))
    C.total(ws, 23, "C", "D", level=3)
    C.add_ampel(ws, "D23", "CF")
    ws["D26"].number_format = C.NUMFMT["mult2"]
    for c in (ws["C24"], ws["D24"]):
        c.border = Border(top=C.side("hair", C.LINE2), bottom=C.side("hair", C.LINE))
        c.alignment = C.align(c.alignment.horizontal, "bottom", 1)
    h[24].height = 28
    chart = charts.get(7)
    if chart is not None:
        set_anchor(chart, "E", 21, "F", 26)
    for i, (word, rest, target) in enumerate(LINKS_S12):
        cell = ws.cell(21 + i, 8)
        rich_link(cell, word, rest, target, tooltip=f"Weiter: {word}")
        row_line(ws, 21 + i, "H", "I")
    return nav_and_footer(ws, n, 29, names)


# ============================================================================ Ablauf
def validation_messages(ws):
    """Datenüberprüfungen mit Fehlermeldung (Stopp) – ungültige Eingaben werden abgewiesen, Logik unverändert."""
    for dv in ws.data_validations.dataValidation:
        dv.showErrorMessage = True
        dv.errorStyle = "stop"
        if dv.type == "list":
            dv.errorTitle = "Bitte aus der Liste wählen"
            dv.error = "Dieser Wert ist nicht vorgesehen. Bitte einen Eintrag aus der Auswahlliste wählen."
        else:
            dv.errorTitle = "Ungültige Eingabe"
            dv.error = "Bitte eine Zahl im zulässigen Bereich eingeben (siehe Hinweis in derselben Zeile)."


def apply(wb):
    names = {int(m.group(1)): ws.title for ws in wb.worksheets for m in [re.match(r"S(\d\d) ", ws.title)] if m}
    for n in sorted(names):
        ws = wb[names[n]]
        old_max = max(ws.max_row, 70)
        grid(ws)
        unmerge_from(ws, 10)
        reset_styles(ws, 10, old_max)
        header(ws, n)
        if n == 8:
            end = page_s08(ws, names)
        elif n == 12:
            end = page_s12(ws, names)
        else:
            sp = SPEC[n]
            end = standard_page(ws, n, sp, names)
            if sp.get("hide"):
                C.hide_rows(ws, *sp["hide"])
                for r in range(sp["hide"][0], sp["hide"][1] + 1):
                    for c in C.iter_cells(ws, "C", r, "F", r):
                        c.font = C.font(C.T_SMALL, False, C.MUTED)
        hide = SPEC.get(n, {}).get("hide")
        clear_rows_after(ws, end, (hide[0] - 1) if hide else old_max)
        ws.sheet_view.showRowColHeaders = False
        validation_messages(ws)
