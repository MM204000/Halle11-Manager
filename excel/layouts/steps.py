"""Leitfaden-Seiten S01–S12 – Runde 2: EINE Komponentensprache aus core.py.

Nur Darstellung. Die Rechenformeln der Vorlage bleiben unverändert. Angepasst werden ausschließlich
nicht referenzierte Anzeigeformeln (Einordnungstexte, Textergebnisse, Kachel-Beschriftungen), statische
Beschriftungen, Stile, Zahlenformate, Zeilenhöhen/Spaltenbreiten, Verbünde, Links und Diagrammanker.

Anatomie einer Schrittseite (festes Zwei-Spalten-Raster)
  links  C:F  Abschnitt „Eingaben“ (C.section 1) · Tabellenkopf (C.section 2) · Eingabezeilen · ggf. Diagramm
  rechts H:I  Abschnitt „Ergebnis dieses Schritts“ · Ergebniszeilen mit Summenstufen (C.sum_row)
              · Einordnungs-Box (C.callout_box, bei Ampel mit Status-Pill „● Urteil · Ist-Wert“ und Statuskante)
              · „Als Nächstes“ (Vorschau auf den nächsten Schritt) – unten bündig mit der linken Spalte
  unten       zwei Leerzeilen unter der tieferen Spalte die Button-Reihe (C.btn_row: Zurück · Übersicht ·
              Weiter), eine Leerzeile darunter der Fuß.
S08/S12 (Ergebnisseiten, „Fintech“): KPI-Kacheln (C.tile, dunkel, Kontextzeile + Status-Chip), rechts
Einordnungs-Boxen je Kennzahl, darunter Herleitung und Diagramm.
"""
import re

from openpyxl.chart import BarChart, Reference, Series
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Alignment, Border
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C

# ============================================================================ Raster und Namen
# Schrittseiten: C Position · D Eingabe (einheitlich breit) · E nur Einheit · F Hinweis · G Rinne · H:I Ergebnis
GRID = (("C", 36), ("D", 37), ("E", 10), ("F", 46), ("G", 3), ("H", 36), ("I", 18))
# Ergebnisseiten S08/S12: zwei gleich breite Kachelspalten C:D | E:F (Kontext links, Status-Chip rechts)
GRID_KPI = (("C", 36), ("D", 29), ("E", 34), ("F", 30), ("G", 3), ("H", 36), ("I", 18))
LONG = ["Objekt", "Kaufpreis & Miete", "Kaufnebenkosten", "Kaufpreisaufteilung", "Maßnahmen & Reserve",
        "Bewirtschaftung", "Finanzierung", "Zwischenergebnis", "Steuern", "Abschreibung", "Prognose & Exit",
        "Ergebnis"]
GAP = C.H_GAP            # genau eine Leerzeile zwischen Komponenten
CHART_MIN = 180          # Mindesthöhe eines Diagramms (pt)
PIE_MIN = 150            # Mindesthöhe eines Kreises in der rechten Spalte (pt)
ROW1, ROW2, ROW3 = C.H_STEP_ROW, C.H_STEP_ROW2, 47
NBSP = " "
SIGNED_EFFECT = '-#,##0" €";+#,##0" €";"–"'   # Steuerwirkung aus Sicht des Cashflows: Erstattung +, Zahlung −

UNITS = {"€": None, "%": None, "€/Monat": "pro Monat", "€ p. a.": "pro Jahr", "€/m²": "je m²",
         "€/m² p. a.": "je m² p. a.", "% Darlehen": None, "% der Miete": "der Miete",
         "% Verkaufspreis": None}


# ============================================================================ Ampel-Boxen (P1-03)
def val_text(kpi):
    """Ist-Wert als Text für die Status-Pill."""
    return {"BMR": 'FIXED(Bruttomietrendite*100,1)&" %"',
            "NMR": 'FIXED(S06_NMR*100,1)&" %"',
            "DSCR": 'FIXED(DSCR_J1,2)&"×"',
            "IRR": 'FIXED(EK_IRR*100,1)&" %"',
            "CFV": 'FIXED(CF_vSt_J1/12,0)&" €"'}[kpi]


VALUE_REF = {"BMR": "Bruttomietrendite", "NMR": "S06_NMR", "DSCR": "DSCR_J1", "IRR": "EK_IRR", "CFV": "CF_vSt_J1"}


def pill_formula(kpi):
    """Status-Pill: „●  kritisch · 0,61×“ – Urteil (core.STATUS_WORDS) und Ist-Wert; reine Anzeigeformel."""
    conds = C.status_conditions(kpi, VALUE_REF[kpi])
    expr = '""'
    for cond, lvl in reversed(conds):
        expr = f'IF({cond},"{C.STATUS_DOT}  {C.STATUS_WORDS[lvl]} · "&{val_text(kpi)},{expr})'
    return "=" + expr, conds


# Texte der Ampel-Boxen: ein Satz zum Urteil mit Zielwert, danach die Einordnung (reine Anzeigeformeln)
TEXT_BMR = ('=IF(Bruttomietrendite<Ampel_BMR_gelb,"Unter "&FIXED(Ampel_BMR_gelb*100,1)&" % trägt sich das Objekt '
            'aus der Miete allein meist nicht – bei üblicher Finanzierung wird der Cashflow vor Steuern negativ. Es '
            'rechnet sich dann nur über Tilgung, Wertentwicklung und Steuereffekte.",IF(Bruttomietrendite<'
            'Ampel_BMR_gruen,"Ausbaufähig: Solide ist eine Bruttomietrendite ab "&FIXED(Ampel_BMR_gruen*100,1)&" %, '
            'zwischen "&FIXED(Ampel_BMR_gelb*100,1)&" und "&FIXED(Ampel_BMR_gruen*100,1)&" % trägt sich die '
            'Finanzierung meist nur mit mehr Eigenkapital. Ob der Cashflow reicht, zeigen die Schritte 6 bis 10.",'
            '"Solide: Das Ziel von "&FIXED(Ampel_BMR_gruen*100,1)&" % ist erreicht. Ob der Cashflow nach '
            'Kapitaldienst und Steuern positiv ist, entscheidet sich in den Schritten 6 bis 10."))')
TEXT_NMR_PREFIX = ('IF(S06_NMR<Ampel_NMR_gruen,"Die Nettomietrendite liegt unter dem Ziel von "&'
                   'FIXED(Ampel_NMR_gruen*100,1)&" %.","Die Nettomietrendite erreicht das Ziel von "&'
                   'FIXED(Ampel_NMR_gruen*100,1)&" %.")&" "')
TEXT_DSCR_S07 = ('=IF(NOT(ISNUMBER(DSCR_J1)),"Ohne Darlehen entfällt die Kapitaldienstdeckung.",IF(DSCR_J1>='
                 'Ampel_DSCR_gruen,"Der Einnahmenüberschuss deckt die Rate an die Bank mit Reserve – Banken '
                 'erwarten mindestens "&FIXED(Ampel_DSCR_gruen,2)&"×.",IF(DSCR_J1>=Ampel_DSCR_gelb,"Knapp: Banken '
                 'erwarten meist mindestens "&FIXED(Ampel_DSCR_gruen,2)&"× – mehr Eigenkapital oder Sicherheiten '
                 'einplanen.","Der Einnahmenüberschuss deckt nur "&FIXED(DSCR_J1*100,0)&" % der Rate. Banken '
                 'erwarten mindestens "&FIXED(Ampel_DSCR_gruen,2)&"× – den Rest tragen Sie aus Ihrem Einkommen.")))')
TEXT_DSCR_S08 = ('="Banken erwarten meist eine Kapitaldienstdeckung von mindestens "&FIXED(Ampel_DSCR_gruen,2)&'
                 '"× (Einnahmenüberschuss ÷ Rate). Darunter trägt der Investor einen Teil der Rate aus seinem '
                 'Einkommen."')
TEXT_CFV_S08 = ('=IF(CF_vSt_J1>=0,"Die Miete deckt Bewirtschaftung, Zinsen und Tilgung: Es bleiben "&'
                'FIXED(CF_vSt_J1/12,0)&" € pro Monat vor Steuern für Rücklagen oder Sondertilgungen.","Es fehlen "&'
                'FIXED(-CF_vSt_J1/12,0)&" € pro Monat, die aus anderen Einkünften zugeschossen werden. Ob sich das '
                'Objekt trotzdem lohnt, zeigen Steuereffekt (Schritte 9–10) und Wertentwicklung (Schritt 11).")')
TEXT_IRR_S12 = ('=IFERROR(IF(EK_IRR>=Ampel_IRR_gruen,"Über "&Haltedauer&" Jahre verzinst sich das Eigenkapital mit "&'
                'FIXED(EK_IRR*100,1)&" % p. a. nach Steuern – das Ziel von "&FIXED(Ampel_IRR_gruen*100,1)&" % ist '
                'erreicht.",IF(EK_IRR>=Ampel_IRR_gelb,"Über "&Haltedauer&" Jahre verzinst sich das Eigenkapital mit "&'
                'FIXED(EK_IRR*100,1)&" % p. a. nach Steuern – solide, aber unter dem Ziel von "&'
                'FIXED(Ampel_IRR_gruen*100,1)&" %. Hebel: Kaufpreis, Miete, Finanzierung und Haltedauer.","Über "&'
                'Haltedauer&" Jahre verzinst sich das Eigenkapital nur mit "&FIXED(EK_IRR*100,1)&" % p. a. nach '
                'Steuern – unter der Mindestschwelle von "&FIXED(Ampel_IRR_gelb*100,1)&" %. Kaufpreis, Miete und '
                'Finanzierung prüfen.")),"Die Eigenkapitalrendite lässt sich mit diesen Annahmen nicht berechnen.")')

# ============================================================================ Blatt-Spezifikation
# results: Zeile → (Rolle, Beschriftung|None, Formel|None, Format|None)
#   Rollen: n normal · s Unterposition · i Zwischensumme (sum_row sub) · f Blockergebnis (sum_row result)
#           · m nachrichtlich · t Textwert
# callout: head = Kopfzeile, src = Zelle mit dem Einordnungstext (wird nach head+1 verschoben), title, kpi
SPEC = {
    1: dict(
        inputs=range(12, 20), text=(12, 13, 14, 19),
        labels={16: "Grundstücksfläche bzw. -anteil", 17: "Baujahr (Fertigstellung)",
                18: "Kaufdatum (Nutzen und Lasten)"},
        hints={14: "Gewerbe: keine § 7b-Sonder-AfA, keine degressive AfA",
               15: "Basis für Miete/m², Rücklage und § 7b-Grenzen",
               16: "ETW: Miteigentumsanteil × Fläche (Teilungserklärung)",
               17: "steuert den AfA-Satz und die Prüfung von § 7b",
               18: "Jahr 1 = die ersten 12 Monate ab Kauf"},
        results={11: ("n", None, None, C.NUMFMT["pct1"]), 12: ("n", None, None, C.NUMFMT["pct1"]),
                 13: ("n", None, None, None),
                 14: ("n", None, '=IF(Wohngebaeude=1,"Ja","Nein (Gewerbe)")', None)},
        callout=dict(head=16, src="H17", title="AfA nach Baujahr"),
        info=(21, "Alle gelben Felder enthalten Beispielwerte – einfach überschreiben. Die Rechtsform wählen Sie "
                  "auf der Startseite  ›"),
        nav=24, foot=26),
    2: dict(
        inputs=range(12, 18), indent2=(13, 15),
        labels={13: "davon bewegliche Gegenstände", 15: "davon Anteil Erhaltungsrücklage",
                17: "Umlagefähige Betriebskosten"},
        hints={12: "inkl. beweglicher Gegenstände und Rücklagenanteil",
               13: "Küche, Möbel · grunderwerbsteuerfrei (§ 2 GrEStG)",
               14: "Küche i. d. R. 10 Jahre, Möbel 8–13 (AfA-Tabelle)",
               15: "nicht abschreibbar, aber grunderwerbsteuerpflichtig",
               16: "Soll-Miete lt. Mietvertrag bzw. Marktmiete",
               17: "Vorauszahlung des Mieters · Durchlaufposten"},
        results={11: ("n", "Kaufpreis ohne Inventar/Rücklage", None, None),
                 12: ("n", None, None, None), 13: ("n", None, None, None),
                 14: ("n", "Kaufpreisfaktor", "=Kaufpreisfaktor", C.NUMFMT["mult1"]),
                 15: ("f", "= Bruttomietrendite", None, C.NUMFMT["pct1"]),
                 16: ("m", "nachrichtlich: Warmmiete pro Monat", "=Miete_Monat+NK_umlagefaehig", C.NUMFMT["eur"])},
        callout=dict(head=18, src="H19", title="Bruttomietrendite", kpi="BMR", text=TEXT_BMR),
        nav=26, foot=28),
    3: dict(
        inputs=range(12, 17), optional=(12,),
        labels={12: "Grunderwerbsteuer abweichend", 16: "Sonstige Erwerbsnebenkosten"},
        hints={12: "optional · leer = Satz des Bundeslands (siehe Ergebnis)",
               16: "Gutachten, Fahrten – sofern objektbezogen"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None), 13: ("n", None, None, None),
                 14: ("n", None, None, None), 15: ("n", None, None, None), 16: ("i", None, None, None),
                 17: ("s", None, None, C.NUMFMT["pct1"]), 18: ("f", None, None, None)},
        callout=dict(head=20, src="H21", title="Steuerliche Behandlung"),
        charts=[dict(side="L", title="Zusammensetzung der Kaufnebenkosten", unit="Anteile in %", pie=True)],
        nav=44, foot=46, hide=(60, 66)),
    4: dict(
        inputs=range(12, 16), text=(12,),
        inactive={14: "KPA_Idx<>2", 15: "KPA_Idx<>3"},
        labels={15: "Gebäudewert (Gutachten / Vertrag)"},
        hints={12: "Standard: BMF-Arbeitshilfe · Gutachten möglich",
               13: "lt. Gutachterausschuss / BORIS (Methode 1)",
               14: "nur Methode „Bodenanteil in %“",
               15: "nur Methode „Gebäudewert“ (ohne Inventar)"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None),
                 13: ("n", None, None, C.NUMFMT["pct1"]),
                 14: ("n", "AK Grund und Boden (inkl. NK)", None, None),
                 15: ("f", "= AfA-Basis Gebäude (inkl. NK)", None, None),
                 16: ("m", "nachrichtlich: bewegliche Gegenstände", None, None)},
        callout=dict(head=18, src="H19", title="Gebäudeanteil und AfA"),
        charts=[dict(side="L", title="Aufteilung der Anschaffungskosten", unit="inkl. NK und Inventar · in %", pie=True)],
        nav=43, foot=45),
    5: dict(
        inputs=range(12, 20), text=(16,),
        inactive={17: "OR(Rechtsform_Idx>=2,Wohngebaeude=0)"},
        labels={12: "Renovierung Jahr 1 (brutto)", 13: "Renovierung Jahr 2 (brutto)",
                14: "Renovierung Jahr 3 (brutto)", 15: f"Sonderumlagen der WEG (Jahr{NBSP}1)",
                16: "Steuerliche Behandlung", 17: "Verteilung Erhaltungsaufwand",
                18: "Wertsteigernder Anteil", 19: "Einmalige Liquiditätsreserve"},
        hints={12: "Teil des anfänglichen Kapitalbedarfs",
               13: "aus laufendem Cashflow / Eigenkapital",
               14: f"Jahre 1–3 zählen für die 15{NBSP}%-Grenze",
               15: f"Erhaltungsaufwand · zählt zur 15{NBSP}%-Grenze",
               16: f"§{NBSP}6 Abs. 1 Nr. 1a EStG · über 15{NBSP}% → AfA",
               17: f"§{NBSP}82b EStDV · 1 = sofort, 2–5 = verteilt (Privat)",
               18: "Anteil der Maßnahmen, der den Objektwert erhöht",
               19: "z. B. Mieterwechsel · erhöht nur das Eigenkapital"},
        units={17: "Jahr(e)"},
        results={11: ("n", "Maßnahmen Jahre 1–3 (netto)", None, None),
                 12: ("n", None, None, None),
                 13: ("t", "Steuerliche Behandlung",
                      '=IF(HK_Flag=1,"Aktiviert (AfA)",IF(Verteilung_eff>1,"Verteilt (§ 82b)","Sofortabzug"))', None),
                 14: ("n", "Maßnahmen + Sonderumlage Jahr 1", None, None),
                 15: ("i", None, None, None), 16: ("f", None, None, None)},
        callout=dict(head=18, src="H19", title=f"Die 15{NBSP}%-Grenze"),
        charts=[dict(side="L", title="Zusammensetzung der Gesamtinvestition", unit="Anteile in %", pie=True)],
        nav=44, foot=46),
    6: dict(
        inputs=range(12, 20), indent2=(13, 14),
        labels={13: "davon nicht umlagefähig", 14: "davon Erhaltungsrücklage", 15: "Miet-/SE-Verwaltung",
                18: "Leerstand zu Beginn", 19: "Sonstige Werbungskosten"},
        hints={13: "Verwaltung WEG, Rücklage, Sonstiges – trägt der Eigentümer",
               14: "Teil von „nicht umlagefähig“ · Abzug erst bei Verausgabung",
               16: "kalkulatorisch, Faustwert 8–15 €/m² · nicht steuerwirksam",
               18: "Renovierung / Neuvermietung · mindert die Miete Jahr 1",
               19: "Steuerberatung, Kontoführung, Fahrten"},
        results={11: ("n", "Nettokaltmiete Ist (nach Ausfall)", None, None),
                 12: ("n", "– Bewirtschaftung inkl. Rücklagen", None, None),
                 13: ("s", "in % der Nettokaltmiete", None, C.NUMFMT["pct1"]),
                 14: ("i", "= Einnahmenüberschuss (NOI)", None, None),
                 15: ("f", "= Nettomietrendite", None, C.NUMFMT["pct1"])},
        band_right="Jahr 1 · pro Monat",
        callout=dict(head=17, src="H18", title="Nettomietrendite", kpi="NMR", prefix=TEXT_NMR_PREFIX),
        charts=[dict(side="L", title="Bewirtschaftungskosten Jahr 1", unit="€ p. a. · Anteile", pie=True)],
        nav=41, foot=43, hide=(60, 66)),
    7: dict(
        inputs=range(12, 22), note=22,
        labels={16: "Darlehen I – Anschlusszins", 19: "Grundschuldbestellung",
                20: "Bearbeitungs- und Vermittlungsgebühr"},
        hints={16: "nach der Zinsbindung · Annahme, die Annuität bleibt",
               19: "Notar + Grundbuch, ca. 0,4–0,6 % des Darlehens",
               20: "Bank, Wertermittlung · sofort abziehbar",
               21: "bis 5 % bei ≥ 5 Jahren Zinsbindung sofort abziehbar"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None), 13: ("i", None, None, None),
                 14: ("n", "Eigenkapitalquote", "=EK_Quote", C.NUMFMT["pct1"]),
                 15: ("n", "Beleihungsauslauf", "=Beleihung", C.NUMFMT["pct1"]),
                 16: ("n", "Restschuld nach Zinsbindung", None, None),
                 17: ("n", None, None, None),
                 18: ("f", "= Rate an die Bank / Monat", "=Kapitaldienst_Monat_J1", C.NUMFMT["eur"])},
        callout=dict(head=20, src="H20", body=21, title="Kapitaldienstdeckung (DSCR)", kpi="DSCR",
                     text=TEXT_DSCR_S07, second=dict(title="Zinsänderungsrisiko")),
        charts=[dict(side="L", title="Restschuld am Jahresende", unit="T€ · Jahresende · Jahre 1–35", chart_col=7),
                dict(side="R", title="Finanzierungsstruktur", unit="Anteile in %", chart_col=2)],
        chart_min=165, nav=44, foot=46),
    9: dict(
        inputs=range(12, 19), text=(12, 13, 16, 17), linked=12, optional=(15,),
        inactive={13: "Rechtsform_Idx>=2", 14: "Rechtsform_Idx>=2", 15: "Rechtsform_Idx>=2",
                  16: "Rechtsform_Idx>=2", 17: "Rechtsform_Idx>=2", 18: "Rechtsform_Idx=1"},
        labels={13: "Veranlagung", 14: "Einkommen ohne dieses Objekt (zvE)",
                15: "Grenzsteuersatz manuell", 16: "Solidaritätszuschlag",
                17: "Kirchensteuer", 18: "Gewerbesteuer-Hebesatz"},
        hints={13: "nur Privatperson",
               14: "nur Privatperson · Basis für den Grenzsteuersatz",
               15: "nur Privat · optional, leer = nach Tarif 2026",
               16: "nur Privat · „Nein“ bei ESt unter der Freigrenze",
               17: "nur Privat · Sonderausgabenabzug berücksichtigt",
               18: "nur GmbH ohne erweiterte Kürzung · z. B. 400 %"},
        results={11: ("n", None, None, None),
                 12: ("n", "Grenzbelastung inkl. Soli / KiSt", None, None),
                 13: ("n", "Steuersatz GmbH Jahr 1", None, None),
                 14: ("i", None, None, None), 15: ("n", None, None, None),
                 16: ("f", "= Steuerwirkung Jahr 1", None, C.NUMFMT["tax_effect"])},
        band_right="Jahr 1",
        callout=dict(head=18, src="H19", title="Besteuerung nach Rechtsform"),
        charts=[dict(side="L", title="Steuerliches Ergebnis und Steuer", unit="Jahre 1–20 · € p. a.", new=True)],
        nav=27, foot=29),
    10: dict(
        inputs=range(12, 18), text=(12, 14, 15, 16),
        inactive={13: "AfA_Idx<>6", 14: "AfA_Idx<>5", 17: 'Denkmal<>"Ja"'},
        labels={13: "Restnutzungsdauer lt. Gutachten", 14: "Degressiv: Wechsel zur linearen AfA",
                15: f"Sonder-AfA §{NBSP}7b (Neubau)", 16: f"Erhöhte AfA §{NBSP}7h / §{NBSP}7i EStG",
                17: f"Begünstigte Kosten §{NBSP}7h / 7i"},
        hints={12: "degressiv nur bei Baubeginn 10/2023 – 09/2029",
               13: "nur Methode Gutachten · kürzere Restnutzungsdauer",
               14: f"nur degressive AfA · Wechsel, sobald vorteilhaft",
               15: "5 % p. a. für 4 Jahre · EH 40 + QNG, Baukostengrenze",
               16: "Sanierungsgebiet / Denkmal: 9 % (J. 1–8), 7 % (J. 9–12)",
               17: "im Kaufpreis enthalten · lt. Bescheinigung"},
        results={11: ("n", "Angewendete AfA-Methode",
                      '=IF(Methode_eff=5,"Degressiv 5 %",IF(Methode_eff=6,"Gutachten (RND)","Linear "&'
                      'FIXED(CHOOSE(Methode_eff,AfA_Satz_auto,0.02,0.025,0.03)*100,1)&" %"))', None),
                 12: ("n", None, None, None), 13: ("n", f"+ Sonder-AfA §{NBSP}7b", None, None),
                 14: ("n", f"+ Erhöhte Absetzungen §{NBSP}7h/7i", None, None),
                 15: ("n", None, None, None), 16: ("f", None, None, None),
                 17: ("m", f"Steuereffekt ggü. Standard-AfA (10{NBSP}J.)", None, None)},
        band_right="Jahr 1",
        callout=dict(head=19, src="H20", title="Wahl der Abschreibung"),
        charts=[dict(side="L", title="Abschreibungen nach Komponenten", unit="Jahre 1–20 · € p. a.")],
        nav=45, foot=47),
    11: dict(
        inputs=range(12, 19), text=(18,), optional=(17,),
        inactive={18: "Rechtsform_Idx>=2"},
        labels={16: "Verkaufskosten", 17: "Verkaufspreis manuell"},
        units={16: None},
        hints={12: "Prognose · Mietpreisbremse und Kappungsgrenze",
               15: "Privat: Verkauf nach über 10 Jahren steuerfrei",
               16: "vom Verkaufspreis · Makler, Notar, Löschung",
               17: "optional · leer = Wertentwicklung lt. Prognose",
               18: "GmbH: immer Verlustvortrag (§ 10d EStG)"},
        results={11: ("n", "Verkaufspreis bei Exit", None, None),
                 12: ("n", None, None, None),
                 13: ("n", "Verkauf steuerpflichtig?", '=IF(Exit_steuerpflichtig=1,"Ja","Nein")', None),
                 14: ("n", None, None, None),
                 15: ("i", "= Nettoerlös nach Steuer und Ablösung", None, None),
                 16: ("n", "Gesamtertrag nach Steuern", None, None),
                 17: ("f", "= Eigenkapitalrendite (IRR n. St.)", None, C.NUMFMT["pct1"])},
        band_right="Verkauf",
        callout=dict(head=19, src="H20", title="Verkauf und Steuern"),
        charts=[dict(side="L", title="Immobilienwert, Restschuld und Nettovermögen", unit="T€ · Jahresende")],
        nav=45, foot=47),
}
TEMPLATE_CF = ["I15", "E15", "D23", "I17", "C12", "C15"]   # Ampel-Regeln der Vorlage (werden neu gesetzt)


# ============================================================================ Messen
def lines(text, width_px, size=C.T_BODY, bold=False, indent=1):
    return C.lines_needed_metric(text, width_px, size, bold, indent) if text not in (None, "") else 1


def _split_args(inner):
    """Argumente einer Funktion auf oberster Ebene trennen (Anführungszeichen und Klammern beachten)."""
    args, depth, cur, q = [], 0, "", False
    for ch in inner:
        if ch == '"':
            q = not q
        if not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                args.append(cur)
                cur = ""
                continue
        cur += ch
    args.append(cur)
    return args


def _concat_parts(expr):
    """Operanden einer &-Verkettung auf oberster Ebene."""
    parts, depth, cur, q = [], 0, "", False
    for ch in expr:
        if ch == '"':
            q = not q
        if not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "&" and depth == 0:
                parts.append(cur)
                cur = ""
                continue
        cur += ch
    parts.append(cur)
    return [p.strip() for p in parts]


def variants(expr, ref_chars=7):
    """Mögliche Anzeigetexte einer Textformel (je IF-Zweig), Bezüge/Funktionen als `ref_chars` Platzhalter."""
    expr = expr.strip()
    if expr.startswith("="):
        expr = expr[1:].strip()
    parts = _concat_parts(expr)
    if len(parts) > 1:
        out = [""]
        for p in parts:
            vs = variants(p, ref_chars)
            out = [a + b for a in out for b in vs][:64]
        return out
    m = re.fullmatch(r"(IFERROR|IF)\((.*)\)", expr, re.S)
    if m:
        args = _split_args(m.group(2))
        branches = args[1:] if m.group(1) == "IF" else args[:2]
        out = []
        for b in branches:
            out += variants(b, ref_chars)
        return out or [""]
    if expr.startswith('"') and expr.endswith('"'):
        return [expr[1:-1].replace('""', '"')]
    if re.fullmatch(r"CHAR\(10\)", expr):
        return ["\n"]
    return ["0" * ref_chars]


def text_bound(value):
    """Längster Anzeigetext (in Zeilen gemessen wird später) – statisch oder je Formelzweig."""
    if value is None:
        return ""
    if not isinstance(value, str) or not C.is_formula(value):
        return str(value)
    vs = variants(value)
    return max(vs, key=len) if vs else ""


def row_height(n):
    return ROW1 if n <= 1 else (ROW2 if n == 2 else ROW3)


def body_need(text, width_px, size=C.T_SMALL, extra=0.0):
    """Höhe (pt) eines oben ausgerichteten Fließtexts: Zeilen × Zeilenpitch + 8 pt Innenraum."""
    n = lines(text, width_px, size, False, 1)
    return C.px_pt(n * C.line_pt(size) + 8 + extra), n


def lit(text, size=200):
    """Excel-Textliteral (Stücke < 255 Zeichen, mit & verbunden)."""
    parts = [text[i:i + size] for i in range(0, len(text), size)] or [""]
    return "&".join('"' + p.replace('"', '""') + '"' for p in parts)


def stack(heights, start, need, fill_max=None):
    """Zeilen ab `start` belegen, bis `need` pt erreicht sind. Gesetzte Zeilen zählen mit ihrer Höhe;
    die erste freie Zeile erhält genau den Rest (mind. 12 pt). Liefert die letzte belegte Zeile."""
    r, acc = start, 0.0
    while True:
        if r in heights:
            acc += heights[r]
            if acc >= need - 1:
                return r
        else:
            heights[r] = max(C.px_pt(need - acc), 12 if fill_max is None else min(fill_max, 12))
            return r
        r += 1


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


def move_formula(ws, src, dst):
    """Nicht referenzierte Anzeigeformel an eine neue Zelle setzen (Formeltext bleibt gleich)."""
    if src == dst:
        return ws[dst].value
    v = ws[src].value
    ws[dst].value = v
    blank(ws[src])
    return v


def head2(ws, row, c1, c2, title, unit=None):
    """Ebene-2-Kopf (Unterabschnitt, Linie) für Diagramm- und Tabellenblöcke: 8 pt Versalien, Einheit rechts."""
    C.section(ws, row, c1, c2, title, level=2, variant="line", meta=unit)


def set_anchor(chart, c1, r1, c2, r2):
    """Diagramm genau in C1:R1 … C2:R2 (Zellgrenzen, ohne Versatz)."""
    a = chart.anchor
    if not hasattr(a, "_from") or not hasattr(a, "to"):
        chart.anchor = a = TwoCellAnchor(_from=AnchorMarker(), to=AnchorMarker())
    a._from.col, a._from.colOff, a._from.row, a._from.rowOff = C.col(c1) - 1, 0, r1 - 1, 0
    a.to.col, a.to.colOff, a.to.row, a.to.rowOff = C.col(c2), 0, r2, 0


def rich_link(cell, word, rest, target_sheet, tooltip=None):
    """Listen-Link: Link-Wort fett Blau, Beschreibung grau, Pfeil am Ende."""
    cell.value = C.rich([(word, C.T_BODY, True, C.BLUE), (f" – {rest}", C.T_BODY, False, C.MUTED),
                         ("  ›", C.T_BODY, True, C.BLUE)])
    cell.font = C.font(C.T_BODY, True, C.BLUE)
    cell.alignment = C.align("left", "center", 1)
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=C.link_loc(target_sheet), display=f"{word} – {rest}",
                               tooltip=tooltip)


def row_line(ws, row, c1, c2, color=C.LINE):
    for c in C.iter_cells(ws, c1, row, c2, row):
        b = c.border
        c.border = Border(left=b.left, right=b.right, top=b.top, bottom=C.side("hair", color))


# ============================================================================ Bausteine der Seite
def grid(ws, spec=GRID):
    for letter, w in spec:
        ws.column_dimensions[letter].width = w


def header(ws, n):
    title = ws["C6"].value if isinstance(ws["C6"].value, str) else LONG[n - 1]
    C.page_header(ws, "C", "I", f"Schritt {n:02d} / 12  ·  {LONG[n - 1]}", title,
                  context=("=Obj_Name", '=Obj_Adresse&IFERROR(IF(Erstellt_fuer="","","  ·  Erstellt für "&Erstellt_fuer),"")'),
                  context_col="H")
    C.safe_merge(ws, "C", 7, "F", 7)
    C.safe_merge(ws, "H", 6, "I", 6)
    C.safe_merge(ws, "H", 7, "I", 7)


def box(ws, heights, head, title, kpi=None, text=None, body=None, fixed_end=None):
    """Einordnungs-Box (core.callout_box) mit Zeilenbedarf nach Textlänge. Liefert die letzte Zeile."""
    body = body or head + 1
    cell = ws.cell(body, 8)
    if text is not None:
        cell.value = text
    heights[head] = max(heights.get(head, 0), C.H_CALLOUT_HEAD)
    need, _ = max((body_need(t, C.span_px(ws, "H", "I")) for t in variants(cell.value)), key=lambda x: x[0]) \
        if C.is_formula(cell.value) else body_need(text_bound(cell.value), C.span_px(ws, "H", "I"))
    end = fixed_end or stack(heights, body, need)
    pill, conds = pill_formula(kpi) if kpi else (False, None)
    C.callout_box(ws, "H", head, "I", body, end, title=title, conditions=conds, pill=pill, fit=None)
    return end


def next_card(ws, heights, head, n, names):
    """„Als Nächstes“: Vorschau auf den nächsten Schritt (Name fett + Kurzbeschreibung), Link auf das Blatt."""
    nxt = names[n + 1]
    desc = ws.parent[nxt]["C7"].value
    desc = desc if isinstance(desc, str) and not C.is_formula(desc) else ""
    head2(ws, head, "H", "I", "Als Nächstes", f"Schritt {n + 1:02d} / 12")
    heights[head] = max(heights.get(head, 0), C.H_HEAD)
    body = head + 1
    for c in C.iter_cells(ws, "H", body, "I", body):
        c.fill, c.border = C.NOFILL, Border()
    cell = ws.cell(body, 8)
    cell.value = C.rich([(LONG[n], C.T_BODY, True, C.NAVY), (f"  ·  {desc}", C.T_SMALL, False, C.MUTED)])
    cell.font = C.font(C.T_SMALL, False, C.MUTED)
    cell.alignment = C.align("left", "top", 1, wrap=True)
    n_l = lines(f"{LONG[n]}  ·  {desc}", C.span_px(ws, "H", "I") * 0.95, C.T_SMALL, False, 1)
    need = C.grid_height(n_l * C.line_pt(C.T_BODY) + 5)
    return body, need


def nav_and_footer(ws, n, row, names):
    """Button-Reihe (Zurück sekundär · Übersicht sekundär · Weiter primär) und Fuß."""
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
    for c in C.iter_cells(ws, "C", row, "I", row):
        c.value, c.hyperlink = None, None
    res = C.btn_row(ws, row, [
        dict(c1="C", c2="C", text=prev_text, target=prev_sheet, kind="secondary", tooltip=prev_tip),
        dict(c1="E", c2="F", text="Übersicht: Leitfaden", target="Leitfaden", kind="secondary",
             tooltip="Alle zwölf Schritte im Überblick"),
        dict(c1="H", c2="I", text=next_text, target=next_sheet, kind="primary", tooltip=next_tip)])
    for text, w, ok in res:
        if not ok:
            print(f"  steps: {ws.title} Button „{text}“ knapp ({w} px)")
    set_h[row + 1].height = GAP
    C.footer(ws, row + 2, "C", "I")
    for c in C.iter_cells(ws, "C", row + 2, "I", row + 2):
        c.border = Border(top=C.side("hair", C.LINE2))
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


def style_input_table(ws, sp):
    """Kopfzeile Z. 11 und Eingabezeilen: Beschriftung, Eingabe (nur Spalte D), Einheit (nur E), Hinweis.
    Liefert den Zeilenbedarf je Zeile (Anzahl Textzeilen)."""
    need = {}
    C.section(ws, 11, "C", "F", None, level=2, labels={"C": ("Position", "left"), "D": ("Eingabe", "right"),
                                                          "E": ("Einheit", "left"), "F": ("Hinweis", "left")})
    text_rows = set(sp.get("text", ()))
    wc, wd, wf = C.span_px(ws, "C", "C"), C.span_px(ws, "D", "D"), C.span_px(ws, "F", "F")
    for r in sp["inputs"]:
        lab, inp, unit, hint = (ws.cell(r, k) for k in (3, 4, 5, 6))
        if r in sp.get("labels", {}):
            C.set_text(lab, sp["labels"][r])
        if r in sp.get("hints", {}):
            C.set_text(hint, sp["hints"][r])
        if r in sp.get("units", {}):
            C.set_text(unit, sp["units"][r])
        elif isinstance(unit.value, str) and unit.value.strip() in UNITS:
            unit.value = UNITS[unit.value.strip()]
        ind = 2 if r in sp.get("indent2", ()) else 1
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", ind, wrap=True)
        unit.font = C.font(C.T_SMALL, False, C.MUTED)
        unit.alignment = C.align("left", "center", 1)
        hint.font = C.font(C.T_SMALL, False, C.MUTED)
        hint.alignment = C.align("left", "center", 1, wrap=True)
        state = "linked" if r == sp.get("linked") else ("optional" if r in sp.get("optional", ()) else "required")
        is_text = r in text_rows
        C.input_style(inp, state)
        inp.font = C.font(C.T_BODY, state != "linked", C.INPUT_FG)
        inp.alignment = C.align("left", "center", 1, wrap=True) if is_text else C.align("right", "center", 1)
        for c in (lab, unit, hint):
            row_line(ws, r, C.L(c.column), C.L(c.column))
        n = max(lines(lab.value, wc, C.T_BODY, False, ind),
                lines(hint.value, wf, C.T_SMALL, False, 1),
                lines(C.display_text(inp), wd, C.T_BODY, True, 1) if is_text else 1)
        need[r] = n
        if n > 1:
            print(f"  steps: {ws.title} Z. {r} zweizeilig ({lab.value!r} | {hint.value!r} | {C.display_text(inp)!r})")
    for r, cond in sp.get("inactive", {}).items():
        C.inactive_when(ws, f"D{r}", cond)
    return need


RESULT_FONT = {"n": (C.T_BODY, False, C.INK), "t": (C.T_BODY, False, C.INK), "s": (C.T_SMALL, False, C.MUTED)}


def style_results(ws, sp):
    need = {}
    wh = C.span_px(ws, "H", "H")
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
            val.font = C.font(sz, b, color)
        elif role == "i":
            C.sum_row(ws, r, "H", "I", "sub", neg=True)
        elif role == "f":
            C.sum_row(ws, r, "H", "I", "result", neg=(fmt != C.NUMFMT["tax_effect"]))
        elif role == "m":
            C.memo(ws, r, "H", "I")
            for c in (lab, val):
                c.border = Border()
        lab.alignment = C.align("left", "center", 2 if role == "s" else 1)
        val.alignment = C.align("right", "center", 1)
        size = C.T_SMALL if role in ("m", "s") else C.T_BODY
        n = 1 if lab.data_type == "f" else lines(lab.value, wh, size, role in ("i", "f"), 1)
        need[r] = n
        if n > 1:
            print(f"  steps: {ws.title} H{r} umbricht ({lab.value})")
    return need


def chart_block(ws, heights, charts, spec, head_row, c1, c2, end_row=None, target=CHART_MIN, chart_c2=None):
    """Ebene-2-Kopf und Diagramm darunter; das Diagramm endet auf `end_row` oder nach `target` pt."""
    head2(ws, head_row, c1, c2, spec["title"], spec.get("unit"))
    heights[head_row] = max(heights.get(head_row, 0), C.H_HEAD)
    if end_row is None:
        r, acc = head_row + 1, 0.0
        while acc < target - 4:
            if r not in heights:
                heights[r] = 15
            acc += heights[r]
            r += 1
        end_row = r - 1
    chart = charts.get(spec.get("chart_col", C.col(c1) - 1))
    if chart is not None:
        set_anchor(chart, c1, head_row + 1, chart_c2 or c2, end_row)
    return end_row


def new_tax_chart(ws):
    """S09: flaches Säulendiagramm „Steuerliches Ergebnis und Steuer“ aus den Diagrammdaten (Jahre 1–20).
    Inhalt/Horizont vereinheitlicht layouts/diagramme.py, das Styling finish_pro.py (Art „steuer“)."""
    dg = ws.parent["Diagramme"]
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap, ch.gapWidth = "col", "clustered", 0, 60
    for row, title in ((196, "Steuerliches Ergebnis"), (197, "Steuer (+ Zahlung / – Erstattung)")):
        s = Series(Reference(dg, min_col=4, max_col=23, min_row=row), title=title)
        ch.series.append(s)
    ch.set_categories(Reference(dg, min_col=4, max_col=23, min_row=178))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    ws.add_chart(ch)
    ch.anchor = TwoCellAnchor(_from=AnchorMarker(), to=AnchorMarker())
    return ch


# ============================================================================ Standardseite
def standard_page(ws, n, sp, names):
    heights = {}
    old_nav, old_foot = sp["nav"], sp["foot"]
    for coord in (f"C{old_nav}", f"E{old_nav}", f"H{old_nav}", f"C{old_foot}", f"C{old_foot + 1}"):
        blank(ws[coord])
    charts = {ch.anchor._from.col: ch for ch in ws._charts}
    for spec in sp.get("charts", []):
        if spec.get("new"):
            charts[C.col("C") - 1] = new_tax_chart(ws)

    # Abschnitte Z. 10
    C.section(ws, 10, "C", "F", "Eingaben")
    C.section(ws, 10, "H", "I", "Ergebnis dieses Schritts", meta=sp.get("band_right"))
    heights[10] = C.H_BAND

    if n == 9:  # Rechtsform: verknüpfte Anzeige, Auswahl auf der Startseite
        ws["D12"].value = '=CHOOSE(Rechtsform_Idx,"Privatperson","vv-GmbH","GmbH / Holding")'
    need_l = style_input_table(ws, sp)
    need_r = style_results(ws, sp)
    if n == 9:
        C.text_link(ws["F12"], "Rechtsform auf der Startseite ändern  ›", "Start", "D23", size=C.T_SMALL, bold=True,
                    tooltip="Privat oder GmbH wählen Sie auf der Startseite („Kauf als“)")
        ws["F12"].alignment = C.align("left", "center", 1)
    inputs_end = max(sp["inputs"])
    results_end = max(sp["results"])
    heights[11] = C.H_HEAD if need_r.get(11, 1) <= 1 else ROW2
    for r in range(12, max(inputs_end, results_end) + 1):
        heights[r] = row_height(max(need_l.get(r, 1), need_r.get(r, 1)))

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

    # Hinweis-Banner zu den Eingaben (nur S01, Stil wie Leitfaden-Banner, Link auf die Rechtsform)
    if sp.get("info"):
        r, text = sp["info"]
        heights.setdefault(r - 1, GAP)
        for c in C.iter_cells(ws, "C", r, "F", r):
            c.fill = C.fill(C.INPUT_BG)
            c.border = Border(left=C.side("thick", C.INPUT_LINE) if c.column == 3 else None)
        C.safe_merge(ws, "C", r, "F", r)
        cell = ws.cell(r, 3)
        C.set_text(cell, text)
        cell.font = C.font(C.T_SMALL, False, C.INK2)
        cell.alignment = C.align("left", "center", 1, wrap=True)
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=C.link_loc("Start", "D23"), display=text,
                                   tooltip="Privat oder GmbH – Auswahl auf der Startseite („Kauf als“)")
        heights[r] = row_height(lines(text, C.span_px(ws, "C", "F"), C.T_SMALL, False, 1))
        left_end = r

    # Diagrammkopf links reservieren (die Einordnung rechts läuft über diese Zeilen)
    ch_specs = sp.get("charts", [])
    chead = None
    if len(ch_specs) == 1:
        chead = left_end + 2
        heights.setdefault(left_end + 1, GAP)
        heights[chead] = max(heights.get(chead, 0), C.H_HEAD)

    # Einordnung (rechts, eine Leerzeile unter dem Ergebnis)
    co = sp["callout"]
    head = co["head"]
    body = co.get("body", head + 1)
    src_value = ws[co["src"]].value
    blank(ws[co["src"]])
    for rr in range(results_end + 1, results_end + 16):   # alte Köpfe/Reste der Vorlage im Einordnungsbereich
        for c in (ws.cell(rr, 8), ws.cell(rr, 9)):
            if c.value is not None and not C.is_formula(c.value):
                blank(c)
    for rr in range(results_end + 1, head):
        heights.setdefault(rr, GAP)
    if co.get("second"):  # S07: Ampel-Box DSCR, darunter neutrale Box mit dem Originaltext
        second_text = src_value.replace('="Zinsänderungsrisiko: Nach ', '="Nach ')
        text = co["text"]
    elif co.get("text"):
        text = co["text"]
    elif co.get("prefix"):
        old = src_value[1:] if C.is_formula(src_value) else lit(str(src_value))
        text = "=" + co["prefix"] + "&" + old
    else:
        text = src_value
    callout_end = box(ws, heights, head, co["title"], co.get("kpi"), text, body)
    if co.get("second"):
        h2 = callout_end + 2
        heights.setdefault(callout_end + 1, GAP)
        ws.cell(h2 + 1, 8).value = second_text
        callout_end = box(ws, heights, h2, co["second"]["title"], None, None, h2 + 1)

    # Diagrammband
    right_end = callout_end
    chart_end = 0
    card = n not in (7, 12)
    if len(ch_specs) == 1:
        # rechte Spalte: flexible Leerzeile + „Als Nächstes“ – unten bündig mit dem Diagramm
        g = max(callout_end + 1, chead + 1)
        while g in heights:
            g += 1
        cbody, cneed = next_card(ws, heights, g + 1, n, names)
        heights[cbody] = max(cneed, heights.get(cbody, 0))
        C.safe_merge(ws, "H", cbody, "I", cbody)
        fixed = sum(heights.get(r, 0) for r in range(chead + 1, cbody + 1) if r != g)
        heights[g] = max(GAP, C.px_pt(sp.get("chart_min", CHART_MIN) - fixed))
        chart_end = chart_block(ws, heights, charts, ch_specs[0], chead, "C", "F", end_row=cbody,
                                chart_c2="E" if ch_specs[0].get("pie") else "F")
        right_end = cbody
    elif len(ch_specs) == 2:
        lhead, rhead = left_end + 2, callout_end + 2
        heights.setdefault(left_end + 1, GAP)
        heights.setdefault(callout_end + 1, GAP)
        heights[lhead] = max(heights.get(lhead, 0), C.H_HEAD)
        heights[rhead] = max(heights.get(rhead, 0), C.H_HEAD)
        target = sp.get("chart_min", CHART_MIN)
        r, acc = lhead + 1, 0.0
        while True:
            heights.setdefault(r, 15)
            acc += heights[r]
            if acc >= target - 4 and r > rhead and \
                    sum(heights[x] for x in range(rhead + 1, r + 1)) >= PIE_MIN - 4:
                break
            r += 1
        chart_end = chart_block(ws, heights, charts, ch_specs[0], lhead, "C", "F", end_row=r)
        chart_block(ws, heights, charts, ch_specs[1], rhead, "H", "I", end_row=r)
        right_end = r
    elif card:
        g = callout_end + 1
        heights.setdefault(g, GAP)
        cbody, cneed = next_card(ws, heights, g + 1, n, names)
        end = stack(heights, cbody, cneed)
        C.safe_merge(ws, "H", cbody, "I", end)
        right_end = end

    content_end = max(left_end, results_end, right_end, chart_end)
    for r, h in heights.items():
        ws.row_dimensions[r].height = h
    return nav_and_footer(ws, n, content_end + 3, names)


# ============================================================================ S08 Zwischenergebnis
def page_s08(ws, names):
    n = 8
    for coord in ("C57", "E57", "H57", "C59", "C60", "H27"):
        blank(ws[coord])
    charts = {ch.anchor._from.col: ch for ch in ws._charts}
    h = ws.row_dimensions

    C.section(ws, 10, "C", "F", "Cashflow vor Steuern – Jahr 1")
    tiles = [
        dict(c1="C", c2="D", lr=11, label="Nettokaltmiete Ist / Monat", sub="nach Mietausfall und Leerstand"),
        dict(c1="E", c2="F", lr=11, label="Einnahmenüberschuss (NOI) / Monat", sub="Miete abzüglich Bewirtschaftung"),
        dict(c1="C", c2="D", lr=14, label="Rate an die Bank / Monat", sub="Zinsen und Tilgung · Jahr 1"),
        dict(c1="E", c2="F", lr=14, label="Cashflow vor Steuern / Monat", kpi="CFV", sub="Ziel ≥ 0 € · vor Steuern"),
    ]
    drop_cf(ws, TEMPLATE_CF + ["D22"])
    for t in tiles:
        C.tile(ws, t["c1"], t["c2"], t["lr"], t["lr"] + 1, t["lr"] + 2, label=t["label"], kpi=t.get("kpi"),
               fmt=C.NUMFMT["eur"], sub=t["sub"], gap_right=(t["c1"] == "C"), gap_top=(t["lr"] > 11))

    # Einordnung rechts: je Kennzahl eine Box (Cashflow · Kapitaldienstdeckung)
    blank(ws["H28"])
    blank(ws["H34"])
    heights = {10: C.H_BAND, 11: C.H_TILE_LABEL, 12: C.H_TILE_VALUE, 13: C.H_TILE_SUB,
               14: C.H_CALLOUT_HEAD, 15: C.H_TILE_VALUE, 16: C.H_TILE_SUB}
    box(ws, heights, 10, "Cashflow vor Steuern", "CFV", TEXT_CFV_S08, 11, fixed_end=12)
    box(ws, heights, 14, "Kapitaldienstdeckung (DSCR)", "DSCR", TEXT_DSCR_S08, 15, fixed_end=16)
    heights[10] = C.H_BAND

    # Herleitung, Diagramm
    head2(ws, 17, "C", "D", "Herleitung pro Monat")
    head2(ws, 17, "E", "F", "Einnahmen vs. Ausgaben", "€ / Monat · vor Steuern")
    labels = {18: "Nettokaltmiete Ist", 19: "– Bewirtschaftung inkl. Rücklagen", 20: "– Zinsen", 21: "– Tilgung",
              22: "= Cashflow vor Steuern / Monat", 23: "Kapitaldienstdeckung (DSCR)"}
    for r, text in labels.items():
        lab, val = ws.cell(r, 3), ws.cell(r, 4)
        C.set_text(lab, text)
        lab.font, val.font = C.font(), C.font()
        lab.alignment = C.align("left", "center", 1)
        val.alignment = C.align("right", "center", 1)
        row_line(ws, r, "C", "D")
        heights[r] = ROW1
    C.sum_row(ws, 22, "C", "D", "result", neg=True)
    ws["D23"].number_format = C.NUMFMT["dscr"]
    for c in (ws["C23"], ws["D23"]):
        c.border = Border(bottom=C.side("hair", C.LINE))
        c.alignment = C.align(c.alignment.horizontal, "bottom", 1)
    heights[17] = C.H_BAND
    heights[23] = ROW2
    for c in C.iter_cells(ws, "C", 17, "F", 17):
        c.alignment = C.align(c.alignment.horizontal or "left", "bottom", 1)
    chart = charts.get(7)
    if chart is not None:
        set_anchor(chart, "E", 18, "F", 23)

    # rechts: „Als Nächstes“ – Kopf auf einer Linie mit „Herleitung“ und „Einnahmen vs. Ausgaben“
    cbody, cneed = next_card(ws, heights, 17, n, names)
    C.safe_merge(ws, "H", cbody, "I", stack(heights, cbody, cneed))
    for c in C.iter_cells(ws, "H", 17, "I", 17):
        c.alignment = C.align(c.alignment.horizontal or "left", "bottom", 1)
    for r, hh in heights.items():
        h[r].height = hh
    return nav_and_footer(ws, n, 26, names)


# ============================================================================ S12 Ergebnis
EINORDNUNG_S12_ZUSATZ = ('"Ab "&IF(Darlehen_Summe>0,Volltilgung_Txt,"sofort")&" entfällt die Rate an die Bank – der '
                         'Cashflow nach Steuern steigt dann auf rund "&FIXED(INDEX(Projektion!$D$37:$AQ$37,MIN(40,'
                         'IFERROR(MATCH(0,Finanzierung!$D$38:$AQ$38,0),40)+1)),0)&" € pro Monat: die Zusatzrente aus '
                         'dem Objekt (Prognosewerte)."')

S12_TAX_OLD = ('abzüglich Rate ("&FIXED(Kapitaldienst_Monat_J1,0)&" €), Bewirtschaftung ("&FIXED(BWK_J1/12,0)&'
               '" €) und Steuer ("&FIXED(Steuer_J1/12,0)&" €).')
S12_TAX_NEW = ('abzüglich Rate ("&FIXED(Kapitaldienst_Monat_J1,0)&" €) und Bewirtschaftung ("&FIXED(BWK_J1/12,0)&'
               '" €), "&IF(Steuer_J1<0,"zuzüglich Steuererstattung (","abzüglich Steuerzahlung (")&'
               'FIXED(ABS(Steuer_J1/12),0)&" €).')

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

    C.section(ws, 10, "C", "F", "Ergebnis der Kalkulation")
    drop_cf(ws, TEMPLATE_CF + ["E12", "D22", "D23"])
    ws["C17"].value = '="NETTOVERMÖGEN NACH "&Haltedauer&" JAHREN"'
    ws["E17"].value = '="ZUSATZRENTE / MONAT AB "&IF(Darlehen_Summe>0,UPPER(Volltilgung_Txt),"SOFORT")'
    tiles = [
        dict(c1="C", c2="D", lr=11, label="Cashflow nach Steuern / Monat", kpi="CF", fmt=C.NUMFMT["eur"],
             sub='="Jahr 1 · ab Jahr 2: "&FIXED(' + C.CF_YEAR2 + ',0)&" €"'),
        dict(c1="E", c2="F", lr=11, label="Steuerwirkung / Monat · Jahr 1", fmt=SIGNED_EFFECT, neg=False,
             sub='=IF(Steuer_J1<0,"Erstattung","Zahlung")&" · Steuersatz Jahr 1: "&FIXED(Steuersatz_J1*100,1)&" %"'),
        dict(c1="C", c2="D", lr=14, label="EK-Rendite Jahr 1", kpi="EKR", fmt=C.NUMFMT["pct1"],
             sub="Vermögenszuwachs Jahr 1 / Eigenkapital"),
        dict(c1="E", c2="F", lr=14, label="IRR nach Steuern", kpi="IRR", fmt=C.NUMFMT["pct1"],
             sub='="Haltedauer "&Haltedauer&" Jahre · Multiple "&FIXED(EK_Multiple,2)&"×"'),
        dict(c1="C", c2="D", lr=17, label=None, fmt=C.NUMFMT["eur"], sub="Immobilienwert abzüglich Restschuld"),
        dict(c1="E", c2="F", lr=17, label=None, fmt=C.NUMFMT["eur"],
             sub="Cashflow nach Steuern nach Volltilgung (Prognose)"),
    ]
    for t in tiles:
        C.tile(ws, t["c1"], t["c2"], t["lr"], t["lr"] + 1, t["lr"] + 2, label=t["label"], kpi=t.get("kpi"),
               fmt=t["fmt"], sub=t["sub"], neg=t.get("neg"), gap_right=(t["c1"] == "C"), gap_top=(t["lr"] > 11))

    # Einordnung rechts: IRR (Ampel) und Cashflow (neutral), bündig mit den Kacheln
    first = ws["H29"].value.replace(S12_TAX_OLD, S12_TAX_NEW)   # Steuerwirkung wie in Kachel/Herleitung (P3-08)
    blank(ws["H29"])
    blank(ws["H37"])
    heights = {10: C.H_BAND, 11: C.H_TILE_LABEL, 12: C.H_TILE_VALUE, 13: C.H_TILE_SUB,
               14: C.H_CALLOUT_HEAD, 15: C.H_TILE_VALUE, 16: C.H_TILE_SUB,
               17: C.H_TILE_LABEL, 18: C.H_TILE_VALUE, 19: C.H_TILE_SUB}
    box(ws, heights, 10, "Eigenkapitalrendite (IRR)", "IRR", TEXT_IRR_S12, 11, fixed_end=12)
    box(ws, heights, 14, "Cashflow nach Steuern", None, first + "&CHAR(10)&CHAR(10)&" + EINORDNUNG_S12_ZUSATZ,
        15, fixed_end=19)
    heights[10] = C.H_BAND

    # Zeile 20: Herleitung · Diagramm · Weiter zur Auswertung
    head2(ws, 20, "C", "D", "Herleitung pro Monat")
    head2(ws, 20, "E", "F", "Cashflow nach Steuern", "Jahre 1–30 · € p. a.")
    head2(ws, 20, "H", "I", "Weiter zur Auswertung")
    heights[20] = ROW2
    for c in C.iter_cells(ws, "C", 20, "I", 20):
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
        heights[r] = ROW1
    ws["D22"].number_format = SIGNED_EFFECT
    C.sum_row(ws, 23, "C", "D", "result", neg=True)
    ws["D26"].number_format = C.NUMFMT["mult2"]
    for c in (ws["C24"], ws["D24"]):
        c.alignment = C.align(c.alignment.horizontal, "bottom", 1)
    heights[24] = ROW2
    chart = charts.get(7)
    if chart is not None:
        set_anchor(chart, "E", 21, "F", 26)
    for i, (word, rest, target) in enumerate(LINKS_S12):
        cell = ws.cell(21 + i, 8)
        rich_link(cell, word, rest, target, tooltip=f"Weiter: {word}")
        row_line(ws, 21 + i, "H", "I")
    for r, hh in heights.items():
        h[r].height = hh
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
        grid(ws, GRID_KPI if n in (8, 12) else GRID)
        unmerge_from(ws, 10)
        reset_styles(ws, 10, old_max)
        header(ws, n)
        if n == 8:
            end = page_s08(ws, names)
        elif n == 12:
            end = page_s12(ws, names)
        else:
            sp = SPEC[n]
            drop_cf(ws, TEMPLATE_CF)
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
        C.cf_close(ws)
