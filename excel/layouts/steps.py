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
from openpyxl.styles import Alignment, Border, Font
from openpyxl.worksheet.hyperlink import Hyperlink

import core as C

# ============================================================================ Raster und Namen
# Schrittseiten: C Position · D Eingabe (einheitlich breit) · E nur Einheit · F Hinweis · G Rinne · H:I Ergebnis
# Runde 3 (P13/P39/P02): ein Drittelraster für alle zwölf Seiten – C:D = E:F(+G) = H:I = 65 Zeichen (455 px).
#   Zurück-Button C:D und Weiter-Button H:I sind pixelgleich breit und stehen auf allen Schrittseiten an derselben
#   Stelle; „Übersicht“ (Textlink) sitzt genau in der Mitte (E:G). Spalte A 2,5 → linke Satzkante wie B-Blätter (A+B = 4,5).
# Runde 4 (P2-08): Eingabefeld über D:E (eine Breite, Einheit im Zahlenformat) – E 11 / F 54 geben Auswahltexten Platz.
GRID = (("A", 2.5), ("B", 2), ("C", 31), ("D", 34), ("E", 11), ("F", 54), ("G", 3), ("H", 43), ("I", 22))
# Ergebnisseiten S08/S12: zwei gleich breite Kachelspalten C:D | E:F (Kontext links, Status-Chip rechts)
GRID_KPI = (("A", 2.5), ("B", 2), ("C", 33), ("D", 32), ("E", 33), ("F", 32), ("G", 3), ("H", 43), ("I", 22))
LONG = ["Objekt", "Kaufpreis & Miete", "Kaufnebenkosten", "Kaufpreisaufteilung", "Maßnahmen & Reserve",
        "Bewirtschaftung", "Finanzierung", "Zwischenergebnis", "Steuern", "Abschreibung", "Prognose & Exit",
        "Ergebnis"]
GAP = C.H_GAP            # genau eine Leerzeile zwischen Komponenten
CHART_MIN = 180          # Mindesthöhe eines Diagramms (pt)
PIE_MIN = 150            # Mindesthöhe eines Kreises in der rechten Spalte (pt)
PIE_CHART_MIN = 190      # P17: Kreisdiagramme links über C:F, Kreis ≥ 200 px
ROW1, ROW2, ROW3 = C.H_STEP_ROW, C.H_STEP_ROW2, 47
NBSP = "\u00a0"
# Steuerwirkung in der Cash-Sicht (P07: + = Geld fließt zu). Die Formel liefert die Steuer-Sicht (+ Zahlung), die Anzeige
# dreht das Vorzeichen nur per Format um: Erstattung „+709 €“, Zahlung „−709 €“ (typografisches Minus, P23).
SIGNED_EFFECT = '"\u2212"#,##0" €";"+"#,##0" €";"–"'

UNITS = {"€": None, "%": None, "€/Monat": "pro Monat", "€ p. a.": "pro Jahr", "€/m²": "je m²",
         "€/m² p. a.": "je m² p. a.", "% Darlehen": None, "% der Miete": "der Miete",
         "% Verkaufspreis": None}


# ============================================================================ Ampel-Boxen (P1-03)
def val_text(kpi):
    """Ist-Wert als Text für die Status-Pill (suffix von core.callout_box): geschütztes Leerzeichen vor der Einheit,
    typografisches Minus."""
    return {"BMR": f'FIXED(Bruttomietrendite*100,1)&"{NBSP}%"',
            "NMR": f'FIXED(S06_NMR*100,1)&"{NBSP}%"',
            "DSCR": 'FIXED(DSCR_J1,2)&"×"',
            "IRR": f'FIXED(EK_IRR*100,1)&"{NBSP}%"',
            "CFV": f'IF(CF_vSt_J1<0,"{C.MINUS}","")&FIXED(ABS(CF_vSt_J1/12),0)&"{NBSP}€"'}[kpi]


VALUE_REF = {"BMR": "Bruttomietrendite", "NMR": "S06_NMR", "DSCR": "DSCR_J1", "IRR": "EK_IRR", "CFV": "CF_vSt_J1"}


def nbsp(text):
    """P45: Zahl und Einheit nie trennen – geschütztes Leerzeichen vor „%“/„€“ in Textbausteinen der Anzeigeformeln
    (auch „…&" % …“ hinter FIXED) und typografisches Minus vor Beträgen in Texten."""
    if not isinstance(text, str):
        return text
    text = re.sub(r'"[ ](%|€)', '"' + NBSP + r'\1', text)
    text = re.sub(r'(\d)[ ](%|€)', r'\1' + NBSP + r'\2', text)
    text = re.sub(r'§[ ](?=\d)', '§' + NBSP, text)          # „§ 82b“ nie am Zeilenende trennen
    return text


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
                'FIXED(CF_vSt_J1/12,0)&" € pro Monat vor Steuern als Puffer für Rücklagen, Mietausfall oder '
                'Sondertilgungen. Was nach Steuern bleibt, zeigen die Schritte 9 bis 12.","Es fehlen "&'
                'FIXED(-CF_vSt_J1/12,0)&" € pro Monat: Der Einnahmenüberschuss deckt die Rate an die Bank nicht, die '
                'Lücke kommt aus anderen Einkünften. Ob sich das Objekt trotzdem lohnt, zeigen Steuereffekt '
                '(Schritte 9–10) und Wertentwicklung (Schritt 11).")')
TEXT_IRR_S12 = ('=IFERROR(IF(EK_IRR>=Ampel_IRR_gruen,"Über "&Haltedauer&" Jahre verzinst sich das Eigenkapital mit "&'
                'FIXED(EK_IRR*100,1)&" % p. a. nach Steuern – das Ziel von "&FIXED(Ampel_IRR_gruen*100,1)&" % ist '
                'erreicht.",IF(EK_IRR>=Ampel_IRR_gelb,"Über "&Haltedauer&" Jahre verzinst sich das Eigenkapital mit "&'
                'FIXED(EK_IRR*100,1)&" % p. a. nach Steuern – solide, aber unter dem Ziel von "&'
                'FIXED(Ampel_IRR_gruen*100,1)&" %. Hebel: Kaufpreis, Miete, Finanzierung und Haltedauer.","Über "&'
                'Haltedauer&" Jahre verzinst sich das Eigenkapital nur mit "&FIXED(EK_IRR*100,1)&" % p. a. nach '
                'Steuern – unter der Mindestschwelle von "&FIXED(Ampel_IRR_gelb*100,1)&" %. Kaufpreis, Miete und '
                'Finanzierung prüfen."))&" Aus 1 € Eigenkapital werden bis zum Verkauf "&FIXED(EK_Multiple,2)&" €.",'
                '"Die Eigenkapitalrendite lässt sich mit diesen Annahmen nicht berechnen.")')

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
        fmts={15: C.NUMFMT["qm"], 16: C.NUMFMT["qm"]}, units={15: None, 16: None},
        band_right="beim Kauf",
        info=(21, "Alle gelben Felder enthalten Beispielwerte – einfach überschreiben."),
        nav=24, foot=26),
    2: dict(
        inputs=range(12, 18), indent2=(13, 15),
        labels={13: "davon bewegliche Gegenstände", 14: "Nutzungsdauer der Gegenstände",
                15: "davon Anteil Erhaltungsrücklage", 17: "Umlagefähige Betriebskosten"},
        fmts={14: C.NUMFMT["years_n"]}, units={14: None}, band_right="Kauf · Jahr 1",
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
        tile=dict(kpi="BMR", value="=Bruttomietrendite", fmt=C.NUMFMT["pct1"],
                  sub='="Kaufpreisfaktor "&FIXED(Kaufpreisfaktor,1)&"×  ·  " & ' + C.threshold_text("BMR")),
        nav=26, foot=28),
    3: dict(
        inputs=range(12, 17), optional=(12,),
        labels={12: "Grunderwerbsteuer abweichend", 14: "Grundbuchamt (Umschreibung)",
                16: "Sonstige Erwerbsnebenkosten"},
        fmts={12: C.NUMFMT["pct2_in"], 13: C.NUMFMT["pct2_in"], 14: C.NUMFMT["pct2_in"], 15: C.NUMFMT["pct2_in"]},
        band_right="beim Kauf",
        hints={12: "optional · leer = Satz des Bundeslands (siehe Ergebnis)",
               16: "Gutachten, Fahrten – sofern objektbezogen"},
        results={11: ("n", None, None, None), 12: ("n", None, None, None), 13: ("n", None, None, None),
                 14: ("n", None, None, None), 15: ("n", None, None, None), 16: ("i", None, None, None),
                 17: ("s", None, None, C.NUMFMT["pct1"]), 18: ("f", None, None, None)},
        callout=dict(head=20, src="H21", title="Steuerliche Behandlung"),
        charts=[dict(side="L", title="Zusammensetzung der Kaufnebenkosten", unit="Anteile in %", pie=True)],
        tile=dict(label="Kaufnebenkosten gesamt", value="=NK_Summe",
                  sub=f'=IFERROR("je m² Wohnfläche: "&FIXED(NK_Summe/Wohnflaeche,0)&"{NBSP}€","")'),
        nav=44, foot=46, hide=(60, 66)),
    4: dict(
        inputs=range(12, 16), text=(12,),
        inactive={14: "KPA_Idx<>2", 15: "KPA_Idx<>3"},
        labels={15: "Gebäudewert (Gutachten / Vertrag)"},
        hints={12: "Standard: BMF-Arbeitshilfe · Gutachten möglich",
               13: "lt. Gutachterausschuss / BORIS (Methode 1)",
               14: "nur Methode „Bodenanteil in %“",
               15: "nur Methode „Gebäudewert“ (ohne Inventar)"},
        results={11: ("n", "Kaufpreis ohne Inventar/Rücklage", None, None), 12: ("n", None, None, None),
                 13: ("n", None, None, C.NUMFMT["pct1"]),
                 14: ("n", "AK Grund und Boden (inkl. NK)", None, None),
                 15: ("f", "= AfA-Basis Gebäude (inkl. NK)", None, None),
                 16: ("m", "nachrichtlich: bewegliche Gegenstände", None, None)},
        callout=dict(head=18, src="H19", title="Gebäudeanteil und AfA"),
        band_right="beim Kauf",
        charts=[dict(side="L", title="Aufteilung des Kaufpreises", unit="inkl. Inventar und Rücklage · Anteile in %",
                     pie=True)],
        tile=dict(label="AfA-Basis Gebäude", value="=AK_Gebaeude",
                  sub=f'="Reguläre Gebäude-AfA Jahr 1: "&FIXED(INDEX(Steuern!$D$49:$AQ$49,1),0)&"{NBSP}€"'),
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
        units={17: None}, fmts={17: C.NUMFMT["years_n"]}, band_right="Jahre 1–3",
        results={11: ("n", "Maßnahmen Jahre 1–3 (netto)", None, None),
                 12: ("n", None, None, None),
                 13: ("t", "Steuerliche Behandlung",
                      '=IF(HK_Flag=1,"Aktiviert (AfA)",IF(Verteilung_eff>1,"Verteilt (§ 82b)","Sofortabzug"))', None),
                 14: ("n", "Maßnahmen + Sonderumlage Jahr 1", None, None),
                 15: ("i", None, None, None), 16: ("f", None, None, None)},
        callout=dict(head=18, src="H19", title=f"Die 15{NBSP}%-Grenze",
                     conditions=[("$I$11>$I$12", "amber"), ("ISNUMBER($I$11)", "green")]),
        charts=[dict(side="L", title="Zusammensetzung der Gesamtinvestition", unit="Anteile in %", pie=True)],
        tile=dict(label=C.KPI_LABELS["EK"], value="=EK_Bedarf_gesamt",
                  sub=f'=IFERROR("Anteil an der Gesamtinvestition "&FIXED(EK_Bedarf_gesamt/Gesamtinvestition*100,1)&"{NBSP}%","")'),
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
        band_right="Jahr 1 · pro Monat", units={18: None}, fmts={18: '[=1]0" Monat";0" Monate"'},
        callout=dict(head=17, src="H18", title="Nettomietrendite", kpi="NMR", prefix=TEXT_NMR_PREFIX),
        charts=[dict(side="L", title="Bewirtschaftungskosten Jahr 1", unit="€ p. a. · Anteile", pie=True)],
        tile=dict(label="Einnahmenüberschuss (NOI) / Monat", value="=INDEX(Projektion!$D$29:$AQ$29,1)/12",
                  sub=f'="Jahr 1 · "&FIXED(INDEX(Projektion!$D$29:$AQ$29,1),0)&"{NBSP}€ p.{NBSP}a."'),
        nav=41, foot=43, hide=(60, 66)),
    7: dict(
        inputs=range(12, 22), note=22,
        labels={16: "Darlehen I – Anschlusszins", 18: "Darlehen I – tilgungsfreie Jahre", 19: "Grundschuldbestellung",
                20: "Bearbeitungs-/Vermittlungsgebühr"},
        units={15: None, 18: None}, band_right="Kauf · Rate Jahr 1",
        fmts={13: C.NUMFMT["pct2_in"], 14: C.NUMFMT["pct2_in"], 15: C.NUMFMT["years_n"], 16: C.NUMFMT["pct2_in"],
              18: C.NUMFMT["years_n"], 19: C.NUMFMT["pct2_in"], 21: C.NUMFMT["pct2_in"]},
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
                     text=TEXT_DSCR_S07),
        # P1-09: Leitwert-Kachel wie S03–S11; das Zinsänderungsrisiko steht kompakt in der Fußzeile
        tile=dict(kpi="RATE", value="=Kapitaldienst_Monat_J1",
                  sub=f'="Jahr 1  ·  ab Jahr "&(Zinsbindung_I+1)&" je +1 %-Pkt. Zins: +"&FIXED(Restschuld_ZB_I*0.01/12,0)&"{NBSP}€ / Monat"'),
        charts=[dict(side="L", title="Restschuld am Jahresende", unit="Jahre 1–35 · T€ · Jahresende", chart_col=7,
                     hz=181, empty=(181, "kein Darlehen II", "MAX")),
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
        charts=[dict(side="L", title="Steuerliches Ergebnis und Steuer", unit="Jahre 1–20 · T€ · Steuer: + Zahlung / − Erstattung",
                     new=True, hz=196)],
        tile=dict(label="Steuerwirkung / Monat (Jahr 1)", value="=Steuer_J1/12", fmt=SIGNED_EFFECT, neg=False,
                  sub=f'=IF(Steuer_J1<0,"Erstattung","Zahlung")&"  ·  Cash-Sicht wie Schritt 12"'),
        nav=27, foot=29),
    10: dict(
        inputs=range(12, 18), text=(12, 14, 15, 16),
        inactive={13: "AfA_Idx<>6", 14: "AfA_Idx<>5", 17: 'Denkmal<>"Ja"'},
        units={13: None}, fmts={13: C.NUMFMT["years_n"]},
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
        charts=[dict(side="L", title="Abschreibungen nach Komponenten", unit="Jahre 1–20 · T€ p. a.", hz=203,
                     empty=(203, "keine Sonder-AfA", "SUM"))],
        tile=dict(label="Abschreibungen / Monat (Jahr 1)", value="=INDEX(Steuern!$D$56:$AQ$56,1)/12",
                  sub=f'="Jahr 1 gesamt: "&FIXED(INDEX(Steuern!$D$56:$AQ$56,1),0)&"{NBSP}€ p.{NBSP}a."'),
        nav=45, foot=47),
    11: dict(
        inputs=range(12, 19), text=(18,), optional=(17,),
        inactive={18: "Rechtsform_Idx>=2"},
        labels={16: "Verkaufskosten", 17: "Verkaufspreis manuell"},
        units={15: None, 16: None}, fmts={15: C.NUMFMT["years_n"]},
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
                 17: ("f", "= IRR n. St. (Eigenkapitalrendite)", None, C.NUMFMT["pct1"])},
        band_right="Verkauf",
        callout=dict(head=19, src="H20", title="Verkauf und Steuern"),
        charts=[dict(side="L", title="Immobilienwert, Restschuld und Nettovermögen", unit="Jahre 1–35 · T€ · Jahresende", hz=179)],
        tile=dict(kpi="IRR", label=C.kpi_label("IRR", caps=True, formula=True), value="=EK_IRR", fmt=C.NUMFMT["pct1"],
                  sub="=" + C.threshold_text("IRR")),
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


# ---------------------------------------------------------------------------- Innenabstand rechts für Formeltexte
# Runde 5: Excel kennt bei linksbündigem Text keinen Einzug rechts – Formeltexte der Einordnungs-Boxen liefen deshalb
# bis an die rechte Boxkante. Die Anzeigeformel wird so umgebaut, dass jeder Zweig feste Umbrüche („\n“ im
# Textliteral) auf Boxbreite − Einzug links − Innenabstand rechts bekommt. Dynamische Teile (FIXED, Namen) werden mit
# einer eher zu breiten Stellvertreterzahl bemessen – echte Werte sind höchstens schmaler, Excel bricht dann nie
# zusätzlich um. IF-Zweige innerhalb einer Verkettung werden dafür ausmultipliziert (A&IF(c,x,y)&B →
# IF(c,A&x&B,A&y&B), anzeigegleich). Nur für nicht referenzierte Anzeigeformeln.
_FN_RE = re.compile(r"(IFERROR|IF)\((.*)\)", re.S)


class _FN:   # noqa: N801 – wie ein re-Muster benutzt: fullmatch nur, wenn die erste Klammer am Ende schließt
    @staticmethod
    def fullmatch(expr):
        m = _FN_RE.fullmatch(expr)
        if not m:
            return None
        depth, q = 0, False
        start = len(m.group(1))
        for k in range(start, len(expr)):
            ch = expr[k]
            if ch == '"':
                q = not q
            elif not q and ch == "(":
                depth += 1
            elif not q and ch == ")":
                depth -= 1
                if depth == 0:
                    return m if k == len(expr) - 1 else None
        return None


def _placeholder(expr):
    """Stellvertreter-Text eines dynamischen Formelteils (eher breit geschätzt)."""
    e = expr.replace(" ", "")
    if e.upper() == "CHAR(10)":
        return "\n"
    m = re.findall(r"FIXED\(.*?,(\d)\)", e)
    if "FIXED(" in e:
        d = int(m[-1]) if m else 0
        neg = "" if "ABS(" in e or "-" not in e and "SUBSTITUTE" not in e else "−"
        if "*100" in e:
            return "00.0" if d >= 1 else "000"
        if d >= 2:
            return "0.00"
        if d == 1:
            return "00.0"
        return neg + ("0,000" if "/12" in e else "000,000")
    if e.endswith("_Txt"):                 # Textnamen, z. B. Volltilgung_Txt „Jahr 32 (2057)“
        return "Jahr 00 (0000)"
    if "YEAR(" in e or re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", e):
        return "0000"
    return "0000000"


def _is_lit(p):
    return len(p) >= 2 and p.startswith('"') and p.endswith('"')


def _expand(expr, budget):
    """Formel → Liste der Zweige als (Aufbau, Blätter): Aufbau ist eine Funktion, die aus den umbrochenen Blättern die
    Formel zurückbaut. Liefert None, wenn die Struktur zu verzweigt oder unbekannt ist."""
    expr = expr.strip()
    parts = _concat_parts(expr)
    if len(parts) == 1:
        m = _FN.fullmatch(expr)
        if m:
            args = _split_args(m.group(2))
            fn = m.group(1)
            if fn == "IF" and len(args) == 3:
                idx = (1, 2)
            elif fn == "IFERROR" and len(args) == 2:
                idx = (0, 1)
            else:
                return None
            subs = []
            for i in idx:
                s = _expand(args[i], budget)
                if s is None:
                    return None
                subs.append(s)
            leaves = [lf for s in subs for lf in s[1]]

            def build(new_leaves, subs=subs, args=args, idx=idx, fn=fn):
                out, k = list(args), 0
                for i, s in zip(idx, subs):
                    n = len(s[1])
                    out[i] = s[0](new_leaves[k:k + n])
                    k += n
                return f"{fn}(" + ",".join(a.strip() for a in out) + ")"
            return build, leaves
        return (lambda nl: "&".join(nl[0])), [[expr]]
    # Verkettung: erstes IF-/IFERROR-Glied ausmultiplizieren
    for i, p in enumerate(parts):
        mm = _FN.fullmatch(p.strip())
        if mm:
            args = _split_args(mm.group(2))
            fn = mm.group(1)
            idx = (1, 2) if fn == "IF" and len(args) == 3 else ((0, 1) if fn == "IFERROR" and len(args) == 2 else None)
            if idx is None:
                return None
            for j in idx:
                top = re.sub(r'"[^"]*"', "", args[j])
                while re.search(r"\([^()]*\)", top):
                    top = re.sub(r"\([^()]*\)", "", top)
                if re.search(r"[<>=]", top):
                    return None
            new_args = list(args)
            for j in idx:
                new_args[j] = "&".join(parts[:i] + [args[j].strip()] + parts[i + 1:])
            budget[0] -= 1
            if budget[0] < 0:
                return None
            return _expand(f"{fn}(" + ",".join(a.strip() for a in new_args) + ")", budget)
    return (lambda nl: "&".join(nl[0])), [parts]


def _wrap_leaf(parts, avail, size):
    """Ein Zweig (Glieder einer Verkettung): Leerzeichen in Textliteralen dort durch „\n“ ersetzen, wo die Zeile die
    Breite `avail` überschreiten würde. Liefert (neue Glieder, Zeilenzahl)."""
    chars = []            # (Zeichen, Glied, Position im Literal | None)
    lits = {}
    for pi, p in enumerate(parts):
        if _is_lit(p):
            s = p[1:-1].replace('""', '"')
            lits[pi] = list(s)
            chars += [(ch, pi, k) for k, ch in enumerate(s)]
        else:
            chars += [(ch, pi, None) for ch in _placeholder(p)]
    n_lines = 0
    para = []
    paras = []
    for c in chars:
        if c[0] == "\n":
            paras.append(para)
            para = []
        else:
            para.append(c)
    paras.append(para)
    for para in paras:
        words, cur = [], []
        for c in para:
            if c[0] == " " and c[2] is not None:
                words.append((cur, c))
                cur = []
            else:
                cur.append(c)
        words.append((cur, None))
        n_lines += 1
        line = ""
        prev_space = None
        for w, space_after in words:
            wt = "".join(ch for ch, _, _ in w)
            cand = (line + " " + wt) if line else wt
            if line and C.text_width(cand, size) > avail and prev_space is not None:
                lits[prev_space[1]][prev_space[2]] = "\n"
                n_lines += 1
                line = wt
            else:
                line = cand
            prev_space = space_after
    out = []
    for pi, p in enumerate(parts):
        if pi in lits:
            out.append('"' + "".join(lits[pi]).replace('"', '""') + '"')
        else:
            out.append(p)
    return out, n_lines


def wrap_formula(value, width_px, size=C.T_SMALL, indent=1, right_pad=None):
    """Anzeigeformel mit festen Umbrüchen (Innenabstand rechts). Liefert (Formel, Zeilenzahl) oder (value, None)."""
    if not (isinstance(value, str) and C.is_formula(value)):
        return value, None
    pad = 9 * max(indent, 1) if right_pad is None else right_pad
    avail = max(C.cell_inner_px(width_px, indent) * 0.98 - pad, 20)
    try:
        res = _expand(value[1:], [12])
        if res is None:
            return value, None
        build, leaves = res
        new, n = [], 0
        for lf in leaves:
            nl, k = _wrap_leaf(lf, avail, size)
            new.append(nl)
            n = max(n, k)
        f = "=" + build(new)
        if len(f) > 7500:
            return value, None
        return f, n
    except Exception as exc:  # noqa: BLE001 – im Zweifel die Formel unverändert lassen
        print(f"  steps: wrap_formula übersprungen ({exc})")
        return value, None


def wrap_runs(runs, avail):
    """Rich-Text-Läufe [(Text, Größe, fett, Farbe)] mit festen Umbrüchen auf `avail` px (Innenabstand rechts).
    Liefert (Läufe, Zeilenzahl)."""
    chars = [(ch, i) for i, r in enumerate(runs) for ch in r[0]]

    def width(seq):
        tot, k = 0.0, 0
        while k < len(seq):
            j = k
            while j < len(seq) and seq[j][1] == seq[k][1]:
                j += 1
            r = runs[seq[k][1]]
            tot += C.text_width("".join(ch for ch, _ in seq[k:j]), r[1], r[2])
            k = j
        return tot
    out = [list(r[0]) for r in runs]
    pos = {}
    k = 0
    for i, r in enumerate(runs):
        for j in range(len(r[0])):
            pos[k] = (i, j)
            k += 1
    line_start, last_space, n = 0, None, 1
    for idx, (ch, _) in enumerate(chars):
        if ch == " ":
            if width(chars[line_start:idx]) > avail and last_space is not None:
                i, j = pos[last_space]
                out[i][j] = "\n"
                line_start = last_space + 1
                n += 1
            last_space = idx
    if width(chars[line_start:]) > avail and last_space is not None and last_space >= line_start:
        i, j = pos[last_space]
        out[i][j] = "\n"
        n += 1
    return [("".join(o),) + tuple(r[1:]) for o, r in zip(out, runs)], n


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


def stack(heights, start, need, tol=3.0):
    """Zeilen ab `start` belegen, bis `need` pt erreicht sind. Gesetzte Zeilen zählen mit ihrer Höhe; die erste freie
    Zeile erhält genau den Rest (mind. 8 pt). Fehlen nur ≤ `tol` pt, endet der Block schon vorher (≈ 4 px weniger
    Innenraum statt einer Zusatzzeile). Liefert die letzte belegte Zeile."""
    r, acc = start, 0.0
    while True:
        if r in heights:
            acc += heights[r]
            if acc >= need - tol:
                return r
        else:
            if r > start and need - acc <= tol:
                return r - 1
            heights[r] = max(C.px_pt(need - acc), 8.0)
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


def head2(ws, row, c1, c2, title, unit=None, unit_formula=None):
    """Ebene-2-Kopf (Unterabschnitt, Linie) für Diagramm- und Tabellenblöcke: 8,5 pt Versalien, Einheit rechts.
    unit_formula: Anzeigeformel für die Meta-Zelle (z. B. Zusatz „ · kein Darlehen II“, wenn eine Reihe leer ist)."""
    C.section(ws, row, c1, c2, title, level=2, variant="line", meta=unit)
    if unit_formula:
        ws[f"{c2}{row}"].value = unit_formula


def horizon_years(row):
    """Jahre einer Zeitreihe auf „Diagramme“ – liest den Horizont aus layouts.diagramme (Agent H), damit der
    Diagrammkopf immer zur Achse passt."""
    try:
        from layouts import diagramme as D
        last = D.HORIZON.get(row)
        return C.col(last) - C.col("D") + 1 if last else None
    except Exception:  # noqa: BLE001 – Kopf fällt auf den Standard zurück
        return None


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


def box_lines(ws, value, c1="H", c2="I"):
    """Textzeilen (9 pt) des Körpers einer Einordnungs-Box: statischer Text wie gesetzt (feste Umbrüche), Formeln je
    Zweig (größter Wert aus eigener Zweigzerlegung und core.display_text)."""
    w = C.span_px(ws, c1, c2)
    if value in (None, ""):
        return 1
    if not C.is_formula(value):
        return C.lines_needed_metric(str(value), w, C.T_SMALL, False, 1)
    n = max(C.lines_needed_metric(t, w, C.T_SMALL, False, 1) for t in variants(value))
    class _Probe:   # Stellvertreter-Zelle nur zum Bemessen (legt keine Zelle im Blatt an)
        parent = ws
    _Probe.value = value
    try:
        t = C.display_text(_Probe)
        if t:
            n = max(n, C.lines_needed_metric(t, w, C.T_SMALL, False, 1))
    except Exception:  # noqa: BLE001
        pass
    return n


def box(ws, heights, head, title, kpi=None, text=None, body=None, fixed_end=None, draw=True, conditions=None,
        c1="H", c2="I"):
    """Einordnungs-Box (core.callout_box). Körperhöhe = Textzeilen × 12,5 + 8 pt (core.callout_height, P2-01): die
    Box belegt ab `body` nur so viele Zeilen, wie der Text braucht. Statischer Text bekommt feste Umbrüche mit
    Innenabstand rechts (core.hard_wrap), Formeltexte das typografische Minus (core.minus_text).
    Ampel-Boxen: Pill „● Urteil  ·  Ist-Wert“ (suffix), Statuskante links. draw=False: nur bemessen."""
    body = body or head + 1
    cell = ws.cell(body, C.col(c1))
    if text is not None:
        cell.value = nbsp(text)
    elif isinstance(cell.value, str):
        cell.value = nbsp(cell.value)
    w = C.span_px(ws, c1, c2)
    if isinstance(cell.value, str):
        cell.value = C.minus_text(cell.value) if C.is_formula(cell.value) else C.hard_wrap(cell.value, w, C.T_SMALL,
                                                                                            indent=1)
    n_wrap = None
    if isinstance(cell.value, str) and C.is_formula(cell.value):   # Runde 5: Innenabstand rechts auch für Formeln
        cell.value, n_wrap = wrap_formula(cell.value, w, C.T_SMALL, indent=1)
    heights[head] = max(heights.get(head, 0), C.H_CALLOUT_HEAD)
    need = C.callout_height(n_wrap or box_lines(ws, cell.value, c1, c2))
    end = fixed_end or stack(heights, body, need)
    if draw:
        kw = dict(kpi=kpi, value_ref=VALUE_REF[kpi], suffix=val_text(kpi)) if kpi else {}
        if conditions:
            kw = dict(conditions=conditions)
        C.callout_box(ws, c1, head, c2, body, end, title=title, fit=None, **kw)
    return end


def next_card(ws, heights, head, n, names):
    """„Als Nächstes“: Vorschau auf den nächsten Schritt (Name fett + Kurzbeschreibung), Link auf das Blatt."""
    nxt = names[n + 1]
    desc = ws.parent[nxt]["C7"].value
    desc = nbsp(desc) if isinstance(desc, str) and not C.is_formula(desc) else ""
    head2(ws, head, "H", "I", "Als Nächstes", f"Schritt {n + 1:02d} / 12")
    heights[head] = max(heights.get(head, 0), C.H_HEAD)
    body = head + 1
    for c in C.iter_cells(ws, "H", body, "I", body):
        c.fill, c.border = C.NOFILL, Border()
    cell = ws.cell(body, 8)
    # Runde 5: feste Umbrüche mit Innenabstand rechts (wie die Einordnungs-Boxen)
    avail = C.cell_inner_px(C.span_px(ws, "H", "I"), 1) * 0.98 - 9
    runs, n_l = wrap_runs([(LONG[n], C.T_BODY, True, C.NAVY), (f"  ·  {desc}", C.T_SMALL, False, C.MUTED)], avail)
    cell.value = C.rich(runs)
    cell.font = C.font(C.T_SMALL, False, C.MUTED)
    cell.alignment = C.align("left", "top", 1, wrap=True)
    need = C.grid_height(n_l * C.line_pt(C.T_BODY) + 5)
    return body, need


def nav_and_footer(ws, n, row, names):
    """Button-Reihe mit klarer Dreistufigkeit (P13), auf allen zwölf Seiten in denselben Zellen und Pixeln:
    Zurück C:D (sekundär, Outline) · „Übersicht: Leitfaden“ E:G (tertiär, Textlink genau in der Mitte) ·
    Weiter H:I (primär, gefüllt). Zurück und Weiter sind gleich breit (Drittelraster)."""
    set_h = ws.row_dimensions
    set_h[row - 2].height = GAP
    set_h[row - 1].height = GAP
    prev_cell = None
    if n == 1:   # S01: zurück zum Pflichtschritt „Kauf als“ (Doppelziel Leitfaden entfällt)
        prev_sheet, prev_text, prev_tip = "Start", "‹  Zurück: Start (Kauf als)", "Privat oder GmbH auf der Startseite wählen"
        prev_cell = "D23"
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
        dict(c1="C", c2="D", text=prev_text, target=prev_sheet, kind="secondary", tooltip=prev_tip, cell=prev_cell),
        dict(c1="E", c2="G", text="Übersicht: Leitfaden", target="Leitfaden", kind="secondary",
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


def unit_format(fmt, qualifier):
    """Bezugsangabe (pro Monat, pro Jahr, je m², der Miete) ins Zahlenformat der Eingabe (P3-09): „950 € / Monat“,
    „3,0 % der Miete“ – die Spalte EINHEIT entfällt, das Eingabefeld reicht über D:E."""
    suffix = {"pro Monat": " / Monat", "pro Jahr": " p. a.", "je m²": " / m²", "je m² p. a.": " / m² p. a.",
              "der Miete": " der Miete"}.get(qualifier)
    if suffix is None:
        return fmt
    if "%" in fmt:
        return f'0.0\\ %"{suffix}";"{C.MINUS}"0.0\\ %"{suffix}"'
    return f'#,##0" €{suffix}";"{C.MINUS}"#,##0" €{suffix}"'


def zero_dash_fmt(fmt):
    """Anzeige „–“ für 0 in inaktiven Feldern (Befund 25) – als bedingtes Zahlenformat, das Feld bleibt Eingabe."""
    if not fmt or fmt in ("General", "@"):
        return None
    if fmt.startswith("[=1]"):
        parts = fmt.split(";")
        return f'{parts[0]};[=0]"–";{parts[-1]}' if len(parts) == 2 else fmt
    try:
        return C.zero_dash(fmt)
    except Exception:  # noqa: BLE001
        return None


def inactive_rule(ws, ref, condition, numfmt=None):
    """Wie core.inactive_when (grau, bleibt editierbar), zusätzlich 0 → „–“."""
    ln = C.side("thin", C.LINE2)
    C.cf_rule(ws, ref, condition, font_=Font(color=C.INACTIVE_FG, bold=False), fill_=C.fill(C.INACTIVE_BG),
              border=Border(left=ln, right=ln, top=ln, bottom=ln), numfmt=numfmt)


def style_input_table(ws, sp):
    """Kopfzeile Z. 11 und Eingabezeilen: Beschriftung (C) · Eingabefeld D:E (eine Breite für alle Felder, Einheit und
    Bezugsangabe im Zahlenformat) · Hinweis (F). Liefert den Zeilenbedarf je Zeile (Anzahl Textzeilen)."""
    need = {}
    for c in C.iter_cells(ws, "C", 11, "F", 11):   # alte Spaltenköpfe (EINGABE/EINHEIT) der Vorlage
        c.value = None
    C.section(ws, 11, "C", "F", None, level=2, labels={"C": ("Position", "left"), "E": ("Eingabe", "right"),
                                                          "F": ("Hinweis", "left")})
    text_rows = set(sp.get("text", ()))
    inactive = sp.get("inactive", {})
    wc, wf = C.span_px(ws, "C", "C"), C.span_px(ws, "F", "F")
    wd = C.span_px(ws, "D", "E")
    for r in sp["inputs"]:
        lab, inp, unit, hint = (ws.cell(r, k) for k in (3, 4, 5, 6))
        if r in sp.get("labels", {}):
            C.set_text(lab, sp["labels"][r])
        if r in sp.get("hints", {}):
            C.set_text(hint, sp["hints"][r])
        qualifier = sp.get("units", {}).get(r, unit.value.strip() if isinstance(unit.value, str) else None)
        qualifier = UNITS.get(qualifier, qualifier)
        if r in sp.get("fmts", {}):   # P38/Befund 29: Einheit im Zahlenformat, eine Genauigkeit je Größe
            inp.number_format = sp["fmts"][r]
        if qualifier and isinstance(inp.value, (int, float)):
            inp.number_format = unit_format(inp.number_format, qualifier)
        if C.is_formula(unit.value):
            print(f"  steps: {ws.title}!E{r} enthält eine Formel – Eingabefeld bleibt einspaltig")
        else:
            blank(unit)
        ind = 2 if r in sp.get("indent2", ()) else 1
        lab.font = C.font(C.T_BODY, False, C.INK)
        lab.alignment = C.align("left", "center", ind, wrap=True)
        hint.font = C.font(C.T_SMALL, False, C.MUTED)
        hint.alignment = C.align("left", "center", 1, wrap=True)
        state = "linked" if r == sp.get("linked") else ("optional" if r in sp.get("optional", ()) else "required")
        is_text = r in text_rows
        for c in (inp, unit):
            C.input_style(c, state)
        size = C.T_BODY
        if is_text:   # P2-08: Auswahltexte einzeilig – erst 10 pt, sonst 9 pt (Listenwerte bleiben unverändert)
            t = C.display_text(inp) or ""
            if not C.fits(t, wd, C.T_BODY, state != "linked", 1, 0.97):
                size = C.T_SMALL
        inp.font = C.font(size, state != "linked", C.INPUT_FG)
        inp.alignment = C.align("left", "center", 1, wrap=True) if is_text else C.align("right", "center", 1)
        if not C.is_formula(unit.value):
            C.safe_merge(ws, "D", r, "E", r)
        for c in (lab, hint):
            row_line(ws, r, C.L(c.column), C.L(c.column))
        if r in inactive:   # inaktiver Zustand: Hinweis beginnt mit „inaktiv – “ (nur solange inaktiv)
            base = hint.value if isinstance(hint.value, str) and not C.is_formula(hint.value) else None
            if base:
                hint.value = f'=IF({inactive[r]},"{C.INACTIVE_PREFIX}","")&{lit(base)}'
        n = max(lines(lab.value, wc, C.T_BODY, False, ind),
                lines(C.display_text(hint) if C.is_formula(hint.value) else hint.value, wf, C.T_SMALL, False, 1),
                lines(C.display_text(inp), wd, size, True, 1) if is_text else 1)
        need[r] = n
        if n > 1:
            print(f"  steps: {ws.title} Z. {r} zweizeilig ({lab.value!r} | {hint.value!r} | {C.display_text(inp)!r})")
    for r, cond in inactive.items():
        inactive_rule(ws, f"D{r}:E{r}", cond, zero_dash_fmt(ws.cell(r, 4).number_format))
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
        elif val.data_type == "f" and val.number_format == "@":
            val.number_format = "General"      # Befund 34: Textformel nie im Textformat (F2 + Enter)
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
    unit, uf = spec.get("unit"), None
    years = horizon_years(spec["hz"]) if spec.get("hz") else None
    if unit and years:
        unit = re.sub(r"Jahre 1–\d+", f"Jahre 1–{years}", unit)
    if spec.get("empty") and years:   # leere Reihe fehlt in der Legende → Hinweis im Kopf (Wunsch H)
        row, word, func = spec["empty"]
        last = C.L(C.col("D") + years - 1)
        uf = f'={lit(unit)}&IF({func}(Diagramme!$D${row}:${last}${row})>0,"","  ·  {word}")'
    head2(ws, head_row, c1, c2, spec["title"], unit, uf)
    heights[head_row] = max(heights.get(head_row, 0), C.H_HEAD)
    for c in C.iter_cells(ws, c1, head_row, c2, head_row):   # Kopf sitzt immer auf seiner Linie (auch in hohen Zeilen)
        c.alignment = C.align(c.alignment.horizontal or "left", "bottom", 1)
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
        C.text_link(ws["F12"], "Kauf als auf der Startseite ändern  ›", "Start", "D23", size=C.T_SMALL, bold=True,
                    tooltip="Privat oder GmbH wählen Sie auf der Startseite („Kauf als“)")
        ws["F12"].alignment = C.align("left", "center", 1)
    inputs_end = max(sp["inputs"])
    results_end = max(sp["results"])
    heights[11] = C.H_HEAD if need_r.get(11, 1) <= 1 else ROW2
    for r in range(12, max(inputs_end, results_end) + 1):
        heights[r] = row_height(max(need_l.get(r, 1), need_r.get(r, 1)))

    # Hinweise unter der Eingabetabelle – EINE Bauform (P34, core.note): gelbe Hinweisleiste, Text ohne Link,
    # Aktionslink in eigener Zelle rechts (F); genau eine Leerzeile Abstand zur Tabelle.
    left_end = inputs_end
    if sp.get("note"):   # S07: zweites Darlehen
        blank(ws.cell(sp["note"], 3))
        r = inputs_end + 2
        link_row = next((rr for rr in range(95, 125) if str(ws.parent["Eingaben"].cell(rr, 2).value or "")
                         .startswith("Darlehen II")), 104)
        heights[r - 1] = GAP
        C.note(ws, r, "C", "F", "Ein zweites Darlehen (z. B. KfW) erfassen Sie auf dem Blatt „Eingaben“.", label="Hinweis", link_text="Darlehen II erfassen  ›",
               target_sheet="Eingaben", target_cell=f"C{link_row}", link_col="F",
               tooltip="Darlehen II wird nur auf dem Blatt „Eingaben“ erfasst")
        heights[r] = C.text_row_height(1)
        left_end = r
    if sp.get("info"):   # S01: Beispielwerte, Link auf „Kauf als“
        r, text = sp["info"]
        heights[r - 1] = GAP
        C.note(ws, r, "C", "F", text, label="Hinweis", link_text="Kauf als auf der Startseite wählen  ›",
               target_sheet="Start", target_cell="D23", link_col="F",
               tooltip="Privat oder GmbH – Auswahl auf der Startseite („Kauf als“)")
        heights[r] = C.text_row_height(1)
        left_end = r

    # Diagrammkopf links reservieren (die Einordnung rechts läuft über diese Zeilen)
    ch_specs = sp.get("charts", [])
    left_chart = next((c for c in ch_specs if c["side"] == "L"), None)
    pie = next((c for c in ch_specs if c["side"] == "R"), None)
    chead = None
    if left_chart:
        chead = left_end + 2
        heights.setdefault(left_end + 1, GAP)
        heights[chead] = max(heights.get(chead, 0), C.H_HEAD)

    # Einordnung (rechts, eine Leerzeile unter dem Ergebnis) – Höhe aus dem Text (P2-01)
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
    if co.get("text"):
        text = co["text"]
    elif co.get("prefix"):
        old = src_value[1:] if C.is_formula(src_value) else lit(str(src_value))
        text = "=" + co["prefix"] + "&" + old
    else:
        text = src_value
    callout_end = box(ws, heights, head, co["title"], co.get("kpi"), text, body, conditions=co.get("conditions"))

    # Rechte Spalte, feste Folge auf allen Schrittseiten: Box · Leitwert-Kachel (core.tile 18/30/20) · [Kreis] ·
    # „Als Nächstes“. Genau eine 12-pt-Fuge zwischen den Blöcken; braucht das Diagramm links mehr Höhe, geht der Rest
    # in die Fuge vor „Als Nächstes“ (P1-02) bzw. in den Kreis (S07) – nie in die Kachel.
    r = callout_end + 1
    heights.setdefault(r, GAP)
    if sp.get("tile"):
        sr = place_tile(ws, heights, r + 1, sp["tile"])
        r = sr + 1
        heights[r] = GAP
    flex = r
    if pie:
        rhead = r + 1
        heights[rhead] = max(heights.get(rhead, 0), C.H_HEAD)
        pie_end = stack(heights, rhead + 1, PIE_MIN)
        chart_block(ws, heights, charts, pie, rhead, "H", "I", end_row=pie_end)
        flex = pie_end
        r = pie_end + 1
        heights[r] = GAP
    cbody, cneed = next_card(ws, heights, r + 1, n, names)
    heights[cbody] = max(cneed, heights.get(cbody, 0))
    C.safe_merge(ws, "H", cbody, "I", cbody)
    right_end = cbody

    chart_end = 0
    if left_chart:
        cmin = sp.get("chart_min", PIE_CHART_MIN if left_chart.get("pie") else CHART_MIN)
        have = sum(heights.get(rr, 15) for rr in range(chead + 1, right_end + 1))
        if have < cmin:
            if flex > chead:
                heights[flex] = C.px_pt(heights[flex] + cmin - have)
            else:
                heights[right_end + 1] = C.px_pt(cmin - have)
                right_end += 1
        chart_end = chart_block(ws, heights, charts, left_chart, chead, "C", "F", end_row=right_end)

    content_end = max(left_end, results_end, right_end, chart_end)
    for r, h in heights.items():
        ws.row_dimensions[r].height = h
    return nav_and_footer(ws, n, content_end + 3, names)


def place_tile(ws, heights, lr, t):
    """Leitwert-Kachel rechts (core.tile, Variante nach Blattrolle): Kopf 18 · Wert 30 · Fuß 20 pt. Liefert die
    Fußzeile."""
    vr, sr = lr + 1, lr + 2
    heights.update({lr: C.H_TILE_LABEL, vr: C.H_TILE_VALUE, sr: C.H_TILE_SUB})
    for c in C.iter_cells(ws, "H", lr, "I", sr):
        blank(c)
    C.tile(ws, "H", "I", lr, vr, sr, label=t.get("label"), value=t["value"], sub=nbsp(t.get("sub")),
           kpi=t.get("kpi"), fmt=t.get("fmt", C.NUMFMT["eur"]), neg=t.get("neg"), variant="light")
    return sr


# ============================================================================ Ergebnisseiten S08/S12
def result_rows(ws, heights, labels, r_head, row_h=C.H_STEP_ROW):
    """Herleitung links (C:D): gleichmäßige Zeilen (Tabellenzeile der Schrittseiten, P3-16), Haarlinien; Kopfzeile
    r_head mit 32 pt (Abstand zu den Kacheln, Text unten)."""
    for r, text in labels.items():
        lab, val = ws.cell(r, 3), ws.cell(r, 4)
        C.set_text(lab, text)
        lab.font, val.font = C.font(), C.font()
        lab.alignment = C.align("left", "center", 1)
        val.alignment = C.align("right", "center", 1)
        row_line(ws, r, "C", "D")
        heights[r] = row_h
    heights[r_head] = ROW2


def link_list(ws, row, links):
    """Linkliste rechts (H:I) – eine Zeile je Ziel, gleiche Zeilenhöhe wie die Herleitung links."""
    for i, (word, rest, target) in enumerate(links):
        cell = ws.cell(row + i, 8)
        rich_link(cell, word, rest, target, tooltip=f"Weiter: {word}")
        row_line(ws, row + i, "H", "I")


def tiles_block(ws, tiles):
    """KPI-Kacheln (core.tile, dunkel): Rinnen automatisch (horizontal = vertikal), Status als Chip rechts in der
    Fußzeile, Wertfarbe = Status (core, P11)."""
    for t in tiles:
        C.tile(ws, t["c1"], t["c2"], t["lr"], t["lr"] + 1, t["lr"] + 2, label=t.get("label"), kpi=t.get("kpi"),
               fmt=t.get("fmt", C.NUMFMT["eur"]), sub=t.get("sub"), neg=t.get("neg"))


def tile_heights(heights, first, rows):
    """Kachelreihen nach core (P1-02): Kopf 18 pt, Wert 30 pt, Fußzeile 20 pt."""
    for i in range(rows):
        lr = first + 3 * i
        heights.update({lr: C.H_TILE_LABEL, lr + 1: C.H_TILE_VALUE, lr + 2: C.H_TILE_SUB})


# ============================================================================ S08 Zwischenergebnis
LINKS_S08 = [("Kaufpreis & Miete", "Miete und Kaufpreis", 2), ("Bewirtschaftung", "nicht umlagefähige Kosten", 6),
             ("Finanzierung", "Eigenkapital, Zins und Tilgung", 7), ("Steuern", "Rechtsform und Steuersatz", 9),
             ("Sensitivität", "Break-even-Miete und Zinsänderung", "Sensitivität"),
             ("Cockpit", "alle Kennzahlen im Detail", "Cockpit")]


def page_s08(ws, names):
    n = 8
    for coord in ("C57", "E57", "H57", "C59", "C60", "H27", "H17", "I17", "H18"):
        blank(ws[coord])
    charts = {ch.anchor._from.col: ch for ch in ws._charts}
    h = ws.row_dimensions

    C.section(ws, 10, "C", "F", "Cashflow vor Steuern – Jahr 1")
    drop_cf(ws, TEMPLATE_CF + ["D22"])
    tiles_block(ws, [
        dict(c1="C", c2="D", lr=11, label="Nettokaltmiete Ist / Monat", sub="nach Mietausfall und Leerstand"),
        dict(c1="E", c2="F", lr=11, label="Einnahmenüberschuss (NOI) / Monat", sub="Miete abzüglich Bewirtschaftung"),
        dict(c1="C", c2="D", lr=14, kpi="RATE", sub="Zinsen und Tilgung · Jahr 1"),
        dict(c1="E", c2="F", lr=14, kpi="CFV", sub=f"Ziel ≥ 0{NBSP}€ · nach Rate, vor Steuern"),
    ])

    # Einordnung rechts: je Kennzahl eine Box (Cashflow · Kapitaldienstdeckung), bündig mit den Kachelreihen
    blank(ws["H28"])
    blank(ws["H34"])
    heights = {10: C.H_BAND}
    tile_heights(heights, 11, 2)
    box(ws, heights, 10, "Cashflow vor Steuern", "CFV", TEXT_CFV_S08, 11, fixed_end=12)
    box(ws, heights, 14, "Kapitaldienstdeckung (DSCR)", "DSCR", TEXT_DSCR_S08, 15, fixed_end=16)
    heights[10] = C.H_BAND

    # Zeile 17: Herleitung · Diagramm · Stellschrauben (drei Spalten, gleiche Zeilen 18–23)
    head2(ws, 17, "C", "D", "Herleitung pro Monat", "Jahr 1")
    head2(ws, 17, "E", "F", "Einnahmen vs. Ausgaben", f"€ / Monat · vor Steuern")
    head2(ws, 17, "H", "I", "Stellschrauben", "Cashflow verbessern")
    result_rows(ws, heights, {18: "Nettokaltmiete Ist", 19: "– Bewirtschaftung inkl. Rücklagen", 20: "– Zinsen",
                              21: "– Tilgung", 22: "= Cashflow vor Steuern / Monat",
                              23: "Kapitaldienstdeckung (DSCR)"}, 17)
    C.sum_row(ws, 22, "C", "D", "final")
    ws["D23"].number_format = C.NUMFMT["dscr"]
    for c in C.iter_cells(ws, "C", 17, "I", 17):
        c.alignment = C.align(c.alignment.horizontal or "left", "bottom", 1)
    chart = charts.get(7)
    if chart is not None:
        set_anchor(chart, "E", 18, "F", 23)
    link_list(ws, 18, [(w, r, names[t] if isinstance(t, int) else t) for w, r, t in LINKS_S08])
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
CF2 = f'IF({C.CF_YEAR2}<0,"{C.MINUS}","")&FIXED(ABS({C.CF_YEAR2}),0)&"{NBSP}€"'


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
    tiles_block(ws, [   # kanonische Labels und Reihenfolge (P20)
        dict(c1="C", c2="D", lr=11, kpi="CF", sub=f'="ab Jahr 2: "&{CF2}'),
        dict(c1="E", c2="F", lr=11, label="Steuerwirkung / Monat (Jahr 1)", fmt=SIGNED_EFFECT, neg=False,
             sub=nbsp('=IF(Steuer_J1<0,"Erstattung","Zahlung")&"  ·  Cash-Sicht  ·  Steuersatz "&'
                      'FIXED(Steuersatz_J1*100,2)&" %"')),
        dict(c1="C", c2="D", lr=14, kpi="EKR", fmt=C.NUMFMT["pct1"],
             sub=nbsp("=" + C.threshold_text("EKR") + '&"  ·  Vermögenszuwachs / Eigenkapital"')),
        dict(c1="E", c2="F", lr=14, kpi="IRR", label=C.kpi_label("IRR", caps=True, formula=True), fmt=C.NUMFMT["pct1"],
             sub=nbsp("=" + C.threshold_text("IRR") + '&"  ·  Multiple "&FIXED(EK_Multiple,2)&"×"')),
        dict(c1="C", c2="D", lr=17, sub="Immobilienwert abzüglich Restschuld"),
        dict(c1="E", c2="F", lr=17, sub="Cashflow nach Steuern nach Volltilgung (Prognose)"),
    ])

    # Einordnung rechts: IRR (Ampel) und Cashflow (neutral), bündig mit den Kacheln
    first = ws["H29"].value.replace(S12_TAX_OLD, S12_TAX_NEW)   # Steuerwirkung wie in Kachel/Herleitung (P3-08)
    blank(ws["H29"])
    blank(ws["H37"])
    heights = {10: C.H_BAND}
    tile_heights(heights, 11, 3)
    box(ws, heights, 10, "IRR n. St. (Eigenkapitalrendite)", "IRR", TEXT_IRR_S12, 11, fixed_end=12)
    box(ws, heights, 14, "Cashflow nach Steuern", None, first + "&CHAR(10)&CHAR(10)&" + EINORDNUNG_S12_ZUSATZ,
        15, fixed_end=19)
    heights[10] = C.H_BAND

    # Zeile 20: Herleitung · Diagramm · Weiter zur Auswertung (gleiche Zeilen 21–26)
    head2(ws, 20, "C", "D", "Herleitung pro Monat", "Jahr 1")
    try:
        from layouts import diagramme as D
        cf_years = C.col(getattr(D, "CF_NACH_LAST", "AG")) - C.col("D") + 1
    except Exception:  # noqa: BLE001
        cf_years = 30
    head2(ws, 20, "E", "F", "Cashflow nach Steuern", f"Jahre 1–{cf_years} · T€ p. a. · rot = Zuschuss")
    head2(ws, 20, "H", "I", "Weiter zur Auswertung")
    result_rows(ws, heights, {21: "Cashflow vor Steuern / Monat", 22: "± Steuerwirkung / Monat (+ Erstattung)",
                              23: "= Cashflow nach Steuern / Monat", 24: C.KPI_LABELS["EK"],
                              25: "Gesamtertrag n. St. bis Verkauf", 26: C.KPI["MULT"]["label"]}, 20,
                row_h=C.H_BTN)   # 6 × 25,5 = 153 pt Plotzeilen für das Cashflow-Diagramm (P1-07)
    for c in C.iter_cells(ws, "C", 20, "I", 20):
        c.alignment = C.align(c.alignment.horizontal or "left", "bottom", 1)
    ws["D22"].number_format = SIGNED_EFFECT
    C.neg_red(ws, "D21")                     # P15: negatives Ergebnis in der Herleitung rot
    C.sum_row(ws, 23, "C", "D", "final")
    ws["D26"].number_format = C.NUMFMT["mult2"]
    chart = charts.get(7)
    if chart is not None:
        set_anchor(chart, "E", 21, "F", 26)
    link_list(ws, 21, LINKS_S12)
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
