# 🎈 Blank app template

A simple Streamlit app template for you to modify!

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://blank-app-template.streamlit.app/)

### How to run it on your own machine

1. Install the requirements

   ```
   $ pip install -r requirements.txt
   ```

2. Run the app

   ```
   $ streamlit run streamlit_app.py
   ```

## Excel: Immobilien-Kalkulation (Cockpit)

`excel/Immobilien-Kalkulation.xlsx` ist die professionalisierte Excel-Umsetzung des Entwurfs „Cockpit – Entwurf v4“.
Sie wird von `excel/build_immobilien_kalkulation.py` erzeugt (`pip install openpyxl`, dann `python excel/build_immobilien_kalkulation.py`).

- **Eingaben** (blaue Felder) und **Konfiguration** (Zielwerte/Ampel) sind die einzigen Eingabeblätter.
- **Cockpit**: Gesamtbewertung, KPI-Kacheln, Kennzahlen-Check mit Ampel, Diagramme, Prüfhinweise, Zeitverlauf.
- **Projektion / Finanzierung / Steuern / AfA-Vergleich / Sensitivität**: vollständige Formelrechnung über 30 Jahre.
- **Bankgespräch**: druckfertige Übersicht inkl. Haushaltsrechnung.

## Excel: Immobilien-Kalkulationstool Pro (Premium-Design)

`excel/Immobilien-Kalkulationstool_Pro_6.xlsx` ist das Kalkulationstool Pro im Premium-Design (Blau-Weiß):
Navigationsleiste mit allen Bereichen auf jedem Blatt, Schritt-Leiste mit Fortschritt und „Weiter“ auf den
zwölf Leitfaden-Seiten, Direktauswahl Privat / Kapitalgesellschaft auf der Startseite („Kauf als“),
Dashboard als Management-Übersicht, 3D-Säulen- und 3D-Kreisdiagramme.

Erzeugung aus der Vorlage `excel/quelle/Immobilien-Kalkulationstool_Pro_1.xlsx` (LibreOffice für die Neuberechnung):

```
pip install openpyxl lxml pillow
python excel/design_pro.py excel/quelle/Immobilien-Kalkulationstool_Pro_1.xlsx excel/Immobilien-Kalkulationstool_Pro_6.xlsx
soffice --headless --convert-to xlsx ...   # bzw. einmal in Excel öffnen und speichern (Neuberechnung)
python excel/navigation.py excel/Immobilien-Kalkulationstool_Pro_6.xlsx   # Reiter- und Schritt-Leiste als Formen
```

`design_pro.py` gestaltet die Darstellung, `dashboard.py` erzeugt das Dashboard, `navigation.py` setzt die
anklickbaren Leisten ein. Formeln und Namen bleiben unverändert; einzige Ausnahme: der Name `Rechtsform`
zeigt auf die Auswahl der Startseite, Schritt 9 übernimmt sie von dort.
