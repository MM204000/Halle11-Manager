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

`excel/Immobilien-Kalkulationstool_Pro_4_blau.xlsx` (empfohlen) und `excel/Immobilien-Kalkulationstool_Pro_4_rot.xlsx`
sind das Kalkulationstool Pro im Premium-Design mit 3D-Säulen- und 3D-Kreisdiagrammen.
Beide werden aus der Vorlage `excel/quelle/Immobilien-Kalkulationstool_Pro_1.xlsx` erzeugt:

```
pip install openpyxl lxml pillow
python excel/design_pro.py excel/quelle/Immobilien-Kalkulationstool_Pro_1.xlsx excel/Immobilien-Kalkulationstool_Pro_4_blau.xlsx blau
python excel/design_pro.py excel/quelle/Immobilien-Kalkulationstool_Pro_1.xlsx excel/Immobilien-Kalkulationstool_Pro_4_rot.xlsx rot
```

Das Skript verändert ausschließlich die Darstellung (Farbthema, Kopfleiste, Typografie, Tabellen, Kacheln,
Schaltflächen, Hinweise, Zahlenformate, Diagramme inkl. 3D, Logo). Formeln, Namen, Blattschutz und
Datenüberprüfungen bleiben unverändert. Nach dem Erzeugen die Mappe einmal in Excel öffnen oder mit
LibreOffice neu berechnen.
