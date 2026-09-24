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
