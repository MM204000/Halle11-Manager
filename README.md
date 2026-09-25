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

`excel/Immobilien-Kalkulationstool_Pro.xlsx` ist das Kalkulationstool Pro im Premium-Design „Midnight & Gold“ (Nachtblau, Gold, warme Flächen; Calibri):
Reiterleiste mit allen Bereichen auf jedem Blatt (plus Zell-Links als Rückfall), Schritt-Leiste mit Fortschritt und
„Weiter“ auf den zwölf Leitfaden-Seiten, Direktauswahl Privat / Kapitalgesellschaft auf der Startseite („Kauf als“),
Feld „Erstellt für“, Dashboard als Management-Übersicht, einheitliche Komponenten (Kacheln, Einordnungs-Boxen,
Abschnittsköpfe, Buttons, Summenstufen, Status-Pills), mehrfarbige Diagramme mit fester Farbbedeutung (3D nur für Kreise und einfache Säulen),
Druckeinrichtung für A4 auf allen Blättern.

Erzeugung aus der Vorlage `excel/quelle/Immobilien-Kalkulationstool_Pro_1.xlsx` (LibreOffice für die Neuberechnung):

```
pip install openpyxl lxml pillow
python excel/build_pro.py --strict --out excel/Immobilien-Kalkulationstool_Pro.xlsx [--preview <ordner>]
```

Ablauf: `design_pro.py` (Grund-Restyle, `global_rules.py`, `layouts/*.py`, `dashboard.py`; Bausteine in `core.py`)
→ Neuberechnung (LibreOffice) → `finish_pro.py` (Diagramme, Blatt-XML) → `navigation.py` (Leisten als Formen)
→ Schema-Prüfung → `lint_pro.py` (Design-Regeln) → `verify_logic.py` (Formeln, Werte und Namen gegen die Vorlage).
Formeln und Namen bleiben unverändert; einzige Ausnahme: der Name `Rechtsform` zeigt auf die Auswahl der
Startseite, Schritt 9 übernimmt sie von dort. Neu ist nur der Name `Erstellt_fuer`.
