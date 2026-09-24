"""Blatt-XML-Korrekturen nach der LibreOffice-Neuberechnung (Registerfarben, Datenüberprüfung, Ansicht).

fix(files) erhält das entpackte Paket als dict {pfad: bytes} und ändert es an Ort und Stelle.

- Registerfarben je Gruppe (LibreOffice verwirft sie teilweise)                       P3-02
- Datenüberprüfung: Fehlermeldung an (stop), klare Fehler- und Eingabetexte             P1-07
- Präsentationsblätter ohne Zeilen-/Spaltenköpfe, Objekte im Blattschutz geschützt      P3-02
- dxf-Einträge ohne Schriftnamen (Altschriften Fraunces/Inter aus der Vorlage)          P3-06
- Datenbalken als Vollton ohne Verlauf (x14-Erweiterung, Excel 2010+)                   Wunsch G
"""
import os
import posixpath
import re
import sys

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import global_rules as G  # noqa: E402

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"


def q(tag):
    return f"{{{NS}}}{tag}"


# Datenüberprüfung: Texte je Typ, Sonderfälle je Blatt/Zelle
ERR_TITLE = "Ungültige Eingabe"
ERR_LIST = "Bitte einen Wert aus der Liste ▾ wählen."
ERR_PCT = "Bitte einen Prozentwert zwischen 0 % und 100 % eingeben."
SPECIAL = {
    ("Start", "D23"): dict(errorTitle="Bitte aus der Liste wählen",
                           error="Privatperson, vermögensverwaltende GmbH oder gewerbliche GmbH/Holding über ▾ auswählen.",
                           promptTitle="Rechtsform"),
    ("S01 Objekt", "D17"): dict(error="Bitte ein Baujahr zwischen 1800 und 2100 eingeben."),
    ("S05 Maßnahmen & Reserve", "D17"): dict(error="Bitte eine ganze Zahl von 1 bis 5 Jahren eingeben."),
}
PROMPTS = {
    ("S01 Objekt", "D14"): ("Objektart", "Bestimmt AfA-Satz und Sonder-AfA-Prüfung. Auswahl über ▾."),
    ("S01 Objekt", "D19"): ("Bundesland", "Bestimmt den Grunderwerbsteuersatz. Auswahl über ▾."),
    ("S04 Kaufpreisaufteilung", "D12"): ("Aufteilungsmethode", "Wie der Kaufpreis auf Boden und Gebäude verteilt wird. Auswahl über ▾."),
    ("S09 Steuern", "D13"): ("Veranlagung", "Einzel- oder Zusammenveranlagung – nur für Privatpersonen. Auswahl über ▾."),
    ("S09 Steuern", "D16"): ("Solidaritätszuschlag", "Ja oder Nein – nur für Privatpersonen. Auswahl über ▾."),
    ("S09 Steuern", "D17"): ("Kirchensteuer", "Satz nach Bundesland – nur für Privatpersonen. Auswahl über ▾."),
    ("S10 Abschreibung", "D12"): ("AfA-Methode", "„Automatisch nach Baujahr“ wählt den gesetzlichen Satz. Auswahl über ▾."),
}


def sheet_paths(files):
    """{Blattname: 'xl/worksheets/sheetN.xml'} aus workbook.xml und seinen Beziehungen."""
    wb = etree.fromstring(files["xl/workbook.xml"])
    rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in rels.iter(f"{{{PNS}}}Relationship")}
    out = {}
    for s in wb.iter(q("sheet")):
        t = target.get(s.get(f"{{{RNS}}}id"))
        if not t:
            continue
        p = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
        out[s.get("name")] = p
    return out


def _first_cell(sqref):
    return sqref.split()[0].split(":")[0]


def fix_validations(name, root):
    dvs = root.find(q("dataValidations"))
    if dvs is None:
        return
    for dv in dvs.findall(q("dataValidation")):
        cell = _first_cell(dv.get("sqref", ""))
        kind = dv.get("type")
        dv.set("showErrorMessage", "1")
        dv.set("errorStyle", "stop")
        if not dv.get("errorTitle"):
            dv.set("errorTitle", ERR_TITLE)
        if kind == "list":
            dv.set("error", ERR_LIST)
        elif kind == "decimal" and (dv.findtext(q("formula1")) or "").strip() == "0" \
                and (dv.findtext(q("formula2")) or "").strip() == "1":
            dv.set("error", ERR_PCT)
        elif not dv.get("error"):
            dv.set("error", "Bitte einen gültigen Wert eingeben.")
        if dv.get("promptTitle") == "Kaufstruktur":
            dv.set("promptTitle", "Rechtsform")
        for k, v in SPECIAL.get((name, cell), {}).items():
            dv.set(k, v)
        if (name, cell) in PROMPTS:
            title, text = PROMPTS[(name, cell)]
            dv.set("showInputMessage", "1")
            dv.set("promptTitle", title)
            dv.set("prompt", text)
        # Reihenfolge der Attribute ist egal; Excel begrenzt Titel auf 32, Texte auf 255 Zeichen
        for k, lim in (("errorTitle", 32), ("promptTitle", 32), ("error", 255), ("prompt", 255)):
            if dv.get(k) and len(dv.get(k)) > lim:
                dv.set(k, dv.get(k)[:lim])


def fix_tab_color(name, root):
    pr = root.find(q("sheetPr"))
    if pr is None:
        pr = etree.Element(q("sheetPr"))
        root.insert(0, pr)
    tc = pr.find(q("tabColor"))
    if tc is None:
        tc = etree.Element(q("tabColor"))
        pr.insert(0, tc)
    for a in list(tc.attrib):
        del tc.attrib[a]
    tc.set("rgb", "FF" + G.tab_color(name))


def fix_view(name, root):
    if name in G.PRESENTATION or G.STEP_RE.match(name):
        for sv in root.iter(q("sheetView")):
            sv.set("showRowColHeaders", "0")
    sp = root.find(q("sheetProtection"))
    if sp is not None and sp.get("sheet") in ("1", "true"):
        sp.set("objects", "1")


X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"


def fix_databars(root):
    """Datenbalken: Vollton statt ausbleichendem Verlauf (gradient=0), kein Rahmen – in Excel 2010+ wirksam,
    ältere Versionen zeigen weiter den klassischen Balken."""
    for db in root.iter(f"{{{X14}}}dataBar"):
        db.set("gradient", "0")
        db.set("border", "0")


def fix_styles(files):
    """dxf ohne Schriftnamen: bedingte Formate erben die Zellschrift (Calibri)."""
    p = "xl/styles.xml"
    if p not in files:
        return
    root = etree.fromstring(files[p])
    dxfs = root.find(q("dxfs"))
    if dxfs is None:
        return
    changed = False
    for fnt in dxfs.iter(q("font")):
        for tag in ("name", "family", "scheme"):
            for el in fnt.findall(q(tag)):
                fnt.remove(el)
                changed = True
    if changed:
        files[p] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fix_defined_names(files):
    """Druckbereiche/-titel: je Blatt nur ein Eintrag (LibreOffice schreibt Drucktitel doppelt), Attribut
    „name“ zuerst (Reihenfolge ist für Excel belanglos; Werkzeuge erkennen die Einträge so zuverlässig)."""
    p = "xl/workbook.xml"
    root = etree.fromstring(files[p])
    dns = root.find(q("definedNames"))
    if dns is None:
        return
    seen = set()
    for dn in list(dns.findall(q("definedName"))):
        key = (dn.get("name"), dn.get("localSheetId"))
        if key[0] in ("_xlnm.Print_Titles", "_xlnm.Print_Area") and key in seen:
            dns.remove(dn)
            continue
        seen.add(key)
        attrs = list(dn.attrib.items())
        for k, _ in attrs:
            del dn.attrib[k]
        dn.set("name", key[0])
        for k, v in attrs:
            if k != "name":
                dn.set(k, v)
    files[p] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fix(files):
    try:
        fix_defined_names(files)
    except Exception as exc:
        print(f"WARNUNG finish_sheets Namen: {exc!r}", file=sys.stderr)
    paths = sheet_paths(files)
    for name, p in paths.items():
        if p not in files:
            continue
        try:
            root = etree.fromstring(files[p])
            fix_tab_color(name, root)
            fix_validations(name, root)
            fix_view(name, root)
            fix_databars(root)
            files[p] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        except Exception as exc:  # ein Blatt darf den Nachlauf nicht abbrechen
            print(f"WARNUNG finish_sheets {name}: {exc!r}", file=sys.stderr)
    try:
        fix_styles(files)
    except Exception as exc:
        print(f"WARNUNG finish_sheets styles: {exc!r}", file=sys.stderr)
