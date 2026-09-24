"""Schema-Prüfung aller Zeichnungsebenen und Diagramme gegen die OOXML-Schemata (ISO/IEC 29500)."""
import glob
import sys
import zipfile

from lxml import etree

cands = glob.glob("/root/.claude/skills/synced/*/xlsx/scripts/office/schemas/ISO-IEC29500-4_2016/")
if not cands:
    print("Schemata nicht gefunden – Prüfung übersprungen")
    sys.exit(0)
S = cands[0]
sd = etree.XMLSchema(etree.parse(S + "dml-spreadsheetDrawing.xsd"))
sc = etree.XMLSchema(etree.parse(S + "dml-chart.xsd"))
z = zipfile.ZipFile(sys.argv[1])
bad = n = 0
for name in sorted(z.namelist()):
    sch = sd if (name.startswith("xl/drawings/drawing") and name.endswith(".xml")) else \
        sc if (name.startswith("xl/charts/chart") and name.endswith(".xml")) else None
    if sch is None:
        continue
    n += 1
    if not sch.validate(etree.fromstring(z.read(name))):
        bad += 1
        print(name, [str(e.message)[:200] for e in sch.error_log][:2])
print(f"Schema: {n} Zeichnungen/Diagramme geprüft, {bad} ungültig")
sys.exit(1 if bad else 0)
