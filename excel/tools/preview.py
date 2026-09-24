"""Bildschirmnahe Vorschau: 100 % Maßstab, A2 quer, Formen sichtbar, je Blatt bis zu 3 Seiten als PNG.

Aufruf: python excel/tools/preview.py MAPPE.xlsx ORDNER   (DPI=72 Standard; Umgebungsvariable DPI)
"""
import sys, os, re, zipfile, subprocess, html, tempfile, shutil
import pymupdf as fitz
src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
for f in os.listdir(out):
    if f.endswith('.png'): os.remove(os.path.join(out, f))
tmp = os.path.join(out, 'screen.xlsx')
zi = zipfile.ZipFile(src)
wbx = zi.read('xl/workbook.xml').decode()
names = [html.unescape(n) for n in re.findall(r'<sheet [^>]*name="([^"]+)"', wbx)]
with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        d = zi.read(it.filename)
        n = it.filename
        if n.startswith('xl/drawings/drawing'):
            d = d.replace(b'fPrintsWithSheet="0"', b'fPrintsWithSheet="1"')
        elif n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
            s = d.decode()
            s = re.sub(r'<pageSetUpPr[^>]*/>', '<pageSetUpPr fitToPage="0"/>', s)
            s = re.sub(r'<pageSetup [^>]*/>', '<pageSetup paperSize="66" scale="100" orientation="landscape"/>', s)
            s = re.sub(r'<pageMargins [^>]*/>', '<pageMargins left="0.1" right="0.1" top="0.1" bottom="0.3" header="0" footer="0.1"/>', s)
            s = re.sub(r'<rowBreaks.*?</rowBreaks>|<colBreaks.*?</colBreaks>', '', s, flags=re.S)
            d = s.encode()
        elif n == 'xl/workbook.xml':
            s = re.sub(r'<definedName name="_xlnm\.(Print_Area|Print_Titles)"[^>]*>.*?</definedName>', '', d.decode())
            d = s.encode()
        zo.writestr(it, d)
prof = tempfile.mkdtemp(prefix='lo_prof_')
subprocess.run(['timeout', '400', 'soffice', '--headless', '--norestore', f'-env:UserInstallation=file://{prof}',
                '--convert-to', 'pdf', '--outdir', out, tmp], capture_output=True)
shutil.rmtree(prof, ignore_errors=True)
doc = fitz.open(os.path.join(out, 'screen.pdf'))
count = {}
cur = None
for i, p in enumerate(doc):
    t = p.get_text()
    m = re.search(r'Immobilien-Kalkulation · ([^\n]+)\n', t)
    name = m.group(1).strip() if m else cur
    cur = name
    if name not in names: continue
    k = count.get(name, 0)
    if k >= 3: continue
    count[name] = k + 1
    idx = names.index(name)
    fn = f"{idx:02d}_{re.sub(r'[^A-Za-z0-9]', '_', name)}_{k + 1}.png"
    p.get_pixmap(dpi=int(os.environ.get('DPI', '72'))).save(os.path.join(out, fn))
print(len(doc), 'Seiten;', len(count), 'Blätter;', sum(count.values()), 'Bilder')
missing = [n for n in names if n not in count]
print('fehlend:', missing)
