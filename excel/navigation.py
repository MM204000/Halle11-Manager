"""Navigationsleiste und Schritt-Leiste als anklickbare Formen (DrawingML) einsetzen.

Aufruf:  python excel/navigation.py MAPPE.xlsx

Wird NACH der Neuberechnung ausgeführt (LibreOffice würde Formen-Links sonst umschreiben).
Jedes Blatt erhält oben eine Reiterleiste mit allen Bereichen (aktiver Bereich hervorgehoben);
die zwölf Leitfaden-Seiten zusätzlich eine Schritt-Leiste mit Fortschritt und „Weiter“-Schaltfläche.
Die Formen werden nicht mitgedruckt. Rechenlogik und Zellinhalte bleiben unberührt.
"""
import posixpath
import re
import sys
import zipfile
from xml.sax.saxutils import escape

EMU = 9525  # je Pixel (96 dpi)

NAVY, NAVY_2, ACC, LIGHT = "0B2A4A", "1D4F8A", "4A86C8", "E7EEF7"
TAB_IDLE, TAB_TXT = "1A3F66", "C8D7EB"
GREY_BG, GREY_TXT = "EEF0F3", "8A9099"

NAV = [("Start", "Start"), ("Leitfaden", "Leitfaden"), ("Dashboard", "Dashboard"), ("Eingaben", "Eingaben"),
       ("Cockpit", "Cockpit"), ("Diagramme", "Diagramme"), ("Projektion", "Projektion"), ("Steuern", "Steuern"),
       ("Finanzierung", "Finanzierung"), ("Sensitivität", "Sensitivität"), ("Bank", "Bankgespräch"),
       ("Anhang", "Hinweise")]
ACTIVE = {"AfA-Vergleich": "Steuern", "Haushaltsrechnung": "Bank", "Vermögensaufstellung": "Bank",
          "Bankgespräch": "Bank", "Hinweise": "Anhang", "Konfiguration": "Anhang"}
STEPS = ["Objekt", "Kaufpreis", "Nebenkosten", "Aufteilung", "Maßnahmen", "Bewirtschaftung", "Finanzierung",
         "Zwischenstand", "Steuern", "Abschreibung", "Prognose", "Ergebnis"]

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_HYPER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def run(text, size, color, bold=False):
    return (f'<a:r><a:rPr lang="de-DE" sz="{int(size * 100)}" b="{1 if bold else 0}">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr>'
            f'<a:t>{escape(text)}</a:t></a:r>')


def shape(sid, name, x, y, w, h, fill, lines, rid=None, radius=22000, line=None, align="ctr"):
    """lines: Liste von (Text, Größe, Farbe, fett)."""
    link = f'<a:hlinkClick r:id="{rid}"/>' if rid else ""
    ln = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
    fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    paras = "".join(f'<a:p><a:pPr algn="{align}"/>{run(*ln_)}</a:p>' for ln_ in lines)
    return (f'<xdr:absoluteAnchor><xdr:pos x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
            f'<xdr:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/>'
            f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{sid}" name="{escape(name)}">{link}</xdr:cNvPr>'
            f'<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
            f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {radius}"/></a:avLst></a:prstGeom>'
            f'{fl}{ln}</xdr:spPr>'
            f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="45720" tIns="0" '
            f'rIns="45720" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>{paras}</xdr:txBody></xdr:sp>'
            f'<xdr:clientData fPrintsWithSheet="0"/></xdr:absoluteAnchor>')


def geometry(sheet_xml):
    m = re.search(r'<sheetFormatPr[^>]*defaultRowHeight="([\d.]+)"', sheet_xml)
    def_row = float(m.group(1)) if m else 15.0
    cols = {}
    for cm in re.finditer(r'<col [^>]*/>', sheet_xml):
        a = dict(re.findall(r'(\w+)="([^"]*)"', cm.group(0)))
        for i in range(int(a["min"]), min(int(a["max"]), 200) + 1):
            cols[i] = float(a.get("width", 8.43)) if a.get("customWidth") in ("1", "true") or "width" in a else 8.43
    rows = {}
    for rm in re.finditer(r'<row [^>]*>', sheet_xml):
        a = dict(re.findall(r'(\w+)="([^"]*)"', rm.group(0)))
        if "ht" in a:
            rows[int(a["r"])] = float(a["ht"])
    col_px = lambda c: int(cols.get(c, 8.43) * 7 + 5)
    row_px = lambda r: rows.get(r, def_row) * 96 / 72
    return col_px, row_px


def inject(path):
    zin = zipfile.ZipFile(path)
    files = {n: zin.read(n) for n in zin.namelist()}
    infos = {n: zin.getinfo(n) for n in zin.namelist()}
    zin.close()

    wb = files["xl/workbook.xml"].decode()
    wb_rels = files["xl/_rels/workbook.xml.rels"].decode()
    rel_target = dict(re.findall(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', wb_rels))
    rel_target.update({i: t for t, i in re.findall(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', wb_rels)})
    sheets = []
    for sm in re.finditer(r'<sheet [^>]*/>', wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', sm.group(0)))
        name = a["name"].replace("&amp;", "&")
        target = rel_target[a["r:id"]]
        sheets.append((name, "xl/" + target.lstrip("/").replace("xl/", "") if not target.startswith("/") else target.lstrip("/")))

    for name, spath in sheets:
        sxml = files[spath].decode()
        srels_path = posixpath.join(posixpath.dirname(spath), "_rels", posixpath.basename(spath) + ".rels")
        srels = files[srels_path].decode()
        dm = re.search(r'Target="([^"]*drawings/[^"]+)"', srels)
        if not dm:
            raise SystemExit(f"Blatt {name} hat keine Zeichnungsebene")
        dpath = posixpath.normpath(posixpath.join(posixpath.dirname(spath), dm.group(1)))
        drels_path = posixpath.join(posixpath.dirname(dpath), "_rels", posixpath.basename(dpath) + ".rels")
        dxml = files[dpath].decode()
        drels = files.get(drels_path, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{NS_REL}"></Relationships>'.encode()).decode()

        col_px, row_px = geometry(sxml)
        ids = [int(i) for i in re.findall(r'<xdr:cNvPr id="(\d+)"', dxml)] or [1]
        next_id = [max(ids) + 100]
        new_rels = []

        def rid_for(target):
            rid = f"rIdNav{len(new_rels) + 1}"
            new_rels.append(f'<Relationship Id="{rid}" Type="{REL_HYPER}" Target="#{escape(target, {chr(39): "&apos;"})}" TargetMode="External"/>')
            return rid

        def sid():
            next_id[0] += 1
            return next_id[0]

        shapes = []
        # ---------------------------------------------------------- Reiterleiste
        top2, h2 = row_px(1), row_px(2)
        th = 28
        ty = top2 + (h2 - th) / 2
        shapes.append(shape(sid(), "Marke", 46, ty - 1, 172, th + 2, None,
                            [("MM HOLDING", 9.5, "FFFFFF", True), ("Immobilien-Kalkulation", 7.5, TAB_TXT, False)],
                            align="l"))
        active = "Leitfaden" if (re.match(r"S\d\d ", name) or name == "Leitfaden") else ACTIVE.get(name, name)
        x = 226
        for label, target in NAV:
            on = label == active
            shapes.append(shape(sid(), f"Nav {label}", x, ty, 94, th, "FFFFFF" if on else TAB_IDLE,
                                [(label, 9, NAVY if on else TAB_TXT, on)], rid=rid_for(f"'{target}'!A1")))
            x += 98

        # ---------------------------------------------------------- Schritt-Leiste (Leitfaden-Seiten)
        m = re.match(r"S(\d\d) ", name)
        if m:
            step = int(m.group(1))
            y8 = sum(row_px(r) for r in range(1, 8)) + 3
            h8 = row_px(8) - 6
            x0 = col_px(1) + col_px(2)
            avail = sum(col_px(c) for c in range(3, 10))
            btn_w, gap = 190, 5
            cw = (avail - btn_w - gap * 12) / 12
            order = [n for n, _ in sheets if re.match(r"S\d\d ", n)]
            for i, label in enumerate(STEPS, start=1):
                done, cur = i < step, i == step
                bg = NAVY if cur else (LIGHT if done else GREY_BG)
                c1 = TAB_TXT if cur else (NAVY_2 if done else GREY_TXT)
                c2 = "FFFFFF" if cur else (NAVY if done else GREY_TXT)
                head = f"✓  SCHRITT {i}" if done else f"SCHRITT {i}"
                shapes.append(shape(sid(), f"Schritt {i}", x0 + (i - 1) * (cw + gap), y8, cw, h8, bg,
                                    [(head, 6.5, c1, True), (label, 8, c2, cur or done)],
                                    rid=rid_for(f"'{order[i - 1]}'!A1"), radius=16000))
            nxt_target = order[step] if step < 12 else "Dashboard"
            nxt_label = f"Schritt {step + 1}: {STEPS[step]}" if step < 12 else "Gesamtübersicht"
            shapes.append(shape(sid(), "Weiter", x0 + 12 * (cw + gap), y8, btn_w, h8, ACC,
                                [("WEITER  ›" if step < 12 else "ZUM DASHBOARD  ›", 9.5, "FFFFFF", True),
                                 (nxt_label, 7.5, "FFFFFF", False)],
                                rid=rid_for(f"'{nxt_target}'!A1"), radius=16000))

        if 'xmlns:r=' not in dxml[:600]:
            dxml = dxml.replace("<xdr:wsDr ", '<xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        dxml = dxml.replace("</xdr:wsDr>", "".join(shapes) + "</xdr:wsDr>")
        drels = drels.replace("</Relationships>", "".join(new_rels) + "</Relationships>")
        files[dpath] = dxml.encode()
        files[drels_path] = drels.encode()

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in files.items():
            info = infos.get(n) or zipfile.ZipInfo(n)
            zout.writestr(info, data)
    print(f"Navigation eingesetzt: {path} ({len(sheets)} Blätter)")


if __name__ == "__main__":
    inject(sys.argv[1])
