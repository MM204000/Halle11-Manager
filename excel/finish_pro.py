"""Nachlauf NACH der LibreOffice-Neuberechnung: Diagramm-Styling und Blatt-XML-Korrekturen.

Aufruf:  python excel/finish_pro.py MAPPE.xlsx   (ändert die Datei an Ort und Stelle)

LibreOffice schreibt beim Roundtrip Diagramme und Teile der Blatt-XML neu (Schriften, Gitterfarben,
Registerfarben, Datenüberprüfungs-Attribute). Deshalb wird all das hier – nach der Neuberechnung – gesetzt.
Zellwerte werden nicht berührt, die berechneten Ergebnisse bleiben gültig.
"""
import os
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from design_pro import INK, INK2, LINE2, MUTED, PIE_MAIN, SANS, SERIES, SERIES_MAP  # noqa: E402

# ============================================================================ Diagramme
NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

def a(tag):
    ns, t = tag.split(":")
    return f"{{{NS[ns]}}}{t}"


def set_run_props(rpr, size, color, bold=None):
    rpr.set("sz", str(size))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    for sf in rpr.findall(a("a:solidFill")):
        rpr.remove(sf)
    sf = etree.Element(a("a:solidFill"))
    etree.SubElement(sf, a("a:srgbClr")).set("val", color)
    rpr.insert(0, sf)
    lat = rpr.find(a("a:latin"))
    if lat is None:
        lat = etree.SubElement(rpr, a("a:latin"))
    lat.set("typeface", SANS)


def _walls(tag):
    el = etree.Element(a(tag))
    etree.SubElement(el, a("c:thickness")).set("val", "0")
    sp = etree.SubElement(el, a("c:spPr"))
    etree.SubElement(sp, a("a:noFill"))
    ln = etree.SubElement(sp, a("a:ln"))
    etree.SubElement(ln, a("a:noFill"))
    return el


def _view3d(rot_x, rot_y, right_angle, perspective=None):
    v = etree.Element(a("c:view3D"))
    etree.SubElement(v, a("c:rotX")).set("val", str(rot_x))
    etree.SubElement(v, a("c:rotY")).set("val", str(rot_y))
    etree.SubElement(v, a("c:depthPercent")).set("val", "100")
    etree.SubElement(v, a("c:rAngAx")).set("val", "1" if right_angle else "0")
    if perspective is not None and not right_angle:
        etree.SubElement(v, a("c:perspective")).set("val", str(perspective))
    return v


def make_3d(root):
    """Säulen- und Kreisdiagramme in schemakonforme 3D-Varianten überführen."""
    chart = root.find(a("c:chart"))
    plot = chart.find(a("c:plotArea"))
    kind = None
    for bc in plot.findall(a("c:barChart")):
        bc.tag = a("c:bar3DChart")
        for tag in ("c:overlap", "c:serLines"):
            for el in bc.findall(a(tag)):
                bc.remove(el)
        for pos in list(bc.iter(a("c:dLblPos"))):
            pos.getparent().remove(pos)
        gw = bc.find(a("c:gapWidth"))
        gd = etree.Element(a("c:gapDepth"))
        gd.set("val", "80")
        shape = etree.Element(a("c:shape"))
        shape.set("val", "box")
        if gw is not None:
            gw.set("val", str(min(int(gw.get("val", "150")), 90)))
            gw.addnext(gd)
        else:
            bc.find(a("c:axId")).addprevious(gd)
        gd.addnext(shape)
        kind = "bar"
    for pc in plot.findall(a("c:pieChart")):
        pc.tag = a("c:pie3DChart")
        for el in pc.findall(a("c:firstSliceAng")):
            pc.remove(el)
        kind = "pie"
    if kind is None:
        return
    view = _view3d(35, 0, False, 20) if kind == "pie" else _view3d(12, 18, True)
    plot.addprevious(view)
    if kind == "bar":
        for tag in ("c:floor", "c:sideWall", "c:backWall"):
            plot.addprevious(_walls(tag))


def style_chart(xml):
    root = etree.fromstring(xml)
    make_3d(root)
    # Serienfarben (Text wird separat gesetzt)
    for clr in root.iter(a("a:srgbClr")):
        v = clr.get("val", "").lower()
        if v in SERIES_MAP:
            clr.set("val", SERIES_MAP[v])
    for lat in root.iter(a("a:latin")):
        lat.set("typeface", SANS)
    # Titel
    for t in root.iter(a("c:title")):
        for rpr in list(t.iter(a("a:defRPr"))) + list(t.iter(a("a:rPr"))):
            set_run_props(rpr, 1050, INK, True)
    # Achsen, Legende, Datenbeschriftungen
    for tag in ("c:catAx", "c:valAx", "c:dateAx", "c:serAx"):
        for ax in root.iter(a(tag)):
            for rpr in ax.iter(a("a:defRPr")):
                set_run_props(rpr, 800, MUTED)
            sp = ax.find(a("c:spPr"))
            if sp is not None:
                for ln in sp.iter(a("a:ln")):
                    for clr in ln.iter(a("a:srgbClr")):
                        clr.set("val", LINE2)
            mg = ax.find(a("c:majorGridlines"))
            if mg is not None:
                for ln in mg.iter(a("a:ln")):
                    ln.set("w", "6350")
                    for clr in ln.iter(a("a:srgbClr")):
                        clr.set("val", "ECEEF0")
    for lg in root.iter(a("c:legend")):
        for rpr in lg.iter(a("a:defRPr")):
            set_run_props(rpr, 800, MUTED)
    for dl in root.iter(a("c:dLbls")):
        for rpr in dl.iter(a("a:defRPr")):
            set_run_props(rpr, 800, INK2)
    # Kreisdiagramme: hellere Hauptfarbe (3D-Schattierung bleibt lesbar)
    for tag in ("c:pieChart", "c:pie3DChart"):
        for pc in root.iter(a(tag)):
            for clr in pc.iter(a("a:srgbClr")):
                if clr.get("val", "").upper() == SERIES[0]:
                    clr.set("val", PIE_MAIN)
    # Kreisdiagramme: Beschriftung außen, gut lesbar
    for tag in ("c:pieChart", "c:pie3DChart", "c:doughnutChart"):
        for pc in root.iter(a(tag)):
            for pos in pc.iter(a("c:dLblPos")):
                if tag != "c:doughnutChart":
                    pos.set("val", "outEnd")
            for rpr in pc.iter(a("a:defRPr")):
                set_run_props(rpr, 800, INK2, True)
            for dl in pc.iter(a("c:dLbls")):
                if dl.find(a("c:dLblPos")) is None and tag != "c:doughnutChart":
                    pos = etree.Element(a("c:dLblPos"))
                    pos.set("val", "outEnd")
                    anchor = dl.find(a("c:txPr")) if dl.find(a("c:txPr")) is not None else dl.find(a("c:spPr"))
                    if anchor is not None:
                        anchor.addnext(pos)
                    else:
                        dl.insert(0, pos)
    # Liniendiagramme: kräftigere, runde Linien
    for lc in root.iter(a("c:lineChart")):
        for ser in lc.findall(a("c:ser")):
            sp = ser.find(a("c:spPr"))
            if sp is not None:
                for ln in sp.findall(a("a:ln")):
                    ln.set("w", str(max(int(ln.get("w", "0") or 0), 31750)))
                    ln.set("cap", "rnd")
    # Diagrammfläche ohne Rahmen
    cs_sp = root.find(a("c:spPr"))
    if cs_sp is None:  # Diagrammfläche ohne Rahmen, Standardschrift klein und grau
        cs_sp = etree.Element(a("c:spPr"))
        etree.SubElement(cs_sp, a("a:noFill"))
        etree.SubElement(cs_sp, a("a:ln"))
        root.find(a("c:chart")).addnext(cs_sp)
    if root.find(a("c:txPr")) is None:
        tx = etree.Element(a("c:txPr"))
        etree.SubElement(tx, a("a:bodyPr"))
        etree.SubElement(tx, a("a:lstStyle"))
        ppr = etree.SubElement(etree.SubElement(tx, a("a:p")), a("a:pPr"))
        set_run_props(etree.SubElement(ppr, a("a:defRPr")), 800, MUTED)
        etree.SubElement(tx.find(a("a:p")), a("a:endParaRPr")).set("lang", "de-DE")
        cs_sp.addnext(tx)
    if cs_sp is not None:
        for ln in cs_sp.findall(a("a:ln")):
            for child in list(ln):
                ln.remove(child)
            etree.SubElement(ln, a("a:noFill"))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)




def finish(path):
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        files = {i.filename: zin.read(i.filename) for i in infos}
    for name in list(files):
        if name.startswith("xl/charts/chart") and name.endswith(".xml"):
            files[name] = style_chart(files[name])
    try:
        import finish_sheets  # Blatt-XML-Korrekturen (Registerfarben, Datenüberprüfung, Ansicht)
        finish_sheets.fix(files)
    except ImportError:
        pass
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in infos:
            zout.writestr(i, files[i.filename])
        for name in files:
            if name not in {i.filename for i in infos}:
                zout.writestr(name, files[name])
    print(f"Nachlauf abgeschlossen: {path}")


if __name__ == "__main__":
    finish(sys.argv[1])
