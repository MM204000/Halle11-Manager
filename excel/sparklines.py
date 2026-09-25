"""Echte Excel-Sparklines (x14:sparklineGroups, Excel 2010+) für openpyxl-Mappen – Runde 6 „Wow-Paket“.

openpyxl kennt keine Sparklines und LibreOffice schreibt beim Neuberechnen nur, was es selbst versteht. Deshalb:

1. Blatt-Module (dashboard.py, layouts/cockpit.py …) melden Sparklines während des Builds an:

       import sparklines as SP
       SP.register(ws, "Q20", "Projektion!D30:AQ30")                      # Linie, Blau, Hoch/Tief Gold
       SP.register(ws, "Q21", "Projektion!D31:AQ31", kind="column")       # Säulen, negative Säulen Rot
       SP.register(ws, "Q20:Q23", "Projektion!D30:AQ33")                  # Block: je Zeile eine Sparkline, EINE Gruppe

2. design_pro.py schreibt die Registry nach dem Speichern als JSON neben die Stage-Datei (save_for).
3. finish_sheets.py fügt sie ganz am Ende der Pipeline (nach Neuberechnung, finish_pro, navigation.py, ScreenTips
   und Umfärbung) als <extLst><ext uri="{05C60535-…}"> ins Blatt-XML ein (inject_file). Vorhandene ext-Einträge
   (x14-Datenbalken/bedingte Formate, Datenüberprüfungen …) bleiben erhalten; die Reihenfolge folgt Excel.

Farben kommen aus core-Tokens: Linie/Säule BLUE (oder INK/NAVY), negativ RED, Hoch-/Tiefpunkt GOLD, erster/letzter
Punkt MUTED2, Achse LINE2. Die Registry ist unabhängig von core – nur die Standardfarben werden dort gelesen.
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import core as C  # noqa: E402
    _TOK = dict(series=C.BLUE, negative=C.RED, markers=C.BLUE, high=C.GOLD, low=C.GOLD, first=C.MUTED2,
                last=C.NAVY, axis=C.LINE2)
except Exception:  # pragma: no cover – Registry bleibt auch ohne core nutzbar
    _TOK = dict(series="1E4E8C", negative="B42318", markers="1E4E8C", high="C9A14A", low="C9A14A",
                first="8A9099", last="0E2238", axis="D6D2C6")

URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
# Reihenfolge der ext-Einträge in <worksheet><extLst>, wie Excel sie schreibt (unbekannte URIs ans Ende)
EXT_ORDER = [
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",   # x14:conditionalFormattings (Datenbalken, Symbolsätze)
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",   # x14:dataValidations
    URI,                                        # x14:sparklineGroups
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}",   # x14:slicerList
    "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}",   # x14:protectedRanges
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}",   # x14:ignoredErrors
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}",   # xlwe:webExtensions
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}",   # x15:slicerList
    "{7E03D99C-DC04-49d9-9315-930204A7B6E9}",   # x15:timelineRefs
]
KINDS = {"line": "line", "column": "column", "winloss": "stacked", "stacked": "stacked"}

REGISTRY = []


# =============================================================================== Anmelden
def _title(ws_or_title):
    return ws_or_title if isinstance(ws_or_title, str) else ws_or_title.title


def _hex(c):
    if c is None:
        return None
    c = str(c).lstrip("#").upper()
    if len(c) == 8:
        c = c[2:]
    if not re.fullmatch(r"[0-9A-F]{6}", c):
        raise ValueError(f"Farbe {c!r} ist kein 6-stelliger Hexwert")
    return c


_RANGE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$")


def _split_ref(ref, default_sheet):
    """„'S12 Ergebnis'!$D$5:$AQ$5“ → ('S12 Ergebnis', 'D', 5, 'AQ', 5)."""
    ref = str(ref).strip().lstrip("=")
    sheet = default_sheet
    if "!" in ref:
        sheet, ref = ref.rsplit("!", 1)
        sheet = sheet.strip()
        if sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
    m = _RANGE.match(ref.replace(" ", "").upper())
    if not m:
        raise ValueError(f"Bereich {ref!r} nicht lesbar")
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3) or m.group(1), int(m.group(4) or m.group(2))
    return sheet, c1, r1, c2, r2


def _col(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def _letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def qsheet(name):
    """Blattname für xm:f – in Hochkommas, sobald er kein schlichter Bezeichner ist."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return name
    return "'" + name.replace("'", "''") + "'"


# Vorlagen (style=…): Farbsemantik der Mappe (DECISIONS Runde 5) – einzelne Argumente überschreiben die Vorlage.
PRESETS = {
    "trend":    dict(kind="line", color=None, high=True, low=True, last=True),           # Blau, Hoch/Tief Gold
    "cashflow": dict(kind="column", color="C_POS", neg="C_NEG", axis=True, negative=True, high=False, low=False),
    "wealth":   dict(kind="line", color="NAVY", markers=True, high=True, low=False, last=True, weight=1.5),
    "value":    dict(kind="line", color="C_AQUA", markers=True, high=False, low=False, last=True),
    "debt":     dict(kind="line", color="C_NEUTRAL", markers=True, high=False, low=False, last=True, min_zero=True),
    "costs":    dict(kind="column", color="C_ORANGE", high=False, low=False),
    "rent":     dict(kind="column", color="C_BLUE", high=False, low=False),
    "winloss":  dict(kind="winloss", color="C_POS", neg="C_NEG", negative=True, high=False, low=False),
}


def _tok(v):
    """Token-Name („C_POS“, „NAVY“) oder Hex → Hex."""
    if isinstance(v, str) and not re.fullmatch(r"#?[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?", v):
        try:
            import core as _C
            return getattr(_C, v)
        except Exception as exc:
            raise ValueError(f"Farbe {v!r} unbekannt") from exc
    return v


_UNSET = object()


def register(ws, location, data_range, kind=_UNSET, color=_UNSET, neg=_UNSET, markers=_UNSET, high=_UNSET,
             low=_UNSET, first=_UNSET, last=_UNSET, negative=_UNSET, weight=_UNSET, axis=_UNSET, same_scale=False,
             min_zero=_UNSET, manual_min=None, manual_max=None, empty="gap", show_hidden=True, marker_color=None,
             high_color=None, low_color=None, first_color=None, last_color=None, axis_color=None, style=None):
    """Meldet eine Sparkline-Gruppe an.

    ws          Arbeitsblatt (oder Blattname), in dessen Zelle(n) die Sparklines stehen.
    location    Zielzelle „Q20“ oder Zielspalte/-zeile „Q20:Q23“ (dann je Zeile bzw. Spalte des Datenblocks eine Sparkline).
    data_range  Datenbereich „Projektion!D30:AQ30“ (ohne Blatt = eigenes Blatt); zu „Q20:Q23“ passt „…!D30:AQ33“.
    kind        'line' | 'column' | 'winloss'.
    style       Vorlage aus PRESETS: 'trend' · 'cashflow' · 'wealth' · 'value' · 'debt' · 'costs' · 'rent' · 'winloss'.
    color       Serienfarbe (Hex oder core-Tokenname wie „C_POS“, Standard core.BLUE); neg = Farbe negativer
                Säulen/Punkte (Standard core.RED).
    markers     Linie: alle Punkte markieren. high/low/first/last: Hoch-/Tief-/Erst-/Letztpunkt hervorheben
                (Standard Linie: Hoch/Tief Gold, letzter Punkt Navy; Säulen: nur negative hervorheben).
    weight      Linienstärke in pt. axis: horizontale Achse (Nulllinie) zeigen – sinnvoll bei Vorzeichenwechsel.
    same_scale  alle Sparklines der Gruppe mit gemeinsamer Min/Max-Achse. min_zero: Achse beginnt bei 0.
    manual_min/manual_max  feste Achsgrenzen. empty: leere Zellen als 'gap' | 'zero' | 'span'.
    show_hidden Daten aus ausgeblendeten Zeilen/Spalten zeigen (Standard an – Hilfsspalten sind oft ausgeblendet).
    Liefert den Registry-Eintrag (dict)."""
    title = _title(ws)
    if style is not None and style not in PRESETS:
        raise ValueError(f"style {style!r} – erlaubt: {', '.join(PRESETS)}")
    pre = dict(PRESETS.get(style, {}))
    given = dict(kind=kind, color=color, neg=neg, markers=markers, high=high, low=low, first=first, last=last,
                 negative=negative, weight=weight, axis=axis, min_zero=min_zero)
    kind_ = given["kind"] if given["kind"] is not _UNSET else pre.get("kind", "line")
    # Standard je Art: Linie → Hoch/Tief Gold + letzter Punkt; Säulen → nur negative Säulen hervorheben
    base = dict(kind=kind_, color=None, neg=None, markers=False, high=kind_ == "line", low=kind_ == "line",
                first=False, last=kind_ == "line", negative=None, weight=1.25, axis=False, min_zero=False)
    base.update(pre)
    base.update({k: v for k, v in given.items() if v is not _UNSET})
    kind, markers, high, low = base["kind"], base["markers"], base["high"], base["low"]
    first, last, negative, weight, axis, min_zero = (base["first"], base["last"], base["negative"], base["weight"],
                                                     base["axis"], base["min_zero"])
    color, neg = _tok(base["color"]), _tok(base["neg"])
    marker_color, high_color, low_color = _tok(marker_color), _tok(high_color), _tok(low_color)
    first_color, last_color, axis_color = _tok(first_color), _tok(last_color), _tok(axis_color)
    if kind not in KINDS:
        raise ValueError(f"kind {kind!r} – erlaubt: {', '.join(KINDS)}")
    ds, dc1, dr1, dc2, dr2 = _split_ref(data_range, title)
    _, lc1, lr1, lc2, lr2 = _split_ref(location, title)
    locs = [(c, r) for r in range(lr1, lr2 + 1) for c in range(_col(lc1), _col(lc2) + 1)]
    if len(locs) == 1:
        pairs = [(f"{_letter(locs[0][0])}{locs[0][1]}", (dc1, dr1, dc2, dr2))]
    elif lc1 == lc2:          # senkrechte Zielspalte: je Datenzeile eine Sparkline
        if dr2 - dr1 + 1 != len(locs):
            raise ValueError(f"{location}: {len(locs)} Zielzellen, aber {dr2 - dr1 + 1} Datenzeilen in {data_range}")
        pairs = [(f"{lc1}{r}", (dc1, dr1 + i, dc2, dr1 + i)) for i, (_, r) in enumerate(locs)]
    elif lr1 == lr2:          # waagerechte Zielzeile: je Datenspalte eine Sparkline
        if _col(dc2) - _col(dc1) + 1 != len(locs):
            raise ValueError(f"{location}: {len(locs)} Zielzellen, aber {_col(dc2) - _col(dc1) + 1} Datenspalten")
        pairs = [(f"{_letter(c)}{lr1}", (_letter(_col(dc1) + i), dr1, _letter(_col(dc1) + i), dr2))
                 for i, (c, _) in enumerate(locs)]
    else:
        raise ValueError(f"{location}: Ziel muss eine Zelle, Zeile oder Spalte sein")
    lines = [{"sqref": sq, "f": f"{qsheet(ds)}!${a}${b}:${c}${d}"} for sq, (a, b, c, d) in pairs]
    ser = _hex(color) or _TOK["series"]
    e = {
        "sheet": title, "kind": KINDS[kind], "lines": lines,
        "colors": {
            "series": ser, "negative": _hex(neg) or _TOK["negative"], "axis": _hex(axis_color) or _TOK["axis"],
            "markers": _hex(marker_color) or (ser if color else _TOK["markers"]),
            "first": _hex(first_color) or _TOK["first"], "last": _hex(last_color) or _TOK["last"],
            "high": _hex(high_color) or _TOK["high"], "low": _hex(low_color) or _TOK["low"],
        },
        "flags": {"markers": bool(markers) and kind == "line", "high": bool(high), "low": bool(low),
                  "first": bool(first), "last": bool(last),
                  "negative": bool(kind != "line" if negative is None else negative),
                  "displayXAxis": bool(axis), "displayHidden": bool(show_hidden)},
        "weight": float(weight), "empty": empty if empty in ("gap", "zero", "span") else "gap",
        "same_scale": bool(same_scale), "min_zero": bool(min_zero),
        "manual_min": manual_min, "manual_max": manual_max,
    }
    # dieselbe Zielzelle zweimal angemeldet → der spätere Eintrag gewinnt (Module dürfen nachbessern)
    targets = {(title, ln["sqref"]) for ln in lines}
    for old in list(REGISTRY):
        old["lines"] = [ln for ln in old["lines"] if (old["sheet"], ln["sqref"]) not in targets]
        if not old["lines"]:
            REGISTRY.remove(old)
    REGISTRY.append(e)
    return e


def clear():
    REGISTRY.clear()


def entries(sheet=None):
    return [e for e in REGISTRY if sheet is None or e["sheet"] == sheet]


# =============================================================================== Seitendatei
def sidecar(xlsx_path):
    """Pfad der Seitendatei neben der Mappe: stage.xlsx → stage.sparklines.json."""
    return os.path.splitext(xlsx_path)[0] + ".sparklines.json"


def save_for(xlsx_path, wb=None):
    """Schreibt die Registry als JSON neben die Mappe. Mit wb werden Einträge auf fehlende Blätter verworfen."""
    items = REGISTRY
    if wb is not None:
        names = set(wb.sheetnames)
        items = [e for e in REGISTRY if e["sheet"] in names]
    with open(sidecar(xlsx_path), "w", encoding="utf-8") as fh:
        json.dump({"version": 1, "groups": items}, fh, ensure_ascii=False, indent=1)
    return len(items)


def load(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as fh:
        return json.load(fh).get("groups", [])


# =============================================================================== Blatt-XML
def _b(v):
    return "1" if v else "0"


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def group_xml(e):
    """<x14:sparklineGroup> eines Registry-Eintrags (Elementreihenfolge nach CT_SparklineGroup)."""
    f, c = e["flags"], e["colors"]
    attrs = []
    if e.get("manual_max") is not None:
        attrs.append(f'manualMax="{float(e["manual_max"])!r}"')
    if e.get("manual_min") is not None:
        attrs.append(f'manualMin="{float(e["manual_min"])!r}"')
    attrs.append(f'lineWeight="{e.get("weight", 1.25)!r}"')
    if e["kind"] != "line":
        attrs.append(f'type="{e["kind"]}"')
    attrs.append(f'displayEmptyCellsAs="{e.get("empty", "gap")}"')
    for k in ("markers", "high", "low", "first", "last", "negative", "displayXAxis", "displayHidden"):
        if f.get(k):
            attrs.append(f'{k}="1"')
    if e.get("manual_min") is not None or e.get("min_zero"):
        attrs.append('minAxisType="custom"')
        if e.get("manual_min") is None:
            attrs[0:0] = ['manualMin="0"']
    elif e.get("same_scale"):
        attrs.append('minAxisType="group"')
    if e.get("manual_max") is not None:
        attrs.append('maxAxisType="custom"')
    elif e.get("same_scale"):
        attrs.append('maxAxisType="group"')
    col = "".join(f'<x14:color{k} rgb="FF{c[k2]}"/>'
                  for k, k2 in (("Series", "series"), ("Negative", "negative"), ("Axis", "axis"),
                                ("Markers", "markers"), ("First", "first"), ("Last", "last"),
                                ("High", "high"), ("Low", "low")))
    sps = "".join(f'<x14:sparkline><xm:f>{_esc(ln["f"])}</xm:f><xm:sqref>{ln["sqref"]}</xm:sqref></x14:sparkline>'
                  for ln in e["lines"])
    return f'<x14:sparklineGroup {" ".join(attrs)}>{col}<x14:sparklines>{sps}</x14:sparklines></x14:sparklineGroup>'


def ext_xml(groups):
    return (f'<ext xmlns:x14="{X14}" uri="{URI}"><x14:sparklineGroups xmlns:xm="{XM}">'
            + "".join(group_xml(e) for e in groups) + "</x14:sparklineGroups></ext>")


_EXTLST = re.compile(r"<extLst>(.*?)</extLst>", re.S)
_EXT = re.compile(r"<ext\b[^>]*?(?:/>|>.*?</ext>)", re.S)


def _uri(ext):
    m = re.search(r'\buri="([^"]+)"', ext[:ext.find(">") + 1])
    return m.group(1) if m else ""


def _rank(uri):
    up = [u.upper() for u in EXT_ORDER]
    return up.index(uri.upper()) if uri.upper() in up else len(up)


def inject_sheet(xml, groups):
    """Setzt die Sparkline-ext in das Blatt-XML (Text); ersetzt eine vorhandene Sparkline-ext, behält alle übrigen
    ext-Einträge und ordnet sie in Excels Reihenfolge. <extLst> ist immer das letzte Kind von <worksheet>."""
    new = ext_xml(groups) if groups else None
    # evtl. Präfix des Hauptnamensraums (LibreOffice schreibt ohne Präfix; sicherheitshalber prüfen)
    m = _EXTLST.search(xml)
    if m and xml.rfind("</worksheet>") > m.end():
        body = m.group(1)
        exts = [x for x in _EXT.findall(body) if _uri(x).upper() != URI.upper()]
        if new:
            exts.append(new)
        exts.sort(key=lambda x: _rank(_uri(x)))
        repl = f"<extLst>{''.join(exts)}</extLst>" if exts else ""
        return xml[:m.start()] + repl + xml[m.end():]
    if not new:
        return xml
    k = xml.rfind("</worksheet>")
    return xml[:k] + f"<extLst>{new}</extLst>" + xml[k:]


def inject(files, groups):
    """files: entpacktes Paket {pfad: bytes}. Liefert die Zahl der eingefügten Sparklines."""
    import finish_sheets as FS   # Blattpfade aus workbook.xml (lazy: finish_sheets importiert nicht zurück)
    paths = FS.sheet_paths(files)
    by_sheet = {}
    for e in groups:
        by_sheet.setdefault(e["sheet"], []).append(e)
    n = 0
    for sheet, gs in by_sheet.items():
        p = paths.get(sheet)
        if not p or p not in files:
            print(f"WARNUNG sparklines: Blatt {sheet!r} fehlt – {sum(len(g['lines']) for g in gs)} Sparklines übersprungen",
                  file=sys.stderr)
            continue
        files[p] = inject_sheet(files[p].decode("utf-8"), gs).encode("utf-8")
        n += sum(len(g["lines"]) for g in gs)
    return n


def inject_file(xlsx_path, json_path=None, groups=None):
    """Sparklines aus der Seitendatei (oder groups) in eine fertige .xlsx schreiben (an Ort und Stelle)."""
    import shutil
    import tempfile
    import zipfile
    if groups is None:
        groups = load(json_path or sidecar(xlsx_path))
    if not groups:
        return 0
    with zipfile.ZipFile(xlsx_path) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i.filename) for i in infos}
    n = inject(files, groups)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for i in infos:
            out.writestr(i, files[i.filename])
    shutil.move(tmp, xlsx_path)
    return n


def check(xlsx_path):
    """Prüft eine fertige Mappe: Anzahl Sparklines je Blatt, wohlgeformtes XML, ext-Reihenfolge, Zielzellen leer
    bzw. ohne Formel. Liefert (anzahl, [befunde])."""
    import zipfile
    from lxml import etree
    import finish_sheets as FS
    import glob
    issues, total = [], 0
    sml = None
    cands = glob.glob("/root/.claude/skills/synced/*/xlsx/scripts/office/schemas/ISO-IEC29500-4_2016/sml.xsd")
    if cands:
        try:
            sml = etree.XMLSchema(etree.parse(cands[0]))
        except Exception:
            sml = None
    order = ["colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh",
             "colorLow", "f", "sparklines"]
    with zipfile.ZipFile(xlsx_path) as z:
        files = {n: z.read(n) for n in z.namelist() if n.endswith(".xml") or n.endswith(".rels")}
    for sheet, p in FS.sheet_paths(files).items():
        xml = files.get(p, b"")
        if URI.encode() not in xml:
            continue
        root = etree.fromstring(xml)
        ext = root.find(f"{{{NS}}}extLst")
        if ext is None or root[-1] is not ext:
            issues.append(f"{sheet}: extLst nicht letztes Element")
        uris = [x.get("uri") for x in ext]
        if [_rank(u) for u in uris] != sorted(_rank(u) for u in uris):
            issues.append(f"{sheet}: ext-Reihenfolge {uris}")
        if sml is not None and not sml.validate(root):
            issues.append(f"{sheet}: Blatt-XML ungültig (sml.xsd): {str(sml.error_log.last_error)[:160]}")
        for g in root.iter(f"{{{X14}}}sparklineGroup"):   # Elementfolge nach CT_SparklineGroup (MS-XLSX 2.6.x)
            kids = [etree.QName(k).localname for k in g]
            if [order.index(k) for k in kids] != sorted(order.index(k) for k in kids) or kids[-1] != "sparklines":
                issues.append(f"{sheet}: Elementfolge der Sparkline-Gruppe {kids}")
            if g.get("type", "line") not in ("line", "column", "stacked"):
                issues.append(f"{sheet}: Sparkline-Typ {g.get('type')}")
        sps = root.findall(f".//{{{X14}}}sparkline")
        total += len(sps)
        for sp in sps:
            f, sq = sp.findtext(f"{{{XM}}}f"), sp.findtext(f"{{{XM}}}sqref")
            if not f or not sq:
                issues.append(f"{sheet}: Sparkline ohne f/sqref")
                continue
            cell = root.find(f".//{{{NS}}}c[@r='{sq}']")
            if cell is not None and cell.find(f"{{{NS}}}f") is not None:
                issues.append(f"{sheet}!{sq}: Zielzelle enthält eine Formel (Sparkline liegt hinter dem Wert)")
    return total, issues


if __name__ == "__main__":   # python sparklines.py check MAPPE.xlsx | inject MAPPE.xlsx [JSON]
    if len(sys.argv) >= 3 and sys.argv[1] == "check":
        n, iss = check(sys.argv[2])
        print(f"sparklines: {n} Sparklines", *(["Befunde:"] + iss if iss else ["ohne Befund"]), sep="\n  ")
        sys.exit(1 if iss else 0)
    if len(sys.argv) >= 3 and sys.argv[1] == "inject":
        print(f"sparklines: {inject_file(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)} eingefügt")
    else:
        print("Aufruf: python sparklines.py check|inject MAPPE.xlsx [JSON]")
