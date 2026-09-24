"""Logik-Wächter: prüft, dass die Umgestaltung die Rechenlogik der Vorlage nicht verändert.

Aufruf:  python excel/verify_logic.py VORLAGE.xlsx ERGEBNIS.xlsx   (ERGEBNIS muss neu berechnet sein)

FEHLER (Exit-Code 1):
  - eine Formel der Vorlage ist geändert/entfernt UND wird von einer Formel, einem Namen oder Diagramm referenziert
  - eine unveränderte Formel liefert ein anderes Ergebnis (außer HEUTE()-abhängige Zellen)
  - ein numerischer Eingabewert der Vorlage wurde verändert
  - ein Name fehlt oder zeigt woanders hin (erlaubt: Rechtsform → Start!$D$23)
HINWEIS: geänderte, aber nirgends referenzierte Anzeigeformeln (reine Darstellung).
"""
import re
import sys
import zipfile
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import column_index_from_string, range_boundaries

ALLOWED_NAME_CHANGES = {"Rechtsform": "Start!$D$23"}
REF_RE = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_\.äöüÄÖÜß\- ]+))!)?(\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?|\$?\d+:\$?\d+)$")


def parse_ref(text, current):
    m = REF_RE.match(text.strip())
    if not m:
        return None
    sheet = (m.group(1) or m.group(2) or current).replace("''", "'")
    r = m.group(3).replace("$", "")
    try:
        if re.match(r"^\d+:\d+$", r):
            a, b = r.split(":")
            return sheet, (1, int(a), 16384, int(b))
        if re.match(r"^[A-Z]+:[A-Z]+$", r):
            a, b = r.split(":")
            return sheet, (column_index_from_string(a), 1, column_index_from_string(b), 1048576)
        c1, r1, c2, r2 = range_boundaries(r)
        return sheet, (c1, r1, c2 or c1, r2 or r1)
    except Exception:
        return None


def collect_refs(wb, zpath):
    """Alle referenzierten Bereiche: {blatt: [(c1,r1,c2,r2,quelle), ...]}."""
    refs = defaultdict(list)
    for ws in wb.worksheets:
        for c in ws._cells.values():
            v = c.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            try:
                toks = Tokenizer(v).items
            except Exception:
                continue
            for t in toks:
                if t.type == Token.OPERAND and t.subtype == Token.RANGE:
                    p = parse_ref(t.value, ws.title)
                    if p:
                        refs[p[0]].append((*p[1], f"{ws.title}!{c.coordinate}"))
    for n, d in wb.defined_names.items():
        for part in re.split(r",(?![^(]*\))", d.attr_text or ""):
            p = parse_ref(part, None)
            if p and p[0]:
                refs[p[0]].append((*p[1], f"Name {n}"))
    with zipfile.ZipFile(zpath) as z:
        for name in z.namelist():
            if name.startswith("xl/charts/chart") and name.endswith(".xml"):
                for f in re.findall(r"<c:f>([^<]+)</c:f>", z.read(name).decode("utf-8", "ignore")):
                    f = f.replace("&apos;", "'").replace("&amp;", "&")
                    for part in f.strip("()").split(","):
                        p = parse_ref(part, None)
                        if p and p[0]:
                            refs[p[0]].append((*p[1], f"Diagramm {name.split('/')[-1]}"))
    return refs


def referenced_by(refs, sheet, row, col, self_ref):
    out = []
    for c1, r1, c2, r2, src in refs.get(sheet, []):
        if src == self_ref:
            continue
        if c1 <= col <= c2 and r1 <= row <= r2:
            out.append(src)
    return out


def main(src, dst):
    of = load_workbook(src)
    ov = load_workbook(src, data_only=True)
    nf = load_workbook(dst)
    nv = load_workbook(dst, data_only=True)
    refs = collect_refs(nf, dst)
    errors, notes = [], []

    for ws in of.worksheets:
        if ws.title not in nf.sheetnames:
            errors.append(f"Blatt fehlt: {ws.title}")
            continue
        wf, wv, owv = nf[ws.title], nv[ws.title], ov[ws.title]
        for c in ws._cells.values():
            v = c.value
            coord = c.coordinate
            new = wf[coord].value
            if isinstance(v, str) and v.startswith("=") and c.data_type == "f":
                if new != v:
                    users = referenced_by(refs, ws.title, c.row, c.column, f"{ws.title}!{coord}")
                    if users:
                        errors.append(f"Formel geändert und referenziert: {ws.title}!{coord} ({v[:50]} → {str(new)[:50]}) "
                                      f"genutzt von {users[:3]}")
                    else:
                        notes.append(f"Anzeigeformel geändert: {ws.title}!{coord}: {v[:40]} → {str(new)[:40]}")
                    continue
                if "TODAY(" in v.upper():
                    continue
                x, y = owv[coord].value, wv[coord].value
                if isinstance(x, (int, float)) and not isinstance(x, bool):
                    if not isinstance(y, (int, float)) or abs(x - y) > 1e-6 * max(1, abs(x)):
                        errors.append(f"Ergebnis abweichend: {ws.title}!{coord}: {x} → {y}")
                elif (x not in (None, "") or y not in (None, "")) and x != y:
                    errors.append(f"Ergebnis abweichend: {ws.title}!{coord}: {str(x)[:40]} → {str(y)[:40]}")
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                if not isinstance(new, (int, float)) or abs(new - v) > 1e-9 * max(1, abs(v)):
                    users = referenced_by(refs, ws.title, c.row, c.column, None)
                    (errors if users else notes).append(f"Eingabewert geändert: {ws.title}!{coord}: {v} → {new}")

    for n, d in of.defined_names.items():
        if n not in nf.defined_names:
            errors.append(f"Name fehlt: {n}")
            continue
        new = nf.defined_names[n].attr_text
        if new != d.attr_text and ALLOWED_NAME_CHANGES.get(n) != new:
            errors.append(f"Name geändert: {n}: {d.attr_text} → {new}")

    print(f"verify_logic: {len(errors)} Fehler, {len(notes)} Hinweise")
    errors.sort(key=lambda e: (e.startswith("Ergebnis abweichend"), e))
    if sum(e.startswith("Ergebnis abweichend") for e in errors) > 1000:
        print("  (sehr viele Ergebnisabweichungen – ist die Datei neu berechnet?)")
    for e in errors[:60]:
        print("  FEHLER", e)
    for n_ in notes[:60]:
        print("  hinweis", n_)
    if len(notes) > 60:
        print(f"  … {len(notes) - 60} weitere Hinweise")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
