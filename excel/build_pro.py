"""Kompletter Build des Immobilien-Kalkulationstools.

Aufruf:  python excel/build_pro.py [--src VORLAGE] [--out ZIEL] [--strict] [--preview ORDNER]

Ablauf: design_pro (Gestaltung, openpyxl) → Neuberechnung (LibreOffice) → finish_pro (Diagramme, Blatt-XML)
→ navigation (Reiter-/Schritt-Leiste als Formen) → Schema-Prüfung → lint_pro → verify_logic → optional Vorschau.
Die Neuberechnung nutzt recalc.py der xlsx-Skill (Umgebungsvariable RECALC_PY oder automatische Suche).
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


def find_recalc():
    if os.environ.get("RECALC_PY"):
        return os.environ["RECALC_PY"]
    hits = glob.glob("/root/.claude/skills/synced/*/xlsx/scripts/recalc.py") + glob.glob(os.path.expanduser("~/.claude/skills/**/recalc.py"), recursive=True)
    return hits[0] if hits else None


def run(cmd, env=None, check=True):
    print("›", " ".join(os.path.basename(c) if i < 2 else c for i, c in enumerate(cmd)))
    r = subprocess.run(cmd, env=env, capture_output=True, text=True)
    out = (r.stdout or "") + (r.stderr or "")
    print(out.strip()[-6000:])
    if check and r.returncode != 0:
        raise SystemExit(f"Schritt fehlgeschlagen ({r.returncode}): {' '.join(cmd)}")
    return r.returncode, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=os.path.join(HERE, "quelle", "Immobilien-Kalkulationstool_Pro_1.xlsx"))
    ap.add_argument("--out", default=os.path.join(HERE, "Immobilien-Kalkulationstool_Pro.xlsx"))
    ap.add_argument("--strict", action="store_true", help="Modulfehler, Lint- und Logikfehler brechen ab")
    ap.add_argument("--preview", help="Ordner für bildschirmnahe Vorschau-PNGs")
    a = ap.parse_args()

    env = dict(os.environ)
    if a.strict:
        env["DESIGN_STRICT"] = "1"
    work = tempfile.mkdtemp(prefix="build_pro_")
    stage = os.path.join(work, "stage.xlsx")
    run([PY, os.path.join(HERE, "design_pro.py"), a.src, stage], env=env)
    recalc = find_recalc()
    if not recalc:
        raise SystemExit("recalc.py nicht gefunden (RECALC_PY setzen)")
    code, out = run([PY, recalc, stage, "240"])
    if '"status": "success"' not in out.replace("\n", " ").replace("  ", " ") and '"status":"success"' not in out.replace(" ", ""):
        print("WARNUNG: Neuberechnung meldet Fehler")
        if a.strict:
            raise SystemExit(1)
    run([PY, os.path.join(HERE, "finish_pro.py"), stage])
    run([PY, os.path.join(HERE, "navigation.py"), stage])
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    shutil.copy(stage, a.out)
    rc_schema, _ = run([PY, os.path.join(HERE, "tools", "validate_xml.py"), a.out], check=False)
    rc_lint, _ = run([PY, os.path.join(HERE, "lint_pro.py"), a.out], check=False)
    rc_logic, _ = run([PY, os.path.join(HERE, "verify_logic.py"), a.src, a.out], check=False)
    if a.preview:
        run([PY, os.path.join(HERE, "tools", "preview.py"), a.out, a.preview], check=False)
    shutil.rmtree(work, ignore_errors=True)
    print(f"\nBuild fertig: {a.out}  | Schema {'OK' if rc_schema == 0 else 'FEHLER'} · Lint {'OK' if rc_lint == 0 else 'Befunde'}"
          f" · Logik {'OK' if rc_logic == 0 else 'FEHLER'}")
    if a.strict and (rc_schema or rc_logic or rc_lint):
        sys.exit(1)


if __name__ == "__main__":
    main()
