"""Icon-System und Illustrationen („Midnight & Gold“, Runde 6).

Ein einheitlicher Satz feiner Linien-Icons (24er-Raster, Strich 1,6 Einheiten, runde Enden und Ecken), per Pillow
mit 4-fachem Supersampling gezeichnet und weich verkleinert – dazu die Cover-Illustration (goldene Linienzeichnung
einer Wohnhaus-/Stadtsilhouette auf Nachtblau) und ein MM-Monogramm.

API (Details: scratchpad/ICONS_API.md)
    icon_path(name, color_hex, px=20, bg=None, pad=None)      → PNG-Pfad (Cache), 2× Auflösung
    place_icon(ws, anchor_cell, name, color_hex, px=20, dx=0, dy=0, bg=None, align=None, valign=None) → openpyxl-Bild
    place_image(ws, anchor_cell, path, w_px, h_px, dx=0, dy=0)                                          → openpyxl-Bild
    cover_art_path(width_px, height_px, background=True, focus="right", scale=2)                       → JPG/PNG-Pfad
    monogram_path(px=120, letters="MM", fg="FFFFFF", ring=GOLD)                                        → PNG-Pfad
    NAMES – alle Icon-Namen;  gallery_path() – Übersicht aller Icons (Test)

Hinweis Farbe: design_pro.postprocess färbt PNGs in xl/media auf die Logo-Farben um (nahe Navy → NAVY, nahe Gold → GOLD,
nahe Creme → Weiß). Icons sind deshalb einfarbig mit Alpha-Kantenglättung gezeichnet (NAVY, GOLD, WHITE, BLUE, SKY … sind
sicher); die Cover-Illustration mit Verlauf wird als JPEG geliefert (wird nicht umgefärbt, keine Streifenbildung).
"""
import hashlib
import math
import os
import random
import tempfile

from PIL import Image, ImageChops, ImageDraw, ImageFont
from PIL.PngImagePlugin import PngInfo

try:                                                     # Tokens aus core (Fallback, falls core nicht importierbar)
    import core as _C
    NAVY, NAVY_2, GOLD, WHITE, BLUE = _C.NAVY, _C.NAVY_2, _C.GOLD, _C.WHITE, _C.BLUE
except Exception:                                        # pragma: no cover
    NAVY, NAVY_2, GOLD, WHITE, BLUE = "0E2238", "1B3553", "C9A14A", "FFFFFF", "1E4E8C"

VERSION = "6.3"
SS = 4                     # Supersampling
STROKE = 1.6               # Strichstärke im 24er-Raster (bei 20 px ≈ 1,33 px – fein wie ein Premium-Set)
EMU = 9525
try:                                                     # Cache-Ordner je Quelltextstand → Änderungen wirken sofort
    with open(os.path.abspath(__file__), "rb") as _f:
        _SRC_HASH = hashlib.md5(_f.read()).hexdigest()[:8]
except OSError:                                          # pragma: no cover
    _SRC_HASH = "0"
_CACHE = os.path.join(tempfile.gettempdir(), f"mm_icons_{VERSION}_{_SRC_HASH}")
_RECOLOR_TARGETS = ((15, 28, 22), (160, 138, 85), (198, 164, 92), (244, 240, 232))


def _rgb(hex_):
    h = str(hex_).lstrip("#")[-6:]
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _cache_file(key, ext):
    os.makedirs(_CACHE, exist_ok=True)
    return os.path.join(_CACHE, hashlib.md5(key.encode()).hexdigest()[:16] + "." + ext)


def _meta():
    info = PngInfo()
    info.add_text("Software", f"mm-icons {VERSION}")
    return info


def recolor_safe(hex_):
    """True, wenn design_pro.recolor_png die Farbe unverändert lässt (weit weg von den Logo-Zielfarben oder exakt das Ziel)."""
    c = _rgb(hex_)
    exact = {_rgb(NAVY), _rgb(GOLD), (255, 255, 255)}
    if c in exact:
        return True
    return all(sum((a - b) ** 2 for a, b in zip(c, t)) >= 60 ** 2 for t in _RECOLOR_TARGETS)


# ============================================================================ Zeichenstift im 24er-Raster
class Pen:
    def __init__(self, draw, scale, color, width=STROKE, ox=0.0, oy=0.0):
        self.d, self.s, self.c, self.ox, self.oy = draw, scale, color, ox, oy
        self.w = width * scale

    def _p(self, x, y):
        return (self.ox + x * self.s, self.oy + y * self.s)

    def _dot(self, x, y, r):
        self.d.ellipse((x - r, y - r, x + r, y + r), fill=self.c)

    def line(self, *pts, closed=False):
        pts = [self._p(*p) for p in pts]
        if closed:
            pts.append(pts[0])
        r = self.w / 2
        for a, b in zip(pts, pts[1:]):
            self.d.line((a, b), fill=self.c, width=max(1, round(self.w)))
        for p in pts:
            self._dot(p[0], p[1], r)

    def poly(self, pts, closed=False):
        self.line(*pts, closed=closed)

    def arc(self, cx, cy, rx, a0, a1, ry=None, n=None):
        """Bogen von a0 nach a1 (Grad, 0° = rechts, 90° = unten – Bildschirmkoordinaten)."""
        ry = rx if ry is None else ry
        n = n or max(12, int(abs(a1 - a0) / 4))
        pts = [(cx + rx * math.cos(math.radians(a0 + (a1 - a0) * i / n)),
                cy + ry * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]
        self.line(*pts)
        return pts

    def circle(self, cx, cy, r, ry=None):
        self.arc(cx, cy, r, 0, 360, ry=ry, n=96)

    def rrect(self, x0, y0, x1, y1, r=2.0):
        pts = []
        for cx, cy, a in ((x1 - r, y0 + r, -90), (x1 - r, y1 - r, 0), (x0 + r, y1 - r, 90), (x0 + r, y0 + r, 180)):
            for i in range(10):
                t = math.radians(a + 9 * i + 0.0)
                pts.append((cx + r * math.cos(t), cy + r * math.sin(t)))
            t = math.radians(a + 90)
            pts.append((cx + r * math.cos(t), cy + r * math.sin(t)))
        self.line(*pts, closed=True)

    def bezier(self, p0, p1, p2, p3=None, n=28):
        pts = []
        for i in range(n + 1):
            t = i / n
            if p3 is None:
                x = (1 - t) ** 2 * p0[0] + 2 * (1 - t) * t * p1[0] + t * t * p2[0]
                y = (1 - t) ** 2 * p0[1] + 2 * (1 - t) * t * p1[1] + t * t * p2[1]
            else:
                x = ((1 - t) ** 3 * p0[0] + 3 * (1 - t) ** 2 * t * p1[0] + 3 * (1 - t) * t * t * p2[0]
                     + t ** 3 * p3[0])
                y = ((1 - t) ** 3 * p0[1] + 3 * (1 - t) ** 2 * t * p1[1] + 3 * (1 - t) * t * t * p2[1]
                     + t ** 3 * p3[1])
            pts.append((x, y))
        self.line(*pts)
        return pts

    def dot(self, x, y, r=1.05):
        px, py = self._p(x, y)
        self._dot(px, py, r * self.s)

    def fill_poly(self, pts):
        self.d.polygon([self._p(*p) for p in pts], fill=self.c)


# ============================================================================ Icons (24er-Raster, Rand 2)
def _haus(p):
    p.line((2.8, 11.2), (12, 3.6), (21.2, 11.2))
    p.line((5.2, 9.4), (5.2, 19.2))
    p.arc(7.2, 19.2, 2, 90, 180)
    p.line((18.8, 9.4), (18.8, 19.2))
    p.arc(16.8, 19.2, 2, 0, 90)
    p.line((7.2, 21.2), (10, 21.2))
    p.line((14, 21.2), (16.8, 21.2))
    p.line((10, 21.2), (10, 15.4), (14, 15.4), (14, 21.2))


def _schluessel(p):
    p.circle(7.6, 16.4, 4.4)
    p.dot(7.6, 16.4, 1.15)
    p.line((10.7, 13.3), (20.4, 3.6))
    p.line((17.2, 6.8), (19.8, 9.4))
    p.line((14.6, 9.4), (16.6, 11.4))


def _muenzen(p):
    """Münzstapel (vorn links) mit aufrecht stehender Münze dahinter (rechts)."""
    cx, rx, ry, top = 9.2, 6.2, 2.4, 10.4
    p.circle(cx, top, rx, ry=ry)
    p.line((cx - rx, top), (cx - rx, 18.6))
    p.line((cx + rx, top), (cx + rx, 18.6))
    for y in (14.5, 18.6):
        p.arc(cx, y, rx, 0, 180, ry=ry)
    # stehende Münze: nur der sichtbare Teil (nicht hinter dem Stapel)
    mx, my, mr = 16.4, 8.0, 5.0
    seg = []
    for i in range(121):
        t = math.radians(-160 + 300 * i / 120)
        x, y = mx + mr * math.cos(t), my + mr * math.sin(t)
        hidden = x < cx + rx + 1.3 and y > top - ry - 1.3
        if hidden:
            if len(seg) > 1:
                p.line(*seg)
            seg = []
        else:
            seg.append((x, y))
    if len(seg) > 1:
        p.line(*seg)
    p.arc(mx, my, 2.3, -150, 20)                          # Prägerand


def _bank(p):
    p.line((2.8, 9.0), (12, 3.4), (21.2, 9.0), (2.8, 9.0))
    for x in (6.0, 10.0, 14.0, 18.0):
        p.line((x, 11.6), (x, 17.4))
    p.line((3.6, 20.2), (20.4, 20.2))
    p.dot(12, 6.9, 0.95)


def _prozent(p):
    p.line((18.6, 5.4), (5.4, 18.6))
    p.circle(7.3, 7.3, 2.6)
    p.circle(16.7, 16.7, 2.6)


def _kalender(p):
    p.rrect(3.4, 5.2, 20.6, 20.8, 2.4)
    p.line((3.4, 10.2), (20.6, 10.2))
    p.line((8.2, 3.0), (8.2, 7.2))
    p.line((15.8, 3.0), (15.8, 7.2))
    for x in (8.0, 12.0, 16.0):
        p.dot(x, 13.9, 0.95)
    for x in (8.0, 12.0):
        p.dot(x, 17.4, 0.95)


def _diagramm(p):
    p.line((3.4, 3.4), (3.4, 18.4))
    p.arc(5.6, 18.4, 2.2, 90, 180)
    p.line((5.6, 20.6), (20.6, 20.6))
    p.line((8.2, 16.6), (8.2, 13.0))
    p.line((12.6, 16.6), (12.6, 7.8))
    p.line((17.0, 16.6), (17.0, 10.8))


def _schild(p):
    p.bezier((12, 21.2), (17.6, 18.6), (19.6, 15.2), (19.6, 11.2))
    p.line((19.6, 11.2), (19.6, 6.2), (12, 3.0), (4.4, 6.2), (4.4, 11.2))
    p.bezier((4.4, 11.2), (4.4, 15.2), (6.4, 18.6), (12, 21.2))
    p.line((8.6, 12.2), (11.0, 14.6), (15.6, 9.8))


def _dokument(p):
    p.line((14.0, 2.8), (6.6, 2.8))
    p.arc(6.6, 5.0, 2.2, 180, 270)
    p.line((4.4, 5.0), (4.4, 19.0))
    p.arc(6.6, 19.0, 2.2, 90, 180)
    p.line((6.6, 21.2), (17.4, 21.2))
    p.arc(17.4, 19.0, 2.2, 0, 90)
    p.line((19.6, 19.0), (19.6, 8.4), (14.0, 2.8))
    p.line((14.0, 2.8), (14.0, 8.4), (19.6, 8.4))
    p.line((8.4, 13.0), (15.6, 13.0))
    p.line((8.4, 16.6), (15.6, 16.6))
    p.line((8.4, 9.4), (10.6, 9.4))


def _paragraf(p):
    """klassisches „§“: oberer Haken, Mittelschlaufe, unterer Haken (punktsymmetrisch um 12/12)."""
    p.arc(12.0, 7.2, 3.6, -20, -270)
    p.bezier((8.4, 7.2), (8.4, 9.6), (15.6, 9.6), (15.6, 13.2))
    p.arc(12.0, 16.8, 3.6, 160, 0)
    p.bezier((15.6, 16.8), (15.6, 14.4), (8.4, 14.4), (8.4, 10.8))


def _ziel(p):
    p.circle(12, 12, 9.0)
    p.circle(12, 12, 5.4)
    p.dot(12, 12, 1.7)


def _trend(p):
    p.line((2.8, 17.6), (8.8, 11.6), (13.0, 15.8), (21.2, 7.6))
    p.line((15.4, 7.6), (21.2, 7.6), (21.2, 13.4))


def _rechner(p):
    p.rrect(4.8, 2.6, 19.2, 21.4, 2.4)
    p.rrect(8.0, 5.8, 16.0, 9.8, 0.9)
    for y in (13.6, 17.6):
        for x in (8.6, 12.0, 15.4):
            p.dot(x, y, 1.0)


def _waage(p):
    p.line((12, 3.6), (12, 20.4))
    p.line((7.6, 20.4), (16.4, 20.4))
    p.line((4.6, 6.6), (19.4, 6.6))
    p.dot(12, 3.9, 1.25)
    for cx in (4.6, 19.4):
        p.line((cx - 3.0, 13.2), (cx, 6.6), (cx + 3.0, 13.2))
        p.arc(cx, 13.2, 3.0, 0, 180, ry=2.3)


def _person(p):
    p.circle(12, 7.6, 4.2)
    p.line((4.6, 21.0), (4.6, 19.2))
    p.arc(9.0, 19.2, 4.4, 180, 270)
    p.line((9.0, 14.8), (15.0, 14.8))
    p.arc(15.0, 19.2, 4.4, 270, 360)
    p.line((19.4, 19.2), (19.4, 21.0))


def _foto(p):
    p.line((8.6, 6.2), (9.8, 4.2), (14.2, 4.2), (15.4, 6.2))
    p.rrect(2.6, 6.2, 21.4, 19.8, 2.4)
    p.circle(12, 12.8, 3.8)
    p.dot(17.8, 9.4, 0.9)


def _bild(p):
    p.rrect(2.8, 4.2, 21.2, 19.8, 2.4)
    p.circle(8.4, 9.4, 1.8)
    p.line((2.8, 17.4), (8.2, 12.8), (11.6, 15.8), (15.2, 11.6), (21.2, 17.0))


def _gebaeude(p):
    p.line((4.4, 21.0), (4.4, 5.0))
    p.arc(6.4, 5.0, 2.0, 180, 270)
    p.line((6.4, 3.0), (13.6, 3.0))
    p.arc(13.6, 5.0, 2.0, 270, 360)
    p.line((15.6, 5.0), (15.6, 21.0))
    p.line((15.6, 10.2), (18.4, 10.2))
    p.arc(18.4, 12.2, 2.0, 270, 360)
    p.line((20.4, 12.2), (20.4, 21.0))
    p.line((2.8, 21.0), (21.2, 21.0))
    for y in (7.2, 11.2, 15.2):
        p.dot(8.2, y, 0.95)
        p.dot(11.8, y, 0.95)
    p.line((8.6, 21.0), (8.6, 18.6), (11.4, 18.6), (11.4, 21.0))


def _euro(p):
    p.arc(13.8, 12, 7.2, 312, 48, n=72)
    p.line((3.6, 10.0), (13.0, 10.0))
    p.line((3.6, 14.0), (13.0, 14.0))


def _check(p):
    p.circle(12, 12, 9.0)
    p.line((7.8, 12.4), (10.8, 15.2), (16.4, 9.2))


def _info(p):
    p.circle(12, 12, 9.0)
    p.line((12, 11.0), (12, 16.6))
    p.dot(12, 7.7, 1.2)


def _warnung(p):
    p.line((12, 3.4), (21.0, 19.4), (3.0, 19.4), closed=True)
    p.line((12, 9.4), (12, 13.8))
    p.dot(12, 16.6, 1.1)


def _pfeil(p):
    p.line((4.0, 12.0), (20.0, 12.0))
    p.line((14.0, 6.0), (20.0, 12.0), (14.0, 18.0))


def _lupe(p):
    p.circle(10.4, 10.4, 6.6)
    p.line((15.2, 15.2), (20.6, 20.6))


def _werkzeug(p):
    """Maulschlüssel: runder Kopf mit Maulöffnung oben rechts, Griff mit gerundetem Ende unten links."""
    C, R = (15.6, 8.4), 5.4
    d = (1 / math.sqrt(2), -1 / math.sqrt(2))              # Achse Griff → Kopf
    n = (1 / math.sqrt(2), 1 / math.sqrt(2))
    E = (5.4, 18.6)

    def at(a, b, base=C):
        return (base[0] + a * d[0] + b * n[0], base[1] + a * d[1] + b * n[1])

    def ang(pt):
        return math.degrees(math.atan2(pt[1] - C[1], pt[0] - C[0]))

    js, hs = 1.7, 1.75
    jd, hd = math.sqrt(R * R - js * js), math.sqrt(R * R - hs * hs)
    p.line(at(0.4, -js), at(0.4, js), at(jd, js))
    a0, a1 = ang(at(jd, js)), ang(at(-hd, hs))
    if a1 < a0:
        a1 += 360
    p.arc(C[0], C[1], R, a0, a1)
    p.line(at(-hd, hs), (E[0] + hs * n[0], E[1] + hs * n[1]))
    a = math.degrees(math.atan2(n[1], n[0]))
    p.arc(E[0], E[1], hs, a, a + 180)
    p.line((E[0] - hs * n[0], E[1] - hs * n[1]), at(-hd, -hs))
    b0, b1 = ang(at(-hd, -hs)), ang(at(jd, -js))
    if b1 < b0:
        b1 += 360
    p.arc(C[0], C[1], R, b0, b1)
    p.line(at(jd, -js), at(0.4, -js))


def _uhr(p):
    p.circle(12, 12, 9.0)
    p.line((12, 6.8), (12, 12), (15.6, 14.2))


def _stift(p):
    p.line((4.0, 20.0), (5.0, 15.6), (15.8, 4.8), (16.2, 4.4))
    p.arc(17.8, 6.0, 2.26, 225, 405, n=30)
    p.line((19.4, 7.6), (8.6, 18.4), (4.0, 20.0))
    p.line((13.8, 6.8), (17.4, 10.4))





ICONS = {
    "haus": _haus, "schluessel": _schluessel, "muenzen": _muenzen, "bank": _bank, "prozent": _prozent,
    "kalender": _kalender, "diagramm": _diagramm, "schild": _schild, "dokument": _dokument, "paragraf": _paragraf,
    "ziel": _ziel, "trend": _trend, "rechner": _rechner, "waage": _waage, "person": _person, "foto": _foto,
    # Ergänzungen
    "bild": _bild, "gebaeude": _gebaeude, "euro": _euro, "check": _check, "info": _info, "warnung": _warnung,
    "pfeil": _pfeil, "lupe": _lupe, "werkzeug": _werkzeug, "uhr": _uhr, "stift": _stift,
}
ALIASES = {"steuer": "paragraf", "rendite": "prozent", "finanzierung": "bank", "cashflow": "muenzen",
           "objekt": "haus", "pruefung": "schild", "bewertung": "waage", "exit": "schluessel", "zeit": "kalender",
           "kamera": "foto", "wohnung": "gebaeude", "massnahmen": "werkzeug", "eingabe": "stift", "geld": "euro"}
NAMES = tuple(ICONS)


def _render(name, color_hex, out_px, bg=None, pad=None, bg_shape="circle"):
    """RGBA-Bild out_px × out_px. bg: Hintergrund-Hex (Kreis/abgerundetes Quadrat), pad: Innenabstand in 24er-Einheiten."""
    fn = ICONS[ALIASES.get(name, name)]
    big = out_px * SS
    im = Image.new("RGBA", (big, big), _rgb(color_hex) + (0,))
    if bg:
        pad = 5.0 if pad is None else pad
        bgl = Image.new("RGBA", (big, big), _rgb(bg) + (0,))
        dd = ImageDraw.Draw(bgl)
        if bg_shape == "circle":
            dd.ellipse((0, 0, big - 1, big - 1), fill=_rgb(bg) + (255,))
        else:
            dd.rounded_rectangle((0, 0, big - 1, big - 1), radius=big * 0.24, fill=_rgb(bg) + (255,))
    else:
        pad = 0.0 if pad is None else pad
    layer = Image.new("L", (big, big), 0)
    d = ImageDraw.Draw(layer)
    scale = big / (24 + 2 * pad)
    fn(Pen(d, scale, 255, ox=pad * scale, oy=pad * scale))
    small = layer.resize((out_px, out_px), Image.LANCZOS)
    glyph = Image.new("RGBA", (out_px, out_px), _rgb(color_hex) + (0,))
    glyph.putalpha(small)
    if bg:
        base = bgl.resize((out_px, out_px), Image.LANCZOS)
        base.alpha_composite(glyph)
        return base
    return glyph


def icon_path(name, color_hex=GOLD, px=20, bg=None, pad=None, bg_shape="circle", res=2):
    """PNG (res-fache Auflösung, Standard 2× für scharfe Darstellung) eines Icons in color_hex; Rückgabe: Dateipfad (Cache).
    bg: optionaler Hintergrund (Kreis bzw. bg_shape="square" abgerundet) – dann wird das Icon mit pad (Standard 5 von 24) eingerückt."""
    key = f"icon|{name}|{color_hex}|{px}|{bg}|{pad}|{bg_shape}|{res}|{VERSION}"
    path = _cache_file(key, "png")
    if not os.path.exists(path):
        img = _render(name, color_hex, px * res, bg=bg, pad=pad, bg_shape=bg_shape)
        tmp = path + f".{os.getpid()}.tmp"
        img.save(tmp, "PNG", pnginfo=_meta(), optimize=True)
        os.replace(tmp, path)
    return path


# ============================================================================ Platzierung (openpyxl)
def _row_px(ws, r):
    d = ws.row_dimensions.get(r) if hasattr(ws.row_dimensions, "get") else ws.row_dimensions[r]
    if d is not None and d.hidden:
        return 0
    h = d.height if d is not None and d.height else (ws.sheet_format.defaultRowHeight or 15)
    return round(h * 96 / 72)


def _col_px(ws, c):
    try:
        import core
        return core.col_px(ws, c)
    except Exception:                                    # pragma: no cover
        from openpyxl.utils import get_column_letter
        d = ws.column_dimensions[get_column_letter(c)]
        return int(((256 * (d.width or 8.43) + int(128 / 7)) / 256) * 7)


def _merged_extent(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col, rng.max_row, rng.max_col
    return row, col, row, col


def _normalize(ws, row, col, dx, dy):
    """Pixelversatz in (Zeile, Spalte, Rest-Offset) auflösen – Excel/LibreOffice mögen Offsets innerhalb der Zelle."""
    while dx >= _col_px(ws, col) > 0 or (dx > 0 and _col_px(ws, col) == 0):
        dx -= _col_px(ws, col)
        col += 1
    while dy >= _row_px(ws, row) > 0 or (dy > 0 and _row_px(ws, row) == 0):
        dy -= _row_px(ws, row)
        row += 1
    return row, col, max(0, dx), max(0, dy)


def place_image(ws, anchor_cell, path, w_px, h_px, dx=0, dy=0, align=None, valign=None):
    """Bild mit OneCellAnchor (feste Größe, verschiebt sich mit der Zelle, keine Verzerrung).
    align: None/"left"/"center"/"right" relativ zur Zelle bzw. zum Verbund ab anchor_cell (dx wirkt zusätzlich),
    valign: None/"top"/"middle"/"bottom"."""
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    letters, row = coordinate_from_string(anchor_cell)
    col = column_index_from_string(letters)
    if align or valign:
        r0, c0, r1, c1 = _merged_extent(ws, row, col)
        row, col = r0, c0
        wspan = sum(_col_px(ws, c) for c in range(c0, c1 + 1))
        hspan = sum(_row_px(ws, r) for r in range(r0, r1 + 1))
        if align == "center":
            dx += (wspan - w_px) / 2
        elif align == "right":
            dx += wspan - w_px
        if valign == "middle":
            dy += (hspan - h_px) / 2
        elif valign == "bottom":
            dy += hspan - h_px
    row, col, dx, dy = _normalize(ws, row, col, round(dx), round(dy))
    img = XLImage(path)
    img.width, img.height = w_px, h_px
    marker = AnchorMarker(col=col - 1, colOff=int(dx * EMU), row=row - 1, rowOff=int(dy * EMU))
    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(int(w_px * EMU), int(h_px * EMU)))
    ws.add_image(img)
    return img


def place_icon(ws, anchor_cell, name, color_hex=GOLD, px=20, dx=0, dy=0, bg=None, align=None, valign=None, pad=None,
               bg_shape="circle"):
    """Icon als Bild ins Blatt setzen (OneCellAnchor, Pixelversatz dx/dy ab linker oberer Ecke von anchor_cell;
    align/valign zentrieren optional in der Zelle bzw. im Verbund). Rückgabe: openpyxl-Image."""
    return place_image(ws, anchor_cell, icon_path(name, color_hex, px, bg=bg, pad=pad, bg_shape=bg_shape), px, px,
                       dx=dx, dy=dy, align=align, valign=valign)


# ============================================================================ Cover-Illustration
def _mix(c1, c2, t):
    return tuple(round(a + (b - a) * t) for a, b in zip(c1, c2))


PALETTES = {"navy": (NAVY, NAVY_2, GOLD, 1.0), "light": ("FBFAF7", "F1EEE6", "B08A3A", 0.85)}


def _gradient(w, h, palette="navy"):
    """Diagonaler Verlauf (navy: NAVY → NAVY_2, light: FBFAF7 → F1EEE6) mit weichem Lichtfeld rechts,
    ruhig und ohne Streifen (feines Korn als Dithering)."""
    a, b = _rgb(PALETTES[palette][0]), _rgb(PALETTES[palette][1])
    gw, gh = max(2, w // 8), max(2, h // 8)
    g = Image.new("RGB", (gw, gh))
    px = g.load()
    for y in range(gh):
        for x in range(gw):
            t = 0.72 * (x / (gw - 1)) + 0.28 * (1 - y / (gh - 1))
            dxl, dyl = (x / (gw - 1) - 0.74) / 0.42, (y / (gh - 1) - 0.62) / 0.75
            glow = max(0.0, 1 - (dxl * dxl + dyl * dyl)) ** 1.6
            t = min(1.0, max(0.0, t * 0.78 + glow * 0.42))
            px[x, y] = _mix(a, b, t)
    g = g.resize((w, h), Image.BICUBIC)
    noise = Image.effect_noise((w, h), 3.2).convert("L")
    n = Image.merge("RGB", (noise, noise, noise))
    g = ImageChops.add(g, n, scale=1.0, offset=-128)
    return g


class _Canvas:
    """Linienzeichnung mit mehreren Deckkraft-Ebenen (Alpha-Masken) in Gold."""

    def __init__(self, w, h, ss):
        self.w, self.h, self.ss = w, h, ss
        self.layers = {}

    def layer(self, alpha):
        if alpha not in self.layers:
            self.layers[alpha] = Image.new("L", (self.w * self.ss, self.h * self.ss), 0)
        return ImageDraw.Draw(self.layers[alpha])

    def line(self, alpha, pts, width=1.0):
        d = self.layer(alpha)
        s = self.ss
        P = [(x * s, y * s) for x, y in pts]
        wpx = max(1, round(width * s))
        d.line(P, fill=255, width=wpx, joint="curve")

    def rect(self, alpha, x0, y0, x1, y1, width=1.0, fill=False):
        d = self.layer(alpha)
        s = self.ss
        if fill:
            d.rectangle((x0 * s, y0 * s, x1 * s, y1 * s), fill=255)
        else:
            self.line(alpha, [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)], width)

    def mask_fill(self, alpha, poly):
        d = self.layer(alpha)
        d.polygon([(x * self.ss, y * self.ss) for x, y in poly], fill=255)

    def erase(self, poly):
        """Hintere Ebenen hinter einem Vordergrundkörper ausblenden (Linienzeichnung ohne Durchscheinen)."""
        for m in self.layers.values():
            ImageDraw.Draw(m).polygon([(x * self.ss, y * self.ss) for x, y in poly], fill=0)

    def composite(self, fade=None):
        """Alle Ebenen zu einer Alpha-Maske (0–255) zusammenführen; fade: Funktion x→Faktor (0–1) für weiches Auslaufen."""
        W, H = self.w, self.h
        acc = Image.new("L", (W, H), 0)
        for alpha, m in sorted(self.layers.items()):
            small = m.resize((W, H), Image.LANCZOS).point(lambda v, a=alpha: round(v * a))
            acc = ImageChops.lighter(acc, small)
        if fade:
            f = Image.new("L", (W, 1))
            f.putdata([round(255 * max(0.0, min(1.0, fade(x / (W - 1))))) for x in range(W)])
            f = f.resize((W, H))
            acc = ImageChops.multiply(acc, f)
        return acc


def _building(cv, rnd, x0, base, w, h, depth, u=1.0, roof="flat", style="grid", lit=0.0, occlude=False):
    """Ein Baukörper in Linienzeichnung. depth 0 = vorn (kräftig), 1 = Mitte, 2 = hinten (zart).
    style: "outline" (nur Kontur + zarte Lisenen), "ribbon" (Fensterbänder), "grid" (Lochfassade),
    "balcony" (modernes Wohnhaus: auskragende Balkonplatten, raumhohe Fenster)."""
    A = {0: 0.80, 1: 0.44, 2: 0.24}[depth]
    AW = {0: 0.42, 1: 0.22, 2: 0.12}[depth]
    lw = {0: 1.4, 1: 1.05, 2: 0.9}[depth] * u
    top = base - h
    face_top = top
    body = [(x0, base), (x0, top), (x0 + w, top), (x0 + w, base)]
    if roof == "gable":
        body = [(x0, base), (x0, top), (x0 + w / 2, top - w * 0.36), (x0 + w, top), (x0 + w, base)]
    elif roof == "shed":
        body = [(x0, base), (x0, top - w * 0.16), (x0 + w, top), (x0 + w, base)]
    elif roof == "setback":
        s = w * 0.2
        face_top = top + h * 0.10
        body = [(x0, base), (x0, face_top), (x0 + s, face_top), (x0 + s, top), (x0 + w - s, top),
                (x0 + w - s, face_top), (x0 + w, face_top), (x0 + w, base)]
    if occlude:
        cv.erase(body)
    cv.line(A, body, lw)
    wl = 0.8 * u
    if style == "outline":
        k = max(1, int(w / (14 * u)))
        for i in range(1, k):
            x = x0 + w * i / k
            cv.line(AW * 0.7, [(x, face_top + 6 * u), (x, base)], wl)
        return
    if style == "ribbon":
        fl = 11 * u
        y = face_top + 6 * u
        while y + fl < base - 4 * u:
            cv.rect(AW, x0 + 3.5 * u, y, x0 + w - 3.5 * u, y + fl * 0.5, wl)
            if rnd.random() < lit:
                seg = rnd.uniform(0.2, 0.45) * (w - 7 * u)
                sx = x0 + 3.5 * u + rnd.uniform(0, w - 7 * u - seg)
                cv.rect(round(min(0.6, AW * 1.9), 2), sx, y, sx + seg, y + fl * 0.5, fill=True)
            y += fl
        return
    if style == "balcony":
        fl = 17 * u
        y = face_top + fl
        mull = max(2, int(w / (12 * u)))
        while y < base - 2 * u:
            cv.line(A, [(x0 - 5 * u, y), (x0 + w + 5 * u, y)], lw)                # Balkonplatte
            cv.line(round(A * 0.55, 2), [(x0 - 5 * u, y - 5.5 * u), (x0 + w + 5 * u, y - 5.5 * u)], wl)  # Geländer
            for i in range(mull):                                                  # raumhohe Fenster
                fx0 = x0 + 2.5 * u + (w - 5 * u) * i / mull
                fx1 = x0 + 2.5 * u + (w - 5 * u) * (i + 1) / mull - 2.5 * u
                if rnd.random() < lit:
                    cv.rect(0.58, fx0, y - fl + 3 * u, fx1, y - 5.5 * u, fill=True)
                cv.rect(AW, fx0, y - fl + 3 * u, fx1, y - 0.5 * u, wl)
            y += fl
        return
    fw, fh, gx, gy = [v * u for v in {0: (5.0, 7.0, 4.5, 6.0), 1: (4.0, 5.5, 4.0, 5.5), 2: (3.2, 4.4, 3.4, 4.6)}[depth]]
    cols = max(1, int((w - gx) // (fw + gx)))
    used = cols * fw + (cols - 1) * gx
    xs = x0 + (w - used) / 2
    y = face_top + gy + 2 * u
    while y + fh < base - gy * 1.4:
        for i in range(cols):
            fx = xs + i * (fw + gx)
            if rnd.random() < lit:
                cv.rect(round(min(0.6, AW * 1.9), 2), fx, y, fx + fw, y + fh, fill=True)
            else:
                cv.rect(AW, fx, y, fx + fw, y + fh, wl)
        y += fh + gy
    if depth == 0:                                        # Eingang
        dw = min(9 * u, w * 0.22)
        cx = x0 + w / 2
        cv.line(A, [(cx - dw / 2, base), (cx - dw / 2, base - 11 * u), (cx + dw / 2, base - 11 * u),
                    (cx + dw / 2, base)], wl * 1.3)


def cover_art_path(width_px, height_px, background=True, focus="right", scale=2, seed=11, palette="navy",
                   intensity=1.0, badge=None):
    """Cover-Illustration: feine goldene Linienzeichnung (Wohnhaus-Fassaden, Stadtsilhouette, Horizont) auf Nachtblau.
    width_px/height_px: Anzeigegröße in Excel-Pixeln; die Datei hat scale-fache Auflösung.
    background=True → JPEG mit Verlauf NAVY → NAVY_2 (wird von design_pro nicht umgefärbt);
    background=False → transparentes PNG nur mit den Goldlinien (für eine navy gefüllte Zellfläche).
    focus: "right" (Motiv rechts, läuft nach links weich aus – Platz für Titel), "center" oder "full".
    palette: "navy" (Gold auf Nachtblau) oder "light" (Bronze-Gold auf warmem Weiß, z. B. Fotoplatzhalter);
    intensity: Deckkraft-Faktor der Linien (0,5 = zurückhaltend für Nebenflächen wie den Dashboard-Hero);
    badge: optionaler Icon-Name, der als Navy-Badge mit Gold-Icon mittig über dem Motiv sitzt."""
    ext = "jpg" if background else "png"
    key = f"cover|{width_px}|{height_px}|{background}|{focus}|{scale}|{seed}|{palette}|{intensity}|{badge}|{VERSION}"
    path = _cache_file(key, ext)
    if os.path.exists(path):
        return path
    W, H = int(width_px * scale), int(height_px * scale)
    u = scale                                              # 1 Einheit = 1 Excel-Pixel
    cv = _Canvas(W, H, 3)
    rnd = random.Random(seed)
    horizon = H * 0.84
    # Motivbereich
    if focus == "right":
        x_lo, x_hi = W * 0.40, W * 1.02
    elif focus == "center":
        x_lo, x_hi = W * 0.12, W * 0.88
    else:
        x_lo, x_hi = -W * 0.02, W * 1.02
    span = x_hi - x_lo

    # Konstruktionslinien (Architekturzeichnung, sehr zart)
    for i in range(1, 9):
        x = x_lo + span * i / 9 + rnd.uniform(-6, 6) * u
        cv.line(0.07, [(x, H * 0.10), (x, horizon)], 0.8 * u)
    for yy in (0.30, 0.52):
        cv.line(0.06, [(x_lo, H * yy), (x_hi, H * yy)], 0.8 * u)

    # Mond/Sonne – dünner Kreis mit Innenring
    mx, my, mr = x_lo + span * 0.70, H * 0.26, H * 0.12
    for rr, al in ((mr, 0.42), (mr * 0.78, 0.16)):
        pts = [(mx + rr * math.cos(t / 90 * math.pi), my + rr * math.sin(t / 90 * math.pi)) for t in range(181)]
        cv.line(al, pts, 1.0 * u)

    # hintere Skyline (zart, dicht)
    x = x_lo - 10 * u
    while x < x_hi:
        w = rnd.uniform(26, 58) * u
        h = rnd.uniform(0.28, 0.58) * H
        roof = rnd.choice(["flat", "flat", "setback", "shed"])
        _building(cv, rnd, x, horizon, w, h, 2, u, roof=roof, style="outline")
        x += w + rnd.uniform(-8, 6) * u

    # mittlere Reihe
    x = x_lo + 6 * u
    mids = []
    while x < x_hi:
        w = rnd.uniform(40, 70) * u
        h = rnd.uniform(0.24, 0.44) * H
        mids.append((x, w, h))
        x += w + rnd.uniform(14, 34) * u
    for (x, w, h) in mids:
        roof = rnd.choice(["flat", "gable", "flat", "shed"])
        style = "grid" if roof == "gable" else rnd.choice(["ribbon", "grid", "ribbon"])
        _building(cv, rnd, x, horizon, w, h, 1, u, roof=roof, style=style, lit=0.10, occlude=True)

    # Vordergrund: modernes Wohnhaus mit Balkonbändern + schmaler Nachbarturm + Giebelhaus
    fx = x_lo + span * (0.40 if focus != "full" else 0.44)
    fw, fh = span * 0.16, H * 0.52
    tower_x, tower_w, tower_h = fx + fw + 14 * u, span * 0.085, H * 0.64
    gable_x, gable_w, gable_h = fx - span * 0.13, span * 0.10, H * 0.30
    _building(cv, rnd, tower_x, horizon, tower_w, tower_h, 0, u, roof="setback", style="ribbon", lit=0.22,
              occlude=True)
    _building(cv, rnd, gable_x, horizon, gable_w, gable_h, 0, u, roof="gable", style="grid", lit=0.12, occlude=True)
    _building(cv, rnd, fx, horizon, fw, fh, 0, u, roof="flat", style="balcony", lit=0.13, occlude=True)
    # Dachgarten-Pergola auf dem Wohnhaus
    top = horizon - fh
    cv.line(0.5, [(fx + fw * 0.12, top), (fx + fw * 0.12, top - 12 * u), (fx + fw * 0.58, top - 12 * u),
                  (fx + fw * 0.58, top)], 1.0 * u)
    for k in range(1, 6):
        xx = fx + fw * 0.12 + (fw * 0.46) * k / 6
        cv.line(0.32, [(xx, top - 12 * u), (xx, top)], 0.8 * u)
    # Bäume (feine Kreise auf Stamm) vor dem Wohnhaus
    for tx in (fx - 12 * u, fx + fw + 6 * u, gable_x - 16 * u):
        r = 8 * u
        cv.erase([(tx - r, horizon), (tx - r, horizon - 26 * u), (tx + r, horizon - 26 * u), (tx + r, horizon)])
        pts = [(tx + r * math.cos(t / 30 * math.pi), horizon - 18 * u + r * 1.15 * math.sin(t / 30 * math.pi))
               for t in range(61)]
        cv.line(0.55, pts, 1.0 * u)
        cv.line(0.55, [(tx, horizon - 18 * u + r * 1.15), (tx, horizon)], 1.0 * u)

    # Horizont und Spiegelung
    cv.line(0.85, [(x_lo - span * 0.1, horizon), (W, horizon)], 1.3 * u)
    for k, al in ((1, 0.22), (2, 0.14), (3, 0.08)):
        yy = horizon + k * 6.5 * u
        cv.line(al, [(x_lo + span * (0.12 + 0.06 * k), yy), (x_hi - span * 0.05 * k, yy)], 0.9 * u)

    if focus == "right":
        def fade(t):
            return min(1.0, max(0.0, (t - 0.34) / 0.30)) ** 1.4
    elif focus == "center":
        def fade(t):
            return min(1.0, max(0.0, min(t, 1 - t) / 0.16))
    else:
        fade = None
    if badge:
        bs = int(min(W, H) * 0.30)
        bx, by = (W - bs) // 2, int(H * 0.44 - bs / 2)
        rr, cx0, cy0 = bs / 2 + 7 * u, bx + bs / 2, by + bs / 2
        cv.erase([(cx0 + rr * math.cos(t * math.pi / 36), cy0 + rr * math.sin(t * math.pi / 36)) for t in range(72)])
    mask = cv.composite(fade)
    line_c, k = PALETTES[palette][2], PALETTES[palette][3] * intensity
    if k != 1.0:
        mask = mask.point(lambda v: min(255, round(v * k)))
    gold = Image.new("RGBA", (W, H), _rgb(line_c) + (0,))
    gold.putalpha(mask)
    if badge:
        b = _render(badge, GOLD, bs, bg=NAVY, pad=6.0)
        gold.alpha_composite(b, (bx, by))
    tmp = path + f".{os.getpid()}.tmp"
    if background:
        base = _gradient(W, H, palette).convert("RGBA")
        base.alpha_composite(gold)
        base.convert("RGB").save(tmp, "JPEG", quality=95, subsampling=0, optimize=True)
    else:
        gold.save(tmp, "PNG", pnginfo=_meta(), optimize=True)
    os.replace(tmp, path)
    return path


def photo_placeholder_path(width_px, height_px, scale=2):
    """Fotoplatzhalter (Exposé): zarte Bronze-Linienzeichnung auf warmem Weiß, mittig ein Navy-Badge mit Kamera."""
    return cover_art_path(width_px, height_px, background=True, focus="center", scale=scale, seed=5,
                          palette="light", intensity=0.75, badge="foto")


# ============================================================================ Monogramm
def _serif_font(size):
    for f in ("/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
              "/usr/share/fonts/truetype/freefont/FreeSerif.ttf"):
        if os.path.exists(f):
            return ImageFont.truetype(f, size)
    return ImageFont.load_default()


def monogram_path(px=120, letters="MM", fg=WHITE, ring=GOLD, res=2):
    """Monogramm: Serifen-Initialen (fg) in doppeltem Goldring (außen 1,4 px, innen 0,7 px) auf transparent."""
    key = f"mono|{px}|{letters}|{fg}|{ring}|{res}|{VERSION}"
    path = _cache_file(key, "png")
    if os.path.exists(path):
        return path
    S = px * res * SS
    ring_m = Image.new("L", (S, S), 0)
    d = ImageDraw.Draw(ring_m)
    u = S / 120.0
    for r, w in ((58.6, 1.5), (53.2, 0.75)):
        d.ellipse((S / 2 - r * u, S / 2 - r * u, S / 2 + r * u, S / 2 + r * u), outline=255, width=max(1, round(w * u)))
    # kleine Rautenpunkte oben/unten im Ringzwischenraum
    for ang in (90, 270):
        cx = S / 2 + 55.9 * u * math.cos(math.radians(ang))
        cy = S / 2 + 55.9 * u * math.sin(math.radians(ang))
        k = 1.5 * u
        d.polygon([(cx, cy - k), (cx + k, cy), (cx, cy + k), (cx - k, cy)], fill=255)
    txt_m = Image.new("L", (S, S), 0)
    dt = ImageDraw.Draw(txt_m)
    font = _serif_font(round(44 * u))
    bbox = dt.textbbox((0, 0), letters, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    dt.text(((S - tw) / 2 - bbox[0], (S - th) / 2 - bbox[1]), letters, font=font, fill=255)
    # feine Linie unter den Initialen
    d.line((S / 2 - 12 * u, S / 2 + th / 2 + 8 * u, S / 2 + 12 * u, S / 2 + th / 2 + 8 * u), fill=255,
           width=max(1, round(0.9 * u)))                   # feine Goldlinie unter den Initialen
    out_px = px * res
    ring_s = ring_m.resize((out_px, out_px), Image.LANCZOS)
    txt_s = txt_m.resize((out_px, out_px), Image.LANCZOS)
    im = Image.new("RGBA", (out_px, out_px), _rgb(ring) + (0,))
    im.putalpha(ring_s)
    t = Image.new("RGBA", (out_px, out_px), _rgb(fg) + (0,))
    t.putalpha(txt_s)
    im.alpha_composite(t)
    tmp = path + f".{os.getpid()}.tmp"
    im.save(tmp, "PNG", pnginfo=_meta())
    os.replace(tmp, path)
    return path


# ============================================================================ Galerie (Test)
def gallery_path(px=24, path=None, colors=None):
    """Übersicht aller Icons: je Farbe eine Zeile auf Weiß bzw. Navy, 1× und 2× Größe, dazu Badges."""
    colors = colors or ((GOLD, WHITE), (NAVY, WHITE), (GOLD, NAVY), (WHITE, NAVY), (BLUE, "FBFAF7"))
    n = len(NAMES)
    cell = 64
    W, H = n * cell + 40, len(colors) * (cell + 34) + 130
    im = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    d = ImageDraw.Draw(im)
    font = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 10)
    y = 10
    for fg, bgc in colors:
        d.rectangle((0, y, W, y + cell + 30), fill=_rgb(bgc))
        for i, nm in enumerate(NAMES):
            x = 20 + i * cell
            ic = Image.open(icon_path(nm, fg, px)).resize((px, px), Image.LANCZOS)   # 1× wie in Excel
            im.alpha_composite(ic, (x, y + 8))
            ic2 = Image.open(icon_path(nm, fg, px))                                   # 2× (Retina)
            im.alpha_composite(ic2, (x, y + 36))
            d.text((x, y + cell + 18), nm[:10], fill=(128, 128, 128), font=font)
        y += cell + 34
    # Badges
    x = 20
    for i, nm in enumerate(NAMES[:14]):
        b = Image.open(icon_path(nm, GOLD, 32, bg=NAVY)).resize((32, 32), Image.LANCZOS)
        im.alpha_composite(b, (x + i * 44, y + 10))
        b = Image.open(icon_path(nm, NAVY, 32, bg="F6EFDF")).resize((32, 32), Image.LANCZOS)
        im.alpha_composite(b, (x + i * 44 + 14 * 44 + 30, y + 10))
    path = path or os.path.join(_CACHE, "gallery.png")
    im.save(path)
    return path


if __name__ == "__main__":                               # python excel/icons.py [ordner] → Galerie + Cover-Art
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else _CACHE
    os.makedirs(out, exist_ok=True)
    print(gallery_path(path=os.path.join(out, "icons_gallery.png")))
    print(cover_art_path(1100, 330))
    print(monogram_path(150))
