"""
Visual / geometric checks for the Whiteboard -> OpenBoard .ubz converters.

For every sample export and both converter scripts this:
  1. converts the export with an *instrumented temporary copy* of the converter (the only
     change is one XML comment recording the page's centre offset, so board coordinates can
     be recovered exactly -- the scripts in the repo are never modified);
  2. renders the Whiteboard HTML in headless Chromium and measures where Whiteboard really
     puts every text, sticker, note, shape outline, connector arrowhead and ink stroke, and
     which colours it paints each note;
  3. renders the .ubz page the way OpenBoard lays it out and measures the same things;
  4. writes tests/out/visual/report.html (side-by-side screenshots + accuracy tables) and
     results.json, and exits non-zero if any tolerance is exceeded.

Only dependency: Playwright (pip install playwright; python -m playwright install chromium).
Run tests/Setup-VisualTests.ps1 once to set that up in tests/.venv.

Usage (from the repo root):
    tests\\.venv\\Scripts\\python tests\\visual\\run_visual_checks.py
    ... --sample AssumptionGrid ImageBoard --version v2_solid
"""
from __future__ import annotations

import argparse
import base64
import html as htmllib
import json
import math
import os
import platform
import re
import shutil
import subprocess
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parents[2]
SVG = "{http://www.w3.org/2000/svg}"
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"
UB_PARENT = "{http://uniboard.mnemis.com/document}parent"
VERSIONS = ("v2_solid", "v3_gradient")

# Pass/fail tolerances in board pixels; note_color is the largest channel difference (0-255)
# at five points of each note, ink_width the stroke width difference in board px.
TOL = {"text": 5.0, "sticker": 0.5, "note": 0.5, "shape": 1.5, "arrow": 0.5,
       "note_color": 3.0, "ink": 0.5, "ink_width": 0.1}


# ----------------------------------------------------------------------------------------
# Conversion with an instrumented copy of the converter
# ----------------------------------------------------------------------------------------
CENTER_ANCHOR = "        [void]$sb.AppendLine(('  <rect fill=\"white\""
CENTER_LINE = ("        [void]$sb.AppendLine(('  <!-- TESTCENTER {0} {1} -->' -f "
               "$centerX.ToString($ci), $centerY.ToString($ci)))")


def instrument(script: Path, dest: Path) -> Path:
    raw = script.read_bytes().decode("utf-8-sig")
    if "TESTCENTER" not in raw:
        nl = "\r\n" if "\r\n" in raw else "\n"
        if raw.count(CENTER_ANCHOR) != 1:
            raise SystemExit(f"Can't instrument {script.name}: page-SVG writer not found.")
        raw = raw.replace(CENTER_ANCHOR, CENTER_LINE + nl + CENTER_ANCHOR, 1)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(raw.encode("utf-8-sig"))   # BOM keeps Windows PowerShell 5.1 on UTF-8
    return dest


def find_powershell(explicit: str | None) -> list[str]:
    if explicit:
        return [explicit]
    candidates = ["powershell.exe", "pwsh"] if platform.system() == "Windows" else ["pwsh"]
    for c in candidates:
        if shutil.which(c):
            return [c]
    raise SystemExit("No PowerShell found (looked for: %s)." % ", ".join(candidates))


def convert(ps: list[str], script: Path, html: Path, out_dir: Path) -> tuple[Path | None, str]:
    out_dir.mkdir(parents=True, exist_ok=True)
    # -Command with the call operator rather than -File: with stdin redirected (CI, IDE and agent
    # shells), "powershell.exe -File" feeds stdin to the script as pipeline input, so the
    # converter's process {} block runs once per stdin line -- zero times for empty stdin.
    q = lambda s: "'" + str(s).replace("'", "''") + "'"
    cmd = ps + ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                f"& {q(script)} -InputPath {q(html)} -OutputDirectory {q(out_dir)} -Force; exit $LASTEXITCODE"]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    ubz = out_dir / (html.stem + ".ubz")
    log = (p.stdout + p.stderr).strip()
    return (ubz if p.returncode == 0 and ubz.exists() else None), log


# ----------------------------------------------------------------------------------------
# Whiteboard side: measure the export in the browser
# ----------------------------------------------------------------------------------------
JS_HTML = r"""() => {
  const cal = [], texts = [], stickers = [], shapes = [], arrows = [], notes = [], noteStyles = [], inks = [];
  const pt = (m, x, y) => [m.a*x + m.c*y + m.e + scrollX, m.b*x + m.d*y + m.f + scrollY];
  document.querySelectorAll('div.anchor[data-whiteboard-type]').forEach(a => {
    const st = a.getAttribute('style') || '', r = a.getBoundingClientRect();
    const L = /left:\s*(-?[\d.]+)px/.exec(st), T = /top:\s*(-?[\d.]+)px/.exec(st);
    if (L && T && !/transform/.test(st)) cal.push([+L[1], +T[1], r.left + scrollX, r.top + scrollY]);
    const type = a.dataset.whiteboardType;
    const spans = [...a.querySelectorAll('span[data-text="true"]')];
    const t = spans.map(s => s.textContent).join('');
    if (t.trim()) {
      const rg = document.createRange(), last = spans[spans.length - 1], lf = last.firstChild || last;
      rg.setStart(spans[0].firstChild || spans[0], 0); rg.setEnd(lf, lf.length || 0);
      const b = rg.getBoundingClientRect();
      texts.push({type, text: t, x: b.left + scrollX, y: b.top + scrollY, w: b.width, h: b.height});
    }
    if (type === 'ReactionStickers') {
      const i = a.querySelector('img').getBoundingClientRect();
      stickers.push([i.left + scrollX, i.top + scrollY, i.width, i.height]);
    }
    if (type === 'Note') {
      const bg = a.querySelector('.textBoxBackground'), b = bg.getBoundingClientRect(), cs = getComputedStyle(bg);
      notes.push([b.left + scrollX, b.top + scrollY, b.width, b.height]);
      noteStyles.push({color: cs.backgroundColor, image: cs.backgroundImage});
    }
    // Ink: each stroke's centreline (the hidden hit-test polyline) and its pen width, twice
    // the radius of the round joins in the visible outline path.
    if (type === 'InkGroup') a.querySelectorAll('g.inkStroke').forEach(g => {
      const pl = g.querySelector('polyline'), p = g.querySelector('path');
      if (!pl) return;
      const m = pl.getScreenCTM(), n = pl.points.numberOfItems, pts = [];
      for (let i = 0; i < n; i++) { const q = pl.points.getItem(i); pts.push(pt(m, q.x, q.y)); }
      const r = /[Aa]\s*([\d.]+)/.exec(p ? p.getAttribute('d') : '');
      inks.push({pts, w: r ? 2 * r[1] * Math.sqrt(Math.abs(m.a * m.d - m.b * m.c)) : null});
    });
    if (type === 'Shape') {
      const svg = a.querySelector('svg.shape'), p = svg && svg.querySelector('g > path');
      if (p) {
        const len = p.getTotalLength(), n = Math.max(200, Math.ceil(len / 2)), m = p.getScreenCTM(), pts = [];
        for (let i = 0; i < n; i++) { const q = p.getPointAtLength(len * i / n); pts.push(pt(m, q.x, q.y)); }
        shapes.push({label: (svg.getAttribute('aria-label') || '').split(',')[0], pts});
      }
    }
    if (type === 'Connector') a.querySelectorAll('svg g path[transform]').forEach(p => {
      const m = p.getScreenCTM(), n = p.getTotalLength();
      arrows.push([0, n / 2, n].map(l => { const q = p.getPointAtLength(l); return pt(m, q.x, q.y); }));
    });
  });
  // The anchor divs themselves are 0 x 0 (their content overflows them), so take the union
  // of everything drawn inside them; comment threads are not board objects.
  const all = [...document.querySelectorAll('div.anchor[data-whiteboard-type]:not([data-whiteboard-type=CommentThread]) *')]
              .map(e => e.getBoundingClientRect()).filter(r => r.width > 0 && r.height > 0);
  const bbox = all.length ? [Math.min(...all.map(r => r.left)) + scrollX, Math.min(...all.map(r => r.top)) + scrollY,
                             Math.max(...all.map(r => r.right)) + scrollX, Math.max(...all.map(r => r.bottom)) + scrollY] : null;
  return {cal, texts, stickers, shapes, arrows, notes, noteStyles, inks, bbox};
}"""


def linfit(xs, ys):
    n = len(xs); mx = sum(xs) / n; my = sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    k = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sxx
    return k, my - k * mx


def measure_html(page, html: Path, shot: Path) -> dict:
    page.goto(html.as_uri()); page.wait_for_timeout(1200)
    d = page.evaluate(JS_HTML)
    if len(d["cal"]) < 2:
        raise RuntimeError("fewer than 2 untransformed anchors; can't map screen to board coordinates")
    kx, ox = linfit([c[0] for c in d["cal"]], [c[2] for c in d["cal"]])
    ky, oy = linfit([c[1] for c in d["cal"]], [c[3] for c in d["cal"]])
    to_b = lambda x, y: ((x - ox) / kx, (y - oy) / ky)
    if d["bbox"]:
        x0, y0, x1, y1 = d["bbox"]; pad = 12
        page.screenshot(path=str(shot), full_page=True,
                        clip={"x": max(0, x0 - pad), "y": max(0, y0 - pad), "width": x1 - x0 + 2 * pad, "height": y1 - y0 + 2 * pad})
    return {
        "texts": [dict(t, bx=to_b(t["x"], t["y"])[0], by=to_b(t["x"], t["y"])[1], bh=t["h"] / ky) for t in d["texts"]],
        "stickers": [(*to_b(s[0], s[1]), s[2] / kx, s[3] / ky) for s in d["stickers"]],
        "notes": [(*to_b(s[0], s[1]), s[2] / kx, s[3] / ky) for s in d["notes"]],
        "note_paints": [note_paint(s) for s in d["noteStyles"]],
        "shapes": [{"label": s["label"], "pts": [to_b(*q) for q in s["pts"]]} for s in d["shapes"]],
        "arrows": [[to_b(*q) for q in a] for a in d["arrows"]],
        "inks": [{"pts": [to_b(*q) for q in s["pts"]], "w": s["w"] / kx if s["w"] else None} for s in d["inks"]],
    }


def parse_rgbs(s: str) -> list[tuple]:
    return [tuple(int(v) for v in m) for m in re.findall(r"rgba?\((\d+),\s*(\d+),\s*(\d+)", s)]


def note_paint(style: dict) -> tuple:
    """A note's colours from the browser's computed style: (solid, gradient), gradient being
    (start, end, angle in CSS degrees) or None. A note with a transparent background-color
    (Whiteboard's newer note style) is expected to use its gradient's end colour when solid."""
    grad = None
    m = re.search(r"linear-gradient\(((?:[^()]|\([^()]*\))*)\)", style["image"])
    if m:
        cols = parse_rgbs(m.group(1)); ang = re.match(r"\s*([-\d.]+)deg", m.group(1))
        if len(cols) >= 2:
            grad = (cols[0], cols[-1], float(ang.group(1)) if ang else 180.0)
    transparent = not parse_rgbs(style["color"]) or re.search(r"rgba\([^)]*,\s*0\)", style["color"])
    solid = (grad[1] if grad else None) if transparent else parse_rgbs(style["color"])[0]
    return solid, grad


def expected_note_color(box, paint, solid_only, x, y):
    solid, grad = paint
    if solid_only or not grad:
        return solid
    a = math.radians(grad[2]); dx, dy = math.sin(a), -math.cos(a)
    t = ((x - box[0] - box[2] / 2) * dx + (y - box[1] - box[3] / 2) * dy) / (abs(box[2] * dx) + abs(box[3] * dy)) + 0.5
    t = min(1.0, max(0.0, t))
    return tuple(s + (e - s) * t for s, e in zip(grad[0], grad[1]))


def point_in_poly(x, y, poly) -> bool:
    inside = False
    for i in range(len(poly)):
        (x1, y1), (x2, y2) = poly[i], poly[i - 1]
        if (y1 > y) != (y2 > y) and x < x1 + (y - y1) * (x2 - x1) / (y2 - y1):
            inside = not inside
    return inside


NOTE_SAMPLES = ((0.08, 0.08), (0.92, 0.08), (0.5, 0.5), (0.08, 0.92), (0.92, 0.92))


def note_color_err(box, paint, solid_only, polys) -> float:
    """Largest channel difference between the colour Whiteboard paints and the colour of the
    topmost output polygon, at five points inside the note."""
    worst = 0.0
    for fx, fy in NOTE_SAMPLES:
        x, y = box[0] + fx * box[2], box[1] + fy * box[3]
        exp = expected_note_color(box, paint, solid_only, x, y)
        got = None
        for pts, rgb in polys:
            if point_in_poly(x, y, pts):
                got = rgb
        if exp is None:
            continue
        if got is None:
            return 255.0
        worst = max(worst, max(abs(e - g) for e, g in zip(exp, got)))
    return round(worst, 1)


def hex_rgb(h: str) -> tuple:
    h = h.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ----------------------------------------------------------------------------------------
# OpenBoard side: read + render the .ubz page
# ----------------------------------------------------------------------------------------
def ubz_page(ubz: Path):
    z = zipfile.ZipFile(ubz)
    raw = z.read("page000.svg").decode("utf-8")
    m = re.search(r"TESTCENTER (\S+) (\S+)", raw)
    center = (float(m.group(1)), float(m.group(2))) if m else None
    return z, raw, center


def browser_svg(z: zipfile.ZipFile, raw: str) -> str:
    """Make page000.svg displayable: inline images, expand OpenBoard's itemTextContent."""
    def inline(m):
        href = m.group(1); ext = href.rsplit(".", 1)[-1].lower()
        mime = {"jpg": "jpeg", "svg": "svg+xml"}.get(ext, ext)
        data = base64.b64encode(z.read(href)).decode()
        return f'xlink:href="data:image/{mime};base64,{data}" preserveAspectRatio="none"'
    s = re.sub(r'xlink:href="(images/[^"]+)"', inline, raw)
    s = re.sub(r"<itemTextContent>(.*?)</itemTextContent>",
               lambda m: '<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;overflow:visible;'
                         'overflow-wrap:break-word">' + htmllib.unescape(m.group(1)).replace("&nbsp;", "&#160;") + "</div>",
               s, flags=re.S)
    return re.sub(r"^<\?xml[^>]*>\s*", "", s)


JS_SVG = r"""() => [...document.querySelectorAll('foreignObject')].map(fo => {
  const p = fo.querySelector('p'); if (!p || !p.firstChild) return null;
  const rg = document.createRange(); rg.selectNodeContents(p.firstChild);
  const r = rg.getBoundingClientRect(), svg = document.querySelector('svg'), m = svg.getScreenCTM().inverse();
  const P = (x, y) => { const q = svg.createSVGPoint(); q.x = x; q.y = y; return q.matrixTransform(m); };
  const a = P(r.left, r.top), b = P(r.right, r.bottom);
  // Line breaks are <br> elements (Qt collapses raw newlines), which textContent drops.
  const text = [...p.childNodes].map(n => n.nodeName.toLowerCase() === 'br' ? '\n' : n.textContent).join('');
  return {text, x: a.x, y: a.y, h: b.y - a.y};
}).filter(Boolean)"""


def measure_ubz(page, ubz: Path, shot: Path) -> dict:
    z, raw, center = ubz_page(ubz)
    if center is None:
        raise RuntimeError("no TESTCENTER marker (converter was not instrumented)")
    cx, cy = center
    root = ET.fromstring(raw.encode("utf-8"))
    vb = [float(v) for v in root.get("viewBox").split()]
    scale = min(1.0, 2400 / max(vb[2], vb[3]))
    w, h = int(vb[2] * scale), int(vb[3] * scale)
    page.set_viewport_size({"width": max(w, 200), "height": max(h, 200)})
    page.set_content("<html><body style='margin:0;background:#fff'>"
                     + browser_svg(z, raw).replace("<svg ", f"<svg width='{w}' height='{h}' ", 1) + "</body></html>")
    page.wait_for_timeout(600)
    page.screenshot(path=str(shot), full_page=True)
    texts = [dict(t, bx=t["x"] + cx, by=t["y"] + cy) for t in page.evaluate(JS_SVG)]

    stickers, groups, lines, notes, note_polys, inks = [], {}, set(), [], [], []
    pts_of = lambda el: [(float(a) + cx, float(b) + cy) for a, b in (xy.split(",") for xy in el.get("points").split())]
    for el in root:
        tag = el.tag.replace(SVG, "")
        if tag == "image" and el.get(XLINK_HREF, "").endswith(".svg"):
            m = [float(v) for v in re.findall(r"-?[\d.]+(?:[eE]-?\d+)?", el.get("transform"))]
            stickers.append((m[4] + cx, m[5] + cy, float(el.get("width")) * m[0], float(el.get("height")) * m[3]))
        par = el.get(UB_PARENT)
        if tag == "polyline" and el.get("stroke-linejoin") == "round":   # ink stroke
            inks.append({"pts": pts_of(el), "w": float(el.get("stroke-width"))})
            continue
        if tag == "polygon" and not par:   # v2 note: the only ungrouped fill polygon
            notes.append(pts_of(el))
            note_polys.append([(pts_of(el), hex_rgb(el.get("fill")))])
        if tag in ("polygon", "polyline") and par:
            pts = [(float(a) + cx, float(b) + cy) for a, b in (xy.split(",") for xy in el.get("points").split())]
            g = groups.setdefault(par, {"polygon": 0, "polyline": 0})
            g[tag] += 1
            g.setdefault(tag + "_pts", pts[:-1] if tag == "polyline" else pts)
        if tag == "line" and par:
            lines.add(par)
    # A shape = a group with at most one fill polygon (v3 gradient notes have many bands).
    shapes = [g.get("polyline_pts", g.get("polygon_pts")) for k, g in groups.items()
              if k not in lines and g["polygon"] <= 1]
    arrows = [g_pts for par in lines for g_pts in
              [pts for el in root if el.tag == SVG + "polyline" and el.get(UB_PARENT) == par
               for pts in [[(float(a) + cx, float(b) + cy) for a, b in (xy.split(",") for xy in el.get("points").split())]]]]
    # v3 note: one group of gradient bands (each band but the last overlaps the next, inside
    # the note).
    for par in {el.get(UB_PARENT) for el in root if el.tag == SVG + "polygon" and el.get(UB_PARENT)}:
        band_els = [el for el in root if el.tag == SVG + "polygon" and el.get(UB_PARENT) == par]
        if len(band_els) > 1:
            bands = [pts_of(el) for el in band_els]
            pts = [p for band in bands for p in band]
            notes.append([(min(p[0] for p in pts), min(p[1] for p in pts)),
                          (max(p[0] for p in pts), max(p[1] for p in pts))])
            note_polys.append([(pts_of(el), hex_rgb(el.get("fill"))) for el in band_els])
    note_boxes = [(min(p[0] for p in n), min(p[1] for p in n),
                   max(p[0] for p in n) - min(p[0] for p in n), max(p[1] for p in n) - min(p[1] for p in n)) for n in notes]
    return {"texts": texts, "stickers": stickers, "notes": note_boxes, "note_polys": note_polys,
            "shapes": shapes, "arrows": arrows, "inks": inks}


# ----------------------------------------------------------------------------------------
# Comparisons
# ----------------------------------------------------------------------------------------
def seg_dist(p, a, b):
    abx, aby = b[0] - a[0], b[1] - a[1]; L = abx * abx + aby * aby
    t = 0 if L == 0 else max(0, min(1, ((p[0] - a[0]) * abx + (p[1] - a[1]) * aby) / L))
    return math.hypot(p[0] - a[0] - t * abx, p[1] - a[1] - t * aby)


def poly_dist(p, poly):
    return min(seg_dist(p, poly[i], poly[(i + 1) % len(poly)]) for i in range(len(poly)))


def path_dist(p, pts):
    """Distance from p to an open polyline."""
    return min(seg_dist(p, pts[i], pts[i + 1]) for i in range(len(pts) - 1)) if len(pts) > 1 else math.dist(p, pts[0])


def compare(src: dict, out: dict, solid_only: bool) -> dict:
    res = {}
    # Texts: match by exact text, in order.
    pool = list(out["texts"]); rows = []
    for s in src["texts"]:
        j = next((i for i, o in enumerate(pool) if o["text"] == s["text"]), None)
        if j is None:
            rows.append({"text": s["text"][:50], "missing": True}); continue
        o = pool.pop(j)
        rows.append({"text": s["text"][:50], "dx": round(o["bx"] - s["bx"], 1), "dy": round(o["by"] - s["by"], 1),
                     "err": round(math.hypot(o["bx"] - s["bx"], o["by"] - s["by"]), 1)})
    errs = [r["err"] for r in rows if "err" in r]
    res["text"] = {"expected": len(src["texts"]), "matched": len(errs),
                   "median": round(sorted(errs)[len(errs) // 2], 1) if errs else None,
                   "max": max(errs) if errs else None, "rows": rows}
    # Stickers: same order in both.
    se = [max(abs(a - b) for a, b in zip(s, o)) for s, o in zip(src["stickers"], out["stickers"])]
    res["sticker"] = {"expected": len(src["stickers"]), "matched": len(out["stickers"]), "max": round(max(se), 2) if se else None}
    # Notes: pair each true note box with the nearest output box; largest edge error.
    # The note's colour at five points: solid fill, or the gradient at the gradient's angle.
    ne, nc, pool = [], [], list(zip(out["notes"], out["note_polys"]))
    for s, paint in zip(src["notes"], src["note_paints"]):
        if not pool: break
        c = (s[0] + s[2] / 2, s[1] + s[3] / 2)
        j = min(range(len(pool)), key=lambda i: math.hypot(pool[i][0][0] + pool[i][0][2] / 2 - c[0], pool[i][0][1] + pool[i][0][3] / 2 - c[1]))
        o, polys = pool.pop(j)
        ne.append(max(abs(s[0] - o[0]), abs(s[1] - o[1]), abs(s[0] + s[2] - o[0] - o[2]), abs(s[1] + s[3] - o[1] - o[3])))
        nc.append(note_color_err(s, paint, solid_only, polys))
    res["note"] = {"expected": len(src["notes"]), "matched": len(out["notes"]), "max": round(max(ne), 2) if ne else None}
    res["note_color"] = {"max": max(nc) if nc else None}
    # Ink: pair each true stroke with the output stroke nearest its ends; largest distance
    # either way between the two centrelines, and the width difference.
    ie, iw, pool = [], [], list(out["inks"])
    for s in src["inks"]:
        if not pool: break
        j = min(range(len(pool)), key=lambda i: math.dist(pool[i]["pts"][0], s["pts"][0]) + math.dist(pool[i]["pts"][-1], s["pts"][-1]))
        o = pool.pop(j)
        ie.append(max(max(path_dist(p, o["pts"]) for p in s["pts"]), max(path_dist(p, s["pts"]) for p in o["pts"])))
        if s["w"] is not None: iw.append(abs(s["w"] - o["w"]))
    res["ink"] = {"expected": len(src["inks"]), "matched": len(out["inks"]), "max": round(max(ie), 2) if ie else None,
                  "width": round(max(iw), 3) if iw else None}
    # Shapes: same order; max distance both ways between true outline and output polygon.
    sh = []
    for s, o in zip(src["shapes"], out["shapes"]):
        d1 = max(poly_dist(p, o) for p in s["pts"]); d2 = max(poly_dist(p, s["pts"]) for p in o)
        sh.append({"label": s["label"], "verts": len(o), "dev": round(max(d1, d2), 2)})
    res["shape"] = {"expected": len(src["shapes"]), "matched": len(out["shapes"]),
                    "max": max((r["dev"] for r in sh), default=None), "rows": sh}
    # Arrowheads: compare each true chevron with the closest output chevron.
    ae = []
    for s in src["arrows"]:
        best = min((max(math.hypot(a[0] - b[0], a[1] - b[1]) for a, b in zip(s, o)) for o in out["arrows"]), default=None)
        if best is not None: ae.append(best)
    res["arrow"] = {"expected": len(src["arrows"]), "matched": len(out["arrows"]), "max": round(max(ae), 2) if ae else None}
    fails = []
    if res["text"]["matched"] < res["text"]["expected"]: fails.append("missing text")
    if res["text"]["max"] is not None and res["text"]["max"] > TOL["text"]: fails.append(f"text off by {res['text']['max']} px")
    for k in ("sticker", "note", "shape", "arrow", "ink"):
        if res[k]["matched"] != res[k]["expected"]: fails.append(f"{k} count {res[k]['matched']}/{res[k]['expected']}")
        if res[k]["max"] is not None and res[k]["max"] > TOL[k]: fails.append(f"{k} off by {res[k]['max']} px")
    if res["note_color"]["max"] is not None and res["note_color"]["max"] > TOL["note_color"]:
        fails.append(f"note colour off by {res['note_color']['max']}/255")
    if res["ink"]["width"] is not None and res["ink"]["width"] > TOL["ink_width"]:
        fails.append(f"ink width off by {res['ink']['width']} px")
    res["fails"] = fails
    return res


# ----------------------------------------------------------------------------------------
# Report
# ----------------------------------------------------------------------------------------
def write_report(out_root: Path, results: list[dict]):
    def cell(r, k):
        v = r["checks"].get(k) if r.get("checks") else None
        if not v: return "<td>—</td>"
        if k == "text":
            return f"<td>{v['matched']}/{v['expected']} · median {v['median']} / max {v['max']}</td>"
        if k == "note_color":
            return f"<td>{v['max'] if v['max'] is not None else '—'}</td>"
        if k == "ink" and not v["expected"] and not v["matched"]:
            return "<td>—</td>"
        if k == "ink":
            return f"<td>{v['matched']}/{v['expected']} · max {v['max']} · width {v['width']}</td>"
        return f"<td>{v['matched']}/{v['expected']}" + (f" · max {v['max']}" if v['max'] is not None else "") + "</td>"
    rows = "".join(
        f"<tr class='{'bad' if r['status'] != 'PASS' else ''}'><td>{r['board']}</td><td>{r['version']}</td><td>{r['status']}</td>"
        + "".join(cell(r, k) for k in ("text", "sticker", "note", "note_color", "shape", "arrow", "ink"))
        + f"<td>{'; '.join(r.get('fails', []))}</td></tr>" for r in results)
    boards = sorted({r["board"] for r in results})
    figs = ""
    for b in boards:
        imgs = f"<figure><img src='shots/{b}_html.png'><figcaption>Whiteboard export</figcaption></figure>"
        for v in VERSIONS:
            if (out_root / "shots" / f"{b}_{v}.png").exists():
                imgs += f"<figure><img src='shots/{b}_{v}.png'><figcaption>{v} → .ubz</figcaption></figure>"
        figs += f"<h2 id='{b}'>{b}</h2><div class=grid>{imgs}</div>"
    doc = f"""<!doctype html><html><head><meta charset=utf-8><title>Visual checks</title><style>
body{{font:14px/1.45 system-ui,sans-serif;margin:24px;color:#1d1b22;background:#fbfbfd}}
table{{border-collapse:collapse;font-variant-numeric:tabular-nums}} td,th{{border:1px solid #ddd;padding:3px 8px}}
th{{background:#f1eff6}} tr.bad td{{background:#fde8e8}} .grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
figure{{margin:0;background:#f1eff6;padding:6px;border-radius:6px}} img{{width:100%;background:#fff}}
figcaption{{font-size:12px;color:#666;text-align:center}}</style></head><body>
<h1>Whiteboard → OpenBoard visual checks</h1>
<p>{time.strftime('%Y-%m-%d %H:%M')} · errors in board px · tolerances: text {TOL['text']}, sticker {TOL['sticker']}, note {TOL['note']}, shape {TOL['shape']}, arrowhead {TOL['arrow']}, ink {TOL['ink']} (width {TOL['ink_width']}); note colour {TOL['note_color']}/255.
Renders show the .ubz page as OpenBoard lays it out (approximation in Chromium).</p>
<table><tr><th>Board</th><th>Script</th><th>Status</th><th>Text</th><th>Stickers</th><th>Notes</th><th>Note colour</th><th>Shapes</th><th>Arrowheads</th><th>Ink</th><th>Problems</th></tr>{rows}</table>
{figs}</body></html>"""
    (out_root / "report.html").write_text(doc, encoding="utf-8")


# ----------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--scripts", type=Path, default=REPO, help="folder with the converter scripts")
    ap.add_argument("--samples", type=Path, default=REPO / "samples")
    ap.add_argument("--out", type=Path, default=REPO / "tests" / "out" / "visual")
    ap.add_argument("--sample", nargs="*", help="only these sample folder names")
    ap.add_argument("--version", nargs="*", choices=VERSIONS, default=list(VERSIONS))
    ap.add_argument("--powershell", help="PowerShell executable (default: powershell.exe on Windows, else pwsh)")
    a = ap.parse_args()

    ps = find_powershell(a.powershell)
    a.out.mkdir(parents=True, exist_ok=True); (a.out / "shots").mkdir(exist_ok=True)
    samples = sorted(d for d in a.samples.iterdir() if d.is_dir() and (not a.sample or d.name in a.sample))
    scripts = {v: instrument(a.scripts / f"Convert-WhiteboardHtmlToOpenBoard-{v}.ps1",
                             a.out / "_instrumented" / f"Convert-WhiteboardHtmlToOpenBoard-{v}.ps1") for v in a.version}
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        hpage = browser.new_page(viewport={"width": 1600, "height": 1000}, device_scale_factor=2)
        upage = browser.new_page()
        for d in samples:
            html = next(d.glob("*.html"), None)
            if not html: continue
            try:
                src = measure_html(hpage, html, a.out / "shots" / f"{d.name}_html.png")
            except Exception as e:
                results.append({"board": d.name, "version": "-", "status": "ERROR", "fails": [f"HTML: {e}"]}); continue
            for v, script in scripts.items():
                r = {"board": d.name, "version": v}
                ubz, log = convert(ps, script, html, a.out / "ubz" / v)
                if not ubz:
                    r.update(status="FAIL", fails=["conversion failed: " + log[-300:]]); results.append(r); continue
                try:
                    out = measure_ubz(upage, ubz, a.out / "shots" / f"{d.name}_{v}.png")
                    r["checks"] = compare(src, out, solid_only=(v == "v2_solid"))
                    r["fails"] = r["checks"].pop("fails")
                    r["status"] = "FAIL" if r["fails"] else "PASS"
                except Exception as e:
                    r.update(status="ERROR", fails=[str(e)])
                results.append(r)
                c = r.get("checks", {})
                print(f"{d.name:24} {v:12} {r['status']:5}  text max {c.get('text', {}).get('max')}  "
                      f"sticker {c.get('sticker', {}).get('max')}  note {c.get('note', {}).get('max')}  shape {c.get('shape', {}).get('max')}  "
                      f"arrow {c.get('arrow', {}).get('max')}  note colour {c.get('note_color', {}).get('max')}  "
                      f"ink {c.get('ink', {}).get('max')}  {'; '.join(r.get('fails', []))}", flush=True)
        browser.close()
    (a.out / "results.json").write_text(json.dumps(results, indent=1), encoding="utf-8")
    write_report(a.out, results)
    bad = [r for r in results if r["status"] != "PASS"]
    print(f"\n{len(results)} check(s): {len(results) - len(bad)} passed, {len(bad)} failed. Report: {a.out / 'report.html'}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
