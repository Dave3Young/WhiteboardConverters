"""
Geometric checks for the Whiteboard -> Excalidraw converters.

For every sample export and both converter scripts this:
  1. converts the export with an *instrumented temporary copy* of the converter (the only
     change is one printed line recording the scene's normalisation offset, so board
     coordinates can be recovered exactly -- the scripts in the repo are never modified);
  2. renders the Whiteboard HTML in headless Chromium and measures where Whiteboard really
     puts every text, sticker, note, image, shape outline, connector, arrowhead and ink stroke,
     and which colours it paints each note;
  3. reads the .excalidraw scene and places the same things the way Excalidraw does (text
     is laid out in Chromium with Excalidraw's font size, line height, width and alignment,
     and -- like Excalidraw -- never re-wrapped: a line wider than its element is clipped in
     Excalidraw, so it fails the check);
  4. draws the scene with Excalidraw's own exporter, for the report only;
  5. writes tests/out/visual/report.html (side-by-side pictures + accuracy tables) and
     results.json, and exits non-zero if any tolerance is exceeded.

Dependencies: Playwright (pip install playwright; python -m playwright install chromium), and
for the Excalidraw pictures @excalidraw/utils in tests/visual/node_modules (npm ci in
tests/visual). Run tests/Setup-VisualTests.ps1 once to set both up. Without @excalidraw/utils
the checks still run and the report shows only the Whiteboard pictures.

Usage (from the repo root):
    tests\\.venv\\Scripts\\python tests\\visual\\run_visual_checks.py
    ... --sample AssumptionGrid ImageBoard --version solid
"""
from __future__ import annotations

import argparse
import html as htmllib
import json
import math
import platform
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

REPO = Path(__file__).resolve().parents[2]
VERSIONS = ("solid", "gradient")

# Pass/fail tolerances in board pixels.
TOL = {"text": 5.0, "sticker": 0.5, "note": 0.5, "image": 0.5, "shape": 1.5, "connector": 0.5, "arrow": 0.5,
       "ink": 0.5, "ink_width": 0.1}
# Largest channel difference (0-255) between Whiteboard's and the scene's note colour, at five
# points of each note.
NOTE_COLOR_TOL = 3.0
ARROW_ANGLE_TOL = 2.0   # degrees between Whiteboard's chevron axis and the arrow's end segment
TEXT_CLIP_TOL = 1.0     # px a text line may extend past its element's width before Excalidraw clips it


# ----------------------------------------------------------------------------------------
# Conversion with an instrumented copy of the converter
# ----------------------------------------------------------------------------------------
OFFSET_ANCHOR = "        $scene = [ordered]@{"
OFFSET_LINE = ("        [Console]::Out.WriteLine(('TESTOFFSET {0} {1} {2}' -f "
               "$minX.ToString([Globalization.CultureInfo]::InvariantCulture), "
               "$minY.ToString([Globalization.CultureInfo]::InvariantCulture), "
               "$CanvasPadding.ToString([Globalization.CultureInfo]::InvariantCulture)))")


def instrument(script: Path, dest: Path) -> Path:
    raw = script.read_bytes().decode("utf-8-sig")
    if "TESTOFFSET" not in raw:
        nl = "\r\n" if "\r\n" in raw else "\n"
        if raw.count(OFFSET_ANCHOR) != 1:
            raise SystemExit(f"Can't instrument {script.name}: scene writer not found.")
        raw = raw.replace(OFFSET_ANCHOR, OFFSET_LINE + nl + OFFSET_ANCHOR, 1)
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


def convert(ps: list[str], script: Path, html: Path, out_dir: Path):
    out_dir.mkdir(parents=True, exist_ok=True)
    # -Command with the call operator rather than -File: with stdin redirected (CI, IDE and agent
    # shells), "powershell.exe -File" feeds stdin to the script as pipeline input, so the
    # converter's process {} block runs once per stdin line -- zero times for empty stdin.
    q = lambda s: "'" + str(s).replace("'", "''") + "'"
    cmd = ps + ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                f"& {q(script)} -InputPath {q(html)} -OutputDirectory {q(out_dir)} -Force; exit $LASTEXITCODE"]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    scene = out_dir / (html.stem + ".excalidraw")
    log = (p.stdout + p.stderr).strip()
    m = re.search(r"TESTOFFSET (\S+) (\S+) (\S+)", p.stdout)
    if p.returncode != 0 or not scene.exists() or not m:
        return None, None, log
    min_x, min_y, pad = (float(v) for v in m.groups())
    # scene = board - min + pad  ->  board = scene + (min - pad)
    return scene, (min_x - pad, min_y - pad), log


# ----------------------------------------------------------------------------------------
# Whiteboard side: measure the export in the browser
# ----------------------------------------------------------------------------------------
# A text element's first-line baseline in screen y, as an unrounded renderer would place it
# (JS function source). Text y is compared there on both sides, not at the Range box top: that
# is the content-area top, which includes CSS half-leading that Chromium rounds to whole px.
# Chromium also rounds the ascent to whole px at the font size it lays out, and Whiteboard lays
# out at 34 px and scales while the stand-in lays out at the final size, so even the drawn
# baselines differ by up to 0.5 px x scale (2.1 px on KWL's 187 px letters). So this takes the
# line top (a zero-size vertical-align:top marker) plus the drawn top-to-baseline distance
# (a zero-size baseline marker), rescaled by the unrounded distance over Chromium's rounded one,
# both read from unscaled probes of the same font (the unrounded one at 1000 px). A probe needs
# text: in quirks mode (no doctype) a line of only empty inline-blocks gets no strut.
# Returns screen [x, y]: the baseline across the line, the Range rect `rect` along it, so a
# label rotated by 90 degrees is measured across its line in x.
BASELINE = r"""(el, rect) => {
  const cs = getComputedStyle(el), F = parseFloat(cs.fontSize);
  const lh = cs.lineHeight === 'normal' ? 'normal' : String(parseFloat(cs.lineHeight) / F);
  const mark = va => { const k = document.createElement('span');
    k.style.cssText = 'display:inline-block;width:0;height:0;padding:0;border:0;margin:0;vertical-align:' + va; return k; };
  // Word joiners keep both markers on the first glyph's line: an inline-block is a break
  // opportunity, and a glyph wider than its column (KWL's "W") would wrap away from them.
  const span = host => { const t = mark('top'), b = mark('baseline');
    const j = [0, 1].map(() => document.createTextNode('\u2060'));
    host.insertBefore(j[1], host.firstChild); host.insertBefore(b, j[1]); host.insertBefore(j[0], b); host.insertBefore(t, j[0]);
    const p = t.getBoundingClientRect(), q = b.getBoundingClientRect();
    [t, b, ...j].forEach(n => n.remove()); return [p.left, p.top, q.left, q.bottom]; };
  const probe = size => { const d = document.createElement('div');
    d.style.cssText = 'position:absolute;left:0;top:0;margin:0;padding:0;border:0;white-space:nowrap;font-family:'
      + cs.fontFamily + ';font-weight:' + cs.fontWeight + ';font-style:' + cs.fontStyle + ';font-size:' + size + 'px;line-height:' + lh;
    d.textContent = 'x'; document.body.append(d); const s = span(d); d.remove(); return s[3] - s[1]; };
  const [tx, ty, bx, by] = span(el), drawn = probe(F), k = drawn > 0 ? probe(1000) / 1000 * F / drawn : 1;
  const x = tx + (bx - tx) * k, y = ty + (by - ty) * k;
  return Math.abs(by - ty) >= Math.abs(bx - tx) ? [rect.left, y] : [x, rect.top];
}"""


JS_HTML = (r"""() => {
  const baseline = BASELINE;
  const cal = [], texts = [], stickers = [], shapes = [], arrows = [], connectors = [], notes = [], images = [];
  const noteStyles = [], inks = [];
  const pt = (m, x, y) => [m.a*x + m.c*y + m.e + scrollX, m.b*x + m.d*y + m.f + scrollY];
  const box = r => [r.left + scrollX, r.top + scrollY, r.width, r.height];
  document.querySelectorAll('div.anchor[data-whiteboard-type]').forEach(a => {
    const st = a.getAttribute('style') || '', r = a.getBoundingClientRect();
    const L = /left:\s*(-?[\d.]+)px/.exec(st), T = /top:\s*(-?[\d.]+)px/.exec(st);
    if (L && T && !/transform/.test(st)) cal.push([+L[1], +T[1], r.left + scrollX, r.top + scrollY]);
    const type = a.dataset.whiteboardType;
    const spans = [...a.querySelectorAll('span[data-text="true"]')];
    // Each Draft.js block (<div data-block>) is a paragraph: one line break between blocks.
    const blocks = [...a.querySelectorAll('div[data-block="true"]')];
    const t = blocks.length ? blocks.map(b => [...b.querySelectorAll('span[data-text="true"]')].map(s => s.textContent).join('')).join('\n')
                            : spans.map(s => s.textContent).join('');
    if (t.trim()) {
      // Union of the visible characters' boxes. A range over the whole text would include the
      // spaces pre-wrap keeps at the end of a wrapped line; they hang past the column and are
      // not used for alignment, so they would shift centred or right-aligned text.
      const rg = document.createRange(); let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      spans.forEach(s => {
        const w = document.createTreeWalker(s, NodeFilter.SHOW_TEXT); let n;
        while ((n = w.nextNode())) for (let i = 0; i < n.length; i++) {
          if (/\s/.test(n.data[i])) continue;
          rg.setStart(n, i); rg.setEnd(n, i + 1);
          for (const c of rg.getClientRects()) {
            x0 = Math.min(x0, c.left); y0 = Math.min(y0, c.top); x1 = Math.max(x1, c.right); y1 = Math.max(y1, c.bottom);
          }
        }
      });
      // y is compared at the first line's baseline (unrounded font metrics, as on the
      // Excalidraw side): the glyph box top depends on the face, and Excalidraw draws every
      // text in Helvetica -- Segoe Print's 1.25 em ascent put its box 15 px above Helvetica's.
      const base = baseline(spans[0], spans[0].getBoundingClientRect())[1];
      texts.push({type, text: t, x: x0 + scrollX, y: y0 + scrollY, w: x1 - x0, h: y1 - y0, base: base + scrollY});
    }
    if (type === 'ReactionStickers') stickers.push(box(a.querySelector('img').getBoundingClientRect()));
    if (['Image', 'AzureImage', 'FluidImage'].includes(type)) images.push(box(a.querySelector('img').getBoundingClientRect()));
    if (type === 'Note') {
      const bg = a.querySelector('.textBoxBackground'), cs = getComputedStyle(bg);
      notes.push(box(bg.getBoundingClientRect()));
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
    if (type === 'Connector') {
      // The whole line, sampled every 0.5 px so samples cut an elbow's corners by no more
      // than 0.35 px (a zero-length line draws nothing).
      const line = a.querySelector('svg g path:not([transform])'), len = line ? line.getTotalLength() : 0;
      if (len > 0) {
        const n = Math.max(2, Math.ceil(len / 0.5)), m = line.getScreenCTM(), pts = [];
        for (let i = 0; i <= n; i++) { const q = line.getPointAtLength(len * i / n); pts.push(pt(m, q.x, q.y)); }
        connectors.push(pts);
      }
      a.querySelectorAll('svg g path[transform]').forEach(p => {
        const m = p.getScreenCTM(), n = p.getTotalLength();
        arrows.push([0, n / 2, n].map(l => { const q = p.getPointAtLength(l); return pt(m, q.x, q.y); }));
      });
    }
  });
  // The anchor divs themselves are 0 x 0 (their content overflows them), so take the union
  // of everything drawn inside them; comment threads are not board objects.
  const all = [...document.querySelectorAll('div.anchor[data-whiteboard-type]:not([data-whiteboard-type=CommentThread]) *')]
              .map(e => e.getBoundingClientRect()).filter(r => r.width > 0 && r.height > 0);
  const bbox = all.length ? [Math.min(...all.map(r => r.left)) + scrollX, Math.min(...all.map(r => r.top)) + scrollY,
                             Math.max(...all.map(r => r.right)) + scrollX, Math.max(...all.map(r => r.bottom)) + scrollY] : null;
  return {cal, texts, stickers, shapes, arrows, connectors, notes, images, noteStyles, inks, bbox};
}""").replace("BASELINE", BASELINE, 1)


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
    if d["bbox"]:
        x0, y0, x1, y1 = d["bbox"]; pad = 12
        page.screenshot(path=str(shot), full_page=True,
                        clip={"x": max(0, x0 - pad), "y": max(0, y0 - pad), "width": x1 - x0 + 2 * pad, "height": y1 - y0 + 2 * pad})
    to_b = lambda x, y: ((x - ox) / kx, (y - oy) / ky)
    to_box = lambda b: (*to_b(b[0], b[1]), b[2] / kx, b[3] / ky)
    return {
        "texts": [dict(t, bx=to_b(t["x"], t["y"])[0], by=to_b(t["x"], t["y"])[1], bw=t["w"] / kx, bh=t["h"] / ky,
                       bbase=to_b(t["x"], t["base"])[1])
                  for t in d["texts"]],
        "stickers": [to_box(s) for s in d["stickers"]],
        "notes": [to_box(s) for s in d["notes"]],
        "images": [to_box(s) for s in d["images"]],
        "shapes": [{"label": s["label"], "pts": [to_b(*q) for q in s["pts"]]} for s in d["shapes"]],
        "connectors": [[to_b(*q) for q in c] for c in d["connectors"]],
        "arrows": [[to_b(*q) for q in a] for a in d["arrows"]],
        "note_paints": [note_paint(s) for s in d["noteStyles"]],
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
    topmost scene element, at five points inside the note."""
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
# Excalidraw side: place the scene's elements the way Excalidraw draws them
# ----------------------------------------------------------------------------------------
# Excalidraw font families (FONT_METADATA in @excalidraw/common): CSS font stack plus the metrics
# Excalidraw lays text out with (unitsPerEm, ascender, descender). Helvetica falls back to Arial
# on Windows, the same face Whiteboard's "sans-serif" resolves to.
FONTS = {1: ("Virgil, 'Segoe UI Emoji'", 1000, 886, -374),
         2: ("Helvetica, sans-serif, 'Segoe UI Emoji'", 2048, 1577, -471),
         3: ("Cascadia, 'Cascadia Code', monospace", 2048, 1900, -480)}

# Excalidraw draws each line with fillText at textBaseline "alphabetic", the first baseline at
# y + (lineHeightPx - fontSize*(ascender - descender)/unitsPerEm) / 2 + fontSize*ascender/unitsPerEm
# (getVerticalOffset). A CSS box puts it at (lineHeightPx - (fontAscent + fontDescent)) / 2 +
# fontAscent instead, so each emulated text is shifted by the difference; the text's range box
# then sits where Excalidraw's glyphs are. Rotation is applied about the box centre, as
# Excalidraw does.
JS_TEXT = r"""(els) => {
  document.body.innerHTML = ''; document.body.style.margin = '0';
  const ctx = document.createElement('canvas').getContext('2d');
  return els.map(e => {
    const lh = e.lineHeight * e.fontSize;
    ctx.font = `${e.fontSize}px ${e.font}`;
    const m = ctx.measureText('Hg');
    const cssBase = (lh - (m.fontBoundingBoxAscent + m.fontBoundingBoxDescent)) / 2 + m.fontBoundingBoxAscent;
    const exBase = (lh - e.fontSize * (e.asc - e.desc) / e.upem) / 2 + e.fontSize * e.asc / e.upem;
    const d = document.createElement('div');
    d.style.cssText = `position:absolute;left:${e.x}px;top:${e.y}px;width:${e.width}px;font-size:${e.fontSize}px;` +
      `font-family:${e.font};line-height:${lh}px;text-align:${e.textAlign};white-space:pre;` +
      `transform-origin:center;` +
      `transform:rotate(${e.angle || 0}rad) translateY(${exBase - cssBase}px);margin:0;padding:0`;
    d.textContent = e.text; document.body.appendChild(d);
    const rg = document.createRange(); rg.selectNodeContents(d);
    const b = rg.getBoundingClientRect();
    // Excalidraw draws "text" line by line with fillText and never re-wraps it: anything
    // past the element's width is clipped.
    const widest = Math.max(...e.text.split('\n').map(l => ctx.measureText(l.trimEnd()).width));
    return {text: e.originalText, x: b.left + scrollX, y: b.top + scrollY, w: b.width, h: b.height, base: e.y + exBase,
            align: e.textAlign, angle: e.angle || 0, overflow: widest - e.width};
  });
}"""


def rotate_about(p, c, angle):
    s, co = math.sin(angle), math.cos(angle)
    dx, dy = p[0] - c[0], p[1] - c[1]
    return (c[0] + dx * co - dy * s, c[1] + dx * s + dy * co)


def outline(e, off):
    x, y, w, h, a = e["x"] + off[0], e["y"] + off[1], e["width"], e["height"], e.get("angle", 0) or 0
    c = (x + w / 2, y + h / 2)
    t = e["type"]
    if t == "rectangle":
        pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
    elif t == "diamond":
        pts = [(x + w / 2, y), (x + w, y + h / 2), (x + w / 2, y + h), (x, y + h / 2)]
    elif t == "ellipse":
        pts = [(c[0] + w / 2 * math.cos(2 * math.pi * i / 180), c[1] + h / 2 * math.sin(2 * math.pi * i / 180)) for i in range(180)]
    else:   # closed polygon "line"
        pts = [(x + p[0], y + p[1]) for p in e["points"][:-1]]
        # (x, y) is a line's first point, not its box corner: Excalidraw rotates a linear
        # element about the centre of its points' bounds.
        xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
        c = ((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2)
    return [rotate_about(p, c, a) for p in pts]


def is_band(e):
    """A gradient note band: a grouped, borderless rectangle or closed line."""
    return (e["type"] == "rectangle" or (e["type"] == "line" and e.get("polygon"))) \
        and bool(e.get("groupIds")) and e.get("strokeColor") == "transparent"


def is_shape(e):
    if e["type"] in ("ellipse", "diamond"):
        return True
    if e["type"] == "rectangle":   # notes are rounded (solid) or grouped bands (gradient)
        return not e.get("roundness") and not e.get("groupIds")
    return e["type"] == "line" and bool(e.get("polygon")) and not is_band(e)


def measure_scene(page, scene_path: Path, off) -> dict:
    scene = json.loads(scene_path.read_text(encoding="utf-8"))
    els = [e for e in scene["elements"] if not e.get("isDeleted")]
    files = scene.get("files", {})

    def bx(e):
        """Board-space bounds of an element, after Excalidraw rotates it about its centre
        (the browser's bounding box of a rotated Whiteboard object is measured the same way)."""
        x, y, w, h = e["x"] + off[0], e["y"] + off[1], e["width"], e["height"]
        a = e.get("angle", 0) or 0
        if abs(a) < 1e-9:
            return (x, y, w, h)
        c = (x + w / 2, y + h / 2)
        pts = [rotate_about(p, c, a) for p in ((x, y), (x + w, y), (x + w, y + h), (x, y + h))]
        x0 = min(p[0] for p in pts); y0 = min(p[1] for p in pts)
        return (x0, y0, max(p[0] for p in pts) - x0, max(p[1] for p in pts) - y0)

    texts_in = []
    for e in els:
        if e["type"] != "text" or not e.get("text", "").strip():
            continue
        font, upem, asc, desc = FONTS.get(e.get("fontFamily"), FONTS[1])
        texts_in.append(dict(e, x=e["x"] + off[0], y=e["y"] + off[1], font=font, upem=upem, asc=asc, desc=desc))
    texts = page.evaluate(JS_TEXT, texts_in)

    images = [e for e in els if e["type"] == "image"]
    svg_img = lambda e: files.get(e["fileId"], {}).get("mimeType") == "image/svg+xml"

    # Notes: a rounded rectangle (solid) or the union of one group's bands (gradient; each band
    # but the last overlaps the next, inside the note). Each keeps its polygons and colours.
    notes, note_polys, groups = [], [], {}
    for e in els:
        if is_band(e):
            groups.setdefault(e["groupIds"][0], []).append((outline(e, off), hex_rgb(e["backgroundColor"])))
        elif e["type"] == "rectangle" and e.get("roundness"):
            notes.append(bx(e)); note_polys.append([(outline(e, off), hex_rgb(e["backgroundColor"]))])
    for g in groups.values():
        pts = [p for poly, _ in g for p in poly]
        x0 = min(p[0] for p in pts); y0 = min(p[1] for p in pts)
        notes.append((x0, y0, max(p[0] for p in pts) - x0, max(p[1] for p in pts) - y0)); note_polys.append(g)

    # Drawn underlines and ink strokes are lines too, marked with customData; they aren't
    # connectors.
    custom = lambda e, k: bool((e.get("customData") or {}).get(k))
    linear = [e for e in els if e["type"] in ("line", "arrow") and not e.get("polygon")
              and not custom(e, "underline") and not custom(e, "ink")]
    inks = [{"pts": [(e["x"] + off[0] + p[0], e["y"] + off[1] + p[1]) for p in e["points"]], "w": e["strokeWidth"]}
            for e in els if custom(e, "ink")]
    conns, arrows = [], []
    for e in linear:
        pts = [(e["x"] + off[0] + p[0], e["y"] + off[1] + p[1]) for p in e["points"]]
        conns.append(pts)
        if e["type"] == "arrow":
            if e.get("startArrowhead"): arrows.append((pts[0], pts[1], e["startArrowhead"]))
            if e.get("endArrowhead"): arrows.append((pts[-1], pts[-2], e["endArrowhead"]))
    return {
        "texts": texts,
        "stickers": [bx(e) for e in images if svg_img(e)],
        "images": [bx(e) for e in images if not svg_img(e)],
        "notes": notes,
        "note_polys": note_polys,
        "shapes": [outline(e, off) for e in els if is_shape(e)],
        "connectors": conns,
        "arrows": arrows,
        "inks": inks,
    }


# ----------------------------------------------------------------------------------------
# Excalidraw side: a picture of the scene for the report
# ----------------------------------------------------------------------------------------
# Excalidraw's own exporter (exportToSvg from @excalidraw/utils, installed into
# tests/visual/node_modules by Setup-VisualTests.ps1) draws the scene, so the picture is what
# Excalidraw shows. The bundle is an ES module, which a file:// page can't import, so the page
# and bundle are served from a made-up origin through Playwright's request routing.
EXCALIDRAW_BUNDLE = Path(__file__).resolve().parent / "node_modules" / "@excalidraw" / "utils" / "dist" / "prod"
RENDER_ORIGIN = "http://excalidraw.test/"
RENDER_PAGE = """<!doctype html><html><head><meta charset=utf-8>
<style>html,body{margin:0;background:#fff}</style></head><body>
<script type=module>
import { exportToSvg } from './index.js';
window.render = async (scene) => {
  const svg = await exportToSvg({
    data: { elements: scene.elements.filter(e => !e.isDeleted), files: scene.files || {},
            appState: { ...(scene.appState || {}), exportBackground: true } },
    config: { canvasBackgroundColor: '#ffffff', padding: 12 } });
  document.body.innerHTML = ''; document.body.appendChild(svg);
  const r = svg.getBoundingClientRect(); return [r.width, r.height];
};
window.ready = true;
</script></body></html>"""


def open_renderer(browser):
    """A page that can draw .excalidraw scenes, or None when @excalidraw/utils isn't installed."""
    if not (EXCALIDRAW_BUNDLE / "index.js").is_file():
        return None
    def serve(route):
        rel = route.request.url[len(RENDER_ORIGIN):].split("?")[0].replace("%20", " ")
        if rel in ("", "render.html"):
            return route.fulfill(body=RENDER_PAGE, content_type="text/html")
        f = EXCALIDRAW_BUNDLE / rel
        if f.is_file():
            return route.fulfill(path=str(f))
        route.fulfill(status=404, body="")
    page = browser.new_page()
    page.route(RENDER_ORIGIN + "**", serve)
    page.goto(RENDER_ORIGIN + "render.html")
    page.wait_for_function("window.ready === true", timeout=60000)
    return page


def render_scene(page, scene_path: Path, shot: Path):
    scene = json.loads(scene_path.read_text(encoding="utf-8"))
    w, h = page.evaluate("s => window.render(s)", scene)
    page.set_viewport_size({"width": max(200, math.ceil(w)), "height": max(200, math.ceil(h))})
    page.wait_for_timeout(300)
    page.screenshot(path=str(shot), full_page=True)


# ----------------------------------------------------------------------------------------
# Comparisons
# ----------------------------------------------------------------------------------------
def seg_dist(p, a, b):
    abx, aby = b[0] - a[0], b[1] - a[1]; L = abx * abx + aby * aby
    t = 0 if L == 0 else max(0, min(1, ((p[0] - a[0]) * abx + (p[1] - a[1]) * aby) / L))
    return math.hypot(p[0] - a[0] - t * abx, p[1] - a[1] - t * aby)


def poly_dist(p, poly):
    return min(seg_dist(p, poly[i], poly[(i + 1) % len(poly)]) for i in range(len(poly)))


def box_err(s, o):
    """Largest edge error between two (x, y, w, h) boxes."""
    return max(abs(s[0] - o[0]), abs(s[1] - o[1]), abs(s[0] + s[2] - o[0] - o[2]), abs(s[1] + s[3] - o[1] - o[3]))


def nearest_boxes(src, out):
    """Pairs every source box with the closest unused output box (by centre)."""
    pool = list(out); errs = []
    for s in src:
        if not pool: break
        c = (s[0] + s[2] / 2, s[1] + s[3] / 2)
        j = min(range(len(pool)), key=lambda i: math.hypot(pool[i][0] + pool[i][2] / 2 - c[0], pool[i][1] + pool[i][3] / 2 - c[1]))
        errs.append(box_err(s, pool.pop(j)))
    return errs


def path_dist(p, pts):
    """Distance from p to an open polyline."""
    return min(seg_dist(p, pts[i], pts[i + 1]) for i in range(len(pts) - 1)) if len(pts) > 1 else math.dist(p, pts[0])


def compare(src: dict, out: dict, solid_only: bool) -> dict:
    res = {}
    # Texts: match by exact text, in order. Excalidraw drops bold and the font face, so line
    # widths differ; x is compared at the alignment anchor (left edge, centre or right edge),
    # y at the first line's baseline. Rotated text is compared by the centre of its bounds.
    pool = list(out["texts"]); rows = []
    for s in src["texts"]:
        j = next((i for i, o in enumerate(pool) if o["text"] == s["text"]), None)
        if j is None:
            rows.append({"text": s["text"][:50], "missing": True}); continue
        o = pool.pop(j)
        if abs(o["angle"]) > 1e-6:
            dx = (o["x"] + o["w"] / 2) - (s["bx"] + s["bw"] / 2); dy = (o["y"] + o["h"] / 2) - (s["by"] + s["bh"] / 2)
        else:
            k = {"left": 0.0, "center": 0.5, "right": 1.0}.get(o["align"], 0.0)
            dx = (o["x"] + k * o["w"]) - (s["bx"] + k * s["bw"]); dy = o["base"] - s["bbase"]
        # Bottom-edge difference: shows line-spacing drift on multi-line text (informational --
        # wrapping can legitimately differ because bold and the font face are dropped).
        rows.append({"text": s["text"][:50], "dx": round(dx, 1), "dy": round(dy, 1), "err": round(math.hypot(dx, dy), 1),
                     "dbottom": round((o["y"] + o["h"]) - (s["by"] + s["bh"]), 1)})
    errs = [r["err"] for r in rows if "err" in r]
    res["text"] = {"expected": len(src["texts"]), "matched": len(errs),
                   "median": round(sorted(errs)[len(errs) // 2], 1) if errs else None,
                   "max": max(errs) if errs else None, "rows": rows,
                   "clipped": [{"text": t["text"][:50], "overflow": round(t["overflow"], 1)}
                               for t in out["texts"] if t["overflow"] > TEXT_CLIP_TOL]}
    for k in ("sticker", "image"):
        e = nearest_boxes(src[k + "s"], out[k + "s"])
        res[k] = {"expected": len(src[k + "s"]), "matched": len(out[k + "s"]), "max": round(max(e), 2) if e else None}
    # Notes: nearest output note by centre; box error, and its colour at five points (solid
    # fill, or the gradient at the gradient's angle).
    ne, nc, pool = [], [], list(zip(out["notes"], out["note_polys"]))
    for s, paint in zip(src["notes"], src["note_paints"]):
        if not pool: break
        c = (s[0] + s[2] / 2, s[1] + s[3] / 2)
        j = min(range(len(pool)), key=lambda i: math.hypot(pool[i][0][0] + pool[i][0][2] / 2 - c[0], pool[i][0][1] + pool[i][0][3] / 2 - c[1]))
        o, polys = pool.pop(j)
        ne.append(box_err(s, o)); nc.append(note_color_err(s, paint, solid_only, polys))
    res["note"] = {"expected": len(src["notes"]), "matched": len(out["notes"]), "max": round(max(ne), 2) if ne else None,
                   "color": max(nc) if nc else None}
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
    # Shapes: nearest output outline; max distance both ways between true and output outline.
    sh = []; pool = list(out["shapes"])
    for s in src["shapes"]:
        if not pool: break
        cx = sum(p[0] for p in s["pts"]) / len(s["pts"]); cy = sum(p[1] for p in s["pts"]) / len(s["pts"])
        j = min(range(len(pool)), key=lambda i: math.hypot(sum(p[0] for p in pool[i]) / len(pool[i]) - cx,
                                                            sum(p[1] for p in pool[i]) / len(pool[i]) - cy))
        o = pool.pop(j)
        d1 = max(poly_dist(p, o) for p in s["pts"]); d2 = max(poly_dist(p, s["pts"]) for p in o)
        # Whiteboard draws an "oval" as 8 quadratic curves, which bulge up to ~0.33% of the
        # radius off a true ellipse; the converter draws a real (editable) Excalidraw ellipse
        # through the same extremes, so ovals may deviate by that much on top of the tolerance.
        allow = TOL["shape"]
        if s["label"].strip().lower() in ("oval", "ellipse", "circle"):
            r = max(max(p[0] for p in s["pts"]) - min(p[0] for p in s["pts"]),
                    max(p[1] for p in s["pts"]) - min(p[1] for p in s["pts"])) / 2
            allow = max(allow, 0.0035 * r)
        sh.append({"label": s["label"], "dev": round(max(d1, d2), 2), "allowed": round(allow, 2)})
    res["shape"] = {"expected": len(src["shapes"]), "matched": len(out["shapes"]),
                    "max": max((r["dev"] for r in sh), default=None),
                    "over": sum(1 for r in sh if r["dev"] > r["allowed"]), "rows": sh}
    # Connectors: each true line against the output line nearest its ends; largest distance
    # either way between the two, so every bend of an elbow connector counts.
    ce, pool = [], list(out["connectors"])
    for s in src["connectors"]:
        if not pool: break
        j = min(range(len(pool)), key=lambda i: min(math.dist(pool[i][0], s[0]) + math.dist(pool[i][-1], s[-1]),
                                                   math.dist(pool[i][0], s[-1]) + math.dist(pool[i][-1], s[0])))
        o = pool.pop(j)
        ce.append(max(max(path_dist(p, o) for p in s), max(path_dist(p, s) for p in o)))
    res["connector"] = {"expected": len(src["connectors"]), "matched": len(out["connectors"]), "max": round(max(ce), 2) if ce else None}
    # Arrowheads: the chevron's tip must be an arrowhead end, pointing the same way.
    ae, angle_bad = [], 0
    for s in src["arrows"]:
        tip = s[1]; base = ((s[0][0] + s[2][0]) / 2, (s[0][1] + s[2][1]) / 2)
        cands = [(math.hypot(o[0][0] - tip[0], o[0][1] - tip[1]), o) for o in out["arrows"]]
        if not cands: continue
        d, o = min(cands, key=lambda c: c[0]); ae.append(d)
        want = math.atan2(tip[1] - base[1], tip[0] - base[0]); got = math.atan2(o[0][1] - o[1][1], o[0][0] - o[1][0])
        if abs((math.degrees(want - got) + 180) % 360 - 180) > ARROW_ANGLE_TOL: angle_bad += 1
    res["arrow"] = {"expected": len(src["arrows"]), "matched": len(out["arrows"]), "max": round(max(ae), 2) if ae else None,
                    "wrong_direction": angle_bad}
    fails = []
    if res["text"]["matched"] < res["text"]["expected"]: fails.append("missing text")
    if res["text"]["max"] is not None and res["text"]["max"] > TOL["text"]: fails.append(f"text off by {res['text']['max']} px")
    if res["text"]["clipped"]: fails.append(f"{len(res['text']['clipped'])} text(s) wider than their element (clipped in Excalidraw)")
    for k in ("sticker", "note", "image", "shape", "connector", "arrow", "ink"):
        if res[k]["matched"] != res[k]["expected"]: fails.append(f"{k} count {res[k]['matched']}/{res[k]['expected']}")
        if k == "shape":
            if res[k]["over"]: fails.append(f"{res[k]['over']} shape(s) off by up to {res[k]['max']} px")
        elif res[k]["max"] is not None and res[k]["max"] > TOL[k]: fails.append(f"{k} off by {res[k]['max']} px")
    if angle_bad: fails.append(f"{angle_bad} arrowhead(s) point the wrong way")
    if res["note"]["color"] is not None and res["note"]["color"] > NOTE_COLOR_TOL:
        fails.append(f"note colour off by {res['note']['color']}/255")
    if res["ink"]["width"] is not None and res["ink"]["width"] > TOL["ink_width"]:
        fails.append(f"ink width off by {res['ink']['width']} px")
    res["fails"] = fails
    return res


# ----------------------------------------------------------------------------------------
# Report
# ----------------------------------------------------------------------------------------
def write_report(out_root: Path, results: list[dict], rendered: bool):
    def cell(r, k):
        v = r["checks"].get(k) if r.get("checks") else None
        if not v: return "<td>—</td>"
        if k == "text":
            return f"<td>{v['matched']}/{v['expected']} · median {v['median']} / max {v['max']}</td>"
        if k == "ink" and not v["expected"] and not v["matched"]:
            return "<td>—</td>"
        extra = ""
        if k == "note" and v.get("color") is not None: extra = f" · colour {v['color']}"
        if k == "ink": extra = f" · width {v['width']}"
        return f"<td>{v['matched']}/{v['expected']}" + (f" · max {v['max']}" if v['max'] is not None else "") + extra + "</td>"
    cols = ("text", "sticker", "note", "image", "shape", "connector", "arrow", "ink")
    rows = "".join(
        f"<tr class='{'bad' if r['status'] != 'PASS' else ''}'><td><a href='#{r['board']}'>{r['board']}</a></td>"
        f"<td>{r['version']}</td><td>{r['status']}</td>"
        + "".join(cell(r, k) for k in cols)
        + f"<td>{htmllib.escape('; '.join(r.get('fails', [])))}</td></tr>" for r in results)
    figs = ""
    for b in sorted({r["board"] for r in results}):
        imgs = f"<figure><img src='shots/{b}_html.png'><figcaption>Whiteboard export</figcaption></figure>"
        for v in VERSIONS:
            if (out_root / "shots" / f"{b}_{v}.png").exists():
                imgs += f"<figure><img src='shots/{b}_{v}.png'><figcaption>{v} → .excalidraw</figcaption></figure>"
        figs += f"<h2 id='{b}'>{b}</h2><div class=grid>{imgs}</div>"
    note = ("Renders are drawn by Excalidraw's own exporter (@excalidraw/utils exportToSvg)." if rendered else
            "Excalidraw renders are missing: run tests\\Setup-VisualTests.ps1 to install @excalidraw/utils.")
    doc = f"""<!doctype html><html><head><meta charset=utf-8><title>Visual checks</title><style>
body{{font:14px/1.45 system-ui,sans-serif;margin:24px;color:#1d1b22;background:#fbfbfd}}
table{{border-collapse:collapse;font-variant-numeric:tabular-nums}} td,th{{border:1px solid #ddd;padding:3px 8px}}
th{{background:#f1eff6}} tr.bad td{{background:#fde8e8}} .grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
figure{{margin:0;background:#f1eff6;padding:6px;border-radius:6px}} img{{width:100%;background:#fff}}
figcaption{{font-size:12px;color:#666;text-align:center}}</style></head><body>
<h1>Whiteboard → Excalidraw visual checks</h1>
<p>{time.strftime('%Y-%m-%d %H:%M')} · errors in board px · tolerances: {', '.join(f'{k} {v}' for k, v in TOL.items())}; note colour {NOTE_COLOR_TOL}/255.
{note}</p>
<table><tr><th>Board</th><th>Script</th><th>Status</th><th>Text</th><th>Stickers</th><th>Notes</th><th>Images</th><th>Shapes</th><th>Connectors</th><th>Arrowheads</th><th>Ink</th><th>Problems</th></tr>{rows}</table>
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
    scripts = {v: instrument(a.scripts / f"Convert-WhiteboardHtmlToExcalidraw-{v}.ps1",
                             a.out / "_instrumented" / f"Convert-WhiteboardHtmlToExcalidraw-{v}.ps1") for v in a.version}
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        hpage = browser.new_page(viewport={"width": 1600, "height": 1000}, device_scale_factor=2)
        tpage = browser.new_page()
        rpage = open_renderer(browser)
        if not rpage:
            print("@excalidraw/utils is not installed (run tests\\Setup-VisualTests.ps1): "
                  "the report will have no Excalidraw renders.", flush=True)
        for d in samples:
            html = next(d.glob("*.html"), None)
            if not html: continue
            try:
                src = measure_html(hpage, html, a.out / "shots" / f"{d.name}_html.png")
            except Exception as e:
                results.append({"board": d.name, "version": "-", "status": "ERROR", "fails": [f"HTML: {e}"]}); continue
            for v, script in scripts.items():
                r = {"board": d.name, "version": v}
                scene, off, log = convert(ps, script, html, a.out / "scenes" / v)
                if not scene:
                    r.update(status="FAIL", fails=["conversion failed: " + log[-300:].replace("\n", " ")])
                else:
                    try:
                        r["checks"] = compare(src, measure_scene(tpage, scene, off), solid_only=(v == "solid"))
                        r["fails"] = r["checks"].pop("fails")
                        r["status"] = "FAIL" if r["fails"] else "PASS"
                    except Exception as e:
                        r.update(status="ERROR", fails=[repr(e)])
                    if rpage:
                        # The picture is for people; a failed render doesn't fail the geometry check.
                        try:
                            render_scene(rpage, scene, a.out / "shots" / f"{d.name}_{v}.png")
                        except Exception as e:
                            print(f"{d.name} {v}: render failed: {e}", flush=True)
                results.append(r)
                c = r.get("checks", {})
                m = lambda k: c.get(k, {}).get("max")
                print(f"{d.name:24} {v:9} {r['status']:5}  text {m('text')}  sticker {m('sticker')}  note {m('note')}  "
                      f"image {m('image')}  shape {m('shape')}  connector {m('connector')}  arrow {m('arrow')}  "
                      f"note colour {c.get('note', {}).get('color')}  ink {m('ink')}  "
                      f"{'; '.join(r.get('fails', []))}", flush=True)
        browser.close()
    (a.out / "results.json").write_text(json.dumps(results, indent=1), encoding="utf-8")
    write_report(a.out, results, rpage is not None)
    bad = [r for r in results if r["status"] != "PASS"]
    print(f"\n{len(results)} check(s): {len(results) - len(bad)} passed, {len(bad)} failed. "
          f"Tolerances (board px): {TOL}. Report: {a.out / 'report.html'}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
