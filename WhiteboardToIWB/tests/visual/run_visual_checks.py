"""
Visual / geometric checks for the Whiteboard -> IWB converter.

For every sample export and both note fills (Solid, Gradient) this:
  1. converts the export with an *instrumented temporary copy* of the converter (the only
     change is one XML comment recording the board-to-page mapping, so board coordinates can
     be recovered exactly -- the script in the repo is never modified);
  2. renders the Whiteboard HTML in headless Chromium and measures where Whiteboard really
     puts every text, sticker/image, note, shape outline, connector arrowhead and ink stroke,
     and which colours it paints each note;
  3. renders the IWB page as a standard SVG renderer shows it (textarea as a wrapped HTML
     box) and measures the same things;
  4. writes tests/out/visual/report.html (side-by-side screenshots + accuracy tables) and
     results.json, and exits non-zero if any tolerance is exceeded.

Errors are reported in board px (page px divided by the page scale).

Only dependency: Playwright (pip install playwright; python -m playwright install chromium).
Run tests/Setup-VisualTests.ps1 once to set that up in tests/.venv.

Usage (from the repo root):
    tests\\.venv\\Scripts\\python tests\\visual\\run_visual_checks.py
    ... --sample AssumptionGrid ImageBoard --fill Solid
"""
from __future__ import annotations

import argparse
import base64
import html as htmllib
import json
import math
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
IWB = "{http://www.imsglobal.org/xsd/iwb_v1p0}"
SVG = "{http://www.w3.org/2000/svg}"
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"
FILLS = ("Solid", "Gradient")
SCRIPT = "Convert-WhiteboardHtmlToIwb.ps1"

# Pass/fail tolerances in board pixels.
# image_content: mean colour difference (0-255) between the two images' 8 x 8 thumbnails.
TOL = {"text": 5.0, "image": 0.5, "note": 0.5, "shape": 1.5, "connector": 0.5, "arrow": 0.5, "image_content": 3.0,
       "note_color": 3.0, "ink": 0.5, "ink_width": 0.1}
# note_color: largest channel difference (0-255) at five points of each note; ink_width in board px.

# An image's 8 x 8 thumbnail drawn on white, as 192 RGB values (JS function source).
# Drawn at 256 x 256 first, then halved five times: each halving averages 2 x 2 pixels, so
# the thumbnail is a true area average. A direct 8 x 8 drawImage samples only a few source
# pixels, which made a 256 px PNG sticker look lighter at its edges than the 64 px SVG it
# was rasterised from.
THUMB = r"""async (src) => { const im = new Image(); im.src = src; await im.decode();
  let n = 256; let c = document.createElement('canvas'); c.width = c.height = n; let g = c.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, n, n); g.drawImage(im, 0, 0, n, n);
  while (n > 8) { n /= 2; const h = document.createElement('canvas'); h.width = h.height = n;
    const hg = h.getContext('2d'); hg.imageSmoothingQuality = 'high'; hg.drawImage(c, 0, 0, n, n); c = h; g = hg; }
  return [...g.getImageData(0, 0, 8, 8).data].filter((v, i) => i % 4 !== 3); }"""


def img_src(z: zipfile.ZipFile, el) -> str:
    href = el.get(XLINK_HREF); ext = href.rsplit(".", 1)[-1].lower()
    mime = {"jpg": "jpeg", "svg": "svg+xml"}.get(ext, ext)
    return f"data:image/{mime};base64,{base64.b64encode(z.read(href)).decode()}"


# ----------------------------------------------------------------------------------------
# Conversion with an instrumented copy of the converter
# ----------------------------------------------------------------------------------------
MAP_ANCHOR = "        [void]$sb.Append('    <svg:pageset>' + \"`r`n\")"
MAP_LINE = ("        [void]$sb.Append(('<!-- TESTMAP {0} {1} {2} -->' -f $s.ToString('R', $ci), "
            "$ox.ToString('R', $ci), $oy.ToString('R', $ci)) + \"`r`n\")")


def instrument(script: Path, dest: Path) -> Path:
    raw = script.read_bytes().decode("utf-8-sig")
    if "TESTMAP" not in raw:
        nl = "\r\n" if "\r\n" in raw else "\n"
        if raw.count(MAP_ANCHOR) != 1:
            raise SystemExit(f"Can't instrument {script.name}: content.xml writer not found.")
        raw = raw.replace(MAP_ANCHOR, MAP_LINE + nl + MAP_ANCHOR, 1)
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


def convert(ps: list[str], script: Path, html: Path, out_dir: Path, fill: str) -> tuple[Path | None, str]:
    out_dir.mkdir(parents=True, exist_ok=True)
    # -Command with the call operator rather than -File: with stdin redirected (CI, IDE and agent
    # shells), "powershell.exe -File" feeds stdin to the script as pipeline input, so the
    # converter's process {} block runs once per stdin line -- zero times for empty stdin.
    q = lambda s: "'" + str(s).replace("'", "''") + "'"
    cmd = ps + ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                f"& {q(script)} -InputPath {q(html)} -OutputDirectory {q(out_dir)} -NoteFill {fill} -Force; exit $LASTEXITCODE"]
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    iwb = out_dir / (html.stem + ".iwb")
    log = (p.stdout + p.stderr).strip()
    return (iwb if p.returncode == 0 and iwb.exists() else None), log


# ----------------------------------------------------------------------------------------
# Whiteboard side: measure the export in the browser
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


# ----------------------------------------------------------------------------------------
JS_HTML = (r"""async () => {
  const cal = [], texts = [], images = [], shapes = [], arrows = [], notes = [], noteStyles = [], inks = [], connectors = [];
  const thumb = THUMB;
  const baseline = BASELINE;
  const pt = (m, x, y) => [m.a*x + m.c*y + m.e + scrollX, m.b*x + m.d*y + m.f + scrollY];
  for (const a of document.querySelectorAll('div.anchor[data-whiteboard-type]')) {
    const st = a.getAttribute('style') || '', r = a.getBoundingClientRect();
    const L = /left:\s*(-?[\d.]+)px/.exec(st), T = /top:\s*(-?[\d.]+)px/.exec(st);
    if (L && T && !/transform/.test(st)) cal.push([+L[1], +T[1], r.left + scrollX, r.top + scrollY]);
    const type = a.dataset.whiteboardType;
    const spans = [...a.querySelectorAll('span[data-text="true"]')];
    // Each Draft.js block (<div data-block>) is a paragraph: one line break between blocks.
    const blocks = [...a.querySelectorAll('div[data-block="true"]')];
    let t = blocks.length ? blocks.map(b => [...b.querySelectorAll('span[data-text="true"]')].map(s => s.textContent).join('')).join('\n')
                          : spans.map(s => s.textContent).join('');
    // A shape's label is clipped by its text box (overflow-y: hidden), so compare the text that
    // shows: cut at the first character (or empty paragraph) whose line is less than half inside
    // the box, as Chromium lays it out, with trailing white space trimmed.
    const tb = type === 'Shape' ? a.querySelector('.textbox.shapeText') : null;
    if (tb && blocks.length && tb.scrollHeight > tb.clientHeight + 1) {
      const bottom = tb.getBoundingClientRect().bottom, rg = document.createRange(), out = r => (r.top + r.bottom) / 2 > bottom;
      let cut = null, off = 0;
      for (const b of blocks) {
        const nodes = [...b.querySelectorAll('span[data-text="true"]')].map(s => s.firstChild).filter(n => n && n.nodeType === 3);
        if (!nodes.length && out(b.getBoundingClientRect())) cut = off;
        for (const n of nodes) {
          for (let i = 0; i < n.length && cut === null; i++) {
            rg.setStart(n, i); rg.setEnd(n, i + 1); const r = rg.getBoundingClientRect();
            if (r.height && out(r)) cut = off + i;
          }
          if (cut !== null) break;
          off += n.length;
        }
        if (cut !== null) break;
        off += 1;   // the line break between blocks
      }
      if (cut !== null) t = t.slice(0, cut).trimEnd();
    }
    if (t.trim()) {
      // The first paragraph only: a Range across several Draft.js blocks also spans the blocks'
      // own boxes, so it is the column's full width (the output side stops at the first break).
      const para = blocks.length ? [...blocks[0].querySelectorAll('span[data-text="true"]')] : [];
      const rg = document.createRange(), last = para.length ? para[para.length - 1] : spans[spans.length - 1], lf = last.firstChild || last;
      rg.setStart(spans[0].firstChild || spans[0], 0); rg.setEnd(lf, lf.length || 0);
      const b = rg.getBoundingClientRect(), [x, y] = baseline(spans[0], b);
      texts.push({type, text: t.replace(/\r\n?/g, '\n'), x: x + scrollX, y: y + scrollY, w: b.width, h: b.height});
    }
    // An image's top-left, top-right and bottom-left corners, read from zero-size markers
    // laid on its corners, so rotation and scale come from the browser's own transform.
    if (['ReactionStickers', 'Image', 'AzureImage', 'FluidImage'].includes(type)) {
      const im = a.querySelector('img');
      if (im) {
        // The used size, unrounded (offsetWidth/offsetHeight are whole px: 578 for 577.535,
        // 2 px at SailboatRetrospective's 4.38x scale).
        const host = im.parentElement; host.style.position = 'relative';
        const W = parseFloat(getComputedStyle(im).width), H = parseFloat(getComputedStyle(im).height);
        const corners = [[0, 0], [1, 0], [0, 1]].map(([u, v]) => {
          const k = document.createElement('div');
          k.style.cssText = 'position:absolute;width:0;height:0;left:' + (im.offsetLeft + u * W) + 'px;top:' + (im.offsetTop + v * H) + 'px';
          host.append(k); const q = k.getBoundingClientRect(); k.remove();
          return [q.left + scrollX, q.top + scrollY];
        });
        images.push([corners, await thumb(im.src)]);
      }
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
    // The connector's line, sampled every 0.5 px so samples cut an elbow's corners by no more
    // than 0.35 px (a zero-length line draws nothing).
    if (type === 'Connector') {
      const p = a.querySelector('svg g > path:not([transform])');
      const len = p ? p.getTotalLength() : 0;
      if (len > 0) {
        const n = Math.max(2, Math.ceil(len / 0.5)), m = p.getScreenCTM(), pts = [];
        for (let i = 0; i <= n; i++) { const q = p.getPointAtLength(len * i / n); pts.push(pt(m, q.x, q.y)); }
        connectors.push(pts);
      }
    }
  }
  // The anchor divs themselves are 0 x 0 (their content overflows them), so take the union
  // of everything drawn inside them; comment threads are not board objects.
  const all = [...document.querySelectorAll('div.anchor[data-whiteboard-type]:not([data-whiteboard-type=CommentThread]) *')]
              .map(e => e.getBoundingClientRect()).filter(r => r.width > 0 && r.height > 0);
  const bbox = all.length ? [Math.min(...all.map(r => r.left)) + scrollX, Math.min(...all.map(r => r.top)) + scrollY,
                             Math.max(...all.map(r => r.right)) + scrollX, Math.max(...all.map(r => r.bottom)) + scrollY] : null;
  return {cal, texts, images, shapes, arrows, notes, noteStyles, inks, connectors, bbox};
}""").replace("THUMB", THUMB, 1).replace("BASELINE", BASELINE, 1)


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
        "texts": [dict(t, bx=to_b(t["x"], t["y"])[0], by=to_b(t["x"], t["y"])[1]) for t in d["texts"]],
        "images": [[to_b(*q) for q in s[0]] for s in d["images"]],
        "thumbs": [s[1] for s in d["images"]],
        "connectors": [[to_b(*q) for q in c] for c in d["connectors"]],
        "notes": [(*to_b(s[0], s[1]), s[2] / kx, s[3] / ky) for s in d["notes"]],
        "shapes": [{"label": s["label"], "pts": [to_b(*q) for q in s["pts"]]} for s in d["shapes"]],
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
    topmost output rect or polygon, at five points inside the note."""
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
# IWB side: read + render the page
# ----------------------------------------------------------------------------------------
def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def iwb_page(iwb: Path):
    z = zipfile.ZipFile(iwb)
    raw = z.read("content.xml").decode("utf-8")
    m = re.search(r"TESTMAP (\S+) (\S+) (\S+)", raw)
    mapping = tuple(float(v) for v in m.groups()) if m else None
    root = ET.fromstring(raw.encode("utf-8"))
    svg = root.find(SVG + "svg")
    w, h = float(svg.get("width")), float(svg.get("height"))
    page = svg.find(f"{SVG}pageset/{SVG}page")
    groups = {}
    for gi, g in enumerate(root.findall(IWB + "group")):
        for e in g.findall(IWB + "element"):
            groups[e.get("ref")] = gi
    return z, root, page, (w, h), mapping, groups


def textarea_lines(el) -> list[str]:
    lines = [el.text or ""]
    for c in el:
        if local(c.tag) == "tbreak":
            lines.append("")
        elif local(c.tag) == "tspan":
            lines[-1] += c.text or ""
        lines[-1] += c.tail or ""
    return lines


def textarea_html(el) -> str:
    """A textarea's content as HTML: <br/> for tbreak, a styled <span> for each tspan."""
    out = [htmllib.escape(el.text or "")]
    for c in el:
        if local(c.tag) == "tbreak":
            out.append("<br/>")
        elif local(c.tag) == "tspan":
            css = "".join(f"{k}:{CSS_WEIGHT.get(c.get(k), c.get(k))};" for k in ("font-weight", "font-style", "text-decoration") if c.get(k))
            out.append(f'<span style="{css}">{htmllib.escape(c.text or "")}</span>')
        out.append(htmllib.escape(c.tail or ""))
    return "".join(out)


# OpenBoard's weight names as CSS weights (demibold isn't a CSS keyword).
CSS_WEIGHT = {"light": "300", "demibold": "600", "black": "900"}


def browser_svg(z: zipfile.ZipFile, page, size, scale: float) -> str:
    """The page as plain SVG: images inlined, each textarea as a wrapped HTML box.

    Textareas are laid out at board scale (size / scale, then drawn through scale(scale)), as a
    viewer zoomed to the board would lay them out. At page scale the fonts can be 3-4 px, and
    Chromium rounds a font's ascent to whole px, so the measured baseline would carry up to
    0.5 page px = 0.5 / scale board px (5 px at scale 0.1) of the stand-in renderer's rounding.
    """
    w, h = size
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{w:g}" height="{h:g}" viewBox="0 0 {w:g} {h:g}">',
           f'<rect width="{w:g}" height="{h:g}" fill="#fff"/>']
    attrs = lambda el, skip=(): " ".join(f'{k}="{htmllib.escape(v)}"' for k, v in el.attrib.items()
                                         if not k.startswith("{") and k not in skip)
    for el in page:
        tag = local(el.tag)
        if tag in ("rect", "polygon", "polyline"):
            out.append(f"<{tag} {attrs(el)}/>")
        elif tag == "image":
            href = el.get(XLINK_HREF); ext = href.rsplit(".", 1)[-1].lower()
            mime = {"jpg": "jpeg", "svg": "svg+xml"}.get(ext, ext)
            data = base64.b64encode(z.read(href)).decode()
            out.append(f'<image {attrs(el)} href="data:image/{mime};base64,{data}" preserveAspectRatio="none"/>')
        elif tag == "textarea":
            align = {"center": "center", "end": "right"}.get(el.get("text-align"), "left")
            fs = re.fullmatch(r"([\d.]+)(\D*)", el.get("font-size"))
            style = (f"font-family:'{el.get('font-family')}';font-size:{float(fs[1]) / scale:g}{fs[2]};font-weight:{CSS_WEIGHT.get(el.get('font-weight'), el.get('font-weight'))};"
                     f"font-style:{el.get('font-style', 'normal')};"
                     f"color:{el.get('fill')};text-align:{align};white-space:pre-wrap;overflow-wrap:break-word;"
                     "line-height:normal;margin:0;width:100%;overflow:visible")
            body = textarea_html(el)
            # line-increment spaces the baselines and (like a renderer laying out SVG Tiny 1.2
            # text) leaves the first line where it is; JS_LINE_INCREMENT applies it.
            inc = f' data-inc="{float(el.get("line-increment")) / scale:g}"' if el.get("line-increment") else ""
            fo = attrs(el, skip=("font-family", "font-size", "font-weight", "font-style", "fill", "text-align",
                                 "x", "y", "width", "height", "transform", "line-increment"))
            g = f"{el.get('transform', '')} translate({el.get('x', '0')},{el.get('y', '0')}) scale({scale!r})".strip()
            out.append(f'<g transform="{g}"><foreignObject {fo} width="{float(el.get("width")) / scale:g}" '
                       f'height="{float(el.get("height")) / scale:g}" overflow="visible"><div xmlns="http://www.w3.org/1999/xhtml" '
                       f'class="ta"{inc} style="{style}">{body}</div></foreignObject></g>')
    out.append("</svg>")
    return "".join(out)


# A textarea's line-increment as the line height, with a negative top margin that cancels the
# half-leading CSS would add above the first line.
JS_LINE_INCREMENT = r"""() => document.querySelectorAll('div.ta[data-inc]').forEach(div => {
  const d = document.createElement('div'); d.textContent = 'x'; div.prepend(d);
  const normal = d.offsetHeight; d.remove();
  const L = parseFloat(div.dataset.inc);
  div.style.lineHeight = L + 'px'; div.style.marginTop = (-(L - normal) / 2) + 'px';
})"""


JS_SVG = r"""() => [...document.querySelectorAll('div.ta')].map(div => {
  if (!div.firstChild) return null;
  // The text up to the first <br>: the first paragraph, as on the HTML side.
  const br = [...div.childNodes].findIndex(n => n.nodeName.toLowerCase() === 'br');
  const rg = document.createRange(); rg.setStart(div, 0); rg.setEnd(div, br < 0 ? div.childNodes.length : br);
  const r = rg.getBoundingClientRect(), svg = document.querySelector('svg'), m = svg.getScreenCTM().inverse();
  const P = (x, y) => { const q = svg.createSVGPoint(); q.x = x; q.y = y; return q.matrixTransform(m); };
  const a = P(...(BASELINE)(div, r));
  const text = [...div.childNodes].map(n => n.nodeName.toLowerCase() === 'br' ? '\n' : n.textContent).join('');
  return {text, x: a.x, y: a.y};
}).filter(Boolean)""".replace("BASELINE", BASELINE, 1)


def pts_of(el):
    return [tuple(float(v) for v in xy.split(",")) for xy in el.get("points").split()]


def measure_iwb(page_, iwb: Path, shot: Path) -> dict:
    z, root, page, size, mapping, groups = iwb_page(iwb)
    if mapping is None:
        raise RuntimeError("no TESTMAP marker (converter was not instrumented)")
    s, ox, oy = mapping
    B = lambda x, y: ((x - ox) / s, (y - oy) / s)
    page_.set_viewport_size({"width": int(size[0]), "height": int(size[1])})
    page_.set_content("<html><body style='margin:0;background:#fff'>" + browser_svg(z, page, size, s) + "</body></html>")
    page_.wait_for_timeout(500)
    page_.evaluate(JS_LINE_INCREMENT)
    page_.screenshot(path=str(shot), full_page=True)
    texts = [dict(t, bx=B(t["x"], t["y"])[0], by=B(t["x"], t["y"])[1]) for t in page_.evaluate(JS_SVG)]

    images, rects, polys, by_group, bands = [], [], [], {}, {}
    for el in page:
        tag = local(el.tag); gid = groups.get(el.get("id"))
        if tag == "image":
            # Top-left, top-right and bottom-left corners, through translate(x,y) rotate(a).
            tr = re.match(r"\s*translate\(([-\d.]+)[ ,]+([-\d.]+)\)\s*rotate\(([-\d.]+)\)", el.get("transform", ""))
            x0, y0, ang = (float(tr[1]), float(tr[2]), math.radians(float(tr[3]))) if tr else (0.0, 0.0, 0.0)
            x0 += float(el.get("x")); y0 += float(el.get("y"))
            w, h, c, sn = float(el.get("width")), float(el.get("height")), math.cos(ang), math.sin(ang)
            images.append([B(x0, y0), B(x0 + c * w, y0 + sn * w), B(x0 - sn * h, y0 + c * h)])
        elif tag == "rect":
            x, y = B(float(el.get("x")), float(el.get("y")))
            box = (x, y, float(el.get("width")) / s, float(el.get("height")) / s)
            rects.append({"box": box, "borderless": el.get("stroke-width") == "0", "group": gid, "fill": el.get("fill")})
        elif tag == "polygon":
            polys.append([B(*p) for p in pts_of(el)])
            if el.get("stroke-width") == "0" and gid is not None:   # an angled gradient band
                bands.setdefault(gid, []).append(([B(*p) for p in pts_of(el)], hex_rgb(el.get("fill"))))
        if tag == "polyline":
            by_group.setdefault(gid if gid is not None else ("solo", el.get("id")), []).append(
                ([B(*p) for p in pts_of(el)], float(el.get("stroke-width") or 0) / s))

    # Notes: every borderless rect, plus the union of each group's gradient bands (borderless
    # rects, and polygons across an angled gradient). Each keeps its polygons and colours.
    rect_poly = lambda b: [(b[0], b[1]), (b[0] + b[2], b[1]), (b[0] + b[2], b[1] + b[3]), (b[0], b[1] + b[3])]
    notes = [r["box"] for r in rects if r["borderless"]]
    note_polys = [[(rect_poly(r["box"]), hex_rgb(r["fill"]))] for r in rects if r["borderless"]]
    for r in rects:
        if r["borderless"] and r["group"] is not None:
            bands.setdefault(r["group"], []).append((rect_poly(r["box"]), hex_rgb(r["fill"])))
    for g in bands.values():
        if len(g) >= 4:
            pts = [p for poly, _ in g for p in poly]
            x0 = min(p[0] for p in pts); y0 = min(p[1] for p in pts)
            notes.append((x0, y0, max(p[0] for p in pts) - x0, max(p[1] for p in pts) - y0)); note_polys.append(g)
    # Shapes: rect corners, polygons, and closed loops of edges (unfilled non-rectangles).
    shapes = [[(b[0], b[1]), (b[0] + b[2], b[1]), (b[0] + b[2], b[1] + b[3]), (b[0], b[1] + b[3])]
              for b in (r["box"] for r in rects)] + polys
    arrows, inks, connectors = [], [], []
    close = lambda p, q: math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-6 / s + 0.01
    for group_lines in by_group.values():
        lines = [l for l, _ in group_lines]
        # Ink: an open chain of many edges (connectors have at most 3, unfilled shapes close).
        chained = all(close(lines[i][1], lines[i + 1][0]) for i in range(len(lines) - 1))
        if len(lines) >= 8 and chained and not close(lines[-1][1], lines[0][0]):
            inks.append({"pts": [l[0] for l in lines] + [lines[-1][1]], "w": group_lines[0][1]})
            continue
        if len(lines) >= 3 and close(lines[-1][1], lines[0][0]) and all(close(lines[i][1], lines[i + 1][0]) for i in range(len(lines) - 1)):
            shapes.append([l[0] for l in lines])
        else:
            # A connector: its edges come first and chain end to end; its arrowheads follow
            # (each starts at an arm, not where the line ends). Underlines and fallback stickers
            # land here too, which is why connectors are matched to the nearest output line.
            chain = [lines[0][0], lines[0][1]]
            for l in lines[1:]:
                if not close(chain[-1], l[0]): break
                chain.append(l[1])
            connectors.append(chain)
        # Arrowheads: two consecutive edges that meet at the tip.
        for a, b in zip(lines, lines[1:]):
            if close(a[1], b[0]):
                arrows.append([a[0], a[1], b[1]])
    thumbs = page_.evaluate("async (srcs) => { const thumb = " + THUMB + "; const r = []; for (const s of srcs) r.push(await thumb(s)); return r; }",
                            [img_src(z, el) for el in page if local(el.tag) == "image"])
    return {"texts": texts, "images": images, "notes": notes, "note_polys": note_polys, "shapes": shapes,
            "arrows": arrows, "inks": inks, "thumbs": thumbs, "connectors": connectors}


# ----------------------------------------------------------------------------------------
# Comparisons
# ----------------------------------------------------------------------------------------
def seg_dist(p, a, b):
    abx, aby = b[0] - a[0], b[1] - a[1]; L = abx * abx + aby * aby
    t = 0 if L == 0 else max(0, min(1, ((p[0] - a[0]) * abx + (p[1] - a[1]) * aby) / L))
    return math.hypot(p[0] - a[0] - t * abx, p[1] - a[1] - t * aby)


def poly_dist(p, poly):
    return min(seg_dist(p, poly[i], poly[(i + 1) % len(poly)]) for i in range(len(poly)))


def bbox(pts):
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


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
    # Images (stickers and pictures): same order in both; largest corner distance (catches
    # rotation and scale).
    ie = [max(math.dist(a, b) for a, b in zip(s, o)) for s, o in zip(src["images"], out["images"])]
    # Content: mean colour difference (0-255) of 8 x 8 thumbnails drawn on white.
    td = [sum(abs(a - b) for a, b in zip(s, o)) / len(s) for s, o in zip(src["thumbs"], out["thumbs"]) if s and o]
    res["image"] = {"expected": len(src["images"]), "matched": len(out["images"]), "max": round(max(ie), 2) if ie else None,
                    "content": round(max(td), 1) if td else None}
    # Notes: pair each true note box with the nearest output box; largest edge error.
    # Its colour at five points: solid fill, or the gradient at the gradient's angle.
    ne, nc, pool = [], [], list(zip(out["notes"], out["note_polys"]))
    for s, paint in zip(src["notes"], src["note_paints"]):
        if not pool: break
        c = (s[0] + s[2] / 2, s[1] + s[3] / 2)
        j = min(range(len(pool)), key=lambda i: math.hypot(pool[i][0][0] + pool[i][0][2] / 2 - c[0], pool[i][0][1] + pool[i][0][3] / 2 - c[1]))
        o, polys = pool.pop(j)
        ne.append(max(abs(s[0] - o[0]), abs(s[1] - o[1]), abs(s[0] + s[2] - o[0] - o[2]), abs(s[1] + s[3] - o[1] - o[3])))
        nc.append(note_color_err(s, paint, solid_only, polys))
    res["note"] = {"expected": len(src["notes"]), "matched": len(ne), "max": round(max(ne), 2) if ne else None}
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
    # Shapes: the closest output outline (by bounding box, then both-way distance).
    sh, cands = [], [(o, bbox(o)) for o in out["shapes"]]
    for s in src["shapes"]:
        sb = bbox(s["pts"])
        near = sorted(cands, key=lambda c: max(abs(a - b) for a, b in zip(c[1], sb)))[:4]
        best = None
        for o, _ in near:
            d1 = max(poly_dist(p, o) for p in s["pts"][::4]); d2 = max(poly_dist(p, s["pts"]) for p in o)
            best = max(d1, d2) if best is None else min(best, max(d1, d2))
        if best is not None:
            sh.append({"label": s["label"], "dev": round(best, 2)})
    res["shape"] = {"expected": len(src["shapes"]), "matched": len(sh),
                    "max": max((r["dev"] for r in sh), default=None), "rows": sh}
    # Arrowheads: compare each true chevron with the closest output chevron.
    ae = []
    for s in src["arrows"]:
        best = min((max(math.hypot(a[0] - b[0], a[1] - b[1]) for a, b in zip(s, o)) for o in out["arrows"]), default=None)
        if best is not None: ae.append(best)
    res["arrow"] = {"expected": len(src["arrows"]), "matched": len(ae), "max": round(max(ae), 2) if ae else None}
    # Connector lines: each true line against the output line nearest its ends; largest distance
    # either way between the two.
    ce, pool = [], list(out["connectors"])
    for s in src["connectors"]:
        if not pool: break
        j = min(range(len(pool)), key=lambda i: min(math.dist(pool[i][0], s[0]) + math.dist(pool[i][-1], s[-1]),
                                                   math.dist(pool[i][0], s[-1]) + math.dist(pool[i][-1], s[0])))
        o = pool.pop(j)
        ce.append(max(max(path_dist(p, o) for p in s), max(path_dist(p, s) for p in o)))
    res["connector"] = {"expected": len(src["connectors"]), "matched": len(ce), "max": round(max(ce), 2) if ce else None}
    fails = []
    if res["text"]["matched"] < res["text"]["expected"]: fails.append("missing text")
    if res["text"]["max"] is not None and res["text"]["max"] > TOL["text"]: fails.append(f"text off by {res['text']['max']} px")
    for k in ("image", "note", "shape", "connector", "arrow", "ink"):
        if res[k]["matched"] != res[k]["expected"]: fails.append(f"{k} count {res[k]['matched']}/{res[k]['expected']}")
        if res[k]["max"] is not None and res[k]["max"] > TOL[k]: fails.append(f"{k} off by {res[k]['max']} px")
    if res["note_color"]["max"] is not None and res["note_color"]["max"] > TOL["note_color"]:
        fails.append(f"note colour off by {res['note_color']['max']}/255")
    if res["ink"]["width"] is not None and res["ink"]["width"] > TOL["ink_width"]:
        fails.append(f"ink width off by {res['ink']['width']} px")
    if res["image"]["content"] is not None and res["image"]["content"] > TOL["image_content"]:
        fails.append(f"image content differs by {res['image']['content']}")
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
        if k in ("ink", "connector") and not v["expected"] and not v["matched"]:
            return "<td>—</td>"
        if k == "ink":
            return f"<td>{v['matched']}/{v['expected']} · max {v['max']} · width {v['width']}</td>"
        extra = f" · content {v['content']}" if v.get('content') is not None else ""
        return f"<td>{v['matched']}/{v['expected']}" + (f" · max {v['max']}" if v['max'] is not None else "") + extra + "</td>"
    rows = "".join(
        f"<tr class='{'bad' if r['status'] != 'PASS' else ''}'><td>{r['board']}</td><td>{r['fill']}</td><td>{r['status']}</td>"
        + "".join(cell(r, k) for k in ("text", "image", "note", "note_color", "shape", "connector", "arrow", "ink"))
        + f"<td>{'; '.join(r.get('fails', []))}</td></tr>" for r in results)
    boards = sorted({r["board"] for r in results})
    figs = ""
    for b in boards:
        imgs = f"<figure><img src='shots/{b}_html.png'><figcaption>Whiteboard export</figcaption></figure>"
        for f in FILLS:
            if (out_root / "shots" / f"{b}_{f}.png").exists():
                imgs += f"<figure><img src='shots/{b}_{f}.png'><figcaption>{f} notes → .iwb</figcaption></figure>"
        figs += f"<h2 id='{b}'>{b}</h2><div class=grid>{imgs}</div>"
    doc = f"""<!doctype html><html><head><meta charset=utf-8><title>Visual checks</title><style>
body{{font:14px/1.45 system-ui,sans-serif;margin:24px;color:#1d1b22;background:#fbfbfd}}
table{{border-collapse:collapse;font-variant-numeric:tabular-nums}} td,th{{border:1px solid #ddd;padding:3px 8px}}
th{{background:#f1eff6}} tr.bad td{{background:#fde8e8}} .grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
figure{{margin:0;background:#f1eff6;padding:6px;border-radius:6px}} img{{width:100%;background:#fff}}
figcaption{{font-size:12px;color:#666;text-align:center}}</style></head><body>
<h1>Whiteboard → IWB visual checks</h1>
<p>{time.strftime('%Y-%m-%d %H:%M')} · errors in board px · tolerances: text {TOL['text']}, image {TOL['image']}, note {TOL['note']}, shape {TOL['shape']}, connector {TOL['connector']}, arrowhead {TOL['arrow']}, ink {TOL['ink']} (width {TOL['ink_width']}); note colour {TOL['note_color']}/255.
Renders show the IWB page as a standard SVG renderer draws it (textarea as a wrapped text box).</p>
<table><tr><th>Board</th><th>Notes</th><th>Status</th><th>Text</th><th>Images</th><th>Notes</th><th>Note colour</th><th>Shapes</th><th>Connectors</th><th>Arrowheads</th><th>Ink</th><th>Problems</th></tr>{rows}</table>
{figs}</body></html>"""
    (out_root / "report.html").write_text(doc, encoding="utf-8")


# ----------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--scripts", type=Path, default=REPO, help="folder with the converter script")
    ap.add_argument("--samples", type=Path, default=REPO / "samples")
    ap.add_argument("--out", type=Path, default=REPO / "tests" / "out" / "visual")
    ap.add_argument("--sample", nargs="*", help="only these sample folder names")
    ap.add_argument("--fill", nargs="*", choices=FILLS, default=list(FILLS))
    ap.add_argument("--powershell", help="PowerShell executable (default: powershell.exe on Windows, else pwsh)")
    a = ap.parse_args()

    ps = find_powershell(a.powershell)
    a.out.mkdir(parents=True, exist_ok=True); (a.out / "shots").mkdir(exist_ok=True)
    samples = sorted(d for d in a.samples.iterdir() if d.is_dir() and (not a.sample or d.name in a.sample))
    script = instrument(a.scripts / SCRIPT, a.out / "_instrumented" / SCRIPT)
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        hpage = browser.new_page(viewport={"width": 1600, "height": 1000}, device_scale_factor=2)
        ipage = browser.new_page(device_scale_factor=2)
        for d in samples:
            html = next(d.glob("*.html"), None)
            if not html: continue
            try:
                src = measure_html(hpage, html, a.out / "shots" / f"{d.name}_html.png")
            except Exception as e:
                results.append({"board": d.name, "fill": "-", "status": "ERROR", "fails": [f"HTML: {e}"]}); continue
            for f in a.fill:
                r = {"board": d.name, "fill": f}
                iwb, log = convert(ps, script, html, a.out / "iwb" / f.lower(), f)
                if not iwb:
                    r.update(status="FAIL", fails=["conversion failed: " + log[-300:]]); results.append(r); continue
                try:
                    out = measure_iwb(ipage, iwb, a.out / "shots" / f"{d.name}_{f}.png")
                    r["checks"] = compare(src, out, solid_only=(f == "Solid"))
                    r["fails"] = r["checks"].pop("fails")
                    r["status"] = "FAIL" if r["fails"] else "PASS"
                except Exception as e:
                    r.update(status="ERROR", fails=[str(e)])
                results.append(r)
                c = r.get("checks", {})
                print(f"{d.name:26} {f:8} {r['status']:5}  text max {c.get('text', {}).get('max')}  "
                      f"image {c.get('image', {}).get('max')}  note {c.get('note', {}).get('max')}  shape {c.get('shape', {}).get('max')}  "
                      f"connector {c.get('connector', {}).get('max')}  arrow {c.get('arrow', {}).get('max')}  note colour {c.get('note_color', {}).get('max')}  "
                      f"ink {c.get('ink', {}).get('max')}  {'; '.join(r.get('fails', []))}", flush=True)
        browser.close()
    (a.out / "results.json").write_text(json.dumps(results, indent=1), encoding="utf-8")
    write_report(a.out, results)
    bad = [r for r in results if r["status"] != "PASS"]
    print(f"\n{len(results)} check(s): {len(results) - len(bad)} passed, {len(bad)} failed. Report: {a.out / 'report.html'}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
