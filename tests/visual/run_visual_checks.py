"""
Visual / geometric checks for the Whiteboard -> IWB converter.

For every sample export and both note fills (Solid, Gradient) this:
  1. converts the export with an *instrumented temporary copy* of the converter (the only
     change is one XML comment recording the board-to-page mapping, so board coordinates can
     be recovered exactly -- the script in the repo is never modified);
  2. renders the Whiteboard HTML in headless Chromium and measures where Whiteboard really
     puts every text, sticker/image, note, shape outline and connector arrowhead;
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
TOL = {"text": 5.0, "image": 0.5, "note": 0.5, "shape": 1.5, "arrow": 0.5, "image_content": 12.0}

# An image's 8 x 8 thumbnail drawn on white, as 192 RGB values (JS function source).
THUMB = r"""async (src) => { const im = new Image(); im.src = src; await im.decode();
  const c = document.createElement('canvas'); c.width = c.height = 8; const g = c.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, 8, 8); g.drawImage(im, 0, 0, 8, 8);
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
# ----------------------------------------------------------------------------------------
JS_HTML = (r"""async () => {
  const cal = [], texts = [], images = [], shapes = [], arrows = [], notes = [];
  const thumb = THUMB;
  const pt = (m, x, y) => [m.a*x + m.c*y + m.e + scrollX, m.b*x + m.d*y + m.f + scrollY];
  for (const a of document.querySelectorAll('div.anchor[data-whiteboard-type]')) {
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
      texts.push({type, text: t.replace(/\r\n?/g, '\n'), x: b.left + scrollX, y: b.top + scrollY, w: b.width, h: b.height});
    }
    if (type === 'ReactionStickers' || type === 'Image' || type === 'AzureImage') {
      const im = a.querySelector('img');
      if (im) { const i = im.getBoundingClientRect(); images.push([i.left + scrollX, i.top + scrollY, i.width, i.height, await thumb(im.src)]); }
    }
    if (type === 'Note') {
      const b = a.querySelector('.textBoxBackground').getBoundingClientRect();
      notes.push([b.left + scrollX, b.top + scrollY, b.width, b.height]);
    }
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
  }
  // The anchor divs themselves are 0 x 0 (their content overflows them), so take the union
  // of everything drawn inside them; comment threads are not board objects.
  const all = [...document.querySelectorAll('div.anchor[data-whiteboard-type]:not([data-whiteboard-type=CommentThread]) *')]
              .map(e => e.getBoundingClientRect()).filter(r => r.width > 0 && r.height > 0);
  const bbox = all.length ? [Math.min(...all.map(r => r.left)) + scrollX, Math.min(...all.map(r => r.top)) + scrollY,
                             Math.max(...all.map(r => r.right)) + scrollX, Math.max(...all.map(r => r.bottom)) + scrollY] : null;
  return {cal, texts, images, shapes, arrows, notes, bbox};
}""").replace("THUMB", THUMB, 1)


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
        "images": [(*to_b(s[0], s[1]), s[2] / kx, s[3] / ky) for s in d["images"]],
        "thumbs": [s[4] for s in d["images"]],
        "notes": [(*to_b(s[0], s[1]), s[2] / kx, s[3] / ky) for s in d["notes"]],
        "shapes": [{"label": s["label"], "pts": [to_b(*q) for q in s["pts"]]} for s in d["shapes"]],
        "arrows": [[to_b(*q) for q in a] for a in d["arrows"]],
    }


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
        lines[-1] += c.tail or ""
    return lines


def browser_svg(z: zipfile.ZipFile, page, size) -> str:
    """The page as plain SVG: images inlined, each textarea as a wrapped HTML box."""
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
            style = (f"font-family:'{el.get('font-family')}';font-size:{el.get('font-size')};font-weight:{el.get('font-weight')};"
                     f"color:{el.get('fill')};text-align:{align};white-space:pre-wrap;overflow-wrap:break-word;"
                     "line-height:normal;margin:0;width:100%;overflow:visible")
            body = "<br/>".join(htmllib.escape(line) for line in textarea_lines(el))
            fo = attrs(el, skip=("font-family", "font-size", "font-weight", "fill", "text-align"))
            out.append(f'<foreignObject {fo} overflow="visible"><div xmlns="http://www.w3.org/1999/xhtml" '
                       f'class="ta" style="{style}">{body}</div></foreignObject>')
    out.append("</svg>")
    return "".join(out)


JS_SVG = r"""() => [...document.querySelectorAll('div.ta')].map(div => {
  if (!div.firstChild) return null;
  const rg = document.createRange(); rg.selectNodeContents(div);
  const r = rg.getBoundingClientRect(), svg = document.querySelector('svg'), m = svg.getScreenCTM().inverse();
  const P = (x, y) => { const q = svg.createSVGPoint(); q.x = x; q.y = y; return q.matrixTransform(m); };
  const a = P(r.left, r.top);
  const text = [...div.childNodes].map(n => n.nodeName.toLowerCase() === 'br' ? '\n' : n.textContent).join('');
  return {text, x: a.x, y: a.y};
}).filter(Boolean)"""


def pts_of(el):
    return [tuple(float(v) for v in xy.split(",")) for xy in el.get("points").split()]


def measure_iwb(page_, iwb: Path, shot: Path) -> dict:
    z, root, page, size, mapping, groups = iwb_page(iwb)
    if mapping is None:
        raise RuntimeError("no TESTMAP marker (converter was not instrumented)")
    s, ox, oy = mapping
    B = lambda x, y: ((x - ox) / s, (y - oy) / s)
    page_.set_viewport_size({"width": int(size[0]), "height": int(size[1])})
    page_.set_content("<html><body style='margin:0;background:#fff'>" + browser_svg(z, page, size) + "</body></html>")
    page_.wait_for_timeout(500)
    page_.screenshot(path=str(shot), full_page=True)
    texts = [dict(t, bx=B(t["x"], t["y"])[0], by=B(t["x"], t["y"])[1]) for t in page_.evaluate(JS_SVG)]

    images, rects, polys, by_group = [], [], [], {}
    for el in page:
        tag = local(el.tag); gid = groups.get(el.get("id"))
        if tag == "image":
            x, y = B(float(el.get("x")), float(el.get("y")))
            images.append((x, y, float(el.get("width")) / s, float(el.get("height")) / s))
        elif tag == "rect":
            x, y = B(float(el.get("x")), float(el.get("y")))
            box = (x, y, float(el.get("width")) / s, float(el.get("height")) / s)
            rects.append({"box": box, "borderless": el.get("stroke-width") == "0", "group": gid})
        elif tag == "polygon":
            polys.append([B(*p) for p in pts_of(el)])
        if tag == "polyline":
            by_group.setdefault(gid if gid is not None else ("solo", el.get("id")), []).append([B(*p) for p in pts_of(el)])

    # Notes: every borderless rect, plus the union of each group's gradient bands.
    notes = [r["box"] for r in rects if r["borderless"]]
    band_groups = {}
    for r in rects:
        if r["borderless"] and r["group"] is not None:
            band_groups.setdefault(r["group"], []).append(r["box"])
    for boxes in band_groups.values():
        if len(boxes) >= 4:
            x0 = min(b[0] for b in boxes); y0 = min(b[1] for b in boxes)
            notes.append((x0, y0, max(b[0] + b[2] for b in boxes) - x0, max(b[1] + b[3] for b in boxes) - y0))
    # Shapes: rect corners, polygons, and closed loops of edges (unfilled non-rectangles).
    shapes = [[(b[0], b[1]), (b[0] + b[2], b[1]), (b[0] + b[2], b[1] + b[3]), (b[0], b[1] + b[3])]
              for b in (r["box"] for r in rects)] + polys
    arrows = []
    close = lambda p, q: math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-6 / s + 0.01
    for lines in by_group.values():
        if len(lines) >= 3 and close(lines[-1][1], lines[0][0]) and all(close(lines[i][1], lines[i + 1][0]) for i in range(len(lines) - 1)):
            shapes.append([l[0] for l in lines])
        # Arrowheads: two consecutive edges that meet at the tip.
        for a, b in zip(lines, lines[1:]):
            if close(a[1], b[0]):
                arrows.append([a[0], a[1], b[1]])
    thumbs = page_.evaluate("async (srcs) => { const thumb = " + THUMB + "; const r = []; for (const s of srcs) r.push(await thumb(s)); return r; }",
                            [img_src(z, el) for el in page if local(el.tag) == "image"])
    return {"texts": texts, "images": images, "notes": notes, "shapes": shapes, "arrows": arrows, "thumbs": thumbs}


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


def compare(src: dict, out: dict) -> dict:
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
    # Images (stickers and pictures): same order in both.
    ie = [max(abs(a - b) for a, b in zip(s, o)) for s, o in zip(src["images"], out["images"])]
    # Content: mean colour difference (0-255) of 8 x 8 thumbnails drawn on white.
    td = [sum(abs(a - b) for a, b in zip(s, o)) / len(s) for s, o in zip(src["thumbs"], out["thumbs"]) if s and o]
    res["image"] = {"expected": len(src["images"]), "matched": len(out["images"]), "max": round(max(ie), 2) if ie else None,
                    "content": round(max(td), 1) if td else None}
    # Notes: pair each true note box with the nearest output box; largest edge error.
    ne, pool = [], list(out["notes"])
    for s in src["notes"]:
        if not pool: break
        c = (s[0] + s[2] / 2, s[1] + s[3] / 2)
        j = min(range(len(pool)), key=lambda i: math.hypot(pool[i][0] + pool[i][2] / 2 - c[0], pool[i][1] + pool[i][3] / 2 - c[1]))
        o = pool.pop(j)
        ne.append(max(abs(s[0] - o[0]), abs(s[1] - o[1]), abs(s[0] + s[2] - o[0] - o[2]), abs(s[1] + s[3] - o[1] - o[3])))
    res["note"] = {"expected": len(src["notes"]), "matched": len(ne), "max": round(max(ne), 2) if ne else None}
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
    fails = []
    if res["text"]["matched"] < res["text"]["expected"]: fails.append("missing text")
    if res["text"]["max"] is not None and res["text"]["max"] > TOL["text"]: fails.append(f"text off by {res['text']['max']} px")
    for k in ("image", "note", "shape", "arrow"):
        if res[k]["matched"] != res[k]["expected"]: fails.append(f"{k} count {res[k]['matched']}/{res[k]['expected']}")
        if res[k]["max"] is not None and res[k]["max"] > TOL[k]: fails.append(f"{k} off by {res[k]['max']} px")
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
        extra = f" · content {v['content']}" if v.get('content') is not None else ""
        return f"<td>{v['matched']}/{v['expected']}" + (f" · max {v['max']}" if v['max'] is not None else "") + extra + "</td>"
    rows = "".join(
        f"<tr class='{'bad' if r['status'] != 'PASS' else ''}'><td>{r['board']}</td><td>{r['fill']}</td><td>{r['status']}</td>"
        + "".join(cell(r, k) for k in ("text", "image", "note", "shape", "arrow"))
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
<p>{time.strftime('%Y-%m-%d %H:%M')} · errors in board px · tolerances: text {TOL['text']}, image {TOL['image']}, note {TOL['note']}, shape {TOL['shape']}, arrowhead {TOL['arrow']}.
Renders show the IWB page as a standard SVG renderer draws it (textarea as a wrapped text box).</p>
<table><tr><th>Board</th><th>Notes</th><th>Status</th><th>Text</th><th>Images</th><th>Notes</th><th>Shapes</th><th>Arrowheads</th><th>Problems</th></tr>{rows}</table>
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
                    r["checks"] = compare(src, out)
                    r["fails"] = r["checks"].pop("fails")
                    r["status"] = "FAIL" if r["fails"] else "PASS"
                except Exception as e:
                    r.update(status="ERROR", fails=[str(e)])
                results.append(r)
                c = r.get("checks", {})
                print(f"{d.name:26} {f:8} {r['status']:5}  text max {c.get('text', {}).get('max')}  "
                      f"image {c.get('image', {}).get('max')}  note {c.get('note', {}).get('max')}  shape {c.get('shape', {}).get('max')}  "
                      f"arrow {c.get('arrow', {}).get('max')}  {'; '.join(r.get('fails', []))}", flush=True)
        browser.close()
    (a.out / "results.json").write_text(json.dumps(results, indent=1), encoding="utf-8")
    write_report(a.out, results)
    bad = [r for r in results if r["status"] != "PASS"]
    print(f"\n{len(results)} check(s): {len(results) - len(bad)} passed, {len(bad)} failed. Report: {a.out / 'report.html'}")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
