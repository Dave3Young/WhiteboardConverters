# WhiteboardToIWB

A PowerShell converter from **Microsoft Whiteboard HTML exports** to **IWB**, the IMS Interactive
Whiteboard Common File Format 1.0. IWB files open in OpenBoard (File > Import) and in other
interactive-whiteboard applications.

| Script | Notes |
| --- | --- |
| `Convert-WhiteboardHtmlToIwb.ps1` | `-NoteFill Solid` (default: the note's fallback `background-color`) or `Gradient` (`-GradientSteps` colour bands) |

The HTML parsing is ported from `Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1` in
`C:\git\dey\WhiteboardToUBZ`. Its CLAUDE.md documents the Whiteboard export format in full, and
the fixes made there apply here too.

The target platform is Windows 10/11 with Windows PowerShell 5.1 (`#requires -Version 5.1`).
- Image downscaling and sticker rasterising use `System.Drawing`.
- Rasterising also uses headless Microsoft Edge, or Chrome as a fallback.

## Conventions
- Keep the script **ASCII-only** and use **CRLF** line endings.
- Keep `Set-StrictMode -Version 2.0` clean.
- **Literals:**
  - Watch the comma operator: `@($a + $b, $c)` parses as `$a + ($b, $c)`.
  - Watch integer literals in `[Math]::Max`/`Min`: `[Math]::Max(1, $w)` rounds a fractional
    `$w`, so write `1.0`.
- **Launching scripts:** under `powershell.exe -File`, `$PSScriptRoot` is empty while `param()`
  defaults are evaluated. Stdin also feeds the `process {}` block. So launch scripts from other
  processes with `-Command "& 'script.ps1' ..."`.
- **Reporting losses:** anything the converter can't reproduce faithfully is **reported in the
  run summary**, never silently dropped. Examples: rotated non-shape objects drawn unrotated,
  stickers left as SVG, unfilled shapes written as edges.
- **Temporary files:** the sticker rasteriser reuses one work folder,
  `%TEMP%\WhiteboardToIwb-render`, and serialises access with a named mutex. Two runs sharing
  that folder once swapped each other's stickers.

## Tests (run before and after every change)
```powershell
.\tests\Invoke-SmokeTest.ps1          # pure PowerShell: structure, OpenBoard-safe subset, content; all samples x both note fills
.\tests\Setup-VisualTests.ps1         # once: Python venv + Playwright Chromium (-InstallPython if needed)
.\tests\Invoke-VisualTests.ps1        # geometry + image content vs. Whiteboard's real layout; opens tests\out\visual\report.html
```
- `samples\` holds 59 real Whiteboard exports; it's the same set as WhiteboardToUBZ.
- **Smoke test:**
  - **Structure:** the ZIP has `content.xml` at its root, and the root is `iwb` version 1.0 in
    the IMS namespace. It contains `meta` elements, then `svg`, then `group` elements. The
    lower-case `viewbox` is present, and the page is a single `pageset/page`.
  - **Ids and images:** ids are unique and every group ref resolves. Images are PNG, JPEG, GIF
    or BMP.
  - **OpenBoard subset:**
    - Every shape has an explicit stroke colour, and polygons have a real fill.
    - Polylines have exactly 2 points.
    - Font sizes are in `pt`, and weights are `normal` or `bold`.
    - There are no exponents, and text transforms are exactly `translate(x,y) rotate(a)`.
  - **Content:**
    - Every PlainText string is present verbatim, with `tbreak` as the line break, and all
      content is inside the page.
    - Ovals are drawn round, arrowheads are all converted, and rotated text and shapes are
      really rotated.
- **Visual checks** (`tests\visual\run_visual_checks.py`):
  - **Method:** they convert with an instrumented copy of the script, which adds one
    `<!-- TESTMAP scale ox oy -->` comment so page coordinates can be mapped back to board
    coordinates. They render the page as a standard SVG renderer would, with each textarea as a
    wrapped HTML box. Then they measure it against the Whiteboard HTML in headless Chromium.
  - **Image content:** images are also compared by content, using the mean colour difference of
    8 × 8 thumbnails.
  - **Tolerances,** in board px: text 5, image 0.5, note 0.5, shape 1.5, arrowhead 0.5, image
    content 12/255.
  - **Current results,** for 59 samples × 2 fills: all pass. Text is ≤ 2.4 px off (median ≤ 0.1),
    images ≤ 0.18, notes ≤ 0.01, shapes ≤ 0.93, arrowheads 0.0, image content ≤ 9.
- **Mutation testing** (2026-09-27): each suite catches the bugs it is meant to catch:
  - a swapped sticker image: visual;
  - line breaks lost: both suites;
  - an unfilled polygon, which OpenBoard paints black: smoke;
  - note text 11 px off: visual;
  - arrowheads as 3-point shapes: both suites.

## Target format: IWB (IMS CFF 1.0)
Verified against OpenBoard's `src/adaptors/UBCFFSubsetAdaptor.cpp` (the importer) and
`plugins/cffadaptor/src/UBCFFAdaptor.cpp` / `UBCFFConstants.h` (its exporter).
- **Package:** a ZIP containing `content.xml` and `images/`.
- **Root element:** `<iwb xmlns="http://www.imsglobal.org/xsd/iwb_v1p0" version="1.0">`, with
  these children in order:
  - `iwb:meta name/content` pairs;
  - `svg:svg width height viewbox`, with **viewbox in lower case**;
  - a `svg:pageset` holding one `svg:page`;
  - then `iwb:group` elements, each wrapping `iwb:element ref="id"` members.
- **Allowed SVG elements:** rect, ellipse, line, polyline, polygon, text, textarea (with
  `tspan`/`tbreak`), image, audio, video.
- **Image formats:** PNG, JPEG, GIF, BMP, TIFF, WMF and EMF. **SVG is not an allowed image
  format.**
- **Importer quirks this converter works around.** OpenBoard's importer is buggy:
  - **Page scale:** it scales the page by `min(its page / IWB viewbox)`. Its own page is
    1280 × 960 (4:3) or 1696 × 960 (16:9).
    - It doesn't scale a text item's x position, so text drifts sideways unless the scale is 1.
    - **So the board is fitted into a 1280 × 960 page, which imports 1:1.** Don't change the
      default page size without a reason.
  - **Unsupported elements:** it ignores `line`, and it treats a `polyline` as a closed, filled
    shape. So every line is a **2-point polyline**.
  - **Fills:** it fills **every polygon**, painting a missing or `none` fill black. `rect` fills
    only a valid colour.
    - An axis-aligned rectangle is a `rect` (`fill="none"` works).
    - A filled outline is a `polygon`.
    - An unfilled non-rectangle is written as its **edges**.
  - **Strokes:** it draws a missing or invalid stroke as a 1 px **black** outline. So every
    shape has a stroke; a borderless one uses its own fill colour at width 0.
  - **Positioning:** `ellipse` placement is buggy (it uses `cx - 2rx`), so ovals are traced
    polygons.
  - **Textarea:**
    - `font-size` is read as **pt**, with the "pt" suffix stripped.
    - `font-weight` is read only by name, from normal, light, demibold, bold and black.
    - Leading and trailing spaces of each text run are trimmed.
    - Text is imported read-only.
  - **Transforms:** it reads only `translate(x,y)` and `rotate(a)` (about the item's own
    corner), split on spaces. It can't read matrices or exponents.
    - Rotated text is written as `x="0" y="0" transform="translate(x,y) rotate(a)"`, which is
      also correct in standard SVG.
  - **Colours:** colours are `#rrggbb`, and `fill-opacity` is ignored.
  - **Dashes:** `stroke-dasharray` is ignored, so dashes import solid. The IWB file keeps them.
  - **Stroke-width offset:** a polygon's y position is off by its integer stroke width, which is
    about 1 px at typical scales.
- **Standard renderers:** a textarea shows only the lines that fit its height, and OpenBoard
  squashes an item whose content is taller than its box. So text heights are generous: the
  estimated line count + 1, at 1.35 em per line.

## Open items
- **Not verified in OpenBoard:** no import has been tested in OpenBoard itself, which isn't
  installed here. Everything above comes from reading its source.
- **Other readers:** nothing has been tested in other IWB readers (SMART Notebook, ActivInspire).
- **Shape text** is unverified: no sample has any.
- **Rotation** is native for PlainText and shapes. Rotated notes, images and stickers are drawn
  unrotated; none appear in the samples.
- **Formatting losses:**
  - Underline and the font face are dropped.
  - Plain text is drawn in Arial, and note text in Segoe UI (Aptos is not assumed).
- **Page size:** there is no 1:1 (unfitted) page option. A large board is scaled down to fit
  the page, so its text becomes small.
