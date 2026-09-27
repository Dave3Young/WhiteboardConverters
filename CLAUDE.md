# WhiteboardToUBZ

PowerShell converters from **Microsoft Whiteboard HTML exports** to **OpenBoard `.ubz`**.

| Script | Sticky notes |
| --- | --- |
| `Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1` | solid fill (Whiteboard's fallback `background-color`) |
| `Convert-WhiteboardHtmlToOpenBoard-v3_gradient.ps1` | CSS `linear-gradient` drawn as 32 grouped colour bands (`-GradientSteps`) |

The two scripts are the same apart from note backgrounds. **Every fix must go into both.**

The target platform is Windows 10/11 with Windows PowerShell 5.1 (`#requires -Version 5.1`).
PowerShell 7 works too, except that image downscaling uses `System.Drawing`, which is
Windows-only.

## Conventions
- Keep the scripts **ASCII-only**. Windows PowerShell 5.1 reads BOM-less files as ANSI.
- Keep **CRLF** line endings, which is how the scripts are stored.
- Keep `Set-StrictMode -Version 2.0` clean.
- Watch the comma operator: `@($a + $b, $c)` parses as `$a + ($b, $c)`. Parenthesise the
  arithmetic, or build point lists with `List[double[]]` and `.Add()`.
- OpenBoard `ub:uuid` / `ub:parent` values must be canonical GUID strings (`New-UbzUuid`).
- Anything the converter can't reproduce faithfully is **reported in the run summary**, never
  silently dropped or approximated. Examples: dashed strokes drawn solid, rotated non-text
  objects drawn unrotated, sticker fallbacks, shapes built from their label.
- Two Windows PowerShell 5.1 traps under `powershell.exe -File`: `$PSScriptRoot` is empty while
  an advanced script's `param()` defaults are evaluated (resolve paths in the body), and with
  stdin redirected the script's `process {}` block runs once per stdin line, i.e. never for
  empty stdin. Launch scripts from other processes with `-Command "& 'script.ps1' ..."`.

## Tests (run before and after every change)
```powershell
.\tests\Invoke-SmokeTest.ps1          # pure PowerShell: structure + content, all samples x both scripts
.\tests\Setup-VisualTests.ps1         # once: Python venv + Playwright Chromium (-InstallPython if needed)
.\tests\Invoke-VisualTests.ps1        # geometry vs. Whiteboard's real layout; opens tests\out\visual\report.html
```
- `samples\` holds 59 real Whiteboard exports. Add new ones as `samples\<Name>\<Name>.html`,
  plus its `-comments.json`.
- **Smoke test** (`Invoke-SmokeTest.ps1`), which judges the output alone:
  - the zip and XML are valid; `ub:version` is 4.8.0; GUIDs are valid; referenced images exist
    and parse; all content is inside the viewBox;
  - every PlainText string is present verbatim, line breaks included (`<br>` in the Qt HTML),
    with no `?` glyphs;
  - stickers are embedded as SVG; ovals are drawn round;
  - arrowheads are grouped with their line; rotated text and shapes are really rotated.
- **Visual checks** (`tests\visual\run_visual_checks.py`):
  - how they work: they convert with a temporary *instrumented copy* of each script (it adds
    one `<!-- TESTCENTER x y -->` comment so board coordinates can be recovered). They then
    measure the Whiteboard HTML and the rendered `.ubz` in headless Chromium.
  - tolerances, in board px: text 5, sticker 0.5, note 0.5, shape 1.5, arrowhead 0.5.
  - current results (59 samples): text ≤ 3.0 (median 0–1), stickers 0.0, notes ≤ 0.01,
    shapes ≤ 0.92, arrowheads 0.01.
- Both suites were mutation-tested against the original pre-fix scripts, and flag every bug
  listed below.

## Source format: Whiteboard HTML export
- Objects are sibling `<div class="anchor" data-whiteboard-type="...">` elements. Split them
  the way the scripts do, with the regex `(?=<div\s+class="anchor\b)`.
- Types seen: `PlainText`, `Shape`, `Note`, `Connector`, `ReactionStickers`, `Image`,
  `AzureImage`. `CommentThread` appears only in CSS and is not an object.
- **Position:**
  - `left`/`top` plus an optional `transform: matrix(a, b, c, d, e, f)` with
    `transform-origin: 0 0`.
  - The scale is the **column length** √(a²+b²), not `|a|`. A rotated object has `a = d = 0`.
  - `align center` anchors (Shape, Image, ReactionStickers) put the object's *centre* at
    left/top. `align topLeft` anchors (PlainText, Note, Connector) use left/top as the corner.
- **PlainText:**
  - The export's stylesheet has `.textbox.plainText > div { padding: 16px }`, and Draft.js
    adds 1 px more on the left. So the text starts at **(left + 17·s, top)**: the matrix's
    translateY of −16·s cancels the top padding.
  - The column width is `outerWidth − 33` (pre-scale). Alignment comes from the
    `DraftEditor-align*` class.
  - The weight is an inline `font-weight` on `.textBoxCore`: template headings are 600, some
    labels 700. The face is `sans-serif`, which renders as Arial.
- **Line breaks:** a soft break inside a span is stored as a raw **CR or CRLF**, which a browser
  reads as a newline under `pre-wrap`. `Get-HtmlText` normalises both to LF.
- **Note text:** at **(+14, +1)** from the note's outer corner (1 px border, 12 px
  `.textBoxCoreWrapper` side padding, 1 px Draft.js gutter; the inline `padding: 0px` on
  `.textBoxCore` overrides the stylesheet's 12/16 px). The column is `width − 25`, top-left
  aligned, **24 px** (`.stickyNote.fixedFontSize`, on every note seen) and **bold** (700).
  The font stack starts with Aptos, and renders as Segoe UI when Aptos is absent.
- **Shape:** the first `<path d>` inside `svg.shape > g` is the exact outline, centred by
  `translate(w/2 h/2)` on the `<g>`.
  - Rectangles are M/L/Z.
  - Ovals are 8 `Q` curves, and are **labelled "oval"**, not ellipse or circle.
  - A second, hidden path is only the hit area.
- **ReactionStickers:** the `<img>` holds Whiteboard's full-colour artwork as
  `data:image/svg+xml;base64` (64×64; the light bulb is 40×40). It is sized by
  `max-width`/`max-height` only, so CSS shrinks it to fit and never enlarges it.
- **Connector:** the first `<path>` is the line.
  - Each arrowhead is a separate open chevron `d="M-5 -9 L0 0 L5 -9"` with
    `transform="translate(x,y), rotate(deg)"` and its tip at the origin.
  - Dashed lines use `stroke-dasharray`.
- **Note:** `.stickyNote { border-width: 1px }` sits outside the CSS `width`/`height`, and the
  background fills it, so a 304 × 304 note shows as **306 × 306** at (left, top).
- **Note background:** the background is `background-color` (the solid fallback) plus
  `background-image: ..., linear-gradient(start, end)`. The fallback can differ visibly from
  the gradient; blue notes are the clearest case.
- **Image / AzureImage:** a base64 data URI. AzureImage declares the MIME type `image/*`, so the
  real format is sniffed from the bytes.

## Target format: OpenBoard `.ubz`
All details below were verified against OpenBoard's own `src/adaptors/UBSvgSubsetAdaptor.cpp`.
- **Archive:** a ZIP containing `metadata.rdf`, `page000.svg` and `images/` at its root.
  - `ub:version` is `4.8.0`.
  - The `ub` namespace is `http://uniboard.mnemis.com/document`.
- **Coordinates:** the scene is centred on the origin, and items are positioned only by
  `transform="matrix(a, b, c, d, e, f)"`.
  - OpenBoard reads the matrix in standard SVG order (`fromSvgTransform`). Text items apply
    it, **rotation included**.
- **Text:** `<foreignObject ub:type="text">` with `<itemTextContent>`, which holds the escaped
  Qt rich-text HTML.
  - Qt collapses raw newlines to spaces, so line breaks are written as `<br />`.
  - The `<p>` style carries `font-weight`, so bold headings keep their width and alignment.
- **Images:** width/height on `<image>` are **ignored**, because pixmap and SVG items aren't
  `UBResizableGraphicsItem`.
  - Set width/height to the file's native size, and put the scale in the matrix:
    display size ÷ native size.
  - The native size comes from the file header (PNG, GIF, BMP, JPEG), or from the SVG's
    width/height or viewBox.
- **SVG images:** an `href` ending in `.svg` loads as a native vector `UBGraphicsSvgItem`, so
  there is no need to rasterise.
- **No shape primitives:**
  - A fill is an ink `<polygon>`. It auto-closes, so don't repeat the first point.
  - A border is a separate `<polyline>`. Repeat the first point to close it.
  - A connector is a `<line>`.
  - Elements that share a `ub:parent` GUID form one movable object: a shape's fill plus its
    outline, a note's gradient bands, a connector plus its arrowheads.
- **No dashed or gradient styles:** dashes are drawn solid (and counted in the summary), and
  gradients become grouped bands.

## History of fixes (details in the claude.ai project "Whiteboard format converter")
1. **Stickers:** 60 of 102 on ImageBoard came out as black `?`. Only 4 labels were
   hand-mapped. The fix embeds the original SVG artwork (`Get-EmbeddedImageData`), with SVG
   size parsing in `Get-RasterImageDimensions`.
2. **Ovals came out square:** the "oval" label wasn't matched. The fix traces every shape
   from its outline path (`Get-ShapePathPoints`), with the label kept only as a fallback that
   now includes `oval`.
3. **Text was placed 17·s px left and 16·s px high:** the fix applies the plain-text insets
   (`$PlainTextInsetLeft/Right/Top`) and narrows the column.
4. **Rotated text crashed the conversion (AssumptionGrid):** the scale was read from the
   matrix diagonal. The fix uses the column length, rotates PlainText natively, and
   counts/warns on other rotated types.
5. **Arrowheads were dropped:** the fix converts each chevron to a polyline grouped with its
   line through `ub:parent`.
6. **Notes were 2 px too small:** the 1 px `.stickyNote` border was ignored. The fix adds
   `2 * $NoteBorder` to the note size; the visual checks now measure notes too.
7. **Line breaks were lost:** raw CR/CRLF went into the Qt `<p>` and collapsed to spaces (e.g.
   "Subject: Instructor: Course Date:" on one line). The fix normalises them in `Get-HtmlText`
   and writes `<br />`.
8. **Note text was 11 px too low, at 20 px regular:** the inset was a guess. The fix uses the
   measured `$NoteTextInset*` constants, 24 px and the note's weight.
9. **Rotated shapes were drawn unrotated:** 180-degree block arrows pointed the wrong way. The
   fix traces the outline in local px and maps it through the full matrix
   (`ConvertTo-BoardPoints`); shape labels turn with the shape.
10. **Bold was dropped:** weight-600 headings drew narrower, so centred and right-aligned ones
    moved 5-31 px. The fix writes `font-weight` (`Get-FontWeight`).

## Open items
- **Shape text** is unverified: no sample has any.
- **Rotation** is native for PlainText and shapes. Other rotated types (notes, images,
  stickers) are drawn unrotated at the correct size; none are rotated in the samples.
- **Elbow/curved connectors:** none seen yet. The main line reads only its first 2 points.
- **Formatting losses:** underline and the font face are dropped (text is Arial). Note text
  is drawn in Arial, not Segoe UI/Aptos, so its lines can wrap a little differently. A space
  after a soft line break is collapsed. Dashed strokes are drawn solid.
- **Option:** v2 notes could use the gradient midpoint instead of the solid fallback colour.
