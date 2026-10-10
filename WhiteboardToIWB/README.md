# WhiteboardToIWB
Converts a **Microsoft Whiteboard HTML export** into an **IWB file**: the IMS Interactive
Whiteboard Common File Format (CFF 1.0). CFF is the vendor-neutral format for exchanging
lessons between interactive-whiteboard applications. The output is tuned to import correctly
into [OpenBoard](https://openboard.ch/) (File > Import).

## Requirements
- Windows 10/11 with Windows PowerShell 5.1.
- Microsoft Edge (or Google Chrome), to convert stickers to PNG. Without a browser the
  conversion still works, and the stickers stay SVG.

## Usage
```powershell
# One board: writes MyBoard.iwb next to MyBoard.html
.\Convert-WhiteboardHtmlToIwb.ps1 -InputPath .\MyBoard.html

# Many boards, with gradient sticky notes
Get-ChildItem .\Exports\*.html |
    .\Convert-WhiteboardHtmlToIwb.ps1 -OutputDirectory .\IWB -NoteFill Gradient -Force
```

| Parameter | Default | Meaning |
| --- | --- | --- |
| `-InputPath` | | The HTML export or exports. It accepts files piped from `Get-ChildItem`. |
| `-OutputPath` | `<input>.iwb` | Output file (for one input). |
| `-OutputDirectory` | | Output folder (for many inputs). |
| `-NoteFill` | `Solid` | `Solid` uses Whiteboard's fallback note colour. `Gradient` draws the note's colour gradient as bands. |
| `-GradientSteps` | 32 | Number of colour bands per note with `-NoteFill Gradient`. |
| `-PageWidth`, `-PageHeight` | 1280, 960 | Page size. The whole board is scaled down to fit. |
| `-PageMargin` | 20 | Space in px between the board and the page edge. |
| `-BrowserPath` | auto | Edge or Chrome executable used to convert stickers to PNG. |
| `-NoRasterize` | | Keep stickers and WebP images as they are instead of converting them to PNG. |
| `-Force` | | Overwrite existing output. |
| `-PassThru` | | Return a result object for each conversion. |

## Why the page is 1280 × 960
IWB is page-based, so the whole board is fitted onto one page, and a large board comes out
small. OpenBoard's own page is 1280 × 960, so an IWB page of that size imports at exactly
1:1. At other sizes, OpenBoard's importer shifts text boxes sideways. Change the page size
only if you're targeting another application.

## What gets converted
| Whiteboard object | In the IWB file |
| --- | --- |
| Plain text | A text area, keeping line breaks, bold, alignment, rotation and the export's font (Arial or Segoe UI). |
| Sticky note | A filled rectangle in the note's colour (solid, or banded along its gradient) with its text, in Segoe UI. |
| Shape | A rectangle, or a polygon traced from Whiteboard's exact outline, so ovals stay round and rotated shapes stay rotated. Unfilled outlines that aren't rectangles are written as their edges. |
| Connector | A line grouped with its arrowheads, following every bend of an elbow or curved connector. Dashes are kept. |
| Freehand ink | Each stroke as lines at the pen's width, grouped together. |
| Reaction sticker | Whiteboard's original artwork, converted to PNG. |
| Image | PNG, JPEG, GIF or BMP. WebP images are converted to PNG, and very large images are downscaled. |

The parts of each object (a shape's fill and border, a note's bands and text, a connector
and its arrowheads) are grouped so they move together.

## Limits
The run summary reports these when they occur.
- **Losses in OpenBoard's importer:** it draws dashed lines solid, and imported text is
  read-only.
- **Stickers kept as SVG:** without Edge/Chrome, or with `-NoRasterize`, stickers stay SVG.
  OpenBoard shows them, but other IWB apps may not, because SVG isn't a CFF image format.
- **Rotated notes, images and stickers** are drawn unrotated. None of the samples has any.
- **Formatting:** note text is drawn in Segoe UI, not Whiteboard's Aptos. It sits at
  Whiteboard's baseline, but a long note can wrap at a different word.
- **Ink:** each stroke is drawn at one width in one colour. A pressure-sensitive, highlighter or
  effect pen is approximated.
- **Underline:** text keeps `text-decoration="underline"`, but OpenBoard's importer ignores it,
  so a single unrotated line also gets a drawn underline grouped with it. Wrapped or rotated
  underlined text has no drawn line, and the run summary says so.
- **Elbow and curved connectors** are traced through every bend, but no sample has one, so
  they are tested only on hand-edited paths. Arcs are drawn straight.
- **Untested imports:** the files haven't yet been tried in OpenBoard itself, or in other IWB
  readers such as SMART Notebook or ActivInspire. The conversions are checked against
  OpenBoard's importer source code and the test suites below.

## Tests
```powershell
.\tests\Invoke-SmokeTest.ps1      # structure, OpenBoard-safe subset and content; all samples x both note fills
.\tests\Setup-VisualTests.ps1     # once: Python venv + Playwright Chromium (-InstallPython if needed)
.\tests\Invoke-VisualTests.ps1    # geometry and image content vs. Whiteboard's real layout
```

- **Samples:** `samples\` holds 60 real Whiteboard exports.
- **Output:** converted files are written to `tests\out\solid` and `tests\out\gradient`.
- **Visual tests:** the report is written to `tests\out\visual\report.html`.
