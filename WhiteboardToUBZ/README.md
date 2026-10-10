# WhiteboardToUBZ
Converts a **Microsoft Whiteboard HTML export** into an **OpenBoard `.ubz` document**.

| Script | Sticky notes |
| --- | --- |
| `Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1` | solid: Whiteboard's fallback colour |
| `Convert-WhiteboardHtmlToOpenBoard-v3_gradient.ps1` | the note's colour gradient, drawn as grouped colour bands |

The two scripts are otherwise identical.

## Requirements
- Windows 10/11 with Windows PowerShell 5.1.
- PowerShell 7 also works, except for downscaling large images, which uses the Windows-only
  `System.Drawing`.

## Usage
```powershell
# One board: writes MyBoard.ubz next to MyBoard.html
.\Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1 -InputPath .\MyBoard.html

# Many boards
Get-ChildItem .\Exports\*.html |
    .\Convert-WhiteboardHtmlToOpenBoard-v3_gradient.ps1 -OutputDirectory .\OpenBoard -Force
```

| Parameter | Default | Meaning |
| --- | --- | --- |
| `-InputPath` | | The HTML export or exports. It accepts files piped from `Get-ChildItem`. |
| `-OutputPath` | `<input>.ubz` | Output file (for one input). |
| `-OutputDirectory` | | Output folder (for many inputs). |
| `-CanvasPadding` | 100 | Space in px added around the board. |
| `-GradientSteps` | 32 | Number of colour bands per note (gradient script only). |
| `-Force` | | Overwrite existing output. |
| `-PassThru` | | Return a result object for each conversion. |

Open the `.ubz` in [OpenBoard](https://openboard.ch/).

## What gets converted
| Whiteboard object | In OpenBoard |
| --- | --- |
| Plain text | Editable text, keeping paragraphs, line breaks, alignment, rotation, line spacing, the export's font (Arial, Segoe UI, Aptos, ...) and bold, italic and underline, including on parts of a text. |
| Sticky note | A filled rectangle in the note's colour (solid, or banded along its gradient) with its text. |
| Shape | A fill and border traced from Whiteboard's exact outline and grouped together. Ovals stay round, and rotated shapes stay rotated. |
| Connector | A line grouped with its arrowheads, following every bend of an elbow or curved connector. |
| Freehand ink | Each stroke as a line at the pen's width, grouped together. |
| Reaction sticker | Whiteboard's original vector artwork (SVG), keeping its rotation. |
| Image | The embedded image, keeping its rotation. Very large images are downscaled. |

## Limits
The run summary reports every item below when it occurs.
- **Dashed lines** are drawn solid, because OpenBoard has no dashed ink style.
- **Rotated notes** are drawn unrotated at the right size. None of the samples has any.
- **Formatting:** note text is drawn in Arial, not Whiteboard's Aptos. Underline is kept.
- **Ink:** each stroke is drawn at one width in one colour. A pressure-sensitive, highlighter or
  effect pen is approximated.
- **Curved connectors** are traced through every bend, but no sample has one, so they are
  tested only on hand-edited paths. Elbow connectors are tested (RotatedTest1).

## Tests
```powershell
.\tests\Invoke-SmokeTest.ps1      # structure and content, all samples x both scripts
.\tests\Setup-VisualTests.ps1     # once: Python venv + Playwright Chromium (-InstallPython if needed)
.\tests\Invoke-VisualTests.ps1    # geometry vs. Whiteboard's real layout
```

- **Samples:** `samples\` holds 61 real Whiteboard exports.
- **Visual tests:** they measure the original board in headless Chromium and compare every
  text, note, shape, sticker, image, connector, arrowhead and ink stroke position, and every note's colour, with
  the converted file. The report is
  written to `tests\out\visual\report.html`.
