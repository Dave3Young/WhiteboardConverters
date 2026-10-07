# WhiteboardToExcalidraw
Converts a **Microsoft Whiteboard HTML export** into an editable **Excalidraw scene**
(`.excalidraw`).

| Script | Sticky notes |
| --- | --- |
| `Convert-WhiteboardHtmlToExcalidraw-solid.ps1` | solid: Whiteboard's fallback colour |
| `Convert-WhiteboardHtmlToExcalidraw-gradient.ps1` | the note's colour gradient, drawn as grouped colour bands |

The two scripts are otherwise identical.

## Requirements
Windows 10/11 with Windows PowerShell 5.1 or later.

## Usage
```powershell
# One board: writes MyBoard.excalidraw next to MyBoard.html
.\Convert-WhiteboardHtmlToExcalidraw-solid.ps1 -InputPath .\MyBoard.html

# Many boards
Get-ChildItem .\Exports\*.html |
    .\Convert-WhiteboardHtmlToExcalidraw-gradient.ps1 -OutputDirectory .\Excalidraw -Force
```

| Parameter | Default | Meaning |
| --- | --- | --- |
| `-InputPath` | | The HTML export or exports. It accepts files piped from `Get-ChildItem`. |
| `-OutputPath` | `<input>.excalidraw` | Output file (for one input). |
| `-OutputDirectory` | | Output folder (for many inputs). |
| `-CanvasPadding` | 100 | Space in px added around the board. |
| `-GradientSteps` | 32 | Number of colour bands per note (gradient script only). |
| `-Force` | | Overwrite existing output. |
| `-PassThru` | | Return a result object for each conversion. |

Open the file at [excalidraw.com](https://excalidraw.com/) or in the Excalidraw desktop app.
- **Desktop app:** opening a file there with Open can show images as grey placeholders. Drag
  the file onto the window instead.
- **Image copies:** each image is embedded in the `.excalidraw` file. A copy is also written
  to a `<name>_images` folder next to it, for your reference.

## What gets converted
| Whiteboard object | In Excalidraw |
| --- | --- |
| Plain text | Editable text, wrapped into the same lines as in Whiteboard and re-wrapped by Excalidraw when you edit it. |
| Sticky note | A filled rectangle in the note's colour (solid, or banded along its gradient) with its text. |
| Shape | Rectangles, ovals and diamonds become native Excalidraw shapes, sized to Whiteboard's exact outline. Any other shape becomes a closed line. Text inside a shape becomes a separate text element. |
| Connector | A native Excalidraw arrow or line, keeping its arrowheads and dashes. Elbow and curved connectors keep every bend as a point. |
| Freehand ink | Each stroke as a line at the pen's width, grouped together. |
| Reaction sticker | Whiteboard's original vector artwork. |
| Image | The embedded image. Very large images are downscaled. |

Rotated objects keep their rotation.

## Limits
The run summary reports every item below when it occurs.
- **Mirrored objects** are drawn unmirrored, because Excalidraw can't mirror.
- **Stickers without artwork** are drawn as a fallback symbol.
- **Untraceable shapes:** a shape whose outline can't be traced is built from its label.
- **Font:** text uses Excalidraw's Helvetica, and bold is dropped. Text set in another face
  in Whiteboard (newer exports use Segoe UI) is narrower, and is counted.
- **Ink:** each stroke is drawn at one width in one colour. A pressure-sensitive, highlighter or
  effect pen is approximated.
- **Elbow and curved connectors** are traced through every bend (curves as many short steps),
  but no sample has one, so they are tested only on hand-edited paths. Arcs are drawn straight.
- **Underline:** Excalidraw text can't be underlined, so a line is drawn under each line of
  underlined text and grouped with it. If you edit the text, the line doesn't follow the new
  wording.

## Tests
```powershell
.\tests\Invoke-SmokeTest.ps1      # structure and content, all samples x both scripts
.\tests\Setup-VisualTests.ps1     # once: Python venv, Playwright Chromium, @excalidraw/utils (-InstallPython if needed)
.\tests\Invoke-VisualTests.ps1    # geometry vs. Whiteboard's real layout
```

- **Samples:** `samples\` holds 60 real Whiteboard exports.
- **Visual tests:** they measure the original board in headless Chromium and compare every
  text, note, shape, sticker, image, connector, arrowhead and ink stroke, and every note's
  colour, with the converted scene. They
  also fail any line of text that Excalidraw would clip. The report is written to
  `tests\out\visual\report.html`, and shows each board next to both converted scenes, drawn
  by Excalidraw's own exporter. Setup installs that exporter with npm, so it needs Node.js.
  Without it the checks still run, but the report shows only the Whiteboard pictures.
