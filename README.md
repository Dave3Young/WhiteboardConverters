# WhiteboardConverters
Scripts for converting Whiteboard exports to other platform file formats.

Each converter reads the self-contained **HTML export of a Microsoft Whiteboard board** and
writes an editable file for another whiteboard application. It converts these objects:
- shapes;
- plain text;
- sticky notes;
- connectors with their arrowheads;
- reaction stickers;
- embedded images.

| Folder | Output | Opens in |
| --- | --- | --- |
| [`WhiteboardToUBZ`](WhiteboardToUBZ) | `.ubz` | [OpenBoard](https://openboard.ch/) |
| [`WhiteboardToExcalidraw`](WhiteboardToExcalidraw) | `.excalidraw` | [Excalidraw](https://excalidraw.com/) |
| [`WhiteboardToIWB`](WhiteboardToIWB) | `.iwb` (IMS Interactive Whiteboard CFF 1.0) | OpenBoard (File > Import) and other IWB-capable whiteboard apps |

All the scripts are PowerShell and need Windows 10/11 with Windows PowerShell 5.1 or later.

## Usage
Every converter takes the same basic parameters:

```powershell
# One file: writes the output next to the input
.\Convert-WhiteboardHtmlToIwb.ps1 -InputPath .\MyBoard.html

# Many files: pipe them in and pick an output folder
Get-ChildItem .\Exports\*.html | .\Convert-WhiteboardHtmlToIwb.ps1 -OutputDirectory .\Out -Force
```

- **`-InputPath`** is the HTML export. It also accepts files piped from `Get-ChildItem`.
- **`-OutputPath`** is the output file for one input.
- **`-OutputDirectory`** is the output folder for many inputs.
- **`-Force`** overwrites existing output.
- **`-PassThru`** returns the created files.

Each run ends with a summary of anything that couldn't be reproduced exactly. For example,
dashed lines drawn solid are listed there, never silently dropped.

## Sticky-note colours
Whiteboard draws sticky notes with a colour gradient and stores a solid fallback colour
alongside it. Most target formats have no gradient fill, so each converter offers both:
- **Solid:** the note uses Whiteboard's fallback colour, and stays a single fill you can
  recolour.
- **Gradient:** the note is a stack of thin colour bands (32 by default, set with
  `-GradientSteps`), grouped so they move together.

## WhiteboardToUBZ: OpenBoard native documents
| Script | Sticky notes |
| --- | --- |
| `Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1` | solid |
| `Convert-WhiteboardHtmlToOpenBoard-v3_gradient.ps1` | gradient bands |

This writes OpenBoard's own `.ubz` document format.
- **Text:** plain text stays editable, keeping its line breaks, bold, alignment and rotation.
- **Stickers:** they keep Whiteboard's original vector artwork.
- **Shapes:** each one is traced from its exact outline, so ovals stay round and rotated
  shapes stay rotated. A shape's fill and border are grouped, and so is a connector with its
  arrowheads.
- **Canvas:** `-CanvasPadding` (default 100 px) adds space around the board.

## WhiteboardToExcalidraw: Excalidraw scenes
| Script | Sticky notes |
| --- | --- |
| `Convert-WhiteboardHtmlToExcalidraw-solid.ps1` | solid |
| `Convert-WhiteboardHtmlToExcalidraw-gradient.ps1` | gradient bands |

This writes an `.excalidraw` JSON scene with native, editable Excalidraw elements.
- **Shapes:** rectangles, ovals and diamonds become Excalidraw's own shapes. Other outlines
  become closed lines.
- **Rotation:** rotated objects keep their rotation.
- **Text:** it is wrapped into the same lines as in Whiteboard.
- **Images and stickers:** they are embedded in the file.
- **Canvas:** `-CanvasPadding` (default 100 px) adds space around the board.

**Tip:** in the Excalidraw desktop app, images can show as grey placeholders when a file is
opened with Open. Drag the file onto the window instead.

## WhiteboardToIWB: IMS Interactive Whiteboard files
| Script | Sticky notes |
| --- | --- |
| `Convert-WhiteboardHtmlToIwb.ps1` | `-NoteFill Solid` (default) or `-NoteFill Gradient` |

This writes an `.iwb` file in the IMS Common File Format (CFF 1.0). CFF is the vendor-neutral
format for exchanging files between interactive-whiteboard applications. Within it, the
script sticks to the subset that OpenBoard's importer reads correctly.

- **Page size:** IWB is page-based. The whole board is scaled to fit one page, 1280 × 960 by
  default. That size imports into OpenBoard at exactly 1:1. Change it with `-PageWidth`,
  `-PageHeight` and `-PageMargin`.
- **Stickers and WebP images:** CFF doesn't allow SVG images, so Whiteboard's vector stickers
  and any WebP images are converted to PNG using headless Microsoft Edge (or Google Chrome).
  - Use `-BrowserPath` to choose the browser.
  - Use `-NoRasterize` to keep them as SVG instead. OpenBoard shows SVG, but other apps may
    not.
- **Known OpenBoard import limits:** dashed lines import solid, and imported text is
  read-only. The file itself keeps the dashes.

## Tests
Each folder has its own `samples\` (59 real Whiteboard exports) and two test suites.

```powershell
.\tests\Invoke-SmokeTest.ps1      # structure and content checks, pure PowerShell
.\tests\Setup-VisualTests.ps1     # once: Python venv + Playwright Chromium
.\tests\Invoke-VisualTests.ps1    # compares geometry against Whiteboard's real layout
```

The visual tests write an HTML report to `tests\out\visual\report.html`. Each folder's
`CLAUDE.md`, where present, documents the source and target formats in detail.

## License
[GPL-3.0](LICENSE)
