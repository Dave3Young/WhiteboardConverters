#requires -Version 5.1
<#
.SYNOPSIS
Converts a Microsoft Whiteboard HTML export into an IWB document (IMS Interactive Whiteboard
Common File Format 1.0), for OpenBoard and other interactive-whiteboard applications.

.DESCRIPTION
The HTML parsing (anchors, positions, CSS matrices, text, shape outlines, stickers, images,
colours) is ported from Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1 / -v3_gradient.ps1 in
WhiteboardToUBZ. The output is an .iwb file: a ZIP with content.xml at its root and an
images/ folder.

content.xml is the IMS CFF shape of SVG: an <iwb> root holding one <svg:svg> with a
<svg:pageset> of one <svg:page>, followed by <iwb:group> elements that tie the parts of one
Whiteboard object together. Only the CFF SVG subset is written, and within it only the forms
that OpenBoard's CFF importer (src/adaptors/UBCFFSubsetAdaptor.cpp) reads correctly:

  - Page: a fixed page (-PageWidth x -PageHeight, default 1280 x 960) with the whole board
    scaled to fit inside it. IWB is page-based, and OpenBoard scales an imported page by
    min(its page size / the IWB page size); its page is 1280 x 960 (4:3) or 1696 x 960 (16:9),
    so a 1280 x 960 IWB page imports at exactly 1:1. At any other scale OpenBoard's importer
    moves text boxes sideways (it doesn't scale a text item's x).
  - Rectangles (axis-aligned) are <svg:rect>, which OpenBoard fills only when a fill is given.
  - Other filled outlines (ovals, block arrows, diamonds, rotated shapes) are <svg:polygon>.
    OpenBoard fills every polygon, so an unfilled non-rectangular outline is written as its
    edges instead.
  - Lines -- connectors, arrowheads, unfilled outlines -- are 2-point <svg:polyline>s:
    OpenBoard ignores <svg:line> and fills a polyline as a closed shape, which only draws a
    2-point polyline correctly.
  - Text is <svg:textarea> with <svg:tbreak/> line breaks; font sizes in pt (the unit
    OpenBoard reads them in) and weights as normal/bold (the only names it maps).
    Rotated text uses transform="translate(x,y) rotate(a)".
  - Images are PNG/JPEG/GIF/BMP files in images/. CFF has no SVG images, so Whiteboard's
    SVG sticker artwork (and WebP images) are rasterised to PNG with headless Microsoft Edge
    or Google Chrome. Without a browser (or with -NoRasterize) they stay SVG, which OpenBoard
    shows but other IWB applications may not; the summary says so.
  - Every stroke is written with an explicit colour and every filled shape with an explicit
    stroke (its own fill colour, width 0, when it has no border): OpenBoard draws a missing
    stroke as a 1 px black outline.

Dashed strokes keep their stroke-dasharray (OpenBoard's importer draws them solid). Sticky
notes are solid (the export's fallback background-color) or, with -NoteFill Gradient, a
stack of -GradientSteps colour bands approximating the CSS linear-gradient at its angle.
Freehand ink (InkGroup) becomes each stroke's centreline at the pen's width, as 2-point
polylines grouped together.

.EXAMPLE
.\Convert-WhiteboardHtmlToIwb.ps1 -InputPath .\sailboat.html

.EXAMPLE
Get-ChildItem .\Exports\*.html | .\Convert-WhiteboardHtmlToIwb.ps1 -OutputDirectory .\IWB -NoteFill Gradient -Force
#>
[CmdletBinding(DefaultParameterSetName = 'File')]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('FullName')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string[]] $InputPath,

    [Parameter(ParameterSetName = 'File')]
    [string] $OutputPath,

    [Parameter(ParameterSetName = 'Directory')]
    [string] $OutputDirectory,

    # Solid: the note's fallback background-color. Gradient: colour bands (-GradientSteps).
    [ValidateSet('Solid', 'Gradient')]
    [string] $NoteFill = 'Solid',

    [ValidateRange(4, 128)]
    [int] $GradientSteps = 32,

    # The IWB page. Keep the height at 960 and the width at most 1280 for a 1:1 OpenBoard import.
    [ValidateRange(320, 20000)]
    [int] $PageWidth = 1280,

    [ValidateRange(240, 20000)]
    [int] $PageHeight = 960,

    # Space kept free around the board on the page, in page px.
    [ValidateRange(0, 200)]
    [double] $PageMargin = 20,

    # msedge.exe / chrome.exe used to rasterise SVG stickers (default: found automatically).
    [string] $BrowserPath,

    # Keep SVG stickers as SVG instead of rasterising them to PNG.
    [switch] $NoRasterize,

    [switch] $Force,
    [switch] $PassThru
)

begin {
    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'

    # ---------------------------------------------------------------------------------
    # IWB constants (IMS CFF 1.0, as written and read by OpenBoard's UBCFFAdaptor.cpp /
    # UBCFFSubsetAdaptor.cpp).
    # ---------------------------------------------------------------------------------
    $IWB_NS   = 'http://www.imsglobal.org/xsd/iwb_v1p0'
    $SVG_NS   = 'http://www.w3.org/2000/svg'
    $XLINK_NS = 'http://www.w3.org/1999/xlink'
    $ci       = [Globalization.CultureInfo]::InvariantCulture

    # Whiteboard plain-text inset (pre-scale px): ".textbox.plainText > div { padding: 16px; }"
    # in the export's stylesheet, plus Draft.js's 1px left caret gutter.
    $PlainTextInsetLeft  = 17
    $PlainTextInsetRight = 16
    $PlainTextInsetTop   = 16

    # ".stickyNote { border-width: 1px }" sits outside the note's CSS width/height, and the
    # note's background fills that border too: a 304 x 304 note shows as 306 x 306.
    $NoteBorder = 1

    # Sticky-note text inset from the note's outer corner (pre-scale px): 1px border + 12px
    # ".textBoxCoreWrapper" side padding + Draft.js's 1px caret gutter on the left, and only the
    # border on top. The column is the CSS width minus 2 x 12 and the gutter. Note text is 24px
    # (".stickyNote.fixedFontSize"), bold, in Aptos with Segoe UI as the installed fallback.
    $NoteTextInsetLeft = 14
    $NoteTextInsetTop  = 1
    $NoteTextColumnInset = 25
    $NoteFontFamily = 'Segoe UI'
    # Aptos isn't assumed to be installed where the file is opened, so note text is written in
    # Segoe UI. Under line-height normal a line's baseline sits one ascent below its top: Aptos
    # 1923/2048 em, Segoe UI 2210/2048 em. Note text is raised by the difference (3.4px at
    # 24px) so its baseline lands where Whiteboard's Aptos baseline does.
    $NoteBaselineRaise = (2210 - 1923) / 2048

    # Whiteboard's newer note style (class "noteVisualUpdate" on .textBoxBackground, first seen
    # in CompareAndContrast2) has no border, and puts a 40px author title bar
    # (".titleBarWithAttribution { height: 40px }") above the CSS-sized text area: a 304 x 265
    # note shows as 304 x 305, with its text 13px in and 40px down. Its colours are given only
    # as a class (e.g. "paleGreenGradient"); see Get-NoteColorTable.
    $NoteTitleBarHeight = 40

    # Fill of a note whose colour can't be found (Whiteboard's default yellow); counted in the
    # run summary.
    $DefaultNoteFill = '#fee15a'

    # Caches for Get-InstalledFontNames and Get-FontLineMetrics.
    $script:InstalledFontNames = $null
    $script:FontLineMetrics = @{}

    # Plain text and shape labels are "sans-serif" in the export, which renders as Arial.
    $TextFontFamily = 'Arial'

    # A drawn underline's distance below the top of its text, in font sizes.
    $UnderlineOffset = 1.05

    # Long side, in px, of a rasterised sticker.
    $StickerRasterSize = 256

    # ===================================================================================
    # Parsing helpers -- from Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1
    # ===================================================================================

    function Get-FirstMatch {
        param([string]$Text, [string]$Pattern, [int]$Group = 1)
        $m = [regex]::Match($Text, $Pattern,
            [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [Text.RegularExpressions.RegexOptions]::Singleline)
        if ($m.Success) { return $m.Groups[$Group].Value }
        return $null
    }

    function Get-CssNumber {
        param([string]$Style, [string]$Name, [double]$Default = 0)
        $value = Get-FirstMatch $Style ('(?:^|;)\s*' + [regex]::Escape($Name) + '\s*:\s*(-?[0-9.]+)px')
        if ($null -eq $value) { return $Default }
        return [double]::Parse($value, [Globalization.CultureInfo]::InvariantCulture)
    }

    function Get-Number {
        param([string]$Value, [double]$Default = 0)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
        $n = 0.0
        if ([double]::TryParse($Value, [Globalization.NumberStyles]::Float,
                [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
        return $Default
    }

    function Get-ScalarDouble {
        param($Value, [double]$Default = 0)
        foreach ($candidate in @($Value)) {
            if ($null -eq $candidate) { continue }
            $number = 0.0
            if ([double]::TryParse([string]$candidate,
                    [Globalization.NumberStyles]::Float,
                    [Globalization.CultureInfo]::InvariantCulture,
                    [ref]$number)) { return [double]$number }
        }
        return [double]$Default
    }

    function Convert-RgbaToHex {
        param([string]$Color, [string]$Default = '#000000', [switch]$TransparentAllowed)
        if ([string]::IsNullOrWhiteSpace($Color) -or $Color -eq 'none') {
            if ($TransparentAllowed) { return 'transparent' }
            return $Default
        }
        if ($Color -match '^#([0-9a-f]{3}|[0-9a-f]{6})$') { return $Color.ToLowerInvariant() }
        if ($Color -match 'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([0-9.]+))?\s*\)') {
            if ($TransparentAllowed -and $Matches[4] -ne '' -and [double]$Matches[4] -eq 0) { return 'transparent' }
            return '#{0:x2}{1:x2}{2:x2}' -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3]
        }
        return $Default
    }

    function Get-StyleValue {
        param([string]$Style, [string]$Name)
        return Get-FirstMatch $Style ('(?:^|;)\s*' + [regex]::Escape($Name) + '\s*:\s*([^;]+)')
    }

    function Get-TextRuns {
        # The text of an anchor as runs of uniformly styled text. Draft.js writes each paragraph
        # as a <div data-block="true"> (an empty one holds only <br data-text="true">), and each
        # run as <span data-offset-key style="..."><span data-text="true">text</span></span>,
        # where the outer style carries per-run bold, italic and underline. Paragraphs are
        # joined by a "`n" run; they used to be run together with no break. Weight is $null
        # when the run has no weight of its own.
        param([string]$Block)
        $opts = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
        $paragraphs = @([regex]::Matches($Block, '<div\b[^>]*\bdata-block="true"[^>]*>(.*?)(?=<div\b[^>]*\bdata-block="true"|$)', $opts) |
            ForEach-Object { $_.Groups[1].Value })
        if ($paragraphs.Count -eq 0) { $paragraphs = @($Block) }
        $runs = New-Object Collections.Generic.List[object]
        for ($p = 0; $p -lt $paragraphs.Count; $p++) {
            if ($p -gt 0) { [void]$runs.Add([pscustomobject]@{ Text = "`n"; Weight = $null; Italic = $false; Underline = $false; Strike = $false }) }
            foreach ($m in [regex]::Matches($paragraphs[$p], '(?:<span\b([^>]*)>\s*)?<span\s+data-text="true"[^>]*>(.*?)</span>', $opts)) {
                # Whiteboard stores soft line breaks inside a span as a bare CR or CRLF (the
                # export keeps them raw); a browser reads both as a newline, so normalise them.
                $text = [Net.WebUtility]::HtmlDecode(($m.Groups[2].Value -replace '<[^>]+>', '')) -replace "`r`n?", "`n"
                $style = [Net.WebUtility]::HtmlDecode([string](Get-FirstMatch $m.Groups[1].Value '\bstyle="([^"]*)"'))
                $weight = $null
                if (Get-StyleValue $style 'font-weight') { $weight = Get-FontWeight $style 400 }
                $decoration = "$(Get-StyleValue $style 'text-decoration')"
                [void]$runs.Add([pscustomobject]@{
                    Text = $text; Weight = $weight
                    Italic = ("$(Get-StyleValue $style 'font-style')" -match '\b(italic|oblique)\b')
                    Underline = ($decoration -match '\bunderline\b'); Strike = ($decoration -match '\bline-through\b')
                })
            }
        }
        return ,($runs.ToArray())
    }

    function Test-StyledRuns {
        # True when any run carries its own bold, italic, underline or strikethrough.
        param([object[]]$Runs)
        foreach ($r in $Runs) { if ($null -ne $r.Weight -or $r.Italic -or $r.Underline -or $r.Strike) { return $true } }
        return $false
    }

    function Get-HtmlText {
        param([string]$Block)
        return -join @((Get-TextRuns $Block) | ForEach-Object { $_.Text })
    }

    function Get-Transform {
        # Reads a CSS matrix(a, b, c, d, e, f). The scale is the length of each column
        # (sqrt(a^2+b^2), sqrt(c^2+d^2)); a 90-degree-rotated object has a = d = 0.
        param([string]$Style)
        $result = @{ ScaleX = 1.0; ScaleY = 1.0; TranslateX = 0.0; TranslateY = 0.0
                     A = 1.0; B = 0.0; C = 0.0; D = 1.0; IsRotated = $false }
        $matrix = Get-FirstMatch $Style 'transform\s*:\s*matrix\(([^)]+)\)'
        if ($matrix) {
            $v = @($matrix -split '\s*,\s*' | ForEach-Object { Get-Number $_ })
            if ($v.Count -ge 6) {
                $sx = [Math]::Sqrt(($v[0] * $v[0]) + ($v[1] * $v[1]))
                $sy = [Math]::Sqrt(($v[2] * $v[2]) + ($v[3] * $v[3]))
                if ($sx -gt 0) { $result.ScaleX = $sx }
                if ($sy -gt 0) { $result.ScaleY = $sy }
                $result.A = $v[0]; $result.B = $v[1]; $result.C = $v[2]; $result.D = $v[3]
                $result.TranslateX = $v[4]; $result.TranslateY = $v[5]
                $result.IsRotated = ([Math]::Abs($v[1]) -gt 1e-6 -or [Math]::Abs($v[2]) -gt 1e-6 -or $v[0] -lt 0 -or $v[3] -lt 0)
            }
        }
        return $result
    }

    function Get-AnchorInfo {
        param([string]$Block)
        $tag = Get-FirstMatch $Block '^(<div\s+class="anchor\b[^>]*>)'
        $style = Get-FirstMatch $tag 'style="([^"]*)"'
        $class = Get-FirstMatch $tag 'class="([^"]*)"'
        $transform = Get-Transform $style
        return @{
            Type = Get-FirstMatch $tag 'data-whiteboard-type="([^"]+)"'
            Key = Get-FirstMatch $tag 'data-apikey="([^"]+)"'
            X = (Get-ScalarDouble (Get-CssNumber $style 'left')) + (Get-ScalarDouble $transform.TranslateX)
            Y = (Get-ScalarDouble (Get-CssNumber $style 'top')) + (Get-ScalarDouble $transform.TranslateY)
            ScaleX = Get-ScalarDouble $transform.ScaleX 1
            ScaleY = Get-ScalarDouble $transform.ScaleY 1
            IsCentered = ($class -match '\balign\s+center\b')
            MA = $transform.A; MB = $transform.B; MC = $transform.C; MD = $transform.D
            IsRotated = $transform.IsRotated
        }
    }

    function Get-ImageSize {
        param([string]$Block)
        $componentStyle = Get-FirstMatch $Block '<div[^>]*class="[^"]*\bimageComponent\b[^"]*"[^>]*style="([^"]*)"'
        if (-not $componentStyle) {
            $componentStyle = Get-FirstMatch $Block '<img[^>]*style="([^"]*)"'
        }
        $fallbackWidth = Get-ScalarDouble (Get-CssNumber $componentStyle 'max-width' 40) 40
        $fallbackHeight = Get-ScalarDouble (Get-CssNumber $componentStyle 'max-height' 40) 40
        return [pscustomobject]@{
            Width = Get-ScalarDouble (Get-CssNumber $componentStyle 'width' $fallbackWidth) $fallbackWidth
            Height = Get-ScalarDouble (Get-CssNumber $componentStyle 'height' $fallbackHeight) $fallbackHeight
        }
    }

    function Get-ImageMimeType {
        param([string]$Base64Data, [string]$Default = 'image/png')
        $usableLength = [Math]::Min(64, $Base64Data.Length - ($Base64Data.Length % 4))
        if ($usableLength -le 0) { return $Default }
        try {
            $bytes = [Convert]::FromBase64String($Base64Data.Substring(0, $usableLength))
        } catch {
            return $Default
        }
        if ($bytes.Length -ge 4 -and $bytes[0] -eq 0x89 -and $bytes[1] -eq 0x50 -and $bytes[2] -eq 0x4e -and $bytes[3] -eq 0x47) { return 'image/png' }
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xd8 -and $bytes[2] -eq 0xff) { return 'image/jpeg' }
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0x47 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46) { return 'image/gif' }
        if ($bytes.Length -ge 2 -and $bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4d) { return 'image/bmp' }
        if ($bytes.Length -ge 4 -and $bytes[0] -eq 0x52 -and $bytes[1] -eq 0x49 -and $bytes[2] -eq 0x46 -and $bytes[3] -eq 0x46) { return 'image/webp' }
        return $Default
    }

    function Get-DownscaledImagePayload {
        param([string]$Base64Payload, [double]$DisplayWidth, [double]$DisplayHeight)
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        } catch {
            return $null
        }
        try {
            $bytes = [Convert]::FromBase64String($Base64Payload)
        } catch {
            return $null
        }
        $capWidth = [Math]::Max(64.0, [Math]::Ceiling($DisplayWidth * 2))
        $capHeight = [Math]::Max(64.0, [Math]::Ceiling($DisplayHeight * 2))
        $ms = New-Object IO.MemoryStream(,$bytes)
        try {
            $img = [System.Drawing.Image]::FromStream($ms)
        } catch {
            $ms.Dispose()
            return $null
        }
        if ($img.Width -le $capWidth -and $img.Height -le $capHeight) {
            $img.Dispose(); $ms.Dispose()
            return $null
        }
        $scale = [Math]::Min($capWidth / $img.Width, $capHeight / $img.Height)
        $newWidth = [Math]::Max(1.0, [Math]::Round($img.Width * $scale))
        $newHeight = [Math]::Max(1.0, [Math]::Round($img.Height * $scale))
        $bmp = New-Object System.Drawing.Bitmap([int]$newWidth, [int]$newHeight)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($img, 0, 0, [int]$newWidth, [int]$newHeight)
        $outMs = New-Object IO.MemoryStream
        $bmp.Save($outMs, [System.Drawing.Imaging.ImageFormat]::Png)
        $result = [Convert]::ToBase64String($outMs.ToArray())
        $g.Dispose(); $bmp.Dispose(); $img.Dispose(); $outMs.Dispose(); $ms.Dispose()
        return $result
    }

    function Get-TextLineCount {
        # Rough line count for sizing a text box: each paragraph wraps on its own.
        param([string]$Text, [double]$Width, [double]$FontSize)
        $charsPerLine = [Math]::Max(1.0, [Math]::Floor($Width / ($FontSize * 0.58)))
        $count = 0
        foreach ($paragraph in ($Text -split "`n")) { $count += [Math]::Max(1.0, [Math]::Ceiling($paragraph.Length / $charsPerLine)) }
        return [Math]::Max(1.0, $count)
    }

    function Get-FontWeight {
        # CSS font-weight from an inline style: numbers as-is, normal/bold as 400/700.
        param([string]$Style, [int]$Default = 400)
        $v = "$(Get-StyleValue $Style 'font-weight')".Trim()
        if ($v -match '^\d+$') { return [int]$v }
        if ($v -eq 'bold') { return 700 }
        if ($v -eq 'normal') { return 400 }
        return $Default
    }

    function Get-StrokeWidthPx {
        # SVG stroke-width as px: Whiteboard writes shapes in pt ("2pt") and connectors unitless.
        param([string]$Value, [double]$Default = 2)
        if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
        $m = [regex]::Match($Value.Trim(), '^([0-9.]+)\s*(px|pt)?$', 'IgnoreCase')
        if (-not $m.Success) { return $Default }
        $n = Get-Number $m.Groups[1].Value $Default
        if ($m.Groups[2].Value -eq 'pt') { $n = $n * 96.0 / 72.0 }
        return $n
    }

    function Get-DashArray {
        # stroke-dasharray as a list of px lengths ("10.6667, 8"); $null for none/solid.
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Trim() -eq 'none') { return $null }
        $parts = @($Value.Trim() -split '[\s,]+' | Where-Object { $_ } | ForEach-Object { Get-Number $_ 0 })
        if ($parts.Count -eq 0 -or @($parts | Where-Object { $_ -gt 0 }).Count -eq 0) { return $null }
        return ,([double[]]$parts)
    }

    function Get-RgbTriplet {
        # rgb()/rgba() or #rgb/#rrggbb -> [r, g, b]; $null for anything else.
        param([string]$Color)
        $m = [regex]::Match($Color, 'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        # Unary comma keeps the three values together as one return object.
        if ($m.Success) { return ,([int[]]@([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value)) }
        $m = [regex]::Match($Color, '#([0-9a-f]{6}|[0-9a-f]{3})\b', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $m.Success) { return $null }
        $hex = $m.Groups[1].Value
        if ($hex.Length -eq 3) { $hex = -join @($hex.ToCharArray() | ForEach-Object { "$_$_" }) }
        return ,([int[]]@([Convert]::ToInt32($hex.Substring(0, 2), 16), [Convert]::ToInt32($hex.Substring(2, 2), 16),
                          [Convert]::ToInt32($hex.Substring(4, 2), 16)))
    }

    function ConvertTo-HexColor {
        param([int[]]$Rgb)
        return '#{0:x2}{1:x2}{2:x2}' -f $Rgb[0], $Rgb[1], $Rgb[2]
    }

    function Get-LinearGradient {
        # The first and last colour and the direction of the first CSS linear-gradient in
        # $Style. Angle is in CSS degrees (0 = to top, 180 = to bottom, the default). A
        # "to <corner>" gradient's angle depends on the box, so it is kept as Corner (x, y),
        # with right/bottom = +1.
        param([string]$Style)
        $gradient = Get-FirstMatch $Style 'linear-gradient\(((?:[^()]|\([^()]*\))*)\)'
        if (-not $gradient) { return $null }
        $colors = [regex]::Matches($gradient, 'rgba?\([^)]*\)|#[0-9a-f]{3,8}\b',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($colors.Count -lt 2) { return $null }
        $start = Get-RgbTriplet $colors[0].Value
        $end = Get-RgbTriplet $colors[$colors.Count - 1].Value
        if ($null -eq $start -or $null -eq $end) { return $null }
        $angle = 180.0; $corner = $null
        $first = ($gradient -split ',')[0].Trim().ToLowerInvariant()
        $m = [regex]::Match($first, '^([-+]?[0-9.]+)(deg|grad|rad|turn)$')
        if ($m.Success) {
            $v = Get-Number $m.Groups[1].Value 0
            $angle = switch ($m.Groups[2].Value) { 'grad' { $v * 0.9 } 'rad' { $v * 180.0 / [Math]::PI } 'turn' { $v * 360.0 } default { $v } }
        } elseif ($first -match '^to\s+(top|bottom|left|right)(?:\s+(top|bottom|left|right))?$') {
            $sides = @(@($Matches[1], $Matches[2]) | Where-Object { $_ })
            if ($sides.Count -eq 1) {
                $angle = @{ top = 0.0; right = 90.0; bottom = 180.0; left = 270.0 }[$sides[0]]
            } else {
                $cx = if ($sides -contains 'right') { 1.0 } else { -1.0 }
                $cy = if ($sides -contains 'bottom') { 1.0 } else { -1.0 }
                $corner = [double[]]@($cx, $cy)
            }
        }
        return [pscustomobject]@{ Start = $start; End = $end; Angle = $angle; Corner = $corner }
    }

    function Get-NoteColorTable {
        # Colours of Whiteboard's newer note style, keyed by colour name ("paleGreen"), from the
        # export's light-theme ":root { --noteBackgroundPaleGreenGradient: linear-gradient(...);
        # --noteBackgroundPaleGreenColor: none; ... }" (".darkTheme:root" and ".highContrast:root"
        # hold other themes). Solid is the Color property when it is a colour, else the
        # gradient's end colour -- the colour the export's own forced-colors fallback uses.
        param([string]$Html)
        $opts = [Text.RegularExpressions.RegexOptions]::IgnoreCase
        $table = @{}
        foreach ($root in [regex]::Matches($Html, '(?:^|[};>])\s*:root\s*\{([^}]*)\}', $opts)) {
            foreach ($m in [regex]::Matches($root.Groups[1].Value, '--noteBackground([A-Za-z]+?)(Gradient|Color)\s*:\s*([^;]+)', $opts)) {
                $name = $m.Groups[1].Value
                if (-not $table.ContainsKey($name)) { $table[$name] = @{ Gradient = $null; Solid = $null } }
                if ($m.Groups[2].Value -eq 'Gradient') {
                    $g = Get-LinearGradient $m.Groups[3].Value
                    if ($null -ne $g) { $table[$name].Gradient = $g }
                } else {
                    $rgb = Get-RgbTriplet $m.Groups[3].Value
                    if ($null -ne $rgb) { $table[$name].Solid = ConvertTo-HexColor $rgb }
                }
            }
        }
        foreach ($entry in $table.Values) {
            if (-not $entry.Solid -and $null -ne $entry.Gradient) { $entry.Solid = ConvertTo-HexColor $entry.Gradient.End }
        }
        return $table
    }

    function Get-NoteLayout {
        # Box size, text inset and colours of one sticky note, in its local px. Classic notes
        # have a 1px border outside their CSS size and inline background colours. Notes in the
        # newer style ("noteVisualUpdate") have no border, may have an author title bar above
        # the text area, and name their colour by class ("paleGreenGradient"), looked up in
        # $NoteColors (Get-NoteColorTable). ColorFound is $false when the default fill is used.
        param([string]$Block, [hashtable]$NoteColors)
        $noteStyle = Get-FirstMatch $Block '<div[^>]*class="[^"]*\btextbox\s+stickyNote\b[^"]*"[^>]*style="([^"]*)"'
        $bgTag = [string](Get-FirstMatch $Block '(<div[^>]*class="[^"]*\btextBoxBackground\b[^"]*"[^>]*>)')
        $bgClass = [string](Get-FirstMatch $bgTag 'class="([^"]*)"')
        $bgStyle = [string](Get-FirstMatch $bgTag 'style="([^"]*)"')
        $cssW = Get-CssNumber $noteStyle 'width' 304; $cssH = Get-CssNumber $noteStyle 'height' 304
        $border = if ($bgClass -match '\bnoteVisualUpdate\b') { 0.0 } else { [double]$NoteBorder }
        $titleBar = if ($Block -match 'class="[^"]*\btitleBarWithAttribution\b') { [double]$NoteTitleBarHeight } else { 0.0 }
        $solid = $null; $gradient = $null
        $inline = Get-StyleValue $bgStyle 'background-color'
        if ($inline) {
            $solid = Convert-RgbaToHex $inline $DefaultNoteFill
            $gradient = Get-LinearGradient $bgStyle
        } else {
            $name = Get-FirstMatch $bgClass '\b([a-z]+)Gradient\b'
            if ($name -and $NoteColors.ContainsKey($name)) {
                $solid = $NoteColors[$name].Solid; $gradient = $NoteColors[$name].Gradient
            }
        }
        return [pscustomobject]@{
            CssWidth = $cssW; CssHeight = $cssH
            Width = $cssW + (2 * $border); Height = $cssH + $titleBar + (2 * $border)
            # The 12px side padding plus Draft.js's 1px gutter, inside the border and title bar.
            TextLeft = $border + ($NoteTextInsetLeft - $NoteBorder); TextTop = $border + $titleBar + ($NoteTextInsetTop - $NoteBorder)
            Border = $border
            Solid = $(if ($solid) { $solid } else { $DefaultNoteFill }); Gradient = $gradient; ColorFound = [bool]$solid
        }
    }

    function Get-ClippedPolygon {
        # The part of convex polygon $Points whose distance along the unit vector ($DX, $DY),
        # measured from ($CX, $CY), lies between $From and $To (Sutherland-Hodgman, two edges).
        param([double[][]]$Points, [double]$CX, [double]$CY, [double]$DX, [double]$DY, [double]$From, [double]$To)
        $poly = $Points
        foreach ($edge in @(@(1.0, $From), @(-1.0, (-$To)))) {
            $out = New-Object Collections.Generic.List[double[]]
            $n = $poly.Count
            for ($i = 0; $i -lt $n; $i++) {
                $p = $poly[$i]; $q = $poly[($i + 1) % $n]
                $sp = ($edge[0] * ((($p[0] - $CX) * $DX) + (($p[1] - $CY) * $DY))) - $edge[1]
                $sq = ($edge[0] * ((($q[0] - $CX) * $DX) + (($q[1] - $CY) * $DY))) - $edge[1]
                if ($sp -ge 0) { [void]$out.Add($p) }
                if (($sp -ge 0) -ne ($sq -ge 0)) {
                    $k = $sp / ($sp - $sq)
                    [void]$out.Add([double[]]@(($p[0] + (($q[0] - $p[0]) * $k)), ($p[1] + (($q[1] - $p[1]) * $k))))
                }
            }
            $poly = $out.ToArray()
            if ($poly.Count -eq 0) { break }
        }
        return ,$poly
    }

    function Get-GradientBands {
        # Splits a $Width x $Height box (local px, origin at its corner) into $Steps colour
        # bands across a CSS linear-gradient (Get-LinearGradient): band i covers gradient
        # positions i/Steps .. (i+1)/Steps, clipped to the box, and is coloured from the start
        # colour (i = 0) to the end colour. As CSS defines it, the gradient line runs through
        # the box centre at the gradient's angle and is |W sin a| + |H cos a| long. Each band
        # but the last reaches half a band into the next, which is drawn over it: a thinner
        # overlap (0.15 px) let anti-aliasing show white seams along angled band edges. At 180 degrees (classic notes) the bands are plain horizontal strips; newer
        # notes use 155 degrees.
        param([double]$Width, [double]$Height, $Gradient, [int]$Steps)
        if ($null -ne $Gradient.Corner) {
            # "to <corner>": perpendicular to the diagonal through the two other corners.
            $dx = $Gradient.Corner[0] * $Height; $dy = $Gradient.Corner[1] * $Width
        } else {
            $rad = $Gradient.Angle * [Math]::PI / 180.0
            $dx = [Math]::Sin($rad); $dy = -[Math]::Cos($rad)
        }
        $len = [Math]::Sqrt(($dx * $dx) + ($dy * $dy)); $dx = $dx / $len; $dy = $dy / $len
        # sin(180 deg) is 1.2e-16, not 0: snap it so axis-aligned bands stay exact rectangles.
        if ([Math]::Abs($dx) -lt 1e-9) { $dx = 0.0 }
        if ([Math]::Abs($dy) -lt 1e-9) { $dy = 0.0 }
        $lineLength = [Math]::Abs($Width * $dx) + [Math]::Abs($Height * $dy)
        $box = [double[][]]@(@(0.0, 0.0), @($Width, 0.0), @($Width, $Height), @(0.0, $Height))
        $bands = New-Object Collections.Generic.List[object]
        for ($i = 0; $i -lt $Steps; $i++) {
            $ratio = if ($Steps -le 1) { 0.0 } else { [double]$i / [double]($Steps - 1) }
            $rgb = for ($c = 0; $c -lt 3; $c++) {
                [int][Math]::Round($Gradient.Start[$c] + (($Gradient.End[$c] - $Gradient.Start[$c]) * $ratio))
            }
            $from = (([double]$i / $Steps) - 0.5) * $lineLength
            $to = ((([double]$i + 1) / $Steps) - 0.5) * $lineLength
            if ($i -lt $Steps - 1) { $to += 0.5 * $lineLength / $Steps }
            $poly = Get-ClippedPolygon $box ($Width / 2) ($Height / 2) $dx $dy $from $to
            if ($poly.Count -ge 3) { [void]$bands.Add([pscustomobject]@{ Points = $poly; Color = (ConvertTo-HexColor $rgb) }) }
        }
        return ,($bands.ToArray())
    }

    function Get-ConnectorPoints {
        # The connector's line -- the first <path d> in its <g> -- as points in the anchor's
        # local px. A straight connector is one M/L pair; an elbow connector adds corners
        # (L/H/V) and a curved one Bezier segments (C/S/Q/T), which are flattened at about one
        # point per 4 px. Only the first two number pairs used to be read, so a bent or curved
        # connector came out as a straight line to its first corner. Exact is $false when the
        # path holds something not traced here: an arc (A, drawn as a straight segment to its
        # end point), a second subpath (joined on), or a malformed command (the rest is
        # dropped). A zero-length line has a single point; $null means no path at all.
        param([string]$Block)
        $d = Get-FirstMatch $Block '<path\s+d="([^"]+)"'
        if (-not $d) { return $null }
        $tokens = @([regex]::Matches($d, '[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?') | ForEach-Object { $_.Value })
        $argCount = @{ M = 2; L = 2; H = 1; V = 1; C = 6; S = 4; Q = 4; T = 2; A = 7; Z = 0 }
        $pts = New-Object Collections.Generic.List[double[]]
        $exact = $true; $i = 0; $cmd = ''; $prev = ''; $subpaths = 0
        $cx = 0.0; $cy = 0.0; $sx = 0.0; $sy = 0.0; $qx = 0.0; $qy = 0.0   # current, subpath start, last control point
        while ($i -lt $tokens.Count) {
            if ($tokens[$i] -match '^[A-Za-z]$') { $cmd = $tokens[$i]; $i++ }
            elseif (-not $cmd) { $exact = $false; break }
            $up = $cmd.ToUpperInvariant()
            if (-not $argCount.ContainsKey($up)) { $exact = $false; break }
            $n = $argCount[$up]
            $v = New-Object double[] $n
            $ok = $true
            for ($k = 0; $k -lt $n; $k++) {
                if ($i + $k -ge $tokens.Count -or $tokens[$i + $k] -match '^[A-Za-z]$') { $ok = $false; break }
                $v[$k] = Get-Number $tokens[$i + $k] 0
            }
            if (-not $ok) { $exact = $false; break }
            $i += $n
            $ox = if ($cmd -ceq $up) { 0.0 } else { $cx }; $oy = if ($cmd -ceq $up) { 0.0 } else { $cy }
            $bx = $null; $by = $null   # Bezier control polygon, starting at the current point
            switch ($up) {
                'M' {
                    $subpaths++
                    if ($subpaths -gt 1) { $exact = $false }
                    $cx = $ox + $v[0]; $cy = $oy + $v[1]; $sx = $cx; $sy = $cy
                    [void]$pts.Add([double[]]@($cx, $cy))
                    $cmd = if ($cmd -ceq 'M') { 'L' } else { 'l' }   # extra pairs after M are line-tos
                }
                'L' { $cx = $ox + $v[0]; $cy = $oy + $v[1]; [void]$pts.Add([double[]]@($cx, $cy)) }
                'H' { $cx = $ox + $v[0]; [void]$pts.Add([double[]]@($cx, $cy)) }
                'V' { $cy = $oy + $v[0]; [void]$pts.Add([double[]]@($cx, $cy)) }
                'C' {
                    $bx = @($cx, ($ox + $v[0]), ($ox + $v[2]), ($ox + $v[4]))
                    $by = @($cy, ($oy + $v[1]), ($oy + $v[3]), ($oy + $v[5]))
                }
                'S' {
                    $rx = $cx; $ry = $cy   # first control point: the previous one reflected
                    if ($prev -eq 'C' -or $prev -eq 'S') { $rx = (2 * $cx) - $qx; $ry = (2 * $cy) - $qy }
                    $bx = @($cx, $rx, ($ox + $v[0]), ($ox + $v[2]))
                    $by = @($cy, $ry, ($oy + $v[1]), ($oy + $v[3]))
                }
                'Q' {
                    $bx = @($cx, ($ox + $v[0]), ($ox + $v[2]))
                    $by = @($cy, ($oy + $v[1]), ($oy + $v[3]))
                }
                'T' {
                    $rx = $cx; $ry = $cy
                    if ($prev -eq 'Q' -or $prev -eq 'T') { $rx = (2 * $cx) - $qx; $ry = (2 * $cy) - $qy }
                    $bx = @($cx, $rx, ($ox + $v[0]))
                    $by = @($cy, $ry, ($oy + $v[1]))
                }
                'A' { $exact = $false; $cx = $ox + $v[5]; $cy = $oy + $v[6]; [void]$pts.Add([double[]]@($cx, $cy)) }
                'Z' { $cx = $sx; $cy = $sy; [void]$pts.Add([double[]]@($cx, $cy)); $cmd = '' }
            }
            if ($null -ne $bx) {
                $deg = $bx.Count - 1
                $qx = $bx[$deg - 1]; $qy = $by[$deg - 1]
                $hull = 0.0
                for ($k = 1; $k -le $deg; $k++) { $hull += [Math]::Sqrt([Math]::Pow($bx[$k] - $bx[$k - 1], 2) + [Math]::Pow($by[$k] - $by[$k - 1], 2)) }
                $segments = [int][Math]::Min(64.0, [Math]::Max(4.0, [Math]::Ceiling($hull / 4.0)))
                for ($k = 1; $k -le $segments; $k++) {
                    $u = $k / $segments; $w = 1 - $u
                    if ($deg -eq 2) {
                        $px = ($w * $w * $bx[0]) + (2 * $w * $u * $bx[1]) + ($u * $u * $bx[2])
                        $py = ($w * $w * $by[0]) + (2 * $w * $u * $by[1]) + ($u * $u * $by[2])
                    } else {
                        $px = ($w * $w * $w * $bx[0]) + (3 * $w * $w * $u * $bx[1]) + (3 * $w * $u * $u * $bx[2]) + ($u * $u * $u * $bx[3])
                        $py = ($w * $w * $w * $by[0]) + (3 * $w * $w * $u * $by[1]) + (3 * $w * $u * $u * $by[2]) + ($u * $u * $u * $by[3])
                    }
                    [void]$pts.Add([double[]]@($px, $py))
                }
                $cx = $bx[$deg]; $cy = $by[$deg]
            }
            $prev = $up
        }
        $clean = New-Object Collections.Generic.List[double[]]   # drop zero-length steps
        foreach ($p in $pts) {
            if ($clean.Count -gt 0) {
                $last = $clean[$clean.Count - 1]
                if ([Math]::Abs($p[0] - $last[0]) -lt 0.001 -and [Math]::Abs($p[1] - $last[1]) -lt 0.001) { continue }
            }
            [void]$clean.Add($p)
        }
        if ($clean.Count -eq 0) { return $null }
        return [pscustomobject]@{ Points = $clean.ToArray(); Exact = $exact }
    }

    function Get-InkStrokes {
        # The strokes of an InkGroup (freehand ink), in board px. Whiteboard draws each stroke
        # as a filled outline <path> inside <g class="inkStroke" transform="matrix(...)"> (in
        # 1/128 px), and also stores its centreline as the invisible hit-test
        # <polyline class="inkHitTestOverlay">. The outline is that centreline widened by the
        # pen radius -- its round joins and caps are arcs of exactly that radius -- so the
        # centreline drawn round-capped at twice the arc radius is the same stroke. Exact is
        # $false when that doesn't hold: no arcs, arcs of varying radius (a pressure-sensitive
        # pen), or a fill that isn't a plain opaque colour (e.g. the rainbow pen).
        param($Anchor, [string]$Block)
        $opts = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
        $numRx = '[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?'
        $svgTag = [string](Get-FirstMatch $Block '(<svg\b[^>]*>)')
        $vb = @([regex]::Matches([string](Get-FirstMatch $svgTag '\bviewBox="([^"]*)"'), $numRx) | ForEach-Object { Get-Number $_.Value 0 })
        $sx = 1.0; $sy = 1.0; $vx = 0.0; $vy = 0.0
        if ($vb.Count -eq 4 -and $vb[2] -gt 0 -and $vb[3] -gt 0) {
            $vx = $vb[0]; $vy = $vb[1]
            $sx = (Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') $vb[2]) / $vb[2]
            $sy = (Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') $vb[3]) / $vb[3]
        }
        $strokes = New-Object Collections.Generic.List[object]
        foreach ($g in [regex]::Matches($Block, '(<g\b[^>]*\bclass="[^"]*\binkStroke\b[^"]*"[^>]*>)(.*?)</g>', $opts)) {
            $m = @(1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
            $gm = @([regex]::Matches([string](Get-FirstMatch $g.Groups[1].Value '\btransform="\s*matrix\(([^)]*)\)'), $numRx) | ForEach-Object { Get-Number $_.Value 0 })
            if ($gm.Count -eq 6) { $m = $gm }
            $pathTag = [string](Get-FirstMatch $g.Groups[2].Value '(<path\b[^>]*>)')
            $d = [string](Get-FirstMatch $pathTag '\sd="([^"]*)"')
            $nums = @([regex]::Matches([string](Get-FirstMatch $g.Groups[2].Value '<polyline\b[^>]*\bpoints="([^"]*)"'), $numRx) | ForEach-Object { Get-Number $_.Value 0 })
            if ($nums.Count -lt 4) { continue }
            $pts = New-Object Collections.Generic.List[double[]]
            for ($k = 0; $k + 1 -lt $nums.Count; $k += 2) {
                # g matrix, then the svg viewBox, then the anchor's own matrix.
                $u = ((($m[0] * $nums[$k]) + ($m[2] * $nums[$k + 1]) + $m[4]) - $vx) * $sx
                $v = ((($m[1] * $nums[$k]) + ($m[3] * $nums[$k + 1]) + $m[5]) - $vy) * $sy
                [void]$pts.Add([double[]]@(($Anchor.X + ($Anchor.MA * $u) + ($Anchor.MC * $v)),
                                           ($Anchor.Y + ($Anchor.MB * $u) + ($Anchor.MD * $v))))
            }
            $radii = @([regex]::Matches($d, "[Aa]\s*($numRx)[\s,]+($numRx)") | ForEach-Object { Get-Number $_.Groups[1].Value 0 } | Sort-Object)
            $exact = $true
            if ($radii.Count -gt 0) {
                $radius = $radii[[int][Math]::Floor($radii.Count / 2)]
                if (($radii[$radii.Count - 1] - $radii[0]) -gt (0.05 * $radius)) { $exact = $false }
            } else {
                $radius = 256.0; $exact = $false   # 2 px at Whiteboard's 1/128 px scale
            }
            $scale = [Math]::Sqrt([Math]::Abs((($m[0] * $m[3]) - ($m[1] * $m[2])) * $sx * $sy *
                                              (($Anchor.MA * $Anchor.MD) - ($Anchor.MB * $Anchor.MC))))
            $fill = [string](Get-FirstMatch $pathTag '\bfill="([^"]*)"')
            $alpha = Get-FirstMatch $fill 'rgba\([^,]+,[^,]+,[^,]+,\s*([0-9.]+)\s*\)'
            $opacity = Get-FirstMatch $pathTag '\bopacity="([0-9.]+)"'
            if ($fill -match '^\s*url\(' -or -not (Get-RgbTriplet $fill) -or ($alpha -and (Get-Number $alpha 1) -lt 1) -or ($opacity -and (Get-Number $opacity 1) -lt 1)) {
                $exact = $false
            }
            [void]$strokes.Add([pscustomobject]@{
                Points = $pts.ToArray(); Width = 2 * $radius * $scale
                Color = (Convert-RgbaToHex $fill '#1f1f1f'); Exact = $exact
            })
        }
        return ,($strokes.ToArray())
    }

    # ===================================================================================
    # Geometry helpers
    # ===================================================================================

    function ConvertTo-BoardPoints {
        # Maps points in an anchor's local (pre-transform) px through its full CSS matrix
        # (transform-origin 0 0), so rotated and mirrored objects keep their orientation.
        param($Anchor, [double[][]]$Points)
        $out = New-Object Collections.Generic.List[double[]]
        foreach ($p in $Points) {
            [void]$out.Add([double[]]@(($Anchor.X + ($Anchor.MA * $p[0]) + ($Anchor.MC * $p[1])),
                                       ($Anchor.Y + ($Anchor.MB * $p[0]) + ($Anchor.MD * $p[1]))))
        }
        return ,($out.ToArray())
    }

    function Get-RectPoints { param($X, $Y, $W, $H)
        $pts = New-Object Collections.Generic.List[double[]]
        [void]$pts.Add(@($X, $Y))
        [void]$pts.Add(@(($X + $W), $Y))
        [void]$pts.Add(@(($X + $W), ($Y + $H)))
        [void]$pts.Add(@($X, ($Y + $H)))
        return $pts.ToArray()
    }
    function Get-DiamondPoints { param($X, $Y, $W, $H)
        $pts = New-Object Collections.Generic.List[double[]]
        [void]$pts.Add(@(($X + $W/2), $Y))
        [void]$pts.Add(@(($X + $W), ($Y + $H/2)))
        [void]$pts.Add(@(($X + $W/2), ($Y + $H)))
        [void]$pts.Add(@($X, ($Y + $H/2)))
        return $pts.ToArray()
    }
    function Get-EllipsePoints { param($X, $Y, $W, $H, [int]$Segments = 40)
        $cx = $X + $W/2; $cy = $Y + $H/2; $rx = $W/2; $ry = $H/2
        $pts = New-Object Collections.Generic.List[double[]]
        for ($i = 0; $i -lt $Segments; $i++) {
            $theta = 2 * [Math]::PI * $i / $Segments
            [void]$pts.Add(@(($cx + $rx * [Math]::Cos($theta)), ($cy + $ry * [Math]::Sin($theta))))
        }
        return $pts.ToArray()
    }

    function Get-ShapePathPoints {
        # Traces the shape's own outline from Whiteboard's SVG <path d="..."> (the first,
        # visible path inside the shape's <g>), flattening curves into points. Returns $null
        # (so the caller falls back to the label-based builders) for arcs (A), smooth
        # shorthands (S/T), more than one subpath, or a <g> transform other than a translate.
        param([string]$Block, [double]$X, [double]$Y, [double]$ScaleX, [double]$ScaleY, [int]$CurveSegments = 8)
        $opts = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
        $m = [regex]::Match($Block, '<svg\b[^>]*\bclass="[^"]*\bshape\b[^"]*"[^>]*>\s*(<g\b[^>]*>)\s*<path\b[^>]*?\sd="([^"]+)"', $opts)
        if (-not $m.Success) { return $null }
        $gTag = $m.Groups[1].Value; $d = $m.Groups[2].Value
        $tx = 0.0; $ty = 0.0
        $gTransform = Get-FirstMatch $gTag '\btransform="([^"]*)"'
        if ($gTransform) {
            $t = [regex]::Match($gTransform, '^\s*translate\(\s*([-+0-9.eE]+)(?:[\s,]+([-+0-9.eE]+))?\s*\)\s*$')
            if (-not $t.Success) { return $null }
            $tx = Get-Number $t.Groups[1].Value 0
            if ($t.Groups[2].Success) { $ty = Get-Number $t.Groups[2].Value 0 }
        }
        $tokens = @([regex]::Matches($d, '[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?') | ForEach-Object { $_.Value })
        $pts = New-Object Collections.Generic.List[double[]]
        $i = 0; $cmd = ''; $cx = 0.0; $cy = 0.0; $sx = 0.0; $sy = 0.0; $subpaths = 0
        $isNum = { param($tok) $tok -notmatch '^[A-Za-z]$' }
        while ($i -lt $tokens.Count) {
            if (-not (& $isNum $tokens[$i])) { $cmd = $tokens[$i]; $i++ }
            elseif (-not $cmd) { return $null }
            $rel = ($cmd -ceq $cmd.ToLowerInvariant())
            $ox = if ($rel) { $cx } else { 0.0 }; $oy = if ($rel) { $cy } else { 0.0 }
            switch -CaseSensitive ($cmd.ToUpperInvariant()) {
                'M' {
                    if ($i + 1 -ge $tokens.Count) { return $null }
                    $subpaths++
                    if ($subpaths -gt 1) { return $null }
                    $cx = $ox + (Get-Number $tokens[$i] 0); $cy = $oy + (Get-Number $tokens[$i + 1] 0); $i += 2
                    $sx = $cx; $sy = $cy
                    [void]$pts.Add([double[]]@($cx, $cy))
                    $cmd = if ($rel) { 'l' } else { 'L' }   # extra pairs after M are line-tos
                }
                'L' {
                    if ($i + 1 -ge $tokens.Count) { return $null }
                    $cx = $ox + (Get-Number $tokens[$i] 0); $cy = $oy + (Get-Number $tokens[$i + 1] 0); $i += 2
                    [void]$pts.Add([double[]]@($cx, $cy))
                }
                'H' { $cx = $ox + (Get-Number $tokens[$i] 0); $i += 1; [void]$pts.Add([double[]]@($cx, $cy)) }
                'V' { $cy = $oy + (Get-Number $tokens[$i] 0); $i += 1; [void]$pts.Add([double[]]@($cx, $cy)) }
                'Q' {
                    if ($i + 3 -ge $tokens.Count) { return $null }
                    $x1 = $ox + (Get-Number $tokens[$i] 0);     $y1 = $oy + (Get-Number $tokens[$i + 1] 0)
                    $x2 = $ox + (Get-Number $tokens[$i + 2] 0); $y2 = $oy + (Get-Number $tokens[$i + 3] 0); $i += 4
                    for ($s = 1; $s -le $CurveSegments; $s++) {
                        $u = $s / $CurveSegments; $v = 1 - $u
                        [void]$pts.Add([double[]]@((($v * $v * $cx) + (2 * $v * $u * $x1) + ($u * $u * $x2)),
                                                   (($v * $v * $cy) + (2 * $v * $u * $y1) + ($u * $u * $y2))))
                    }
                    $cx = $x2; $cy = $y2
                }
                'C' {
                    if ($i + 5 -ge $tokens.Count) { return $null }
                    $x1 = $ox + (Get-Number $tokens[$i] 0);     $y1 = $oy + (Get-Number $tokens[$i + 1] 0)
                    $x2 = $ox + (Get-Number $tokens[$i + 2] 0); $y2 = $oy + (Get-Number $tokens[$i + 3] 0)
                    $x3 = $ox + (Get-Number $tokens[$i + 4] 0); $y3 = $oy + (Get-Number $tokens[$i + 5] 0); $i += 6
                    for ($s = 1; $s -le $CurveSegments; $s++) {
                        $u = $s / $CurveSegments; $v = 1 - $u
                        [void]$pts.Add([double[]]@((($v * $v * $v * $cx) + (3 * $v * $v * $u * $x1) + (3 * $v * $u * $u * $x2) + ($u * $u * $u * $x3)),
                                                   (($v * $v * $v * $cy) + (3 * $v * $v * $u * $y1) + (3 * $v * $u * $u * $y2) + ($u * $u * $u * $y3))))
                    }
                    $cx = $x3; $cy = $y3
                }
                'Z' { $cx = $sx; $cy = $sy; $cmd = '' }
                default { return $null }   # A, S, T and anything unexpected -> label fallback
            }
        }
        if ($pts.Count -ge 2) {   # drop the closing duplicate; polygons auto-close
            $first = $pts[0]; $last = $pts[$pts.Count - 1]
            if ([Math]::Abs($first[0] - $last[0]) -lt 0.01 -and [Math]::Abs($first[1] - $last[1]) -lt 0.01) { $pts.RemoveAt($pts.Count - 1) }
        }
        if ($pts.Count -lt 3) { return $null }
        $board = New-Object Collections.Generic.List[double[]]
        foreach ($p in $pts) {
            [void]$board.Add([double[]]@(($X + (($p[0] + $tx) * $ScaleX)), ($Y + (($p[1] + $ty) * $ScaleY))))
        }
        return ,($board.ToArray())
    }

    function Get-AxisAlignedBox {
        # The bounding box of a 4-point outline whose edges are all horizontal or vertical
        # (an unrotated rectangle); $null for any other outline.
        param([double[][]]$Points)
        if ($Points.Count -ne 4) { return $null }
        for ($i = 0; $i -lt 4; $i++) {
            $p = $Points[$i]; $q = $Points[($i + 1) % 4]
            if ([Math]::Abs($p[0] - $q[0]) -gt 0.01 -and [Math]::Abs($p[1] - $q[1]) -gt 0.01) { return $null }
        }
        $xs = @($Points | ForEach-Object { $_[0] }); $ys = @($Points | ForEach-Object { $_[1] })
        $minX = ($xs | Measure-Object -Minimum).Minimum; $maxX = ($xs | Measure-Object -Maximum).Maximum
        $minY = ($ys | Measure-Object -Minimum).Minimum; $maxY = ($ys | Measure-Object -Maximum).Maximum
        if (($maxX - $minX) -lt 0.01 -or ($maxY - $minY) -lt 0.01) { return $null }
        return [pscustomobject]@{ X = $minX; Y = $minY; W = ($maxX - $minX); H = ($maxY - $minY) }
    }

    # ===================================================================================
    # IWB graphic descriptors, in board coordinates. New-IwbContentXml maps them onto the
    # page at the end. Group is a key shared by the parts of one Whiteboard object.
    # ===================================================================================

    function Add-IwbRect {
        param([Collections.Generic.List[object]]$Graphics, [double]$X, [double]$Y, [double]$W, [double]$H,
              [string]$Fill, [string]$Stroke, [double]$StrokeWidth = 0, [double[]]$Dash = $null, [string]$Group = $null)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Rect'; X = $X; Y = $Y; W = $W; H = $H; Fill = $Fill; Stroke = $Stroke
            StrokeWidth = $StrokeWidth; Dash = $Dash; Group = $Group
        })
    }

    function Add-IwbPolygon {
        param([Collections.Generic.List[object]]$Graphics, [double[][]]$Points,
              [string]$Fill, [string]$Stroke, [double]$StrokeWidth = 0, [double[]]$Dash = $null, [string]$Group = $null)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Polygon'; Points = $Points; Fill = $Fill; Stroke = $Stroke
            StrokeWidth = $StrokeWidth; Dash = $Dash; Group = $Group
        })
    }

    function Add-IwbLine {
        param([Collections.Generic.List[object]]$Graphics, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
              [string]$Stroke, [double]$StrokeWidth = 2, [double[]]$Dash = $null, [string]$Group = $null,
              [double]$DashOffset = 0)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Line'; X1 = $X1; Y1 = $Y1; X2 = $X2; Y2 = $Y2; Stroke = $Stroke
            StrokeWidth = $StrokeWidth; Dash = $Dash; DashOffset = $DashOffset; Group = $Group
        })
    }

    function Add-IwbPolyline {
        # An open or closed run of points as its individual edges (2-point polylines). A dashed
        # run gives each edge the distance travelled so far as its dash offset, so the pattern
        # continues round corners instead of restarting on every edge.
        param([Collections.Generic.List[object]]$Graphics, [double[][]]$Points, [switch]$Closed,
              [string]$Stroke, [double]$StrokeWidth = 2, [double[]]$Dash = $null, [string]$Group = $null)
        $n = $Points.Count
        $last = if ($Closed) { $n } else { $n - 1 }
        $travelled = 0.0
        for ($i = 0; $i -lt $last; $i++) {
            $p = $Points[$i]; $q = $Points[($i + 1) % $n]
            Add-IwbLine $Graphics $p[0] $p[1] $q[0] $q[1] $Stroke $StrokeWidth $Dash $Group $travelled
            $travelled += [Math]::Sqrt([Math]::Pow($q[0] - $p[0], 2) + [Math]::Pow($q[1] - $p[1], 2))
        }
    }

    function Add-IwbText {
        param(
            [Collections.Generic.List[object]]$Graphics, [string]$Text,
            [double]$X, [double]$Y, [double]$Width, [double]$Height,
            [double]$FontSize, [string]$Color, [string]$Align = 'left',
            # Optional unit rotation (no scale): the box's local x axis is (M11, M12) and its
            # local y axis is (M21, M22), in SVG matrix order.
            [double]$M11 = 1, [double]$M12 = 0, [double]$M21 = 0, [double]$M22 = 1,
            [int]$Weight = 400, [string]$Family = $TextFontFamily, [string]$Group = $null,
            [bool]$Underline = $false, [bool]$Italic = $false,
            # CSS line-height as a multiple of the font size (Get-LineHeight); 0 = normal.
            [double]$LineHeight = 0,
            # Runs with their own styling (Get-TextRuns), or $null for uniformly styled text.
            [object[]]$Runs = $null
        )
        if ([string]::IsNullOrEmpty($Text)) { return }
        $FontSize = Get-ScalarDouble $FontSize 20
        if ($Width -le 1) { $Width = [Math]::Max(1.0, [Math]::Min(4000.0, $Text.Length * $FontSize * 0.62)) }
        # A textarea shows only the lines that fit its height (SVG Tiny 1.2), and OpenBoard
        # squashes a text item whose content is taller than it, so err on the tall side.
        $lineFactor = [Math]::Max(1.35, $LineHeight)
        $needed = ((Get-TextLineCount $Text $Width $FontSize) + 1) * $FontSize * $lineFactor
        $Height = [Math]::Max($Height, $needed)
        # A set line height moves the first line by half the difference from the font's normal
        # line height (CSS half-leading), so the box moves with it along its own y axis; later
        # lines are spaced by line-increment. Segoe Print's normal line is 1.77 em, so at 140%
        # the text was 5.8px low.
        $lineIncrement = 0.0
        $metrics = if ($LineHeight -gt 0) { Get-FontLineMetrics $Family $Weight } else { $null }
        if ($null -ne $metrics) {
            $shift = ($LineHeight - $metrics.Normal) * $FontSize / 2
            $X += $M21 * $shift; $Y += $M22 * $shift
            $lineIncrement = $LineHeight * $FontSize
        }
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Text'; X = $X; Y = $Y; Width = $Width; Height = $Height
            FontSize = $FontSize; Color = $Color; Align = $Align; Family = $Family; Weight = $Weight
            Text = $Text; M11 = $M11; M12 = $M12; M21 = $M21; M22 = $M22; Group = $Group
            Underline = $Underline; Italic = $Italic; LineIncrement = $lineIncrement; Runs = $Runs
        })
    }

    function Get-InstalledFontNames {
        # Names of the font families installed here (empty where System.Drawing is missing).
        if ($null -eq $script:InstalledFontNames) {
            $set = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            try {
                Add-Type -AssemblyName System.Drawing -ErrorAction Stop
                foreach ($f in (New-Object Drawing.Text.InstalledFontCollection).Families) { [void]$set.Add($f.Name) }
            } catch { }
            $script:InstalledFontNames = $set
        }
        return ,$script:InstalledFontNames
    }

    function Get-FontFamily {
        # The face for an inline CSS font-family list: the first installed family, with
        # Whiteboard's web-font aliases mapped to the installed fonts' names ("AptosSerif" is
        # "Aptos Serif", "InkFreeFont" is "Ink Free") and generic families to Windows' faces.
        # Older exports' plain text is "sans-serif, "Segoe UI"", and the generic sans-serif
        # renders as Arial; newer exports (CompareAndContrast2) name "Segoe UI" itself, which is
        # wider -- drawing it in Arial moved centred labels up to 21px. Where the installed
        # fonts can't be listed, the first family is used.
        param([string]$Style, [string]$Default = 'Arial')
        $v = "$(Get-StyleValue ([Net.WebUtility]::HtmlDecode($Style)) 'font-family')"
        $aliases = @{ 'AptosSerif' = 'Aptos Serif'; 'AptosDisplay' = 'Aptos Display'; 'AptosMono' = 'Aptos Mono'
                      'AptosNarrow' = 'Aptos Narrow'; 'InkFreeFont' = 'Ink Free'; 'sans-serif' = $Default
                      'serif' = 'Times New Roman'; 'monospace' = 'Courier New' }
        $names = @(($v -split ',') | ForEach-Object { $_.Trim().Trim([char[]]@([char]34, [char]39)).Trim() } | Where-Object { $_ } |
            ForEach-Object { if ($aliases.ContainsKey($_)) { $aliases[$_] } else { $_ } })
        if ($names.Count -eq 0) { return $Default }
        $installed = Get-InstalledFontNames
        foreach ($n in $names) { if ($installed.Contains($n)) { return $n } }
        return $names[0]
    }

    function Get-LineHeight {
        # An inline CSS line-height as a multiple of the font size; 0 for "normal" or none.
        param([string]$Style, [double]$FontSize)
        $v = "$(Get-StyleValue $Style 'line-height')".Trim()
        if ($v -match '^([0-9.]+)%$') { return (Get-Number $Matches[1] 0) / 100.0 }
        if ($v -match '^([0-9.]+)px$' -and $FontSize -gt 0) { return (Get-Number $Matches[1] 0) / $FontSize }
        if ($v -match '^[0-9.]+$') { return Get-Number $v 0 }
        return 0.0
    }

    function Get-FontLineMetrics {
        # A face's line metrics as multiples of the font size: Normal is the browser's
        # "line-height: normal" (ascent + descent + line gap) and Content the ascent + descent,
        # which Qt's proportional line height is a percentage of. $null if the face isn't
        # installed or System.Drawing is missing.
        param([string]$Family, [int]$Weight = 400)
        $key = "$Family|$([int]($Weight -ge 600))"
        if (-not $script:FontLineMetrics.ContainsKey($key)) {
            $metrics = $null
            if ((Get-InstalledFontNames).Contains($Family)) {
                try {
                    $ff = New-Object Drawing.FontFamily $Family
                    $style = if ($Weight -ge 600 -and $ff.IsStyleAvailable([Drawing.FontStyle]::Bold)) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }
                    $em = [double]$ff.GetEmHeight($style)
                    $metrics = [pscustomobject]@{
                        Normal = $ff.GetLineSpacing($style) / $em
                        Content = ($ff.GetCellAscent($style) + $ff.GetCellDescent($style)) / $em
                    }
                    $ff.Dispose()
                } catch { $metrics = $null }
            }
            $script:FontLineMetrics[$key] = $metrics
        }
        return $script:FontLineMetrics[$key]
    }

    function Get-TextMeasurer {
        # A scriptblock giving a string's advance width in px in the given face (System.Drawing,
        # typographic layout). Without the face or System.Drawing, 0.58 em per character.
        param([string]$Family, [double]$FontSize, [int]$Weight = 400, [bool]$Italic = $false)
        $font = $null
        if ((Get-InstalledFontNames).Contains($Family)) {
            try {
                $style = [Drawing.FontStyle]::Regular
                if ($Weight -ge 600) { $style = $style -bor [Drawing.FontStyle]::Bold }
                if ($Italic) { $style = $style -bor [Drawing.FontStyle]::Italic }
                $font = New-Object Drawing.Font($Family, [single]$FontSize, $style, [Drawing.GraphicsUnit]::Pixel)
                if ($null -eq $script:MeasureGraphics) {
                    $script:MeasureGraphics = [Drawing.Graphics]::FromImage((New-Object Drawing.Bitmap 1, 1))
                    $script:MeasureGraphics.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAlias
                    $script:MeasureFormat = [Drawing.StringFormat]::GenericTypographic.Clone()
                    $script:MeasureFormat.FormatFlags = $script:MeasureFormat.FormatFlags -bor [Drawing.StringFormatFlags]::MeasureTrailingSpaces
                }
            } catch { $font = $null }
        }
        if ($null -eq $font) {
            $em = $FontSize * 0.58
            return { param([string]$s) $s.Length * $em }.GetNewClosure()
        }
        $g = $script:MeasureGraphics; $f = $script:MeasureFormat; $origin = New-Object Drawing.PointF 0, 0
        return { param([string]$s) if ($s.Length -eq 0) { 0.0 } else { [double]$g.MeasureString($s, $font, $origin, $f).Width } }.GetNewClosure()
    }

    function Get-TextLineStarts {
        # Where each line of a text box starts, as character offsets into $Text, wrapped as a
        # browser wraps it: each paragraph on its own, greedily at spaces (trailing spaces
        # hang), and a word wider than the column broken between characters.
        param([string]$Text, [double]$Width, [scriptblock]$Measure)
        $starts = New-Object Collections.Generic.List[int]
        $offset = 0
        foreach ($paragraph in ($Text -split "`n")) {
            $starts.Add($offset)
            $lineStart = 0; $pos = 0
            while ($pos -lt $paragraph.Length) {
                $wordEnd = $pos
                while ($wordEnd -lt $paragraph.Length -and $paragraph[$wordEnd] -ne ' ') { $wordEnd++ }
                if ((& $Measure $paragraph.Substring($lineStart, $wordEnd - $lineStart)) -le $Width + 0.01) {
                    $pos = $wordEnd
                    while ($pos -lt $paragraph.Length -and $paragraph[$pos] -eq ' ') { $pos++ }
                } elseif ($pos -gt $lineStart) {
                    $starts.Add($offset + $pos); $lineStart = $pos
                } else {
                    $end = $lineStart + 1
                    while ($end -lt $wordEnd -and (& $Measure $paragraph.Substring($lineStart, $end + 1 - $lineStart)) -le $Width + 0.01) { $end++ }
                    $starts.Add($offset + $end); $lineStart = $end; $pos = $end
                }
            }
            $offset += $paragraph.Length + 1
        }
        return ,($starts.ToArray())
    }

    function Select-TextRuns {
        # The runs cut to their first $Length characters.
        param([object[]]$Runs, [int]$Length)
        $kept = New-Object Collections.Generic.List[object]
        $left = $Length
        foreach ($r in $Runs) {
            if ($left -le 0) { break }
            $copy = $r.PSObject.Copy()
            if ($copy.Text.Length -gt $left) { $copy.Text = $copy.Text.Substring(0, $left) }
            $left -= $copy.Text.Length
            [void]$kept.Add($copy)
        }
        return ,($kept.ToArray())
    }

    function Test-Underline {
        # Whiteboard underlines a whole text box with text-decoration on its textBoxCore div.
        param([string]$Style)
        return "$(Get-StyleValue $Style 'text-decoration')" -match '\bunderline\b'
    }

    function Test-Italic {
        param([string]$Style)
        return "$(Get-StyleValue $Style 'font-style')" -match '\b(italic|oblique)\b'
    }

    function Add-IwbUnderline {
        # OpenBoard's IWB importer ignores text-decoration, so an underline is also drawn as a
        # line grouped with its text. Only for one unrotated line: the line is measured with
        # the same font, and wrapped lines can't be placed without the renderer's line breaks.
        # Returns $false when no line was drawn.
        param([Collections.Generic.List[object]]$Graphics, [string]$Text, [double]$X, [double]$Y,
              [double]$Width, [double]$FontSize, [string]$Color, [string]$Align, [int]$Weight,
              [string]$Family, [string]$Group)
        if ($Text.Contains("`n")) { return $false }
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $style = if ($Weight -ge 600) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }
            $font = New-Object Drawing.Font($Family, [single]$FontSize, $style, [Drawing.GraphicsUnit]::Pixel)
            $bmp = New-Object Drawing.Bitmap(1, 1)
            $gfx = [Drawing.Graphics]::FromImage($bmp)
            try { $textWidth = $gfx.MeasureString($Text, $font, [Drawing.PointF]::Empty, [Drawing.StringFormat]::GenericTypographic).Width }
            finally { $gfx.Dispose(); $bmp.Dispose(); $font.Dispose() }
        } catch { return $false }
        if ($textWidth -le 0 -or $textWidth -gt $Width) { return $false }
        $left = switch ($Align) { 'center' { $X + (($Width - $textWidth) / 2) } 'right' { $X + $Width - $textWidth } default { $X } }
        $lineY = $Y + ($FontSize * $UnderlineOffset)
        Add-IwbLine $Graphics $left $lineY ($left + $textWidth) $lineY $Color ([Math]::Max(1.0, $FontSize * 0.07)) $null $Group
        return $true
    }

    function Add-IwbImage {
        # Raster is a pending rasterisation (SVG sticker / WebP image) that replaces Bytes with
        # a PNG after all objects are read; $null when the file is already a CFF image format.
        param([Collections.Generic.List[object]]$Graphics, [double]$X, [double]$Y, [double]$Width, [double]$Height,
              [byte[]]$Bytes, [string]$Extension, $Raster = $null, [string]$Group = $null,
              # Clockwise rotation in degrees about the top-left corner (X, Y).
              [double]$Angle = 0)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Image'; X = $X; Y = $Y; Width = $Width; Height = $Height; Angle = $Angle
            Bytes = $Bytes; Extension = $Extension; Raster = $Raster; Path = $null; Group = $Group
        })
    }

    function Get-ImagePlacement {
        # Where an image or sticker of $LocalWidth x $LocalHeight (the anchor's own px) goes: its
        # board-space top-left corner, mapped through the anchor's matrix, and the rotation about
        # that corner. A mirrored matrix can't be written as IWB's translate/rotate, so it is
        # placed unrotated at the right size (Rotated = $false).
        param($Anchor, [double]$LocalWidth, [double]$LocalHeight)
        $u = 0.0; $v = 0.0
        if ($Anchor.IsCentered) { $u = -$LocalWidth / 2; $v = -$LocalHeight / 2 }
        $mirrored = (($Anchor.MA * $Anchor.MD) - ($Anchor.MB * $Anchor.MC)) -lt 0
        if (-not $Anchor.IsRotated -or $mirrored) {
            return [pscustomobject]@{ X = $Anchor.X + ($u * $Anchor.ScaleX); Y = $Anchor.Y + ($v * $Anchor.ScaleY); Angle = 0.0; Rotated = $false }
        }
        $corner = ConvertTo-BoardPoints $Anchor @(,[double[]]@($u, $v))
        return [pscustomobject]@{
            X = $corner[0][0]; Y = $corner[0][1]; Rotated = $true
            Angle = [Math]::Atan2($Anchor.MB, $Anchor.MA) * 180.0 / [Math]::PI
        }
    }

    function Add-IwbShape {
        # Picks the element OpenBoard's importer draws correctly: <rect> for an unrotated
        # rectangle, a filled <polygon> otherwise, and edges for an unfilled non-rectangle.
        # Returns 'Rect', 'Polygon', 'Edges' or $null (nothing visible).
        param([Collections.Generic.List[object]]$Graphics, [double[][]]$Points,
              [string]$Fill, [string]$Stroke, [double]$StrokeWidth, [double[]]$Dash, [string]$Group)
        $hasFill = ($Fill -and $Fill -ne 'transparent')
        $hasStroke = ($Stroke -and $Stroke -ne 'transparent' -and $StrokeWidth -gt 0)
        if (-not $hasFill -and -not $hasStroke) { return $null }
        $fillValue = if ($hasFill) { $Fill } else { $null }
        $strokeValue = if ($hasStroke) { $Stroke } else { $null }
        $width = if ($hasStroke) { $StrokeWidth } else { 0.0 }
        $dashValue = if ($hasStroke) { $Dash } else { $null }
        $box = Get-AxisAlignedBox $Points
        if ($null -ne $box) {
            Add-IwbRect $Graphics $box.X $box.Y $box.W $box.H $fillValue $strokeValue $width $dashValue $Group
            return 'Rect'
        }
        if ($hasFill) {
            Add-IwbPolygon $Graphics $Points $fillValue $strokeValue $width $dashValue $Group
            return 'Polygon'
        }
        Add-IwbPolyline $Graphics $Points -Closed $strokeValue $width $dashValue $Group
        return 'Edges'
    }

    function Add-IwbReactionSticker {
        # Fallback only (a sticker with no embedded artwork): a best-effort stand-in.
        param([Collections.Generic.List[object]]$Graphics, [string]$Label,
              [double]$X, [double]$Y, [double]$Width, [double]$Height, [string]$Group)
        $size = [Math]::Max(16.0, [Math]::Min($Width, $Height))
        if ($Label -match 'red cross') {
            Add-IwbLine $Graphics $X $Y ($X + $size) ($Y + $size) '#e62e55' 4 $null $Group
            Add-IwbLine $Graphics ($X + $size) $Y $X ($Y + $size) '#e62e55' 4 $null $Group
            return
        }
        if ($Label -match 'green check') {
            $pts = [double[][]]@(@($X, ($Y + ($size * 0.55))), @(($X + ($size * 0.35)), ($Y + $size)), @(($X + $size), $Y))
            Add-IwbPolyline $Graphics $pts '#4fc24f' 4 $null $Group
            return
        }
        $symbol = if ($Label -match 'star') { [char]0x2605 }
                  elseif ($Label -match 'heart') { [char]0x2764 }
                  else { [char]0x25CF }
        $color = if ($Label -match 'star') { '#f5b700' } elseif ($Label -match 'heart') { '#e8336d' } else { '#1f1f1f' }
        Add-IwbText $Graphics ([string]$symbol) $X $Y $size $size ($size * 0.9) $color 'center' -Group $Group
    }

    function Get-RasterImageDimensions {
        # Native size of an image file from its header (PNG, GIF, BMP, JPEG) or, for SVG, from
        # its root width/height or viewBox. Used to keep a sticker's aspect ratio.
        param([byte[]]$Bytes)
        if ($Bytes.Length -ge 24 -and $Bytes[0] -eq 0x89 -and $Bytes[1] -eq 0x50 -and $Bytes[2] -eq 0x4e -and $Bytes[3] -eq 0x47) {
            $w = ([int]$Bytes[16] -shl 24) -bor ([int]$Bytes[17] -shl 16) -bor ([int]$Bytes[18] -shl 8) -bor [int]$Bytes[19]
            $h = ([int]$Bytes[20] -shl 24) -bor ([int]$Bytes[21] -shl 16) -bor ([int]$Bytes[22] -shl 8) -bor [int]$Bytes[23]
            if ($w -gt 0 -and $h -gt 0) { return [pscustomobject]@{ Width = $w; Height = $h } }
        }
        elseif ($Bytes.Length -ge 10 -and $Bytes[0] -eq 0x47 -and $Bytes[1] -eq 0x49 -and $Bytes[2] -eq 0x46) {
            $w = [int]$Bytes[6] -bor ([int]$Bytes[7] -shl 8)
            $h = [int]$Bytes[8] -bor ([int]$Bytes[9] -shl 8)
            if ($w -gt 0 -and $h -gt 0) { return [pscustomobject]@{ Width = $w; Height = $h } }
        }
        elseif ($Bytes.Length -ge 26 -and $Bytes[0] -eq 0x42 -and $Bytes[1] -eq 0x4d) {
            $w = [BitConverter]::ToInt32($Bytes, 18)
            $h = [BitConverter]::ToInt32($Bytes, 22)
            if ($w -gt 0 -and $h -ne 0) { return [pscustomobject]@{ Width = $w; Height = [Math]::Abs($h) } }
        }
        elseif ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0xff -and $Bytes[1] -eq 0xd8) {
            $i = 2
            while ($i -lt $Bytes.Length - 9) {
                if ($Bytes[$i] -ne 0xff) { $i++; continue }
                $marker = $Bytes[$i + 1]
                if ($marker -eq 0xd8 -or $marker -eq 0x01 -or ($marker -ge 0xd0 -and $marker -le 0xd7)) { $i += 2; continue }
                if ($i + 3 -ge $Bytes.Length) { break }
                $segLen = ([int]$Bytes[$i + 2] -shl 8) -bor [int]$Bytes[$i + 3]
                if ($marker -ge 0xc0 -and $marker -le 0xcf -and $marker -ne 0xc4 -and $marker -ne 0xc8 -and $marker -ne 0xcc) {
                    if ($i + 9 -lt $Bytes.Length) {
                        $h = ([int]$Bytes[$i + 5] -shl 8) -bor [int]$Bytes[$i + 6]
                        $w = ([int]$Bytes[$i + 7] -shl 8) -bor [int]$Bytes[$i + 8]
                        if ($w -gt 0 -and $h -gt 0) { return [pscustomobject]@{ Width = $w; Height = $h } }
                    }
                    break
                }
                if ($marker -eq 0xd9 -or $segLen -le 0) { break }
                $i += 2 + $segLen
            }
        }
        $headLength = [Math]::Min($Bytes.Length, 8192)
        if ($headLength -gt 0) {
            $head = [Text.Encoding]::UTF8.GetString($Bytes, 0, $headLength)
            $svgTag = [regex]::Match($head, '<svg\b[^>]*>',
                [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
                [Text.RegularExpressions.RegexOptions]::Singleline)
            if ($svgTag.Success) {
                $tag = $svgTag.Value
                $svgW = 0.0; $svgH = 0.0; $vbW = 0.0; $vbH = 0.0
                $wm = [regex]::Match($tag, '\swidth\s*=\s*["'']\s*([0-9.]+)\s*(px)?\s*["'']', 'IgnoreCase')
                $hm = [regex]::Match($tag, '\sheight\s*=\s*["'']\s*([0-9.]+)\s*(px)?\s*["'']', 'IgnoreCase')
                if ($wm.Success) { $svgW = Get-Number $wm.Groups[1].Value 0 }
                if ($hm.Success) { $svgH = Get-Number $hm.Groups[1].Value 0 }
                $vb = [regex]::Match($tag, '\sviewBox\s*=\s*["'']([^"'']+)["'']', 'IgnoreCase')
                if ($vb.Success) {
                    $parts = @($vb.Groups[1].Value.Trim() -split '[\s,]+')
                    if ($parts.Count -eq 4) { $vbW = Get-Number $parts[2] 0; $vbH = Get-Number $parts[3] 0 }
                }
                if ($svgW -le 0 -and $svgH -gt 0 -and $vbW -gt 0 -and $vbH -gt 0) { $svgW = $svgH * $vbW / $vbH }
                if ($svgH -le 0 -and $svgW -gt 0 -and $vbW -gt 0 -and $vbH -gt 0) { $svgH = $svgW * $vbH / $vbW }
                if ($svgW -le 0 -or $svgH -le 0) { $svgW = $vbW; $svgH = $vbH }
                if ($svgW -gt 0 -and $svgH -gt 0) { return [pscustomobject]@{ Width = $svgW; Height = $svgH } }
            }
        }
        return $null
    }

    function Get-EmbeddedImageData {
        # Decodes the first <img src="data:..."> in an anchor block into raw bytes, a file
        # extension and its MIME type (sniffed when the declared type is a placeholder such as
        # image/*). Returns $null when there is no usable inline image.
        param([string]$Block)
        $src = Get-FirstMatch $Block '<img\b[^>]*\bsrc="([^"]+)"'
        if (-not $src) { return $null }
        $src = [Net.WebUtility]::HtmlDecode($src)
        $m = [regex]::Match($src, '^data:([^;,]*)((?:;[^,]*)?),(.*)$',
            [Text.RegularExpressions.RegexOptions]::Singleline)
        if (-not $m.Success) { return $null }
        $mime = $m.Groups[1].Value.ToLowerInvariant()
        $isBase64 = $m.Groups[2].Value -match ';\s*base64'
        $data = $m.Groups[3].Value
        try {
            if ($isBase64) { $bytes = [Convert]::FromBase64String($data) }
            else { $bytes = [Text.Encoding]::UTF8.GetBytes([Uri]::UnescapeDataString($data)) }
        } catch { return $null }
        if ($null -eq $bytes -or $bytes.Length -eq 0) { return $null }
        $isSvg = ($mime -eq 'image/svg+xml')
        if (-not $isSvg -and $mime -notmatch '^image/(png|jpe?g|gif|bmp|webp)$') {
            $head = [Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min($bytes.Length, 1024))
            if ($head -match '<svg\b') { $isSvg = $true }
            else { $mime = Get-ImageMimeType ([Convert]::ToBase64String($bytes, 0, [Math]::Min($bytes.Length, 48))) }
        }
        if ($isSvg) { $mime = 'image/svg+xml' }
        $extension = if ($isSvg) { 'svg' } else {
            switch -Regex ($mime) { 'jpe?g' { 'jpg' } 'gif' { 'gif' } 'bmp' { 'bmp' } 'webp' { 'webp' } default { 'png' } }
        }
        return [pscustomobject]@{ Bytes = $bytes; Extension = $extension; Mime = $mime }
    }

    # ===================================================================================
    # Rasterising SVG stickers / WebP images with a headless browser
    # ===================================================================================

    function Find-HeadlessBrowser {
        if ($BrowserPath) {
            if (Test-Path -LiteralPath $BrowserPath -PathType Leaf) { return $BrowserPath }
            return $null
        }
        $candidates = New-Object Collections.Generic.List[string]
        foreach ($root in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA)) {
            if (-not $root) { continue }
            $candidates.Add((Join-Path $root 'Microsoft\Edge\Application\msedge.exe'))
            $candidates.Add((Join-Path $root 'Google\Chrome\Application\chrome.exe'))
        }
        foreach ($c in $candidates) { if (Test-Path -LiteralPath $c -PathType Leaf) { return $c } }
        return $null
    }

    function ConvertTo-PngTiles {
        # Renders each request (@{ Key; Mime; Bytes; Width; Height }, integer px) into one
        # screenshot of a grid of <img> elements on a transparent page, then crops the tiles.
        # Returns Key -> PNG bytes for every tile that rendered; an empty table without a browser.
        param([object[]]$Requests)
        $result = @{}
        if ($NoRasterize -or -not $Requests -or $Requests.Count -eq 0) { return $result }
        $browser = Find-HeadlessBrowser
        if (-not $browser) { return $result }
        try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch { return $result }

        # One fixed work folder, overwritten on every run; a named mutex keeps two conversions
        # running at once from overwriting each other's tile page and screenshot.
        $mutex = New-Object Threading.Mutex($false, 'Local\WhiteboardToIwb-render')
        $owned = $false
        try { $owned = $mutex.WaitOne(300000) } catch [Threading.AbandonedMutexException] { $owned = $true }
        if (-not $owned) { $mutex.Dispose(); return $result }
        try { return (ConvertTo-PngTilesLocked $Requests $browser) }
        finally { $mutex.ReleaseMutex(); $mutex.Dispose() }
    }

    function ConvertTo-PngTilesLocked {
        param([object[]]$Requests, [string]$browser)
        $result = @{}
        $work = Join-Path ([IO.Path]::GetTempPath()) 'WhiteboardToIwb-render'
        if (-not (Test-Path -LiteralPath $work)) { [void](New-Item -ItemType Directory -Path $work) }
        $cell = 1
        foreach ($r in $Requests) { $cell = [Math]::Max($cell, [Math]::Max([int]$r.Width, [int]$r.Height)) }
        $cols = [int][Math]::Max(1, [Math]::Min($Requests.Count, [Math]::Floor(4096 / $cell)))
        $rows = [int][Math]::Ceiling($Requests.Count / $cols)
        $sb = New-Object Text.StringBuilder
        [void]$sb.Append('<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;background:transparent;overflow:hidden}img{position:absolute;display:block}</style></head><body>')
        for ($i = 0; $i -lt $Requests.Count; $i++) {
            $r = $Requests[$i]
            $uri = 'data:' + $r.Mime + ';base64,' + [Convert]::ToBase64String([byte[]]$r.Bytes)
            [void]$sb.Append(('<img src="{0}" style="left:{1}px;top:{2}px;width:{3}px;height:{4}px">' -f `
                $uri, (($i % $cols) * $cell), ([Math]::Floor($i / $cols) * $cell), [int]$r.Width, [int]$r.Height))
        }
        [void]$sb.Append('</body></html>')
        $htmlFile = Join-Path $work 'tiles.html'
        $shot = Join-Path $work 'tiles.png'
        [IO.File]::WriteAllText($htmlFile, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
        $started = [DateTime]::Now.AddSeconds(-1)
        $arguments = @('--headless', '--disable-gpu', '--hide-scrollbars', '--no-first-run', '--no-default-browser-check',
                       '--disable-extensions', '--force-device-scale-factor=1', '--default-background-color=00000000',
                       ('"--user-data-dir={0}"' -f (Join-Path $work 'profile')),
                       ('--window-size={0},{1}' -f ($cols * $cell), ($rows * $cell)),
                       ('"--screenshot={0}"' -f $shot), ([Uri]$htmlFile).AbsoluteUri)
        try {
            $p = Start-Process -FilePath $browser -ArgumentList $arguments -PassThru -WindowStyle Hidden
            if (-not $p.WaitForExit(60000)) { try { $p.Kill() } catch { }; return $result }
        } catch { return $result }
        if (-not (Test-Path -LiteralPath $shot) -or (Get-Item -LiteralPath $shot).LastWriteTime -lt $started) { return $result }

        $ms = New-Object IO.MemoryStream(,[IO.File]::ReadAllBytes($shot))
        $bmp = $null
        try {
            $bmp = New-Object Drawing.Bitmap($ms)
            for ($i = 0; $i -lt $Requests.Count; $i++) {
                $r = $Requests[$i]
                $rect = New-Object Drawing.Rectangle((($i % $cols) * $cell), ([int][Math]::Floor($i / $cols) * $cell), [int]$r.Width, [int]$r.Height)
                if ($rect.Right -gt $bmp.Width -or $rect.Bottom -gt $bmp.Height) { continue }
                $tile = $bmp.Clone($rect, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
                try {
                    # A tile with no visible pixel means the image didn't render.
                    $visible = $false
                    for ($sy = 0; $sy -lt 16 -and -not $visible; $sy++) {
                        for ($sx = 0; $sx -lt 16 -and -not $visible; $sx++) {
                            $px = $tile.GetPixel([int](($sx + 0.5) * $tile.Width / 16), [int](($sy + 0.5) * $tile.Height / 16))
                            if ($px.A -gt 0) { $visible = $true }
                        }
                    }
                    if (-not $visible) { continue }
                    $out = New-Object IO.MemoryStream
                    $tile.Save($out, [Drawing.Imaging.ImageFormat]::Png)
                    $result[$r.Key] = $out.ToArray()
                    $out.Dispose()
                } finally { $tile.Dispose() }
            }
        } catch {
            return @{}
        } finally {
            if ($bmp) { $bmp.Dispose() }
            $ms.Dispose()
        }
        return $result
    }

    # ===================================================================================
    # content.xml
    # ===================================================================================

    function ConvertTo-XmlText {
        param([string] $Text)
        if ($null -eq $Text) { return '' }
        return $Text.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&apos;')
    }

    function Format-Number {
        # Plain decimals only: OpenBoard's transform parser doesn't read exponents.
        param([double]$Value)
        $r = [Math]::Round($Value, 3)
        if ($r -eq 0) { return '0' }
        return $r.ToString('0.###', $ci)
    }

    function Get-PageMapping {
        # Fits the board's bounding box into the page (never enlarging it) and centres it.
        param([object[]]$Graphics)
        $minX = [double]::PositiveInfinity; $minY = [double]::PositiveInfinity
        $maxX = [double]::NegativeInfinity; $maxY = [double]::NegativeInfinity
        foreach ($g in $Graphics) {
            $pts = switch ($g.Kind) {
                'Rect' { ,@(@($g.X, $g.Y), @(($g.X + $g.W), ($g.Y + $g.H))) }
                'Polygon' { ,$g.Points }
                'Line' { ,@(@($g.X1, $g.Y1), @($g.X2, $g.Y2)) }
                'Image' {
                    $rad = $g.Angle * [Math]::PI / 180.0; $c = [Math]::Cos($rad); $sn = [Math]::Sin($rad)
                    ,@(foreach ($uv in @(@(0, 0), @($g.Width, 0), @(0, $g.Height), @($g.Width, $g.Height))) {
                        ,@(($g.X + ($c * $uv[0]) - ($sn * $uv[1])), ($g.Y + ($sn * $uv[0]) + ($c * $uv[1])))
                    })
                }
                'Text' {
                    # All four corners, so a rotated text box is bounded correctly too.
                    ,@(foreach ($uv in @(@(0, 0), @($g.Width, 0), @(0, $g.Height), @($g.Width, $g.Height))) {
                        ,@(($g.X + ($g.M11 * $uv[0]) + ($g.M21 * $uv[1])), ($g.Y + ($g.M12 * $uv[0]) + ($g.M22 * $uv[1])))
                    })
                }
            }
            foreach ($p in $pts) {
                $minX = [Math]::Min($minX, $p[0]); $minY = [Math]::Min($minY, $p[1])
                $maxX = [Math]::Max($maxX, $p[0]); $maxY = [Math]::Max($maxY, $p[1])
            }
        }
        if ([double]::IsInfinity($minX)) { $minX = 0.0; $minY = 0.0; $maxX = 1.0; $maxY = 1.0 }
        $bw = [Math]::Max(1.0, $maxX - $minX); $bh = [Math]::Max(1.0, $maxY - $minY)
        $availW = [Math]::Max(1.0, $PageWidth - (2 * $PageMargin)); $availH = [Math]::Max(1.0, $PageHeight - (2 * $PageMargin))
        $scale = [Math]::Min(1.0, [Math]::Min($availW / $bw, $availH / $bh))
        return [pscustomobject]@{
            Scale = $scale
            OffsetX = (($PageWidth - ($bw * $scale)) / 2) - ($minX * $scale)
            OffsetY = (($PageHeight - ($bh * $scale)) / 2) - ($minY * $scale)
        }
    }

    function New-IwbContentXml {
        param([object[]]$Graphics, $Mapping, [string]$Title, [string]$SourceName)
        $s = $Mapping.Scale; $ox = $Mapping.OffsetX; $oy = $Mapping.OffsetY
        $px = { param($x) Format-Number (($x * $s) + $ox) }
        $py = { param($y) Format-Number (($y * $s) + $oy) }
        $len = { param($v) Format-Number ($v * $s) }
        $strokeAttrs = {
            param($g, [string]$FillColor)
            # OpenBoard draws a missing or invalid stroke as a 1 px black outline, so a borderless
            # shape gets its own fill colour with width 0 (no stroke in a standard renderer).
            if ($g.Stroke) {
                $a = ' stroke="{0}" stroke-width="{1}"' -f $g.Stroke, (& $len $g.StrokeWidth)
                if ($g.Dash) {
                    $a += ' stroke-dasharray="{0}"' -f ((@($g.Dash) | ForEach-Object { & $len $_ }) -join ',')
                    if ($g.PSObject.Properties['DashOffset'] -and $g.DashOffset -gt 0) { $a += ' stroke-dashoffset="{0}"' -f (& $len $g.DashOffset) }
                }
                return $a
            }
            return ' stroke="{0}" stroke-width="0"' -f $FillColor
        }

        $sb = New-Object Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8"?>' + "`r`n")
        [void]$sb.Append(('<iwb xmlns="{0}" xmlns:iwb="{0}" xmlns:svg="{1}" xmlns:xlink="{2}" version="1.0">' -f $IWB_NS, $SVG_NS, $XLINK_NS) + "`r`n")
        [void]$sb.Append(('  <iwb:meta name="title" content="{0}"/>' -f (ConvertTo-XmlText $Title)) + "`r`n")
        [void]$sb.Append(('  <iwb:meta name="source" content="{0}"/>' -f (ConvertTo-XmlText $SourceName)) + "`r`n")
        [void]$sb.Append(('  <iwb:meta name="date" content="{0}"/>' -f [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', $ci)) + "`r`n")
        [void]$sb.Append('  <iwb:meta name="generator" content="Convert-WhiteboardHtmlToIwb.ps1"/>' + "`r`n")
        # "viewbox" is lower-case in IWB (it is what OpenBoard's importer reads).
        [void]$sb.Append(('  <svg:svg width="{0}" height="{1}" viewbox="0 0 {0} {1}">' -f $PageWidth, $PageHeight) + "`r`n")
        [void]$sb.Append('    <svg:pageset>' + "`r`n")
        [void]$sb.Append('      <svg:page id="page1">' + "`r`n")

        $groups = [ordered]@{}
        $n = 0
        foreach ($g in $Graphics) {
            $n++
            $id = 'e' + $n
            if ($g.Group) {
                if (-not $groups.Contains($g.Group)) { $groups[$g.Group] = New-Object Collections.Generic.List[string] }
                $groups[$g.Group].Add($id)
            }
            [void]$sb.Append('        ')
            switch ($g.Kind) {
                'Rect' {
                    $fill = if ($g.Fill) { $g.Fill } else { 'none' }
                    [void]$sb.Append(('<svg:rect id="{0}" x="{1}" y="{2}" width="{3}" height="{4}" fill="{5}"' -f `
                        $id, (& $px $g.X), (& $py $g.Y), (& $len $g.W), (& $len $g.H), $fill))
                    [void]$sb.Append((& $strokeAttrs $g $fill))
                    [void]$sb.Append('/>')
                }
                'Polygon' {
                    $pointStr = (@($g.Points) | ForEach-Object { '{0},{1}' -f (& $px $_[0]), (& $py $_[1]) }) -join ' '
                    [void]$sb.Append(('<svg:polygon id="{0}" points="{1}" fill="{2}"' -f $id, $pointStr, $g.Fill))
                    [void]$sb.Append((& $strokeAttrs $g $g.Fill))
                    [void]$sb.Append(' stroke-linejoin="round"/>')
                }
                'Line' {
                    [void]$sb.Append(('<svg:polyline id="{0}" points="{1},{2} {3},{4}" fill="none"' -f `
                        $id, (& $px $g.X1), (& $py $g.Y1), (& $px $g.X2), (& $py $g.Y2)))
                    [void]$sb.Append((& $strokeAttrs $g $g.Stroke))
                    [void]$sb.Append(' stroke-linecap="round"/>')
                }
                'Text' {
                    $fontPt = Format-Number ($g.FontSize * $s * 0.75)
                    # OpenBoard maps only weight names; 600 (Whiteboard's headings) is demibold, which
                    # matters for faces with a semibold, e.g. Segoe UI: drawn bold, centred labels moved 8px.
                    $weight = if ($g.Weight -ge 700) { 'bold' } elseif ($g.Weight -ge 600) { 'demibold' } else { 'normal' }
                    $align = switch ($g.Align) { 'center' { 'center' } 'right' { 'end' } default { 'start' } }
                    $angle = [Math]::Atan2($g.M12, $g.M11) * 180.0 / [Math]::PI
                    if ([Math]::Abs($angle) -gt 0.001) {
                        # OpenBoard reads only "translate(x,y)" and "rotate(a)" (rotating about the
                        # item's own corner), separated by a space; so the box sits at 0,0.
                        $position = ' x="0" y="0" transform="translate({0},{1}) rotate({2})"' -f `
                            (& $px $g.X), (& $py $g.Y), (Format-Number $angle)
                    } else {
                        $position = ' x="{0}" y="{1}"' -f (& $px $g.X), (& $py $g.Y)
                    }
                    $decoration = if ($g.Underline) { ' text-decoration="underline"' } else { '' }
                    if ($g.Italic) { $decoration += ' font-style="italic"' }
                    if ($g.LineIncrement -gt 0) { $decoration += ' line-increment="{0}"' -f (& $len $g.LineIncrement) }
                    [void]$sb.Append(('<svg:textarea id="{0}"{1} width="{2}" height="{3}" font-family="{4}" font-size="{5}pt" font-weight="{6}" fill="{7}" text-align="{8}"{9} xml:space="preserve">' -f `
                        $id, $position, (& $len $g.Width), (& $len $g.Height), $g.Family, $fontPt, $weight, $g.Color, $align, $decoration))
                    if ($null -eq $g.Runs) {
                        [void]$sb.Append(((@($g.Text -split "`n") | ForEach-Object { ConvertTo-XmlText $_ }) -join '<svg:tbreak/>'))
                    } else {
                        # Every run is a tspan that sets its weight and style. OpenBoard's importer
                        # applies each tspan to one shared format and never restores it, so text
                        # after a tspan (or a tspan without its own weight) took the previous run's
                        # bold or italic: "Profess|i|onal" imported as "Prof" + bold "essi" + bold
                        # italic "onal".
                        foreach ($r in $g.Runs) {
                            $runWeight = if ($null -eq $r.Weight) { $weight } elseif ($r.Weight -ge 700) { 'bold' } elseif ($r.Weight -ge 600) { 'demibold' } else { 'normal' }
                            $attrs = ' font-weight="{0}" font-style="{1}"' -f $runWeight, $(if ($r.Italic -or $g.Italic) { 'italic' } else { 'normal' })
                            $lines = @()
                            if ($r.Underline) { $lines += 'underline' }
                            if ($r.Strike) { $lines += 'line-through' }
                            if ($lines.Count) { $attrs += ' text-decoration="{0}"' -f ($lines -join ' ') }
                            $pieces = @($r.Text -split "`n" | ForEach-Object {
                                $safe = ConvertTo-XmlText $_
                                if ($attrs -and $safe) { '<svg:tspan{0}>{1}</svg:tspan>' -f $attrs, $safe } else { $safe }
                            })
                            [void]$sb.Append($pieces -join '<svg:tbreak/>')
                        }
                    }
                    [void]$sb.Append('</svg:textarea>')
                }
                'Image' {
                    if ([Math]::Abs($g.Angle) -gt 0.001) {
                        # Rotated about its top-left corner, in the only form OpenBoard reads.
                        $position = 'x="0" y="0" transform="translate({0},{1}) rotate({2})"' -f (& $px $g.X), (& $py $g.Y), (Format-Number $g.Angle)
                    } else {
                        $position = 'x="{0}" y="{1}"' -f (& $px $g.X), (& $py $g.Y)
                    }
                    [void]$sb.Append(('<svg:image id="{0}" {1} width="{2}" height="{3}" xlink:href="{4}"/>' -f `
                        $id, $position, (& $len $g.Width), (& $len $g.Height), $g.Path))
                }
            }
            [void]$sb.Append("`r`n")
        }

        [void]$sb.Append('      </svg:page>' + "`r`n")
        [void]$sb.Append('    </svg:pageset>' + "`r`n")
        [void]$sb.Append('  </svg:svg>' + "`r`n")
        # The parts of one Whiteboard object (a note and its text, a shape and its label, a
        # connector and its arrowheads) move as one object.
        foreach ($key in $groups.Keys) {
            $members = $groups[$key]
            if ($members.Count -lt 2) { continue }
            [void]$sb.Append('  <iwb:group>')
            foreach ($m in $members) { [void]$sb.Append(('<iwb:element ref="{0}"/>' -f $m)) }
            [void]$sb.Append('</iwb:group>' + "`r`n")
        }
        [void]$sb.Append('</iwb>' + "`r`n")
        return $sb.ToString()
    }

    function Write-IwbArchive {
        param([string] $OutFile, [string] $ContentXml, [Collections.Specialized.OrderedDictionary] $BinaryFiles)
        Add-Type -AssemblyName System.IO.Compression | Out-Null
        $fs = [IO.File]::Open($OutFile, [IO.FileMode]::Create)
        try {
            $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
            try {
                $entry = $zip.CreateEntry('content.xml', [IO.Compression.CompressionLevel]::Optimal)
                $st = $entry.Open()
                $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($ContentXml)
                $st.Write($bytes, 0, $bytes.Length); $st.Dispose()
                foreach ($name in $BinaryFiles.Keys) {
                    $entry = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
                    $st = $entry.Open()
                    $bytes = [byte[]]$BinaryFiles[$name]
                    $st.Write($bytes, 0, $bytes.Length); $st.Dispose()
                }
            } finally { $zip.Dispose() }
        } finally { $fs.Dispose() }
    }

    # ===================================================================================
    # Main per-file conversion
    # ===================================================================================

    function Convert-OneFile {
        param([string]$SourcePath, [string]$DestinationPath, [string]$Title)
        $html = [IO.File]::ReadAllText($SourcePath, [Text.Encoding]::UTF8)
        if ($html -notmatch 'data-whiteboard-type=') {
            throw "'$SourcePath' does not appear to be a Microsoft Whiteboard HTML export."
        }

        $starts = [regex]::Matches($html, '(?=<div\s+class="anchor\b)',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $blocks = for ($i = 0; $i -lt $starts.Count; $i++) {
            $start = $starts[$i].Index
            $end = if ($i + 1 -lt $starts.Count) { $starts[$i + 1].Index } else { $html.Length }
            $html.Substring($start, $end - $start)
        }

        $graphics = [Collections.Generic.List[object]]::new()
        $unsupported = [Collections.Generic.List[string]]::new()
        $sourceTypes = @{}
        $dashedCount = 0
        $stickerImageCount = 0; $stickerFallbackCount = 0
        $shapeTracedCount = 0; $shapeFallbackCount = 0; $shapeEdgesCount = 0; $shapeTextClippedCount = 0; $shapeLinesHiddenCount = 0
        $underlineDrawn = 0; $underlineNotDrawn = 0
        $rotatedTextCount = 0; $rotatedShapeCount = 0; $rotationIgnored = [Collections.Generic.List[string]]::new()
        $arrowheadCount = 0
        $groupNumber = 0
        $noteColors = Get-NoteColorTable $html
        $noteColorMissing = [Collections.Generic.List[string]]::new()
        $inkStrokeCount = 0; $inkApproximated = 0
        $connectorBentCount = 0; $connectorApprox = [Collections.Generic.List[string]]::new()
        $connectorEmptyCount = 0
        $styledRunTexts = 0
        # FluidImage (first seen in RotatedTest1) is an ordinary embedded image.
        $imageTypes = @('Image', 'AzureImage', 'FluidImage')
        $rotatedImageCount = 0
        # Shape text has no inline font-family: the stylesheet's ".textbox.shapeText" uses
        # var(--fontFamilySimple), "Aptos","Segoe UI",... in every export seen.
        $simpleFamily = Get-FontFamily ('font-family: ' + [string](Get-FirstMatch $html '--fontFamilySimple\s*:\s*([^;}]+)')) $TextFontFamily
        foreach ($block in $blocks) {
            $a = Get-AnchorInfo $block
            if (-not $a.Type) { continue }
            if (-not $sourceTypes.ContainsKey($a.Type)) { $sourceTypes[$a.Type] = 0 }
            $sourceTypes[$a.Type]++
            # PlainText, shapes, connectors and ink keep their rotation, and images and stickers
            # are rotated about their corner; rotated notes (and mirrored images, which IWB's
            # translate/rotate can't express) keep their size but are drawn unrotated -- counted
            # so the summary can say so.
            $mirrored = (($a.MA * $a.MD) - ($a.MB * $a.MC)) -lt 0
            if ($a.IsRotated -and $a.Type -eq 'Shape') { $rotatedShapeCount++ }
            elseif ($a.IsRotated -and $a.Type -in $imageTypes -and -not $mirrored) { $rotatedImageCount++ }
            elseif ($a.IsRotated -and $a.Type -notin 'PlainText','InkGroup','Connector','ReactionStickers') { [void]$rotationIgnored.Add("$($a.Type):$($a.Key)") }
            $groupNumber++
            $group = 'g' + $groupNumber

            switch ($a.Type) {
                'Shape' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*\bclass="[^"]*\bshape\b[^"]*"[^>]*>)'
                    if (-not $svgTag) { $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)' }
                    # Everything below is in the anchor's local (pre-transform) px; the outline is
                    # then mapped through the anchor's full matrix.
                    $lw = Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 100
                    $lh = Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 100
                    $u0 = 0.0; $v0 = 0.0
                    if ($a.IsCentered) { $u0 = -$lw / 2; $v0 = -$lh / 2 }
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $fill = Convert-RgbaToHex (Get-FirstMatch $gTag '\bfill="([^"]+)"') '#ffffff' -TransparentAllowed
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f' -TransparentAllowed
                    $strokeWidth = (Get-StrokeWidthPx (Get-FirstMatch $gTag '\bstroke-width="([^"]+)"') 2) * $a.ScaleX
                    $dash = Get-DashArray (Get-FirstMatch $gTag '\bstroke-dasharray="([^"]+)"')
                    if ($dash) {
                        $dashedCount++
                        $dash = [double[]]@($dash | ForEach-Object { $_ * $a.ScaleX })
                    }
                    $shapeName = Get-FirstMatch $svgTag 'aria-label="\s*([^,."]+)'
                    # Prefer Whiteboard's own outline path (exact geometry for any shape); fall
                    # back to the aria-label only when the path can't be traced. Whiteboard labels
                    # circles/ellipses "oval".
                    $local = Get-ShapePathPoints $block $u0 $v0 1 1
                    if ($null -ne $local) { $shapeTracedCount++ }
                    else {
                        $shapeFallbackCount++
                        $local = if ($shapeName -match 'ellipse|circle|oval') { Get-EllipsePoints $u0 $v0 $lw $lh }
                                 elseif ($shapeName -match 'diamond') { Get-DiamondPoints $u0 $v0 $lw $lh }
                                 else { Get-RectPoints $u0 $v0 $lw $lh }
                    }
                    $drawn = Add-IwbShape $graphics (ConvertTo-BoardPoints $a $local) $fill $stroke $strokeWidth $dash $group
                    if ($drawn -eq 'Edges') { $shapeEdgesCount++ }

                    $shapeRuns = Get-TextRuns $block
                    $shapeText = -join @($shapeRuns | ForEach-Object { $_.Text })
                    if (-not [string]::IsNullOrEmpty($shapeText)) {
                        $shapeTextStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+shapeText\b[^"]*"[^>]*style="([^"]*)"'
                        $shapeCoreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                        $fontLocal = Get-CssNumber $shapeTextStyle 'font-size' 20
                        $innerWidth = Get-CssNumber $shapeTextStyle 'width' ([Math]::Max(1.0, $lw - 26))
                        $innerHeight = Get-CssNumber $shapeTextStyle 'height' ([Math]::Max(1.0, $lh - 26))
                        $shapeWeight = Get-FontWeight $shapeCoreStyle 700
                        $shapeFamily = Get-FontFamily $shapeCoreStyle $simpleFamily
                        # The label is a flex column centred in the shape's text box, which has
                        # overflow-y: hidden. Text taller than the box starts at its top instead,
                        # and Whiteboard shows only the lines that fit; it used to be drawn whole,
                        # running far past the shape. Lines at least half inside are kept.
                        $shapeLineHeight = Get-LineHeight $shapeCoreStyle $fontLocal
                        if ($shapeLineHeight -le 0) {
                            $metrics = Get-FontLineMetrics $shapeFamily $shapeWeight
                            $shapeLineHeight = if ($null -ne $metrics) { $metrics.Normal } else { 1.25 }
                        }
                        $linePx = $shapeLineHeight * $fontLocal
                        $lineStarts = Get-TextLineStarts $shapeText $innerWidth (Get-TextMeasurer $shapeFamily $fontLocal $shapeWeight (Test-Italic $shapeCoreStyle))
                        $textHeight = $lineStarts.Count * $linePx
                        if ($textHeight -gt $innerHeight + 0.01) {
                            $textHeight = $innerHeight
                            $visibleLines = [Math]::Max(1, [int][Math]::Floor(($innerHeight / $linePx) + 0.5))
                            if ($visibleLines -lt $lineStarts.Count) {
                                $shapeText = $shapeText.Substring(0, $lineStarts[$visibleLines]).TrimEnd()
                                $shapeRuns = Select-TextRuns $shapeRuns $shapeText.Length
                                $shapeTextClippedCount++
                                $shapeLinesHiddenCount += $lineStarts.Count - $visibleLines
                            }
                        }
                        if (-not $shapeText) { break }
                        $origin = ConvertTo-BoardPoints $a @(,[double[]]@(($u0 + (($lw - $innerWidth) / 2)), ($v0 + (($lh - $textHeight) / 2))))
                        $textColor = Convert-RgbaToHex (Get-StyleValue $shapeCoreStyle 'color') '#000000'
                        $textAlign = if ($block -match 'DraftEditor-alignRight') { 'right' } elseif ($block -match 'DraftEditor-alignCenter') { 'center' } else { 'left' }
                        $shapeUnderline = Test-Underline $shapeCoreStyle
                        $shapeStyled = Test-StyledRuns $shapeRuns
                        if ($shapeStyled) { $styledRunTexts++ }
                        # The text box turns with the shape.
                        Add-IwbText $graphics $shapeText $origin[0][0] $origin[0][1] ($innerWidth * $a.ScaleX) ($textHeight * $a.ScaleY) `
                            ($fontLocal * $a.ScaleY) $textColor $textAlign `
                            ($a.MA / $a.ScaleX) ($a.MB / $a.ScaleX) ($a.MC / $a.ScaleY) ($a.MD / $a.ScaleY) `
                            -Weight $shapeWeight -Family $shapeFamily -Group $group -Underline $shapeUnderline `
                            -Italic (Test-Italic $shapeCoreStyle) -LineHeight (Get-LineHeight $shapeCoreStyle $fontLocal) -Runs $(if ($shapeStyled) { $shapeRuns } else { $null })
                        if ($a.IsRotated) { $rotatedTextCount++ }
                        if ($shapeUnderline) {
                            if (-not $a.IsRotated -and (Add-IwbUnderline $graphics $shapeText $origin[0][0] $origin[0][1] ($innerWidth * $a.ScaleX) `
                                    ($fontLocal * $a.ScaleY) $textColor $textAlign $shapeWeight $shapeFamily $group)) { $underlineDrawn++ }
                            else { $underlineNotDrawn++ }
                        }
                    }
                }
                'PlainText' {
                    $runs = Get-TextRuns $block
                    $text = -join @($runs | ForEach-Object { $_.Text })
                    $styled = Test-StyledRuns $runs
                    if ($styled) { $styledRunTexts++ }
                    $textBoxStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+plainText\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $outerStyle = Get-FirstMatch $block '<div\s+style="([^"]*width:[^"]*display:\s*flex[^"]*)"'
                    $font = (Get-CssNumber $textBoxStyle 'font-size' 20) * $a.ScaleY
                    # The glyphs start at (left + 17*s, top): the stylesheet's 16px padding plus
                    # Draft.js's 1px gutter on the left, and the matrix's translateY (-16*s) cancels
                    # the top padding. The column is the outer width minus both side insets.
                    $insetLeft = $PlainTextInsetLeft; $insetRight = $PlainTextInsetRight; $insetTop = $PlainTextInsetTop
                    $width = Get-CssNumber $outerStyle 'width' 0
                    if ($width -le 0) { $width = Get-CssNumber $textBoxStyle 'max-width' 0 }
                    if ($width -gt 0) { $width = [Math]::Max(1.0, $width - $insetLeft - $insetRight) }
                    else { $width = [Math]::Max(20.0, $text.Length * $font * 0.58 / $a.ScaleX) }
                    $width *= $a.ScaleX
                    $weight = Get-FontWeight $coreStyle 400
                    $underline = Test-Underline $coreStyle
                    $family = Get-FontFamily $coreStyle $TextFontFamily
                    # Plain text is grouped only with a drawn underline.
                    $textGroup = if ($underline) { $group } else { $null }
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $align = if ($block -match 'DraftEditor-alignCenter') { 'center' } elseif ($block -match 'DraftEditor-alignRight') { 'right' } else { 'left' }
                    # Text origin = anchor + matrix * (insetLeft, insetTop); for rotated text the
                    # inset rotates with it.
                    $textX = $a.X + ($a.MA * $insetLeft) + ($a.MC * $insetTop)
                    $textY = $a.Y + ($a.MB * $insetLeft) + ($a.MD * $insetTop)
                    $extra = @{ Weight = $weight; Family = $family; Group = $textGroup; Underline = $underline; Italic = (Test-Italic $coreStyle)
                                LineHeight = (Get-LineHeight $coreStyle ($font / $a.ScaleY)); Runs = $(if ($styled) { $runs } else { $null }) }
                    if ($a.IsRotated) {
                        $rotatedTextCount++
                        Add-IwbText $graphics $text $textX $textY $width 0 $font $color $align `
                            ($a.MA / $a.ScaleX) ($a.MB / $a.ScaleX) ($a.MC / $a.ScaleY) ($a.MD / $a.ScaleY) @extra
                        if ($underline) { $underlineNotDrawn++ }
                    } else {
                        Add-IwbText $graphics $text $textX $textY $width 0 $font $color $align @extra
                        if ($underline) {
                            if (Add-IwbUnderline $graphics $text $textX $textY $width $font $color $align $weight $family $group) { $underlineDrawn++ }
                            else { $underlineNotDrawn++ }
                        }
                    }
                }
                'Note' {
                    $noteRuns = Get-TextRuns $block
                    $text = -join @($noteRuns | ForEach-Object { $_.Text })
                    $noteStyled = Test-StyledRuns $noteRuns
                    if ($noteStyled -and $text.Trim()) { $styledRunTexts++ }
                    $noteStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+stickyNote\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $note = Get-NoteLayout $block $noteColors
                    if (-not $note.ColorFound) { [void]$noteColorMissing.Add("$($a.Type):$($a.Key)") }
                    $cssW = $note.CssWidth; $cssH = $note.CssHeight
                    $w = $note.Width * $a.ScaleX
                    $h = $note.Height * $a.ScaleY
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $gradient = if ($NoteFill -eq 'Gradient') { $note.Gradient } else { $null }
                    if ($null -ne $gradient) {
                        # Colour bands across the gradient (Get-GradientBands): a horizontal strip
                        # (classic notes) is a rect, a band across an angled gradient a polygon.
                        foreach ($band in (Get-GradientBands $note.Width $note.Height $gradient $GradientSteps)) {
                            $pts = @($band.Points | ForEach-Object { ,([double[]]@(($a.X + ($_[0] * $a.ScaleX)), ($a.Y + ($_[1] * $a.ScaleY)))) })
                            $box = Get-AxisAlignedBox $pts
                            if ($null -ne $box) { Add-IwbRect $graphics $box.X $box.Y $box.W $box.H $band.Color $null 0 $null $group }
                            else { Add-IwbPolygon $graphics $pts $band.Color $null 0 $null $group }
                        }
                    } else {
                        Add-IwbRect $graphics $a.X $a.Y $w $h $note.Solid $null 0 $null $group
                    }
                    if ($text.Trim()) {
                        $noteFont = Get-CssNumber $noteStyle 'font-size' 24
                        $noteWeight = Get-FontWeight $coreStyle 700
                        $noteUnderline = Test-Underline $coreStyle
                        $noteRaise = $noteFont * $NoteBaselineRaise
                        $noteTextX = $a.X + ($note.TextLeft * $a.ScaleX); $noteTextY = $a.Y + (($note.TextTop - $noteRaise) * $a.ScaleY)
                        $noteTextW = ($cssW - $NoteTextColumnInset) * $a.ScaleX
                        Add-IwbText $graphics $text $noteTextX $noteTextY $noteTextW (($cssH - $note.Border - 12 + $noteRaise) * $a.ScaleY) `
                            ($noteFont * $a.ScaleY) $color 'left' -Weight $noteWeight -Family $NoteFontFamily -Group $group -Underline $noteUnderline `
                            -Italic (Test-Italic $coreStyle) -Runs $(if ($noteStyled) { $noteRuns } else { $null })
                        if ($noteUnderline) {
                            if (Add-IwbUnderline $graphics $text $noteTextX $noteTextY $noteTextW ($noteFont * $a.ScaleY) $color 'left' `
                                    $noteWeight $NoteFontFamily $group) { $underlineDrawn++ }
                            else { $underlineNotDrawn++ }
                        }
                    }
                }
                'InkGroup' {
                    $strokes = @(Get-InkStrokes $a $block)
                    if ($strokes.Count -eq 0) { [void]$unsupported.Add("$($a.Type):$($a.Key) (no ink strokes)"); break }
                    # Each stroke's centreline as round-capped 2-point polylines (the only line
                    # OpenBoard's importer draws), all in the object's group.
                    foreach ($s in $strokes) {
                        Add-IwbPolyline $graphics $s.Points -Stroke $s.Color -StrokeWidth $s.Width -Group $group
                        $inkStrokeCount++
                        if (-not $s.Exact) { $inkApproximated++ }
                    }
                }
                'Connector' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)'
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $line = Get-ConnectorPoints $block
                    # Elbow connectors wrap their <svg> in <div style="position: relative; left: 0px;
                    # top: -16px;">, which moves the whole drawing: -16px on most, -1220px on one
                    # in RotatedTest1. Ignoring it put the line and arrowheads that far too low.
                    $offset = [regex]::Match($block, '<div\s+style="\s*position:\s*relative;\s*left:\s*(-?[0-9.]+)px;\s*top:\s*(-?[0-9.]+)px;?\s*">\s*<svg\b',
                        [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    $offX = 0.0; $offY = 0.0
                    if ($offset.Success) { $offX = Get-Number $offset.Groups[1].Value 0; $offY = Get-Number $offset.Groups[2].Value 0 }
                    $linePts = $null
                    if ($line -and $line.Points.Count -ge 2) {
                        $linePts = $line.Points
                        if (-not $line.Exact) { [void]$connectorApprox.Add("$($a.Type):$($a.Key)") }
                    } elseif ($line) {
                        # A zero-length line ("M11 11L11 11") draws nothing in Whiteboard.
                        $connectorEmptyCount++
                    } else {
                        # No readable path: the diagonal of the connector's SVG box.
                        $w = Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 10
                        $h = Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 10
                        $linePts = @([double[]]@(0.0, 0.0), [double[]]@($w, $h))
                        [void]$connectorApprox.Add("$($a.Type):$($a.Key)")
                    }
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f'
                    $strokeWidth = (Get-StrokeWidthPx (Get-FirstMatch $gTag '\bstroke-width="([^"]+)"') 2) * $a.ScaleX
                    $dash = Get-DashArray (Get-FirstMatch $gTag '\bstroke-dasharray="([^"]+)"')
                    if ($dash) {
                        $dashedCount++
                        $dash = [double[]]@($dash | ForEach-Object { $_ * $a.ScaleX })
                    }
                    # Elbow and curved connectors become one edge per bend (or flattened step).
                    if ($linePts) {
                        $shifted = New-Object Collections.Generic.List[double[]]
                        foreach ($p in $linePts) { [void]$shifted.Add([double[]]@(($p[0] + $offX), ($p[1] + $offY))) }
                        $boardPts = ConvertTo-BoardPoints $a $shifted.ToArray()
                        Add-IwbPolyline $graphics $boardPts -Stroke $stroke -StrokeWidth $strokeWidth -Dash $dash -Group $group
                        if ($boardPts.Count -gt 2) { $connectorBentCount++ }
                    }
                    # Arrowheads: each is its own open chevron path inside the connector's <g>,
                    # e.g. d="M-5 -9 L0 0 L5 -9" transform="translate(x,y), rotate(deg)" (tip at the
                    # local origin), drawn solid. Each becomes its two edges, grouped with the line.
                    $arrowOpts = [Text.RegularExpressions.RegexOptions]::IgnoreCase
                    $arrowRx = '<path\s+d="([^"]+)"\s+transform="\s*translate\(\s*([-+0-9.eE]+)[\s,]+([-+0-9.eE]+)\s*\)\s*,?\s*rotate\(\s*([-+0-9.eE]+)\s*\)\s*"'
                    foreach ($am in [regex]::Matches($block, $arrowRx, $arrowOpts)) {
                        $aNums = @([regex]::Matches($am.Groups[1].Value, '[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?') | ForEach-Object { Get-Number $_.Value })
                        if ($aNums.Count -lt 4 -or ($aNums.Count % 2) -ne 0) { continue }
                        $tx = Get-Number $am.Groups[2].Value 0; $ty = Get-Number $am.Groups[3].Value 0
                        $rad = (Get-Number $am.Groups[4].Value 0) * [Math]::PI / 180.0
                        $cos = [Math]::Cos($rad); $sin = [Math]::Sin($rad)
                        $arrowPts = New-Object Collections.Generic.List[double[]]
                        for ($k = 0; $k -lt $aNums.Count; $k += 2) {
                            # SVG "translate(t), rotate(r)": rotate the point, then translate it.
                            $lx = ($aNums[$k] * $cos) - ($aNums[$k + 1] * $sin) + $tx + $offX
                            $ly = ($aNums[$k] * $sin) + ($aNums[$k + 1] * $cos) + $ty + $offY
                            [void]$arrowPts.Add([double[]]@($lx, $ly))
                        }
                        Add-IwbPolyline $graphics (ConvertTo-BoardPoints $a $arrowPts.ToArray()) $stroke $strokeWidth $null $group
                        $arrowheadCount++
                    }
                }
                'ReactionStickers' {
                    $label = Get-FirstMatch $block 'aria-label="([^"]+)"'
                    $size = Get-ImageSize $block
                    $stickerData = Get-EmbeddedImageData $block
                    $stickerNative = if ($stickerData) { Get-RasterImageDimensions $stickerData.Bytes } else { $null }
                    $boxW = $size.Width; $boxH = $size.Height
                    if ($stickerNative) {
                        # The sticker <img> is usually sized only by max-width/max-height, so CSS
                        # shows it at its own size, shrunk (never enlarged) to fit that box.
                        $imgStyle = Get-FirstMatch $block '<img[^>]*style="([^"]*)"'
                        $hasExplicitSize = ($null -ne (Get-FirstMatch $imgStyle '(?:^|;)\s*width\s*:\s*[0-9.]+px')) -or
                                           ($null -ne (Get-FirstMatch $imgStyle '(?:^|;)\s*height\s*:\s*[0-9.]+px'))
                        if (-not $hasExplicitSize) {
                            $fit = [Math]::Min(1.0, [Math]::Min($boxW / $stickerNative.Width, $boxH / $stickerNative.Height))
                            $boxW = $stickerNative.Width * $fit; $boxH = $stickerNative.Height * $fit
                        }
                    }
                    $w = $boxW * $a.ScaleX; $h = $boxH * $a.ScaleY
                    $place = Get-ImagePlacement $a $boxW $boxH
                    $x = $place.X; $y = $place.Y
                    if ($stickerData) {
                        # Whiteboard's own sticker artwork, rasterised to PNG below.
                        $raster = $null
                        if ($stickerData.Extension -eq 'svg' -or $stickerData.Extension -eq 'webp') {
                            $aspect = if ($stickerNative -and $stickerNative.Height -gt 0) { $stickerNative.Width / $stickerNative.Height } else { $w / [Math]::Max(0.001, $h) }
                            $rw = if ($aspect -ge 1) { $StickerRasterSize } else { [Math]::Max(1.0, [Math]::Round($StickerRasterSize * $aspect)) }
                            $rh = if ($aspect -ge 1) { [Math]::Max(1.0, [Math]::Round($StickerRasterSize / $aspect)) } else { $StickerRasterSize }
                            $raster = [pscustomobject]@{ Mime = $stickerData.Mime; Width = [int]$rw; Height = [int]$rh }
                        }
                        Add-IwbImage $graphics $x $y $w $h $stickerData.Bytes $stickerData.Extension $raster -Angle $place.Angle
                        $stickerImageCount++
                        if ($a.IsRotated) { if ($place.Rotated) { $rotatedImageCount++ } else { [void]$rotationIgnored.Add("$($a.Type):$($a.Key)") } }
                    } else {
                        $x = $a.X; $y = $a.Y
                        if ($a.IsCentered) { $x -= $w / 2; $y -= $h / 2 }
                        Add-IwbReactionSticker $graphics $label $x $y $w $h $group
                        $stickerFallbackCount++
                        if ($a.IsRotated) { [void]$rotationIgnored.Add("$($a.Type):$($a.Key)") }
                    }
                }
                {$_ -in $imageTypes} {
                    $src = Get-FirstMatch $block '<img\b[^>]*\bsrc="([^"]+)"'
                    if (-not $src) { [void]$unsupported.Add("$($a.Type):$($a.Key) (missing image data)"); break }
                    $src = [Net.WebUtility]::HtmlDecode($src)
                    $size = Get-ImageSize $block
                    $w = $size.Width * $a.ScaleX; $h = $size.Height * $a.ScaleY
                    $place = Get-ImagePlacement $a $size.Width $size.Height
                    $x = $place.X; $y = $place.Y
                    $mime = Get-FirstMatch $src '^data:([^;,]+)'
                    if (-not $mime) { [void]$unsupported.Add("$($a.Type):$($a.Key) (external image URL)"); break }
                    $commaIndex = $src.IndexOf(',')
                    $payload = if ($commaIndex -gt 0) { $src.Substring($commaIndex + 1) } else { $null }
                    if (-not $payload) { [void]$unsupported.Add("$($a.Type):$($a.Key) (empty image payload)"); break }
                    if ($mime -notmatch '^image/(png|jpe?g|gif|bmp|webp|svg\+xml)$') {
                        $mime = Get-ImageMimeType $payload
                    }
                    if ($mime -match '^image/(png|jpe?g|gif|bmp)$') {
                        $downscaled = Get-DownscaledImagePayload $payload $w $h
                        if ($downscaled) { $payload = $downscaled; $mime = 'image/png' }
                    }
                    $extension = switch -Regex ($mime) {
                        'jpe?g' { 'jpg' } 'gif' { 'gif' } 'bmp' { 'bmp' } 'webp' { 'webp' } 'svg' { 'svg' }
                        default { 'png' }
                    }
                    $raster = $null
                    if ($extension -eq 'svg' -or $extension -eq 'webp') {
                        $raster = [pscustomobject]@{ Mime = $mime; Width = [int][Math]::Min(4096.0, [Math]::Max(1.0, [Math]::Ceiling($w * 2)))
                                                     Height = [int][Math]::Min(4096.0, [Math]::Max(1.0, [Math]::Ceiling($h * 2))) }
                    }
                    Add-IwbImage $graphics $x $y $w $h ([Convert]::FromBase64String($payload)) $extension $raster -Angle $place.Angle
                }
                default { [void]$unsupported.Add("$($a.Type):$($a.Key)") }
            }
        }

        if ($graphics.Count -eq 0) { throw "No convertible objects were found in '$SourcePath'." }

        # CFF images are PNG/JPEG/GIF/BMP: rasterise SVG and WebP images in one browser run,
        # once per distinct image and size.
        $sha = [Security.Cryptography.SHA1]::Create()
        $requests = [ordered]@{}
        foreach ($g in $graphics) {
            if ($g.Kind -ne 'Image' -or $null -eq $g.Raster) { continue }
            $key = ([BitConverter]::ToString($sha.ComputeHash([byte[]]$g.Bytes)) -replace '-', '') + "_$($g.Raster.Width)x$($g.Raster.Height)"
            $g.Raster | Add-Member -NotePropertyName Key -NotePropertyValue $key -Force
            if (-not $requests.Contains($key)) {
                $requests[$key] = [pscustomobject]@{ Key = $key; Mime = $g.Raster.Mime; Bytes = $g.Bytes; Width = $g.Raster.Width; Height = $g.Raster.Height }
            }
        }
        $pngs = ConvertTo-PngTiles @($requests.Values)
        $keptAsSvg = 0; $keptAsWebp = 0
        foreach ($g in $graphics) {
            if ($g.Kind -ne 'Image' -or $null -eq $g.Raster) { continue }
            if ($pngs.ContainsKey($g.Raster.Key)) { $g.Bytes = $pngs[$g.Raster.Key]; $g.Extension = 'png' }
            elseif ($g.Extension -eq 'svg') { $keptAsSvg++ }
            else { $keptAsWebp++ }
        }

        # One file per distinct image: a board's many copies of one sticker share it.
        $imageFiles = [ordered]@{}
        $byHash = @{}
        foreach ($g in $graphics) {
            if ($g.Kind -ne 'Image') { continue }
            $hash = ([BitConverter]::ToString($sha.ComputeHash([byte[]]$g.Bytes)) -replace '-', '') + '.' + $g.Extension
            if (-not $byHash.ContainsKey($hash)) {
                $name = 'images/image{0}.{1}' -f ($byHash.Count + 1), $g.Extension
                $byHash[$hash] = $name
                $imageFiles[$name] = [byte[]]$g.Bytes
            }
            $g.Path = $byHash[$hash]
        }
        $sha.Dispose()

        $mapping = Get-PageMapping $graphics
        $content = New-IwbContentXml $graphics $mapping $Title ([IO.Path]::GetFileName($SourcePath))
        Write-IwbArchive -OutFile $DestinationPath -ContentXml $content -BinaryFiles $imageFiles

        [pscustomobject]@{
            InputPath = $SourcePath; OutputPath = $DestinationPath
            SourceObjects = $blocks.Count; IwbElements = $graphics.Count
            EmbeddedFiles = $imageFiles.Count; PageScale = $mapping.Scale
            UnsupportedObjects = $unsupported.Count; UnsupportedDetails = @($unsupported); SourceTypeCounts = $sourceTypes
            DashedStrokes = $dashedCount
            StickersEmbedded = $stickerImageCount; StickersFallback = $stickerFallbackCount
            ImagesKeptAsSvg = $keptAsSvg; ImagesKeptAsWebp = $keptAsWebp
            ShapesTraced = $shapeTracedCount; ShapesFromLabel = $shapeFallbackCount; ShapesAsEdges = $shapeEdgesCount
            ShapeTextClipped = $shapeTextClippedCount; ShapeTextLinesHidden = $shapeLinesHiddenCount
            RotatedTextConverted = $rotatedTextCount; RotatedShapesConverted = $rotatedShapeCount
            RotationIgnored = $rotationIgnored.Count; RotationIgnoredDetails = @($rotationIgnored)
            ArrowheadsConverted = $arrowheadCount
            UnderlinesDrawn = $underlineDrawn; UnderlinesNotDrawn = $underlineNotDrawn
            NoteColorMissing = $noteColorMissing.Count; NoteColorMissingDetails = @($noteColorMissing)
            InkStrokes = $inkStrokeCount; InkApproximated = $inkApproximated
            ConnectorsBent = $connectorBentCount; ConnectorsEmpty = $connectorEmptyCount
            RotatedImagesConverted = $rotatedImageCount; StyledRunTexts = $styledRunTexts
            ConnectorsApproximated = $connectorApprox.Count; ConnectorsApproximatedDetails = @($connectorApprox)
        }
    }
}

process {
    foreach ($item in $InputPath) {
        $source = (Resolve-Path -LiteralPath $item).Path
        if ($OutputPath) {
            if ($InputPath.Count -gt 1) { throw '-OutputPath can only be used with one input file.' }
            $destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        } elseif ($OutputDirectory) {
            $directory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
            if (-not (Test-Path -LiteralPath $directory)) { [void](New-Item -ItemType Directory -Path $directory) }
            $destination = Join-Path $directory (([IO.Path]::GetFileNameWithoutExtension($source)) + '.iwb')
        } else {
            $destination = [IO.Path]::ChangeExtension($source, '.iwb')
        }
        if ((Test-Path -LiteralPath $destination) -and -not $Force) {
            throw "Output exists: '$destination'. Use -Force to overwrite it."
        }
        $title = [IO.Path]::GetFileNameWithoutExtension($source)
        $result = Convert-OneFile $source $destination $title
        Write-Host ("Converted {0} source objects to {1} IWB elements (board scaled by {2}): {3}" -f `
            $result.SourceObjects, $result.IwbElements, $result.PageScale.ToString('0.###', [Globalization.CultureInfo]::InvariantCulture), $destination)
        if ($result.ArrowheadsConverted -gt 0) {
            Write-Host ("  {0} connector arrowhead(s) converted." -f $result.ArrowheadsConverted)
        }
        if ($result.ConnectorsBent -gt 0) {
            Write-Host ("  {0} elbow/curved connector(s) traced through every bend." -f $result.ConnectorsBent)
        }
        if ($result.ConnectorsApproximated -gt 0) {
            Write-Warning ("{0} connector(s) had a line this script could not fully trace (an arc, a second subpath or no path); those parts were drawn straight: {1}" -f `
                $result.ConnectorsApproximated, ($result.ConnectorsApproximatedDetails -join ', '))
        }
        if ($result.RotatedTextConverted -gt 0) {
            Write-Host ("  {0} rotated text box(es) converted with their rotation." -f $result.RotatedTextConverted)
        }
        if ($result.RotatedShapesConverted -gt 0) {
            Write-Host ("  {0} rotated shape(s) converted with their rotation." -f $result.RotatedShapesConverted)
        }
        if ($result.RotatedImagesConverted -gt 0) {
            Write-Host ("  {0} rotated image(s)/sticker(s) converted with their rotation." -f $result.RotatedImagesConverted)
        }
        if ($result.ShapeTextClipped -gt 0) {
            Write-Host ("  {0} shape label(s) taller than their shape were cut to the lines Whiteboard shows ({1} line(s) hidden)." -f $result.ShapeTextClipped, $result.ShapeTextLinesHidden)
        }
        if ($result.ConnectorsEmpty -gt 0) {
            Write-Host ("  {0} zero-length connector(s) skipped (Whiteboard draws nothing for them)." -f $result.ConnectorsEmpty)
        }
        if ($result.RotationIgnored -gt 0) {
            Write-Warning ("{0} rotated non-text object(s) were drawn unrotated (correct size): {1}" -f `
                $result.RotationIgnored, ($result.RotationIgnoredDetails -join ', '))
        }
        if ($result.ShapesTraced -gt 0 -or $result.ShapesFromLabel -gt 0) {
            Write-Host ("  {0} shape(s) traced from their Whiteboard outline; {1} built from the shape label." -f $result.ShapesTraced, $result.ShapesFromLabel)
        }
        if ($result.ShapesAsEdges -gt 0) {
            Write-Host ("  {0} unfilled non-rectangular shape(s) written as their edges (OpenBoard fills every polygon)." -f $result.ShapesAsEdges)
        }
        if ($result.StickersEmbedded -gt 0 -or $result.StickersFallback -gt 0) {
            Write-Host ("  {0} reaction sticker(s) embedded from the original artwork; {1} drawn with a fallback symbol." -f $result.StickersEmbedded, $result.StickersFallback)
        }
        if ($result.ImagesKeptAsSvg -gt 0) {
            Write-Warning ("{0} SVG image(s)/sticker(s) were kept as SVG (no headless Edge/Chrome, or -NoRasterize): OpenBoard shows them, other IWB applications may not." -f $result.ImagesKeptAsSvg)
        }
        if ($result.ImagesKeptAsWebp -gt 0) {
            Write-Warning ("{0} WebP image(s) were kept as WebP (no headless Edge/Chrome, or -NoRasterize): not a CFF image format." -f $result.ImagesKeptAsWebp)
        }
        if ($result.DashedStrokes -gt 0) {
            Write-Host ("  {0} dashed stroke(s) kept as dashes (OpenBoard's IWB import draws them solid)." -f $result.DashedStrokes)
        }
        if ($result.InkStrokes -gt 0) {
            Write-Host ("  {0} ink stroke(s) converted." -f $result.InkStrokes)
        }
        if ($result.InkApproximated -gt 0) {
            Write-Warning ("{0} ink stroke(s) were drawn at one width in one solid colour, which differs from the original (a pressure-sensitive, translucent or effect pen)." -f $result.InkApproximated)
        }
        if ($result.NoteColorMissing -gt 0) {
            Write-Warning ("{0} sticky note(s) had no colour this script could find and were drawn in the default yellow: {1}" -f `
                $result.NoteColorMissing, ($result.NoteColorMissingDetails -join ', '))
        }
        if ($result.UnderlinesDrawn -gt 0) {
            Write-Host ("  {0} underlined text(s) also given a drawn underline (OpenBoard's IWB import ignores text-decoration)." -f $result.UnderlinesDrawn)
        }
        if ($result.UnderlinesNotDrawn -gt 0) {
            Write-Warning ("{0} underlined text(s) are wrapped or rotated, so only carry text-decoration: OpenBoard shows them without an underline." -f $result.UnderlinesNotDrawn)
        }
        if ($result.StyledRunTexts -gt 0) {
            Write-Warning ("{0} text(s) are partly bold, italic or underlined: kept as tspans. OpenBoard's importer shows their bold and italic, but not underlining, and drops the spaces at the edges of each run." -f $result.StyledRunTexts)
        }
        if ($result.UnsupportedObjects -gt 0) {
            Write-Warning ("{0} object(s) were unsupported. Details: {1}" -f `
                $result.UnsupportedObjects, ($result.UnsupportedDetails -join ', '))
        }
        if ($PassThru) { $result }
    }
}
