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
stack of -GradientSteps colour bands approximating the CSS linear-gradient.

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

    # Plain text and shape labels are "sans-serif" in the export, which renders as Arial.
    $TextFontFamily = 'Arial'

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

    function Get-HtmlText {
        param([string]$Block)
        $matches = [regex]::Matches($Block, '<span\s+data-text="true"[^>]*>(.*?)</span>',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [Text.RegularExpressions.RegexOptions]::Singleline)
        $values = foreach ($m in $matches) {
            [Net.WebUtility]::HtmlDecode(($m.Groups[1].Value -replace '<[^>]+>', ''))
        }
        # Whiteboard stores soft line breaks inside a span as a bare CR or CRLF (the export
        # keeps them raw); a browser reads both as a newline, so normalise them to LF here.
        return (($values -join '') -replace "`r`n?", "`n")
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
        param([string]$Color)
        $m = [regex]::Match($Color, 'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $m.Success) { return $null }
        # Unary comma keeps the three values together as one return object.
        return ,([int[]]@([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value))
    }

    function Get-LinearGradient {
        param([string]$Style)
        $gradient = Get-FirstMatch $Style 'linear-gradient\(([^)]*\([^)]*\)[^)]*\([^)]*\)[^)]*)\)'
        if (-not $gradient) { return $null }
        $colors = [regex]::Matches($gradient, 'rgba?\([^)]*\)',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($colors.Count -lt 2) { return $null }
        $start = Get-RgbTriplet $colors[0].Value
        $end = Get-RgbTriplet $colors[$colors.Count - 1].Value
        if ($null -eq $start -or $null -eq $end) { return $null }
        return [pscustomobject]@{ Start = $start; End = $end }
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
              [string]$Stroke, [double]$StrokeWidth = 2, [double[]]$Dash = $null, [string]$Group = $null)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Line'; X1 = $X1; Y1 = $Y1; X2 = $X2; Y2 = $Y2; Stroke = $Stroke
            StrokeWidth = $StrokeWidth; Dash = $Dash; Group = $Group
        })
    }

    function Add-IwbPolyline {
        # An open or closed run of points as its individual edges (2-point polylines).
        param([Collections.Generic.List[object]]$Graphics, [double[][]]$Points, [switch]$Closed,
              [string]$Stroke, [double]$StrokeWidth = 2, [double[]]$Dash = $null, [string]$Group = $null)
        $n = $Points.Count
        $last = if ($Closed) { $n } else { $n - 1 }
        for ($i = 0; $i -lt $last; $i++) {
            $p = $Points[$i]; $q = $Points[($i + 1) % $n]
            Add-IwbLine $Graphics $p[0] $p[1] $q[0] $q[1] $Stroke $StrokeWidth $Dash $Group
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
            [int]$Weight = 400, [string]$Family = $TextFontFamily, [string]$Group = $null
        )
        if ([string]::IsNullOrEmpty($Text)) { return }
        $FontSize = Get-ScalarDouble $FontSize 20
        if ($Width -le 1) { $Width = [Math]::Max(1.0, [Math]::Min(4000.0, $Text.Length * $FontSize * 0.62)) }
        # A textarea shows only the lines that fit its height (SVG Tiny 1.2), and OpenBoard
        # squashes a text item whose content is taller than it, so err on the tall side.
        $needed = ((Get-TextLineCount $Text $Width $FontSize) + 1) * $FontSize * 1.35
        $Height = [Math]::Max($Height, $needed)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Text'; X = $X; Y = $Y; Width = $Width; Height = $Height
            FontSize = $FontSize; Color = $Color; Align = $Align; Family = $Family; Weight = $Weight
            Text = $Text; M11 = $M11; M12 = $M12; M21 = $M21; M22 = $M22; Group = $Group
        })
    }

    function Add-IwbImage {
        # Raster is a pending rasterisation (SVG sticker / WebP image) that replaces Bytes with
        # a PNG after all objects are read; $null when the file is already a CFF image format.
        param([Collections.Generic.List[object]]$Graphics, [double]$X, [double]$Y, [double]$Width, [double]$Height,
              [byte[]]$Bytes, [string]$Extension, $Raster = $null, [string]$Group = $null)
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Image'; X = $X; Y = $Y; Width = $Width; Height = $Height
            Bytes = $Bytes; Extension = $Extension; Raster = $Raster; Path = $null; Group = $Group
        })
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
                'Image' { ,@(@($g.X, $g.Y), @(($g.X + $g.Width), ($g.Y + $g.Height))) }
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
                if ($g.Dash) { $a += ' stroke-dasharray="{0}"' -f ((@($g.Dash) | ForEach-Object { & $len $_ }) -join ',') }
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
                    $weight = if ($g.Weight -ge 600) { 'bold' } else { 'normal' }
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
                    [void]$sb.Append(('<svg:textarea id="{0}"{1} width="{2}" height="{3}" font-family="{4}" font-size="{5}pt" font-weight="{6}" fill="{7}" text-align="{8}" xml:space="preserve">' -f `
                        $id, $position, (& $len $g.Width), (& $len $g.Height), $g.Family, $fontPt, $weight, $g.Color, $align))
                    [void]$sb.Append(((@($g.Text -split "`n") | ForEach-Object { ConvertTo-XmlText $_ }) -join '<svg:tbreak/>'))
                    [void]$sb.Append('</svg:textarea>')
                }
                'Image' {
                    [void]$sb.Append(('<svg:image id="{0}" x="{1}" y="{2}" width="{3}" height="{4}" xlink:href="{5}"/>' -f `
                        $id, (& $px $g.X), (& $py $g.Y), (& $len $g.Width), (& $len $g.Height), $g.Path))
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
        $shapeTracedCount = 0; $shapeFallbackCount = 0; $shapeEdgesCount = 0
        $rotatedTextCount = 0; $rotatedShapeCount = 0; $rotationIgnored = [Collections.Generic.List[string]]::new()
        $arrowheadCount = 0
        $groupNumber = 0

        foreach ($block in $blocks) {
            $a = Get-AnchorInfo $block
            if (-not $a.Type) { continue }
            if (-not $sourceTypes.ContainsKey($a.Type)) { $sourceTypes[$a.Type] = 0 }
            $sourceTypes[$a.Type]++
            # PlainText and shapes keep their rotation; other rotated objects keep their size but
            # are drawn unrotated -- counted so the summary can say so.
            if ($a.IsRotated -and $a.Type -eq 'Shape') { $rotatedShapeCount++ }
            elseif ($a.IsRotated -and $a.Type -ne 'PlainText') { [void]$rotationIgnored.Add("$($a.Type):$($a.Key)") }
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

                    $shapeText = Get-HtmlText $block
                    if (-not [string]::IsNullOrEmpty($shapeText)) {
                        $shapeTextStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+shapeText\b[^"]*"[^>]*style="([^"]*)"'
                        $shapeCoreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                        $fontLocal = Get-CssNumber $shapeTextStyle 'font-size' 20
                        $innerWidth = Get-CssNumber $shapeTextStyle 'width' ([Math]::Max(1.0, $lw - 26))
                        $innerHeight = Get-CssNumber $shapeTextStyle 'height' ([Math]::Max(1.0, $lh - 26))
                        $textHeight = [Math]::Min($innerHeight, $fontLocal * 1.25 * (Get-TextLineCount $shapeText $innerWidth $fontLocal))
                        $origin = ConvertTo-BoardPoints $a @(,[double[]]@(($u0 + (($lw - $innerWidth) / 2)), ($v0 + (($lh - $textHeight) / 2))))
                        $textColor = Convert-RgbaToHex (Get-StyleValue $shapeCoreStyle 'color') '#000000'
                        $textAlign = if ($block -match 'DraftEditor-alignRight') { 'right' } elseif ($block -match 'DraftEditor-alignCenter') { 'center' } else { 'left' }
                        # The text box turns with the shape.
                        Add-IwbText $graphics $shapeText $origin[0][0] $origin[0][1] ($innerWidth * $a.ScaleX) ($textHeight * $a.ScaleY) `
                            ($fontLocal * $a.ScaleY) $textColor $textAlign `
                            ($a.MA / $a.ScaleX) ($a.MB / $a.ScaleX) ($a.MC / $a.ScaleY) ($a.MD / $a.ScaleY) `
                            -Weight (Get-FontWeight $shapeCoreStyle 700) -Group $group
                        if ($a.IsRotated) { $rotatedTextCount++ }
                    }
                }
                'PlainText' {
                    $text = Get-HtmlText $block
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
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $align = if ($block -match 'DraftEditor-alignCenter') { 'center' } elseif ($block -match 'DraftEditor-alignRight') { 'right' } else { 'left' }
                    # Text origin = anchor + matrix * (insetLeft, insetTop); for rotated text the
                    # inset rotates with it.
                    $textX = $a.X + ($a.MA * $insetLeft) + ($a.MC * $insetTop)
                    $textY = $a.Y + ($a.MB * $insetLeft) + ($a.MD * $insetTop)
                    if ($a.IsRotated) {
                        $rotatedTextCount++
                        Add-IwbText $graphics $text $textX $textY $width 0 $font $color $align `
                            ($a.MA / $a.ScaleX) ($a.MB / $a.ScaleX) ($a.MC / $a.ScaleY) ($a.MD / $a.ScaleY) -Weight $weight
                    } else {
                        Add-IwbText $graphics $text $textX $textY $width 0 $font $color $align -Weight $weight
                    }
                }
                'Note' {
                    $text = Get-HtmlText $block
                    $noteStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+stickyNote\b[^"]*"[^>]*style="([^"]*)"'
                    $backgroundStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxBackground\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $cssW = Get-CssNumber $noteStyle 'width' 304; $cssH = Get-CssNumber $noteStyle 'height' 304
                    $w = ($cssW + (2 * $NoteBorder)) * $a.ScaleX
                    $h = ($cssH + (2 * $NoteBorder)) * $a.ScaleY
                    $bg = Convert-RgbaToHex (Get-StyleValue $backgroundStyle 'background-color') '#fee15a'
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $gradient = if ($NoteFill -eq 'Gradient') { Get-LinearGradient $backgroundStyle } else { $null }
                    if ($null -ne $gradient) {
                        # Colour bands from the gradient's start to its end colour; each overlaps
                        # the next slightly so no hairline gaps show between them.
                        $bandHeight = $h / $GradientSteps
                        for ($i = 0; $i -lt $GradientSteps; $i++) {
                            $ratio = if ($GradientSteps -le 1) { 0.0 } else { [double]$i / [double]($GradientSteps - 1) }
                            $rgb = for ($c = 0; $c -lt 3; $c++) { [int][Math]::Round($gradient.Start[$c] + (($gradient.End[$c] - $gradient.Start[$c]) * $ratio)) }
                            $band = '#{0:x2}{1:x2}{2:x2}' -f $rgb[0], $rgb[1], $rgb[2]
                            $bandH = if ($i -lt $GradientSteps - 1) { $bandHeight + 0.15 } else { $bandHeight }
                            Add-IwbRect $graphics $a.X ($a.Y + ($i * $bandHeight)) $w $bandH $band $null 0 $null $group
                        }
                    } else {
                        Add-IwbRect $graphics $a.X $a.Y $w $h $bg $null 0 $null $group
                    }
                    if ($text.Trim()) {
                        $noteFont = Get-CssNumber $noteStyle 'font-size' 24
                        Add-IwbText $graphics $text ($a.X + ($NoteTextInsetLeft * $a.ScaleX)) ($a.Y + ($NoteTextInsetTop * $a.ScaleY)) `
                            (($cssW - $NoteTextColumnInset) * $a.ScaleX) (($cssH - $NoteTextInsetTop - 12) * $a.ScaleY) `
                            ($noteFont * $a.ScaleY) $color 'left' -Weight (Get-FontWeight $coreStyle 700) -Family $NoteFontFamily -Group $group
                    }
                }
                'Connector' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)'
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $path = Get-FirstMatch $block '<path\s+d="([^"]+)"'
                    $nums = @([regex]::Matches($path, '-?[0-9.]+') | ForEach-Object { Get-Number $_.Value })
                    if ($nums.Count -ge 4) {
                        $x1 = $nums[0]; $y1 = $nums[1]; $x2 = $nums[2]; $y2 = $nums[3]
                    } else {
                        $w = Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 10
                        $h = Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 10
                        $x1 = 0; $y1 = 0; $x2 = $w; $y2 = $h
                    }
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f'
                    $strokeWidth = (Get-StrokeWidthPx (Get-FirstMatch $gTag '\bstroke-width="([^"]+)"') 2) * $a.ScaleX
                    $dash = Get-DashArray (Get-FirstMatch $gTag '\bstroke-dasharray="([^"]+)"')
                    if ($dash) {
                        $dashedCount++
                        $dash = [double[]]@($dash | ForEach-Object { $_ * $a.ScaleX })
                    }
                    Add-IwbLine $graphics ($a.X + ($x1 * $a.ScaleX)) ($a.Y + ($y1 * $a.ScaleY)) ($a.X + ($x2 * $a.ScaleX)) ($a.Y + ($y2 * $a.ScaleY)) `
                        $stroke $strokeWidth $dash $group
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
                            $lx = ($aNums[$k] * $cos) - ($aNums[$k + 1] * $sin) + $tx
                            $ly = ($aNums[$k] * $sin) + ($aNums[$k + 1] * $cos) + $ty
                            [void]$arrowPts.Add([double[]]@(($a.X + ($lx * $a.ScaleX)), ($a.Y + ($ly * $a.ScaleY))))
                        }
                        Add-IwbPolyline $graphics $arrowPts.ToArray() $stroke $strokeWidth $null $group
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
                    $x = $a.X; $y = $a.Y
                    if ($a.IsCentered) { $x -= $w / 2; $y -= $h / 2 }
                    if ($stickerData) {
                        # Whiteboard's own sticker artwork, rasterised to PNG below.
                        $raster = $null
                        if ($stickerData.Extension -eq 'svg' -or $stickerData.Extension -eq 'webp') {
                            $aspect = if ($stickerNative -and $stickerNative.Height -gt 0) { $stickerNative.Width / $stickerNative.Height } else { $w / [Math]::Max(0.001, $h) }
                            $rw = if ($aspect -ge 1) { $StickerRasterSize } else { [Math]::Max(1.0, [Math]::Round($StickerRasterSize * $aspect)) }
                            $rh = if ($aspect -ge 1) { [Math]::Max(1.0, [Math]::Round($StickerRasterSize / $aspect)) } else { $StickerRasterSize }
                            $raster = [pscustomobject]@{ Mime = $stickerData.Mime; Width = [int]$rw; Height = [int]$rh }
                        }
                        Add-IwbImage $graphics $x $y $w $h $stickerData.Bytes $stickerData.Extension $raster
                        $stickerImageCount++
                    } else {
                        Add-IwbReactionSticker $graphics $label $x $y $w $h $group
                        $stickerFallbackCount++
                    }
                }
                {$_ -in 'Image', 'AzureImage'} {
                    $src = Get-FirstMatch $block '<img\b[^>]*\bsrc="([^"]+)"'
                    if (-not $src) { [void]$unsupported.Add("$($a.Type):$($a.Key) (missing image data)"); break }
                    $src = [Net.WebUtility]::HtmlDecode($src)
                    $size = Get-ImageSize $block
                    $w = $size.Width * $a.ScaleX; $h = $size.Height * $a.ScaleY
                    $x = $a.X; $y = $a.Y
                    if ($a.IsCentered) { $x -= $w / 2; $y -= $h / 2 }
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
                    Add-IwbImage $graphics $x $y $w $h ([Convert]::FromBase64String($payload)) $extension $raster
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
            RotatedTextConverted = $rotatedTextCount; RotatedShapesConverted = $rotatedShapeCount
            RotationIgnored = $rotationIgnored.Count; RotationIgnoredDetails = @($rotationIgnored)
            ArrowheadsConverted = $arrowheadCount
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
        if ($result.RotatedTextConverted -gt 0) {
            Write-Host ("  {0} rotated text box(es) converted with their rotation." -f $result.RotatedTextConverted)
        }
        if ($result.RotatedShapesConverted -gt 0) {
            Write-Host ("  {0} rotated shape(s) converted with their rotation." -f $result.RotatedShapesConverted)
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
        if ($result.UnsupportedObjects -gt 0) {
            Write-Warning ("{0} object(s) were unsupported. Details: {1}" -f `
                $result.UnsupportedObjects, ($result.UnsupportedDetails -join ', '))
        }
        if ($PassThru) { $result }
    }
}
