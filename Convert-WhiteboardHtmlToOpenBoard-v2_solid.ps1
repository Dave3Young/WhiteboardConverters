#requires -Version 5.1
<#
.SYNOPSIS
Converts a Microsoft Whiteboard HTML export into an OpenBoard .ubz document (v2, solid fills).

.DESCRIPTION
This is a renamed/re-targeted port of Convert-WhiteboardHtmlToExcalidraw-v2_solid_c.ps1: the
HTML parsing (anchor/position/transform reading, text extraction, image decoding/downscaling,
color parsing) is unchanged from that script. What differs is the OUTPUT side: instead of
building an Excalidraw JSON scene, this script builds an OpenBoard page in OpenBoard's native
on-disk format and packages it as a .ubz (a ZIP containing metadata.rdf + page000.svg + an
images/ folder), using the same format this session's Convert-WhiteboardToOpenBoard.ps1 was
built and verified against (OpenBoard's own UBSvgSubsetAdaptor.cpp source).

OpenBoard has no native "rectangle/ellipse/diamond with a fill and a separate stroke color"
shape primitive -- it is an ink-and-annotation tool, not a diagramming one. Shapes, sticky
notes and gradient bands are represented the way OpenBoard itself represents any solid-color
region: as an OpenBoard "ink" <polygon> (for the fill) with a "ink" <polyline> traced around
the same outline (for the stroke), both tagged with a shared ub:parent id so OpenBoard groups
them as one movable/selectable object. Connectors become an OpenBoard <line>. Straight text
becomes an OpenBoard editable text box (foreignObject), exactly like the plain .ubz converter.
Dashed strokes have no OpenBoard equivalent and are drawn solid (reported in the summary).

Object types handled: Shape, PlainText, Note (sticky note, solid fill in this v2 script --
see the _v3_gradient port for banded gradient notes), Connector, ReactionStickers (Whiteboard's
own sticker artwork embedded as an OpenBoard SVG image; native vector stand-ins only as a
fallback when a sticker has no inline artwork), Image/AzureImage (embedded, downscaled
the same way the Excalidraw port downscales them). Unlike the Excalidraw port, no standalone
backup copy of each image is written next to the output -- that workaround existed only
because Excalidraw can't reference an external file; a .ubz's images/ folder already *is*
the image's real, permanent home inside the archive.

.EXAMPLE
.\Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1 -InputPath .\sailboat.html

.EXAMPLE
Get-ChildItem .\Exports\*.html | .\Convert-WhiteboardHtmlToOpenBoard-v2_solid.ps1 `
    -OutputDirectory .\OpenBoard -Force
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

    [ValidateRange(0, 10000)]
    [double] $CanvasPadding = 100,

    [switch] $Force,
    [switch] $PassThru
)

begin {
    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'

    # ---------------------------------------------------------------------------------
    # OpenBoard format constants (from OpenBoard-org/OpenBoard source: UBSvgSubsetAdaptor.cpp,
    # UBMetadataDcSubsetAdaptor.cpp, UBSettings.cpp) -- unchanged from the plain .ubz converter.
    # ---------------------------------------------------------------------------------
    $UB_NS        = 'http://uniboard.mnemis.com/document'
    $SVG_NS       = 'http://www.w3.org/2000/svg'
    $XLINK_NS     = 'http://www.w3.org/1999/xlink'
    $XHTML_NS     = 'http://www.w3.org/1999/xhtml'
    $RDF_NS       = 'http://www.w3.org/1999/02/22-rdf-syntax-ns#'
    $DC_NS        = 'http://purl.org/dc/elements/1.1/'
    $FILE_VERSION = '4.8.0'
    $DEFAULT_DPI  = 96
    $GRID_SIZE    = 50
    $ci           = [Globalization.CultureInfo]::InvariantCulture

    # Whiteboard plain-text inset (pre-scale px): ".textbox.plainText > div { padding: 16px; }"
    # in the export's stylesheet, plus Draft.js's 1px left caret gutter. Measured identical on
    # every PlainText object in five real exports; see the PlainText branch below.
    $PlainTextInsetLeft  = 17
    $PlainTextInsetRight = 16
    $PlainTextInsetTop   = 16

    # ===================================================================================
    # Parsing helpers -- verbatim from Convert-WhiteboardHtmlToExcalidraw-v2_solid_c.ps1
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

    function New-Id {
        return ([guid]::NewGuid().ToString('N').Substring(0, 20))
    }

    function New-UbzUuid {
        # OpenBoard's ub:uuid / ub:parent attributes are parsed with QUuid(str), which
        # needs canonical GUID form -- unlike Excalidraw's opaque 20-char New-Id.
        return [guid]::NewGuid().ToString('D')
    }

    function Get-HtmlText {
        param([string]$Block)
        $matches = [regex]::Matches($Block, '<span\s+data-text="true"[^>]*>(.*?)</span>',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [Text.RegularExpressions.RegexOptions]::Singleline)
        $values = foreach ($m in $matches) {
            [Net.WebUtility]::HtmlDecode(($m.Groups[1].Value -replace '<[^>]+>', ''))
        }
        return ($values -join '')
    }

    function Get-Transform {
        # Reads a CSS matrix(a, b, c, d, e, f). The scale is the length of each column
        # (sqrt(a^2+b^2), sqrt(c^2+d^2)) rather than |a| and |d|: those are only the scale
        # when there is no rotation -- a 90-degree-rotated object has a = d = 0, which used
        # to give it a zero font size and crash the text-height estimate with Infinity.
        # The raw matrix is kept so rotated objects can be placed and rotated correctly.
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
                # Rotated (or mirrored) whenever the off-diagonal terms are non-zero or a
                # diagonal term is negative; a plain uniform scale stays "not rotated".
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
            # Raw CSS matrix terms (a, b, c, d) for rotated objects; identity-scaled otherwise.
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
        $capWidth = [Math]::Max(64, [Math]::Ceiling($DisplayWidth * 2))
        $capHeight = [Math]::Max(64, [Math]::Ceiling($DisplayHeight * 2))
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
        $newWidth = [Math]::Max(1, [Math]::Round($img.Width * $scale))
        $newHeight = [Math]::Max(1, [Math]::Round($img.Height * $scale))
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

    # ===================================================================================
    # OpenBoard graphic descriptors -- replace v2's New-BaseElement / Add-TextElement /
    # Add-ReactionSticker with equivalents that build an OpenBoard-shaped descriptor
    # instead of an Excalidraw element. Geometry stays in *global* (un-padded, un-centered)
    # whiteboard coordinates here; New-PageSvg re-bases everything at the end, the same
    # way Convert-WhiteboardToOpenBoard.ps1 does.
    # ===================================================================================

    function ConvertTo-XmlText {
        param([string] $Text)
        if ($null -eq $Text) { return '' }
        return $Text.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&apos;')
    }

    function Add-UbzText {
        param(
            [Collections.Generic.List[object]]$Graphics, [string]$Text,
            [double]$X, [double]$Y, [double]$Width, [double]$Height,
            [double]$FontSize, [string]$Color, [string]$Align = 'left',
            # Optional unit rotation (no scale) for rotated text: the box's local x axis is
            # (M11, M12) and its local y axis is (M21, M22), in SVG matrix order.
            [double]$M11 = 1, [double]$M12 = 0, [double]$M21 = 0, [double]$M22 = 1
        )
        if ([string]::IsNullOrEmpty($Text)) { return }
        $X = Get-ScalarDouble $X; $Y = Get-ScalarDouble $Y
        $Width = Get-ScalarDouble $Width 1; $Height = Get-ScalarDouble $Height 1
        $FontSize = Get-ScalarDouble $FontSize 20
        $estimatedWidth = [Math]::Max(1, [Math]::Min($Width, $Text.Length * $FontSize * 0.58))
        if ($Width -le 1) { $Width = $estimatedWidth }
        if ($Height -le 1) {
            $charsPerLine = [Math]::Max(1, [Math]::Floor($Width / ($FontSize * 0.58)))
            $lines = [Math]::Max(1, [Math]::Ceiling($Text.Length / $charsPerLine))
            $Height = $lines * $FontSize * 1.25
        }
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Text'; X = $X; Y = $Y; Width = $Width; Height = $Height
            FontSize = $FontSize; Color = $Color; Align = $Align; Family = 'Arial'
            Text = $Text; M11 = $M11; M12 = $M12; M21 = $M21; M22 = $M22
        })
    }

    function Add-UbzFilledPolygon {
        # A solid-color region -- OpenBoard's closest native equivalent to a filled shape.
        param(
            [Collections.Generic.List[object]]$Graphics,
            [double[][]]$Points, [string]$Fill, [double]$FillOpacity = 1.0,
            [string]$ParentGroup = $null
        )
        if ($Fill -eq 'transparent') { return }
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Polygon'; Points = $Points; Fill = $Fill; FillOpacity = $FillOpacity
            ParentGroup = $ParentGroup
        })
    }

    function Add-UbzOutline {
        # A closed-loop ink stroke traced around a shape's own points, standing in for
        # a separate border color -- OpenBoard polygons have no stroke of their own color
        # (see UBSvgSubsetAdaptor::polygonItemToSvgPolygon in the OpenBoard source).
        param(
            [Collections.Generic.List[object]]$Graphics,
            [double[][]]$Points, [string]$Stroke, [double]$StrokeWidth = 2,
            [string]$ParentGroup = $null
        )
        if ($Stroke -eq 'transparent' -or $Points.Count -lt 2) { return }
        $closed = New-Object Collections.Generic.List[double[]]
        foreach ($p in $Points) { [void]$closed.Add($p) }
        [void]$closed.Add($Points[0])   # repeat the first point: polylines don't auto-close
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Polyline'; Points = @($closed.ToArray()); Stroke = $Stroke
            StrokeWidth = $StrokeWidth; ParentGroup = $ParentGroup
        })
    }

    function Add-UbzLine {
        param(
            [Collections.Generic.List[object]]$Graphics,
            [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
            [string]$Stroke, [double]$StrokeWidth = 2, [string]$ParentGroup = $null
        )
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Line'; X1 = $X1; Y1 = $Y1; X2 = $X2; Y2 = $Y2
            Stroke = $Stroke; StrokeWidth = $StrokeWidth; ParentGroup = $ParentGroup
        })
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
        # visible path inside the shape's <g>), flattening curves into points, and maps it
        # to board coordinates. This keeps the true geometry of ovals, rounded rectangles,
        # triangles etc. instead of guessing it from the aria-label. Returns $null (so the
        # caller falls back to the label-based rectangle/ellipse/diamond builders) when the
        # path uses something not handled here: arcs (A), smooth shorthands (S/T), more than
        # one subpath, or a <g> transform other than a plain translate.
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
        if ($pts.Count -ge 2) {   # drop the closing duplicate; OpenBoard polygons auto-close
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

    function Add-UbzShape {
        # Fill + outline sharing one ub:parent group, so OpenBoard treats the shape as one
        # selectable/movable object instead of two independent ink pieces.
        param(
            [Collections.Generic.List[object]]$Graphics,
            [double[][]]$Points, [string]$Fill, [string]$Stroke, [double]$StrokeWidth = 2
        )
        $groupId = New-UbzUuid
        Add-UbzFilledPolygon $Graphics $Points $Fill 1.0 $groupId
        Add-UbzOutline $Graphics $Points $Stroke $StrokeWidth $groupId
    }

    function Get-RasterImageDimensions {
        # OpenBoard's .ubz reader does NOT resize an <image> element from its SVG
        # width/height attributes -- UBGraphicsPixmapItem doesn't inherit
        # UBResizableGraphicsItem, so the resize() call in graphicsItemFromSvg()
        # (UBSvgSubsetAdaptor.cpp) is skipped for images. OpenBoard always displays the
        # embedded file at its own native pixel size and then applies the transform
        # matrix's scale on top. So the transform must carry scaleX/scaleY =
        # intended-display-size / native-pixel-size, which means the native pixel size
        # has to be known up front. Parsed directly from the file header (no
        # System.Drawing dependency, so this works the same on every platform).
        param([byte[]]$Bytes)
        if ($Bytes.Length -ge 24 -and $Bytes[0] -eq 0x89 -and $Bytes[1] -eq 0x50 -and $Bytes[2] -eq 0x4e -and $Bytes[3] -eq 0x47) {
            # PNG: width/height are big-endian Int32 at offset 16 / 20 (the IHDR chunk).
            $w = ([int]$Bytes[16] -shl 24) -bor ([int]$Bytes[17] -shl 16) -bor ([int]$Bytes[18] -shl 8) -bor [int]$Bytes[19]
            $h = ([int]$Bytes[20] -shl 24) -bor ([int]$Bytes[21] -shl 16) -bor ([int]$Bytes[22] -shl 8) -bor [int]$Bytes[23]
            if ($w -gt 0 -and $h -gt 0) { return [pscustomobject]@{ Width = $w; Height = $h } }
        }
        elseif ($Bytes.Length -ge 10 -and $Bytes[0] -eq 0x47 -and $Bytes[1] -eq 0x49 -and $Bytes[2] -eq 0x46) {
            # GIF: width/height are little-endian UInt16 at offset 6 / 8.
            $w = [int]$Bytes[6] -bor ([int]$Bytes[7] -shl 8)
            $h = [int]$Bytes[8] -bor ([int]$Bytes[9] -shl 8)
            if ($w -gt 0 -and $h -gt 0) { return [pscustomobject]@{ Width = $w; Height = $h } }
        }
        elseif ($Bytes.Length -ge 26 -and $Bytes[0] -eq 0x42 -and $Bytes[1] -eq 0x4d) {
            # BMP: width/height are little-endian Int32 at offset 18 / 22 (height may be
            # negative for a top-down bitmap -- the magnitude is still the pixel count).
            $w = [BitConverter]::ToInt32($Bytes, 18)
            $h = [BitConverter]::ToInt32($Bytes, 22)
            if ($w -gt 0 -and $h -ne 0) { return [pscustomobject]@{ Width = $w; Height = [Math]::Abs($h) } }
        }
        elseif ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0xff -and $Bytes[1] -eq 0xd8) {
            # JPEG: scan markers for the first SOFn segment (0xC0-0xCF, excluding the
            # DHT/JPG/DAC marker numbers 0xC4/0xC8/0xCC) -- its payload holds height
            # then width as big-endian UInt16, 3 bytes in.
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

        # SVG (added for embedded reaction-sticker artwork): OpenBoard loads an <image>
        # whose href ends in .svg as a UBGraphicsSvgItem (UBSvgSubsetAdaptor.cpp, the
        # href.endsWith(".svg") branch). Like UBGraphicsPixmapItem it is NOT a
        # UBResizableGraphicsItem, so it is drawn at the SVG's own default size (Qt's
        # QSvgRenderer::defaultSize: the root width/height, else the viewBox size) and
        # scaled only by the transform matrix. Same rule as rasters, so return that size.
        $headLength = [Math]::Min($Bytes.Length, 8192)
        if ($headLength -gt 0) {
            $head = [Text.Encoding]::UTF8.GetString($Bytes, 0, $headLength)
            $svgTag = [regex]::Match($head, '<svg\b[^>]*>',
                [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
                [Text.RegularExpressions.RegexOptions]::Singleline)
            if ($svgTag.Success) {
                $tag = $svgTag.Value
                $svgW = 0.0; $svgH = 0.0; $vbW = 0.0; $vbH = 0.0
                # Only plain user units / px are usable; %, em etc. fall back to the viewBox.
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
        # Decodes the first <img src="data:..."> in an anchor block into raw bytes plus an
        # OpenBoard-friendly file extension. Handles both base64 and URL-encoded (plain
        # text) data URIs, and sniffs the real format when the declared MIME type is a
        # placeholder such as image/* . Returns $null when there is no usable inline image.
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
        $extension = if ($isSvg) { 'svg' } else {
            switch -Regex ($mime) { 'jpe?g' { 'jpg' } 'gif' { 'gif' } 'bmp' { 'bmp' } 'webp' { 'webp' } default { 'png' } }
        }
        return [pscustomobject]@{ Bytes = $bytes; Extension = $extension }
    }

    function Add-UbzImage {
        param(
            [Collections.Generic.List[object]]$Graphics, [Collections.Specialized.OrderedDictionary]$ImageFiles,
            [double]$X, [double]$Y, [double]$Width, [double]$Height,
            [byte[]]$Bytes, [string]$Extension
        )
        $fname = (New-UbzUuid) + '.' + $Extension
        $ImageFiles["images/$fname"] = $Bytes
        # OpenBoard sizes an <image> from native-pixel-size * transform-scale, not from
        # the SVG width/height attributes (see Get-RasterImageDimensions) -- so the
        # native pixel size of the embedded file has to travel with the graphic
        # descriptor for New-UbzPageSvg to compute the right scale factor.
        $native = Get-RasterImageDimensions $Bytes
        $nativeWidth  = if ($native) { $native.Width }  else { $Width }
        $nativeHeight = if ($native) { $native.Height } else { $Height }
        [void]$Graphics.Add([pscustomobject]@{
            Kind = 'Image'; X = $X; Y = $Y; Width = $Width; Height = $Height
            NativeWidth = $nativeWidth; NativeHeight = $nativeHeight; File = $fname
        })
    }

    function Add-UbzReactionSticker {
        param(
            [Collections.Generic.List[object]]$Graphics, [string]$Label,
            [double]$X, [double]$Y, [double]$Width, [double]$Height
        )
        $X = Get-ScalarDouble $X; $Y = Get-ScalarDouble $Y
        $Width = Get-ScalarDouble $Width 40; $Height = Get-ScalarDouble $Height 40
        $size = [Math]::Max(16, [Math]::Min($Width, $Height))

        if ($Label -match 'red cross') {
            Add-UbzLine $Graphics $X $Y ($X + $size) ($Y + $size) '#e62e55' 4
            Add-UbzLine $Graphics ($X + $size) $Y $X ($Y + $size) '#e62e55' 4
            return
        }
        if ($Label -match 'green check') {
            $checkMidY = $Y + ($size * 0.55)
            $checkMidX = $X + ($size * 0.35)
            $pts = @(@($X, $checkMidY), @($checkMidX, ($Y + $size)), @(($X + $size), $Y))
            [void]$Graphics.Add([pscustomobject]@{
                Kind = 'Polyline'; Points = $pts; Stroke = '#4fc24f'; StrokeWidth = 4; ParentGroup = $null
            })
            return
        }
        # Fallback only (used when a sticker has no embedded artwork): best-effort glyph.
        $symbol = if ($Label -match 'star') { [char]0x2605 }
                  elseif ($Label -match 'trophy') { [char]::ConvertFromUtf32(0x1F3C6) }
                  elseif ($Label -match 'heart') { [char]0x2764 }
                  elseif ($Label -match 'fire') { [char]::ConvertFromUtf32(0x1F525) }
                  elseif ($Label -match 'like|thumb') { [char]::ConvertFromUtf32(0x1F44D) }
                  elseif ($Label -match 'light ?bulb|idea') { [char]::ConvertFromUtf32(0x1F4A1) }
                  else { [char]0x25CF }
        $color = if ($Label -match 'star') { '#f5b700' } elseif ($Label -match 'trophy') { '#636363' }
                 elseif ($Label -match 'heart') { '#e8336d' } elseif ($Label -match 'fire') { '#f26b1d' }
                 elseif ($Label -match 'like|thumb') { '#f2b400' } else { '#1f1f1f' }
        Add-UbzText $Graphics ([string]$symbol) $X $Y $size $size ($size * 0.9) $color 'center'
    }

    # ===================================================================================
    # OpenBoard page/document serialization (from this session's already-verified
    # Convert-WhiteboardToOpenBoard.ps1, extended with polygon/polyline/line writers
    # cross-checked against UBSvgSubsetAdaptor::polygonItemToSvgPolygon /
    # polygonItemToSvgLine / strokeToSvgPolyline in the OpenBoard source).
    # ===================================================================================

    function Get-UbzPageGeometry {
        param([object[]] $Graphics, [double]$Padding)
        $minX = [double]::PositiveInfinity; $minY = [double]::PositiveInfinity
        $maxX = [double]::NegativeInfinity; $maxY = [double]::NegativeInfinity
        foreach ($g in $Graphics) {
            switch ($g.Kind) {
                'Text' {
                    # All four corners, so a rotated text box is bounded correctly too.
                    foreach ($uv in @(@(0, 0), @($g.Width, 0), @(0, $g.Height), @($g.Width, $g.Height))) {
                        $px = $g.X + ($g.M11 * $uv[0]) + ($g.M21 * $uv[1])
                        $py = $g.Y + ($g.M12 * $uv[0]) + ($g.M22 * $uv[1])
                        $minX = [Math]::Min($minX, $px); $minY = [Math]::Min($minY, $py)
                        $maxX = [Math]::Max($maxX, $px); $maxY = [Math]::Max($maxY, $py)
                    }
                }
                'Image' {
                    $minX = [Math]::Min($minX, $g.X); $minY = [Math]::Min($minY, $g.Y)
                    $maxX = [Math]::Max($maxX, $g.X + $g.Width); $maxY = [Math]::Max($maxY, $g.Y + $g.Height)
                }
                { $_ -in 'Polygon','Polyline' } {
                    foreach ($p in $g.Points) {
                        $minX = [Math]::Min($minX, $p[0]); $minY = [Math]::Min($minY, $p[1])
                        $maxX = [Math]::Max($maxX, $p[0]); $maxY = [Math]::Max($maxY, $p[1])
                    }
                }
                'Line' {
                    $minX = [Math]::Min($minX, [Math]::Min($g.X1, $g.X2)); $minY = [Math]::Min($minY, [Math]::Min($g.Y1, $g.Y2))
                    $maxX = [Math]::Max($maxX, [Math]::Max($g.X1, $g.X2)); $maxY = [Math]::Max($maxY, [Math]::Max($g.Y1, $g.Y2))
                }
            }
        }
        if ([double]::IsInfinity($minX)) { $minX = 0; $minY = 0; $maxX = 1280; $maxY = 960 }
        $centerX = ($minX + $maxX) / 2.0
        $centerY = ($minY + $maxY) / 2.0
        $pageW = [int][Math]::Max(1280, [Math]::Ceiling(($maxX - $minX) + (2 * $Padding)))
        $pageH = [int][Math]::Max(960,  [Math]::Ceiling(($maxY - $minY) + (2 * $Padding)))
        return [pscustomobject]@{ CenterX = $centerX; CenterY = $centerY; PageW = $pageW; PageH = $pageH }
    }

    function Get-Luminance {
        param([string] $Hex)
        $h = $Hex.TrimStart('#')
        if ($h.Length -ne 6) { return 0 }
        $r = [Convert]::ToInt32($h.Substring(0,2),16)
        $g = [Convert]::ToInt32($h.Substring(2,2),16)
        $b = [Convert]::ToInt32($h.Substring(4,2),16)
        return (0.299 * $r + 0.587 * $g + 0.114 * $b)
    }

    function New-UbzPageSvg {
        param([object[]] $Graphics, [pscustomobject] $Geometry)

        $centerX = $Geometry.CenterX; $centerY = $Geometry.CenterY
        $pageW = $Geometry.PageW; $pageH = $Geometry.PageH
        $vbX = [int](-1 * ($pageW / 2)); $vbY = [int](-1 * ($pageH / 2))

        $sb = New-Object Text.StringBuilder
        [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
        [void]$sb.Append('<svg xmlns="' + $SVG_NS + '" xmlns:xlink="' + $XLINK_NS + '" xmlns:ub="' + $UB_NS + '" xmlns:xhtml="' + $XHTML_NS + '"')
        [void]$sb.Append(' version="1.1" ub:version="' + $FILE_VERSION + '" ub:uuid="' + (New-UbzUuid) + '"')
        [void]$sb.Append((' viewBox="{0} {1} {2} {3}"' -f $vbX, $vbY, $pageW, $pageH))
        [void]$sb.Append((' ub:nominal-size="{0}x{1}"' -f $pageW, $pageH))
        [void]$sb.Append(' ub:dark-background="false" ub:crossed-background="false" ub:ruled-background="false"')
        [void]$sb.Append((' ub:grid-size="{0}" pageDpi="{1}">' -f $GRID_SIZE, $DEFAULT_DPI))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(('  <rect fill="white" x="{0}" y="{1}" width="{2}" height="{3}"/>' -f $vbX, $vbY, $pageW, $pageH))

        $z = 0
        foreach ($g in $Graphics) {
            $z++
            $zval = ($z).ToString('0.000000', $ci)
            $uuid = New-UbzUuid

            switch ($g.Kind) {
                'Text' {
                    $sx = ($g.X - $centerX).ToString($ci); $sy = ($g.Y - $centerY).ToString($ci)
                    $w = $g.Width.ToString($ci); $h = $g.Height.ToString($ci); $fs = $g.FontSize.ToString($ci)
                    $safe = ConvertTo-XmlText $g.Text
                    $html = '<p style="margin:0px; text-align:{4}; font-family:''{0}''; font-size:{1}px; color:{2};">{3}</p>' -f `
                        $g.Family, $fs, $g.Color, $safe, $g.Align
                    $innerEscaped = ConvertTo-XmlText $html
                    $colorDark = if ((Get-Luminance $g.Color) -lt 128) { '#ffffff' } else { $g.Color }
                    [void]$sb.Append('  <foreignObject ub:type="text"')
                    [void]$sb.Append((' x="0" y="0" width="{0}" height="{1}"' -f $w, $h))
                    # OpenBoard reads matrix(a, b, c, d, e, f) in standard SVG order and applies
                    # it to text items (UBSvgSubsetAdaptor::fromSvgTransform + graphicsItemFromSvg),
                    # so a rotated Whiteboard label stays a real, editable, rotated text box.
                    [void]$sb.Append((' transform="matrix({0}, {1}, {2}, {3}, {4}, {5})"' -f `
                        $g.M11.ToString($ci), $g.M12.ToString($ci), $g.M21.ToString($ci), $g.M22.ToString($ci), $sx, $sy))
                    [void]$sb.Append((' ub:z-value="{0}" ub:background="false" ub:uuid="{1}" ub:layer="0" ub:hidden-on-display="false"' -f $zval, $uuid))
                    [void]$sb.Append((' ub:width="{0}" ub:height="{1}" ub:pixels-per-point="0"' -f $w, $h))
                    [void]$sb.Append((' ub:fill-on-dark-background="{0}" ub:fill-on-light-background="{1}">' -f $colorDark, $g.Color))
                    [void]$sb.Append(('<itemTextContent>{0}</itemTextContent>' -f $innerEscaped))
                    [void]$sb.AppendLine('</foreignObject>')
                }
                'Image' {
                    $sx = ($g.X - $centerX).ToString($ci); $sy = ($g.Y - $centerY).ToString($ci)
                    $nw = [Math]::Max(1, $g.NativeWidth); $nh = [Math]::Max(1, $g.NativeHeight)
                    $scaleX = ($g.Width / $nw).ToString($ci); $scaleY = ($g.Height / $nh).ToString($ci)
                    [void]$sb.Append('  <image')
                    [void]$sb.Append((' xlink:href="images/{0}"' -f $g.File))
                    # width/height are set to the embedded file's own native pixel size (a
                    # no-op box for any generic SVG viewer) -- it's the transform's scale,
                    # not width/height, that OpenBoard itself actually sizes the image by
                    # (see Get-RasterImageDimensions / Add-UbzImage above).
                    [void]$sb.Append((' x="0" y="0" width="{0}" height="{1}"' -f $nw, $nh))
                    [void]$sb.Append((' transform="matrix({0}, 0, 0, {1}, {2}, {3})"' -f $scaleX, $scaleY, $sx, $sy))
                    [void]$sb.Append((' ub:z-value="{0}" ub:background="false" ub:uuid="{1}" ub:layer="0" ub:hidden-on-display="false"/>' -f $zval, $uuid))
                    [void]$sb.AppendLine('')
                }
                'Polygon' {
                    $pointStr = ($g.Points | ForEach-Object {
                        '{0},{1}' -f ($_[0]-$centerX).ToString('0.00',$ci), ($_[1]-$centerY).ToString('0.00',$ci)
                    }) -join ' '
                    $colorDark = if ((Get-Luminance $g.Fill) -lt 128) { '#ffffff' } else { $g.Fill }
                    [void]$sb.Append('  <polygon')
                    [void]$sb.Append((' points="{0}"' -f $pointStr))
                    [void]$sb.Append((' fill="{0}" fill-opacity="{1}" fill-rule="nonzero"' -f $g.Fill, $g.FillOpacity.ToString('0.00',$ci)))
                    [void]$sb.Append((' ub:z-value="{0}" ub:fill-on-dark-background="{1}" ub:fill-on-light-background="{2}" ub:uuid="{3}"' -f $zval, $colorDark, $g.Fill, $uuid))
                    if ($g.ParentGroup) { [void]$sb.Append((' ub:parent="{0}"' -f $g.ParentGroup)) }
                    [void]$sb.AppendLine('/>')
                }
                'Polyline' {
                    $pointStr = ($g.Points | ForEach-Object {
                        '{0},{1}' -f ($_[0]-$centerX).ToString('0.00',$ci), ($_[1]-$centerY).ToString('0.00',$ci)
                    }) -join ' '
                    $colorDark = if ((Get-Luminance $g.Stroke) -lt 128) { '#ffffff' } else { $g.Stroke }
                    [void]$sb.Append('  <polyline')
                    [void]$sb.Append((' points="{0}"' -f $pointStr))
                    [void]$sb.Append(' fill="none"')
                    [void]$sb.Append((' stroke-width="{0}" stroke="{1}" stroke-opacity="1" stroke-linecap="round"' -f $g.StrokeWidth.ToString($ci), $g.Stroke))
                    [void]$sb.Append((' ub:z-value="{0}" ub:fill-on-dark-background="{1}" ub:fill-on-light-background="{2}" ub:uuid="{3}"' -f $zval, $colorDark, $g.Stroke, $uuid))
                    if ($g.ParentGroup) { [void]$sb.Append((' ub:parent="{0}"' -f $g.ParentGroup)) }
                    [void]$sb.AppendLine('/>')
                }
                'Line' {
                    $x1 = ($g.X1-$centerX).ToString('0.00',$ci); $y1 = ($g.Y1-$centerY).ToString('0.00',$ci)
                    $x2 = ($g.X2-$centerX).ToString('0.00',$ci); $y2 = ($g.Y2-$centerY).ToString('0.00',$ci)
                    $colorDark = if ((Get-Luminance $g.Stroke) -lt 128) { '#ffffff' } else { $g.Stroke }
                    [void]$sb.Append('  <line')
                    [void]$sb.Append((' x1="{0}" y1="{1}" x2="{2}" y2="{3}"' -f $x1, $y1, $x2, $y2))
                    [void]$sb.Append((' stroke-width="{0}" stroke="{1}" stroke-linecap="round"' -f $g.StrokeWidth.ToString($ci), $g.Stroke))
                    [void]$sb.Append((' ub:z-value="{0}" ub:fill-on-dark-background="{1}" ub:fill-on-light-background="{2}" ub:uuid="{3}"' -f $zval, $colorDark, $g.Stroke, $uuid))
                    # A shared ub:parent groups a connector with its arrowheads (OpenBoard wraps
                    # <line>/<polyline>/<polygon> into one UBGraphicsStrokesGroup per parent id).
                    if ($g.ParentGroup) { [void]$sb.Append((' ub:parent="{0}"' -f $g.ParentGroup)) }
                    [void]$sb.AppendLine('/>')
                }
            }
        }

        [void]$sb.AppendLine('</svg>')
        return $sb.ToString()
    }

    function New-UbzMetadataRdf {
        param([string] $DocId, [string] $DocTitle, [int] $PageW, [int] $PageH)
        $now = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        $t = ConvertTo-XmlText $DocTitle
        @"
<?xml version="1.0" encoding="UTF-8"?>
<RDF xmlns="$RDF_NS" xmlns:dc="$DC_NS" xmlns:ub="$UB_NS">
    <Description about="$DocId">
        <dc:title>$t</dc:title>
        <dc:type>Subject</dc:type>
        <dc:date>$now</dc:date>
        <dc:format>image/svg+xml</dc:format>
        <dc:identifier>$DocId</dc:identifier>
        <ub:version>$FILE_VERSION</ub:version>
        <ub:size>${PageW}x${PageH}</ub:size>
        <ub:updated-at>$now</ub:updated-at>
    </Description>
</RDF>
"@
    }

    function Write-UbzArchive {
        param([string] $OutFile, [Collections.Specialized.OrderedDictionary] $TextFiles, [Collections.Specialized.OrderedDictionary] $BinaryFiles)
        Add-Type -AssemblyName System.IO.Compression | Out-Null
        if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force }
        $fs = [IO.File]::Open($OutFile, [IO.FileMode]::Create)
        try {
            $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
            try {
                $enc = New-Object Text.UTF8Encoding($false)
                foreach ($name in $TextFiles.Keys) {
                    $entry = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
                    $s = $entry.Open()
                    $bytes = $enc.GetBytes([string]$TextFiles[$name])
                    $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
                }
                foreach ($name in $BinaryFiles.Keys) {
                    $entry = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
                    $s = $entry.Open()
                    $bytes = [byte[]]$BinaryFiles[$name]
                    $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
                }
            } finally { $zip.Dispose() }
        } finally { $fs.Dispose() }
    }

    # ===================================================================================
    # Main per-file conversion -- object-type switch ported from v2_solid_c.ps1;
    # each case now appends OpenBoard graphic descriptors instead of Excalidraw elements.
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
        $imageFiles = [ordered]@{}
        $unsupported = [Collections.Generic.List[string]]::new()
        $sourceTypes = @{}
        $dashedCount = 0
        $stickerImageCount = 0; $stickerFallbackCount = 0
        $shapeTracedCount = 0; $shapeFallbackCount = 0
        $rotatedTextCount = 0; $rotationIgnored = [Collections.Generic.List[string]]::new()
        $arrowheadCount = 0

        foreach ($block in $blocks) {
            $a = Get-AnchorInfo $block
            if (-not $a.Type) { continue }
            if (-not $sourceTypes.ContainsKey($a.Type)) { $sourceTypes[$a.Type] = 0 }
            $sourceTypes[$a.Type]++
            # Only PlainText is rotated natively so far; other rotated objects keep their
            # correct size but are drawn unrotated -- counted so the summary can say so.
            if ($a.IsRotated -and $a.Type -ne 'PlainText') { [void]$rotationIgnored.Add("$($a.Type):$($a.Key)") }

            switch ($a.Type) {
                'Shape' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*\bclass="[^"]*\bshape\b[^"]*"[^>]*>)'
                    if (-not $svgTag) { $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)' }
                    $w = (Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 100) * $a.ScaleX
                    $h = (Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 100) * $a.ScaleY
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $fill = Convert-RgbaToHex (Get-FirstMatch $gTag '\bfill="([^"]+)"') '#ffffff' -TransparentAllowed
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f' -TransparentAllowed
                    $dash = Get-FirstMatch $gTag '\bstroke-dasharray="([^"]+)"'
                    if ($dash) { $dashedCount++ }   # OpenBoard ink strokes have no dash style; drawn solid
                    $x = $a.X; $y = $a.Y
                    if ($a.IsCentered) { $x -= $w / 2; $y -= $h / 2 }
                    $shapeName = Get-FirstMatch $svgTag 'aria-label="\s*([^,."]+)'
                    # Prefer Whiteboard's own outline path (exact geometry for any shape);
                    # fall back to the aria-label only when the path can't be traced.
                    # Whiteboard labels circles/ellipses "oval" -- previously unmatched here,
                    # which turned every oval into a rectangle.
                    $points = Get-ShapePathPoints $block $x $y $a.ScaleX $a.ScaleY
                    if ($null -ne $points) { $shapeTracedCount++ }
                    else {
                        $shapeFallbackCount++
                        $points = if ($shapeName -match 'ellipse|circle|oval') { Get-EllipsePoints $x $y $w $h }
                                  elseif ($shapeName -match 'diamond') { Get-DiamondPoints $x $y $w $h }
                                  else { Get-RectPoints $x $y $w $h }
                    }
                    Add-UbzShape $graphics $points $fill $stroke 2

                    $shapeText = Get-HtmlText $block
                    if (-not [string]::IsNullOrEmpty($shapeText)) {
                        $shapeTextStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+shapeText\b[^"]*"[^>]*style="([^"]*)"'
                        $shapeCoreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                        $font = (Get-CssNumber $shapeTextStyle 'font-size' 20) * $a.ScaleY
                        $innerWidth = (Get-CssNumber $shapeTextStyle 'width' ([Math]::Max(1, $w - 26))) * $a.ScaleX
                        $innerHeight = (Get-CssNumber $shapeTextStyle 'height' ([Math]::Max(1, $h - 26))) * $a.ScaleY
                        $textHeight = [Math]::Min($innerHeight, $font * 1.25)
                        $textX = $x + (($w - $innerWidth) / 2)
                        $textY = $y + (($h - $textHeight) / 2)
                        $textColor = Convert-RgbaToHex (Get-StyleValue $shapeCoreStyle 'color') '#000000'
                        $textAlign = if ($block -match 'DraftEditor-alignRight') { 'right' } elseif ($block -match 'DraftEditor-alignCenter') { 'center' } else { 'left' }
                        Add-UbzText $graphics $shapeText $textX $textY $innerWidth $textHeight $font $textColor $textAlign
                    }
                }
                'PlainText' {
                    $text = Get-HtmlText $block
                    $textBoxStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+plainText\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $outerStyle = Get-FirstMatch $block '<div\s+style="([^"]*width:[^"]*display:\s*flex[^"]*)"'
                    $font = (Get-CssNumber $textBoxStyle 'font-size' 20) * $a.ScaleY
                    # Whiteboard insets the text inside every plain-text box: the export's own
                    # stylesheet has ".textbox.plainText > div { padding: 16px; }", and Draft.js's
                    # editor adds 1px more on the left. Those are pre-scale px, and the anchor's
                    # matrix translateY (-16 x scale) is what cancels the top padding -- so the
                    # glyphs actually start at (left + 17*s, top), not at (left, top - 16*s).
                    # Placing the box at the anchor point put every text ~17*s px left and
                    # ~16*s px high, and measuring right/center alignment against the outer
                    # edges instead of the padded ones shifted those texts too.
                    $insetLeft = $PlainTextInsetLeft; $insetRight = $PlainTextInsetRight; $insetTop = $PlainTextInsetTop
                    $width = Get-CssNumber $outerStyle 'width' 0
                    if ($width -le 0) { $width = Get-CssNumber $textBoxStyle 'max-width' 0 }
                    if ($width -gt 0) { $width = [Math]::Max(1, $width - $insetLeft - $insetRight) }
                    else { $width = [Math]::Max(20, $text.Length * $font * 0.58 / $a.ScaleX) }
                    $width *= $a.ScaleX
                    $height = [Math]::Max($font * 1.25, [Math]::Ceiling(($text.Length * $font * 0.58) / [Math]::Max(1, $width)) * $font * 1.25)
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $align = if ($block -match 'DraftEditor-alignCenter') { 'center' } elseif ($block -match 'DraftEditor-alignRight') { 'right' } else { 'left' }
                    # Text origin = anchor + matrix * (insetLeft, insetTop). With no rotation this
                    # is exactly (left + 17*s, top); for rotated text the inset rotates with it.
                    $textX = $a.X + ($a.MA * $insetLeft) + ($a.MC * $insetTop)
                    $textY = $a.Y + ($a.MB * $insetLeft) + ($a.MD * $insetTop)
                    if ($a.IsRotated) {
                        $rotatedTextCount++
                        Add-UbzText $graphics $text $textX $textY $width $height $font $color $align `
                            ($a.MA / $a.ScaleX) ($a.MB / $a.ScaleX) ($a.MC / $a.ScaleY) ($a.MD / $a.ScaleY)
                    } else {
                        Add-UbzText $graphics $text $textX $textY $width $height $font $color $align
                    }
                }
                'Note' {
                    $text = Get-HtmlText $block
                    $noteStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+stickyNote\b[^"]*"[^>]*style="([^"]*)"'
                    $backgroundStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxBackground\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $w = (Get-CssNumber $noteStyle 'width' 304) * $a.ScaleX
                    $h = (Get-CssNumber $noteStyle 'height' 304) * $a.ScaleY
                    $bg = Convert-RgbaToHex (Get-StyleValue $backgroundStyle 'background-color') '#fee15a'
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    # v2: solid fill only, even if the source note used a CSS gradient
                    # background (see the _v3_gradient port for banded gradient notes).
                    Add-UbzFilledPolygon $graphics (Get-RectPoints $a.X $a.Y $w $h) $bg 1.0
                    Add-UbzText $graphics $text ($a.X + 12) ($a.Y + 12) ($w - 24) ($h - 24) (20 * $a.ScaleY) $color 'left'
                }
                'Connector' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)'
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $path = Get-FirstMatch $block '<path\s+d="([^"]+)"'
                    $nums = @([regex]::Matches($path, '-?[0-9.]+') | ForEach-Object { Get-Number $_.Value })
                    if ($nums.Count -ge 4) {
                        $x1=$nums[0]; $y1=$nums[1]; $x2=$nums[2]; $y2=$nums[3]
                    } else {
                        $w = Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 10
                        $h = Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 10
                        $x1=0; $y1=0; $x2=$w; $y2=$h
                    }
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f'
                    if ($gTag -match 'stroke-dasharray') { $dashedCount++ }
                    $connectorGroup = New-UbzUuid
                    Add-UbzLine $graphics ($a.X + ($x1 * $a.ScaleX)) ($a.Y + ($y1 * $a.ScaleY)) ($a.X + ($x2 * $a.ScaleX)) ($a.Y + ($y2 * $a.ScaleY)) $stroke 2 $connectorGroup
                    # Arrowheads: Whiteboard draws each as its own open chevron path inside the
                    # connector's <g>, e.g. d="M-5 -9 L0 0 L5 -9" transform="translate(x,y), rotate(deg)"
                    # (tip at the local origin). Previously only the first <path> was read, so
                    # every arrowhead was dropped. Each one becomes a 3-point OpenBoard polyline in
                    # the connector's color, grouped with the line so they move as one object.
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
                        [void]$graphics.Add([pscustomobject]@{
                            Kind = 'Polyline'; Points = @($arrowPts.ToArray()); Stroke = $stroke
                            StrokeWidth = 2; ParentGroup = $connectorGroup
                        })
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
                        # The sticker <img> is usually sized only by max-width/max-height, so
                        # CSS shows it at its own size, shrunk (never enlarged) to fit that box.
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
                        # Embed Whiteboard's own sticker artwork (an SVG for every built-in
                        # sticker) so any sticker type -- heart, fire, like, light bulb, and any
                        # added later -- converts faithfully, instead of mapping known labels
                        # to hand-drawn stand-ins and printing '?' for everything else.
                        Add-UbzImage $graphics $imageFiles $x $y $w $h $stickerData.Bytes $stickerData.Extension
                        $stickerImageCount++
                    } else {
                        Add-UbzReactionSticker $graphics $label $x $y $w $h
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
                    Add-UbzImage $graphics $imageFiles $x $y $w $h ([Convert]::FromBase64String($payload)) $extension
                }
                default { [void]$unsupported.Add("$($a.Type):$($a.Key)") }
            }
        }

        if ($graphics.Count -eq 0) { throw "No convertible objects were found in '$SourcePath'." }

        $geometry = Get-UbzPageGeometry -Graphics $graphics -Padding $CanvasPadding
        $docId = New-UbzUuid
        $textFiles = [ordered]@{
            'metadata.rdf' = New-UbzMetadataRdf -DocId $docId -DocTitle $Title -PageW $geometry.PageW -PageH $geometry.PageH
            'page000.svg'  = New-UbzPageSvg -Graphics $graphics -Geometry $geometry
        }
        Write-UbzArchive -OutFile $DestinationPath -TextFiles $textFiles -BinaryFiles $imageFiles

        [pscustomobject]@{
            InputPath=$SourcePath; OutputPath=$DestinationPath
            SourceObjects=$blocks.Count; OpenBoardGraphics=$graphics.Count
            EmbeddedFiles=$imageFiles.Count; UnsupportedObjects=$unsupported.Count
            UnsupportedDetails=@($unsupported); SourceTypeCounts=$sourceTypes
            DashedStrokesFlattened=$dashedCount
            StickersEmbedded=$stickerImageCount; StickersFallback=$stickerFallbackCount
            ShapesTraced=$shapeTracedCount; ShapesFromLabel=$shapeFallbackCount
            RotatedTextConverted=$rotatedTextCount; RotationIgnored=$rotationIgnored.Count
            ArrowheadsConverted=$arrowheadCount
            RotationIgnoredDetails=@($rotationIgnored)
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
            $destination = Join-Path $directory (([IO.Path]::GetFileNameWithoutExtension($source)) + '.ubz')
        } else {
            $destination = [IO.Path]::ChangeExtension($source, '.ubz')
        }
        if ((Test-Path -LiteralPath $destination) -and -not $Force) {
            throw "Output exists: '$destination'. Use -Force to overwrite it."
        }
        $title = [IO.Path]::GetFileNameWithoutExtension($source)
        $result = Convert-OneFile $source $destination $title
        Write-Host ("Converted {0} source objects to {1} OpenBoard graphics: {2}" -f `
            $result.SourceObjects, $result.OpenBoardGraphics, $destination)
        if ($result.ArrowheadsConverted -gt 0) {
            Write-Host ("  {0} connector arrowhead(s) converted." -f $result.ArrowheadsConverted)
        }
        if ($result.RotatedTextConverted -gt 0) {
            Write-Host ("  {0} rotated text box(es) converted with their rotation." -f $result.RotatedTextConverted)
        }
        if ($result.RotationIgnored -gt 0) {
            Write-Warning ("{0} rotated non-text object(s) were drawn unrotated (correct size): {1}" -f `
                $result.RotationIgnored, ($result.RotationIgnoredDetails -join ', '))
        }
        if ($result.ShapesTraced -gt 0 -or $result.ShapesFromLabel -gt 0) {
            Write-Host ("  {0} shape(s) traced from their Whiteboard outline; {1} built from the shape label." -f $result.ShapesTraced, $result.ShapesFromLabel)
        }
        if ($result.StickersEmbedded -gt 0 -or $result.StickersFallback -gt 0) {
            Write-Host ("  {0} reaction sticker(s) embedded from the original artwork; {1} drawn with a fallback symbol." -f $result.StickersEmbedded, $result.StickersFallback)
        }
        if ($result.DashedStrokesFlattened -gt 0) {
            Write-Host ("  {0} dashed stroke(s) were drawn solid (OpenBoard has no dashed ink style)." -f $result.DashedStrokesFlattened)
        }
        if ($result.UnsupportedObjects -gt 0) {
            Write-Warning ("{0} object(s) were unsupported. Details: {1}" -f `
                $result.UnsupportedObjects, ($result.UnsupportedDetails -join ', '))
        }
        if ($PassThru) { $result }
    }
}