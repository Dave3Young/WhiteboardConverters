#requires -Version 5.1
<#
.SYNOPSIS
Converts a Microsoft Whiteboard HTML export into an editable Excalidraw scene (v2).

.DESCRIPTION
Reads the self-contained HTML produced by Export-WhiteboardHtml and translates
the board objects exposed in that document into Excalidraw JSON. The converter
currently handles shapes, plain text, sticky notes, connectors (with their
arrowheads), reaction stickers and embedded images. Stickers embed Whiteboard's
own SVG artwork. Shapes are traced from their outline path: rectangles, ovals and
diamonds become native Excalidraw shapes sized to that outline, anything else a
closed polygon. Rotated objects keep their rotation (Excalidraw's element angle).
Shape-contained text is emitted as a separate editable text element.

The HTML export is a rendered representation, not Microsoft's native board
schema. Consequently, CSS-only texture/gradient effects are reduced to solid
colors. Anything drawn approximately (mirrored objects drawn unmirrored, stickers
without artwork, shapes built from their label) is reported in the run summary.

.EXAMPLE
.\Convert-WhiteboardHtmlToExcalidraw.ps1 -InputPath .\Templateboard.html

.EXAMPLE
Get-ChildItem .\Exports\*.html | .\Convert-WhiteboardHtmlToExcalidraw.ps1 `
    -OutputDirectory .\Excalidraw -Force
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

    # Whiteboard plain-text inset (pre-scale px): ".textbox.plainText > div { padding: 16px; }"
    # in the export's stylesheet, plus Draft.js's 1px left caret gutter. The anchor's matrix
    # translateY (-16 x scale) cancels the top padding, so glyphs start at (left + 17*s, top).
    $PlainTextInsetLeft  = 17
    $PlainTextInsetRight = 16
    $PlainTextInsetTop   = 16

    # Excalidraw lays text out with its own font metrics: for fontFamily 2 (Helvetica, drawn
    # with Arial on Windows) the first baseline sits at y + ((lineHeight - 1) / 2 + 1577/2048)
    # * fontSize. Whiteboard's "sans-serif" is Arial at line-height normal (1.15 x font
    # size), with the baseline 1854/2048 x font size below the line top. Using Excalidraw's
    # own Helvetica line height (1.15) and moving each text down by the difference puts the
    # first baseline -- and every later line -- where Whiteboard draws it.
    $TextLineHeight    = 1.15
    $TextBaselineShift = (1854 / 2048) - ((($TextLineHeight - 1) / 2) + (1577 / 2048))

    # Underlines (drawn as lines, see Add-Underlines): centre below the baseline and thickness,
    # in font sizes, as Chromium draws Arial's.
    $UnderlineOffset    = 217 / 2048
    $UnderlineThickness = 0.08

    # ".stickyNote { border-width: 1px }" sits outside the note's CSS width/height, and the
    # note's background fills that border too: a 304 x 304 note shows as 306 x 306.
    $NoteBorder = 1

    # Sticky-note text inset from the note's outer corner (pre-scale px), measured on 12 typed
    # notes in four real exports: 1px border + 12px ".textBoxCoreWrapper" side padding + Draft.js's
    # 1px caret gutter on the left, and only the border on top -- the inline "padding: 0px" on
    # .textBoxCore overrides the stylesheet's 12/16px. The column is the CSS width minus 2 x 12
    # and the gutter. Note text is 24px (".stickyNote.fixedFontSize") in the note font stack
    # (Aptos, "Segoe UI", ...). Whiteboard draws it in Aptos, Microsoft's current default font:
    # line-height normal is 2500/2048 x font size (Aptos's ascent 1923 + descent 577), with
    # the baseline 1923/2048 below the line top. (The first measurements were taken without
    # Aptos installed, in Segoe UI, whose baseline sits 0.14 x font size lower -- 3.4px at
    # 24px.) The note's text element uses that line height, shifted so its first baseline
    # lands where Whiteboard's does.
    $NoteTextInsetLeft   = 14
    $NoteTextInsetTop    = 1
    $NoteTextColumnInset = 25
    $NoteLineHeight      = 2500 / 2048
    $NoteBaselineShift   = (1923 / 2048) - ((($NoteLineHeight - 1) / 2) + (1577 / 2048))

    # Whiteboard's newer note style (class "noteVisualUpdate" on .textBoxBackground, first seen
    # in CompareAndContrast2) has no border, and puts a 40px author title bar
    # (".titleBarWithAttribution { height: 40px }") above the CSS-sized text area: a 304 x 265
    # note shows as 304 x 305, with its text 13px in and 40px down. Its colours are given only
    # as a class (e.g. "paleGreenGradient"); see Get-NoteColorTable.
    $NoteTitleBarHeight = 40

    # Fill of a note whose colour can't be found (Whiteboard's default yellow); counted in the
    # run summary.
    $DefaultNoteFill = '#fee15a'

    # Excalidraw draws each text element's "text" exactly as stored, one line per newline:
    # it never re-wraps a free-standing text on load, and it clips whatever overflows the
    # element's width. Whiteboard wraps to the column instead, so the converter breaks the
    # lines itself, the way Chromium lays them out. Widths are Arial's advance widths (in
    # 1/2048 em, code points 32..126) plus Arial's kerning pairs; Excalidraw's Helvetica is
    # drawn with Arial on Windows and has the same metrics elsewhere.
    $ArialAdvance = @(
        569, 569, 727, 1139, 1139, 1821, 1366, 391, 682, 682, 797, 1196, 569, 682, 569, 569,
        1139, 1139, 1139, 1139, 1139, 1139, 1139, 1139, 1139, 1139, 569, 569, 1196, 1196, 1196, 1139,
        2079, 1366, 1366, 1479, 1479, 1366, 1251, 1593, 1479, 569, 1024, 1366, 1139, 1706, 1479, 1593,
        1366, 1593, 1479, 1366, 1251, 1479, 1366, 1933, 1366, 1366, 1251, 569, 569, 569, 961, 1139,
        682, 1139, 1139, 1024, 1139, 1139, 569, 1139, 1139, 455, 455, 1024, 455, 1706, 1139, 1139,
        1139, 1139, 682, 1024, 569, 1139, 1024, 1479, 1024, 1024, 1024, 684, 532, 684, 1196
    )
    $ArialDefaultAdvance = 1139
    $ArialKerning = New-Object 'Collections.Generic.Dictionary[string,int]' ([StringComparer]::Ordinal)
    foreach ($pair in (@(
        '11-152 AT-152 AV-152 AW-76 AY-152 Av-37 Aw-37 Ay-37 F,-227 F.-227 FA-113',
        'LT-152 LV-152 LW-152 LY-152 Ly-76 P,-264 P.-264 PA-152 RT-37 RV-37 RW-37',
        'RY-37 T,-227 T--113 T.-227 T:-227 T;-227 TA-152 TO-37 Ta-227 Tc-227 Te-227',
        'Ti-76 To-227 Tr-76 Ts-227 Tu-76 Tw-113 Ty-113 V,-188 V--113 V.-188 V:-76',
        'V;-76 VA-152 Va-152 Ve-113 Vi-37 Vo-113 Vr-76 Vu-76 Vy-76 W,-113 W--37',
        'W.-113 W:-37 W;-37 WA-76 Wa-76 We-37 Wo-37 Wr-37 Wu-37 Wy-18 Y,-264',
        'Y--188 Y.-264 Y:-113 Y;-133 YA-152 Ya-152 Ye-188 Yi-76 Yo-188 Yp-152 Yq-188',
        'Yu-113 Yv-113 ff-37 r,-113 r.-113 v,-152 v.-152 w,-113 w.-113 y,-152 y.-152'
    ) -join ' ').Split(' ')) {
        $ArialKerning[$pair.Substring(0, 2)] = [int]$pair.Substring(2)
    }

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
        # PowerShell can preserve command output as Object[] even when only one
        # numeric value is expected. Normalize it before doing any arithmetic.
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

    function Get-StyleValue {
        param([string]$Style, [string]$Name)
        return Get-FirstMatch $Style ('(?:^|;)\s*' + [regex]::Escape($Name) + '\s*:\s*([^;]+)')
    }

    function New-Id {
        return ([guid]::NewGuid().ToString('N').Substring(0, 20))
    }

    function New-BaseElement {
        param([string]$Type, [double]$X, [double]$Y, [double]$Width, [double]$Height,
              [string]$Stroke = '#1f1f1f', [string]$Background = 'transparent',
              [string]$StrokeStyle = 'solid', [int]$StrokeWidth = 2)
        return [ordered]@{
            # 1.0, not 1: with an int first argument PowerShell picks Math.Max(int, int) and
            # rounds the size to whole px (a 2794.93 px image became 2795).
            id = New-Id; type = $Type; x = $X; y = $Y
            width = [Math]::Max(1.0, $Width); height = [Math]::Max(1.0, $Height)
            angle = 0; strokeColor = $Stroke; backgroundColor = $Background
            fillStyle = 'solid'; strokeWidth = $StrokeWidth; strokeStyle = $StrokeStyle
            roughness = 0; opacity = 100; groupIds = @(); frameId = $null
            index = $null; roundness = $null; seed = Get-Random -Minimum 1 -Maximum 2147483646
            version = 1; versionNonce = Get-Random -Minimum 1 -Maximum 2147483646
            isDeleted = $false; boundElements = @(); updated = 1; link = $null; locked = $false
        }
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
        # (sqrt(a^2+b^2), sqrt(c^2+d^2)) rather than |a| and |d|: those are only the scale
        # when there is no rotation -- a 90-degree-rotated object has a = d = 0, which used
        # to give it a zero font size and crash the text-height estimate with Infinity.
        # The raw matrix is kept so rotated objects can be placed and rotated correctly.
        param([string]$Style)
        $result = @{ ScaleX = 1.0; ScaleY = 1.0; TranslateX = 0.0; TranslateY = 0.0
                     A = 1.0; B = 0.0; C = 0.0; D = 1.0; IsRotated = $false; IsMirrored = $false }
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
                # A negative determinant is a mirror image, which Excalidraw can't show.
                $result.IsMirrored = ((($v[0] * $v[3]) - ($v[1] * $v[2])) -lt 0)
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
            # Raw CSS matrix terms (a, b, c, d); identity when there is no transform.
            MA = $transform.A; MB = $transform.B; MC = $transform.C; MD = $transform.D
            IsRotated = $transform.IsRotated; IsMirrored = $transform.IsMirrored
        }
    }

    function Get-BoardPoint {
        # Maps a point in the anchor's local (pre-transform) px to board coordinates.
        param($Anchor, [double]$U, [double]$V)
        return ,([double[]]@(($Anchor.X + ($Anchor.MA * $U) + ($Anchor.MC * $V)),
                             ($Anchor.Y + ($Anchor.MB * $U) + ($Anchor.MD * $V))))
    }

    function Get-BoxPlacement {
        # Places a box given in the anchor's local px (U, V = top-left, unrotated). Excalidraw
        # rotates an element about its centre by "angle" (radians, clockwise), so the box
        # centre goes through the anchor's matrix and the scaled box is laid around it.
        # Without rotation this reduces to (X + U*s, Y + V*s).
        param($Anchor, [double]$U, [double]$V, [double]$LocalWidth, [double]$LocalHeight)
        $c = Get-BoardPoint $Anchor ($U + ($LocalWidth / 2)) ($V + ($LocalHeight / 2))
        $w = $LocalWidth * $Anchor.ScaleX; $h = $LocalHeight * $Anchor.ScaleY
        $angle = 0.0
        if ($Anchor.IsRotated) {
            $angle = [Math]::Atan2($Anchor.MB, $Anchor.MA)
            if ($angle -lt 0) { $angle += 2 * [Math]::PI }
        }
        return [pscustomobject]@{ X = ($c[0] - ($w / 2)); Y = ($c[1] - ($h / 2)); Width = $w; Height = $h; Angle = $angle }
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
        # Whiteboard's AzureImage export declares a placeholder "image/*" MIME
        # in the data URI, which Excalidraw won't render. Sniff the real type
        # from the decoded header bytes instead of trusting the declared MIME.
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
        # Whiteboard embeds images at their original source resolution (which can
        # be tens of megapixels) even when the on-canvas element displays at a
        # few hundred units. Excalidraw's own upload flow downscales images
        # before embedding; a hand-built scene file that skips that step can
        # embed something large enough that Excalidraw fails to decode/render
        # it and falls back to a broken-image placeholder. Cap the embedded
        # resolution at 2x the display size (still sharp at 100% zoom) so the
        # scene stays well inside what Excalidraw's normal flow would produce.
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

    function Measure-TextWidth {
        param([string]$Text, [double]$FontSize)
        $units = 0
        for ($i = 0; $i -lt $Text.Length; $i++) {
            $code = [int]$Text[$i]
            if ($code -ge 32 -and $code -le 126) { $units += $ArialAdvance[$code - 32] }
            else { $units += $ArialDefaultAdvance }
            if ($i -gt 0) {
                $kern = 0
                if ($ArialKerning.TryGetValue($Text.Substring($i - 1, 2), [ref]$kern)) { $units += $kern }
            }
        }
        return $units * $FontSize / 2048
    }

    function Get-WrappedLines {
        # Breaks at spaces and after hyphens; trailing spaces hang past the column edge; a
        # word longer than the column is broken between characters (overflow-wrap).
        param([string]$Text, [double]$MaxWidth, [double]$FontSize)
        $lines = New-Object 'Collections.Generic.List[string]'
        $limit = $MaxWidth + 0.01
        foreach ($paragraph in ($Text -split "\r?\n")) {
            $line = ''
            foreach ($m in [regex]::Matches($paragraph, '[^\s-]*-+(?=[^\s-])|\S+|\s+')) {
                $token = $m.Value
                if ($token -match '^\s+$') { $line += $token; continue }
                if ((Measure-TextWidth ($line + $token).TrimEnd() $FontSize) -le $limit) { $line += $token; continue }
                if ($line.Trim().Length -gt 0) { [void]$lines.Add($line.TrimEnd()); $line = '' }
                foreach ($ch in $token.ToCharArray()) {
                    if ($line.Length -gt 0 -and (Measure-TextWidth ($line + $ch) $FontSize) -gt $limit) {
                        [void]$lines.Add($line); $line = ''
                    }
                    $line += $ch
                }
            }
            [void]$lines.Add($line.TrimEnd())
        }
        return ,$lines
    }

    function Get-TextBlockHeight {
        param([string]$Text, [double]$Width, [double]$FontSize)
        $count = (Get-WrappedLines $Text $Width $FontSize).Count
        return [Math]::Max(1, $count) * $FontSize * $TextLineHeight
    }

    function Add-TextElement {
        param(
            [Collections.Generic.List[object]]$Elements, [string]$Text,
            [double]$X, [double]$Y, [double]$Width, [double]$Height,
            [double]$FontSize, [string]$Color, [string]$Align = 'left',
            [double]$Angle = 0,
            [double]$LineHeight = 0,  # 0 = $TextLineHeight (Whiteboard's Arial texts)
            [bool]$Underline = $false
        )
        if ([string]::IsNullOrEmpty($Text)) { return }
        $X = Get-ScalarDouble $X; $Y = Get-ScalarDouble $Y
        $Width = Get-ScalarDouble $Width 1; $Height = Get-ScalarDouble $Height 1
        $FontSize = Get-ScalarDouble $FontSize 20
        $lineHeight = if ($LineHeight -gt 0) { $LineHeight } else { $TextLineHeight }
        $fontFamily = 2 # Helvetica/sans-serif in Excalidraw
        $estimatedWidth = [Math]::Max(1.0, [Math]::Min($Width, $Text.Length * $FontSize * 0.58))
        if ($Width -le 1) { $Width = $estimatedWidth }
        if ($Height -le 1) {
            $charsPerLine = [Math]::Max(1, [Math]::Floor($Width / ($FontSize * 0.58)))
            $lines = [Math]::Max(1, [Math]::Ceiling($Text.Length / $charsPerLine))
            $Height = $lines * $FontSize * $lineHeight
        }
        $wrapped = (Get-WrappedLines $Text $Width $FontSize) -join "`n"
        if ($wrapped.Trim().Length -eq 0) { $wrapped = $Text }   # whitespace-only boxes keep their spaces
        # A line that can't wrap (a single glyph wider than the column, e.g. a huge "W") overflows
        # its column in Whiteboard, but Excalidraw clips text to the element's width. Widen the
        # element to the widest line around its alignment anchor; the centre moves along the
        # element's own x axis, so rotated text stays put too.
        $widest = 0.0
        foreach ($line in ($wrapped -split "`n")) { $widest = [Math]::Max($widest, (Measure-TextWidth $line $FontSize)) }
        if ($widest -gt $Width) {
            $extra = $widest - $Width
            $shift = switch ($Align) { 'center' { 0.0 } 'right' { -$extra / 2 } default { $extra / 2 } }
            $cx = $X + ($Width / 2) + ($shift * [Math]::Cos($Angle))
            $cy = $Y + ($Height / 2) + ($shift * [Math]::Sin($Angle))
            $Width = $widest
            $X = $cx - ($Width / 2); $Y = $cy - ($Height / 2)
        }
        $e = New-BaseElement 'text' $X $Y $Width $Height $Color 'transparent' 'solid' 1
        $e.fontSize = $FontSize; $e.fontFamily = $fontFamily
        $e.text = $wrapped; $e.rawText = $Text; $e.originalText = $Text
        $e.textAlign = $Align; $e.verticalAlign = 'top'; $e.containerId = $null
        $e.autoResize = $false; $e.lineHeight = $lineHeight; $e.angle = $Angle
        if ($Underline) { Add-Underlines $Elements $e $wrapped $Align }   # sets $e.groupIds, so first
        [void]$Elements.Add([pscustomobject]$e)
    }

    function Get-FontFamily {
        # The first family of an inline CSS font-family list. Older exports' plain text is
        # "sans-serif, "Segoe UI"", and the generic sans-serif renders as Arial; newer exports
        # (CompareAndContrast2) name "Segoe UI" itself, which is wider -- drawing it in Arial
        # moved centred labels up to 21px.
        param([string]$Style, [string]$Default = 'Arial')
        $v = "$(Get-StyleValue ([Net.WebUtility]::HtmlDecode($Style)) 'font-family')"
        $first = ($v -split ',')[0].Trim().Trim([char[]]@([char]34, [char]39)).Trim()
        if (-not $first -or $first -eq 'sans-serif') { return $Default }
        return $first
    }

    function Test-Underline {
        # Whiteboard underlines a whole text box with text-decoration on its textBoxCore div.
        param([string]$Style)
        return "$(Get-StyleValue $Style 'text-decoration')" -match '\bunderline\b'
    }

    function Add-Underlines {
        # Excalidraw text has no underline, so each wrapped line gets a drawn one, grouped with
        # the text so they move together. Chromium centres Arial's underline 217/2048 x font size
        # below the baseline ($UnderlineOffset); the first baseline sits where the comment on
        # $TextBaselineShift says. Each line is placed in the text's own frame and turned about
        # the text's centre, so rotated text keeps its underline. customData marks the lines so
        # the visual checks don't count them as connectors.
        param([Collections.Generic.List[object]]$Elements, $Text, [string]$Wrapped, [string]$Align)
        $fs = [double]$Text.fontSize; $lh = [double]$Text.lineHeight
        $groupId = New-Id
        $Text.groupIds = @($groupId)
        $cos = [Math]::Cos($Text.angle); $sin = [Math]::Sin($Text.angle)
        $cx = $Text.x + ($Text.width / 2); $cy = $Text.y + ($Text.height / 2)
        $lines = @($Wrapped -split "`n")
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $w = Measure-TextWidth $lines[$i].TrimEnd() $fs
            if ($w -le 0) { continue }
            $left = switch ($Align) { 'center' { ($Text.width - $w) / 2 } 'right' { $Text.width - $w } default { 0.0 } }
            $ly = (((($lh - 1) / 2) + (1577 / 2048)) * $fs) + ($i * $lh * $fs) + ($UnderlineOffset * $fs)
            $dx = $left + ($w / 2) - ($Text.width / 2); $dy = $ly - ($Text.height / 2)
            $mx = $cx + ($dx * $cos) - ($dy * $sin); $my = $cy + ($dx * $sin) + ($dy * $cos)
            $u = New-BaseElement 'line' ($mx - ($w / 2)) $my $w 0 $Text.strokeColor
            $u.height = 0; $u.strokeWidth = [Math]::Max(1.0, $fs * $UnderlineThickness); $u.angle = $Text.angle
            $pointList = New-Object Collections.ArrayList
            [void]$pointList.Add([double[]]@(0, 0)); [void]$pointList.Add([double[]]@($w, 0))
            $u.groupIds = @($groupId); $u.points = $pointList
            $u.lastCommittedPoint = $null; $u.startBinding = $null; $u.endBinding = $null
            $u.startArrowhead = $null; $u.endArrowhead = $null; $u.customData = @{ underline = $true }
            [void]$Elements.Add([pscustomobject]$u)
        }
    }

    function Add-ReactionSticker {
        param(
            [Collections.Generic.List[object]]$Elements, [string]$Label,
            [double]$X, [double]$Y, [double]$Width, [double]$Height
        )
        $X = Get-ScalarDouble $X; $Y = Get-ScalarDouble $Y
        $Width = Get-ScalarDouble $Width 40; $Height = Get-ScalarDouble $Height 40
        $groupId = New-Id
        $size = [Math]::Max(16.0, [Math]::Min($Width, $Height))

        if ($Label -match 'red cross') {
            foreach ($direction in @('down', 'up')) {
                $e = New-BaseElement 'line' $X $Y $size $size '#e62e55' 'transparent' 'solid' 4
                $pointList = New-Object Collections.ArrayList
                if ($direction -eq 'down') {
                    [void]$pointList.Add([double[]]@(0, 0)); [void]$pointList.Add([double[]]@($size, $size))
                } else {
                    [void]$pointList.Add([double[]]@($size, 0)); [void]$pointList.Add([double[]]@(0, $size))
                }
                $e.groupIds = @($groupId); $e.points = $pointList
                $e.lastCommittedPoint = $null; $e.startBinding = $null; $e.endBinding = $null
                $e.startArrowhead = $null; $e.endArrowhead = $null
                [void]$Elements.Add([pscustomobject]$e)
            }
            return
        }

        if ($Label -match 'green check') {
            $e = New-BaseElement 'line' $X $Y $size $size '#4fc24f' 'transparent' 'solid' 4
            $e.groupIds = @($groupId)
            # Precompute these values. In Windows PowerShell, a multiplication
            # inside an array expression such as @($size * .35, $size) can be
            # parsed as multiplication by the two-item Object[] on its right.
            $checkMidY = [double]($size * 0.55)
            $checkMidX = [double]($size * 0.35)
            $pointList = New-Object Collections.ArrayList
            [void]$pointList.Add([double[]]@(0, $checkMidY))
            [void]$pointList.Add([double[]]@($checkMidX, $size))
            [void]$pointList.Add([double[]]@($size, 0))
            $e.points = $pointList
            $e.lastCommittedPoint = $null; $e.startBinding = $null; $e.endBinding = $null
            $e.startArrowhead = $null; $e.endArrowhead = $null
            [void]$Elements.Add([pscustomobject]$e)
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
        Add-TextElement $Elements ([string]$symbol) $X $Y $size $size ($size * 0.9) $color 'center'
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

    function New-InkElement {
        # One ink stroke (Get-InkStrokes) as an open Excalidraw line through its centreline, at
        # the pen's width; customData marks it so the visual checks can find it.
        param($Stroke)
        $pts = $Stroke.Points
        $x0 = $pts[0][0]; $y0 = $pts[0][1]
        $pointList = New-Object Collections.ArrayList
        foreach ($p in $pts) { [void]$pointList.Add([double[]]@(($p[0] - $x0), ($p[1] - $y0))) }
        $xs = @($pts | ForEach-Object { $_[0] }); $ys = @($pts | ForEach-Object { $_[1] })
        $bx = ($xs | Measure-Object -Minimum -Maximum); $by = ($ys | Measure-Object -Minimum -Maximum)
        $e = New-BaseElement 'line' $x0 $y0 ($bx.Maximum - $bx.Minimum) ($by.Maximum - $by.Minimum) $Stroke.Color
        $e.strokeWidth = $Stroke.Width
        $e.points = $pointList; $e.lastCommittedPoint = $null
        $e.startBinding = $null; $e.endBinding = $null; $e.startArrowhead = $null; $e.endArrowhead = $null
        $e.customData = @{ ink = $true }
        return $e
    }
    function Get-ShapePathPoints {
        # Traces the shape's own outline from Whiteboard's SVG <path d="..."> (the first,
        # visible path inside the shape's <g>), flattening curves into points, and maps it
        # with (X, Y) + point * scale. This keeps the true geometry of ovals, triangles etc.
        # instead of guessing it from the aria-label. Returns $null (so the caller falls back
        # to the label) when the path uses something not handled here: arcs (A), smooth
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
        if ($pts.Count -ge 2) {   # drop the closing duplicate; callers close the outline themselves
            $first = $pts[0]; $last = $pts[$pts.Count - 1]
            if ([Math]::Abs($first[0] - $last[0]) -lt 0.01 -and [Math]::Abs($first[1] - $last[1]) -lt 0.01) { $pts.RemoveAt($pts.Count - 1) }
        }
        if ($pts.Count -lt 3) { return $null }
        $mapped = New-Object Collections.Generic.List[double[]]
        foreach ($p in $pts) {
            [void]$mapped.Add([double[]]@(($X + (($p[0] + $tx) * $ScaleX)), ($Y + (($p[1] + $ty) * $ScaleY))))
        }
        return ,($mapped.ToArray())
    }

    function Get-RasterImageDimensions {
        # Native size of an embedded image, read from the file header (PNG, GIF, BMP, JPEG)
        # or from an SVG's root width/height (else its viewBox). Used to size reaction
        # stickers, whose <img> is shrunk by CSS max-width/max-height but never enlarged.
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

        # SVG: the root width/height (plain user units / px only), else the viewBox size.
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
        # Decodes the first <img src="data:..."> in an anchor block into raw bytes plus a MIME
        # type Excalidraw accepts. Handles both base64 and URL-encoded (plain text) data URIs,
        # and sniffs the real format when the declared MIME type is a placeholder such as
        # image/* . Returns $null when there is no usable inline image.
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
        if ($mime -ne 'image/svg+xml' -and $mime -notmatch '^image/(png|jpe?g|gif|bmp|webp)$') {
            $head = [Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min($bytes.Length, 1024))
            if ($head -match '<svg\b') { $mime = 'image/svg+xml' }
            else { $mime = Get-ImageMimeType ([Convert]::ToBase64String($bytes, 0, [Math]::Min($bytes.Length, 48))) }
        }
        if ($mime -eq 'image/jpg') { $mime = 'image/jpeg' }
        return [pscustomobject]@{ Bytes = $bytes; Mime = $mime }
    }

    function Convert-OneFile {
        param([string]$SourcePath, [string]$DestinationPath)
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

        $elements = [Collections.Generic.List[object]]::new()
        $files = [ordered]@{}
        $unsupported = [Collections.Generic.List[string]]::new()
        $sourceTypes = @{}
        $imageIndex = 0
        $imageFilesWritten = [Collections.Generic.List[string]]::new()
        $destBaseName = [IO.Path]::GetFileNameWithoutExtension($DestinationPath)
        $imageOutputDir = Join-Path (Split-Path -Parent $DestinationPath) "${destBaseName}_images"
        $stickerImageCount = 0; $stickerFallbackCount = 0
        $shapeTracedCount = 0; $shapeFallbackCount = 0
        $rotatedCount = 0; $underlinedCount = 0; $mirrorIgnored = [Collections.Generic.List[string]]::new()
        $arrowheadCount = 0
        $noteColors = Get-NoteColorTable $html
        $noteColorMissing = [Collections.Generic.List[string]]::new()
        $inkStrokeCount = 0; $inkApproximated = 0; $fontSubstituted = 0

        foreach ($block in $blocks) {
            $a = Get-AnchorInfo $block
            if (-not $a.Type) { continue }
            if (-not $sourceTypes.ContainsKey($a.Type)) { $sourceTypes[$a.Type] = 0 }
            $sourceTypes[$a.Type]++
            # Excalidraw rotates any element (its "angle"), but can't mirror one: a mirrored
            # object is drawn with the right rotation and size but unmirrored.
            if ($a.IsRotated) { $rotatedCount++ }
            # Ink is mapped through the full matrix, so it keeps its mirroring too.
            if ($a.IsMirrored -and $a.Type -ne 'InkGroup') { [void]$mirrorIgnored.Add("$($a.Type):$($a.Key)") }

            switch ($a.Type) {
                'Shape' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*\bclass="[^"]*\bshape\b[^"]*"[^>]*>)'
                    if (-not $svgTag) { $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)' }
                    # Everything below is in the anchor's local (pre-transform) px; the anchor's
                    # matrix is applied by Get-BoxPlacement / Get-BoardPoint.
                    $lw = Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 100
                    $lh = Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 100
                    $u0 = 0.0; $v0 = 0.0
                    if ($a.IsCentered) { $u0 = -$lw / 2; $v0 = -$lh / 2 }
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $fill = Convert-RgbaToHex (Get-FirstMatch $gTag '\bfill="([^"]+)"') '#ffffff' -TransparentAllowed
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f' -TransparentAllowed
                    $dash = Get-FirstMatch $gTag '\bstroke-dasharray="([^"]+)"'
                    $strokeStyle = if ($dash) { 'dashed' } else { 'solid' }
                    $shapeName = Get-FirstMatch $svgTag 'aria-label="\s*([^,."]+)'
                    # Whiteboard labels circles/ellipses "oval" -- previously unmatched here,
                    # which turned every oval into a rectangle.
                    $kind = if ($shapeName -match 'ellipse|circle|oval') { 'ellipse' }
                            elseif ($shapeName -match 'diamond') { 'diamond' }
                            elseif ($shapeName -match 'rectangle|square') { 'rectangle' }
                            else { $null }
                    # Prefer Whiteboard's own outline path: its bounds are the shape's true
                    # geometry (the <svg> box also includes the stroke), and shapes Excalidraw
                    # has no primitive for are drawn from it as a closed polygon.
                    $outline = Get-ShapePathPoints $block $u0 $v0 1 1
                    if ($null -ne $outline) {
                        $shapeTracedCount++
                        $xs = @($outline | ForEach-Object { $_[0] }); $ys = @($outline | ForEach-Object { $_[1] })
                        $bu = ($xs | Measure-Object -Minimum -Maximum); $bv = ($ys | Measure-Object -Minimum -Maximum)
                        $boxU = $bu.Minimum; $boxV = $bv.Minimum
                        $boxW = $bu.Maximum - $bu.Minimum; $boxH = $bv.Maximum - $bv.Minimum
                    } else {
                        $shapeFallbackCount++
                        $boxU = $u0; $boxV = $v0; $boxW = $lw; $boxH = $lh
                        if (-not $kind) { $kind = 'rectangle' }
                    }
                    if ($kind) {
                        $p = Get-BoxPlacement $a $boxU $boxV $boxW $boxH
                        $e = New-BaseElement $kind $p.X $p.Y $p.Width $p.Height $stroke $fill $strokeStyle 2
                        $e.angle = $p.Angle
                    } elseif (-not $a.IsMirrored) {
                        # Closed polygon through the traced outline, built unrotated (scaled only)
                        # inside the outline's box and turned by "angle" -- Excalidraw rotates a
                        # line about the centre of its points' bounds, which is the box centre --
                        # so its selection box and rotate handle follow it like any rotated shape.
                        $p = Get-BoxPlacement $a $boxU $boxV $boxW $boxH
                        $local = @($outline | ForEach-Object { ,([double[]]@((($_[0] - $boxU) * $a.ScaleX), (($_[1] - $boxV) * $a.ScaleY))) })
                        $pointList = New-Object Collections.ArrayList
                        foreach ($lp in $local) { [void]$pointList.Add([double[]]@(($lp[0] - $local[0][0]), ($lp[1] - $local[0][1]))) }
                        [void]$pointList.Add([double[]]@(0, 0))
                        $e = New-BaseElement 'line' ($p.X + $local[0][0]) ($p.Y + $local[0][1]) $p.Width $p.Height $stroke $fill $strokeStyle 2
                        $e.angle = $p.Angle
                        $e.points = $pointList; $e.polygon = $true; $e.lastCommittedPoint = $null
                        $e.startBinding = $null; $e.endBinding = $null; $e.startArrowhead = $null; $e.endArrowhead = $null
                    } else {
                        # Mirrored: Excalidraw can't mirror an element, so the outline is mapped
                        # through the full matrix instead (right geometry, axis-aligned box).
                        $board = @($outline | ForEach-Object { Get-BoardPoint $a $_[0] $_[1] })
                        $x0 = $board[0][0]; $y0 = $board[0][1]
                        $pointList = New-Object Collections.ArrayList
                        foreach ($bp in $board) { [void]$pointList.Add([double[]]@(($bp[0] - $x0), ($bp[1] - $y0))) }
                        [void]$pointList.Add([double[]]@(0, 0))
                        $pxs = @($board | ForEach-Object { $_[0] }); $pys = @($board | ForEach-Object { $_[1] })
                        $pw = ($pxs | Measure-Object -Minimum -Maximum); $ph = ($pys | Measure-Object -Minimum -Maximum)
                        $e = New-BaseElement 'line' $x0 $y0 ($pw.Maximum - $pw.Minimum) ($ph.Maximum - $ph.Minimum) $stroke $fill $strokeStyle 2
                        $e.points = $pointList; $e.polygon = $true; $e.lastCommittedPoint = $null
                        $e.startBinding = $null; $e.endBinding = $null; $e.startArrowhead = $null; $e.endArrowhead = $null
                    }
                    [void]$elements.Add([pscustomobject]$e)

                    # Whiteboard stores labels such as the blue-box "test" text
                    # inside the Shape block. V1 converted only the outline.
                    $shapeText = Get-HtmlText $block
                    if (-not [string]::IsNullOrEmpty($shapeText)) {
                        $shapeTextStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+shapeText\b[^"]*"[^>]*style="([^"]*)"'
                        $shapeCoreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                        $fontLocal = Get-CssNumber $shapeTextStyle 'font-size' 20
                        if ((Get-FontFamily $shapeCoreStyle) -ne 'Arial') { $fontSubstituted++ }
                        $innerWidth = Get-CssNumber $shapeTextStyle 'width' ([Math]::Max(1.0, $lw - 26))
                        $innerHeight = Get-CssNumber $shapeTextStyle 'height' ([Math]::Max(1.0, $lh - 26))
                        $textHeight = [Math]::Min($innerHeight, (Get-TextBlockHeight $shapeText $innerWidth $fontLocal))
                        $p = Get-BoxPlacement $a ($u0 + (($lw - $innerWidth) / 2)) ($v0 + (($lh - $textHeight) / 2)) $innerWidth $textHeight
                        $textColor = Convert-RgbaToHex (Get-StyleValue $shapeCoreStyle 'color') '#000000'
                        $textAlign = if ($block -match 'DraftEditor-alignRight') { 'right' } elseif ($block -match 'DraftEditor-alignCenter') { 'center' } else { 'left' }
                        $underline = Test-Underline $shapeCoreStyle
                        if ($underline) { $underlinedCount++ }
                        Add-TextElement $elements $shapeText $p.X $p.Y $p.Width $p.Height ($fontLocal * $a.ScaleY) $textColor $textAlign $p.Angle -Underline $underline
                    }
                }
                'PlainText' {
                    $text = Get-HtmlText $block
                    $textBoxStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+plainText\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $outerStyle = Get-FirstMatch $block '<div\s+style="([^"]*width:[^"]*display:\s*flex[^"]*)"'
                    $fontLocal = Get-CssNumber $textBoxStyle 'font-size' 20
                    if ($text.Trim() -and (Get-FontFamily $coreStyle) -ne 'Arial') { $fontSubstituted++ }
                    # Whiteboard insets the text inside every plain-text box (see
                    # $PlainTextInsetLeft): the glyphs start at local (17, 16), and the column
                    # is the outer width minus both side insets. Placing the box at the anchor
                    # point put every text ~17*s px left and ~16*s px high.
                    $width = Get-CssNumber $outerStyle 'width' 0
                    if ($width -le 0) { $width = Get-CssNumber $textBoxStyle 'max-width' 0 }
                    if ($width -gt 0) { $width = [Math]::Max(1.0, $width - $PlainTextInsetLeft - $PlainTextInsetRight) }
                    else { $width = [Math]::Max(20.0, $text.Length * $fontLocal * 0.58) }
                    $height = Get-TextBlockHeight $text $width $fontLocal
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $align = if ($block -match 'DraftEditor-alignCenter') { 'center' } elseif ($block -match 'DraftEditor-alignRight') { 'right' } else { 'left' }
                    $p = Get-BoxPlacement $a $PlainTextInsetLeft ($PlainTextInsetTop + ($TextBaselineShift * $fontLocal)) $width $height
                    $underline = Test-Underline $coreStyle
                    if ($underline -and $text.Trim()) { $underlinedCount++ }
                    Add-TextElement $elements $text $p.X $p.Y $p.Width $p.Height ($fontLocal * $a.ScaleY) $color $align $p.Angle -Underline $underline
                }
                'Note' {
                    $text = Get-HtmlText $block
                    $noteStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+stickyNote\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $note = Get-NoteLayout $block $noteColors
                    if (-not $note.ColorFound) { [void]$noteColorMissing.Add("$($a.Type):$($a.Key)") }
                    $lw = $note.CssWidth; $lh = $note.CssHeight
                    $boxW = $note.Width; $boxH = $note.Height
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $p = Get-BoxPlacement $a 0 0 $boxW $boxH
                    $e = New-BaseElement 'rectangle' $p.X $p.Y $p.Width $p.Height 'transparent' $note.Solid 'solid' 1
                    $e.angle = $p.Angle
                    $e.roundness = @{ type = 3 }; [void]$elements.Add([pscustomobject]$e)
                    if ($text.Trim()) {
                        $noteFont = Get-CssNumber $noteStyle 'font-size' 24
                        $noteColumn = $lw - $NoteTextColumnInset
                        $noteLines = (Get-WrappedLines $text $noteColumn $noteFont).Count
                        $p = Get-BoxPlacement $a $note.TextLeft ($note.TextTop + ($NoteBaselineShift * $noteFont)) $noteColumn `
                            ([Math]::Min($lh - $note.Border - 12, [Math]::Max(1, $noteLines) * $noteFont * $NoteLineHeight))
                        $underline = Test-Underline $coreStyle
                        if ($underline) { $underlinedCount++ }
                        Add-TextElement $elements $text $p.X $p.Y $p.Width $p.Height ($noteFont * $a.ScaleY) $color 'left' $p.Angle -LineHeight $NoteLineHeight -Underline $underline
                    }
                }
                'InkGroup' {
                    $strokes = @(Get-InkStrokes $a $block)
                    if ($strokes.Count -eq 0) { [void]$unsupported.Add("$($a.Type):$($a.Key) (no ink strokes)"); break }
                    # One group for all strokes of the ink object, so they move together.
                    $inkGroup = New-Id
                    foreach ($s in $strokes) {
                        $e = New-InkElement $s
                        if ($strokes.Count -gt 1) { $e.groupIds = @($inkGroup) }
                        [void]$elements.Add([pscustomobject]$e)
                        $inkStrokeCount++
                        if (-not $s.Exact) { $inkApproximated++ }
                    }
                }
                'Connector' {
                    $svgTag = Get-FirstMatch $block '(<svg\b[^>]*>)'
                    $w = Get-Number (Get-FirstMatch $svgTag '\bwidth="([0-9.]+)"') 10
                    $h = Get-Number (Get-FirstMatch $svgTag '\bheight="([0-9.]+)"') 10
                    $gTag = Get-FirstMatch $block '(<g\b[^>]*>)'
                    $path = Get-FirstMatch $block '<path\s+d="([^"]+)"'
                    $nums = @([regex]::Matches($path, '-?[0-9.]+') | ForEach-Object { Get-Number $_.Value })
                    if ($nums.Count -ge 4) {
                        $x1=$nums[0]; $y1=$nums[1]; $x2=$nums[2]; $y2=$nums[3]
                    } else { $x1=0; $y1=0; $x2=$w; $y2=$h }
                    # Endpoints go through the anchor's matrix, so scaled and rotated
                    # connectors land where Whiteboard draws them. Excalidraw measures a linear
                    # element's points from its first point, which sits at (x, y).
                    $p1 = Get-BoardPoint $a $x1 $y1; $p2 = Get-BoardPoint $a $x2 $y2
                    $dx = $p2[0] - $p1[0]; $dy = $p2[1] - $p1[1]
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f'
                    $style = if ($gTag -match 'stroke-dasharray') { 'dashed' } else { 'solid' }
                    # Arrowheads: Whiteboard draws each as its own open chevron path inside the
                    # connector's <g>, e.g. d="M-5 -9 L0 0 L5 -9" transform="translate(x,y), rotate(deg)"
                    # (tip at the local origin). Previously only the first <path> was read, so
                    # every arrowhead was dropped. Each tip is matched to the nearer endpoint
                    # and becomes that end's Excalidraw arrowhead ("arrow" = open chevron).
                    $startHead = $null; $endHead = $null
                    $arrowRx = '<path\s+d="[^"]+"\s+transform="\s*translate\(\s*([-+0-9.eE]+)[\s,]+([-+0-9.eE]+)\s*\)\s*,?\s*rotate\(\s*[-+0-9.eE]+\s*\)\s*"'
                    foreach ($am in [regex]::Matches($block, $arrowRx, [Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                        $tipX = Get-Number $am.Groups[1].Value 0; $tipY = Get-Number $am.Groups[2].Value 0
                        $toStart = [Math]::Pow($tipX - $x1, 2) + [Math]::Pow($tipY - $y1, 2)
                        $toEnd = [Math]::Pow($tipX - $x2, 2) + [Math]::Pow($tipY - $y2, 2)
                        if ($toStart -lt $toEnd) { $startHead = 'arrow' } else { $endHead = 'arrow' }
                        $arrowheadCount++
                    }
                    # Excalidraw only draws arrowheads on "arrow" elements, not on "line".
                    $lineType = if ($startHead -or $endHead) { 'arrow' } else { 'line' }
                    $e = New-BaseElement $lineType $p1[0] $p1[1] ([Math]::Abs($dx)) ([Math]::Abs($dy)) $stroke 'transparent' $style 2
                    $pointList = New-Object Collections.ArrayList
                    [void]$pointList.Add([double[]]@(0,0)); [void]$pointList.Add([double[]]@($dx,$dy))
                    $e.points = $pointList; $e.lastCommittedPoint = $null
                    $e.startBinding=$null; $e.endBinding=$null; $e.startArrowhead=$startHead; $e.endArrowhead=$endHead
                    if ($lineType -eq 'arrow') { $e.elbowed = $false }
                    [void]$elements.Add([pscustomobject]$e)
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
                    $u0 = 0.0; $v0 = 0.0
                    if ($a.IsCentered) { $u0 = -$boxW / 2; $v0 = -$boxH / 2 }
                    $p = Get-BoxPlacement $a $u0 $v0 $boxW $boxH
                    if ($stickerData) {
                        # Embed Whiteboard's own sticker artwork (an SVG for every built-in
                        # sticker) so any sticker type -- heart, fire, like, light bulb, and any
                        # added later -- converts faithfully, instead of mapping known labels
                        # to hand-drawn stand-ins and printing '?' for everything else.
                        $fileId = New-Id
                        $dataUrl = 'data:{0};base64,{1}' -f $stickerData.Mime, [Convert]::ToBase64String($stickerData.Bytes)
                        $files[$fileId] = [ordered]@{ mimeType=$stickerData.Mime; id=$fileId; dataURL=$dataUrl; created=1; lastRetrieved=1 }
                        $e = New-BaseElement 'image' $p.X $p.Y $p.Width $p.Height 'transparent' 'transparent' 'solid' 1
                        $e.angle = $p.Angle
                        $e.fileId=$fileId; $e.status='saved'; $e.scale=@(1,1); $e.crop=$null
                        [void]$elements.Add([pscustomobject]$e)
                        $stickerImageCount++
                    } else {
                        Add-ReactionSticker $elements $label $p.X $p.Y $p.Width $p.Height
                        $stickerFallbackCount++
                    }
                }
                {$_ -in 'Image', 'AzureImage'} {
                    # Newer Whiteboard exports tag images as "AzureImage" rather
                    # than "Image", but use the same imageComponent/img markup.
                    $src = Get-FirstMatch $block '<img\b[^>]*\bsrc="([^"]+)"'
                    if (-not $src) { [void]$unsupported.Add("$($a.Type):$($a.Key) (missing image data)"); break }
                    $src = [Net.WebUtility]::HtmlDecode($src)
                    $size = Get-ImageSize $block
                    $u0 = 0.0; $v0 = 0.0
                    if ($a.IsCentered) { $u0 = -$size.Width / 2; $v0 = -$size.Height / 2 }
                    $p = Get-BoxPlacement $a $u0 $v0 $size.Width $size.Height
                    $x = $p.X; $y = $p.Y; $w = $p.Width; $h = $p.Height
                    $mime = Get-FirstMatch $src '^data:([^;,]+)'
                    if (-not $mime) { [void]$unsupported.Add("$($a.Type):$($a.Key) (external image URL)"); break }
                    $commaIndex = $src.IndexOf(',')
                    $payload = if ($commaIndex -gt 0) { $src.Substring($commaIndex + 1) } else { $null }
                    if ($payload -and $mime -notmatch '^image/(png|jpe?g|gif|bmp|webp|svg\+xml)$') {
                        $mime = Get-ImageMimeType $payload
                    }
                    if ($payload -and $mime -match '^image/(png|jpe?g|gif|bmp)$') {
                        $downscaled = Get-DownscaledImagePayload $payload $w $h
                        if ($downscaled) { $payload = $downscaled; $mime = 'image/png' }
                    }
                    if ($payload) { $src = "data:$mime;base64,$payload" }
                    $fileId = New-Id
                    $files[$fileId] = [ordered]@{ mimeType=$mime; id=$fileId; dataURL=$src; created=1; lastRetrieved=1 }
                    $e = New-BaseElement 'image' $x $y $w $h 'transparent' 'transparent' 'solid' 1
                    $e.angle = $p.Angle
                    $e.fileId=$fileId; $e.status='saved'; $e.scale=@(1,1); $e.crop=$null
                    [void]$elements.Add([pscustomobject]$e)

                    # Also drop a standalone copy of the same bytes next to the
                    # .excalidraw output. Excalidraw itself can't reference an
                    # external file (only an embedded data URI, which is what
                    # actually makes the image show up above) - this copy is
                    # purely for the user's own reference/backup.
                    if ($payload) {
                        $imageIndex++
                        $extension = switch -Regex ($mime) {
                            'jpe?g' { 'jpg' } 'gif' { 'gif' } 'bmp' { 'bmp' } 'webp' { 'webp' } 'svg' { 'svg' }
                            default { 'png' }
                        }
                        if (-not (Test-Path -LiteralPath $imageOutputDir)) {
                            [void](New-Item -ItemType Directory -Path $imageOutputDir)
                        }
                        $imageFilePath = Join-Path $imageOutputDir ("{0}_image{1}.{2}" -f $destBaseName, $imageIndex, $extension)
                        [IO.File]::WriteAllBytes($imageFilePath, [Convert]::FromBase64String($payload))
                        [void]$imageFilesWritten.Add($imageFilePath)
                    }
                }
                default { [void]$unsupported.Add("$($a.Type):$($a.Key)") }
            }
        }

        if ($elements.Count -eq 0) { throw "No convertible objects were found in '$SourcePath'." }
        $minX = ($elements | Measure-Object -Property x -Minimum).Minimum
        $minY = ($elements | Measure-Object -Property y -Minimum).Minimum
        foreach ($e in $elements) {
            $e.x = [Math]::Round(([double]$e.x - $minX + $CanvasPadding), 3)
            $e.y = [Math]::Round(([double]$e.y - $minY + $CanvasPadding), 3)
            $e.width = [Math]::Round([double]$e.width, 3)
            $e.height = [Math]::Round([double]$e.height, 3)
        }

        $scene = [ordered]@{
            type='excalidraw'; version=2; source='https://github.com/excalidraw/excalidraw'
            elements=@($elements)
            appState=[ordered]@{ gridSize=$null; viewBackgroundColor='#ffffff' }
            files=$files
        }
        $json = $scene | ConvertTo-Json -Depth 100
        $utf8NoBom = New-Object Text.UTF8Encoding($false)
        [IO.File]::WriteAllText($DestinationPath, $json, $utf8NoBom)

        [pscustomobject]@{
            InputPath=$SourcePath; OutputPath=$DestinationPath
            SourceObjects=$blocks.Count; ExcalidrawElements=$elements.Count
            EmbeddedFiles=$files.Count; UnsupportedObjects=$unsupported.Count
            UnsupportedDetails=@($unsupported); SourceTypeCounts=$sourceTypes
            ImageFiles=@($imageFilesWritten)
            StickersEmbedded=$stickerImageCount; StickersFallback=$stickerFallbackCount
            ShapesTraced=$shapeTracedCount; ShapesFromLabel=$shapeFallbackCount
            RotatedObjects=$rotatedCount; ArrowheadsConverted=$arrowheadCount
            UnderlinedTexts=$underlinedCount
            MirrorIgnored=$mirrorIgnored.Count; MirrorIgnoredDetails=@($mirrorIgnored)
            NoteColorMissing=$noteColorMissing.Count; NoteColorMissingDetails=@($noteColorMissing)
            InkStrokes=$inkStrokeCount; InkApproximated=$inkApproximated; FontSubstituted=$fontSubstituted
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
            $destination = Join-Path $directory (([IO.Path]::GetFileNameWithoutExtension($source)) + '.excalidraw')
        } else {
            $destination = [IO.Path]::ChangeExtension($source, '.excalidraw')
        }
        if ((Test-Path -LiteralPath $destination) -and -not $Force) {
            throw "Output exists: '$destination'. Use -Force to overwrite it."
        }
        $result = Convert-OneFile $source $destination
        Write-Host ("Converted {0} source objects to {1} Excalidraw elements: {2}" -f `
            $result.SourceObjects, $result.ExcalidrawElements, $destination)
        if ($result.ImageFiles.Count -gt 0) {
            Write-Host ("  Also wrote {0} standalone image file(s) to: {1}" -f `
                $result.ImageFiles.Count, (Split-Path -Parent $result.ImageFiles[0]))
        }
        if ($result.UnderlinedTexts -gt 0) {
            Write-Host ("  {0} underlined text(s) were given a drawn underline (Excalidraw text has no underline)." -f $result.UnderlinedTexts)
        }
        if ($result.StickersFallback -gt 0) {
            Write-Warning ("{0} reaction sticker(s) had no embedded artwork and were drawn with a fallback symbol." -f $result.StickersFallback)
        }
        if ($result.ShapesFromLabel -gt 0) {
            Write-Warning ("{0} shape(s) had an outline that couldn't be traced and were built from their label." -f $result.ShapesFromLabel)
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
        if ($result.FontSubstituted -gt 0) {
            Write-Warning ("{0} text(s) use a font other than Arial (e.g. Segoe UI) in Whiteboard; Excalidraw draws them in its Helvetica, so they are narrower and may wrap differently." -f $result.FontSubstituted)
        }
        if ($result.MirrorIgnored -gt 0) {            Write-Warning ("{0} mirrored object(s) were drawn unmirrored (Excalidraw can't mirror): {1}" -f `
                $result.MirrorIgnored, ($result.MirrorIgnoredDetails -join ', '))
        }
        if ($result.UnsupportedObjects -gt 0) {
            Write-Warning ("{0} object(s) were unsupported. Details: {1}" -f `
                $result.UnsupportedObjects, ($result.UnsupportedDetails -join ', '))
        }
        if ($PassThru) { $result }
    }
}
