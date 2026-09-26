#requires -Version 5.1
<#
.SYNOPSIS
Converts a Microsoft Whiteboard HTML export into an editable Excalidraw scene (v2).

.DESCRIPTION
Reads the self-contained HTML produced by Export-WhiteboardHtml and translates
the board objects exposed in that document into Excalidraw JSON. The converter
currently handles shapes, plain text, sticky notes, connectors, and embedded
images. Reaction stickers are translated to native Excalidraw vectors/text so
that they remain visible when Excalidraw rejects Microsoft embedded SVG data.
Shape-contained text is emitted as a separate editable text element.

The HTML export is a rendered representation, not Microsoft's native board
schema. Consequently, CSS-only texture/gradient effects are reduced to solid
colors and complex SVG paths are represented by their bounding box unless they
are embedded images.

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
            id = New-Id; type = $Type; x = $X; y = $Y
            width = [Math]::Max(1, $Width); height = [Math]::Max(1, $Height)
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
        return ($values -join '')
    }

    function Get-Transform {
        param([string]$Style)
        $result = @{ ScaleX = 1.0; ScaleY = 1.0; TranslateX = 0.0; TranslateY = 0.0 }
        $matrix = Get-FirstMatch $Style 'transform\s*:\s*matrix\(([^)]+)\)'
        if ($matrix) {
            $v = @($matrix -split '\s*,\s*' | ForEach-Object { Get-Number $_ })
            if ($v.Count -ge 6) {
                $result.ScaleX = [Math]::Abs($v[0]); $result.ScaleY = [Math]::Abs($v[3])
                $result.TranslateX = $v[4]; $result.TranslateY = $v[5]
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

    function Add-TextElement {
        param(
            [Collections.Generic.List[object]]$Elements, [string]$Text,
            [double]$X, [double]$Y, [double]$Width, [double]$Height,
            [double]$FontSize, [string]$Color, [string]$Align = 'left',
            [switch]$Bold
        )
        if ([string]::IsNullOrEmpty($Text)) { return }
        $X = Get-ScalarDouble $X; $Y = Get-ScalarDouble $Y
        $Width = Get-ScalarDouble $Width 1; $Height = Get-ScalarDouble $Height 1
        $FontSize = Get-ScalarDouble $FontSize 20
        $lineHeight = 1.25
        $fontFamily = 2 # Helvetica/sans-serif in Excalidraw
        $estimatedWidth = [Math]::Max(1, [Math]::Min($Width, $Text.Length * $FontSize * 0.58))
        if ($Width -le 1) { $Width = $estimatedWidth }
        if ($Height -le 1) {
            $charsPerLine = [Math]::Max(1, [Math]::Floor($Width / ($FontSize * 0.58)))
            $lines = [Math]::Max(1, [Math]::Ceiling($Text.Length / $charsPerLine))
            $Height = $lines * $FontSize * $lineHeight
        }
        $e = New-BaseElement 'text' $X $Y $Width $Height $Color 'transparent' 'solid' 1
        $e.fontSize = $FontSize; $e.fontFamily = $fontFamily
        $e.text = $Text; $e.rawText = $Text; $e.originalText = $Text
        $e.textAlign = $Align; $e.verticalAlign = 'top'; $e.containerId = $null
        $e.autoResize = $false; $e.lineHeight = $lineHeight
        [void]$Elements.Add([pscustomobject]$e)
    }

    function Add-ReactionSticker {
        param(
            [Collections.Generic.List[object]]$Elements, [string]$Label,
            [double]$X, [double]$Y, [double]$Width, [double]$Height
        )
        $X = Get-ScalarDouble $X; $Y = Get-ScalarDouble $Y
        $Width = Get-ScalarDouble $Width 40; $Height = Get-ScalarDouble $Height 40
        $groupId = New-Id
        $size = [Math]::Max(16, [Math]::Min($Width, $Height))

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

        $symbol = if ($Label -match 'star') { [char]0x2605 } elseif ($Label -match 'trophy') { [char]::ConvertFromUtf32(0x1F3C6) } else { '?' }
        $color = if ($Label -match 'star') { '#f5b700' } elseif ($Label -match 'trophy') { '#636363' } else { '#1f1f1f' }
        Add-TextElement $Elements ([string]$symbol) $X $Y $size $size ($size * 0.9) $color 'center'
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

        foreach ($block in $blocks) {
            $a = Get-AnchorInfo $block
            if (-not $a.Type) { continue }
            if (-not $sourceTypes.ContainsKey($a.Type)) { $sourceTypes[$a.Type] = 0 }
            $sourceTypes[$a.Type]++

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
                    $strokeStyle = if ($dash) { 'dashed' } else { 'solid' }
                    $x = $a.X; $y = $a.Y
                    if ($a.IsCentered) { $x -= $w / 2; $y -= $h / 2 }
                    $shapeName = Get-FirstMatch $svgTag 'aria-label="\s*([^,."]+)'
                    $kind = if ($shapeName -match 'ellipse|circle') { 'ellipse' } elseif ($shapeName -match 'diamond') { 'diamond' } else { 'rectangle' }
                    $e = New-BaseElement $kind $x $y $w $h $stroke $fill $strokeStyle 2
                    [void]$elements.Add([pscustomobject]$e)

                    # Whiteboard stores labels such as the blue-box "test" text
                    # inside the Shape block. V1 converted only the outline.
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
                        Add-TextElement $elements $shapeText $textX $textY $innerWidth $textHeight $font $textColor $textAlign
                    }
                }
                'PlainText' {
                    $text = Get-HtmlText $block
                    $textBoxStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextbox\s+plainText\b[^"]*"[^>]*style="([^"]*)"'
                    $coreStyle = Get-FirstMatch $block '<div[^>]*class="[^"]*\btextBoxCore\b[^"]*"[^>]*style="([^"]*)"'
                    $outerStyle = Get-FirstMatch $block '<div\s+style="([^"]*width:[^"]*display:\s*flex[^"]*)"'
                    $font = (Get-CssNumber $textBoxStyle 'font-size' 20) * $a.ScaleY
                    $width = Get-CssNumber $outerStyle 'width' 0
                    if ($width -le 0) { $width = Get-CssNumber $textBoxStyle 'max-width' 0 }
                    if ($width -le 0) { $width = [Math]::Max(20, $text.Length * $font * 0.58 / $a.ScaleX) }
                    $width *= $a.ScaleX
                    $height = [Math]::Max($font * 1.25, [Math]::Ceiling(($text.Length * $font * 0.58) / [Math]::Max(1, $width)) * $font * 1.25)
                    $color = Convert-RgbaToHex (Get-StyleValue $coreStyle 'color') '#000000'
                    $align = if ($block -match 'DraftEditor-alignCenter') { 'center' } elseif ($block -match 'DraftEditor-alignRight') { 'right' } else { 'left' }
                    Add-TextElement $elements $text $a.X $a.Y $width $height $font $color $align
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
                    $e = New-BaseElement 'rectangle' $a.X $a.Y $w $h 'transparent' $bg 'solid' 1
                    $e.roundness = @{ type = 3 }; [void]$elements.Add([pscustomobject]$e)
                    Add-TextElement $elements $text ($a.X + 12) ($a.Y + 12) ($w - 24) ($h - 24) (20 * $a.ScaleY) $color 'left'
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
                    $minX=[Math]::Min($x1,$x2); $minY=[Math]::Min($y1,$y2)
                    $dx=$x2-$x1; $dy=$y2-$y1
                    $stroke = Convert-RgbaToHex (Get-FirstMatch $gTag '\bstroke="([^"]+)"') '#1f1f1f'
                    $style = if ($gTag -match 'stroke-dasharray') { 'dashed' } else { 'solid' }
                    $e = New-BaseElement 'line' ($a.X+$minX) ($a.Y+$minY) ([Math]::Abs($dx)) ([Math]::Abs($dy)) $stroke 'transparent' $style 2
                    $pointList = New-Object Collections.ArrayList
                    [void]$pointList.Add([double[]]@(0,0)); [void]$pointList.Add([double[]]@($dx,$dy))
                    $e.points = $pointList; $e.lastCommittedPoint = $null
                    $e.startBinding=$null; $e.endBinding=$null; $e.startArrowhead=$null; $e.endArrowhead=$null
                    [void]$elements.Add([pscustomobject]$e)
                }
                'ReactionStickers' {
                    $label = Get-FirstMatch $block 'aria-label="([^"]+)"'
                    $size = Get-ImageSize $block
                    $w = $size.Width * $a.ScaleX; $h = $size.Height * $a.ScaleY
                    $x = $a.X; $y = $a.Y
                    if ($a.IsCentered) { $x -= $w / 2; $y -= $h / 2 }
                    Add-ReactionSticker $elements $label $x $y $w $h
                }
                {$_ -in 'Image', 'AzureImage'} {
                    # Newer Whiteboard exports tag images as "AzureImage" rather
                    # than "Image", but use the same imageComponent/img markup.
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
        if ($result.UnsupportedObjects -gt 0) {
            Write-Warning ("{0} object(s) were unsupported. Details: {1}" -f `
                $result.UnsupportedObjects, ($result.UnsupportedDetails -join ', '))
        }
        if ($PassThru) { $result }
    }
}
