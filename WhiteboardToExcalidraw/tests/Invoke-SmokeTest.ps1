#requires -Version 5.1
<#
.SYNOPSIS
Pure-PowerShell smoke test for the Whiteboard -> Excalidraw converters.

.DESCRIPTION
Converts every Microsoft Whiteboard HTML export under samples\ with both converter scripts
(solid and gradient) and checks each .excalidraw scene without any tools beyond PowerShell:

  Structure  - the file is JSON with type "excalidraw"; element ids are unique; every
               element has finite x/y/width/height; every image element's fileId is in
               "files" with a MIME type Excalidraw accepts and a dataURL that decodes;
               every text element's text/originalText is set.
  Coverage   - every source object produced output (the converter reports no unsupported
               objects); every PlainText string appears verbatim in a text element; no "?"
               fallback glyphs; every Shape is drawn with the right Excalidraw kind (oval ->
               ellipse); every reaction sticker is drawn from its original artwork (an image
               element) rather than a stand-in; every connector arrowhead is on an "arrow"
               element; dashed connectors stay dashed; rotated text is rotated; no element
               collapsed to the 1 px minimum size.

Exit code 0 when nothing FAILs (WARNs are allowed), 1 otherwise. A JSON copy of the results
is written to tests\out\smoke-results.json.

.EXAMPLE
.\tests\Invoke-SmokeTest.ps1

.EXAMPLE
.\tests\Invoke-SmokeTest.ps1 -Sample AssumptionGrid, ImageBoard -Version solid
#>
[CmdletBinding()]
param(
    # Folder holding the two Convert-WhiteboardHtmlToExcalidraw-*.ps1 scripts (default: repo root).
    [string] $ScriptDirectory,
    # Folder holding one sub-folder per sample, each with a Whiteboard .html export.
    [string] $SampleDirectory,
    # Where the .excalidraw files and smoke-results.json are written (default: tests\out).
    [string] $OutputDirectory,
    # Limit the run to these sample names (folder names); default is all.
    [string[]] $Sample,
    [ValidateSet('solid', 'gradient')]
    [string[]] $Version = @('solid', 'gradient')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
# Defaults are resolved here, not in param(): under "powershell.exe -File", Windows PowerShell
# 5.1 leaves $PSScriptRoot empty while an advanced script's parameter defaults are evaluated.
if (-not $ScriptDirectory) { $ScriptDirectory = Split-Path -Parent $PSScriptRoot }
if (-not $SampleDirectory) { $SampleDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) 'samples' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $PSScriptRoot 'out' }
$RxOpts = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
# MIME types Excalidraw's image loader accepts (IMAGE_MIME_TYPES in @excalidraw/common).
$ExcalidrawMimes = @('image/svg+xml', 'image/png', 'image/jpeg', 'image/gif', 'image/webp',
                     'image/bmp', 'image/x-icon', 'image/avif', 'image/jfif')

# ------------------------------------------------------------------------------------------
# What the export says should be there (parsed the same way the converters split anchors).
# ------------------------------------------------------------------------------------------
function Get-SourceExpectations {
    param([string] $HtmlPath)
    $html = [IO.File]::ReadAllText($HtmlPath, [Text.Encoding]::UTF8)
    $starts = [regex]::Matches($html, '(?=<div\s+class="anchor\b)', 'IgnoreCase')
    $texts = New-Object Collections.Generic.List[string]
    $types = @{}
    $counts = @{ Shape = 0; Oval = 0; Sticker = 0; Arrowhead = 0; DashedConnector = 0; Rotated = 0; RotatedText = 0 }
    for ($i = 0; $i -lt $starts.Count; $i++) {
        $start = $starts[$i].Index
        $end = if ($i + 1 -lt $starts.Count) { $starts[$i + 1].Index } else { $html.Length }
        $block = $html.Substring($start, $end - $start)
        $tag = [regex]::Match($block, '^<div\s+class="anchor\b[^>]*>', $RxOpts).Value
        $type = [regex]::Match($tag, 'data-whiteboard-type="([^"]+)"').Groups[1].Value
        if (-not $type) { continue }
        if (-not $types.ContainsKey($type)) { $types[$type] = 0 }
        $types[$type]++
        $m = [regex]::Match($tag, 'matrix\(([^)]+)\)')
        if ($m.Success) {
            $v = @($m.Groups[1].Value -split '\s*,\s*' | ForEach-Object { [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture) })
            if ([Math]::Abs($v[1]) -gt 1e-6 -or [Math]::Abs($v[2]) -gt 1e-6 -or $v[0] -lt 0 -or $v[3] -lt 0) {
                if ($type -eq 'PlainText') { $counts.RotatedText++ } else { $counts.Rotated++ }
            }
        }
        switch ($type) {
            'PlainText' {
                $spans = [regex]::Matches($block, '<span\s+data-text="true"[^>]*>(.*?)</span>', $RxOpts)
                # As a browser reads it: Whiteboard keeps soft line breaks as a raw CR or CRLF.
                $t = (@($spans | ForEach-Object { [Net.WebUtility]::HtmlDecode(($_.Groups[1].Value -replace '<[^>]+>', '')) }) -join '') -replace "`r`n?", "`n"
                if ($t.Trim()) { $texts.Add($t) }
            }
            'Shape' {
                $counts.Shape++
                $label = [regex]::Match($block, '<svg\b[^>]*\baria-label="\s*([^,."]+)', $RxOpts).Groups[1].Value
                if ($label -match 'oval|ellipse|circle') { $counts.Oval++ }
            }
            'ReactionStickers' { if ($block -match '<img\b[^>]*\bsrc="data:') { $counts.Sticker++ } }
            'Connector' {
                $counts.Arrowhead += [regex]::Matches($block, '<path\s+d="[^"]+"\s+transform="\s*translate\([^)]*\)\s*,?\s*rotate\(', $RxOpts).Count
                # Dashed = a real dash array on the connector's <g> or line; solid double-arrow
                # connectors carry stroke-dasharray="none" on their arrowhead paths.
                if ($block -match 'stroke-dasharray="(?!none")[^"]+"') { $counts.DashedConnector++ }
            }
        }
    }
    [pscustomobject]@{
        Types = $types; Texts = $texts; Shapes = $counts.Shape; Ovals = $counts.Oval; Stickers = $counts.Sticker
        Arrowheads = $counts.Arrowhead; DashedConnectors = $counts.DashedConnector
        RotatedNonText = $counts.Rotated; RotatedText = $counts.RotatedText
    }
}

# ------------------------------------------------------------------------------------------
# Structural + content checks on one .excalidraw scene
# ------------------------------------------------------------------------------------------
function Test-Scene {
    param([string] $ScenePath, $Expect, $RunResult)
    $issues = New-Object Collections.Generic.List[object]
    function Add-Issue([string]$Level, [string]$Message) { $issues.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
    $summary = [pscustomobject]@{ Issues = $issues; Elements = 0; Texts = 0; Images = 0 }

    try { $scene = [IO.File]::ReadAllText($ScenePath, [Text.Encoding]::UTF8) | ConvertFrom-Json }
    catch { Add-Issue 'FAIL' "scene is not valid JSON: $($_.Exception.Message)"; return $summary }
    if ($scene.type -ne 'excalidraw') { Add-Issue 'FAIL' "type is '$($scene.type)', expected 'excalidraw'" }
    $elements = @($scene.elements)
    $summary.Elements = $elements.Count

    # Ids, numbers
    $ids = @{}; $dupes = 0; $badNum = 0; $tiny = 0
    foreach ($e in $elements) {
        if ($ids.ContainsKey($e.id)) { $dupes++ } else { $ids[$e.id] = $e }
        foreach ($p in 'x', 'y', 'width', 'height', 'angle') {
            $n = $e.$p
            if ($null -eq $n -or $n -isnot [ValueType] -or [double]::IsNaN([double]$n) -or [double]::IsInfinity([double]$n)) { $badNum++; break }
        }
        # The converters clamp sizes to >= 1 px; anything at exactly 1 x 1 collapsed (e.g. a zero scale).
        if ($e.type -ne 'line' -and $e.type -ne 'arrow' -and [double]$e.width -le 1 -and [double]$e.height -le 1) { $tiny++ }
    }
    if ($dupes) { Add-Issue 'FAIL' "$dupes duplicate element id(s)" }
    if ($badNum) { Add-Issue 'FAIL' "$badNum element(s) with a missing or non-finite x/y/width/height/angle" }
    if ($tiny) { Add-Issue 'FAIL' "$tiny element(s) collapsed to 1 x 1 px" }

    # Images: fileId present, MIME accepted, payload decodes
    $files = @{}
    if ($scene.files) { foreach ($p in $scene.files.PSObject.Properties) { $files[$p.Name] = $p.Value } }
    $images = @($elements | Where-Object { $_.type -eq 'image' })
    $summary.Images = $images.Count
    foreach ($img in $images) {
        if (-not $files.ContainsKey($img.fileId)) { Add-Issue 'FAIL' "image $($img.id) references missing file $($img.fileId)"; continue }
        $f = $files[$img.fileId]
        if ($ExcalidrawMimes -notcontains $f.mimeType) { Add-Issue 'FAIL' "image file $($img.fileId) has MIME '$($f.mimeType)', which Excalidraw won't load" }
        $m = [regex]::Match([string]$f.dataURL, '^data:([^;,]+);base64,(.*)$', $RxOpts)
        if (-not $m.Success) { Add-Issue 'FAIL' "image file $($img.fileId) dataURL is not a base64 data URI"; continue }
        if ($m.Groups[1].Value -ne $f.mimeType) { Add-Issue 'FAIL' "image file $($img.fileId): dataURL MIME '$($m.Groups[1].Value)' differs from mimeType '$($f.mimeType)'" }
        try { $bytes = [Convert]::FromBase64String($m.Groups[2].Value) } catch { Add-Issue 'FAIL' "image file $($img.fileId) payload is not valid base64"; continue }
        if ($bytes.Length -lt 24) { Add-Issue 'FAIL' "image file $($img.fileId) is empty or truncated" }
        if ($f.mimeType -eq 'image/svg+xml') {
            try { [xml]([Text.Encoding]::UTF8.GetString($bytes)) | Out-Null } catch { Add-Issue 'FAIL' "embedded SVG $($img.fileId) does not parse" }
        }
    }

    # Text preservation
    $textEls = @($elements | Where-Object { $_.type -eq 'text' })
    $summary.Texts = $textEls.Count
    $outTexts = @($textEls | ForEach-Object { [string]$_.originalText })
    $brokenBreaks = 0
    foreach ($t in $textEls) {
        if ([string]::IsNullOrEmpty($t.text) -or [string]::IsNullOrEmpty($t.originalText)) { Add-Issue 'FAIL' "text element $($t.id) has empty text/originalText" }
        # Excalidraw breaks lines only at LF: a CR left in the text is not a line break, and the
        # wrapped text must keep every line break of the original.
        elseif ("$($t.text)$($t.originalText)".Contains("`r") -or
                ([regex]::Matches([string]$t.text, "`n").Count -lt [regex]::Matches([string]$t.originalText, "`n").Count)) { $brokenBreaks++ }
    }
    if ($brokenBreaks) { Add-Issue 'FAIL' "$brokenBreaks text element(s) lose a line break (raw CR, or fewer lines than the original)" }
    $missing = @($Expect.Texts | Where-Object { $outTexts -notcontains $_ })
    if ($missing.Count) { Add-Issue 'FAIL' ("{0} source text(s) not found verbatim: {1}" -f $missing.Count, (($missing | Select-Object -First 3) -join ' | ')) }
    $qmarks = @($outTexts | Where-Object { $_ -eq '?' }).Count
    if ($qmarks) { Add-Issue 'FAIL' "$qmarks '?' fallback glyph(s) present" }

    # Shapes: ovals must be ellipses. (Notes are rectangles too, so only ellipses are counted.)
    $ellipses = @($elements | Where-Object { $_.type -eq 'ellipse' }).Count
    if ($ellipses -lt $Expect.Ovals) { Add-Issue 'FAIL' "ovals: export has $($Expect.Ovals), only $ellipses drawn as ellipses" }

    # Stickers: each one with inline artwork should be an image element, not a stand-in.
    $stickerImages = @($images | Where-Object { $files.ContainsKey($_.fileId) -and $files[$_.fileId].mimeType -eq 'image/svg+xml' }).Count
    if ($stickerImages -lt $Expect.Stickers) {
        Add-Issue 'FAIL' "stickers: export has $($Expect.Stickers) with artwork, only $stickerImages embedded as SVG images (the rest are stand-ins or '?')"
    }

    # Connectors: arrowheads live on "arrow" elements; Excalidraw ignores them on "line".
    $heads = 0; $ignoredHeads = 0
    foreach ($e in $elements | Where-Object { $_.type -eq 'arrow' -or $_.type -eq 'line' }) {
        foreach ($p in 'startArrowhead', 'endArrowhead') {
            if ($e.PSObject.Properties[$p] -and $e.$p) { if ($e.type -eq 'arrow') { $heads++ } else { $ignoredHeads++ } }
        }
    }
    if ($heads -ne $Expect.Arrowheads) { Add-Issue 'FAIL' "arrowheads: export has $($Expect.Arrowheads), scene has $heads on arrow elements" }
    if ($ignoredHeads) { Add-Issue 'FAIL' "$ignoredHeads arrowhead(s) set on 'line' elements, which Excalidraw does not draw" }
    $dashed = @($elements | Where-Object { ($_.type -eq 'arrow' -or $_.type -eq 'line') -and $_.strokeStyle -eq 'dashed' }).Count
    if ($dashed -lt $Expect.DashedConnectors) { Add-Issue 'FAIL' "dashed connectors: export has $($Expect.DashedConnectors), scene has $dashed" }

    # Rotation: rotated text must carry an angle.
    $rotated = @($textEls | Where-Object { [Math]::Abs([double]$_.angle) -gt 1e-6 }).Count
    if ($rotated -ne $Expect.RotatedText) { Add-Issue 'FAIL' "rotated text: export has $($Expect.RotatedText), scene has $rotated" }
    $rotatedOther = @($elements | Where-Object { $_.type -ne 'text' -and [Math]::Abs([double]$_.angle) -gt 1e-6 } |
                      ForEach-Object { if ($_.groupIds -and @($_.groupIds).Count) { @($_.groupIds)[0] } else { $_.id } } | Sort-Object -Unique).Count
    if ($rotatedOther -lt $Expect.RotatedNonText) { Add-Issue 'FAIL' "rotated objects: export has $($Expect.RotatedNonText) non-text, scene rotates $rotatedOther" }

    # Converter's own report (counters are read only when this converter version has them)
    $prop = { param($name) if ($RunResult -and $RunResult.PSObject.Properties[$name]) { $RunResult.$name } else { $null } }
    if ((& $prop 'UnsupportedObjects') -gt 0) { Add-Issue 'FAIL' ("unsupported objects: " + ($RunResult.UnsupportedDetails -join ', ')) }
    $traced = & $prop 'ShapesTraced'; $fromLabel = & $prop 'ShapesFromLabel'
    if ($null -ne $traced -and $null -ne $fromLabel -and $Expect.Shapes -ne ($traced + $fromLabel)) {
        Add-Issue 'FAIL' "shapes: export has $($Expect.Shapes), converter produced $($traced + $fromLabel)"
    }
    if ($fromLabel -gt 0) { Add-Issue 'WARN' "$fromLabel shape(s) built from the label (outline not traceable)" }
    $fallback = & $prop 'StickersFallback'
    if ($fallback -gt 0) { Add-Issue 'WARN' "$fallback sticker(s) drawn with a fallback symbol" }
    $mirrored = & $prop 'MirrorIgnored'
    if ($mirrored -gt 0) { Add-Issue 'WARN' "$mirrored mirrored object(s) drawn unmirrored" }
    return $summary
}

# ------------------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory | Out-Null }
$htmlFiles = @(Get-ChildItem -LiteralPath $SampleDirectory -Directory | Sort-Object Name | ForEach-Object {
    if ($Sample -and ($Sample -notcontains $_.Name)) { return }
    Get-ChildItem -LiteralPath $_.FullName -Filter '*.html' | Select-Object -First 1
})
if (-not $htmlFiles.Count) { throw "No Whiteboard .html exports found under '$SampleDirectory'." }

$rows = New-Object Collections.Generic.List[object]
foreach ($ver in $Version) {
    $script = Join-Path $ScriptDirectory "Convert-WhiteboardHtmlToExcalidraw-$ver.ps1"
    if (-not (Test-Path -LiteralPath $script)) { throw "Converter not found: $script" }
    $outDir = Join-Path $OutputDirectory $ver
    foreach ($html in $htmlFiles) {
        $board = $html.BaseName
        $expect = Get-SourceExpectations $html.FullName
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $status = 'PASS'; $messages = @(); $check = $null; $run = $null
        try {
            $run = & $script -InputPath $html.FullName -OutputDirectory $outDir -Force -PassThru 3>$null 6>$null
            $run = @($run)[-1]
            $check = Test-Scene -ScenePath $run.OutputPath -Expect $expect -RunResult $run
            if (@($check.Issues | Where-Object Level -eq 'FAIL').Count) { $status = 'FAIL' }
            elseif (@($check.Issues | Where-Object Level -eq 'WARN').Count) { $status = 'WARN' }
            $messages = @($check.Issues | ForEach-Object { "$($_.Level): $($_.Message)" })
        } catch {
            $status = 'FAIL'; $messages = @("FAIL: conversion threw: $($_.Exception.Message)")
        }
        $sw.Stop()
        $rows.Add([pscustomobject]@{
            Board = $board; Version = $ver; Status = $status; Ms = $sw.ElapsedMilliseconds
            Objects = ($expect.Types.Values | Measure-Object -Sum).Sum
            Elements = $(if ($check) { $check.Elements } else { '' })
            Texts = $(if ($check) { '{0}/{1}' -f $expect.Texts.Count, $check.Texts } else { '' })
            Shapes = $expect.Shapes; Ovals = $expect.Ovals; Stickers = $expect.Stickers; Arrowheads = $expect.Arrowheads
            Messages = $messages
        })
    }
}

$rows | Format-Table Board, Version, Status, Ms, Objects, Elements, @{ n = 'Texts(src/out)'; e = { $_.Texts } }, Shapes, Ovals, Stickers, Arrowheads -AutoSize | Out-String -Width 200 | Write-Host
foreach ($r in $rows | Where-Object { $_.Messages }) {
    Write-Host ("{0} / {1}:" -f $r.Board, $r.Version)
    foreach ($m in $r.Messages) { Write-Host "    $m" }
}
$rows | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'smoke-results.json') -Encoding UTF8

$fails = @($rows | Where-Object Status -eq 'FAIL').Count
$warns = @($rows | Where-Object Status -eq 'WARN').Count
Write-Host ("{0} conversion(s): {1} passed, {2} warned, {3} failed." -f $rows.Count, ($rows.Count - $fails - $warns), $warns, $fails)
if ($fails) { exit 1 } else { exit 0 }
