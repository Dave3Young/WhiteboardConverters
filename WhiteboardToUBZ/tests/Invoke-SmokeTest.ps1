#requires -Version 5.1
<#
.SYNOPSIS
Pure-PowerShell smoke test for the Whiteboard -> OpenBoard .ubz converters.

.DESCRIPTION
Converts every Microsoft Whiteboard HTML export under samples\ with both converter scripts
(v2_solid and v3_gradient) and checks each .ubz without any tools beyond PowerShell:

  Structure  - the archive opens; metadata.rdf and page000.svg exist and parse as XML;
               ub:version is 4.8.0; every ub:uuid / ub:parent is a valid GUID; every
               <image> href exists in the archive; embedded SVG images parse; raster images
               have a readable size; all content lies inside the page's viewBox.
  Content    - every PlainText string from the export appears verbatim in a text box;
               no "?" fallback glyphs; every Shape traced from its outline; every reaction
               sticker embedded from its original artwork; every connector arrowhead
               converted; no rotated objects silently left unrotated.

Needs nothing but Windows PowerShell 5.1 (or PowerShell 7). Runs on the same machine the
converters target, so System.Drawing image downscaling is exercised for real.

Exit code 0 when nothing FAILs (WARNs are allowed), 1 otherwise. A JSON copy of the results
is written to tests\out\smoke-results.json.

.EXAMPLE
.\tests\Invoke-SmokeTest.ps1

.EXAMPLE
.\tests\Invoke-SmokeTest.ps1 -Sample AssumptionGrid, ImageBoard -Version v2_solid
#>
[CmdletBinding()]
param(
    # Folder holding the two Convert-WhiteboardHtmlToOpenBoard-*.ps1 scripts (default: repo root).
    [string] $ScriptDirectory,
    # Folder holding one sub-folder per sample, each with a Whiteboard .html export.
    [string] $SampleDirectory,
    # Where the .ubz files and smoke-results.json are written (default: tests\out).
    [string] $OutputDirectory,
    # Limit the run to these sample names (folder names); default is all.
    [string[]] $Sample,
    [ValidateSet('v2_solid', 'v3_gradient')]
    [string[]] $Version = @('v2_solid', 'v3_gradient')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
# Defaults are resolved here, not in param(): under "powershell.exe -File", Windows PowerShell
# 5.1 leaves $PSScriptRoot empty while an advanced script's parameter defaults are evaluated.
if (-not $ScriptDirectory) { $ScriptDirectory = Split-Path -Parent $PSScriptRoot }
if (-not $SampleDirectory) { $SampleDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) 'samples' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $PSScriptRoot 'out' }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$UB_NS    = 'http://uniboard.mnemis.com/document'
$SVG_NS   = 'http://www.w3.org/2000/svg'
$XLINK_NS = 'http://www.w3.org/1999/xlink'
$RxOpts   = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline

# ------------------------------------------------------------------------------------------
# What the export says should be there (parsed the same way the converters split anchors).
# ------------------------------------------------------------------------------------------
function Get-SourceExpectations {
    param([string] $HtmlPath)
    $html = [IO.File]::ReadAllText($HtmlPath, [Text.Encoding]::UTF8)
    $starts = [regex]::Matches($html, '(?=<div\s+class="anchor\b)', 'IgnoreCase')
    $texts = New-Object Collections.Generic.List[string]
    $counts = @{ Shape = 0; Oval = 0; Sticker = 0; Arrowhead = 0; Rotated = 0; RotatedText = 0; RotatedShape = 0; RotatedImage = 0 }
    # Span text as a browser reads it: Whiteboard keeps soft line breaks as a raw CR or CRLF,
    # and each Draft.js block (<div data-block="true">) is a paragraph of its own.
    $spanText = { param($b)
        $paras = @([regex]::Matches($b, '<div\b[^>]*\bdata-block="true"[^>]*>(.*?)(?=<div\b[^>]*\bdata-block="true"|$)', $RxOpts) | ForEach-Object { $_.Groups[1].Value })
        if ($paras.Count -eq 0) { $paras = @($b) }
        (@($paras | ForEach-Object {
            $spans = [regex]::Matches($_, '<span\s+data-text="true"[^>]*>(.*?)</span>', $RxOpts)
            (@($spans | ForEach-Object { [Net.WebUtility]::HtmlDecode(($_.Groups[1].Value -replace '<[^>]+>', '')) }) -join '')
        }) -join "`n") -replace "`r`n?", "`n" }
    for ($i = 0; $i -lt $starts.Count; $i++) {
        $start = $starts[$i].Index
        $end = if ($i + 1 -lt $starts.Count) { $starts[$i + 1].Index } else { $html.Length }
        $block = $html.Substring($start, $end - $start)
        $tag = [regex]::Match($block, '^<div\s+class="anchor\b[^>]*>', $RxOpts).Value
        $type = [regex]::Match($tag, 'data-whiteboard-type="([^"]+)"').Groups[1].Value
        if (-not $type) { continue }
        $m = [regex]::Match($tag, 'matrix\(([^)]+)\)')
        if ($m.Success) {
            $v = @($m.Groups[1].Value -split '\s*,\s*' | ForEach-Object { [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture) })
            if ([Math]::Abs($v[1]) -gt 1e-6 -or [Math]::Abs($v[2]) -gt 1e-6 -or $v[0] -lt 0 -or $v[3] -lt 0) {
                # PlainText and shapes are drawn rotated (a shape's label turns with it).
                if ($type -eq 'PlainText') { $counts.RotatedText++ }
                elseif ($type -eq 'Shape') { $counts.RotatedShape++; if ((& $spanText $block).Trim()) { $counts.RotatedText++ } }
                # Images and stickers are drawn rotated too, and connectors and ink are mapped
                # through the full matrix; only notes are left unrotated.
                elseif ($type -in 'Image', 'AzureImage', 'FluidImage', 'ReactionStickers') { $counts.RotatedImage++ }
                elseif ($type -notin 'Connector', 'InkGroup') { $counts.Rotated++ }
            }
        }
        switch ($type) {
            'PlainText' {
                $t = & $spanText $block
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
            }
        }
    }
    [pscustomobject]@{ Texts = $texts; Shapes = $counts.Shape; Ovals = $counts.Oval; Stickers = $counts.Sticker
                       Arrowheads = $counts.Arrowhead; RotatedNonText = $counts.Rotated; RotatedText = $counts.RotatedText
                       RotatedShapes = $counts.RotatedShape; RotatedImages = $counts.RotatedImage }
}

# ------------------------------------------------------------------------------------------
# Structural + content checks on one .ubz
# ------------------------------------------------------------------------------------------
function Test-Ubz {
    param([string] $UbzPath, $Expect, $RunResult)
    $issues = New-Object Collections.Generic.List[object]
    function Add-Issue([string]$Level, [string]$Message) { $issues.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }

    $zip = [IO.Compression.ZipFile]::OpenRead($UbzPath)
    try {
        $entries = @{}
        foreach ($e in $zip.Entries) { $entries[$e.FullName] = $e }
        foreach ($required in 'metadata.rdf', 'page000.svg') {
            if (-not $entries.ContainsKey($required)) { Add-Issue 'FAIL' "missing $required" }
        }
        $empty = { [pscustomobject]@{ Issues = $issues; Texts = 0; SvgImages = 0; RasterImages = 0 } }
        if (-not $entries.ContainsKey('page000.svg')) { return (& $empty) }

        function Read-Entry($entry) {
            $s = $entry.Open(); try { $r = New-Object IO.StreamReader($s, [Text.Encoding]::UTF8); $r.ReadToEnd() } finally { $s.Dispose() }
        }
        function Read-EntryBytes($entry) {
            $s = $entry.Open(); try { $ms = New-Object IO.MemoryStream; $s.CopyTo($ms); , $ms.ToArray() } finally { $s.Dispose() }
        }

        try { [xml](Read-Entry $entries['metadata.rdf']) | Out-Null } catch { Add-Issue 'FAIL' "metadata.rdf is not valid XML: $($_.Exception.Message)" }
        try { $svg = [xml](Read-Entry $entries['page000.svg']) } catch { Add-Issue 'FAIL' "page000.svg is not valid XML: $($_.Exception.Message)"; return (& $empty) }

        $root = $svg.DocumentElement
        if ($root.GetAttribute('version', $UB_NS) -ne '4.8.0') { Add-Issue 'FAIL' "ub:version is '$($root.GetAttribute('version', $UB_NS))', expected 4.8.0" }
        $vb = @($root.GetAttribute('viewBox') -split '\s+' | Where-Object { $_ } | ForEach-Object { [double]$_ })
        if ($vb.Count -ne 4) { Add-Issue 'FAIL' 'viewBox does not have 4 numbers'; return (& $empty) }
        $vbMinX = $vb[0]; $vbMinY = $vb[1]; $vbMaxX = $vb[0] + $vb[2]; $vbMaxY = $vb[1] + $vb[3]

        $ns = New-Object Xml.XmlNamespaceManager($svg.NameTable)
        $ns.AddNamespace('s', $SVG_NS); $ns.AddNamespace('ub', $UB_NS); $ns.AddNamespace('xlink', $XLINK_NS)

        # GUIDs
        $badGuid = 0
        foreach ($attr in $svg.SelectNodes('//@ub:uuid | //@ub:parent', $ns)) {
            $g = [guid]::Empty
            if (-not [guid]::TryParse($attr.Value, [ref]$g)) { $badGuid++ }
        }
        if ($badGuid) { Add-Issue 'FAIL' "$badGuid ub:uuid/ub:parent value(s) are not valid GUIDs" }

        # Out-of-page content (tolerance for stroke width / rounding)
        $tol = 5.0; $outside = 0
        function Test-Pt([double]$x, [double]$y) {
            ($x -lt ($vbMinX - $tol) -or $x -gt ($vbMaxX + $tol) -or $y -lt ($vbMinY - $tol) -or $y -gt ($vbMaxY + $tol))
        }
        foreach ($n in $svg.SelectNodes('//s:polygon | //s:polyline', $ns)) {
            foreach ($pt in ($n.GetAttribute('points') -split '\s+' | Where-Object { $_ })) {
                $xy = $pt -split ','
                if (Test-Pt ([double]$xy[0]) ([double]$xy[1])) { $outside++; break }
            }
        }
        foreach ($n in $svg.SelectNodes('//s:line', $ns)) {
            if ((Test-Pt ([double]$n.GetAttribute('x1')) ([double]$n.GetAttribute('y1'))) -or
                (Test-Pt ([double]$n.GetAttribute('x2')) ([double]$n.GetAttribute('y2')))) { $outside++ }
        }
        foreach ($n in $svg.SelectNodes('//s:foreignObject | //s:image', $ns)) {
            $t = [regex]::Match($n.GetAttribute('transform'), 'matrix\(([^)]+)\)')
            if ($t.Success) {
                $m = @($t.Groups[1].Value -split '\s*,\s*' | ForEach-Object { [double]$_ })
                if (Test-Pt $m[4] $m[5]) { $outside++ }
            }
        }
        if ($outside) { Add-Issue 'FAIL' "$outside element(s) lie outside the page viewBox" }

        # Images: present, readable
        $svgImages = 0; $rasterImages = 0
        foreach ($img in $svg.SelectNodes('//s:image', $ns)) {
            $href = $img.GetAttribute('href', $XLINK_NS)
            if (-not $entries.ContainsKey($href)) { Add-Issue 'FAIL' "image '$href' is referenced but not in the archive"; continue }
            $bytes = Read-EntryBytes $entries[$href]
            if ($href -like '*.svg') {
                $svgImages++
                try { [xml]([Text.Encoding]::UTF8.GetString($bytes)) | Out-Null } catch { Add-Issue 'FAIL' "embedded SVG '$href' does not parse" }
            } else {
                $rasterImages++
                if ($bytes.Length -lt 24) { Add-Issue 'FAIL' "raster image '$href' is empty or truncated" }
            }
        }

        # Text preservation
        $outTexts = New-Object Collections.Generic.List[string]
        foreach ($fo in $svg.SelectNodes('//s:foreignObject', $ns)) {
            $inner = $fo.InnerText   # itemTextContent: the (already XML-decoded) Qt rich-text HTML
            # Qt collapses raw newlines to spaces, so only <br> counts as a line break; any raw
            # CR/LF left in the text reads as a space. Line breaks must survive verbatim.
            $inner = ($inner -replace "[`r`n]+", ' ') -replace '<br\s*/?>', "`n"
            $outTexts.Add([Net.WebUtility]::HtmlDecode(($inner -replace '<[^>]+>', '')))
        }
        $missing = @($Expect.Texts | Where-Object { $outTexts -notcontains $_ })
        if ($missing.Count) { Add-Issue 'FAIL' ("{0} source text(s) not found verbatim: {1}" -f $missing.Count, (($missing | Select-Object -First 3) -join ' | ')) }
        $qmarks = @($outTexts | Where-Object { $_ -eq '?' }).Count
        if ($qmarks) { Add-Issue 'FAIL' "$qmarks '?' fallback glyph(s) present" }

        # ---- Content checks derived from the .ubz itself (work for any converter version) ----
        # Stickers: each one with inline artwork should be an embedded .svg image.
        if ($svgImages -lt $Expect.Stickers) { Add-Issue 'FAIL' "stickers: export has $($Expect.Stickers) with artwork, only $svgImages embedded as SVG" }

        # Ovals: a traced oval is a 40+ point outline; a rectangle stand-in has 4 points.
        $roundGroups = @{}
        foreach ($n in $svg.SelectNodes('//s:polygon[@ub:parent] | //s:polyline[@ub:parent]', $ns)) {
            if (@($n.GetAttribute('points') -split '\s+' | Where-Object { $_ }).Count -ge 32) { $roundGroups[$n.GetAttribute('parent', $UB_NS)] = $true }
        }
        if ($roundGroups.Count -lt $Expect.Ovals) { Add-Issue 'FAIL' "ovals: export has $($Expect.Ovals), only $($roundGroups.Count) drawn round" }

        # Arrowheads: polylines sharing a ub:parent with a connector -- a <line>, or (for an
        # elbow or curved connector) the group's first polyline, when the group has no fill
        # and only open polylines (shape outlines are closed).
        $lineGroups = @{}
        foreach ($l in $svg.SelectNodes('//s:line[@ub:parent]', $ns)) { $lineGroups[$l.GetAttribute('parent', $UB_NS)] = $true }
        $polyGroups = [ordered]@{}; $filled = @{}
        foreach ($n in $svg.SelectNodes('//s:polygon[@ub:parent] | //s:polyline[@ub:parent]', $ns)) {
            $par = $n.GetAttribute('parent', $UB_NS)
            if ($n.LocalName -eq 'polygon') { $filled[$par] = $true; continue }
            if ($n.GetAttribute('stroke-linejoin') -eq 'round') { continue }   # ink
            if (-not $polyGroups.Contains($par)) { $polyGroups[$par] = New-Object Collections.Generic.List[object] }
            $polyGroups[$par].Add($n)
        }
        $arrows = 0
        foreach ($par in $polyGroups.Keys) {
            $pls = $polyGroups[$par]
            if ($lineGroups.ContainsKey($par)) { $arrows += $pls.Count; continue }
            if ($filled.ContainsKey($par)) { continue }
            $open = @($pls | Where-Object { $p = @($_.GetAttribute('points') -split '\s+' | Where-Object { $_ }); $p[0] -ne $p[$p.Count - 1] }).Count
            if ($open -eq $pls.Count) { $arrows += $pls.Count - 1 }
        }
        if ($arrows -ne $Expect.Arrowheads) { Add-Issue 'FAIL' "arrowheads: export has $($Expect.Arrowheads), .ubz has $arrows" }

        # Rotated text: a text box whose matrix has off-diagonal terms.
        $rotated = 0
        foreach ($fo in $svg.SelectNodes('//s:foreignObject', $ns)) {
            $t = [regex]::Match($fo.GetAttribute('transform'), 'matrix\(([^)]+)\)')
            if ($t.Success) {
                $m = @($t.Groups[1].Value -split '\s*,\s*' | ForEach-Object { [double]$_ })
                if ([Math]::Abs($m[1]) -gt 1e-6 -or [Math]::Abs($m[2]) -gt 1e-6) { $rotated++ }
            }
        }
        if ($rotated -ne $Expect.RotatedText) { Add-Issue 'FAIL' "rotated text: export has $($Expect.RotatedText), .ubz has $rotated" }
        # Rotated images and stickers: an image whose matrix has off-diagonal terms.
        $rotatedImages = 0
        foreach ($img in $svg.SelectNodes('//s:image', $ns)) {
            $t = [regex]::Match($img.GetAttribute('transform'), 'matrix\(([^)]+)\)')
            if ($t.Success) {
                $m = @($t.Groups[1].Value -split '\s*,\s*' | ForEach-Object { [double]$_ })
                if ([Math]::Abs($m[1]) -gt 1e-6 -or [Math]::Abs($m[2]) -gt 1e-6 -or $m[0] -lt 0 -or $m[3] -lt 0) { $rotatedImages++ }
            }
        }
        if ($rotatedImages -ne $Expect.RotatedImages) { Add-Issue 'FAIL' "rotated images/stickers: export has $($Expect.RotatedImages), .ubz has $rotatedImages" }
        if ($Expect.RotatedNonText -gt 0) { Add-Issue 'WARN' "$($Expect.RotatedNonText) rotated note(s) in the export (drawn unrotated)" }

        # ---- Converter's own counters, when this version reports them ----
        $prop = { param($name) if ($RunResult -and $RunResult.PSObject.Properties[$name]) { $RunResult.$name } else { $null } }
        $traced = & $prop 'ShapesTraced'; $fromLabel = & $prop 'ShapesFromLabel'
        if ($null -ne $traced -and $null -ne $fromLabel -and $Expect.Shapes -ne ($traced + $fromLabel)) {
            Add-Issue 'FAIL' "shapes: export has $($Expect.Shapes), converter produced $($traced + $fromLabel)"
        }
        if ($fromLabel -gt 0) { Add-Issue 'WARN' "$fromLabel shape(s) built from the label (outline not traceable)" }
        $rotShapes = & $prop 'RotatedShapesConverted'
        if ($null -ne $rotShapes -and $rotShapes -ne $Expect.RotatedShapes) {
            Add-Issue 'FAIL' "rotated shapes: export has $($Expect.RotatedShapes), converter rotated $rotShapes"
        } elseif ($null -eq $rotShapes -and $Expect.RotatedShapes -gt 0) {
            Add-Issue 'WARN' "$($Expect.RotatedShapes) rotated shape(s) in the export (this converter version doesn't report rotating them)"
        }
        $fallback = & $prop 'StickersFallback'
        if ($fallback -gt 0) { Add-Issue 'WARN' "$fallback sticker(s) drawn with a fallback symbol" }
        $unsupported = & $prop 'UnsupportedObjects'
        if ($unsupported -gt 0) { Add-Issue 'WARN' ("unsupported objects: " + ($RunResult.UnsupportedDetails -join ', ')) }

        [pscustomobject]@{
            Issues = $issues; Texts = $outTexts.Count; SvgImages = $svgImages; RasterImages = $rasterImages
        }
    } finally { $zip.Dispose() }
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
    $script = Join-Path $ScriptDirectory "Convert-WhiteboardHtmlToOpenBoard-$ver.ps1"
    if (-not (Test-Path -LiteralPath $script)) { throw "Converter not found: $script" }
    $outDir = Join-Path $OutputDirectory $ver
    foreach ($html in $htmlFiles) {
        $board = $html.BaseName
        $expect = Get-SourceExpectations $html.FullName
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $status = 'PASS'; $messages = @(); $graphics = $null; $check = $null
        try {
            $run = & $script -InputPath $html.FullName -OutputDirectory $outDir -Force -PassThru 3>$null 6>$null
            $run = @($run)[-1]
            $graphics = $run.OpenBoardGraphics
            $check = Test-Ubz -UbzPath $run.OutputPath -Expect $expect -RunResult $run
            if (@($check.Issues | Where-Object Level -eq 'FAIL').Count) { $status = 'FAIL' }
            elseif (@($check.Issues | Where-Object Level -eq 'WARN').Count) { $status = 'WARN' }
            $messages = @($check.Issues | ForEach-Object { "$($_.Level): $($_.Message)" })
        } catch {
            $status = 'FAIL'; $messages = @("FAIL: conversion threw: $($_.Exception.Message)")
        }
        $sw.Stop()
        $rows.Add([pscustomobject]@{
            Board = $board; Version = $ver; Status = $status; Ms = $sw.ElapsedMilliseconds
            Graphics = $graphics
            Texts = $(if ($check) { '{0}/{1}' -f $expect.Texts.Count, $check.Texts } else { '' })
            Stickers = $expect.Stickers; Shapes = $expect.Shapes; Ovals = $expect.Ovals; Arrowheads = $expect.Arrowheads
            Messages = $messages
        })
    }
}

$rows | Format-Table Board, Version, Status, Ms, Graphics, @{ n = 'Texts(src/out)'; e = { $_.Texts } }, Shapes, Ovals, Stickers, Arrowheads -AutoSize | Out-String -Width 200 | Write-Host
foreach ($r in $rows | Where-Object { $_.Messages }) {
    Write-Host ("{0} / {1}:" -f $r.Board, $r.Version)
    foreach ($m in $r.Messages) { Write-Host "    $m" }
}
$rows | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'smoke-results.json') -Encoding UTF8

$fails = @($rows | Where-Object Status -eq 'FAIL').Count
$warns = @($rows | Where-Object Status -eq 'WARN').Count
Write-Host ("{0} conversion(s): {1} passed, {2} warned, {3} failed." -f $rows.Count, ($rows.Count - $fails - $warns), $warns, $fails)
if ($fails) { exit 1 } else { exit 0 }
