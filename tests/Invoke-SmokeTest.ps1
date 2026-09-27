#requires -Version 5.1
<#
.SYNOPSIS
Pure-PowerShell smoke test for the Whiteboard -> IWB converter.

.DESCRIPTION
Converts every Microsoft Whiteboard HTML export under samples\ with Convert-WhiteboardHtmlToIwb.ps1
(once per -NoteFill variant) and checks each .iwb without any tools beyond PowerShell:

  Structure  - the archive opens and has content.xml at its root; content.xml parses; the root
               is <iwb version="1.0"> in the IMS CFF namespace, with meta / svg / group
               children in that order; <svg:svg> has width, height and a lower-case viewbox,
               and holds a pageset with one page; ids are unique; every group reference
               resolves; every image file exists and is a CFF image format (PNG, JPEG, GIF,
               BMP); all content lies inside the page.
  OpenBoard  - only the element forms OpenBoard's CFF importer reads correctly: rect, polygon,
               2-point polyline, textarea (with tbreak), image; every shape has an explicit
               stroke colour; polygons have a real fill; font sizes are in pt and weights are
               normal/bold; numbers have no exponents; a text transform is exactly
               "translate(x,y) rotate(a)".
  Content    - every PlainText string from the export appears verbatim in a textarea (line
               breaks as <svg:tbreak/>); no "?" fallback glyphs; every Shape traced from its
               outline; ovals drawn round; every sticker embedded as an image; every connector
               arrowhead converted; rotated text and shapes really rotated.

Exit code 0 when nothing FAILs (WARNs are allowed), 1 otherwise. A JSON copy of the results
is written to tests\out\smoke-results.json.

.EXAMPLE
.\tests\Invoke-SmokeTest.ps1

.EXAMPLE
.\tests\Invoke-SmokeTest.ps1 -Sample AssumptionGrid, ImageBoard -NoteFill Solid
#>
[CmdletBinding()]
param(
    # Folder holding Convert-WhiteboardHtmlToIwb.ps1 (default: repo root).
    [string] $ScriptDirectory,
    # Folder holding one sub-folder per sample, each with a Whiteboard .html export.
    [string] $SampleDirectory,
    # Where the .iwb files and smoke-results.json are written (default: tests\out).
    [string] $OutputDirectory,
    # Limit the run to these sample names (folder names); default is all.
    [string[]] $Sample,
    [ValidateSet('Solid', 'Gradient')]
    [string[]] $NoteFill = @('Solid', 'Gradient')
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

$IWB_NS   = 'http://www.imsglobal.org/xsd/iwb_v1p0'
$SVG_NS   = 'http://www.w3.org/2000/svg'
$XLINK_NS = 'http://www.w3.org/1999/xlink'
$RxOpts   = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
$ci       = [Globalization.CultureInfo]::InvariantCulture
$PageElements = @('rect', 'polygon', 'polyline', 'textarea', 'image')

# ------------------------------------------------------------------------------------------
# What the export says should be there (parsed the same way the converter splits anchors).
# ------------------------------------------------------------------------------------------
function Get-SourceExpectations {
    param([string] $HtmlPath)
    $html = [IO.File]::ReadAllText($HtmlPath, [Text.Encoding]::UTF8)
    $starts = [regex]::Matches($html, '(?=<div\s+class="anchor\b)', 'IgnoreCase')
    $texts = New-Object Collections.Generic.List[string]
    $counts = @{ Shape = 0; Oval = 0; Sticker = 0; Arrowhead = 0; Connector = 0; Rotated = 0; RotatedText = 0; RotatedShape = 0 }
    # Span text as a browser reads it: Whiteboard keeps soft line breaks as a raw CR or CRLF.
    $spanText = { param($b)
        $spans = [regex]::Matches($b, '<span\s+data-text="true"[^>]*>(.*?)</span>', $RxOpts)
        ((@($spans | ForEach-Object { [Net.WebUtility]::HtmlDecode(($_.Groups[1].Value -replace '<[^>]+>', '')) }) -join '') -replace "`r`n?", "`n") }
    for ($i = 0; $i -lt $starts.Count; $i++) {
        $start = $starts[$i].Index
        $end = if ($i + 1 -lt $starts.Count) { $starts[$i + 1].Index } else { $html.Length }
        $block = $html.Substring($start, $end - $start)
        $tag = [regex]::Match($block, '^<div\s+class="anchor\b[^>]*>', $RxOpts).Value
        $type = [regex]::Match($tag, 'data-whiteboard-type="([^"]+)"').Groups[1].Value
        if (-not $type) { continue }
        $m = [regex]::Match($tag, 'matrix\(([^)]+)\)')
        if ($m.Success) {
            $v = @($m.Groups[1].Value -split '\s*,\s*' | ForEach-Object { [double]::Parse($_, $ci) })
            if ([Math]::Abs($v[1]) -gt 1e-6 -or [Math]::Abs($v[2]) -gt 1e-6 -or $v[0] -lt 0 -or $v[3] -lt 0) {
                # PlainText and shapes are drawn rotated (a shape's label turns with it).
                if ($type -eq 'PlainText') { $counts.RotatedText++ }
                elseif ($type -eq 'Shape') { $counts.RotatedShape++; if ((& $spanText $block).Trim()) { $counts.RotatedText++ } }
                else { $counts.Rotated++ }
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
                $counts.Connector++
                $counts.Arrowhead += [regex]::Matches($block, '<path\s+d="[^"]+"\s+transform="\s*translate\([^)]*\)\s*,?\s*rotate\(', $RxOpts).Count
            }
        }
    }
    [pscustomobject]@{ Texts = $texts; Shapes = $counts.Shape; Ovals = $counts.Oval; Stickers = $counts.Sticker
                       Connectors = $counts.Connector; Arrowheads = $counts.Arrowhead; RotatedNonText = $counts.Rotated
                       RotatedText = $counts.RotatedText; RotatedShapes = $counts.RotatedShape }
}

# ------------------------------------------------------------------------------------------
# Structural + content checks on one .iwb
# ------------------------------------------------------------------------------------------
function Get-ImageFormat {
    param([byte[]] $B)
    if ($B.Length -ge 8 -and $B[0] -eq 0x89 -and $B[1] -eq 0x50 -and $B[2] -eq 0x4e -and $B[3] -eq 0x47) { return 'png' }
    if ($B.Length -ge 3 -and $B[0] -eq 0xff -and $B[1] -eq 0xd8 -and $B[2] -eq 0xff) { return 'jpg' }
    if ($B.Length -ge 6 -and $B[0] -eq 0x47 -and $B[1] -eq 0x49 -and $B[2] -eq 0x46) { return 'gif' }
    if ($B.Length -ge 2 -and $B[0] -eq 0x42 -and $B[1] -eq 0x4d) { return 'bmp' }
    $head = [Text.Encoding]::UTF8.GetString($B, 0, [Math]::Min($B.Length, 1024))
    if ($head -match '<svg\b') { return 'svg' }
    return 'unknown'
}

function Test-Iwb {
    param([string] $IwbPath, $Expect, $RunResult)
    $issues = New-Object Collections.Generic.List[object]
    function Add-Issue([string]$Level, [string]$Message) { $issues.Add([pscustomobject]@{ Level = $Level; Message = $Message }) }
    $prop = { param($name) if ($RunResult -and $RunResult.PSObject.Properties[$name]) { $RunResult.$name } else { $null } }

    $zip = [IO.Compression.ZipFile]::OpenRead($IwbPath)
    try {
        $entries = @{}
        foreach ($e in $zip.Entries) {
            $entries[$e.FullName] = $e
            if ($e.FullName -match '\\') { Add-Issue 'FAIL' "archive entry '$($e.FullName)' uses a backslash" }
        }
        $empty = [pscustomobject]@{ Issues = $issues; Texts = 0; Images = 0 }
        if (-not $entries.ContainsKey('content.xml')) { Add-Issue 'FAIL' 'missing content.xml at the archive root'; return $empty }

        function Read-EntryBytes($entry) {
            $s = $entry.Open(); try { $ms = New-Object IO.MemoryStream; $s.CopyTo($ms); , $ms.ToArray() } finally { $s.Dispose() }
        }
        $raw = [Text.Encoding]::UTF8.GetString((Read-EntryBytes $entries['content.xml']))
        try { $doc = [xml]$raw } catch { Add-Issue 'FAIL' "content.xml is not valid XML: $($_.Exception.Message)"; return $empty }

        $root = $doc.DocumentElement
        if ($root.LocalName -ne 'iwb' -or $root.NamespaceURI -ne $IWB_NS) { Add-Issue 'FAIL' "root is {$($root.NamespaceURI)}$($root.LocalName), expected {$IWB_NS}iwb" }
        if ($root.GetAttribute('version') -ne '1.0') { Add-Issue 'FAIL' "iwb version is '$($root.GetAttribute('version'))', expected 1.0" }
        # Top-level order: meta*, svg, then group/element (a group can only refer to parsed items).
        $order = @($root.ChildNodes | Where-Object { $_.NodeType -eq 'Element' } | ForEach-Object { $_.LocalName })
        $svgIndex = [array]::IndexOf($order, 'svg')
        if ($svgIndex -lt 0 -or @($order | Where-Object { $_ -eq 'svg' }).Count -ne 1) { Add-Issue 'FAIL' 'content.xml needs exactly one svg:svg'; return $empty }
        if (@($order[0..$svgIndex] | Where-Object { $_ -notin 'meta', 'svg' }).Count) { Add-Issue 'FAIL' 'only iwb:meta may come before svg:svg' }
        if ($svgIndex -lt $order.Count - 1 -and @($order[($svgIndex + 1)..($order.Count - 1)] | Where-Object { $_ -notin 'group', 'element' }).Count) {
            Add-Issue 'FAIL' 'only iwb:group / iwb:element may follow svg:svg'
        }

        $ns = New-Object Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('i', $IWB_NS); $ns.AddNamespace('s', $SVG_NS); $ns.AddNamespace('xlink', $XLINK_NS)
        $svg = $doc.SelectSingleNode('/i:iwb/s:svg', $ns)
        if (-not $svg) { Add-Issue 'FAIL' 'svg:svg is not in the SVG namespace'; return $empty }
        $vb = @($svg.GetAttribute('viewbox') -split '\s+' | Where-Object { $_ } | ForEach-Object { [double]::Parse($_, $ci) })
        if ($vb.Count -ne 4) { Add-Issue 'FAIL' 'svg:svg has no 4-number lower-case viewbox (the attribute OpenBoard reads)'; return $empty }
        $pageW = $vb[2]; $pageH = $vb[3]
        if ([double]$svg.GetAttribute('width') -ne $pageW -or [double]$svg.GetAttribute('height') -ne $pageH) { Add-Issue 'FAIL' 'svg:svg width/height differ from its viewbox' }
        if ($pageH -ne 960 -or $pageW -gt 1280) { Add-Issue 'WARN' "page is ${pageW}x${pageH}: OpenBoard imports only a 960-high, <= 1280-wide page at 1:1" }
        $pages = $svg.SelectNodes('s:pageset/s:page', $ns)
        if ($pages.Count -ne 1) { Add-Issue 'FAIL' "expected one svg:page inside svg:pageset, found $($pages.Count)"; return $empty }
        $page = $pages[0]

        # ---- Page elements: the OpenBoard-safe subset, ids, numbers ----
        $ids = @{}; $numberRx = '^-?\d+(\.\d+)?$'
        $bad = @{ Numbers = 0 }; $outside = 0; $tol = 2.0
        $inPage = { param([double]$x, [double]$y) ($x -ge -$tol -and $x -le $pageW + $tol -and $y -ge -$tol -and $y -le $pageH + $tol) }
        $num = { param($el, $name) $v = $el.GetAttribute($name); if ($v -notmatch $numberRx) { $bad.Numbers++ }; [double]::Parse($(if ($v) { $v } else { '0' }), [Globalization.NumberStyles]::Float, $ci) }
        $elements = @($page.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })
        $outTexts = New-Object Collections.Generic.List[string]
        $rotatedText = 0; $roundCount = 0; $images = 0; $polylines = 0
        $edgeGroups = @{}
        foreach ($el in $elements) {
            $tag = $el.LocalName
            if ($el.NamespaceURI -ne $SVG_NS -or $PageElements -notcontains $tag) { Add-Issue 'FAIL' "unsupported page element <$($el.Name)>"; continue }
            $id = $el.GetAttribute('id')
            if (-not $id) { Add-Issue 'FAIL' "<$tag> without an id" } elseif ($ids.ContainsKey($id)) { Add-Issue 'FAIL' "duplicate id '$id'" } else { $ids[$id] = $el }
            if ($tag -in 'rect', 'polygon', 'polyline') {
                if ($el.GetAttribute('stroke') -notmatch '^#[0-9a-f]{6}$') { Add-Issue 'FAIL' "<$tag id=$id> has no explicit stroke colour (OpenBoard draws it black)" }
                $sw = $el.GetAttribute('stroke-width'); if ($sw -notmatch $numberRx) { $bad.Numbers++ }
                $dash = $el.GetAttribute('stroke-dasharray')
                if ($dash -and $dash -notmatch '^\d+(\.\d+)?(,\d+(\.\d+)?)+$') { Add-Issue 'FAIL' "<$tag id=$id> has a malformed stroke-dasharray '$dash'" }
            }
            switch ($tag) {
                'rect' {
                    $x = & $num $el 'x'; $y = & $num $el 'y'; $w = & $num $el 'width'; $h = & $num $el 'height'
                    if (-not ((& $inPage $x $y) -and (& $inPage ($x + $w) ($y + $h)))) { $outside++ }
                    if ($el.GetAttribute('fill') -notmatch '^(none|#[0-9a-f]{6})$') { Add-Issue 'FAIL' "<rect id=$id> fill '$($el.GetAttribute('fill'))'" }
                }
                { $_ -in 'polygon', 'polyline' } {
                    $pts = @($el.GetAttribute('points') -split '\s+' | Where-Object { $_ })
                    foreach ($p in $pts) {
                        $xy = $p -split ','
                        if ($xy.Count -ne 2 -or $xy[0] -notmatch $numberRx -or $xy[1] -notmatch $numberRx) { $bad.Numbers++; continue }
                        if (-not (& $inPage ([double]::Parse($xy[0], $ci)) ([double]::Parse($xy[1], $ci)))) { $outside++; break }
                    }
                    if ($tag -eq 'polygon') {
                        if ($el.GetAttribute('fill') -notmatch '^#[0-9a-f]{6}$') { Add-Issue 'FAIL' "<polygon id=$id> has no real fill (OpenBoard fills it black)" }
                        if ($pts.Count -lt 3) { Add-Issue 'FAIL' "<polygon id=$id> has $($pts.Count) points" }
                        if ($pts.Count -ge 32) { $roundCount++ }
                    } else {
                        $polylines++
                        # OpenBoard fills a polyline as a closed shape: only 2 points draw a line.
                        if ($pts.Count -ne 2) { Add-Issue 'FAIL' "<polyline id=$id> has $($pts.Count) points (OpenBoard fills it)" }
                    }
                }
                'textarea' {
                    $x = & $num $el 'x'; $y = & $num $el 'y'; [void](& $num $el 'width'); [void](& $num $el 'height')
                    if ($el.GetAttribute('font-size') -notmatch '^\d+(\.\d+)?pt$') { Add-Issue 'FAIL' "<textarea id=$id> font-size '$($el.GetAttribute('font-size'))' is not in pt" }
                    if ($el.GetAttribute('font-weight') -notin 'normal', 'bold') { Add-Issue 'FAIL' "<textarea id=$id> font-weight '$($el.GetAttribute('font-weight'))' (OpenBoard maps only names)" }
                    if ($el.GetAttribute('text-align') -notin 'start', 'center', 'end') { Add-Issue 'FAIL' "<textarea id=$id> text-align '$($el.GetAttribute('text-align'))'" }
                    $tr = $el.GetAttribute('transform')
                    if ($tr) {
                        $tm = [regex]::Match($tr, '^translate\((-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)\) rotate\((-?\d+(?:\.\d+)?)\)$')
                        if (-not $tm.Success -or $x -ne 0 -or $y -ne 0) { Add-Issue 'FAIL' "<textarea id=$id> transform '$tr' is not 'translate(x,y) rotate(a)' at 0,0" }
                        else {
                            $rotatedText++
                            if (-not (& $inPage ([double]::Parse($tm.Groups[1].Value, $ci)) ([double]::Parse($tm.Groups[2].Value, $ci)))) { $outside++ }
                        }
                    } elseif (-not (& $inPage $x $y)) { $outside++ }
                    $text = New-Object Text.StringBuilder
                    foreach ($c in $el.ChildNodes) {
                        if ($c.NodeType -eq 'Text' -or $c.NodeType -eq 'SignificantWhitespace' -or $c.NodeType -eq 'Whitespace') { [void]$text.Append($c.Value) }
                        elseif ($c.NodeType -eq 'Element' -and $c.LocalName -eq 'tbreak' -and $c.NamespaceURI -eq $SVG_NS) { [void]$text.Append("`n") }
                        else { Add-Issue 'FAIL' "<textarea id=$id> contains <$($c.Name)>" }
                    }
                    $outTexts.Add($text.ToString())
                }
                'image' {
                    $images++
                    $x = & $num $el 'x'; $y = & $num $el 'y'; $w = & $num $el 'width'; $h = & $num $el 'height'
                    if (-not ((& $inPage $x $y) -and (& $inPage ($x + $w) ($y + $h)))) { $outside++ }
                    $href = $el.GetAttribute('href', $XLINK_NS)
                    if (-not $entries.ContainsKey($href)) { Add-Issue 'FAIL' "image '$href' is referenced but not in the archive"; continue }
                    $fmt = Get-ImageFormat (Read-EntryBytes $entries[$href])
                    if ($fmt -notin 'png', 'jpg', 'gif', 'bmp') {
                        $level = if ($fmt -eq 'svg' -and (& $prop 'ImagesKeptAsSvg') -gt 0) { 'WARN' } else { 'FAIL' }
                        Add-Issue $level "image '$href' is $fmt, not a CFF image format"
                    } elseif ($href -notmatch "\.$fmt$") { Add-Issue 'FAIL' "image '$href' holds $fmt data" }
                }
            }
        }
        if ($bad.Numbers) { Add-Issue 'FAIL' "$($bad.Numbers) number(s) are not plain decimals (OpenBoard can't read exponents)" }
        if ($outside) { Add-Issue 'FAIL' "$outside element(s) lie outside the page" }

        # ---- Groups ----
        $grouped = @{}
        foreach ($g in $doc.SelectNodes('/i:iwb/i:group', $ns)) {
            $refs = @($g.SelectNodes('i:element', $ns) | ForEach-Object { $_.GetAttribute('ref') })
            if ($refs.Count -lt 2) { Add-Issue 'FAIL' 'a group with fewer than 2 members' }
            $tags = @()
            foreach ($r in $refs) {
                if (-not $ids.ContainsKey($r)) { Add-Issue 'FAIL' "group refers to unknown id '$r'"; continue }
                if ($grouped.ContainsKey($r)) { Add-Issue 'FAIL' "id '$r' is in two groups" }
                $grouped[$r] = $true; $tags += $ids[$r].LocalName
            }
            # An unfilled oval traced as its edges: a group of 32+ polylines.
            if (@($tags | Where-Object { $_ -eq 'polyline' }).Count -ge 32) { $roundCount++ }
        }
        foreach ($x in $doc.SelectNodes('/i:iwb/i:element', $ns)) {
            if (-not $ids.ContainsKey($x.GetAttribute('ref'))) { Add-Issue 'FAIL' "iwb:element refers to unknown id '$($x.GetAttribute('ref'))'" }
        }

        # ---- Content ----
        $missing = @($Expect.Texts | Where-Object { $outTexts -notcontains $_ })
        if ($missing.Count) { Add-Issue 'FAIL' ("{0} source text(s) not found verbatim: {1}" -f $missing.Count, (($missing | Select-Object -First 3) -join ' | ')) }
        $qmarks = @($outTexts | Where-Object { $_ -eq '?' }).Count
        if ($qmarks) { Add-Issue 'FAIL' "$qmarks '?' fallback glyph(s) present" }
        if ($images -lt $Expect.Stickers) { Add-Issue 'FAIL' "stickers: export has $($Expect.Stickers) with artwork, only $images image(s) in the page" }
        if ($roundCount -lt $Expect.Ovals) { Add-Issue 'FAIL' "ovals: export has $($Expect.Ovals), only $roundCount drawn round" }
        # Each connector is one polyline plus two per arrowhead.
        if ($polylines -lt ($Expect.Connectors + (2 * $Expect.Arrowheads))) {
            Add-Issue 'FAIL' "connectors: expected at least $($Expect.Connectors + (2 * $Expect.Arrowheads)) polylines, found $polylines"
        }
        if ($rotatedText -ne $Expect.RotatedText) { Add-Issue 'FAIL' "rotated text: export has $($Expect.RotatedText), .iwb has $rotatedText" }
        if ($Expect.RotatedNonText -gt 0) { Add-Issue 'WARN' "$($Expect.RotatedNonText) rotated non-shape object(s) in the export (drawn unrotated)" }

        # ---- Converter's own counters ----
        $traced = & $prop 'ShapesTraced'; $fromLabel = & $prop 'ShapesFromLabel'
        if ($null -ne $traced -and $null -ne $fromLabel -and $Expect.Shapes -ne ($traced + $fromLabel)) {
            Add-Issue 'FAIL' "shapes: export has $($Expect.Shapes), converter produced $($traced + $fromLabel)"
        }
        if ($fromLabel -gt 0) { Add-Issue 'WARN' "$fromLabel shape(s) built from the label (outline not traceable)" }
        $arrows = & $prop 'ArrowheadsConverted'
        if ($arrows -ne $Expect.Arrowheads) { Add-Issue 'FAIL' "arrowheads: export has $($Expect.Arrowheads), converter converted $arrows" }
        $rotShapes = & $prop 'RotatedShapesConverted'
        if ($rotShapes -ne $Expect.RotatedShapes) { Add-Issue 'FAIL' "rotated shapes: export has $($Expect.RotatedShapes), converter rotated $rotShapes" }
        $fallback = & $prop 'StickersFallback'
        if ($fallback -gt 0) { Add-Issue 'WARN' "$fallback sticker(s) drawn with a fallback symbol" }
        $unsupported = & $prop 'UnsupportedObjects'
        if ($unsupported -gt 0) { Add-Issue 'WARN' ("unsupported objects: " + ($RunResult.UnsupportedDetails -join ', ')) }

        [pscustomobject]@{ Issues = $issues; Texts = $outTexts.Count; Images = $images }
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

$script = Join-Path $ScriptDirectory 'Convert-WhiteboardHtmlToIwb.ps1'
if (-not (Test-Path -LiteralPath $script)) { throw "Converter not found: $script" }
$rows = New-Object Collections.Generic.List[object]
foreach ($fill in $NoteFill) {
    $outDir = Join-Path $OutputDirectory $fill.ToLowerInvariant()
    foreach ($html in $htmlFiles) {
        $board = $html.BaseName
        $expect = Get-SourceExpectations $html.FullName
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $status = 'PASS'; $messages = @(); $elements = $null; $check = $null
        try {
            $run = & $script -InputPath $html.FullName -OutputDirectory $outDir -NoteFill $fill -Force -PassThru 3>$null 6>$null
            $run = @($run)[-1]
            $elements = $run.IwbElements
            $check = Test-Iwb -IwbPath $run.OutputPath -Expect $expect -RunResult $run
            if (@($check.Issues | Where-Object Level -eq 'FAIL').Count) { $status = 'FAIL' }
            elseif (@($check.Issues | Where-Object Level -eq 'WARN').Count) { $status = 'WARN' }
            $messages = @($check.Issues | ForEach-Object { "$($_.Level): $($_.Message)" })
        } catch {
            $status = 'FAIL'; $messages = @("FAIL: conversion threw: $($_.Exception.Message)")
        }
        $sw.Stop()
        $rows.Add([pscustomobject]@{
            Board = $board; NoteFill = $fill; Status = $status; Ms = $sw.ElapsedMilliseconds
            Elements = $elements
            Texts = $(if ($check) { '{0}/{1}' -f $expect.Texts.Count, $check.Texts } else { '' })
            Stickers = $expect.Stickers; Shapes = $expect.Shapes; Ovals = $expect.Ovals; Arrowheads = $expect.Arrowheads
            Messages = $messages
        })
    }
}

$rows | Format-Table Board, NoteFill, Status, Ms, Elements, @{ n = 'Texts(src/out)'; e = { $_.Texts } }, Shapes, Ovals, Stickers, Arrowheads -AutoSize | Out-String -Width 200 | Write-Host
foreach ($r in $rows | Where-Object { $_.Messages }) {
    Write-Host ("{0} / {1}:" -f $r.Board, $r.NoteFill)
    foreach ($m in $r.Messages) { Write-Host "    $m" }
}
$rows | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'smoke-results.json') -Encoding UTF8

$fails = @($rows | Where-Object Status -eq 'FAIL').Count
$warns = @($rows | Where-Object Status -eq 'WARN').Count
Write-Host ("{0} conversion(s): {1} passed, {2} warned, {3} failed." -f $rows.Count, ($rows.Count - $fails - $warns), $warns, $fails)
if ($fails) { exit 1 } else { exit 0 }
