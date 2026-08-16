param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [string]$Output,
    [double]$TopMinPoints = -10000.0,
    [double]$TopMaxPoints = 100.0,
    [double]$MinimumFontPoints = 24.0,
    [double]$MaxWidthRatio = 1.0,
    [double]$HeightTolerancePoints = 0.8,
    [string]$ShapeName,
    [int[]]$Slides
)

$ErrorActionPreference = 'Stop'
$resolvedInput = (Resolve-Path -LiteralPath $Pptx).Path
$powerPoint = $null
$presentation = $null
$records = [System.Collections.Generic.List[object]]::new()
$issues = [System.Collections.Generic.List[object]]::new()

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($resolvedInput, $true, $false, $false)
    foreach ($slide in $presentation.Slides) {
        if ($Slides -and $Slides.Count -gt 0 -and $Slides -notcontains [int]$slide.SlideIndex) { continue }
        foreach ($shape in $slide.Shapes) {
            try {
                if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame2.HasText -ne -1) { continue }
                if ([double]$shape.Top -lt $TopMinPoints -or [double]$shape.Top -ge $TopMaxPoints) { continue }
                if (-not [string]::IsNullOrWhiteSpace($ShapeName) -and ([string]$shape.Name) -ne $ShapeName) { continue }
                $text = ([string]$shape.TextFrame2.TextRange.Text).Trim()
                if ([string]::IsNullOrWhiteSpace($text)) { continue }
                $fontSize = [double]$shape.TextFrame2.TextRange.Font.Size
                if ($fontSize -lt $MinimumFontPoints) { continue }
                $availableWidth = [double]$shape.Width - [double]$shape.TextFrame2.MarginLeft - [double]$shape.TextFrame2.MarginRight
                $availableHeight = [double]$shape.Height - [double]$shape.TextFrame2.MarginTop - [double]$shape.TextFrame2.MarginBottom
                $boundWidth = [double]$shape.TextFrame2.TextRange.BoundWidth
                $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight
                $lineCount = [int]$shape.TextFrame2.TextRange.Lines().Count
                $widthRatio = if ($availableWidth -gt 0) { $boundWidth / $availableWidth } else { 999.0 }
                $heightFits = $boundHeight -le ($availableHeight + $HeightTolerancePoints)
                $record = [pscustomobject]@{
                    slide = [int]$slide.SlideIndex
                    shape = [string]$shape.Name
                    text = $text
                    visible_chars = ($text -replace '\s', '').Length
                    lines = $lineCount
                    font_size = [math]::Round($fontSize, 2)
                    bound_width = [math]::Round($boundWidth, 2)
                    available_width = [math]::Round($availableWidth, 2)
                    width_ratio = [math]::Round($widthRatio, 4)
                    bound_height = [math]::Round($boundHeight, 2)
                    available_height = [math]::Round($availableHeight, 2)
                    height_fits = $heightFits
                }
                $records.Add($record)
                if ($lineCount -ne 1 -or $widthRatio -gt $MaxWidthRatio -or -not $heightFits) {
                    $issues.Add($record)
                }
            } catch {
                $issues.Add([pscustomobject]@{
                    slide = [int]$slide.SlideIndex
                    shape = [string]$shape.Name
                    error = $_.Exception.Message
                })
            }
        }
    }
    $result = [pscustomobject]@{
        ok = ($issues.Count -eq 0)
        deck = $resolvedInput
        title_count = $records.Count
        top_min_points = $TopMinPoints
        top_max_points = $TopMaxPoints
        shape_name = $ShapeName
        slides = $Slides
        max_width_ratio = $MaxWidthRatio
        issue_count = $issues.Count
        issues = $issues
        titles = $records
    }
    $json = $result | ConvertTo-Json -Depth 8
    if ($Output) {
        $outputPath = if ([System.IO.Path]::IsPathRooted($Output)) {
            [System.IO.Path]::GetFullPath($Output)
        } else {
            [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Output))
        }
        [System.IO.File]::WriteAllText($outputPath, $json + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
    }
    $json
    if ($issues.Count -ne 0) { exit 1 }
}
finally {
    if ($presentation) { $presentation.Close() }
    if ($powerPoint) { $powerPoint.Quit() }
    if ($presentation) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($powerPoint) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
