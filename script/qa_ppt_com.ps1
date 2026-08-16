param(
    [Parameter(Mandatory = $true)]
    [string]$Pptx,
    [string]$Output,
    [string]$ShapePrefix,
    [double]$TolerancePoints = 0.8,
    [int[]]$Slides,
    [switch]$AllowDuplicateNames
)

$ErrorActionPreference = 'Stop'
$resolvedInput = (Resolve-Path -LiteralPath $Pptx).Path
$powerPoint = $null
$presentation = $null
$issues = [System.Collections.Generic.List[object]]::new()

function Add-Issue {
    param([int]$Slide, [string]$Kind, [string]$Shape, [string]$Detail)
    $issues.Add([pscustomobject]@{
        slide = $Slide
        kind = $Kind
        shape = $Shape
        detail = $Detail
    })
}

function Test-ShapeText {
    param($Shape, [int]$SlideNumber, [string]$Path)
    $inScope = [string]::IsNullOrWhiteSpace($ShapePrefix) -or ([string]$Shape.Name).StartsWith($ShapePrefix)
    try {
        if ($inScope -and $Shape.HasTextFrame -eq -1 -and $Shape.TextFrame2.HasText -eq -1) {
            $availableWidth = [double]$Shape.Width - [double]$Shape.TextFrame2.MarginLeft - [double]$Shape.TextFrame2.MarginRight
            $availableHeight = [double]$Shape.Height - [double]$Shape.TextFrame2.MarginTop - [double]$Shape.TextFrame2.MarginBottom
            $boundWidth = [double]$Shape.TextFrame2.TextRange.BoundWidth
            $boundHeight = [double]$Shape.TextFrame2.TextRange.BoundHeight
            if ($boundWidth -gt $availableWidth + $TolerancePoints) {
                Add-Issue $SlideNumber 'text_width_overflow' $Path ("bound={0:N2}, available={1:N2}" -f $boundWidth, $availableWidth)
            }
            if ($boundHeight -gt $availableHeight + $TolerancePoints) {
                Add-Issue $SlideNumber 'text_height_overflow' $Path ("bound={0:N2}, available={1:N2}" -f $boundHeight, $availableHeight)
            }
        }
    } catch {
        Add-Issue $SlideNumber 'text_check_error' $Path $_.Exception.Message
    }

    try {
        if ([int]$Shape.Type -eq 6) {
            for ($index = 1; $index -le $Shape.GroupItems.Count; $index++) {
                $child = $Shape.GroupItems.Item($index)
                Test-ShapeText $child $SlideNumber ("$Path/$($child.Name)")
            }
        }
    } catch {
        Add-Issue $SlideNumber 'group_check_error' $Path $_.Exception.Message
    }
}

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($resolvedInput, $true, $false, $false)
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight

    foreach ($slide in $presentation.Slides) {
        if ($Slides -and $Slides.Count -gt 0 -and $Slides -notcontains [int]$slide.SlideIndex) { continue }
        $seenNames = @{}
        foreach ($shape in $slide.Shapes) {
            $name = [string]$shape.Name
            $inScope = [string]::IsNullOrWhiteSpace($ShapePrefix) -or $name.StartsWith($ShapePrefix)
            if ($seenNames.ContainsKey($name)) {
                $seenNames[$name] = [int]$seenNames[$name] + 1
            } else {
                $seenNames[$name] = 1
            }
            $shapePath = "$name#$($seenNames[$name])"
            if ($seenNames[$name] -gt 1 -and $inScope -and -not $AllowDuplicateNames) {
                Add-Issue $slide.SlideIndex 'duplicate_shape_name' $shapePath 'duplicate top-level name'
            }

            $left = [double]$shape.Left
            $top = [double]$shape.Top
            $right = $left + [double]$shape.Width
            $bottom = $top + [double]$shape.Height
            if ($inScope -and ($left -lt -$TolerancePoints -or $top -lt -$TolerancePoints -or $right -gt $slideWidth + $TolerancePoints -or $bottom -gt $slideHeight + $TolerancePoints)) {
                Add-Issue $slide.SlideIndex 'shape_out_of_bounds' $name ("L={0:N2}, T={1:N2}, R={2:N2}, B={3:N2}" -f $left, $top, $right, $bottom)
            }
            Test-ShapeText $shape $slide.SlideIndex $shapePath
        }
    }

    $result = [pscustomobject]@{
        ok = ($issues.Count -eq 0)
        deck = $resolvedInput
        slide_count = $presentation.Slides.Count
        issue_count = $issues.Count
        issues = $issues
    }
    $json = $result | ConvertTo-Json -Depth 8
    if ($Output) {
        $outputPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Output))
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
