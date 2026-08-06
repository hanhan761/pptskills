param(
    [Parameter(Mandatory = $true)][string]$InputPptx,
    [Parameter(Mandatory = $true)][string]$NotesJson,
    [Parameter(Mandatory = $true)][string]$OutputPptx
)

$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputPptx).Path
$notesPath = (Resolve-Path -LiteralPath $NotesJson).Path
$outputPath = [System.IO.Path]::GetFullPath($OutputPptx)
$outputDir = [System.IO.Path]::GetDirectoryName($outputPath)
if (-not [System.IO.Directory]::Exists($outputDir)) {
    [System.IO.Directory]::CreateDirectory($outputDir) | Out-Null
}
if ($inputPath -ne $outputPath) {
    Copy-Item -LiteralPath $inputPath -Destination $outputPath -Force
}

$payload = Get-Content -LiteralPath $notesPath -Raw -Encoding UTF8 | ConvertFrom-Json
$notes = @($payload.notes)
if ($notes.Count -eq 0) {
    throw 'Notes JSON contains no notes.'
}

$powerPoint = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $powerPoint.Presentations.Open($outputPath, $false, $false, $false)
    if ($presentation.Slides.Count -ne $notes.Count) {
        throw "Slide/note count mismatch: $($presentation.Slides.Count) slides, $($notes.Count) notes."
    }

    foreach ($entry in $notes) {
        $slideIndex = [int]$entry.slide
        if ($slideIndex -lt 1 -or $slideIndex -gt $presentation.Slides.Count) {
            throw "Invalid slide number in notes JSON: $slideIndex"
        }
        $slide = $presentation.Slides.Item($slideIndex)
        $body = $null
        foreach ($shape in $slide.NotesPage.Shapes) {
            if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -eq 2) {
                $body = $shape
                break
            }
        }
        if ($null -eq $body) {
            throw "Slide $slideIndex has no notes body placeholder."
        }
        $body.TextFrame.TextRange.Text = [string]$entry.text
    }
    $presentation.Save()
    $presentation.Close()
    $presentation = $null

    $presentation = $powerPoint.Presentations.Open($outputPath, $true, $false, $false)
    $verified = 0
    $mismatches = @()
    foreach ($entry in $notes) {
        $slide = $presentation.Slides.Item([int]$entry.slide)
        $actual = ''
        foreach ($shape in $slide.NotesPage.Shapes) {
            if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -eq 2) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $actual = $shape.TextFrame.TextRange.Text.Trim()
                }
                break
            }
        }
        if ($actual -eq ([string]$entry.text).Trim()) {
            $verified++
        } else {
            $mismatches += [int]$entry.slide
        }
    }
    [pscustomobject]@{
        output = $outputPath
        slides = $presentation.Slides.Count
        notes_verified = $verified
        mismatches = $mismatches
        passed = ($verified -eq $notes.Count)
    } | ConvertTo-Json -Depth 4
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
    }
    try { $powerPoint.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
