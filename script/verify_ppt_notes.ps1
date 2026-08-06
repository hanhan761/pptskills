param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$NotesJson
)

$ErrorActionPreference = 'Stop'
$pptxPath = (Resolve-Path -LiteralPath $Pptx).Path
$notesPath = (Resolve-Path -LiteralPath $NotesJson).Path
$payload = Get-Content -LiteralPath $notesPath -Raw -Encoding UTF8 | ConvertFrom-Json
$notes = @($payload.notes)
if ($notes.Count -eq 0) {
    throw 'Notes JSON contains no notes.'
}

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($pptxPath, $true, $false, $false)
    if ($presentation.Slides.Count -ne $notes.Count) {
        throw "Slide/note count mismatch: $($presentation.Slides.Count) slides, $($notes.Count) notes."
    }

    $verified = 0
    $mismatches = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $notes) {
        $slideIndex = [int]$entry.slide
        $expected = ([string]$entry.text).Trim()
        $actual = ''
        $status = 'notes_body_missing'
        try {
            $slide = $presentation.Slides.Item($slideIndex)
            foreach ($shape in $slide.NotesPage.Shapes) {
                if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -eq 2) {
                    $status = 'empty'
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $actual = $shape.TextFrame.TextRange.Text.Trim()
                        $status = 'read'
                    }
                    break
                }
            }
        } catch {
            $status = "read_error: $($_.Exception.Message)"
        }

        if ($actual -eq $expected) {
            $verified++
        } else {
            $mismatches.Add([pscustomobject]@{
                slide = $slideIndex
                status = $status
                expected = $expected
                actual = $actual
            })
        }
    }

    [pscustomobject]@{
        deck = $pptxPath
        slides = $presentation.Slides.Count
        notes_verified = $verified
        mismatch_count = $mismatches.Count
        mismatches = $mismatches
        passed = ($verified -eq $notes.Count)
    } | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
