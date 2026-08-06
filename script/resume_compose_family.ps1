param(
    [Parameter(Mandatory = $true)][string]$Plan,
    [Parameter(Mandatory = $true)][string]$Output,
    [int]$ExpectedMounts = 0,
    [int]$MaxAttempts = 30
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$planPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Plan))
$outputPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Output))
$statePath = "$outputPath.compose-state.json"
$provenancePath = "$outputPath.provenance.json"

for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    if (Test-Path -LiteralPath $provenancePath) {
        exit 0
    }
    if ((Test-Path -LiteralPath $statePath) -and $ExpectedMounts -gt 0) {
        $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([int]$state.mounts_done -ge $ExpectedMounts) {
            # One final invocation writes provenance and removes resume files.
            $ExpectedMounts = 0
        }
    }
    $arguments = @(
        'script\template_ppt.py', 'compose-family',
        '--plan', $planPath,
        '--output', $outputPath,
        '--overwrite'
    )
    $process = Start-Process -FilePath python -ArgumentList $arguments -WorkingDirectory $repoRoot -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -eq 0 -and (Test-Path -LiteralPath $provenancePath)) {
        exit 0
    }
    Start-Sleep -Seconds 2
}

exit 1
