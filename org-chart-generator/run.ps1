param(
    [Parameter(Mandatory = $false)]
    [string]$Settings = ".\settings.json",

    [Parameter(Mandatory = $false)]
    [ValidateSet("PowerShell", "Python")]
    [string]$Extractor = "PowerShell"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$settingsPath = if ([System.IO.Path]::IsPathRooted($Settings)) {
    [System.IO.Path]::GetFullPath($Settings)
} else {
    [System.IO.Path]::GetFullPath((Join-Path $scriptDir $Settings))
}

if (-not (Test-Path -LiteralPath $settingsPath)) {
    throw "Settings file not found: $settingsPath. Copy settings.example.json to settings.json and edit it first."
}

$buildDir = Join-Path $scriptDir ".build"
New-Item -ItemType Directory -Force -Path $buildDir | Out-Null

$dataPath = Join-Path $buildDir "org-data.json"
if ($Extractor -eq "Python") {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($null -eq $pythonCommand) { throw "Python was not found on PATH. Install Python or use -Extractor PowerShell." }
    & $pythonCommand.Source (Join-Path $scriptDir "extract_workbook.py") --settings $settingsPath --output $dataPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
} else {
    & (Join-Path $scriptDir "extract_workbook.ps1") -Settings $settingsPath -Output $dataPath
}

& (Join-Path $scriptDir "build_deck.ps1") -Settings $settingsPath -Data $dataPath
& (Join-Path $scriptDir "build_hr_handoff.ps1") -Settings $settingsPath -Data $dataPath
