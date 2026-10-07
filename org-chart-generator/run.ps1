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

$runtimeRoot = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies"
$nodeExe = Join-Path $runtimeRoot "node\bin\node.exe"
$nodeModules = Join-Path $runtimeRoot "node\node_modules"

if (-not (Test-Path -LiteralPath $nodeExe)) { throw "Codex Node runtime not found: $nodeExe" }
if (-not (Test-Path -LiteralPath $nodeModules)) { throw "Codex Node packages not found: $nodeModules" }

$buildDir = Join-Path $scriptDir ".build"
New-Item -ItemType Directory -Force -Path $buildDir | Out-Null
$nodeLink = Join-Path $scriptDir "node_modules"
if (-not (Test-Path -LiteralPath $nodeLink)) {
    New-Item -ItemType Junction -Path $nodeLink -Target $nodeModules | Out-Null
}

$dataPath = Join-Path $buildDir "org-data.json"
if ($Extractor -eq "Python") {
    $pythonExe = Join-Path $runtimeRoot "python\python.exe"
    if (-not (Test-Path -LiteralPath $pythonExe)) { throw "Codex Python runtime not found: $pythonExe" }
    & $pythonExe (Join-Path $scriptDir "extract_workbook.py") --settings $settingsPath --output $dataPath
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
} else {
    & (Join-Path $scriptDir "extract_workbook.ps1") -Settings $settingsPath -Output $dataPath
}

$env:ORG_SETTINGS = $settingsPath
$env:ORG_DATA = $dataPath
& $nodeExe (Join-Path $scriptDir "build_deck.mjs")
exit $LASTEXITCODE
