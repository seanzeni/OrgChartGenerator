param(
    [Parameter(Mandatory = $true)]
    [string]$Settings,

    [Parameter(Mandatory = $true)]
    [string]$Data
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-ConfiguredPath {
    param([string]$Value, [string]$BaseDirectory)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "A required path is blank in settings.json" }
    if ([System.IO.Path]::IsPathRooted($Value)) { return [System.IO.Path]::GetFullPath($Value) }
    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $Value))
}

function Get-Text {
    param($Record, [string]$Property)
    if ($null -eq $Record -or $Record.PSObject.Properties.Name -notcontains $Property) { return "" }
    return ([string]$Record.$Property).Trim()
}

function Set-RangeValues {
    param($Sheet, [int]$StartRow, [int]$StartColumn, [object[]]$Rows)
    if ($Rows.Count -eq 0) { return }
    $columnCount = $Rows[0].Count
    $matrix = New-Object 'object[,]' $Rows.Count, $columnCount
    for ($row = 0; $row -lt $Rows.Count; $row++) {
        for ($column = 0; $column -lt $columnCount; $column++) {
            $matrix[$row, $column] = $Rows[$row][$column]
        }
    }
    $range = $Sheet.Range($Sheet.Cells.Item($StartRow, $StartColumn), $Sheet.Cells.Item($StartRow + $Rows.Count - 1, $StartColumn + $columnCount - 1))
    $range.Value2 = $matrix
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) | Out-Null
}

function Format-Header {
    param($Range)
    $Range.Interior.Color = 6567715
    $Range.Font.Color = 16777215
    $Range.Font.Bold = $true
    $Range.HorizontalAlignment = -4108
    $Range.VerticalAlignment = -4108
    $Range.WrapText = $true
    $Range.RowHeight = 30
}

$settingsPath = [System.IO.Path]::GetFullPath($Settings)
$dataPath = [System.IO.Path]::GetFullPath($Data)
$settingsData = Get-Content -Raw -LiteralPath $settingsPath | ConvertFrom-Json
$orgData = Get-Content -Raw -LiteralPath $dataPath | ConvertFrom-Json
$settingsDirectory = Split-Path -Parent $settingsPath

$pptOutput = Resolve-ConfiguredPath ([string]$settingsData.outputPath) $settingsDirectory
$configuredHrPath = if ($settingsData.PSObject.Properties.Name -contains "hrOutputPath") { [string]$settingsData.hrOutputPath } else { "" }
$hrOutputPath = if ($configuredHrPath) {
    Resolve-ConfiguredPath $configuredHrPath $settingsDirectory
} else {
    Join-Path ([System.IO.Path]::GetDirectoryName($pptOutput)) (([System.IO.Path]::GetFileNameWithoutExtension($pptOutput)) + " HR Handoff.xlsx")
}

$currentRecords = @($orgData.current)
$futureRecords = @($orgData.future)
$matchedFutureRows = [System.Collections.Generic.HashSet[int]]::new()
$pairs = [System.Collections.Generic.List[object]]::new()
foreach ($current in $currentRecords) {
    $currentPcn = Get-Text $current "pcn"
    $currentEmployeeId = Get-Text $current "employeeId"
    $future = $null
    if ($currentPcn) {
        $future = @($futureRecords | Where-Object {
            -not $matchedFutureRows.Contains([int]$_.sourceRow) -and
            (Get-Text $_ "pcn").Equals($currentPcn, [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1)
        if ($future.Count -gt 0) { $future = $future[0] } else { $future = $null }
    }
    if ($null -eq $future -and $currentEmployeeId) {
        $future = @($futureRecords | Where-Object {
            -not $matchedFutureRows.Contains([int]$_.sourceRow) -and
            (Get-Text $_ "employeeId").Equals($currentEmployeeId, [System.StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1)
        if ($future.Count -gt 0) { $future = $future[0] } else { $future = $null }
    }
    if ($null -ne $future) { $null = $matchedFutureRows.Add([int]$future.sourceRow) }
    $pairs.Add([pscustomobject]@{ current = $current; future = $future })
}
foreach ($future in $futureRecords) {
    if (-not $matchedFutureRows.Contains([int]$future.sourceRow)) {
        $pairs.Add([pscustomobject]@{ current = $null; future = $future })
    }
}

$movements = [System.Collections.Generic.List[object]]::new()
foreach ($pair in $pairs) {
    $current = $pair.current
    $future = $pair.future
    $currentEmployeeId = Get-Text $current "employeeId"
    $futureEmployeeId = Get-Text $future "employeeId"
    $currentManagerId = Get-Text $current "managerId"
    $futureManagerId = Get-Text $future "managerId"
    $currentRole = Get-Text $current "role"
    $futureRole = Get-Text $future "role"

    $changeType = ""
    if ($null -eq $current) {
        $changeType = if ($futureEmployeeId) { "Added employee" } else { "Added vacant position" }
    } elseif ($null -eq $future) {
        $changeType = if ($currentEmployeeId) { "Removed employee" } else { "Removed vacant position" }
    } elseif ($currentEmployeeId -ne $futureEmployeeId) {
        if (-not $currentEmployeeId -and $futureEmployeeId) { $changeType = "Vacancy filled" }
        elseif ($currentEmployeeId -and -not $futureEmployeeId) { $changeType = "Position vacated" }
        else { $changeType = "Position staffing change" }
    } elseif ($currentManagerId -ne $futureManagerId) {
        $changeType = "Manager transfer"
    } elseif ($currentRole -ne $futureRole) {
        $changeType = "Role change"
    } else {
        continue
    }

    $losingManagerId = if ($currentManagerId -and ($null -eq $future -or $currentManagerId -ne $futureManagerId)) { $currentManagerId } else { "" }
    $losingManagerName = if ($losingManagerId) { Get-Text $current "managerName" } else { "" }
    $gainingManagerId = if ($futureManagerId -and ($null -eq $current -or $currentManagerId -ne $futureManagerId)) { $futureManagerId } else { "" }
    $gainingManagerName = if ($gainingManagerId) { Get-Text $future "managerName" } else { "" }
    $pcn = Get-Text $future "pcn"
    if (-not $pcn) { $pcn = Get-Text $current "pcn" }

    $movements.Add([pscustomobject]@{
        changeType = $changeType
        pcn = $pcn
        currentEmployeeId = $currentEmployeeId
        currentEmployeeName = Get-Text $current "employeeName"
        futureEmployeeId = $futureEmployeeId
        futureEmployeeName = Get-Text $future "employeeName"
        losingManagerId = $losingManagerId
        losingManagerName = $losingManagerName
        gainingManagerId = $gainingManagerId
        gainingManagerName = $gainingManagerName
        currentRole = $currentRole
        futureRole = $futureRole
        currentExplanation = Get-Text $current "explanation"
        futureExplanation = Get-Text $future "explanation"
    })
}

$managerStats = @{}
foreach ($movement in $movements) {
    foreach ($side in @("losing", "gaining")) {
        $idProperty = "${side}ManagerId"
        $nameProperty = "${side}ManagerName"
        $managerId = [string]$movement.$idProperty
        if (-not $managerId) { continue }
        if (-not $managerStats.ContainsKey($managerId)) {
            $managerStats[$managerId] = [ordered]@{
                managerId = $managerId
                managerName = [string]$movement.$nameProperty
                gainedPositions = 0
                lostPositions = 0
                gainedEmployees = 0
                lostEmployees = 0
                gainedVacancies = 0
                lostVacancies = 0
            }
        }
        $stats = $managerStats[$managerId]
        if ($side -eq "gaining") {
            $stats.gainedPositions++
            if ($movement.futureEmployeeId) { $stats.gainedEmployees++ } else { $stats.gainedVacancies++ }
        } else {
            $stats.lostPositions++
            if ($movement.currentEmployeeId) { $stats.lostEmployees++ } else { $stats.lostVacancies++ }
        }
    }
}

$excel = $null
$workbook = $null
try {
    try { $excel = New-Object -ComObject Excel.Application }
    catch { throw "Microsoft Excel desktop is required to create the HR handoff workbook." }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) { $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete() }

    $changesSheet = $workbook.Worksheets.Item(1)
    $changesSheet.Name = "Employee Changes"
    $summarySheet = $workbook.Worksheets.Add()
    $summarySheet.Name = "Manager Summary"
    $summarySheet.Move($workbook.Worksheets.Item(1))

    $summarySheet.Cells.Item(1, 1).Value2 = "Manager gains and losses"
    $summarySheet.Cells.Item(1, 1).Font.Size = 16
    $summarySheet.Cells.Item(1, 1).Font.Bold = $true
    $summarySheet.Cells.Item(2, 1).Value2 = "Compared Current with Future from $([System.IO.Path]::GetFileName([string]$orgData.workbookPath))"
    $summarySheet.Cells.Item(2, 1).Font.Italic = $true
    $summaryHeaders = @("Manager ID", "Manager Name", "Positions Gained", "Positions Lost", "Net Change", "Employees Gained", "Employees Lost", "Vacancies Gained", "Vacancies Lost")
    Set-RangeValues $summarySheet 4 1 (,[object[]]$summaryHeaders)
    $summaryRows = @($managerStats.Values | Sort-Object { $_.managerName } | ForEach-Object {
        ,([object[]]@($_.managerId, $_.managerName, $_.gainedPositions, $_.lostPositions, ($_.gainedPositions - $_.lostPositions), $_.gainedEmployees, $_.lostEmployees, $_.gainedVacancies, $_.lostVacancies))
    })
    if ($summaryRows.Count -gt 0) { Set-RangeValues $summarySheet 5 1 $summaryRows }
    $summaryHeaderRange = $summarySheet.Range("A4:I4")
    Format-Header $summaryHeaderRange
    $summaryDataEnd = [Math]::Max(5, 4 + $summaryRows.Count)
    $summarySheet.Range("A4:I$summaryDataEnd").AutoFilter() | Out-Null
    $summarySheet.Range("A:I").EntireColumn.AutoFit() | Out-Null
    $summarySheet.Columns.Item(2).ColumnWidth = 24
    $summarySheet.Range("C5:I$summaryDataEnd").HorizontalAlignment = -4108
    $summarySheet.Application.ActiveWindow.SplitRow = 4
    $summarySheet.Application.ActiveWindow.FreezePanes = $true

    $changesSheet.Cells.Item(1, 1).Value2 = "Proposed organization changes"
    $changesSheet.Cells.Item(1, 1).Font.Size = 16
    $changesSheet.Cells.Item(1, 1).Font.Bold = $true
    $changesSheet.Cells.Item(2, 1).Value2 = "$($movements.Count) proposed changes requiring HR review"
    $changesSheet.Cells.Item(2, 1).Font.Italic = $true
    $changeHeaders = @("Change Type", "PCN", "Current Employee ID", "Current Employee Name", "Future Employee ID", "Future Employee Name", "Losing Manager ID", "Losing Manager Name", "Gaining Manager ID", "Gaining Manager Name", "Current Role", "Future Role", "Current Explanation", "Future Explanation")
    Set-RangeValues $changesSheet 4 1 (,[object[]]$changeHeaders)
    $changeRows = @($movements | Sort-Object changeType, futureEmployeeName, currentEmployeeName | ForEach-Object {
        ,([object[]]@($_.changeType, $_.pcn, $_.currentEmployeeId, $_.currentEmployeeName, $_.futureEmployeeId, $_.futureEmployeeName, $_.losingManagerId, $_.losingManagerName, $_.gainingManagerId, $_.gainingManagerName, $_.currentRole, $_.futureRole, $_.currentExplanation, $_.futureExplanation))
    })
    if ($changeRows.Count -gt 0) { Set-RangeValues $changesSheet 5 1 $changeRows }
    $changesHeaderRange = $changesSheet.Range("A4:N4")
    Format-Header $changesHeaderRange
    $changesDataEnd = [Math]::Max(5, 4 + $changeRows.Count)
    $changesSheet.Range("A4:N$changesDataEnd").AutoFilter() | Out-Null
    $changesSheet.Range("A:N").EntireColumn.AutoFit() | Out-Null
    $changesSheet.Columns.Item(1).ColumnWidth = 24
    $changesSheet.Columns.Item(4).ColumnWidth = 24
    $changesSheet.Columns.Item(6).ColumnWidth = 24
    $changesSheet.Columns.Item(8).ColumnWidth = 24
    $changesSheet.Columns.Item(10).ColumnWidth = 24
    $changesSheet.Columns.Item(13).ColumnWidth = 42
    $changesSheet.Columns.Item(14).ColumnWidth = 42
    $changesSheet.Range("M5:N$changesDataEnd").WrapText = $true
    $changesSheet.Range("A5:N$changesDataEnd").VerticalAlignment = -4160
    $changesSheet.Application.ActiveWindow.SplitRow = 4
    $changesSheet.Application.ActiveWindow.FreezePanes = $true

    $outputDirectory = Split-Path -Parent $hrOutputPath
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    if (Test-Path -LiteralPath $hrOutputPath) { Remove-Item -LiteralPath $hrOutputPath -Force }
    $workbook.SaveAs($hrOutputPath, 51)
    Write-Output "Created $hrOutputPath"
    Write-Output "HR changes: $($movements.Count)"
    Write-Output "Managers affected: $($managerStats.Count)"
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
