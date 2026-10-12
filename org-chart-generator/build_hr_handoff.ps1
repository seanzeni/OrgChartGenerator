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

function Get-PersonName {
    param([string]$Id, $OrgData)
    if (-not $Id) { return "" }
    $person = @($OrgData.employees | Where-Object { $_.id -eq $Id } | Select-Object -First 1)
    if ($person.Count -gt 0) { return [string]$person[0].name }
    $record = @($OrgData.future + $OrgData.current | Where-Object { $_.managerId -eq $Id } | Select-Object -First 1)
    if ($record.Count -gt 0) { return [string]$record[0].managerName }
    return $Id
}

function New-ParentMap {
    param($Records)
    $map = @{}
    foreach ($record in $Records) {
        $employeeId = Get-Text $record "employeeId"
        if ($employeeId) { $map[$employeeId] = Get-Text $record "managerId" }
    }
    return $map
}

function Resolve-DirectorForRecord {
    param($Record, [hashtable]$ParentMap, [string[]]$DirectorIds)
    if ($null -eq $Record) { return "" }
    $current = Get-Text $Record "managerId"
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    while ($current -and $visited.Add($current)) {
        if ($DirectorIds -contains $current) { return $current }
        if (-not $ParentMap.ContainsKey($current)) { return $current }
        $current = [string]$ParentMap[$current]
    }
    return ""
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
$configuredDirectors = @()
if ($settingsData.options.PSObject.Properties.Name -contains "directorIds" -and $null -ne $settingsData.options.directorIds) {
    $configuredDirectors = @($settingsData.options.directorIds | ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } | Where-Object { $_ })
} elseif ($settingsData.options.PSObject.Properties.Name -contains "directorId" -and $settingsData.options.directorId) {
    $configuredDirectors = @(([string]$settingsData.options.directorId).Trim().ToUpperInvariant())
}
$directorIds = @($configuredDirectors | Sort-Object -Unique)
if ($directorIds.Count -eq 0) {
    $roots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($stateLabel in @("current", "future")) {
        $records = if ($stateLabel -eq "current") { @($currentRecords) } else { @($futureRecords) }
        $stateManagerIds = @($records.managerId | Sort-Object -Unique)
        $managedManagers = @($records | Where-Object { $stateManagerIds -contains $_.employeeId } | ForEach-Object { $_.employeeId } | Sort-Object -Unique)
        foreach ($root in @($stateManagerIds | Where-Object { $managedManagers -notcontains $_ })) { $null = $roots.Add([string]$root) }
    }
    $directorIds = @($roots | Sort-Object)
}
$currentParentMap = New-ParentMap $currentRecords
$futureParentMap = New-ParentMap $futureRecords
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
    $currentDirectorId = Resolve-DirectorForRecord $current $currentParentMap $directorIds
    $futureDirectorId = Resolve-DirectorForRecord $future $futureParentMap $directorIds
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
    } elseif ($currentDirectorId -ne $futureDirectorId) {
        $changeType = "Director transfer"
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
        currentDirectorId = $currentDirectorId
        currentDirectorName = Get-PersonName $currentDirectorId $orgData
        futureDirectorId = $futureDirectorId
        futureDirectorName = Get-PersonName $futureDirectorId $orgData
        currentRole = $currentRole
        futureRole = $futureRole
        currentExplanation = Get-Text $current "explanation"
        futureExplanation = Get-Text $future "explanation"
    })
}

$managerStats = @{}
$directorStats = @{}
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
    if ($movement.currentDirectorId -and $movement.currentDirectorId -ne $movement.futureDirectorId) {
        if (-not $directorStats.ContainsKey($movement.currentDirectorId)) {
            $directorStats[$movement.currentDirectorId] = [ordered]@{ directorId = $movement.currentDirectorId; directorName = $movement.currentDirectorName; gainedPositions = 0; lostPositions = 0; gainedEmployees = 0; lostEmployees = 0; gainedVacancies = 0; lostVacancies = 0 }
        }
        $stats = $directorStats[$movement.currentDirectorId]
        $stats.lostPositions++
        if ($movement.currentEmployeeId) { $stats.lostEmployees++ } else { $stats.lostVacancies++ }
    }
    if ($movement.futureDirectorId -and $movement.currentDirectorId -ne $movement.futureDirectorId) {
        if (-not $directorStats.ContainsKey($movement.futureDirectorId)) {
            $directorStats[$movement.futureDirectorId] = [ordered]@{ directorId = $movement.futureDirectorId; directorName = $movement.futureDirectorName; gainedPositions = 0; lostPositions = 0; gainedEmployees = 0; lostEmployees = 0; gainedVacancies = 0; lostVacancies = 0 }
        }
        $stats = $directorStats[$movement.futureDirectorId]
        $stats.gainedPositions++
        if ($movement.futureEmployeeId) { $stats.gainedEmployees++ } else { $stats.gainedVacancies++ }
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
    $directorSheet = $workbook.Worksheets.Add()
    $directorSheet.Name = "Director Summary"
    $directorSheet.Move($changesSheet)

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

    $directorSheet.Cells.Item(1, 1).Value2 = "Director gains and losses"
    $directorSheet.Cells.Item(1, 1).Font.Size = 16
    $directorSheet.Cells.Item(1, 1).Font.Bold = $true
    $directorSheet.Cells.Item(2, 1).Value2 = "Includes cross-director moves resolved through each employee's reporting chain"
    $directorSheet.Cells.Item(2, 1).Font.Italic = $true
    $directorHeaders = @("Director ID", "Director Name", "Positions Gained", "Positions Lost", "Net Change", "Employees Gained", "Employees Lost", "Vacancies Gained", "Vacancies Lost")
    Set-RangeValues $directorSheet 4 1 (,[object[]]$directorHeaders)
    $directorRows = @($directorStats.Values | Sort-Object { $_.directorName } | ForEach-Object {
        ,([object[]]@($_.directorId, $_.directorName, $_.gainedPositions, $_.lostPositions, ($_.gainedPositions - $_.lostPositions), $_.gainedEmployees, $_.lostEmployees, $_.gainedVacancies, $_.lostVacancies))
    })
    if ($directorRows.Count -gt 0) { Set-RangeValues $directorSheet 5 1 $directorRows }
    $directorHeaderRange = $directorSheet.Range("A4:I4")
    Format-Header $directorHeaderRange
    $directorDataEnd = [Math]::Max(5, 4 + $directorRows.Count)
    $directorSheet.Range("A4:I$directorDataEnd").AutoFilter() | Out-Null
    $directorSheet.Range("A:I").EntireColumn.AutoFit() | Out-Null
    $directorSheet.Columns.Item(2).ColumnWidth = 24
    $directorSheet.Range("C5:I$directorDataEnd").HorizontalAlignment = -4108

    $changesSheet.Cells.Item(1, 1).Value2 = "Proposed organization changes"
    $changesSheet.Cells.Item(1, 1).Font.Size = 16
    $changesSheet.Cells.Item(1, 1).Font.Bold = $true
    $changesSheet.Cells.Item(2, 1).Value2 = "$($movements.Count) proposed changes requiring HR review"
    $changesSheet.Cells.Item(2, 1).Font.Italic = $true
    $changeHeaders = @("Change Type", "PCN", "Current Employee ID", "Current Employee Name", "Future Employee ID", "Future Employee Name", "Losing Manager ID", "Losing Manager Name", "Gaining Manager ID", "Gaining Manager Name", "Current Director ID", "Current Director Name", "Future Director ID", "Future Director Name", "Current Role", "Future Role", "Current Explanation", "Future Explanation")
    Set-RangeValues $changesSheet 4 1 (,[object[]]$changeHeaders)
    $changeRows = @($movements | Sort-Object changeType, futureEmployeeName, currentEmployeeName | ForEach-Object {
        ,([object[]]@($_.changeType, $_.pcn, $_.currentEmployeeId, $_.currentEmployeeName, $_.futureEmployeeId, $_.futureEmployeeName, $_.losingManagerId, $_.losingManagerName, $_.gainingManagerId, $_.gainingManagerName, $_.currentDirectorId, $_.currentDirectorName, $_.futureDirectorId, $_.futureDirectorName, $_.currentRole, $_.futureRole, $_.currentExplanation, $_.futureExplanation))
    })
    if ($changeRows.Count -gt 0) { Set-RangeValues $changesSheet 5 1 $changeRows }
    $changesHeaderRange = $changesSheet.Range("A4:R4")
    Format-Header $changesHeaderRange
    $changesDataEnd = [Math]::Max(5, 4 + $changeRows.Count)
    $changesSheet.Range("A4:R$changesDataEnd").AutoFilter() | Out-Null
    $changesSheet.Range("A:R").EntireColumn.AutoFit() | Out-Null
    $changesSheet.Columns.Item(1).ColumnWidth = 24
    $changesSheet.Columns.Item(4).ColumnWidth = 24
    $changesSheet.Columns.Item(6).ColumnWidth = 24
    $changesSheet.Columns.Item(8).ColumnWidth = 24
    $changesSheet.Columns.Item(10).ColumnWidth = 24
    $changesSheet.Columns.Item(12).ColumnWidth = 24
    $changesSheet.Columns.Item(14).ColumnWidth = 24
    $changesSheet.Columns.Item(17).ColumnWidth = 42
    $changesSheet.Columns.Item(18).ColumnWidth = 42
    $changesSheet.Range("Q5:R$changesDataEnd").WrapText = $true
    $changesSheet.Range("A5:R$changesDataEnd").VerticalAlignment = -4160
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
