param(
    [Parameter(Mandatory = $true)]
    [string]$Settings,

    [Parameter(Mandatory = $true)]
    [string]$Output
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-NormalizedText {
    param($Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Assert-FourCharacterId {
    param(
        [string]$Value,
        [string]$Label,
        [string]$SheetName,
        [int]$RowNumber
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Sheet '$SheetName', row $RowNumber`: $Label is blank"
    }
    if ($Value -notmatch '^[A-Za-z0-9]{4}$') {
        throw "Sheet '$SheetName', row $RowNumber`: $Label '$Value' must contain exactly four letters or numbers"
    }
}

function Get-Worksheet {
    param($Workbook, [string]$Name)
    try {
        return $Workbook.Worksheets.Item($Name)
    } catch {
        $available = @($Workbook.Worksheets | ForEach-Object { $_.Name }) -join ", "
        throw "Sheet '$Name' was not found. Available sheets: $available"
    }
}

function Get-HeaderMap {
    param($Worksheet, [string[]]$RequiredHeaders)
    $map = @{}
    $columnCount = [int]$Worksheet.UsedRange.Columns.Count
    for ($column = 1; $column -le $columnCount; $column++) {
        $header = Get-NormalizedText $Worksheet.Cells.Item(1, $column).Value2
        if ($header) { $map[$header.ToLowerInvariant()] = $column }
    }
    $missing = @($RequiredHeaders | Where-Object { -not $map.ContainsKey($_.ToLowerInvariant()) })
    if ($missing.Count -gt 0) {
        throw "Sheet '$($Worksheet.Name)' is missing columns: $($missing -join ', ')"
    }
    return $map
}

function Convert-ExcelColorToHex {
    param($Cell)
    $none = -4142
    if ([int]$Cell.Interior.ColorIndex -eq $none) {
        throw "Sheet '$($Cell.Worksheet.Name)', cell $($Cell.Address($false, $false)): choose a solid fill color"
    }
    $packed = [long]$Cell.Interior.Color
    $red = $packed -band 0xFF
    $green = ($packed -shr 8) -band 0xFF
    $blue = ($packed -shr 16) -band 0xFF
    return ('#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue)
}

function Read-Employees {
    param($Workbook, [string]$SheetName)
    $sheet = Get-Worksheet $Workbook $SheetName
    $headers = Get-HeaderMap $sheet @("Employee ID", "Employee Name", "Manager")
    $employees = @{}
    $lastRow = [int]$sheet.UsedRange.Rows.Count
    for ($row = 2; $row -le $lastRow; $row++) {
        $id = (Get-NormalizedText $sheet.Cells.Item($row, $headers["employee id"]).Value2).ToUpperInvariant()
        $name = Get-NormalizedText $sheet.Cells.Item($row, $headers["employee name"]).Value2
        $managerValue = Get-NormalizedText $sheet.Cells.Item($row, $headers["manager"]).Value2
        if (-not $id -and -not $name -and -not $managerValue) { continue }
        Assert-FourCharacterId $id "employee ID" $SheetName $row
        if (-not $name) { throw "Sheet '$SheetName', row $row`: employee name is blank" }
        if ($employees.ContainsKey($id)) { throw "Sheet '$SheetName' has duplicate employee ID '$id'" }
        $managerMarkers = @("☑ yes", "yes", "true", "1", "x", "checked", "manager")
        $isManager = $managerMarkers -contains $managerValue.ToLowerInvariant()
        $employees[$id] = [ordered]@{
            id = $id
            name = $name
            isManager = $isManager
            sourceRow = $row
        }
    }
    if ($employees.Count -eq 0) { throw "Sheet '$SheetName' does not contain any employees" }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) | Out-Null
    return $employees
}

function Read-RoleColors {
    param($Workbook, [string]$SheetName)
    $sheet = Get-Worksheet $Workbook $SheetName
    $headers = Get-HeaderMap $sheet @("Role Type", "Slide Color")
    $roles = @{}
    $lastRow = [int]$sheet.UsedRange.Rows.Count
    for ($row = 2; $row -le $lastRow; $row++) {
        $role = Get-NormalizedText $sheet.Cells.Item($row, $headers["role type"]).Value2
        if (-not $role) { continue }
        $key = $role.ToLowerInvariant()
        if ($roles.ContainsKey($key)) { throw "Sheet '$SheetName' has duplicate role type '$role'" }
        $roles[$key] = [ordered]@{
            name = $role
            color = Convert-ExcelColorToHex $sheet.Cells.Item($row, $headers["slide color"])
        }
    }
    if ($roles.Count -eq 0) { throw "Sheet '$SheetName' does not contain any role types" }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) | Out-Null
    return $roles
}

function Read-StateSheet {
    param(
        $Workbook,
        [string]$SheetName,
        $Columns,
        [hashtable]$Employees,
        [hashtable]$RoleColors
    )
    $sheet = Get-Worksheet $Workbook $SheetName
    $employeeHeader = [string]$Columns.employeeId
    $managerHeader = [string]$Columns.managerId
    $roleHeader = [string]$Columns.role
    $headers = Get-HeaderMap $sheet @($employeeHeader, $managerHeader, $roleHeader)
    $records = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    $lastRow = [int]$sheet.UsedRange.Rows.Count
    for ($row = 2; $row -le $lastRow; $row++) {
        $employeeId = (Get-NormalizedText $sheet.Cells.Item($row, $headers[$employeeHeader.ToLowerInvariant()]).Value2).ToUpperInvariant()
        $managerId = (Get-NormalizedText $sheet.Cells.Item($row, $headers[$managerHeader.ToLowerInvariant()]).Value2).ToUpperInvariant()
        $role = Get-NormalizedText $sheet.Cells.Item($row, $headers[$roleHeader.ToLowerInvariant()]).Value2
        if (-not $employeeId -and -not $managerId -and -not $role) { continue }
        Assert-FourCharacterId $employeeId "employee ID" $SheetName $row
        Assert-FourCharacterId $managerId "manager ID" $SheetName $row
        if (-not $Employees.ContainsKey($employeeId)) {
            throw "Sheet '$SheetName', row $row`: employee ID '$employeeId' is not on Employees"
        }
        if (-not $Employees.ContainsKey($managerId)) {
            throw "Sheet '$SheetName', row $row`: manager ID '$managerId' is not on Employees"
        }
        if (-not [bool]$Employees[$managerId].isManager) {
            throw "Sheet '$SheetName', row $row`: '$managerId' is not marked as a manager"
        }
        if (-not $role) { $role = "Unspecified" }
        $roleKey = $role.ToLowerInvariant()
        if (-not $RoleColors.ContainsKey($roleKey)) {
            throw "Sheet '$SheetName', row $row`: role '$role' is not on Role Settings"
        }
        if ($seen.ContainsKey($employeeId)) {
            throw "Sheet '$SheetName' has duplicate employee ID '$employeeId' on rows $($seen[$employeeId]) and $row"
        }
        $seen[$employeeId] = $row
        $records.Add([ordered]@{
            employeeId = $employeeId
            employeeName = $Employees[$employeeId].name
            managerId = $managerId
            managerName = $Employees[$managerId].name
            role = $RoleColors[$roleKey].name
            sourceRow = $row
        })
    }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) | Out-Null
    return $records
}

$settingsPath = [System.IO.Path]::GetFullPath($Settings)
$settingsData = Get-Content -Raw -LiteralPath $settingsPath | ConvertFrom-Json
$settingsDirectory = Split-Path -Parent $settingsPath
$workbookPath = [string]$settingsData.workbookPath
if (-not [System.IO.Path]::IsPathRooted($workbookPath)) {
    $workbookPath = Join-Path $settingsDirectory $workbookPath
}
$workbookPath = [System.IO.Path]::GetFullPath($workbookPath)
if (-not (Test-Path -LiteralPath $workbookPath)) { throw "Workbook not found: $workbookPath" }

$excel = $null
$workbook = $null
try {
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw "Microsoft Excel desktop is required for the PowerShell extractor. Use -Extractor Python as a fallback."
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $employeesSheet = if ($settingsData.sheets.employees) { [string]$settingsData.sheets.employees } else { "Employees" }
    $rolesSheet = if ($settingsData.sheets.roles) { [string]$settingsData.sheets.roles } else { "Role Settings" }
    $employeeMap = Read-Employees $workbook $employeesSheet
    $roleMap = Read-RoleColors $workbook $rolesSheet
    $roleOutput = [ordered]@{}
    foreach ($item in $roleMap.Values) { $roleOutput[$item.name] = $item.color }
    $payload = [ordered]@{
        workbookPath = $workbookPath
        roleColors = $roleOutput
        employees = @($employeeMap.Values)
        current = @(Read-StateSheet $workbook ([string]$settingsData.sheets.current) $settingsData.columns $employeeMap $roleMap)
        future = @(Read-StateSheet $workbook ([string]$settingsData.sheets.future) $settingsData.columns $employeeMap $roleMap)
    }
    $outputPath = [System.IO.Path]::GetFullPath($Output)
    $payload | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $outputPath -Encoding UTF8
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
