param(
    [Parameter(Mandatory = $true)]
    [string]$Settings,

    [Parameter(Mandatory = $true)]
    [string]$Data
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$msoFalse = 0
$msoTrue = -1
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeOval = 9
$ppLayoutBlank = 12
$ppSaveAsOpenXMLPresentation = 24
$ppAlignLeft = 1
$ppAlignCenter = 2
$ppAlignRight = 3
$ppMouseClick = 1
$ppActionHyperlink = 7
$msoAnchorMiddle = 3
$msoAutoSizeTextToFitShape = 2

function Resolve-ConfiguredPath {
    param([string]$Value, [string]$BaseDirectory)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "A required path is blank in settings.json" }
    if ([System.IO.Path]::IsPathRooted($Value)) { return [System.IO.Path]::GetFullPath($Value) }
    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $Value))
}

function Convert-HexToOfficeColor {
    param([string]$Hex)
    $value = $Hex.Trim().TrimStart('#')
    if ($value -notmatch '^[0-9A-Fa-f]{6}$') { $value = "5B7FA3" }
    $red = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return $red + ($green -shl 8) + ($blue -shl 16)
}

function Get-TextColorHex {
    param([string]$Background)
    $value = $Background.Trim().TrimStart('#')
    if ($value -notmatch '^[0-9A-Fa-f]{6}$') { return "#FFFFFF" }
    $red = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($value.Substring(4, 2), 16)
    if ((0.299 * $red + 0.587 * $green + 0.114 * $blue) -gt 155) { return "#172033" }
    return "#FFFFFF"
}

function Add-TextBox {
    param(
        $Slide,
        [string]$Text,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [double]$FontSize = 12,
        [string]$Color = "#172033",
        [bool]$Bold = $false,
        [int]$Alignment = 1,
        [string]$FontName = "Aptos"
    )
    $shape = $Slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $Left, $Top, $Width, $Height)
    $shape.Fill.Visible = $msoFalse
    $shape.Line.Visible = $msoFalse
    $shape.TextFrame2.MarginLeft = 0
    $shape.TextFrame2.MarginRight = 0
    $shape.TextFrame2.MarginTop = 0
    $shape.TextFrame2.MarginBottom = 0
    $shape.TextFrame2.WordWrap = $msoTrue
    $shape.TextFrame2.AutoSize = $msoAutoSizeTextToFitShape
    $shape.TextFrame2.VerticalAnchor = $msoAnchorMiddle
    $shape.TextFrame2.TextRange.Text = $Text
    $shape.TextFrame2.TextRange.Font.Name = $FontName
    $shape.TextFrame2.TextRange.Font.Size = $FontSize
    $shape.TextFrame2.TextRange.Font.Bold = if ($Bold) { $msoTrue } else { $msoFalse }
    $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexToOfficeColor $Color
    $shape.TextFrame2.TextRange.ParagraphFormat.Alignment = $Alignment
    return $shape
}

function Add-FilledShape {
    param(
        $Slide,
        [int]$ShapeType,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Fill,
        [string]$Line = $Fill
    )
    $shape = $Slide.Shapes.AddShape($ShapeType, $Left, $Top, $Width, $Height)
    $shape.Fill.Visible = $msoTrue
    $shape.Fill.Solid()
    $shape.Fill.ForeColor.RGB = Convert-HexToOfficeColor $Fill
    $shape.Line.Visible = $msoTrue
    $shape.Line.ForeColor.RGB = Convert-HexToOfficeColor $Line
    $shape.Line.Weight = 0.75
    return $shape
}

function Get-Initials {
    param([string]$Name)
    $parts = @($Name -split '\s+' | Where-Object { $_ })
    if ($parts.Count -eq 0) { return "?" }
    return (($parts | Select-Object -First 2 | ForEach-Object { $_.Substring(0, 1).ToUpperInvariant() }) -join "")
}

function Add-Portrait {
    param(
        $Slide,
        $Person,
        [double]$Left,
        [double]$Top,
        [double]$Size,
        [string]$Accent,
        [hashtable]$PhotoIndex,
        [System.Collections.Generic.HashSet[string]]$MissingPhotos,
        [string]$FontName
    )
    $personId = if ($Person.PSObject.Properties.Name -contains "id") { [string]$Person.id } else { [string]$Person.employeeId }
    $personName = if ($Person.PSObject.Properties.Name -contains "name") { [string]$Person.name } else { [string]$Person.employeeName }
    $isVacant = $Person.PSObject.Properties.Name -contains "isVacant" -and [bool]$Person.isVacant
    if ($isVacant) {
        $portrait = Add-FilledShape $Slide $msoShapeOval $Left $Top $Size $Size "#FFFFFF" $Accent
        $portrait.Line.DashStyle = 2
        return $portrait
    }
    $photoPath = if ($PhotoIndex.ContainsKey($personId)) { $PhotoIndex[$personId] } else { $null }
    $portrait = $null
    if ($photoPath) {
        try {
            $portrait = $Slide.Shapes.AddShape($msoShapeOval, $Left, $Top, $Size, $Size)
            $portrait.Line.Visible = $msoFalse
            $portrait.Fill.UserPicture($photoPath)
        } catch {
            if ($null -ne $portrait) { $portrait.Delete() }
            $portrait = $null
        }
    }
    if ($null -eq $portrait) {
        $null = $MissingPhotos.Add("$personId`t$personName")
        $portrait = Add-FilledShape $Slide $msoShapeOval $Left $Top $Size $Size $Accent $Accent
        $null = Add-TextBox $Slide (Get-Initials $personName) $Left $Top $Size $Size ([Math]::Max(9, $Size * 0.34)) (Get-TextColorHex $Accent) $true $ppAlignCenter $FontName
    }
    return $portrait
}

function Add-ExplanationCue {
    param($Slide, $Shape, [string]$Explanation, [bool]$HighlightOutline = $true)
    if ([string]::IsNullOrWhiteSpace($Explanation) -or $null -eq $Shape) { return }
    if ($HighlightOutline) {
        $Shape.Line.Visible = $msoTrue
        $Shape.Line.ForeColor.RGB = Convert-HexToOfficeColor "#D62828"
        $Shape.Line.Weight = 2.25
    }
    $Shape.AlternativeText = $Explanation
    try {
        $subAddress = "$($Slide.SlideID),$($Slide.SlideIndex),$($Slide.Name)"
        $action = $Shape.ActionSettings.Item($ppMouseClick)
        $action.Action = $ppActionHyperlink
        $action.Hyperlink.Address = ""
        $action.Hyperlink.SubAddress = $subAddress
        $action.Hyperlink.ScreenTip = $Explanation
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($action) | Out-Null
    } catch {
        # AlternativeText remains available if this PowerPoint version rejects shape ScreenTips.
    }
}

function New-PhotoIndex {
    param([string]$Folder, [string[]]$Extensions)
    if (-not (Test-Path -LiteralPath $Folder)) { throw "Photo folder not found: $Folder" }
    $index = @{}
    foreach ($file in Get-ChildItem -LiteralPath $Folder -File) {
        if ($Extensions -notcontains $file.Extension.ToLowerInvariant()) { continue }
        foreach ($match in [regex]::Matches($file.BaseName, '[A-Za-z0-9]{4}')) {
            $id = $match.Value.ToUpperInvariant()
            if (-not $index.ContainsKey($id)) { $index[$id] = $file.FullName }
        }
    }
    return $index
}

function Get-ManagerName {
    param([string]$Id, $CurrentRecords, $FutureRecords)
    $record = @($FutureRecords | Where-Object { $_.managerId -eq $Id } | Select-Object -First 1)
    if ($record.Count -eq 0) { $record = @($CurrentRecords | Where-Object { $_.managerId -eq $Id } | Select-Object -First 1) }
    if ($record.Count -gt 0) { return [string]$record[0].managerName }
    return $Id
}

function Get-ManagerOrder {
    param($CurrentRecords, $FutureRecords, [string]$ConfiguredDirector)
    $managerIds = @($CurrentRecords.managerId + $FutureRecords.managerId | Sort-Object -Unique)
    if ($managerIds.Count -eq 0) { throw "No managers were found in Current or Future" }
    $director = $ConfiguredDirector.Trim().ToUpperInvariant()
    if ($director -and $managerIds -notcontains $director) { throw "Configured directorId '$director' is not a manager" }
    $allRecords = @($FutureRecords) + @($CurrentRecords)
    if (-not $director) {
        $managedManagers = @($allRecords | Where-Object { $managerIds -contains $_.employeeId } | ForEach-Object { $_.employeeId } | Sort-Object -Unique)
        $roots = @($managerIds | Where-Object { $managedManagers -notcontains $_ })
        if ($roots.Count -eq 1) { $director = $roots[0] }
        else {
            $director = $managerIds | Sort-Object @{ Expression = {
                $id = $_
                -(@($allRecords | Where-Object { $_.managerId -eq $id -and $managerIds -contains $_.employeeId }).Count)
            } }, @{ Expression = { $_ } } | Select-Object -First 1
        }
    }
    $children = @{}
    foreach ($id in $managerIds) { $children[$id] = [System.Collections.Generic.List[string]]::new() }
    foreach ($id in $managerIds) {
        if ($id -eq $director) { continue }
        $employeeRecord = @($FutureRecords | Where-Object { $_.employeeId -eq $id } | Select-Object -First 1)
        if ($employeeRecord.Count -eq 0) { $employeeRecord = @($CurrentRecords | Where-Object { $_.employeeId -eq $id } | Select-Object -First 1) }
        if ($employeeRecord.Count -gt 0) {
            $parent = [string]$employeeRecord[0].managerId
            if ($children.ContainsKey($parent) -and $parent -ne $id) { $children[$parent].Add($id) }
        }
    }
    foreach ($id in $managerIds) {
        $sorted = @($children[$id] | Sort-Object { Get-ManagerName $_ $CurrentRecords $FutureRecords })
        $children[$id].Clear()
        foreach ($child in $sorted) { $children[$id].Add($child) }
    }
    $ordered = [System.Collections.Generic.List[string]]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($director)
    while ($queue.Count -gt 0) {
        $id = $queue.Dequeue()
        if (-not $visited.Add($id)) { continue }
        $ordered.Add($id)
        foreach ($child in $children[$id]) { $queue.Enqueue($child) }
    }
    foreach ($id in $managerIds | Sort-Object { Get-ManagerName $_ $CurrentRecords $FutureRecords }) {
        if ($visited.Add($id)) { $ordered.Add($id) }
    }
    return [pscustomobject]@{ Director = $director; Ordered = @($ordered) }
}

function Get-RolePages {
    param($Members, [int]$MaxCategories, [int]$MaxPeoplePerRow, [int]$MaxRowsPerCategory)
    $roleGroups = @($Members | Group-Object role | Sort-Object Name)
    if ($roleGroups.Count -eq 0) { return @([pscustomobject]@{ Panels = @() }) }
    $pages = [System.Collections.Generic.List[object]]::new()
    for ($start = 0; $start -lt $roleGroups.Count; $start += $MaxCategories) {
        $end = [Math]::Min($start + $MaxCategories - 1, $roleGroups.Count - 1)
        $roleBatch = @($roleGroups[$start..$end])
        $rowsAvailable = if ($roleBatch.Count -le 2) { $MaxRowsPerCategory } else { [Math]::Min(2, $MaxRowsPerCategory) }
        $peoplePerCategory = $MaxPeoplePerRow * $rowsAvailable
        $batchPageCount = 1
        foreach ($group in $roleBatch) {
            $batchPageCount = [Math]::Max($batchPageCount, [Math]::Ceiling($group.Count / [double]$PeoplePerCategory))
        }
        for ($pageIndex = 0; $pageIndex -lt $batchPageCount; $pageIndex++) {
            $panels = [System.Collections.Generic.List[object]]::new()
            foreach ($group in $roleBatch) {
                $memberStart = $pageIndex * $PeoplePerCategory
                if ($memberStart -ge $group.Count) { continue }
                $memberEnd = [Math]::Min($memberStart + $PeoplePerCategory - 1, $group.Count - 1)
                $allMembers = @($group.Group | Sort-Object name)
                $panels.Add([pscustomobject]@{
                    Role = $group.Name
                    Members = @($allMembers[$memberStart..$memberEnd])
                    TotalCount = $group.Count
                })
            }
            $pages.Add([pscustomobject]@{ Panels = @($panels) })
        }
    }
    return @($pages)
}

function Get-RoleColor {
    param($RoleColors, [string]$Role)
    foreach ($property in $RoleColors.PSObject.Properties) {
        if ($property.Name.Equals($Role, [System.StringComparison]::OrdinalIgnoreCase)) { return [string]$property.Value }
    }
    foreach ($property in $RoleColors.PSObject.Properties) {
        if ($property.Name.Equals("Unspecified", [System.StringComparison]::OrdinalIgnoreCase)) { return [string]$property.Value }
    }
    return "#5B7FA3"
}

function Add-TeamSlide {
    param(
        $Presentation,
        [string]$ManagerId,
        [string]$ManagerName,
        [string]$ManagerRole,
        $Members,
        [string]$StateLabel,
        [string]$StateTitle,
        [int]$CurrentCount,
        [int]$FutureCount,
        [int]$PageNumber,
        [int]$TotalPages,
        $Panels,
        [string]$DirectorId,
        $RoleColors,
        [hashtable]$PhotoIndex,
        [System.Collections.Generic.HashSet[string]]$MissingPhotos,
        $Options
    )
    $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, $ppLayoutBlank)
    $background = Add-FilledShape $slide $msoShapeRectangle 0 0 960 540 "#F5F7FA" "#F5F7FA"
    $background.ZOrder(1)
    $titleSuffix = if ($PageNumber -gt 1) { " (continued $PageNumber/$TotalPages)" } else { "" }
    $titleFontSize = if ($PageNumber -gt 1) { 18 } else { 22 }
    $null = Add-TextBox $slide "$ManagerName, $StateTitle$titleSuffix" 36 10 690 32 $titleFontSize "#172033" $true $ppAlignLeft $Options.fontFamily
    $count = $Members.Count
    $vacantCount = @($Members | Where-Object { $_.PSObject.Properties.Name -contains "isVacant" -and [bool]$_.isVacant }).Count
    $countLabel = if ($vacantCount -gt 0) { "$count positions, $vacantCount vacant" } else { "$count employees" }
    $countText = if ($StateLabel -eq "future") {
        $delta = $FutureCount - $CurrentCount
        $deltaText = if ($delta -gt 0) { "+$delta" } else { [string]$delta }
        if ($vacantCount -gt 0) { "$countLabel ($deltaText)" } else { "$countLabel ($deltaText vs current)" }
    } else { $countLabel }
    $countColor = "#445064"
    if ($StateLabel -eq "future" -and $FutureCount -gt $CurrentCount) { $countColor = "#1B7F45" }
    if ($StateLabel -eq "future" -and $FutureCount -lt $CurrentCount) { $countColor = "#B54237" }
    $null = Add-TextBox $slide $countText 730 10 194 32 14 $countColor $true $ppAlignRight $Options.fontFamily

    $null = Add-FilledShape $slide $msoShapeRoundedRectangle 36 50 888 70 "#FFFFFF" "#D9E0E8"
    $managerAccent = if ($ManagerId -eq $DirectorId) { "#203864" } else { "#44546A" }
    $managerPerson = [pscustomobject]@{ id = $ManagerId; name = $ManagerName }
    $null = Add-Portrait $slide $managerPerson 51 59 52 $managerAccent $PhotoIndex $MissingPhotos $Options.fontFamily
    $null = Add-TextBox $slide $ManagerName 119 57 523 25 18 "#172033" $true $ppAlignLeft $Options.fontFamily
    $roleText = if ($ManagerId -eq $DirectorId) { "Director" } else { $ManagerRole }
    $null = Add-TextBox $slide "$roleText, Employee ID $ManagerId" 119 84 523 20 9 "#5C6778" $false $ppAlignLeft $Options.fontFamily

    if ($Panels.Count -eq 0) {
        $null = Add-TextBox $slide "No employees assigned to this manager in this state" 36 205 888 70 17 "#6B7280" $false $ppAlignCenter $Options.fontFamily
        return
    }

    $contentLeft = 36.0
    $contentTop = 132.0
    $contentWidth = 888.0
    $contentHeight = 396.0
    $panelGap = 8.0
    $headerHeight = 27.0
    $maxPerRow = [int]$Options.maxPeoplePerRow
    $nameFontSize = [double]$Options.employeeNameFontSize

    $orderedPanels = @($Panels)
    if ($orderedPanels.Count -eq 3) {
        $largest = $orderedPanels | Sort-Object @{ Expression = { $_.TotalCount }; Descending = $true }, @{ Expression = { $_.Role } } | Select-Object -First 1
        $orderedPanels = @($orderedPanels | Where-Object { $_ -ne $largest } | Sort-Object Role) + @($largest)
    }

    $panelPositions = [System.Collections.Generic.List[object]]::new()
    switch ($orderedPanels.Count) {
        1 {
            $panelPositions.Add([pscustomobject]@{ Left = $contentLeft; Top = $contentTop; Width = $contentWidth; Height = $contentHeight })
        }
        2 {
            $width = ($contentWidth - $panelGap) / 2
            $panelPositions.Add([pscustomobject]@{ Left = $contentLeft; Top = $contentTop; Width = $width; Height = $contentHeight })
            $panelPositions.Add([pscustomobject]@{ Left = $contentLeft + $width + $panelGap; Top = $contentTop; Width = $width; Height = $contentHeight })
        }
        3 {
            $width = ($contentWidth - $panelGap) / 2
            $height = ($contentHeight - $panelGap) / 2
            $panelPositions.Add([pscustomobject]@{ Left = $contentLeft; Top = $contentTop; Width = $width; Height = $height })
            $panelPositions.Add([pscustomobject]@{ Left = $contentLeft + $width + $panelGap; Top = $contentTop; Width = $width; Height = $height })
            $panelPositions.Add([pscustomobject]@{ Left = $contentLeft; Top = $contentTop + $height + $panelGap; Width = $contentWidth; Height = $height })
        }
        default {
            $width = ($contentWidth - $panelGap) / 2
            $height = ($contentHeight - $panelGap) / 2
            for ($positionIndex = 0; $positionIndex -lt 4; $positionIndex++) {
                $rowIndex = [Math]::Floor($positionIndex / 2)
                $columnIndex = $positionIndex % 2
                $panelPositions.Add([pscustomobject]@{
                    Left = $contentLeft + $columnIndex * ($width + $panelGap)
                    Top = $contentTop + $rowIndex * ($height + $panelGap)
                    Width = $width
                    Height = $height
                })
            }
        }
    }

    for ($panelIndex = 0; $panelIndex -lt $orderedPanels.Count; $panelIndex++) {
        $panel = $orderedPanels[$panelIndex]
        $position = $panelPositions[$panelIndex]
        $left = [double]$position.Left
        $panelTop = [double]$position.Top
        $panelWidth = [double]$position.Width
        $panelHeight = [double]$position.Height
        $color = Get-RoleColor $RoleColors $panel.Role
        $null = Add-FilledShape $slide $msoShapeRoundedRectangle $left $panelTop $panelWidth $panelHeight "#FFFFFF" "#D9E0E8"
        $null = Add-FilledShape $slide $msoShapeRoundedRectangle $left $panelTop $panelWidth $headerHeight $color $color
        $null = Add-TextBox $slide $panel.Role ($left + 5) ($panelTop + 1) ($panelWidth - 10) ($headerHeight - 2) 11 (Get-TextColorHex $color) $true $ppAlignCenter $Options.fontFamily

        $innerLeft = $left + 6
        $innerTop = $panelTop + $headerHeight + 7
        $innerWidth = $panelWidth - 12
        $innerHeight = $panelHeight - $headerHeight - 13
        $maxRows = if ($panelHeight -gt 260) { [int]$Options.maxRowsPerCategory } else { [Math]::Min(2, [int]$Options.maxRowsPerCategory) }
        $columnWidth = $innerWidth / $maxPerRow
        $rowHeight = $innerHeight / $maxRows
        $photoSize = [Math]::Min(39, [Math]::Max(25, $columnWidth - 6))
        for ($index = 0; $index -lt $panel.Members.Count; $index++) {
            $row = [Math]::Floor($index / $maxPerRow)
            $column = $index % $maxPerRow
            $tileLeft = $innerLeft + $column * $columnWidth
            $tileTop = $innerTop + $row * $rowHeight
            $photoLeft = $tileLeft + ($columnWidth - $photoSize) / 2
            $member = $panel.Members[$index]
            $portrait = Add-Portrait $slide $member $photoLeft $tileTop $photoSize $color $PhotoIndex $MissingPhotos $Options.fontFamily
            $explanation = if ($member.PSObject.Properties.Name -contains "explanation") { [string]$member.explanation } else { "" }
            Add-ExplanationCue $slide $portrait $explanation
            $isVacant = $member.PSObject.Properties.Name -contains "isVacant" -and [bool]$member.isVacant
            $pcn = if ($member.PSObject.Properties.Name -contains "pcn") { [string]$member.pcn } else { "" }
            $label = if ($isVacant) { "Vacant`nPCN $pcn" } else { [string]$member.employeeName }
            if (-not $isVacant -and [bool]$Options.showEmployeeIds) { $label = "$label`n$($member.employeeId)" }
            if ($isVacant -or [bool]$Options.showEmployeeNames -or [bool]$Options.showEmployeeIds) {
                $labelShape = Add-TextBox $slide $label ($tileLeft + 1) ($tileTop + $photoSize + 2) ($columnWidth - 2) ([Math]::Min(28, $rowHeight - $photoSize - 3)) $nameFontSize "#303A4A" $false $ppAlignCenter $Options.fontFamily
                Add-ExplanationCue $slide $labelShape $explanation $false
            }
        }
    }
}

$settingsPath = [System.IO.Path]::GetFullPath($Settings)
$dataPath = [System.IO.Path]::GetFullPath($Data)
$settingsData = Get-Content -Raw -LiteralPath $settingsPath | ConvertFrom-Json
$orgData = Get-Content -Raw -LiteralPath $dataPath | ConvertFrom-Json
$settingsDirectory = Split-Path -Parent $settingsPath
$photoFolder = Resolve-ConfiguredPath ([string]$settingsData.photoFolder) $settingsDirectory
$outputPath = Resolve-ConfiguredPath ([string]$settingsData.outputPath) $settingsDirectory
$extensions = @($settingsData.options.photoExtensions | ForEach-Object { ([string]$_).ToLowerInvariant() })
$photoIndex = New-PhotoIndex $photoFolder $extensions
$missingPhotos = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$options = [pscustomobject]@{
    fontFamily = if ($settingsData.options.fontFamily) { [string]$settingsData.options.fontFamily } else { "Aptos" }
    maxPeoplePerRow = if ($settingsData.options.maxPeoplePerRow) { [int]$settingsData.options.maxPeoplePerRow } else { 5 }
    maxCategoriesPerSlide = if ($settingsData.options.maxCategoriesPerSlide) { [int]$settingsData.options.maxCategoriesPerSlide } else { 4 }
    maxRowsPerCategory = if ($settingsData.options.maxRowsPerCategory) { [int]$settingsData.options.maxRowsPerCategory } else { 3 }
    employeeNameFontSize = if ($settingsData.options.employeeNameFontSize) { [double]$settingsData.options.employeeNameFontSize } else { 7 }
    showEmployeeNames = if ($null -ne $settingsData.options.showEmployeeNames) { [bool]$settingsData.options.showEmployeeNames } else { $true }
    showEmployeeIds = if ($null -ne $settingsData.options.showEmployeeIds) { [bool]$settingsData.options.showEmployeeIds } else { $false }
}
if ($options.maxPeoplePerRow -lt 1 -or $options.maxPeoplePerRow -gt 5) { throw "maxPeoplePerRow must be between 1 and 5" }
if ($options.maxCategoriesPerSlide -lt 1 -or $options.maxCategoriesPerSlide -gt 4) { throw "maxCategoriesPerSlide must be between 1 and 4" }
if ($options.maxRowsPerCategory -lt 1 -or $options.maxRowsPerCategory -gt 4) { throw "maxRowsPerCategory must be between 1 and 4" }

$currentRecords = @($orgData.current)
$futureRecords = @($orgData.future)
$configuredDirector = if ($settingsData.options.directorId) { [string]$settingsData.options.directorId } else { "" }
$orderInfo = Get-ManagerOrder $currentRecords $futureRecords $configuredDirector
$powerPoint = $null
$presentation = $null
try {
    try { $powerPoint = New-Object -ComObject PowerPoint.Application }
    catch { throw "Microsoft PowerPoint desktop is required to build the deck." }
    $powerPoint.Visible = $msoTrue
    $presentation = $powerPoint.Presentations.Add()
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540

    foreach ($managerId in $orderInfo.Ordered) {
        $managerName = Get-ManagerName $managerId $currentRecords $futureRecords
        foreach ($stateLabel in @("current", "future")) {
            $records = if ($stateLabel -eq "current") { $currentRecords } else { $futureRecords }
            $members = @($records | Where-Object { $_.managerId -eq $managerId } | Sort-Object role, employeeName)
            $managerRecord = @($records | Where-Object { $_.employeeId -eq $managerId } | Select-Object -First 1)
            $managerRole = if ($managerRecord.Count -gt 0) { [string]$managerRecord[0].role } else { "Manager" }
            $currentCount = @($currentRecords | Where-Object { $_.managerId -eq $managerId }).Count
            $futureCount = @($futureRecords | Where-Object { $_.managerId -eq $managerId }).Count
            $pages = @(Get-RolePages $members $options.maxCategoriesPerSlide $options.maxPeoplePerRow $options.maxRowsPerCategory)
            $stateTitle = if ($stateLabel -eq "current") { [string]$settingsData.titles.current } else { [string]$settingsData.titles.future }
            for ($pageIndex = 0; $pageIndex -lt $pages.Count; $pageIndex++) {
                Add-TeamSlide $presentation $managerId $managerName $managerRole $members $stateLabel $stateTitle $currentCount $futureCount ($pageIndex + 1) $pages.Count $pages[$pageIndex].Panels $orderInfo.Director $orgData.roleColors $photoIndex $missingPhotos $options
            }
        }
    }
    $outputDirectory = Split-Path -Parent $outputPath
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    if ([bool]$settingsData.options.failOnMissingPhotos -and $missingPhotos.Count -gt 0) {
        throw "Missing photos for $($missingPhotos.Count) people. Set failOnMissingPhotos to false to use initials."
    }
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
    $presentation.SaveAs($outputPath, $ppSaveAsOpenXMLPresentation)
    $reportPath = Join-Path ([System.IO.Path]::GetDirectoryName($outputPath)) (([System.IO.Path]::GetFileNameWithoutExtension($outputPath)) + ".missing-photos.txt")
    if ($missingPhotos.Count -gt 0) {
        @("Missing photos ($($missingPhotos.Count))", "") + @($missingPhotos | Sort-Object) | Set-Content -LiteralPath $reportPath -Encoding UTF8
    } else { "No missing photos." | Set-Content -LiteralPath $reportPath -Encoding UTF8 }
    Write-Output "Created $outputPath"
    Write-Output "Slides: $($presentation.Slides.Count)"
    Write-Output "Missing-photo report: $reportPath"
} finally {
    if ($null -ne $presentation) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) | Out-Null
    }
    if ($null -ne $powerPoint) {
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
