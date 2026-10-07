# Organization chart PowerPoint generator

This tool creates a manager-by-manager PowerPoint from an Excel workbook:

1. Director, current organization
2. Director, future organization
3. First manager, current organization
4. First manager, future organization
5. Remaining managers in reporting order, with current immediately followed by future

Each slide shows one manager and that manager's employees. Employees appear inside role-colored groups with their photos. Future slides show the headcount change from the same manager's current team. Photo filenames must contain the employee's four-character ID, such as `A123 Jane Smith.jpg` or `A123.jpg`.

## Excel structure

The current and future sheets must each have one header row and one row per employee. The default expected headers are:

- `Employee ID`
- `Employee Name`
- `Manager ID`
- `Manager Name`
- `Role Type`

Employee and manager IDs must contain exactly four characters. Header names can be changed in `settings.json`.

## Setup

1. Copy `settings.example.json` to `settings.json`.
2. Set `workbookPath`, `photoFolder`, `outputPath`, and the two sheet names.
3. Update the column mappings and role colors if needed.
4. Run from PowerShell:

```powershell
.\run.ps1
```

To use a differently named settings file:

```powershell
.\run.ps1 -Settings .\my-settings.json
```

Relative paths in the settings file resolve from the folder containing that settings file.

## Important settings

- `includeManagersInEmployeeCount`: Adds the manager to the displayed team count when the manager also appears as an employee in the same sheet.
- `directorId`: Optional four-character ID for the director who should appear first. Leave blank to infer the top manager from the reporting relationships.
- `showEmployeeNames`: Shows employee names beneath photos.
- `showEmployeeIds`: Shows IDs with names.
- `maximumEmployeesPerManager`: Stops generation if a team exceeds the defined capacity.
- `failOnMissingPhotos`: Stops instead of using an initials placeholder.
- `minimumEmployeeTileWidth`: Controls the smallest permitted employee tile before the slide is considered too dense.

The generator writes a missing-photo report next to the PowerPoint using the suffix `.missing-photos.txt`.
