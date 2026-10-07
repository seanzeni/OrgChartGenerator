# Organization chart PowerPoint generator

This tool creates a manager-by-manager PowerPoint from an Excel workbook:

1. Director, current organization
2. Director, future organization
3. First manager, current organization
4. First manager, future organization
5. Remaining managers in reporting order, with current immediately followed by future

Each slide shows one manager and that manager's employees. Employees appear inside role-colored groups with their photos. Future slides show the headcount change from the same manager's current team. Photo filenames must contain the employee's four-character ID, such as `A123 Jane Smith.jpg` or `A123.jpg`.

## Excel structure

- `Employees`: Enter each person once with Employee ID, Employee Name, and the Manager selector. Mark managers with `☑ Yes`.
- `Current`: Select an employee, manager, and role for the current organization.
- `Future`: Select an employee, manager, and role for the future organization.
- `Role Settings`: Maintain role names and recolor the Slide Color cells. The generator reads the actual Excel fill color.

Employee and manager IDs must contain exactly four letters or numbers. Names are looked up from the Employees sheet, so they do not need to be typed again on Current or Future.

## Setup

1. Enter your organization data in `Org Chart Data Template.xlsx`.
2. Copy `settings.example.json` to `settings.json`.
3. Set `photoFolder`. The example settings already point to the included workbook and create `Org Chart Output.pptx` in this folder.
4. Update the column mappings only if you rename the Current or Future input headers.
5. Run from PowerShell:

```powershell
.\run.ps1
```

The default workflow uses PowerShell and Microsoft Excel desktop to read the workbook. Python is not required.

To use a differently named settings file:

```powershell
.\run.ps1 -Settings .\my-settings.json
```

Optional Python fallback for a machine without Microsoft Excel desktop:

```powershell
.\run.ps1 -Extractor Python
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
