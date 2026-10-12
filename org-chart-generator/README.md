# Organization chart PowerPoint generator

This tool creates a manager-by-manager PowerPoint from an Excel workbook:

1. Directors overview
2. Each director's current and future organization
3. Managers overview grouped by future-state director
4. Each manager's current and future organization

The deck begins with a Directors overview, followed by each director's Current and Future slides. A Managers overview then groups all managers under their future-state director, followed by each manager's Current and Future slides. Employees appear inside role-colored groups with their photos. Future slides show the headcount change from the same manager's current team. Photo filenames must contain the employee's four-character ID, such as `A123 Jane Smith.jpg` or `A123.jpg`.

## Excel structure

- `Employees`: Enter each person once with Employee ID, Employee Name, and the Manager selector. Mark managers with `☑ Yes`.
- `Current`: Select an employee name from the alphabetical dropdown; the Employee ID fills automatically. Then select the manager and role for the current organization. For an empty seat, leave Employee Name blank and enter its PCN.
- `Future`: Select an employee name from the alphabetical dropdown; the Employee ID fills automatically. Then select the manager and role for the future organization. For an empty seat, leave Employee Name blank and enter its PCN.
- `Role Settings`: Maintain role names and recolor the Slide Color cells. The generator reads the actual Excel fill color.

Employee and manager IDs must contain exactly four letters or numbers. Employee names are selected from an alphabetical dropdown on Current and Future, and the corresponding Employee ID is filled automatically.

`PCN` identifies a position and is required when no employee is assigned. `Explanation` is optional. When present, the generated slide gives that employee or vacant position a red circle outline. Hover over the circle or its label in PowerPoint to see the explanation as a ScreenTip.

## Setup

1. Enter your organization data in `Org Chart Data Template.xlsx`.
2. Copy `settings.example.json` to `settings.json`.
3. Set `photoFolder`. The example settings already point to the included workbook and create `Org Chart Output.pptx` in this folder.
4. Update the column mappings only if you rename the Current or Future input headers.
5. Run from PowerShell:

```powershell
.\run.ps1
```

You can also right-click `run.ps1` and choose **Run with PowerShell**. The scripts are compatible with Windows PowerShell 5.1 and PowerShell 7. If Windows blocks local scripts, open PowerShell in this folder and use `powershell.exe -ExecutionPolicy Bypass -File .\run.ps1`.

`Org Chart Example.xlsx` contains a populated sample with a director, three managers, multiple roles, and current-to-future reporting changes. To generate a deck from it, copy `settings.example.json` to `settings.json` and change `workbookPath` to `./Org Chart Example.xlsx`.

The default workflow uses PowerShell plus Microsoft Excel and PowerPoint desktop automation. Python, Node.js, Codex, and private packages are not required.

To use a differently named settings file:

```powershell
.\run.ps1 -Settings .\my-settings.json
```

Optional Python fallback for a machine without Microsoft Excel desktop:

```powershell
python -m pip install -r .\requirements.txt
.\run.ps1 -Extractor Python
```

The Python fallback only replaces Excel data extraction. PowerPoint desktop still creates the deck.

Relative paths in the settings file resolve from the folder containing that settings file.

## Important settings

- `includeManagersInEmployeeCount`: Adds the manager to the displayed team count when the manager also appears as an employee in the same sheet.
- `directorIds`: Optional array of four-character director IDs, such as `["D001", "D002"]`. Leave it empty to infer every top-level director from the Current and Future reporting relationships. The legacy single `directorId` setting is still accepted.
- `showEmployeeNames`: Shows employee names beneath photos.
- `showEmployeeIds`: Shows IDs with names.
- `maxPeoplePerRow`: Maximum employee photos per category row. Limited to 5.
- `maxCategoriesPerSlide`: Maximum role categories per slide. Limited to 4.
- `maxRowsPerCategory`: Number of employee rows available per category before a continuation slide is added.
- `employeeNameFontSize`: Font size used beneath employee photos.
- `failOnMissingPhotos`: Stops instead of using an initials placeholder.

Large teams and teams with more than four categories automatically continue onto additional slides. Current slides remain directly followed by their current continuations, then the manager's future slides and future continuations. Managers are grouped using their future-state director, so a manager moving between directors appears in the destination director's section with a move note on the Managers overview.

Category layout adapts automatically:

- One category uses the full grouping area.
- Two categories split left and right.
- Three categories place two panels across the top and the largest category full-width below.
- Four categories use an equal 2-by-2 grid.

The generator writes a missing-photo report next to the PowerPoint using the suffix `.missing-photos.txt`.

It also creates a separate HR handoff workbook at `hrOutputPath`. `Director Summary` and `Manager Summary` list positions, employees, and vacancies gained or lost. `Employee Changes` lists current and future director assignments as well as manager transfers, director transfers, additions, removals, staffing changes, vacancy changes, and role changes by comparing Current with Future. If `hrOutputPath` is omitted, the file is created beside the PowerPoint with ` HR Handoff.xlsx` appended to the PowerPoint name.
