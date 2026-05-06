# VBA Source Modules

These `.bas` files are the canonical VBA source for the model. They live outside the workbook so they can be diffed and reviewed in git.

## Importing into the workbook

1. Run `python ../../scripts/build_workbook.py` from the repo root to generate `model/ER_Visit_Reduction_Model.xlsx`.
2. Open the workbook in Excel and press **Alt + F11** to open the VBA editor.
3. **File → Import File...** and import each `.bas` module from this directory:
   - `modConfig.bas`
   - `modScenarios.bas`
   - `modDashboard.bas`
4. Save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)** at `model/ER_Visit_Reduction_Model.xlsm`.
5. On the `Dashboard` sheet, add a Form Control button and assign it to `modDashboard.RunAnalysis`.

## Module overview

| Module | Responsibility |
|--------|----------------|
| `modConfig` | Reads typed values from the `Config` sheet by key |
| `modScenarios` | Single and combined scenario calculations |
| `modDashboard` | `RunAnalysis` and `RunSensitivity` entry points |
