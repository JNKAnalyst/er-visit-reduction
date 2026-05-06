# ER Visit Reduction — HCE Case Study

> Automated Excel/VBA model for multi-scenario ER cost and savings analysis

## Overview

This project designs a fully automated Excel/VBA model for multi-scenario cost and savings analysis, eliminating manual recalculation and improving evaluation efficiency. Four intervention scenarios were modeled to identify the highest-value cost-reduction strategies for healthcare decision-makers.

**Key Result:** Combined intervention strategy projected a **31% reduction** in avoidable ER utilization across 4 modeled cost-reduction paths.

## Tools & Technologies

| Tool | Purpose |
|------|---------|
| Microsoft Excel | Data modeling and scenario analysis |
| VBA (Visual Basic for Applications) | Automation and macro development |
| Scenario Manager | Multi-path intervention comparison |

## Repository Structure

```
er-visit-reduction/
├── README.md
├── LICENSE
├── CHANGELOG.md
├── requirements.txt                       # Environment notes
├── docs/
│   └── methodology.md                     # Analysis methodology, formulas, limitations
├── model/
│   ├── ER_Visit_Reduction_Model.xlsx      # Generated workbook skeleton
│   └── vba/                               # Canonical VBA source (.bas modules)
│       ├── modConfig.bas
│       ├── modScenarios.bas
│       └── modDashboard.bas
├── templates/
│   ├── config_template.csv                # Editable scenario parameters
│   └── baseline_visits_template.csv       # Editable baseline utilization
└── scripts/
    ├── build_workbook.py                  # Regenerates the .xlsx from templates
    └── test_calculations.py               # Sanity check of the scenario math
```

The VBA modules are kept as plain `.bas` files under `model/vba/` so they can be reviewed and diffed in git. After generating the `.xlsx` skeleton, import the modules in Excel's VBA editor and save as `.xlsm` (see `model/vba/README.md`).

## Environment Setup

### Requirements
- Microsoft Excel 2016 or later (Microsoft 365 recommended)
- Macros must be enabled

### Setup Instructions

1. **Clone the repository**
   ```bash
   git clone https://github.com/JNKAnalyst/er-visit-reduction.git
   cd er-visit-reduction
   ```

2. **Generate the workbook skeleton**
   ```bash
   python scripts/build_workbook.py
   ```
   This produces `model/ER_Visit_Reduction_Model.xlsx` from the CSV templates.

3. **Import the VBA modules**
   - Open the generated workbook in Excel and press **Alt + F11**
   - **File → Import File...** for each `.bas` under `model/vba/`
   - Save as **Excel Macro-Enabled Workbook (`.xlsm`)**
   - When reopening, click **Enable Content** to allow macros

4. **Run the model**
   - Navigate to the **Dashboard** sheet
   - Edit values on the `Config` sheet (or paste from `templates/config_template.csv`)
   - Run `modDashboard.RunAnalysis` (assign it to a Form Control button on the Dashboard)
   - Run `modDashboard.RunSensitivity` for the parameter sweep

## Configuration

### Scenario Parameters (`Config` sheet)
| Parameter | Default | Description |
|-----------|---------|-------------|
| Baseline ER Visits | 1,200/month | Starting utilization count |
| Cost per Visit | $1,850 | Average ER visit cost |
| Intervention A Reduction | 17% | Telehealth triage impact |
| Intervention B Reduction | 24% | Care coordination impact |
| Intervention C Reduction | 12% | Patient education impact |

To modify parameters, update values directly in the `Config` sheet — all scenario sheets auto-recalculate.

## Deployment

This is a desktop Excel model — no server deployment required.

- **Share**: Distribute the `.xlsm` file directly
- **Version control**: Commit updated `.xlsm` files to this repo after changes
- **Compatibility**: Tested on Excel 2016, 2019, and Microsoft 365 (Windows & Mac)

> **Note for Mac users:** Some VBA features may behave differently on Excel for Mac. Microsoft 365 on Mac is recommended.

## Methodology

See [`docs/methodology.md`](docs/methodology.md) for inputs, formulas, the multiplicative combined-reduction calculation, and modeling limitations.

Quick summary:
- Baseline ER utilization modeled from HCE case study inputs (defaults shipped here are illustrative, not proprietary case data)
- Each intervention scenario independently modeled then combined multiplicatively
- Savings calculated as `(Baseline − Post-Intervention Visits) × Cost per Visit`
- Sensitivity sweep available via `modDashboard.RunSensitivity`

## Tests

```bash
python scripts/test_calculations.py
```

Verifies the scenario math used by `modScenarios.bas` matches the formulas documented in `docs/methodology.md`.

## Author

**Joash** | MS Business Analytics  
[GitHub](https://github.com/JNKAnalyst) | [Portfolio](https://jnkanalyst.github.io/portfolio/)
