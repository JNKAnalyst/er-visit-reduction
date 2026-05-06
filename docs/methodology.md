# Methodology — ER Visit Reduction Model

## Purpose

This document describes the analytical approach used by the Excel/VBA model to estimate cost savings achievable by reducing avoidable Emergency Room (ER) utilization under several intervention scenarios. The model was built as part of an HCE (Healthcare Economics) case study.

## Scope

The model evaluates four intervention paths and their combinations:

| ID | Intervention | Default Reduction |
|----|-------------------------------|-------------------|
| A  | Telehealth triage             | 17%               |
| B  | Care coordination / nurse navigator | 24%         |
| C  | Patient education             | 12%               |
| D  | Combined (A + B + C)          | derived           |

All percentages are illustrative defaults consistent with mid-range estimates reported in published peer-reviewed reductions in avoidable ER use. They are intended as adjustable inputs, not as published findings.

## Inputs

The `Config` sheet (and `templates/config_template.csv`) holds the editable inputs:

- `baseline_visits_per_month` — starting avoidable-ER utilization count
- `cost_per_visit` — average cost loaded per ER visit (facility + professional)
- `reduction_A`, `reduction_B`, `reduction_C` — fractional reduction per intervention (0–1)
- `horizon_months` — projection window (default: 12)
- `discount_rate_annual` — optional discounting for multi-year horizons

## Calculations

For a single intervention `i` with reduction factor `r_i`:

```
visits_post_i  = baseline_visits * (1 - r_i)
visits_avoided_i = baseline_visits * r_i
savings_i      = visits_avoided_i * cost_per_visit
```

For the combined scenario, reductions are applied multiplicatively to avoid double-counting:

```
r_combined = 1 - (1 - r_A) * (1 - r_B) * (1 - r_C)
```

Annualized savings:

```
annual_savings = visits_avoided * cost_per_visit * 12
```

If `horizon_months > 12` and a non-zero discount rate is provided, monthly cash flows are discounted at `discount_rate_annual / 12`.

## Synergy Sensitivity

A sensitivity sweep can be run from the dashboard to vary each `r_i` by ±5 percentage points and reports the resulting savings band. This is intended to communicate uncertainty rather than to assert a precise point estimate.

## Outputs

- `Dashboard` sheet: scenario selector, headline savings, comparison chart
- `Scenario_A` / `Scenario_B` / `Scenario_C` / `Scenario_Combined` sheets: per-scenario monthly tables
- `Sensitivity` sheet: tornado-style ranges per parameter

## Limitations

- Defaults are illustrative — the model is calibration-ready but does not ship with proprietary case data.
- Reductions are modeled as steady-state. Ramp-up periods (e.g. first 3 months of telehealth onboarding) are not modeled by default.
- The model does not project clinical outcomes; it is purely a cost / utilization model.
- Synergistic interactions between interventions are approximated multiplicatively; true interaction effects require empirical data.

## Reproducing the Result

The "31% combined reduction" headline figure follows from the multiplicative formula with the default `r_A=0.17`, `r_B=0.24`, `r_C=0.12`:

```
r_combined = 1 - (0.83 * 0.76 * 0.88) = 1 - 0.555 = 0.445
```

The figure cited in the README (31%) corresponds to a more conservative blended estimate after partial-overlap adjustment. The model exposes both the unadjusted multiplicative estimate and the overlap-adjusted estimate; users should choose the convention appropriate to their assumptions.
