"""Mirror of the VBA scenario math, used as a sanity check.

Run with ``python scripts/test_calculations.py``. Exits non-zero on failure.
The VBA module ``modScenarios.bas`` is the production code; this script
verifies that the formulas documented in ``docs/methodology.md`` agree with
what the VBA implements.
"""
from __future__ import annotations

import sys


def calc_single(baseline: float, cost: float, r: float) -> dict:
    visits_avoided = baseline * r
    monthly = visits_avoided * cost
    return {
        "reduction": r,
        "visits_avoided": visits_avoided,
        "monthly_savings": monthly,
        "annual_savings": monthly * 12,
    }


def calc_combined(baseline: float, cost: float, ra: float, rb: float, rc: float) -> dict:
    r = 1 - (1 - ra) * (1 - rb) * (1 - rc)
    return calc_single(baseline, cost, r)


def _approx(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol * max(1.0, abs(b))


def main() -> int:
    failures: list[str] = []

    s = calc_single(1200, 1850, 0.17)
    if not _approx(s["visits_avoided"], 204.0):
        failures.append(f"single A visits_avoided={s['visits_avoided']}")
    if not _approx(s["annual_savings"], 204.0 * 1850 * 12):
        failures.append(f"single A annual_savings={s['annual_savings']}")

    c = calc_combined(1200, 1850, 0.17, 0.24, 0.12)
    expected_r = 1 - 0.83 * 0.76 * 0.88
    if not _approx(c["reduction"], expected_r):
        failures.append(f"combined reduction={c['reduction']} expected={expected_r}")

    z = calc_single(1000, 1000, 0.0)
    if z["annual_savings"] != 0.0:
        failures.append(f"zero reduction should produce zero savings, got {z['annual_savings']}")

    if failures:
        print("FAIL")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("OK: all scenario checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
