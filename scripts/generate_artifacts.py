from __future__ import annotations

import json
import sys
from fractions import Fraction
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.lp2.graphical import format_fraction, solve_graphically
from src.lp2.plotting import plot_feasible_region
from src.lp2.problem import VARIANT_16
from src.lp2.simplex import SimplexResult, SimplexStep, solve_with_simplex, standardize


def fraction(value: float) -> str:
    candidate = Fraction(value).limit_denominator(1000)
    if abs(float(candidate) - value) <= 1e-8:
        return str(candidate)
    return f"{value:.6f}"


def write_vertices() -> pd.DataFrame:
    solution = solve_graphically(VARIANT_16)
    rows = []
    for idx, vertex in enumerate(solution.vertices, start=1):
        rows.append(
            {
                "point": f"A{idx}",
                "x1": fraction(vertex.x1),
                "x2": fraction(vertex.x2),
                "Z": fraction(vertex.value),
                "active_constraints": "; ".join(vertex.active_constraints),
            }
        )
    df = pd.DataFrame(rows)
    df.to_csv(ROOT / "tables" / "vertices.csv", index=False, encoding="utf-8-sig")
    return df


def step_rows(result: SimplexResult) -> pd.DataFrame:
    rows = []
    for step in [*result.phase1, *result.phase2]:
        rows.append(
            {
                "task": result.task,
                "phase": step.phase,
                "iteration": step.iteration,
                "basis": ", ".join(step.basis),
                "rhs": "; ".join(fraction(value) for value in step.rhs),
                "objective": fraction(step.objective_value),
                "entering": step.entering or "",
                "leaving": step.leaving or "",
                "reduced_costs": "; ".join(
                    f"{name}={fraction(value)}" for name, value in step.reduced_costs.items()
                ),
            }
        )
    return pd.DataFrame(rows)


def write_tableau(step: SimplexStep, filename: str) -> None:
    names = list(step.reduced_costs.keys()) + ["rhs"]
    rows = []
    for values in step.tableau:
        rows.append({name: fraction(value) for name, value in zip(names, values)})
    df = pd.DataFrame(rows, columns=names)
    df.insert(0, "basis", list(step.basis))
    df.to_csv(ROOT / "tables" / filename, index=False, encoding="utf-8-sig")


def write_simplex() -> tuple[SimplexResult, SimplexResult]:
    max_result = solve_with_simplex(VARIANT_16, "max")
    min_result = solve_with_simplex(VARIANT_16, "min")
    step_rows(max_result).to_csv(ROOT / "tables" / "simplex_max_steps.csv", index=False, encoding="utf-8-sig")
    step_rows(min_result).to_csv(ROOT / "tables" / "simplex_min_steps.csv", index=False, encoding="utf-8-sig")
    write_tableau(max_result.phase2[-1], "simplex_max_final_tableau.csv")
    write_tableau(min_result.phase2[-1], "simplex_min_final_tableau.csv")
    return max_result, min_result


def write_standard_form() -> None:
    form = standardize(VARIANT_16)
    rows = []
    for row_idx, constraint in enumerate(form.normalized_constraints, start=1):
        rows.append(
            {
                "row": row_idx,
                "normalized_constraint": constraint.label,
                "basis_start": form.names[form.basis[row_idx - 1]],
                "rhs": fraction(form.rhs[row_idx - 1]),
            }
        )
    pd.DataFrame(rows).to_csv(ROOT / "tables" / "standard_form.csv", index=False, encoding="utf-8-sig")


def write_summary(max_result: SimplexResult, min_result: SimplexResult) -> None:
    graph = solve_graphically(VARIANT_16)
    payload = {
        "variant": VARIANT_16.variant,
        "objective": VARIANT_16.objective_label,
        "constraints": [constraint.label for constraint in VARIANT_16.constraints],
        "vertices": [
            {
                "x1": vertex.x1,
                "x2": vertex.x2,
                "Z": vertex.value,
                "x1_fraction": format_fraction(vertex.x1),
                "x2_fraction": format_fraction(vertex.x2),
                "Z_fraction": format_fraction(vertex.value),
            }
            for vertex in graph.vertices
        ],
        "min": {
            "point": graph.minimum.point.tolist(),
            "value": graph.minimum.value,
            "point_fraction": [format_fraction(graph.minimum.x1), format_fraction(graph.minimum.x2)],
            "value_fraction": format_fraction(graph.minimum.value),
        },
        "max": {
            "points": [vertex.point.tolist() for vertex in graph.maximum_points],
            "value": graph.maximum.value,
            "value_fraction": format_fraction(graph.maximum.value),
            "simplex_point": list(max_result.solution),
            "alternate_optimum": max_result.alternate_optimum,
        },
        "simplex_min": {
            "point": list(min_result.solution),
            "value": min_result.original_objective_value,
        },
    }
    (ROOT / "tables" / "summary.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> None:
    (ROOT / "figures").mkdir(exist_ok=True)
    (ROOT / "tables").mkdir(exist_ok=True)
    graph = solve_graphically(VARIANT_16)
    plot_feasible_region(VARIANT_16, graph, str(ROOT / "figures" / "feasible_region.png"))
    write_vertices()
    write_standard_form()
    max_result, min_result = write_simplex()
    write_summary(max_result, min_result)
    print("Generated tables and figures for variant 16.")


if __name__ == "__main__":
    main()
