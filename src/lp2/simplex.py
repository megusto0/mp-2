from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from .problem import Constraint, LinearProgram

TOLERANCE = 1e-9


@dataclass(frozen=True)
class StandardForm:
    names: tuple[str, ...]
    matrix: np.ndarray
    rhs: np.ndarray
    basis: tuple[int, ...]
    artificial: tuple[int, ...]
    normalized_constraints: tuple[Constraint, ...]


@dataclass(frozen=True)
class SimplexStep:
    iteration: int
    phase: str
    basis: tuple[str, ...]
    rhs: tuple[float, ...]
    objective_value: float
    reduced_costs: dict[str, float]
    entering: str | None
    leaving: str | None
    tableau: np.ndarray


@dataclass(frozen=True)
class SimplexResult:
    task: str
    transformed_objective: tuple[float, float]
    original_objective_value: float
    transformed_objective_value: float
    solution: tuple[float, float]
    phase1: tuple[SimplexStep, ...]
    phase2: tuple[SimplexStep, ...]
    final_basis: tuple[str, ...]
    alternate_optimum: bool


def standardize(lp: LinearProgram) -> StandardForm:
    names: list[str] = ["x1", "x2"]
    rows: list[list[float]] = []
    rhs: list[float] = []
    basis: list[int] = []
    artificial: list[int] = []
    normalized: list[Constraint] = []

    for idx, original in enumerate(lp.constraints, start=1):
        constraint = original.normalized()
        normalized.append(constraint)
        row = [0.0] * len(names)
        row[0], row[1] = constraint.coefficients

        if constraint.sense == "<=":
            names.append(f"s{idx}")
            row.append(1.0)
            for old in rows:
                old.append(0.0)
            basis.append(len(names) - 1)
        elif constraint.sense == ">=":
            names.append(f"e{idx}")
            row.append(-1.0)
            for old in rows:
                old.append(0.0)
            names.append(f"a{idx}")
            row.append(1.0)
            for old in rows:
                old.append(0.0)
            basis.append(len(names) - 1)
            artificial.append(len(names) - 1)
        elif constraint.sense == "=":
            names.append(f"a{idx}")
            row.append(1.0)
            for old in rows:
                old.append(0.0)
            basis.append(len(names) - 1)
            artificial.append(len(names) - 1)
        else:
            raise ValueError(f"Unsupported constraint sense: {constraint.sense}")

        rows.append(row)
        rhs.append(float(constraint.rhs))

    return StandardForm(
        names=tuple(names),
        matrix=np.array(rows, dtype=float),
        rhs=np.array(rhs, dtype=float),
        basis=tuple(basis),
        artificial=tuple(artificial),
        normalized_constraints=tuple(normalized),
    )


def _canonical(matrix: np.ndarray, rhs: np.ndarray, basis: list[int], objective: np.ndarray):
    basis_matrix = matrix[:, basis]
    inverse = np.linalg.inv(basis_matrix)
    tableau = inverse @ matrix
    right_side = inverse @ rhs
    basis_costs = objective[basis]
    reduced_costs = objective - basis_costs @ tableau
    objective_value = float(basis_costs @ right_side)
    return tableau, right_side, reduced_costs, objective_value


def _run_simplex(
    matrix: np.ndarray,
    rhs: np.ndarray,
    basis: list[int],
    objective: np.ndarray,
    names: tuple[str, ...],
    phase: str,
) -> tuple[tuple[SimplexStep, ...], list[int]]:
    steps: list[SimplexStep] = []

    for iteration in range(50):
        tableau, right_side, reduced_costs, objective_value = _canonical(matrix, rhs, basis, objective)
        nonbasis = [idx for idx in range(len(names)) if idx not in basis]
        candidates = [idx for idx in nonbasis if reduced_costs[idx] > TOLERANCE]

        entering = max(candidates, key=lambda idx: (reduced_costs[idx], -idx)) if candidates else None
        leaving_row: int | None = None

        if entering is not None:
            ratios = [
                (right_side[row] / tableau[row, entering], row)
                for row in range(matrix.shape[0])
                if tableau[row, entering] > TOLERANCE
            ]
            if not ratios:
                raise ValueError(f"The problem is unbounded in {phase}.")
            _, leaving_row = min(ratios, key=lambda item: (item[0], item[1]))

        steps.append(
            SimplexStep(
                iteration=iteration,
                phase=phase,
                basis=tuple(names[idx] for idx in basis),
                rhs=tuple(float(value) for value in right_side),
                objective_value=float(objective_value),
                reduced_costs={names[idx]: float(reduced_costs[idx]) for idx in range(len(names))},
                entering=names[entering] if entering is not None else None,
                leaving=names[basis[leaving_row]] if leaving_row is not None else None,
                tableau=np.column_stack([tableau, right_side]),
            )
        )

        if entering is None:
            return tuple(steps), basis
        basis[leaving_row] = entering

    raise RuntimeError(f"Simplex did not converge in {phase}.")


def _remove_artificial(form: StandardForm, basis: list[int]) -> tuple[tuple[str, ...], np.ndarray, list[int]]:
    keep = [idx for idx in range(len(form.names)) if idx not in form.artificial]
    old_to_new = {old: new for new, old in enumerate(keep)}
    if any(idx in form.artificial for idx in basis):
        raise ValueError("Artificial variable remained in the basis after phase I.")
    names = tuple(form.names[idx] for idx in keep)
    matrix = form.matrix[:, keep]
    mapped_basis = [old_to_new[idx] for idx in basis]
    return names, matrix, mapped_basis


def solve_with_simplex(lp: LinearProgram, task: str) -> SimplexResult:
    if task not in {"max", "min"}:
        raise ValueError("task must be either 'max' or 'min'.")

    form = standardize(lp)
    phase1_objective = np.zeros(len(form.names), dtype=float)
    for idx in form.artificial:
        phase1_objective[idx] = -1.0

    phase1_steps, phase1_basis = _run_simplex(
        form.matrix,
        form.rhs,
        list(form.basis),
        phase1_objective,
        form.names,
        "phase I",
    )
    if abs(phase1_steps[-1].objective_value) > 1e-7:
        raise ValueError("No feasible solution found in phase I.")

    names, matrix, phase2_basis = _remove_artificial(form, phase1_basis)
    objective = np.zeros(len(names), dtype=float)
    objective[:2] = lp.c if task == "max" else -lp.c

    phase2_steps, final_basis = _run_simplex(
        matrix,
        form.rhs,
        phase2_basis,
        objective,
        names,
        "phase II",
    )

    final_step = phase2_steps[-1]
    solution_values = {name: 0.0 for name in names}
    for name, value in zip(final_step.basis, final_step.rhs):
        solution_values[name] = value
    solution = (solution_values["x1"], solution_values["x2"])
    transformed_value = final_step.objective_value
    original_value = transformed_value if task == "max" else -transformed_value

    alternate = any(
        abs(cost) <= 1e-7 and name not in final_step.basis
        for name, cost in final_step.reduced_costs.items()
    )

    return SimplexResult(
        task=task,
        transformed_objective=tuple(float(value) for value in objective[:2]),
        original_objective_value=float(original_value),
        transformed_objective_value=float(transformed_value),
        solution=tuple(float(value) for value in solution),
        phase1=phase1_steps,
        phase2=phase2_steps,
        final_basis=tuple(names[idx] for idx in final_basis),
        alternate_optimum=alternate,
    )
