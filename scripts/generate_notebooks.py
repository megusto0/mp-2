from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NOTEBOOK_DIR = ROOT / "notebooks"
REPO = "megusto0/mp-2"
BRANCH = "main"


def md(source: str) -> dict:
    return {"cell_type": "markdown", "metadata": {}, "source": source.strip().splitlines(True)}


def code(source: str) -> dict:
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": source.strip().splitlines(True),
    }


COMMON_SETUP = r"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

c = np.array([-3.0, 6.0])
constraints = [
    (np.array([5.0, -2.0]), "<=", 4.0, "5*x1 - 2*x2 <= 4"),
    (np.array([1.0, -2.0]), ">=", -4.0, "x1 - 2*x2 >= -4"),
    (np.array([1.0, 1.0]), ">=", 4.0, "x1 + x2 >= 4"),
]

def objective(point):
    return float(c @ point)

def feasible(point, tol=1e-9):
    if point[0] < -tol or point[1] < -tol:
        return False
    for a, sense, b, _ in constraints:
        left = float(a @ point)
        if sense == "<=" and left > b + tol:
            return False
        if sense == ">=" and left < b - tol:
            return False
    return True

def active(point, tol=1e-7):
    names = []
    if abs(point[0]) <= tol:
        names.append("x1 = 0")
    if abs(point[1]) <= tol:
        names.append("x2 = 0")
    for a, _, b, label in constraints:
        if abs(float(a @ point) - b) <= tol:
            names.append(label)
    return "; ".join(names)
"""


GRAPHICAL_CODE = COMMON_SETUP + r"""
from itertools import combinations

boundaries = constraints + [
    (np.array([1.0, 0.0]), ">=", 0.0, "x1 >= 0"),
    (np.array([0.0, 1.0]), ">=", 0.0, "x2 >= 0"),
]

points = []
for first, second in combinations(boundaries, 2):
    a1, _, b1, _ = first
    a2, _, b2, _ = second
    matrix = np.vstack([a1, a2])
    if abs(np.linalg.det(matrix)) < 1e-9:
        continue
    point = np.linalg.solve(matrix, np.array([b1, b2]))
    if feasible(point):
        if not any(np.linalg.norm(point - old, ord=np.inf) < 1e-7 for old in points):
            points.append(point)

rows = []
for index, point in enumerate(sorted(points, key=lambda p: (p[0], p[1])), start=1):
    rows.append({
        "point": f"A{index}",
        "x1": point[0],
        "x2": point[1],
        "Z": objective(point),
        "active constraints": active(point),
    })

vertices = pd.DataFrame(rows)
vertices
"""


PLOT_CODE = r"""
polygon = vertices[["x1", "x2"]].to_numpy()
center = polygon.mean(axis=0)
order = np.argsort(np.arctan2(polygon[:, 1] - center[1], polygon[:, 0] - center[0]))
polygon = polygon[order]

x_values = np.linspace(0, 3, 300)
fig, ax = plt.subplots(figsize=(7, 5), dpi=130)
ax.fill(polygon[:, 0], polygon[:, 1], alpha=0.35, color="#8ecae6", label="feasible region")
ax.plot(np.r_[polygon[:, 0], polygon[0, 0]], np.r_[polygon[:, 1], polygon[0, 1]], color="#126782")

for a, _, b, label in constraints:
    if abs(a[1]) > 1e-12:
        ax.plot(x_values, (b - a[0] * x_values) / a[1], label=label)

max_value = vertices["Z"].max()
min_value = vertices["Z"].min()
max_rows = vertices[abs(vertices["Z"] - max_value) < 1e-7]
min_row = vertices.loc[vertices["Z"].idxmin()]
ax.scatter(max_rows["x1"], max_rows["x2"], marker="s", color="#fb8500", label="max Z")
ax.scatter([min_row["x1"]], [min_row["x2"]], marker="o", color="#111111", label="min Z")

ax.set_xlim(0, 2.4)
ax.set_ylim(2.0, 3.2)
ax.set_xlabel("x1")
ax.set_ylabel("x2")
ax.grid(alpha=0.25)
ax.legend(fontsize=8)
plt.show()
"""


SIMPLEX_CODE = r"""
import numpy as np
import pandas as pd

raw_constraints = [
    (np.array([5.0, -2.0]), "<=", 4.0),
    (np.array([1.0, -2.0]), ">=", -4.0),
    (np.array([1.0, 1.0]), ">=", 4.0),
]

def normalize(a, sense, b):
    if b >= 0:
        return a.astype(float), sense, float(b)
    return -a.astype(float), {"<=": ">=", ">=": "<=", "=": "="}[sense], float(-b)

def standardize(raw):
    names = ["x1", "x2"]
    rows, rhs, basis, artificial = [], [], [], []
    for idx, item in enumerate(raw, start=1):
        a, sense, b = normalize(*item)
        row = [0.0] * len(names)
        row[0], row[1] = a
        if sense == "<=":
            names.append(f"s{idx}")
            row.append(1.0)
            for old in rows:
                old.append(0.0)
            basis.append(len(names) - 1)
        elif sense == ">=":
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
        rows.append(row)
        rhs.append(b)
    return names, np.array(rows), np.array(rhs), basis, artificial

def canonical(A, b, basis, objective):
    B = A[:, basis]
    inv = np.linalg.inv(B)
    table = inv @ A
    right = inv @ b
    reduced = objective - objective[basis] @ table
    value = float(objective[basis] @ right)
    return table, right, reduced, value

def simplex(A, b, basis, objective, names, phase):
    steps = []
    basis = list(basis)
    for iteration in range(30):
        table, right, reduced, value = canonical(A, b, basis, objective)
        nonbasis = [idx for idx in range(len(names)) if idx not in basis]
        candidates = [idx for idx in nonbasis if reduced[idx] > 1e-9]
        entering = max(candidates, key=lambda idx: (reduced[idx], -idx)) if candidates else None
        leaving_row = None
        if entering is not None:
            ratios = [(right[row] / table[row, entering], row) for row in range(A.shape[0]) if table[row, entering] > 1e-9]
            leaving_row = min(ratios)[1]
        steps.append({
            "phase": phase,
            "iteration": iteration,
            "basis": ", ".join(names[idx] for idx in basis),
            "rhs": "; ".join(f"{v:.6g}" for v in right),
            "objective": value,
            "entering": names[entering] if entering is not None else "",
            "leaving": names[basis[leaving_row]] if leaving_row is not None else "",
        })
        if entering is None:
            return pd.DataFrame(steps), basis, table, right, reduced, value
        basis[leaving_row] = entering
    raise RuntimeError("Simplex did not converge")

names, A, b, basis, artificial = standardize(raw_constraints)
phase1_objective = np.zeros(len(names))
for idx in artificial:
    phase1_objective[idx] = -1.0

phase1, phase1_basis, *_ = simplex(A, b, basis, phase1_objective, names, "phase I")
keep = [idx for idx in range(len(names)) if idx not in artificial]
new_index = {old: new for new, old in enumerate(keep)}
names2 = [names[idx] for idx in keep]
A2 = A[:, keep]
basis2 = [new_index[idx] for idx in phase1_basis]
"""


def simplex_notebook(task: str) -> list[dict]:
    objective = "np.array([-3.0, 6.0, 0.0, 0.0, 0.0])" if task == "max" else "np.array([3.0, -6.0, 0.0, 0.0, 0.0])"
    label = "максимума Z" if task == "max" else "минимума Z через максимум -Z"
    final_text = "В последней строке видно, что для максимума есть нулевая оценка у небазисной переменной. Это означает не единственную точку максимума." if task == "max" else "Вторая фаза сразу останавливается: точка, найденная после фазы I, уже дает минимум исходной функции."
    return [
        md(f"[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/{REPO}/blob/{BRANCH}/notebooks/0{'2' if task == 'max' else '3'}_simplex_{task}.ipynb)\n\n# Симплекс-метод для {label}\n\nНоутбук показывает двухфазный симплекс-метод без клонирования репозитория."),
        code(SIMPLEX_CODE),
        md("## Фаза I\n\nИскусственная переменная нужна только для ограничения `x1 + x2 >= 4`. Фаза I ищет начальную допустимую базисную точку."),
        code("phase1"),
        md("## Фаза II\n\nТеперь целевая функция заменяется на нужную для текущей задачи."),
        code(f"objective2 = {objective}\nphase2, final_basis, table, right, reduced, value = simplex(A2, b, basis2, objective2, names2, 'phase II')\nphase2"),
        code("solution = dict.fromkeys(names2, 0.0)\nfor name, val in zip([names2[idx] for idx in final_basis], right):\n    solution[name] = val\nx = np.array([solution['x1'], solution['x2']])\nZ = -3*x[0] + 6*x[1]\nprint('x* =', x)\nprint('Z =', Z)\nprint('transformed objective =', value)\nprint('reduced costs =', dict(zip(names2, reduced)))"),
        md(final_text),
    ]


def write_notebook(filename: str, cells: list[dict]) -> None:
    notebook = {
        "cells": cells,
        "metadata": {
            "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
            "language_info": {"name": "python", "pygments_lexer": "ipython3"},
            "colab": {"provenance": []},
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }
    (NOTEBOOK_DIR / filename).write_text(json.dumps(notebook, ensure_ascii=False, indent=1), encoding="utf-8")


def main() -> None:
    NOTEBOOK_DIR.mkdir(exist_ok=True)
    write_notebook(
        "01_graphical_method.ipynb",
        [
            md(f"[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/{REPO}/blob/{BRANCH}/notebooks/01_graphical_method.ipynb)\n\n# Графический метод\n\nНоутбук строит допустимую область, проверяет вершины и показывает, где достигаются минимум и максимум."),
            code(GRAPHICAL_CODE),
            code("print('minimum')\ndisplay(vertices.loc[[vertices['Z'].idxmin()]])\nprint('maximum')\ndisplay(vertices[abs(vertices['Z'] - vertices['Z'].max()) < 1e-7])"),
            code(PLOT_CODE),
        ],
    )
    write_notebook("02_simplex_max.ipynb", simplex_notebook("max"))
    write_notebook("03_simplex_min.ipynb", simplex_notebook("min"))
    write_notebook(
        "04_summary.ipynb",
        [
            md(f"[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/{REPO}/blob/{BRANCH}/notebooks/04_summary.ipynb)\n\n# Сводная проверка\n\nКороткая проверка, что графический метод и симплекс-метод дают одинаковые ответы."),
            code(GRAPHICAL_CODE),
            code("min_row = vertices.loc[vertices['Z'].idxmin()]\nmax_rows = vertices[abs(vertices['Z'] - vertices['Z'].max()) < 1e-7]\nsummary = pd.DataFrame([\n    {'task': 'minimum', 'point': f\"({min_row.x1:.6g}; {min_row.x2:.6g})\", 'Z': min_row.Z},\n    {'task': 'maximum edge start', 'point': f\"({max_rows.iloc[0].x1:.6g}; {max_rows.iloc[0].x2:.6g})\", 'Z': max_rows.iloc[0].Z},\n    {'task': 'maximum edge end', 'point': f\"({max_rows.iloc[-1].x1:.6g}; {max_rows.iloc[-1].x2:.6g})\", 'Z': max_rows.iloc[-1].Z},\n])\nsummary"),
            md("Для максимума получается не одна точка, а отрезок. Это нормально: линии уровня целевой функции параллельны одной стороне допустимого треугольника."),
        ],
    )
    print("Generated notebooks.")


if __name__ == "__main__":
    main()
