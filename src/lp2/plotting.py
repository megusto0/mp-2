from __future__ import annotations

import matplotlib.pyplot as plt
import numpy as np

from .graphical import GraphicalSolution
from .problem import LinearProgram


def _line_y(coefficients: tuple[float, float], rhs: float, x_values: np.ndarray) -> np.ndarray:
    a1, a2 = coefficients
    if abs(a2) < 1e-12:
        return np.full_like(x_values, np.nan)
    return (rhs - a1 * x_values) / a2


def plot_feasible_region(lp: LinearProgram, solution: GraphicalSolution, output_path: str) -> None:
    vertices = np.array([[vertex.x1, vertex.x2] for vertex in solution.vertices], dtype=float)
    center = vertices.mean(axis=0)
    order = np.argsort(np.arctan2(vertices[:, 1] - center[1], vertices[:, 0] - center[0]))
    polygon = vertices[order]

    fig, ax = plt.subplots(figsize=(7.2, 5.0), dpi=160)
    ax.fill(polygon[:, 0], polygon[:, 1], color="#cfe8ff", alpha=0.7, label="Допустимая область")
    ax.plot(np.r_[polygon[:, 0], polygon[0, 0]], np.r_[polygon[:, 1], polygon[0, 1]], color="#1f77b4")

    x_values = np.linspace(0, 3.0, 300)
    colors = ["#d62728", "#2ca02c", "#9467bd"]
    for constraint, color in zip(lp.constraints, colors):
        y_values = _line_y(constraint.coefficients, constraint.rhs, x_values)
        ax.plot(x_values, y_values, color=color, linewidth=1.3, label=constraint.label)

    min_point = np.array(solution.minimum.point)
    ax.scatter([min_point[0]], [min_point[1]], color="#111111", marker="o", zorder=5, label="min Z")
    for max_vertex in solution.maximum_points:
        ax.scatter([max_vertex.x1], [max_vertex.x2], color="#ff7f0e", marker="s", zorder=5)
    if len(solution.maximum_points) > 1:
        max_points = np.array([[v.x1, v.x2] for v in solution.maximum_points])
        ax.plot(max_points[:, 0], max_points[:, 1], color="#ff7f0e", linewidth=3.0, label="max Z")
    else:
        ax.scatter([], [], color="#ff7f0e", marker="s", label="max Z")

    for vertex in solution.vertices:
        ax.annotate(
            f"({vertex.x1:.2f}; {vertex.x2:.2f})",
            (vertex.x1, vertex.x2),
            xytext=(5, 5),
            textcoords="offset points",
            fontsize=8,
        )

    ax.set_xlim(0, 2.4)
    ax.set_ylim(2.0, 3.2)
    ax.set_xlabel("x1")
    ax.set_ylabel("x2")
    ax.grid(alpha=0.25)
    ax.legend(loc="upper left", fontsize=8, framealpha=0.9)
    fig.tight_layout()
    fig.savefig(output_path)
    plt.close(fig)
