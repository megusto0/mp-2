from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations

import numpy as np

from .problem import Constraint, LinearProgram

TOLERANCE = 1e-9


@dataclass(frozen=True)
class Vertex:
    x1: float
    x2: float
    value: float
    active_constraints: tuple[str, ...]

    @property
    def point(self) -> np.ndarray:
        return np.array([self.x1, self.x2], dtype=float)


@dataclass(frozen=True)
class GraphicalSolution:
    vertices: tuple[Vertex, ...]
    minimum: Vertex
    maximum: Vertex
    maximum_points: tuple[Vertex, ...]
    minimum_points: tuple[Vertex, ...]


def objective_value(lp: LinearProgram, point: np.ndarray) -> float:
    return float(lp.c @ point)


def is_feasible(lp: LinearProgram, point: np.ndarray, tolerance: float = TOLERANCE) -> bool:
    for constraint in lp.all_constraints():
        left = float(np.array(constraint.coefficients, dtype=float) @ point)
        if constraint.sense == "<=" and left > constraint.rhs + tolerance:
            return False
        if constraint.sense == ">=" and left < constraint.rhs - tolerance:
            return False
        if constraint.sense == "=" and abs(left - constraint.rhs) > tolerance:
            return False
    return True


def active_constraints(lp: LinearProgram, point: np.ndarray, tolerance: float = 1e-7) -> tuple[str, ...]:
    active: list[str] = []
    for constraint in lp.all_constraints():
        left = float(np.array(constraint.coefficients, dtype=float) @ point)
        if abs(left - constraint.rhs) <= tolerance:
            active.append(constraint.label)
    return tuple(active)


def enumerate_vertices(lp: LinearProgram) -> tuple[Vertex, ...]:
    boundaries: list[Constraint] = list(lp.all_constraints())
    points: list[np.ndarray] = []

    for first, second in combinations(boundaries, 2):
        a1, b1 = first.boundary()
        a2, b2 = second.boundary()
        matrix = np.vstack([a1, a2])
        if abs(np.linalg.det(matrix)) <= TOLERANCE:
            continue
        point = np.linalg.solve(matrix, np.array([b1, b2], dtype=float))
        if is_feasible(lp, point):
            points.append(point)

    unique: list[np.ndarray] = []
    for point in points:
        if not any(np.linalg.norm(point - existing, ord=np.inf) <= 1e-7 for existing in unique):
            unique.append(point)

    vertices = [
        Vertex(
            x1=float(point[0]),
            x2=float(point[1]),
            value=objective_value(lp, point),
            active_constraints=active_constraints(lp, point),
        )
        for point in unique
    ]
    vertices.sort(key=lambda item: (item.x1, item.x2))
    return tuple(vertices)


def solve_graphically(lp: LinearProgram) -> GraphicalSolution:
    vertices = enumerate_vertices(lp)
    if not vertices:
        raise ValueError("No feasible vertices were found.")

    min_value = min(vertex.value for vertex in vertices)
    max_value = max(vertex.value for vertex in vertices)
    minimum_points = tuple(v for v in vertices if abs(v.value - min_value) <= 1e-7)
    maximum_points = tuple(v for v in vertices if abs(v.value - max_value) <= 1e-7)

    return GraphicalSolution(
        vertices=vertices,
        minimum=minimum_points[0],
        maximum=maximum_points[0],
        minimum_points=minimum_points,
        maximum_points=maximum_points,
    )


def format_fraction(value: float) -> str:
    from fractions import Fraction

    fraction = Fraction(value).limit_denominator(1000)
    if abs(float(fraction) - value) > 1e-8:
        return f"{value:.4f}"
    if fraction.denominator == 1:
        return str(fraction.numerator)
    return f"{fraction.numerator}/{fraction.denominator}"
