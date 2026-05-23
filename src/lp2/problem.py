from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import numpy as np


@dataclass(frozen=True)
class Constraint:
    """Linear inequality a1*x1 + a2*x2 (sense) b."""

    coefficients: tuple[float, float]
    sense: str
    rhs: float
    label: str

    def normalized(self) -> "Constraint":
        """Return an equivalent inequality with non-negative right side."""
        if self.rhs >= 0:
            return self
        flipped = {"<=": ">=", ">=": "<=", "=": "="}[self.sense]
        return Constraint(
            coefficients=tuple(-value for value in self.coefficients),
            sense=flipped,
            rhs=-self.rhs,
            label=f"-({self.label})",
        )

    def as_leq(self) -> tuple[np.ndarray, float]:
        """Return the inequality in the form a @ x <= b."""
        a = np.array(self.coefficients, dtype=float)
        if self.sense == "<=":
            return a, float(self.rhs)
        if self.sense == ">=":
            return -a, -float(self.rhs)
        raise ValueError("Equality constraints are not supported by as_leq().")

    def boundary(self) -> tuple[np.ndarray, float]:
        return np.array(self.coefficients, dtype=float), float(self.rhs)


@dataclass(frozen=True)
class LinearProgram:
    variant: int
    objective: tuple[float, float]
    objective_label: str
    constraints: tuple[Constraint, ...]

    @property
    def c(self) -> np.ndarray:
        return np.array(self.objective, dtype=float)

    def all_constraints(self) -> Iterable[Constraint]:
        yield from self.constraints
        yield Constraint((1.0, 0.0), ">=", 0.0, "x1 >= 0")
        yield Constraint((0.0, 1.0), ">=", 0.0, "x2 >= 0")


VARIANT_16 = LinearProgram(
    variant=16,
    objective=(-3.0, 6.0),
    objective_label="Z = -3*x1 + 6*x2",
    constraints=(
        Constraint((5.0, -2.0), "<=", 4.0, "5*x1 - 2*x2 <= 4"),
        Constraint((1.0, -2.0), ">=", -4.0, "x1 - 2*x2 >= -4"),
        Constraint((1.0, 1.0), ">=", 4.0, "x1 + x2 >= 4"),
    ),
)
