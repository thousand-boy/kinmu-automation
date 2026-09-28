"""職員・休日数と、最適化処理で受け渡すデータ。"""

from __future__ import annotations

from dataclasses import dataclass

from ortools.sat.python import cp_model

CellKey = tuple[int, int]


@dataclass(frozen=True)
class EmployeeMaster:
    employee_id: str
    display_name: str
    ward: int
    is_dedicated: bool
    is_assigned: bool
    is_ot: bool
    is_active: bool


@dataclass(frozen=True)
class Employee:
    employee_id: str
    name: str
    row: int
    ward: int
    is_dedicated: bool
    is_assigned: bool
    is_ot: bool


@dataclass(frozen=True)
class HolidayTarget:
    public_half_units: int
    paid_half_units: int

    @property
    def total_half_units(self) -> int:
        return self.public_half_units + self.paid_half_units


@dataclass(frozen=True)
class Penalty:
    label: str
    variable: cp_model.IntVar
    weight: int
    max_value: int


@dataclass(frozen=True)
class FourDayWindow:
    employee_name: str
    start_label: str
    end_label: str
    keys: tuple[CellKey, ...]
    fixed_only: bool


@dataclass
class ModelData:
    model: cp_model.CpModel
    full_work: dict[CellKey, cp_model.IntVar]
    fixed_shifts: dict[CellKey, str]
    holiday_targets: dict[str, HolidayTarget]
    holiday_off: dict[str, cp_model.LinearExpr]
    priority_penalties: list[Penalty]
    principle_penalties: list[Penalty]
    preference_penalties: list[Penalty]
    four_day_windows: list[FourDayWindow]
    relaxed_four_consecutive: bool


@dataclass(frozen=True)
class SolveResult:
    solver: cp_model.CpSolver
    data: ModelData
    mode: str
    status: int
