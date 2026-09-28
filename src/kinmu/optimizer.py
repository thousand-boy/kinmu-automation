"""CP-SATによる勤務割り当てと、優先順位付きの制約定義。"""

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet
from ortools.sat.python import cp_model

from .constants import (
    MON_TO_SAT,
    WARD_STAFF_LIMITS,
    WEIGHT_EXISTING_SHIFT,
    WEIGHT_FRAGMENTED_WORK,
    WEIGHT_ISOLATED_WORK,
    WEIGHT_LEADER,
    WEIGHT_OT_MINIMUM,
    WEIGHT_OT_SECOND,
    WEIGHT_THIRD_OFF,
    WEIGHT_TOO_MANY_OFF,
    WEIGHT_WARD_STAFF,
)
from .models import (
    Employee,
    FourDayWindow,
    HolidayTarget,
    ModelData,
    Penalty,
    SolveResult,
)
from .roster import normalize
from .validation import is_fixed_shift
from .workbook import day_label, weekday


def add_penalty(
    penalties: list[Penalty],
    label: str,
    variable: cp_model.IntVar,
    weight: int,
    max_value: int,
) -> None:
    penalties.append(Penalty(label, variable, weight, max_value))


def shortage(
    model: cp_model.CpModel,
    expression,
    target: int,
    upper_bound: int,
    name: str,
) -> cp_model.IntVar:
    """目的関数で最小化する、不足数の下限変数を作る。"""
    variable = model.new_int_var(0, upper_bound, name)
    model.add(variable >= target - expression)
    return variable


def excess(
    model: cp_model.CpModel,
    expression,
    limit: int,
    upper_bound: int,
    name: str,
) -> cp_model.IntVar:
    """目的関数で最小化する、超過数の下限変数を作る。"""
    variable = model.new_int_var(0, upper_bound, name)
    model.add(variable >= expression - limit)
    return variable


def _create_shift_variables(sheet, employees, columns, data):
    model = data.model
    presence = {}
    off_half_units = {}

    for employee in employees:
        for column in columns:
            key = (employee.row, column)
            work = normalize(sheet.cell(employee.row, column).value)
            reason = normalize(sheet.cell(employee.row + 1, column).value)
            work_var = model.new_bool_var(f"work_{employee.row}_{column}")
            data.full_work[key] = work_var

            if is_fixed_shift(work, reason):
                data.fixed_shifts[key] = work
                model.add(work_var == int(work == "出"))
                # 半休は終日勤務の連続を切るが、配置人数には1人と数える。
                presence[key] = int(work != "休")
                off_half_units[key] = {"出": 0, "休": 2, "半/": 1, "/半": 1}[work]
                continue

            presence[key] = work_var
            off_half_units[key] = 2 - 2 * work_var
            if work not in {"出", "休"}:
                continue

            changed = model.new_bool_var(
                f"change_existing_{employee.row}_{column}"
            )
            if work == "出":
                model.add(changed + work_var == 1)
            else:
                model.add(changed == work_var)
            add_penalty(
                data.preference_penalties,
                f"{employee.name} {day_label(sheet, column)}日 既存「{work}」を変更",
                changed,
                WEIGHT_EXISTING_SHIFT,
                1,
            )

    return presence, off_half_units


def _add_holiday_constraints(employees, columns, off_half_units, data):
    for employee in employees:
        actual_off = sum(
            off_half_units[employee.row, column] for column in columns
        )
        target = data.holiday_targets[employee.employee_id].total_half_units
        data.model.add(actual_off == target)
        data.holiday_off[employee.employee_id] = actual_off


def _add_four_day_constraints(sheet, employees, columns, data):
    for employee in employees:
        for index in range(len(columns) - 3):
            window = columns[index:index + 4]
            if window != list(range(window[0], window[0] + 4)):
                continue

            keys = tuple((employee.row, column) for column in window)
            fixed_only = all(data.fixed_shifts.get(key) == "出" for key in keys)
            start = day_label(sheet, window[0])
            end = day_label(sheet, window[-1])
            data.four_day_windows.append(
                FourDayWindow(employee.name, start, end, keys, fixed_only)
            )
            # 本人希望など、固定勤務だけで成立する4連勤は変更しない。
            if fixed_only:
                continue

            work_count = sum(data.full_work[key] for key in keys)
            if not data.relaxed_four_consecutive:
                data.model.add(work_count <= 3)
                continue

            violation = excess(
                data.model, work_count, 3, 1,
                f"four_consecutive_{employee.row}_{window[0]}",
            )
            add_penalty(
                data.priority_penalties,
                f"{employee.name}: {start}〜{end}日が4連続勤務",
                violation,
                1,
                1,
            )


def _presence_count(presence, members, column):
    return sum(presence[employee.row, column] for employee in members)


def _add_role_constraints(ward_members, ots, column, label, presence, data):
    for ward, members in ward_members.items():
        leaders = [e for e in members if e.is_dedicated or e.is_assigned]
        missing = shortage(
            data.model,
            _presence_count(presence, leaders, column),
            1,
            1,
            f"leader_shortage_{ward}_{column}",
        )
        add_penalty(
            data.principle_penalties,
            f"{label} {ward}病棟 専従・専任不在",
            missing,
            WEIGHT_LEADER,
            1,
        )

    ot_count = _presence_count(presence, ots, column)
    ot_minimum = shortage(data.model, ot_count, 1, 1, f"ot_min_{column}")
    add_penalty(
        data.principle_penalties,
        f"{label} OTが0人",
        ot_minimum,
        WEIGHT_OT_MINIMUM,
        1,
    )
    ot_second = shortage(data.model, ot_count, 2, 2, f"ot_second_{column}")
    add_penalty(
        data.preference_penalties,
        f"{label} OTが2人未満",
        ot_second,
        WEIGHT_OT_SECOND,
        2,
    )


def _add_day_off_constraints(ward_members, column, label, presence, data):
    for ward, members in ward_members.items():
        off_count = len(members) - _presence_count(presence, members, column)
        over_two = excess(
            data.model, off_count, 2, len(members),
            f"off_over_two_{ward}_{column}",
        )
        over_three = excess(
            data.model, off_count, 3, len(members),
            f"off_over_three_{ward}_{column}",
        )
        # 4人目以降には両方の重みを課し、3人目より強く避ける。
        add_penalty(
            data.principle_penalties,
            f"{label} {ward}病棟 休み2人超過",
            over_two,
            WEIGHT_THIRD_OFF,
            len(members),
        )
        add_penalty(
            data.principle_penalties,
            f"{label} {ward}病棟 休み3人超過",
            over_three,
            WEIGHT_TOO_MANY_OFF,
            len(members),
        )


def _add_headcount_constraints(ward_members, column, label, presence, data):
    for ward, members in ward_members.items():
        minimum, maximum = WARD_STAFF_LIMITS[ward]
        count = _presence_count(presence, members, column)
        if minimum is not None:
            too_few = shortage(
                data.model, count, minimum, len(members),
                f"ward_{ward}_few_{column}",
            )
            add_penalty(
                data.principle_penalties,
                f"{label} {ward}病棟 {minimum}人未満",
                too_few,
                WEIGHT_WARD_STAFF,
                len(members),
            )
        if maximum is not None:
            too_many = excess(
                data.model, count, maximum, len(members),
                f"ward_{ward}_many_{column}",
            )
            add_penalty(
                data.principle_penalties,
                f"{label} {ward}病棟 {maximum}人超過",
                too_many,
                WEIGHT_WARD_STAFF,
                len(members),
            )


def _add_daily_constraints(sheet, employees, columns, presence, data):
    wards = sorted({employee.ward for employee in employees})
    ward_members = {
        ward: [employee for employee in employees if employee.ward == ward]
        for ward in wards
    }
    ots = [employee for employee in employees if employee.is_ot]

    for column in columns:
        current_weekday = weekday(sheet, column)
        label = f"{day_label(sheet, column)}日({current_weekday})"
        if current_weekday == "日":
            for members in ward_members.values():
                data.model.add(_presence_count(presence, members, column) == 2)
        if current_weekday in MON_TO_SAT:
            _add_role_constraints(
                ward_members, ots, column, label, presence, data
            )
        if current_weekday in {"水", "木", "金"}:
            _add_day_off_constraints(
                ward_members, column, label, presence, data
            )
        if current_weekday in {"月", "火", "土"}:
            # 月・火・土は人数目安、水・木・金は休み人数を評価する。
            _add_headcount_constraints(
                ward_members, column, label, presence, data
            )


def _add_shift_rhythm_preferences(sheet, employees, columns, data):
    for employee in employees:
        windows = zip(columns, columns[1:], columns[2:])
        for previous, current, following in windows:
            if following - previous != 2:
                continue
            previous_work = data.full_work[employee.row, previous]
            current_work = data.full_work[employee.row, current]
            next_work = data.full_work[employee.row, following]
            day = day_label(sheet, current)

            isolated = data.model.new_bool_var(
                f"isolated_work_{employee.row}_{current}"
            )
            data.model.add(
                isolated >= current_work - previous_work - next_work
            )
            add_penalty(
                data.preference_penalties,
                f"{employee.name}: {day}日が1日だけ勤務",
                isolated,
                WEIGHT_ISOLATED_WORK,
                1,
            )

            fragmented = data.model.new_bool_var(
                f"fragmented_work_{employee.row}_{current}"
            )
            data.model.add(
                fragmented >= previous_work - current_work + next_work - 1
            )
            add_penalty(
                data.preference_penalties,
                f"{employee.name}: {day}日を挟んで勤務が分断",
                fragmented,
                WEIGHT_FRAGMENTED_WORK,
                1,
            )


def build_model(
    sheet: Worksheet,
    employees: list[Employee],
    columns: list[int],
    holiday_targets: dict[str, HolidayTarget],
    *,
    relax_four_consecutive: bool,
) -> ModelData:
    """固定勤務・休日数・日曜各病棟2名を維持して勤務モデルを作る。"""
    data = ModelData(
        model=cp_model.CpModel(),
        full_work={},
        fixed_shifts={},
        holiday_targets=holiday_targets,
        holiday_off={},
        priority_penalties=[],
        principle_penalties=[],
        preference_penalties=[],
        four_day_windows=[],
        relaxed_four_consecutive=relax_four_consecutive,
    )
    presence, off_half_units = _create_shift_variables(
        sheet, employees, columns, data
    )
    _add_holiday_constraints(employees, columns, off_half_units, data)
    _add_four_day_constraints(sheet, employees, columns, data)
    _add_daily_constraints(sheet, employees, columns, presence, data)
    _add_shift_rhythm_preferences(sheet, employees, columns, data)
    return data


def penalty_cost(penalties: list[Penalty]):
    return sum(penalty.variable * penalty.weight for penalty in penalties)


def penalty_upper_bound(penalties: list[Penalty]) -> int:
    return sum(penalty.weight * penalty.max_value for penalty in penalties)


def build_objective(data: ModelData):
    """4連勤の削減 > 原則条件 > 希望条件の優先順位を保証する。"""
    preference_max = penalty_upper_bound(data.preference_penalties)
    principle_max = penalty_upper_bound(data.principle_penalties)
    # 上位1ポイントの改善が、下位の全ペナルティの悪化に必ず勝る係数。
    principle_multiplier = preference_max + 1
    priority_multiplier = (
        principle_max * principle_multiplier + preference_max + 1
    )
    return (
        penalty_cost(data.priority_penalties) * priority_multiplier
        + penalty_cost(data.principle_penalties) * principle_multiplier
        + penalty_cost(data.preference_penalties)
    )


def new_solver(max_seconds: float) -> cp_model.CpSolver:
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max_seconds
    solver.parameters.random_seed = 42
    solver.parameters.num_search_workers = 1
    return solver


def solve_model(
    data: ModelData, *, max_seconds: float
) -> tuple[cp_model.CpSolver, int]:
    data.model.minimize(build_objective(data))
    solver = new_solver(max_seconds)
    return solver, solver.solve(data.model)


def solve_schedule(
    sheet: Worksheet,
    employees: list[Employee],
    columns: list[int],
    holiday_targets: dict[str, HolidayTarget],
    *,
    max_seconds: float = 60,
) -> SolveResult:
    """通常モードで解が見つからなければ、4連勤のみを緩和して再試行する。"""
    for mode in ("strict", "relaxed"):
        data = build_model(
            sheet, employees, columns, holiday_targets,
            relax_four_consecutive=mode == "relaxed",
        )
        solver, status = solve_model(data, max_seconds=max_seconds)
        if status in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
            return SolveResult(solver, data, mode, status)
        mode_label = "通常" if mode == "strict" else "緩和"
        if status == cp_model.MODEL_INVALID:
            raise RuntimeError(f"{mode_label}モードのOR-Toolsモデルが不正です。")

        if mode == "strict":
            reason = (
                "自動4連勤禁止を含む条件は同時に成立しないため"
                if status == cp_model.INFEASIBLE
                else "通常モードでは制限時間内に勤務表が見つからなかったため"
            )
            print(f"⚠ {reason}、4連勤のみ緩和して暫定勤務表を作成します。")

    if status == cp_model.INFEASIBLE:
        raise RuntimeError(
            "4連勤を緩和しても勤務表を作成できません。\n"
            "固定勤務・必要休日数・日曜日各病棟2名の条件を確認してください。"
        )
    raise RuntimeError("緩和モードでも制限時間内に勤務表を作成できませんでした。")
