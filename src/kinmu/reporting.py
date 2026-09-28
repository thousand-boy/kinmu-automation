"""最適化結果と、確認が必要な条件を利用者へ表示する。"""

from ortools.sat.python import cp_model

from .models import Employee, ModelData, Penalty, SolveResult


def half_units_to_days(value: int) -> float:
    return value / 2


def four_day_violations(
    solver: cp_model.CpSolver, data: ModelData
) -> list[tuple[str, bool]]:
    return [
        (
            f"{window.employee_name}: "
            f"{window.start_label}〜{window.end_label}日が4連続勤務",
            window.fixed_only,
        )
        for window in data.four_day_windows
        if all(solver.value(data.full_work[key]) == 1 for key in window.keys)
    ]


def print_penalty_violations(
    title: str, solver: cp_model.CpSolver, penalties: list[Penalty]
) -> None:
    violations = []
    for penalty in penalties:
        value = int(solver.value(penalty.variable))
        if value > 0:
            violations.append((penalty.weight, value, penalty.label))

    print(f"\n===== {title} =====")
    if not violations:
        print("✅ すべて満たしています。")
        return
    for _, value, label in sorted(violations, reverse=True):
        print(f"⚠ {label}（差: {value}）")


def print_result(result: SolveResult, employees: list[Employee]) -> None:
    solver, data = result.solver, result.data
    print("\n===== 実行モード =====")
    if result.mode == "strict":
        print("✅ 通常モード（自動配置による4連勤禁止）")
    else:
        print("⚠ 緩和モード（休日数・日曜2名を守り、4連勤を削減）")

    if result.status == cp_model.OPTIMAL:
        print("✅ 最適性まで証明されました。")
    else:
        print("⚠ 実行可能解です。制限時間内に最適性までは証明されていません。")

    print("\n===== 個人別 休日数 =====")
    for employee in employees:
        target = data.holiday_targets[employee.employee_id].total_half_units
        actual = int(solver.value(data.holiday_off[employee.employee_id]))
        print(
            f"{employee.name}: 必要{half_units_to_days(target):g}日 / "
            f"割当{half_units_to_days(actual):g}日 / "
            f"残り{half_units_to_days(target - actual):g}日"
        )

    print("\n===== 4連勤 =====")
    violations = four_day_violations(solver, data)
    if not violations:
        print("✅ 該当なし")
    for label, fixed_only in violations:
        note = "本人希望等の固定勤務" if fixed_only else "緩和モードで発生・要目視修正"
        print(f"⚠ {label}（{note}）")

    print_penalty_violations("満たせなかった原則条件", solver, data.principle_penalties)
    print_penalty_violations("満たせなかった希望条件", solver, data.preference_penalties)
