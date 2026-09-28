"""勤務表全体の統合テストとは独立した、最適化の優先順位と終了状態の検証。"""

import pytest
from ortools.sat.python import cp_model

from kinmu import optimizer
from kinmu.models import ModelData, Penalty


def empty_model():
    return ModelData(
        model=cp_model.CpModel(),
        full_work={},
        fixed_shifts={},
        holiday_targets={},
        holiday_off={},
        priority_penalties=[],
        principle_penalties=[],
        preference_penalties=[],
        four_day_windows=[],
        relaxed_four_consecutive=False,
    )


@pytest.mark.parametrize(
    ("higher_tier", "lower_tier"),
    [
        ("priority_penalties", "principle_penalties"),
        ("priority_penalties", "preference_penalties"),
        ("principle_penalties", "preference_penalties"),
    ],
)
def test_one_higher_priority_point_beats_all_lower_penalties(
    higher_tier, lower_tier
):
    data = empty_model()
    higher = data.model.new_bool_var("higher")
    lower = data.model.new_int_var(0, 1000, "lower")
    data.model.add(lower == 1000 * (1 - higher))
    getattr(data, higher_tier).append(Penalty("higher", higher, 1, 1))
    getattr(data, lower_tier).append(Penalty("lower", lower, 1000, 1000))

    solver, status = optimizer.solve_model(data, max_seconds=1)

    assert status == cp_model.OPTIMAL
    assert solver.value(higher) == 0
    assert solver.value(lower) == 1000


def mock_solver_statuses(monkeypatch, statuses):
    calls = []
    remaining_statuses = iter(statuses)

    def build_model(*args, relax_four_consecutive):
        data = empty_model()
        data.relaxed_four_consecutive = relax_four_consecutive
        return data

    def solve_model(data, *, max_seconds):
        calls.append((data.relaxed_four_consecutive, max_seconds))
        return cp_model.CpSolver(), next(remaining_statuses)

    monkeypatch.setattr(optimizer, "build_model", build_model)
    monkeypatch.setattr(optimizer, "solve_model", solve_model)
    return calls


@pytest.mark.parametrize("status", [cp_model.OPTIMAL, cp_model.FEASIBLE])
def test_strict_solution_does_not_try_relaxed_mode(monkeypatch, status):
    calls = mock_solver_statuses(monkeypatch, [status])

    result = optimizer.solve_schedule(None, [], [], {}, max_seconds=7)

    assert result.mode == "strict"
    assert result.status == status
    assert calls == [(False, 7)]


@pytest.mark.parametrize(
    ("strict_status", "message"),
    [
        (cp_model.UNKNOWN, "制限時間内に勤務表が見つからなかった"),
        (cp_model.INFEASIBLE, "条件は同時に成立しない"),
    ],
)
def test_fallback_distinguishes_timeout_from_infeasibility(
    monkeypatch, capsys, strict_status, message
):
    calls = mock_solver_statuses(
        monkeypatch, [strict_status, cp_model.FEASIBLE]
    )

    result = optimizer.solve_schedule(None, [], [], {}, max_seconds=3)

    output = capsys.readouterr().out
    assert message in output
    if strict_status == cp_model.UNKNOWN:
        assert "成立しない" not in output
    assert result.mode == "relaxed"
    assert result.status == cp_model.FEASIBLE
    assert calls == [(False, 3), (True, 3)]


@pytest.mark.parametrize(
    ("statuses", "message"),
    [
        ([cp_model.MODEL_INVALID], "通常モードのOR-Toolsモデルが不正"),
        (
            [cp_model.INFEASIBLE, cp_model.MODEL_INVALID],
            "緩和モードのOR-Toolsモデルが不正",
        ),
        (
            [cp_model.INFEASIBLE, cp_model.INFEASIBLE],
            "4連勤を緩和しても勤務表を作成できません",
        ),
        (
            [cp_model.UNKNOWN, cp_model.UNKNOWN],
            "緩和モードでも制限時間内に勤務表を作成できません",
        ),
    ],
)
def test_unsuccessful_search_reports_cause(monkeypatch, statuses, message):
    calls = mock_solver_statuses(monkeypatch, statuses)

    with pytest.raises(RuntimeError, match=message):
        optimizer.solve_schedule(None, [], [], {})

    assert len(calls) == len(statuses)


def test_new_solver_uses_one_worker_and_fixed_seed():
    solver = optimizer.new_solver(2.5)

    assert solver.parameters.max_time_in_seconds == 2.5
    assert solver.parameters.num_search_workers == 1
    assert solver.parameters.random_seed == 42
