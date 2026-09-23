import pytest

from src.main import (
    HolidayTarget,
    employee_name,
    is_fixed_shift,
    normalize,
    numeric_to_half_units,
    ward_for,
)


def test_normalize():
    assert normalize(" 休 ") == "休"
    assert normalize("") is None
    assert normalize("   ") is None
    assert normalize(None) is None
    assert normalize(10) == 10


@pytest.mark.parametrize(
    ("raw_value", "expected"),
    [
        ("A（専従）", "A"),
        ("B（専任）", "B"),
        ("D（作業療法士）", "D"),
        ("A(専従)", "A"),
        (" C ", "C"),
        (None, None),
    ],
)
def test_employee_name(raw_value, expected):
    assert employee_name(raw_value) == expected


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("A", 1),
        ("D", 1),
        ("E", 2),
        ("K", 2),
        ("L", 3),
        ("P", 3),
    ],
)
def test_ward_for(name, expected):
    assert ward_for(name) == expected


def test_ward_for_unknown_employee():
    with pytest.raises(
        ValueError,
        match="病棟を特定できません",
    ):
        ward_for("Z")


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (10, 20),
        (1, 2),
        (0.5, 1),
        (1.5, 3),
        (0, 0),
    ],
)
def test_numeric_to_half_units(value, expected):
    assert (
        numeric_to_half_units(
            value,
            "テスト",
        )
        == expected
    )


def test_numeric_to_half_units_rejects_quarter_day():
    with pytest.raises(
        ValueError,
        match="0.5日単位",
    ):
        numeric_to_half_units(
            1.25,
            "テスト",
        )


@pytest.mark.parametrize(
    ("work", "reason", "expected"),
    [
        ("休", "希望", True),
        ("出", "委員会", True),
        ("半/", "年/", True),
        ("/半", "/年", True),
        ("休", None, False),
        (None, "希望", False),
        (None, None, False),
    ],
)
def test_is_fixed_shift(
    work,
    reason,
    expected,
):
    assert (
        is_fixed_shift(
            work,
            reason,
        )
        is expected
    )


def test_holiday_target_total():
    target = HolidayTarget(
        public_half_units=20,
        paid_half_units=3,
    )

    assert target.total_half_units == 23