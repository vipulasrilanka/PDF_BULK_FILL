"""Unit tests for legacy text preprocessors, comb line splitting, and validation."""

from __future__ import annotations

from datetime import datetime

import pytest

from pdf_bulk_fill import (
    PreprocessError,
    apply_preprocess,
    extract_number,
    format_currency,
    format_date,
    format_number,
    split_for_comb,
    validate_field_value,
)


def test_extract_number() -> None:
    assert extract_number(42, int) == 42
    assert extract_number(3.14, float) == 3.14
    assert extract_number("Rs. 12,345.67", float, expected_currency="LKR") == 12345.67
    with pytest.raises(PreprocessError):
        extract_number("no digits", float)


def test_format_date() -> None:
    assert format_date("2024-05-18 00:00:00", "%d/%m/%Y") == "18/05/2024"
    assert format_date(datetime(2024, 5, 18), "%Y/%m/%d") == "2024/05/18"
    assert format_date("2024/05/18", None) == "2024/05/18"
    assert format_date("not-a-date", "%Y/%m/%d") == "not-a-date"
    assert format_date(12345, "%Y/%m/%d") == "12345"


def test_format_number_and_currency() -> None:
    assert format_number("12500.5", 2, "LKR ", "/=") == "LKR 12,500.50/="
    assert format_currency("12500.5", 2, "LKR") == "LKR 12,500.50"


def test_apply_preprocess_all_legacy_functions() -> None:
    assert apply_preprocess("hello", None) == "hello"
    assert apply_preprocess(None, None) == ""
    assert apply_preprocess("hello", {"Function": "not_a_dict"}) == "hello"

    assert (
        apply_preprocess(
            "kamal perera", {"Function": {"name": "changeTextCase", "param2": "UPPER"}}
        )
        == "KAMAL PERERA"
    )
    assert (
        apply_preprocess(
            "KAMAL PERERA", {"Function": {"name": "changeTextCase", "param2": "lower"}}
        )
        == "kamal perera"
    )
    assert (
        apply_preprocess(
            "kamal perera", {"Function": {"name": "changeTextCase", "param2": "Title"}}
        )
        == "Kamal Perera"
    )
    assert (
        apply_preprocess(
            "Kamal", {"Function": {"name": "changeTextCase", "param2": "Unknown"}}
        )
        == "Kamal"
    )

    assert (
        apply_preprocess("abc", {"Function": {"name": "AddSpace", "param2": 3}})
        == "abc   "
    )
    assert (
        apply_preprocess(
            "125.6", {"Function": {"name": "NumberToText", "param2": "Integer"}}
        )
        == "One hundred and twenty-six"
    )
    assert (
        apply_preprocess(
            "12.5", {"Function": {"name": "NumberToText", "param2": "Floating Point"}}
        )
        == "Twelve point five"
    )
    assert (
        apply_preprocess(
            "2024-03-15 00:00:00",
            {"Function": {"name": "FormatDate", "param2": "%d-%m-%Y"}},
        )
        == "15-03-2024"
    )
    assert (
        apply_preprocess(
            "5000",
            {"Function": {"name": "NumberToCurrency", "param2": 2, "param3": "Rs."}},
            default_locale="en_LK",
        )
        == "Rs. 5,000.00"
    )
    assert (
        apply_preprocess(
            "5000",
            {
                "Function": {
                    "name": "FormatNumber",
                    "param2": 2,
                    "param3": "Rs. ",
                    "param4": "/-",
                }
            },
        )
        == "Rs. 5,000.00/-"
    )
    assert (
        apply_preprocess("keep", {"Function": {"name": "UnsupportedFunc"}}) == "keep"
    )


def test_split_for_comb_rr_form_full_name() -> None:
    full_name = "don kamal bandara wickramasinghe mudiyanselage perera"
    lines = split_for_comb(full_name, [22, 31, 31])
    assert len(lines) == 3
    assert all(len(line) <= cap for line, cap in zip(lines, [22, 31, 31]))
    assert lines[0] == "DON KAMAL BANDARA"
    assert lines[1] == "WICKRAMASINGHE MUDIYANSELAGE"
    assert lines[2] == "PERERA"


def test_split_for_comb_initials_and_empty() -> None:
    assert split_for_comb("", [20, 31]) == ["", ""]
    assert split_for_comb("D. K. B. Perera", [20, 31]) == ["D. K. B. PERERA", ""]
    assert split_for_comb("D. K. B. Perera", [20, 31], uppercase=False) == [
        "D. K. B. Perera",
        "",
    ]


def test_split_for_comb_long_single_word_and_overflow() -> None:
    # Single word exceeding first line length splits cleanly across lines
    lines = split_for_comb("ABCDEFGHIJKLMNOPQRSTUVWXYZ", [10, 10, 10])
    assert lines == ["ABCDEFGHIJ", "KLMNOPQRST", "UVWXYZ"]

    # Non-strict overflow truncates cleanly
    truncated = split_for_comb("ONE TWO THREE FOUR FIVE", [5, 5], strict=False)
    assert truncated == ["ONE", "TWO"]

    # Strict overflow raises ValueError
    with pytest.raises(ValueError, match="exceeds comb capacity"):
        split_for_comb("ONE TWO THREE FOUR FIVE", [5, 5], strict=True)

    # Invalid box lengths
    with pytest.raises(ValueError, match="positive integers"):
        split_for_comb("ABC", [])
    with pytest.raises(ValueError, match="positive integers"):
        split_for_comb("ABC", [10, 0])


def test_validate_field_value() -> None:
    assert validate_field_value(" 0771234567 ", max_len=10, digits_only=True) == "0771234567"
    assert validate_field_value("KAMAL PERERA", max_len=22, reject_digits=True) == "KAMAL PERERA"

    with pytest.raises(ValueError, match="exceeds maximum length"):
        validate_field_value("12345678901", max_len=10)

    with pytest.raises(ValueError, match="digits only"):
        validate_field_value("077-123456", digits_only=True)

    with pytest.raises(ValueError, match="must not contain numeric digits"):
        validate_field_value("KAMAL 123", reject_digits=True)
