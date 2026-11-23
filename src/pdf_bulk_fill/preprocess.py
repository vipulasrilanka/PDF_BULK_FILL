"""Text preprocessing, legacy formatters, and fixed-box comb line splitting."""

from __future__ import annotations

import math
import re
from datetime import datetime
from typing import Any, Callable, Sequence

from num2words import num2words

from pdf_bulk_fill.errors import PreprocessError

# Matches signed numbers with optional comma grouping, leading decimals (.50),
# and scientific notation (1e-05, 2.5E+3). Requires at least one digit.
_NUMBER_PATTERN = re.compile(
    r"[+-]?(?:\d[\d,]*(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?"
)


def extract_number(value: Any, cast: Callable[[float], Any] = float) -> Any:
    """Extract a numeric value from an int, float, or formatted string.

    Raises:
        PreprocessError: If ``value`` does not contain a valid finite number.
    """
    if isinstance(value, bool) or value is None:
        raise PreprocessError(f"Invalid numeric value: {value!r}")
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise PreprocessError(f"Non-finite numeric value: {value!r}")
        return cast(value)

    raw_text = str(value).strip()
    if not raw_text:
        raise PreprocessError("Empty string is not a valid numeric value.")

    # Fast path: clean direct numeric string (handles 1e-05, .50, -1,234.50, etc.)
    cleaned_direct = raw_text.replace(",", "")
    if ",," not in raw_text and not raw_text.startswith(",") and not raw_text.endswith(","):
        try:
            direct_val = float(cleaned_direct)
            if not math.isfinite(direct_val):
                raise PreprocessError(f"Non-finite numeric value: {value!r}")
            return cast(direct_val)
        except ValueError:
            pass

    # Fallback: extract single embedded numeric token (e.g., "Rs. 1,234.50", "USD .50")
    matches = _NUMBER_PATTERN.findall(raw_text)
    if len(matches) != 1:
        raise PreprocessError(f"Could not extract a valid number from {value!r}.")

    token = matches[0]
    if ",," in token or token.endswith(","):
        raise PreprocessError(f"Malformed numeric grouping in {value!r}.")

    try:
        parsed = float(token.replace(",", ""))
    except ValueError as exc:
        raise PreprocessError(f"Invalid numeric value {value!r}: {exc}") from exc

    if not math.isfinite(parsed):
        raise PreprocessError(f"Non-finite numeric value: {value!r}")
    return cast(parsed)


def format_date(date_string: Any, fmt: str | None = "%Y/%m/%d") -> str:
    """Parse common date representations (including openpyxl timestamps) and format them."""
    if isinstance(date_string, datetime):
        target_fmt = fmt if isinstance(fmt, str) and fmt else "%Y/%m/%d"
        return date_string.strftime(target_fmt)

    if not isinstance(date_string, str):
        return str(date_string)

    target_fmt = fmt if isinstance(fmt, str) and fmt else "%Y/%m/%d"
    candidate_formats = [
        target_fmt,
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
    ]
    for candidate in candidate_formats:
        try:
            parsed = datetime.strptime(date_string.strip(), candidate)
            return parsed.strftime(target_fmt)
        except ValueError:
            continue
    return date_string


def _parse_decimal_places(decimal_points: Any) -> int:
    try:
        decimals = int(decimal_points)
    except (TypeError, ValueError) as exc:
        raise PreprocessError(f"Invalid decimal places specification: {decimal_points!r}") from exc
    if decimals < 0:
        raise PreprocessError(f"Decimal places must be non-negative: {decimals}")
    return decimals


def format_number(
    value: Any,
    decimal_points: int = 2,
    prefix: str = "",
    suffix: str = "",
) -> str:
    """Format a number with thousands separators, fixed decimals, prefix, and suffix."""
    num = extract_number(value, float)
    decimals = _parse_decimal_places(decimal_points)
    return f"{prefix}{num:,.{decimals}f}{suffix}"


def format_currency(
    value: Any,
    decimal_points: int = 2,
    currency: str = "",
) -> str:
    """Format a number as currency with thousands separators."""
    num = extract_number(value, float)
    decimals = _parse_decimal_places(decimal_points)
    return f"{currency} {num:,.{decimals}f}".strip()


def apply_preprocess(text: Any, preprocess_spec: dict[str, Any] | None) -> str:
    """Apply a legacy template preprocess specification to ``text``."""
    text_str = "" if text is None else str(text)
    if not preprocess_spec or "Function" not in preprocess_spec:
        return text_str

    func = preprocess_spec["Function"]
    if not isinstance(func, dict):
        return text_str

    name = func.get("name")
    if name == "NumberToText":
        mode = func.get("param2")
        if mode == "Integer":
            number = extract_number(text_str, round)
        else:
            number = extract_number(text_str, float)
        return num2words(number, to="cardinal").capitalize()

    if name == "AddSpace":
        count = max(0, int(extract_number(func.get("param2", 1), int)))
        return text_str + (" " * count)

    if name == "FormatDate":
        return format_date(text, func.get("param2"))

    if name == "NumberToCurrency":
        return format_currency(
            text_str,
            decimal_points=func.get("param2", 2),
            currency=str(func.get("param3", "")),
        )

    if name == "FormatNumber":
        return format_number(
            text_str,
            decimal_points=func.get("param2", 2),
            prefix=str(func.get("param3", "")),
            suffix=str(func.get("param4", "")),
        )

    if name == "changeTextCase":
        case_type = func.get("param2")
        if case_type == "lower":
            return text_str.lower()
        if case_type == "UPPER":
            return text_str.upper()
        if case_type == "Title":
            return text_str.title()
        return text_str

    return text_str


def split_for_comb(
    text: str,
    box_lengths: Sequence[int],
    *,
    uppercase: bool = True,
    strict: bool = False,
) -> list[str]:
    """Split text across multiple fixed-box comb lines (e.g., [22, 31, 31]).

    Packs words cleanly at word boundaries where possible. If a single word
    exceeds the current line's box count, it is split across lines.

    Args:
        text: Input string to split.
        box_lengths: Sequence of positive box counts per line.
        uppercase: If True, converts text to uppercase (block letters).
        strict: If True, raises ValueError when text exceeds total box capacity;
            otherwise truncates cleanly at the end of the final line.
    """
    if not box_lengths or any(length <= 0 for length in box_lengths):
        raise ValueError("box_lengths must contain positive integers.")

    normalized = " ".join((text or "").split())
    if uppercase:
        normalized = normalized.upper()

    if not normalized:
        return [""] * len(box_lengths)

    words = normalized.split(" ")
    lines: list[str] = []
    line_idx = 0
    current = ""
    word_idx = 0

    while word_idx < len(words) and line_idx < len(box_lengths):
        limit = box_lengths[line_idx]
        word = words[word_idx]

        if not current:
            if len(word) <= limit:
                current = word
                word_idx += 1
            else:
                # Word is longer than the entire line: slice to fit and carry remainder
                lines.append(word[:limit])
                words[word_idx] = word[limit:]
                line_idx += 1
        else:
            candidate = f"{current} {word}"
            if len(candidate) <= limit:
                current = candidate
                word_idx += 1
            else:
                lines.append(current)
                current = ""
                line_idx += 1

    if current and line_idx < len(box_lengths):
        lines.append(current)
        line_idx += 1

    # Check for leftover words (overflow)
    if word_idx < len(words):
        if strict:
            raise ValueError(
                f"Text exceeds comb capacity {list(box_lengths)}: remaining {words[word_idx:]!r}"
            )

    while len(lines) < len(box_lengths):
        lines.append("")

    return lines


def validate_field_value(
    value: str,
    *,
    max_len: int | None = None,
    digits_only: bool = False,
    reject_digits: bool = False,
) -> str:
    """Validate a field value against length and character constraints."""
    cleaned = value.strip()
    if max_len is not None and len(cleaned) > max_len:
        raise ValueError(
            f"Value {cleaned!r} exceeds maximum length of {max_len} characters."
        )
    if digits_only and cleaned and not cleaned.isdigit():
        raise ValueError(f"Value {cleaned!r} must contain digits only.")
    if reject_digits and any(ch.isdigit() for ch in cleaned):
        raise ValueError(f"Value {cleaned!r} must not contain numeric digits.")
    return cleaned
