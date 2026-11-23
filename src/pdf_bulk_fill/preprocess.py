"""Text preprocessing, legacy formatters, and fixed-box comb line splitting."""

from __future__ import annotations

from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Callable, Sequence

from num2words import num2words

from pdf_bulk_fill.currency_parser import parse_currency_amount
from pdf_bulk_fill.errors import PreprocessError


def extract_number(
    value: Any,
    cast: Callable[[Any], Any] = float,
    *,
    expected_currency: str | None = None,
    input_locale: str | None = None,
) -> Any:
    """Extract and validate a finite numeric value via the CLDR/Decimal currency parser.

    Raises:
        PreprocessError: If ``value`` is malformed, ambiguous, or has a currency mismatch.
    """
    parsed = parse_currency_amount(
        value,
        expected_currency=expected_currency,
        input_locale=input_locale,
    )
    if cast is Decimal:
        return parsed.amount
    if cast is round:
        return int(parsed.amount.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    return cast(parsed.amount)


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
    *,
    expected_currency: str | None = None,
    input_locale: str | None = None,
) -> str:
    """Format a number as exact Decimal with thousands separators, prefix, and suffix."""
    parsed = parse_currency_amount(
        value,
        expected_currency=expected_currency,
        input_locale=input_locale,
    )
    decimals = _parse_decimal_places(decimal_points)
    return f"{prefix}{parsed.amount:,.{decimals}f}{suffix}"


def format_currency(
    value: Any,
    decimal_points: int = 2,
    currency: str = "",
    *,
    input_locale: str | None = None,
) -> str:
    """Validate and format a currency amount using exact Decimal precision."""
    expected = currency.strip() if isinstance(currency, str) and currency.strip() else None
    parsed = parse_currency_amount(
        value,
        expected_currency=expected,
        input_locale=input_locale,
    )
    decimals = _parse_decimal_places(decimal_points)
    label = currency.strip() if isinstance(currency, str) else ""
    return f"{label} {parsed.amount:,.{decimals}f}".strip()


def _clean_opt_str(val: Any) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None


def apply_preprocess(
    text: Any,
    preprocess_spec: dict[str, Any] | None,
    *,
    default_currency: str | None = None,
    default_locale: str | None = None,
) -> str:
    """Apply a template preprocess specification to ``text`` with optional currency/locale defaults."""
    text_str = "" if text is None else str(text)
    if not preprocess_spec or "Function" not in preprocess_spec:
        return text_str

    func = preprocess_spec["Function"]
    if not isinstance(func, dict):
        return text_str

    # Per-field <Locale=...> / <InputLocale=...> and <Currency=...> keys in PreProcess cell
    spec_locale = (
        _clean_opt_str(preprocess_spec.get("InputLocale"))
        or _clean_opt_str(preprocess_spec.get("Locale"))
        or default_locale
    )
    spec_currency = _clean_opt_str(preprocess_spec.get("Currency")) or default_currency

    name = func.get("name")
    if name == "NumberToText":
        mode = func.get("param2")
        eff_locale = _clean_opt_str(func.get("param3")) or spec_locale
        eff_currency = _clean_opt_str(func.get("param4")) or spec_currency
        if mode == "Integer":
            number = extract_number(
                text_str,
                round,
                expected_currency=eff_currency,
                input_locale=eff_locale,
            )
        else:
            number = extract_number(
                text_str,
                float,
                expected_currency=eff_currency,
                input_locale=eff_locale,
            )
        return num2words(number, to="cardinal").capitalize()

    if name == "AddSpace":
        count = max(0, int(extract_number(func.get("param2", 1), int)))
        return text_str + (" " * count)

    if name == "FormatDate":
        return format_date(text, func.get("param2"))

    if name == "NumberToCurrency":
        eff_currency = _clean_opt_str(func.get("param3")) or spec_currency or ""
        eff_locale = _clean_opt_str(func.get("param4")) or spec_locale
        return format_currency(
            text_str,
            decimal_points=func.get("param2", 2),
            currency=eff_currency,
            input_locale=eff_locale,
        )

    if name == "FormatNumber":
        eff_locale = _clean_opt_str(func.get("param5")) or spec_locale
        eff_currency = _clean_opt_str(func.get("param6")) or spec_currency
        return format_number(
            text_str,
            decimal_points=func.get("param2", 2),
            prefix=str(func.get("param3", "")),
            suffix=str(func.get("param4", "")),
            expected_currency=eff_currency,
            input_locale=eff_locale,
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
