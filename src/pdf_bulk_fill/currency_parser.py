"""CLDR/Babel-backed currency and locale-aware Decimal number parser."""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from typing import Any

from babel import Locale, UnknownLocaleError
from babel.core import parse_locale
from babel.numbers import (
    get_decimal_symbol,
    get_group_symbol,
    get_territory_currencies,
)

from pdf_bulk_fill.errors import PreprocessError

# Space characters allowed between currency markers and amounts
_CURRENCY_SPACE_CHARS = " \t\u00a0\u202f\u2007\u2009"

# Valid ISO 4217 currency codes from Unicode CLDR via Babel
ISO_CURRENCIES: frozenset[str] = frozenset(Locale("en").currencies.keys())

# Recognized currency symbols & aliases -> set of candidate ISO 4217 codes.
# Single-element sets are unambiguous; multi-element sets require expected_currency or locale.
_SYMBOL_CANDIDATES: dict[str, frozenset[str]] = {
    # Unambiguous symbols & prefixed symbols
    "€": frozenset({"EUR"}),
    "£": frozenset({"GBP"}),
    "₹": frozenset({"INR"}),
    "රු.": frozenset({"LKR"}),
    "රු": frozenset({"LKR"}),
    "US$": frozenset({"USD"}),
    "CA$": frozenset({"CAD"}),
    "C$": frozenset({"CAD"}),
    "AU$": frozenset({"AUD"}),
    "A$": frozenset({"AUD"}),
    "NZ$": frozenset({"NZD"}),
    "S$": frozenset({"SGD"}),
    "HK$": frozenset({"HKD"}),
    "SLR": frozenset({"LKR"}),
    "SLRs": frozenset({"LKR"}),
    # Ambiguous symbols/aliases resolved by template currency or input locale
    "$": frozenset({"USD", "CAD", "AUD", "NZD", "SGD", "HKD", "MXN", "TWD", "BND", "FJD"}),
    "Rs.": frozenset({"LKR", "INR", "PKR", "NPR", "MUR", "SCR"}),
    "Rs": frozenset({"LKR", "INR", "PKR", "NPR", "MUR", "SCR"}),
    "RS.": frozenset({"LKR", "INR", "PKR", "NPR", "MUR", "SCR"}),
    "RS": frozenset({"LKR", "INR", "PKR", "NPR", "MUR", "SCR"}),
    "¥": frozenset({"JPY", "CNY"}),
}

# Pre-sorted longest-first for greedy boundary matching
_SORTED_SYMBOLS: tuple[str, ...] = tuple(
    sorted(_SYMBOL_CANDIDATES.keys(), key=len, reverse=True)
)


@dataclass(frozen=True)
class LocaleSpec:
    """Resolved CLDR number/currency characteristics for an input locale."""

    locale_id: str | None
    decimal_symbol: str
    group_symbol: str
    territory_currencies: tuple[str, ...] = ()


@dataclass(frozen=True)
class ParsedAmount:
    """Validated Decimal amount and resolved ISO 4217 currency code (if any)."""

    amount: Decimal
    currency: str | None = None


@lru_cache(maxsize=64)
def resolve_locale_spec(locale_id: str | None) -> LocaleSpec:
    """Resolve CLDR decimal/group separators and territory currencies via Babel.

    Supports language-territory IDs not directly bundled as standalone CLDR
    locales (e.g. ``en_LK`` -> language ``en`` number format + territory ``LK``
    currency ``LKR``).
    """
    if locale_id is None or not str(locale_id).strip():
        return LocaleSpec(
            locale_id=None,
            decimal_symbol=".",
            group_symbol=",",
            territory_currencies=(),
        )

    cleaned = str(locale_id).strip().replace("-", "_")
    try:
        lang, territory, _script, _variant = parse_locale(cleaned)
    except ValueError as exc:
        raise PreprocessError(f"Unknown or invalid locale: {locale_id!r}") from exc

    if not lang:
        raise PreprocessError(f"Unknown or invalid locale: {locale_id!r}")

    if territory and (
        territory == "ZZ"
        or (
            territory not in Locale("en").territories
            and not get_territory_currencies(territory)
        )
    ):
        raise PreprocessError(f"Unknown locale territory in {locale_id!r}")

    try:
        loc = Locale.parse(cleaned)
    except UnknownLocaleError:
        try:
            loc = Locale.parse(lang)
        except UnknownLocaleError as exc:
            raise PreprocessError(f"Unknown locale: {locale_id!r}") from exc

    eff_territory = territory or loc.territory
    terr_currencies = (
        tuple(get_territory_currencies(eff_territory)) if eff_territory else ()
    )

    return LocaleSpec(
        locale_id=cleaned,
        decimal_symbol=get_decimal_symbol(loc),
        group_symbol=get_group_symbol(loc),
        territory_currencies=terr_currencies,
    )


def _match_leading_marker(text: str) -> tuple[str, frozenset[str], str] | None:
    """Try matching a recognized currency symbol or uppercase ISO 4217 code at the start."""
    for sym in _SORTED_SYMBOLS:
        if text.startswith(sym):
            rest = text[len(sym) :]
            # For alphabetic symbols (Rs, SLR), require non-letter boundary after symbol
            if sym[-1].isalpha() and rest and rest[0].isalpha():
                continue
            return sym, _SYMBOL_CANDIDATES[sym], rest

    if len(text) >= 3 and text[:3].isupper() and text[:3].isascii() and text[:3].isalpha():
        code = text[:3]
        if code in ISO_CURRENCIES:
            rest = text[3:]
            # Avoid matching 4+ letter uppercase words unless followed by space/digit/sign/decimal
            if rest and rest[0].isalpha():
                return None
            return code, frozenset({code}), rest
    return None


def _match_trailing_marker(text: str) -> tuple[str, frozenset[str], str] | None:
    """Try matching a recognized currency symbol or uppercase ISO 4217 code at the end."""
    for sym in _SORTED_SYMBOLS:
        if text.endswith(sym):
            head = text[: len(text) - len(sym)]
            if sym[0].isalpha() and head and head[-1].isalpha():
                continue
            return sym, _SYMBOL_CANDIDATES[sym], head

    if len(text) >= 3 and text[-3:].isupper() and text[-3:].isascii() and text[-3:].isalpha():
        code = text[-3:]
        if code in ISO_CURRENCIES:
            head = text[:-3]
            if head and head[-1].isalpha():
                return None
            return code, frozenset({code}), head
    return None


def _resolve_expected_currency(
    expected_currency: str | None,
) -> tuple[str | None, frozenset[str] | None]:
    """Return (canonical_iso_code, allowed_candidates) for the template's expected currency."""
    if expected_currency is None:
        return None, None
    cleaned = str(expected_currency).strip()
    if not cleaned:
        return None, None

    upper = cleaned.upper()
    if upper in ISO_CURRENCIES:
        return upper, frozenset({upper})
    if cleaned in _SYMBOL_CANDIDATES:
        cands = _SYMBOL_CANDIDATES[cleaned]
        canonical = next(iter(cands)) if len(cands) == 1 else None
        return canonical, cands
    # Custom/unrecognized label (treat literally if no conflict)
    return None, None


def _resolve_currency(
    marker: str | None,
    candidates: frozenset[str] | None,
    *,
    expected_currency: str | None,
    locale_spec: LocaleSpec,
    raw_value: Any,
) -> str | None:
    """Validate and resolve the input currency marker against expected currency and locale."""
    exp_iso, exp_candidates = _resolve_expected_currency(expected_currency)

    if marker is None or candidates is None:
        if exp_iso is not None:
            return exp_iso
        return None

    # 1. If template configured an expected ISO currency, input marker MUST match it.
    if exp_iso is not None:
        if exp_iso in candidates:
            return exp_iso
        found_label = next(iter(candidates)) if len(candidates) == 1 else marker
        raise PreprocessError(
            f"Currency mismatch in {raw_value!r}: input specifies '{found_label}', "
            f"but template expects '{exp_iso}'."
        )

    # 2. If template configured a symbol alias (e.g. legacy 'Rs.'), check overlap.
    if exp_candidates is not None:
        overlap = candidates & exp_candidates
        if not overlap:
            found_label = next(iter(candidates)) if len(candidates) == 1 else marker
            raise PreprocessError(
                f"Currency mismatch in {raw_value!r}: input specifies '{found_label}', "
                f"but template expects '{expected_currency}'."
            )
        if len(overlap) == 1:
            return next(iter(overlap))
        # Narrow with locale if available; otherwise accept legacy template symbol
        for terr_curr in locale_spec.territory_currencies:
            if terr_curr in overlap:
                return terr_curr
        return None

    # 3. No expected currency: unambiguous marker resolves directly.
    if len(candidates) == 1:
        return next(iter(candidates))

    # 4. Ambiguous marker (e.g. '$' or 'Rs.'): resolve via input locale territory if unique.
    locale_matches = [c for c in locale_spec.territory_currencies if c in candidates]
    if len(locale_matches) == 1:
        return locale_matches[0]

    raise PreprocessError(
        f"Ambiguous currency marker '{marker}' in {raw_value!r}: configure an expected "
        "currency or input locale to resolve it."
    )


def _validate_integer_grouping(int_part: str, group_sep: str, raw_value: Any) -> str:
    """Validate integer digits with optional locale grouping separators and return plain digits."""
    if not int_part:
        return ""

    if int_part.isascii() and int_part.isdigit():
        return int_part

    # For whitespace-grouped locales (e.g. fr_FR U+202F, fr_CA U+00A0), accept the
    # locale's grouping space as well as standard NBSP/NNBSP/space, while strictly
    # enforcing 3-digit grouping so "10 20" is always rejected.
    if group_sep.isspace():
        normalized = (
            int_part.replace("\u202f", " ")
            .replace("\u00a0", " ")
            .replace("\u2009", " ")
        )
        pattern = re.compile(r"^\d{1,3}(?: \d{3})+$")
        if not pattern.match(normalized):
            raise PreprocessError(f"Malformed number grouping in {raw_value!r}.")
        return normalized.replace(" ", "")

    esc_grp = re.escape(group_sep)
    standard_grp = re.compile(rf"^\d{{1,3}}(?:{esc_grp}\d{{3}})+$")
    if standard_grp.match(int_part):
        return int_part.replace(group_sep, "")

    # Also support South Asian lakh/crore grouping (e.g. 1,00,000) when group_sep is ','
    if group_sep == ",":
        south_asian_grp = re.compile(r"^\d{1,2}(?:,\d{2})*,\d{3}$")
        if south_asian_grp.match(int_part):
            return int_part.replace(",", "")

    raise PreprocessError(f"Malformed number grouping in {raw_value!r}.")


def _parse_localized_decimal(
    body: str,
    *,
    leading_sign: str,
    locale_spec: LocaleSpec,
    raw_value: Any,
) -> Decimal:
    """Validate the entire numeric body strictly under ``locale_spec`` and return ``Decimal``."""
    s = body.strip(_CURRENCY_SPACE_CHARS)
    if not s:
        raise PreprocessError(f"Missing numeric amount in {raw_value!r}.")

    sign = leading_sign
    if s[0] in "+-":
        if sign:
            raise PreprocessError(f"Duplicate sign in numeric value {raw_value!r}.")
        sign = s[0]
        s = s[1:]
        if not s:
            raise PreprocessError(f"Invalid numeric value {raw_value!r}.")

    dec_sep = locale_spec.decimal_symbol
    grp_sep = locale_spec.group_symbol

    # Scientific notation (must be complete and require at least one exponent digit)
    if "e" in s or "E" in s:
        esc_dec = re.escape(dec_sep)
        sci_pattern = re.compile(rf"^(?:\d+(?:{esc_dec}\d+)?|{esc_dec}\d+)[eE][+-]?\d+$")
        if not sci_pattern.match(s):
            raise PreprocessError(
                f"Invalid or incomplete scientific notation in {raw_value!r}."
            )
        canonical = sign + s.replace(dec_sep, ".")
        try:
            result = Decimal(canonical)
        except InvalidOperation as exc:
            raise PreprocessError(f"Invalid numeric value {raw_value!r}: {exc}") from exc
        if not result.is_finite():
            raise PreprocessError(f"Non-finite numeric value: {raw_value!r}")
        return result

    # Standard decimal / grouped integer
    if dec_sep in s:
        parts = s.split(dec_sep)
        if len(parts) != 2:
            raise PreprocessError(f"Malformed decimal number in {raw_value!r}.")
        int_raw, frac_raw = parts
        if not frac_raw or not (frac_raw.isascii() and frac_raw.isdigit()):
            raise PreprocessError(f"Malformed fractional digits in {raw_value!r}.")
        int_digits = _validate_integer_grouping(int_raw, grp_sep, raw_value)
        if not int_digits:
            int_digits = "0"
        canonical = f"{sign}{int_digits}.{frac_raw}"
    else:
        int_digits = _validate_integer_grouping(s, grp_sep, raw_value)
        if not int_digits:
            raise PreprocessError(f"Invalid numeric value {raw_value!r}.")
        canonical = f"{sign}{int_digits}"

    try:
        result = Decimal(canonical)
    except InvalidOperation as exc:
        raise PreprocessError(f"Invalid numeric value {raw_value!r}: {exc}") from exc

    if not result.is_finite():
        raise PreprocessError(f"Non-finite numeric value: {raw_value!r}")
    return result


def parse_currency_amount(
    value: Any,
    *,
    expected_currency: str | None = None,
    input_locale: str | None = None,
) -> ParsedAmount:
    """Parse and validate a numeric or currency string into an exact ``Decimal`` amount.

    Recognizes ISO 4217 codes and common currency symbols on either side of the
    amount (with or without standard/NBSP/narrow-NBSP whitespace), resolves
    locale-specific decimal and grouping separators via Babel/CLDR, and validates
    the entire remaining amount as a ``Decimal``.
    """
    locale_spec = resolve_locale_spec(input_locale)
    exp_iso, _ = _resolve_expected_currency(expected_currency)

    if isinstance(value, bool) or value is None:
        raise PreprocessError(f"Invalid numeric value: {value!r}")

    if isinstance(value, Decimal):
        if not value.is_finite():
            raise PreprocessError(f"Non-finite numeric value: {value!r}")
        return ParsedAmount(amount=value, currency=exp_iso)

    if isinstance(value, int):
        return ParsedAmount(amount=Decimal(value), currency=exp_iso)

    if isinstance(value, float):
        try:
            dec = Decimal(str(value))
        except InvalidOperation as exc:
            raise PreprocessError(f"Non-finite numeric value: {value!r}") from exc
        if not dec.is_finite():
            raise PreprocessError(f"Non-finite numeric value: {value!r}")
        return ParsedAmount(amount=dec, currency=exp_iso)

    text = str(value).strip(_CURRENCY_SPACE_CHARS)
    if not text:
        raise PreprocessError("Empty string is not a valid numeric value.")

    # Optional Sri Lankan trailing terminator '/=' or '/-' (e.g. '100/=', 'Rs. 100/=')
    if text.endswith("/=") or text.endswith("/-"):
        text = text[:-2].rstrip(_CURRENCY_SPACE_CHARS)
        if not text:
            raise PreprocessError(f"Invalid numeric value {value!r}.")

    # Optional leading sign placed before a leading currency marker (e.g. '-$100', '-USD 100')
    outer_sign = ""
    work = text
    if work[0] in "+-":
        # Check if sign precedes a currency marker rather than directly preceding digits
        after_sign = work[1:].lstrip(_CURRENCY_SPACE_CHARS)
        if _match_leading_marker(after_sign) is not None:
            outer_sign = work[0]
            work = after_sign

    lead_match = _match_leading_marker(work)
    marker: str | None = None
    candidates: frozenset[str] | None = None
    body = work

    if lead_match is not None:
        marker, candidates, body = lead_match
        body = body.strip(_CURRENCY_SPACE_CHARS)
        # Reject conflicting/duplicate trailing currency marker (e.g. 'USD 100 EUR')
        if _match_trailing_marker(body) is not None:
            raise PreprocessError(
                f"Multiple currency markers found in {value!r}."
            )
    else:
        trail_match = _match_trailing_marker(work)
        if trail_match is not None:
            marker, candidates, body = trail_match
            body = body.strip(_CURRENCY_SPACE_CHARS)

    resolved_currency = _resolve_currency(
        marker,
        candidates,
        expected_currency=expected_currency,
        locale_spec=locale_spec,
        raw_value=value,
    )

    amount = _parse_localized_decimal(
        body,
        leading_sign=outer_sign,
        locale_spec=locale_spec,
        raw_value=value,
    )

    return ParsedAmount(amount=amount, currency=resolved_currency)
