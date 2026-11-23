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
    get_currency_symbol,
    get_decimal_symbol,
    get_group_symbol,
    get_territory_currencies,
)

from pdf_bulk_fill.errors import PreprocessError

# Space characters allowed between currency markers and amounts
_CURRENCY_SPACE_CHARS = " \t\u00a0\u202f\u2007\u2009"

# Valid ISO 4217 currency codes from Unicode CLDR via Babel
ISO_CURRENCIES: frozenset[str] = frozenset(Locale("en").currencies.keys())

# Base currency aliases / symbols (merged dynamically with Babel CLDR locale symbols).
_BASE_SYMBOL_CANDIDATES: dict[str, frozenset[str]] = {
    # Unambiguous symbols & prefixed symbols
    "€": frozenset({"EUR"}),
    "£": frozenset({"GBP"}),
    "₹": frozenset({"INR"}),
    "₩": frozenset({"KRW"}),
    "zł": frozenset({"PLN"}),
    "R$": frozenset({"BRL"}),
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
    "kr": frozenset({"SEK", "NOK", "DKK", "ISK"}),
}


@lru_cache(maxsize=1)
def _build_global_cldr_symbol_map() -> dict[str, frozenset[str]]:
    """Build a symbol -> ISO candidate set map combining CLDR data with common aliases."""
    merged: dict[str, set[str]] = {
        sym: set(codes) for sym, codes in _BASE_SYMBOL_CANDIDATES.items()
    }
    en_locale = Locale("en")
    for code, sym in en_locale.currency_symbols.items():
        if sym and code in ISO_CURRENCIES and sym != code:
            merged.setdefault(sym, set()).add(code)
    return {sym: frozenset(codes) for sym, codes in merged.items()}


@dataclass(frozen=True)
class LocaleSpec:
    """Resolved CLDR number/currency characteristics for an input locale."""

    locale_id: str | None
    decimal_symbol: str
    group_symbol: str
    grouping: tuple[int, int] = (3, 3)
    territory_currencies: tuple[str, ...] = ()
    babel_locale_name: str | None = None


@dataclass(frozen=True)
class ParsedAmount:
    """Validated Decimal amount and resolved ISO 4217 currency code (if any)."""

    amount: Decimal
    currency: str | None = None


@lru_cache(maxsize=64)
def resolve_locale_spec(locale_id: str | None) -> LocaleSpec:
    """Resolve CLDR decimal/group separators, grouping sizes, and territory currencies via Babel.

    Supports language-territory IDs not directly bundled as standalone CLDR
    locales (e.g. ``en_LK`` -> language ``en`` number format + territory ``LK``
    currency ``LKR``).
    """
    if locale_id is None or not str(locale_id).strip():
        return LocaleSpec(
            locale_id=None,
            decimal_symbol=".",
            group_symbol=",",
            grouping=(3, 3),
            territory_currencies=(),
            babel_locale_name=None,
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

    fmt = loc.decimal_formats.get(None)
    raw_grouping = getattr(fmt, "grouping", (3, 3)) if fmt is not None else (3, 3)
    grouping = (int(raw_grouping[0]), int(raw_grouping[1]))

    return LocaleSpec(
        locale_id=cleaned,
        decimal_symbol=get_decimal_symbol(loc),
        group_symbol=get_group_symbol(loc),
        grouping=grouping,
        territory_currencies=terr_currencies,
        babel_locale_name=str(loc),
    )


def _effective_symbol_map(
    locale_spec: LocaleSpec,
    expected_iso: str | None = None,
) -> dict[str, frozenset[str]]:
    """Build symbol lookup table including locale-specific CLDR symbols from Babel."""
    merged: dict[str, set[str]] = {
        sym: set(codes) for sym, codes in _build_global_cldr_symbol_map().items()
    }
    if locale_spec.babel_locale_name:
        loc = Locale.parse(locale_spec.babel_locale_name)
        for code, sym in loc.currency_symbols.items():
            if sym and code in ISO_CURRENCIES and sym != code:
                merged.setdefault(sym, set()).add(code)
        for code in (*locale_spec.territory_currencies, *([expected_iso] if expected_iso else [])):
            if code and code in ISO_CURRENCIES:
                sym = get_currency_symbol(code, locale=loc)
                if sym and sym != code:
                    merged.setdefault(sym, set()).add(code)
    elif expected_iso and expected_iso in ISO_CURRENCIES:
        sym = get_currency_symbol(expected_iso, locale="en")
        if sym and sym != expected_iso:
            merged.setdefault(sym, set()).add(expected_iso)

    return {sym: frozenset(codes) for sym, codes in merged.items()}


def validate_configured_currency(
    expected_currency: str | None,
    *,
    input_locale: str | None = None,
) -> str | None:
    """Validate and resolve a configured currency (ISO 4217 code or resolvable symbol).

    Raises:
        PreprocessError: If ``expected_currency`` is unknown (e.g., ``USX``) or an
            ambiguous symbol (e.g., ``$``, ``Rs.``) with no locale that resolves it.
    """
    locale_spec = resolve_locale_spec(input_locale)
    if expected_currency is None:
        return None
    cleaned = str(expected_currency).strip()
    if not cleaned:
        return None

    upper = cleaned.upper()
    if upper in ISO_CURRENCIES:
        return upper

    symbol_map = _effective_symbol_map(locale_spec)
    if cleaned in symbol_map:
        cands = symbol_map[cleaned]
        if len(cands) == 1:
            return next(iter(cands))
        locale_matches = [c for c in locale_spec.territory_currencies if c in cands]
        if len(locale_matches) == 1:
            return locale_matches[0]
        raise PreprocessError(
            f"Ambiguous configured currency symbol {cleaned!r}: specify an ISO 4217 "
            "currency code (e.g. USD, CAD, LKR) or an input locale that resolves it."
        )

    raise PreprocessError(
        f"Unknown configured currency {cleaned!r}: expected a valid ISO 4217 "
        "currency code or recognized currency symbol."
    )


def _match_leading_marker(
    text: str,
    symbol_map: dict[str, frozenset[str]],
) -> tuple[str, frozenset[str], str] | None:
    """Try matching a recognized currency symbol or uppercase ISO 4217 code at the start."""
    for sym in sorted(symbol_map.keys(), key=len, reverse=True):
        if text.startswith(sym):
            rest = text[len(sym) :]
            if sym[-1].isalpha() and rest and rest[0].isalpha():
                continue
            return sym, symbol_map[sym], rest

    if len(text) >= 3 and text[:3].isupper() and text[:3].isascii() and text[:3].isalpha():
        code = text[:3]
        if code in ISO_CURRENCIES:
            rest = text[3:]
            if rest and rest[0].isalpha():
                return None
            return code, frozenset({code}), rest
    return None


def _match_trailing_marker(
    text: str,
    symbol_map: dict[str, frozenset[str]],
) -> tuple[str, frozenset[str], str] | None:
    """Try matching a recognized currency symbol or uppercase ISO 4217 code at the end."""
    for sym in sorted(symbol_map.keys(), key=len, reverse=True):
        if text.endswith(sym):
            head = text[: len(text) - len(sym)]
            if sym[0].isalpha() and head and head[-1].isalpha():
                continue
            return sym, symbol_map[sym], head

    if len(text) >= 3 and text[-3:].isupper() and text[-3:].isascii() and text[-3:].isalpha():
        code = text[-3:]
        if code in ISO_CURRENCIES:
            head = text[:-3]
            if head and head[-1].isalpha():
                return None
            return code, frozenset({code}), head
    return None


def _resolve_currency(
    marker: str | None,
    candidates: frozenset[str] | None,
    *,
    expected_iso: str | None,
    locale_spec: LocaleSpec,
    raw_value: Any,
) -> str | None:
    """Validate and resolve the input currency marker against expected ISO currency and locale."""
    if marker is None or candidates is None:
        return expected_iso

    # 1. If template configured an expected ISO currency, input marker MUST match it.
    if expected_iso is not None:
        if expected_iso in candidates:
            return expected_iso
        found_label = next(iter(candidates)) if len(candidates) == 1 else marker
        raise PreprocessError(
            f"Currency mismatch in {raw_value!r}: input specifies '{found_label}', "
            f"but template expects '{expected_iso}'."
        )

    # 2. No expected currency: unambiguous marker resolves directly.
    if len(candidates) == 1:
        return next(iter(candidates))

    # 3. Ambiguous marker (e.g. '$' or 'Rs.'): resolve via input locale territory if unique.
    locale_matches = [c for c in locale_spec.territory_currencies if c in candidates]
    if len(locale_matches) == 1:
        return locale_matches[0]

    raise PreprocessError(
        f"Ambiguous currency marker '{marker}' in {raw_value!r}: configure an expected "
        "currency or input locale to resolve it."
    )


def _validate_integer_grouping(
    int_part: str,
    group_sep: str,
    grouping: tuple[int, int],
    raw_value: Any,
) -> str:
    """Validate integer digits strictly against the locale's CLDR grouping pattern."""
    if not int_part:
        return ""

    if int_part.isascii() and int_part.isdigit():
        return int_part

    primary, secondary = grouping

    if group_sep.isspace():
        normalized = (
            int_part.replace("\u202f", " ")
            .replace("\u00a0", " ")
            .replace("\u2009", " ")
        )
        esc_grp = " "
        target = normalized
    else:
        esc_grp = re.escape(group_sep)
        target = int_part

    if primary == secondary:
        pattern = re.compile(rf"^\d{{1,{primary}}}(?:{esc_grp}\d{{{primary}}})+$")
    else:
        # e.g. en_IN (3, 2): 1,234 | 12,345 | 1,00,000 | 12,34,567
        pattern = re.compile(
            rf"^\d{{1,{secondary}}}(?:{esc_grp}\d{{{secondary}}})*{esc_grp}\d{{{primary}}}$"
        )

    if not pattern.match(target):
        raise PreprocessError(f"Malformed number grouping in {raw_value!r}.")

    return target.replace(" " if group_sep.isspace() else group_sep, "")


def _parse_localized_decimal(
    body: str,
    *,
    leading_sign: str,
    locale_spec: LocaleSpec,
    raw_value: Any,
) -> Decimal:
    """Validate the entire numeric body strictly under ``locale_spec`` and return ``Decimal``."""
    s = body.strip(_CURRENCY_SPACE_CHARS).replace("\u2212", "-")
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
        int_digits = _validate_integer_grouping(
            int_raw, grp_sep, locale_spec.grouping, raw_value
        )
        if not int_digits:
            int_digits = "0"
        canonical = f"{sign}{int_digits}.{frac_raw}"
    else:
        int_digits = _validate_integer_grouping(
            s, grp_sep, locale_spec.grouping, raw_value
        )
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
    """Parse and validate a numeric or currency value into an exact ``Decimal`` amount.

    Native numeric inputs (``int``, ``float``, ``Decimal``) preserve their numeric
    value directly without locale string re-parsing. String inputs recognize ISO 4217
    codes and CLDR currency symbols on either side (with or without spaces including
    NBSP/narrow-NBSP), resolve locale-specific decimal & grouping rules via Babel,
    and validate the entire remaining amount as ``Decimal``.
    """
    locale_spec = resolve_locale_spec(input_locale)
    expected_iso = validate_configured_currency(
        expected_currency, input_locale=input_locale
    )

    if isinstance(value, bool) or value is None:
        raise PreprocessError(f"Invalid numeric value: {value!r}")

    if isinstance(value, Decimal):
        if not value.is_finite():
            raise PreprocessError(f"Non-finite numeric value: {value!r}")
        return ParsedAmount(amount=value, currency=expected_iso)

    if isinstance(value, int):
        return ParsedAmount(amount=Decimal(value), currency=expected_iso)

    if isinstance(value, float):
        try:
            dec = Decimal(str(value))
        except InvalidOperation as exc:
            raise PreprocessError(f"Non-finite numeric value: {value!r}") from exc
        if not dec.is_finite():
            raise PreprocessError(f"Non-finite numeric value: {value!r}")
        return ParsedAmount(amount=dec, currency=expected_iso)

    text = str(value).strip(_CURRENCY_SPACE_CHARS).replace("\u2212", "-")
    if not text:
        raise PreprocessError("Empty string is not a valid numeric value.")

    # Optional Sri Lankan trailing terminator '/=' or '/-' (e.g. '100/=', 'Rs. 100/=')
    if text.endswith("/=") or text.endswith("/-"):
        text = text[:-2].rstrip(_CURRENCY_SPACE_CHARS)
        if not text:
            raise PreprocessError(f"Invalid numeric value {value!r}.")

    symbol_map = _effective_symbol_map(locale_spec, expected_iso=expected_iso)

    # Optional leading sign placed before a leading currency marker (e.g. '-$100', '-R$ 100')
    outer_sign = ""
    work = text
    if work[0] in "+-":
        after_sign = work[1:].lstrip(_CURRENCY_SPACE_CHARS)
        if _match_leading_marker(after_sign, symbol_map) is not None:
            outer_sign = work[0]
            work = after_sign

    lead_match = _match_leading_marker(work, symbol_map)
    marker: str | None = None
    candidates: frozenset[str] | None = None
    body = work

    if lead_match is not None:
        marker, candidates, body = lead_match
        body = body.strip(_CURRENCY_SPACE_CHARS)
        if _match_trailing_marker(body, symbol_map) is not None:
            raise PreprocessError(f"Multiple currency markers found in {value!r}.")
    else:
        trail_match = _match_trailing_marker(work, symbol_map)
        if trail_match is not None:
            marker, candidates, body = trail_match
            body = body.strip(_CURRENCY_SPACE_CHARS)

    resolved_currency = _resolve_currency(
        marker,
        candidates,
        expected_iso=expected_iso,
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
