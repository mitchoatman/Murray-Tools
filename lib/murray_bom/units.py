# -*- coding: utf-8 -*-
"""Length parsing and formatting. Port of the helpers in ExportXLS.js (parseInches, toFeetInches,
toFractionalInches) and exportTypes.js (SIZE_NUMBERS). Same rounding rules and result fields."""
import re
from math import gcd

FRACTION_DENOMINATOR = 16


def strip_id(s):
    """'PIPE CI-NH [123456]' -> 'PIPE CI-NH'"""
    return re.sub(r'\s*\[\d+\]\s*$', '', '' if s is None else str(s)).strip()


def _inch_part(text):
    q = re.sub(r'"\s*$', '', text)
    q = re.sub(r'\s*(?:inches|inch|in)\s*$', '', q, flags=re.I).strip()
    if q == '':
        return 0.0
    m = re.match(r'^(\d+)[\s-]+(\d+)/(\d+)$', q)
    if m:
        return (int(m.group(1)) + int(m.group(2)) / int(m.group(3))) if int(m.group(3)) else None
    m = re.match(r'^(\d+)/(\d+)$', q)
    if m:
        return (int(m.group(1)) / int(m.group(2))) if int(m.group(2)) else None
    m = re.match(r'^(\d+(?:\.\d+)?|\.\d+)$', q)
    return float(m.group(1)) if m else None


def parse_inches(value):
    """Returns (inches, assumed, tol) or None. assumed: no unit given, read as inches.
    tol: how far (inches) the value may be from the true figure because few decimals were given."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return (float(value), True, 0.0)
    t = '' if value is None else str(value).strip()
    t = re.sub('[\u2033\u201c\u201d]', '"', t)
    t = re.sub('[\u2032\u2018\u2019]', "'", t)
    if not t:
        return None

    m = re.match(r"^(-?)\s*(\d+(?:\.\d+)?)\s*'(?:\s*-?\s*(.+))?$", t)
    if m:
        rest = _inch_part(m.group(3)) if m.group(3) else 0.0
        if rest is None:
            return None
        sign = -1 if m.group(1) else 1
        return (sign * (float(m.group(2)) * 12 + rest), False, 0.0)

    m = re.match(r'^(-?)\s*(.*?)\s*([a-z"][a-z"\-]*)?$', t, re.I)
    if not m:
        return None
    core = m.group(2)
    v = _inch_part(core)
    if v is None:
        return None
    unit = (m.group(3) or '').lower()
    if unit == '':
        factor = 1.0
    elif re.match(r'^(ft|feet|foot)(-|$)', unit):
        factor = 12.0
    elif unit == '"' or re.match(r'^(in|inch|inches|fractional-in|fractional-inch|fractional-inches)$', unit):
        factor = 1.0
    elif unit == 'mm':
        factor = 1 / 25.4
    elif unit == 'cm':
        factor = 1 / 2.54
    elif unit == 'm':
        factor = 100 / 2.54
    else:
        return None
    dec = re.match(r'^\d*\.(\d+)$', core)
    tol = 0.5 * 10 ** (-len(dec.group(1))) * factor if dec else 0.0
    sign = -1 if m.group(1) else 1
    return (sign * v * factor, unit == '', tol)


def _result(text, rounded=False, assumed=False, ok=True):
    return {'text': text, 'rounded': rounded, 'assumed': assumed, 'ok': ok}


def _round_units(inches, tol, denom):
    a = abs(inches)
    units = int(a * denom + 0.5)           # JS Math.round on positive numbers
    rounded = abs(a - units / denom) > tol + 1e-9
    return units, rounded


def to_feet_inches(value, denom=None):
    """Value -> text like 1' - 8 1/2" (as Revit shows it)."""
    denom = denom or FRACTION_DENOMINATOR
    if value is None or value == '':
        return _result('')
    p = parse_inches(value)
    if not p:
        return _result(str(value), ok=False)
    inches, assumed, tol = p
    units, rounded = _round_units(inches, tol, denom)
    per_foot = 12 * denom
    feet = units // per_foot
    rem = units - feet * per_foot
    whole = rem // denom
    num = rem % denom
    den = denom
    if num:
        g = gcd(num, den)
        num //= g
        den //= g
    ins = str(whole) if whole else ''
    if num:
        ins += (' ' if ins else '') + '%d/%d' % (num, den)
    if not ins:
        ins = '0'
    text = "%d' - %s\"" % (feet, ins)
    if inches < 0 and units > 0:
        text = '-' + text
    return _result(text, rounded, assumed)


def to_fractional_inches(value, denom=None):
    """Value -> text like 4 9/16"."""
    denom = denom or FRACTION_DENOMINATOR
    if value is None or value == '':
        return _result('')
    p = parse_inches(value)
    if not p:
        return _result(str(value), ok=False)
    inches, assumed, tol = p
    units, rounded = _round_units(inches, tol, denom)
    whole = units // denom
    num = units % denom
    den = denom
    if num:
        g = gcd(num, den)
        num //= g
        den //= g
    if whole and num:
        text = '%d %d/%d"' % (whole, num, den)
    elif whole:
        text = '%d"' % whole
    elif num:
        text = '%d/%d"' % (num, den)
    else:
        text = '0"'
    if inches < 0 and units > 0:
        text = '-' + text
    return _result(text, rounded, assumed)


def size_numbers(size):
    """Numbers in a size such as 1/2", 1 1/4" or 2" x 1", for sorting small to big. No number sorts last."""
    t = '' if size is None else str(size)
    out = []
    for m in re.finditer(r'(\d+)\s+(\d+)/(\d+)|(\d+)/(\d+)|(\d+(?:\.\d+)?)', t):
        if m.group(1) is not None:
            d = int(m.group(3))
            out.append(int(m.group(1)) + int(m.group(2)) / d if d else int(m.group(1)))
        elif m.group(4) is not None:
            d = int(m.group(5))
            out.append(int(m.group(4)) / d if d else 0)
        else:
            out.append(float(m.group(6)))
    return out or [1e9]


def natural_key(s):
    """Natural sort key (numbers by value, case-insensitive), like localeCompare numeric."""
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r'(\d+)', '' if s is None else str(s))]
