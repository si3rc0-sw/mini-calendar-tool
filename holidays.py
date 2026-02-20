"""Holiday definitions for Switzerland, Germany, and China."""

from __future__ import annotations

from datetime import date, timedelta


def _easter(year: int) -> date:
    """Compute Easter Sunday (Anonymous Gregorian algorithm)."""
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    el = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * el) // 451
    month, day = divmod(h + el - 7 * m + 114, 31)
    return date(year, month, day + 1)


# --- date generators --------------------------------------------------------

def _fixed(m: int, d: int):
    return lambda year: [date(year, m, d)]


def _fixed_range(m: int, d: int, n: int):
    return lambda year: [date(year, m, d) + timedelta(days=i) for i in range(n)]


def _easter_rel(offset: int):
    return lambda year: [_easter(year) + timedelta(days=offset)]


def _from_table(table: dict[int, tuple[int, int]]):
    def fn(year):
        entry = table.get(year)
        return [date(year, entry[0], entry[1])] if entry else []
    return fn


def _from_table_multi(table: dict[int, tuple[int, int]], n: int):
    def fn(year):
        entry = table.get(year)
        if not entry:
            return []
        start = date(year, entry[0], entry[1])
        return [start + timedelta(days=i) for i in range(n)]
    return fn


# --- Chinese lunar calendar lookup tables (2024-2036) -----------------------

_SPRING_FESTIVAL = {
    2024: (2, 10), 2025: (1, 29), 2026: (2, 17), 2027: (2, 6),
    2028: (1, 26), 2029: (2, 13), 2030: (2, 3),  2031: (1, 23),
    2032: (2, 11), 2033: (1, 31), 2034: (2, 19), 2035: (2, 8),
    2036: (1, 28),
}

_QINGMING = {
    2024: (4, 4), 2025: (4, 4), 2026: (4, 5), 2027: (4, 5),
    2028: (4, 4), 2029: (4, 4), 2030: (4, 5), 2031: (4, 5),
    2032: (4, 4), 2033: (4, 4), 2034: (4, 5), 2035: (4, 5),
    2036: (4, 4),
}

_DRAGON_BOAT = {
    2024: (6, 10), 2025: (5, 31), 2026: (6, 19), 2027: (6, 9),
    2028: (5, 28), 2029: (6, 16), 2030: (6, 5),  2031: (6, 24),
    2032: (6, 13), 2033: (6, 2),  2034: (6, 22), 2035: (6, 11),
    2036: (5, 31),
}

_MID_AUTUMN = {
    2024: (9, 17), 2025: (10, 6), 2026: (9, 25), 2027: (9, 15),
    2028: (10, 3), 2029: (9, 22), 2030: (9, 12), 2031: (10, 1),
    2032: (9, 19), 2033: (9, 8),  2034: (9, 28), 2035: (9, 16),
    2036: (9, 5),
}


# --- Holiday registry: (key, name, country, dates_fn) -----------------------

HOLIDAYS: list[tuple] = [
    # Switzerland
    ("ch_neujahr",        "Neujahr",        "CH", _fixed(1, 1)),
    ("ch_berchtoldstag",  "Berchtoldstag",  "CH", _fixed(1, 2)),
    ("ch_karfreitag",     "Karfreitag",     "CH", _easter_rel(-2)),
    ("ch_ostermontag",    "Ostermontag",    "CH", _easter_rel(1)),
    ("ch_tag_der_arbeit", "Tag der Arbeit", "CH", _fixed(5, 1)),
    ("ch_auffahrt",       "Auffahrt",       "CH", _easter_rel(39)),
    ("ch_pfingstmontag",  "Pfingstmontag",  "CH", _easter_rel(49)),
    ("ch_bundesfeier",    "Bundesfeier",    "CH", _fixed(8, 1)),
    ("ch_weihnachten",    "Weihnachten",    "CH", _fixed(12, 25)),
    # Germany
    ("de_neujahr",            "Neujahr",                   "DE", _fixed(1, 1)),
    ("de_karfreitag",         "Karfreitag",                "DE", _easter_rel(-2)),
    ("de_ostermontag",        "Ostermontag",               "DE", _easter_rel(1)),
    ("de_tag_der_arbeit",     "Tag der Arbeit",            "DE", _fixed(5, 1)),
    ("de_christi_himmelfahrt", "Christi Himmelfahrt",       "DE", _easter_rel(39)),
    ("de_pfingstmontag",      "Pfingstmontag",             "DE", _easter_rel(49)),
    ("de_tag_dt_einheit",     "Tag der Deutschen Einheit", "DE", _fixed(10, 3)),
    ("de_weihnachten1",       "1. Weihnachtstag",          "DE", _fixed(12, 25)),
    ("de_weihnachten2",       "2. Weihnachtstag",          "DE", _fixed(12, 26)),
    # China
    ("cn_neujahr",         "New Year's Day",       "CN", _fixed(1, 1)),
    ("cn_spring_festival", "Spring Festival",      "CN", _from_table_multi(_SPRING_FESTIVAL, 3)),
    ("cn_qingming",        "Qingming Festival",    "CN", _from_table(_QINGMING)),
    ("cn_labour_day",      "Labour Day",           "CN", _fixed(5, 1)),
    ("cn_dragon_boat",     "Dragon Boat Festival", "CN", _from_table(_DRAGON_BOAT)),
    ("cn_mid_autumn",      "Mid-Autumn Festival",  "CN", _from_table(_MID_AUTUMN)),
    ("cn_national_day",    "National Day",         "CN", _fixed_range(10, 1, 3)),
]

_BY_KEY = {h[0]: h for h in HOLIDAYS}

# Country list for UI grouping
COUNTRIES: list[tuple[str, str]] = [
    ("CH", "Switzerland"),
    ("DE", "Germany"),
    ("CN", "China"),
]


def holidays_by_country(country: str) -> list[tuple[str, str]]:
    """Return [(key, name), ...] for the given country code."""
    return [(h[0], h[1]) for h in HOLIDAYS if h[2] == country]


def holidays_for_year(
    year: int, enabled_keys: set[str],
) -> dict[date, list[tuple[str, str]]]:
    """Return {date: [(name, country), ...]} for all enabled holidays in a year."""
    result: dict[date, list[tuple[str, str]]] = {}
    for key in enabled_keys:
        entry = _BY_KEY.get(key)
        if entry is None:
            continue
        _key, name, country, dates_fn = entry
        for d in dates_fn(year):
            result.setdefault(d, []).append((name, country))
    return result
