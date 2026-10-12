"""Helpers for building an RElist table: enzyme names and the 10xCutSmart buffer.

The RElist sheet of the HiFi template has four columns (Enzyme, Buffer, Reaction
temp., Heat inactivation). These two functions normalise an enzyme name and
derive the Buffer column from a table of percent activities. They read no file
and use no network.
"""

from __future__ import annotations

import re

# Buffers in the order they are preferred when several give 100% activity.
BUFFER_ORDER = ("CutSmart", "r1.1", "r2.1", "r3.1")

# A "<10%" activity is below the lowest figure the vendor tables give.
BELOW_TABLE = 5

_PERCENT = re.compile(r"(\d+(?:\.\d+)?)\s*%")


def normalize_enzyme_name(name: object) -> str:
    """Return the enzyme name as a design record writes it.

    Registered-trademark and trademark signs are removed and the ends are
    trimmed, so ``BbsI-HF(R)`` and ``BsaI-HF(R)v2`` become ``BbsI-HF`` and
    ``BsaI-HFv2``. Anything else is left as it is.
    """
    text = "" if name is None else str(name)
    for mark in ("®", "™"):
        text = text.replace(mark, "")
    return text.strip()


def parse_activity(value: object) -> float | None:
    """Percent activity from a table cell: ``100%`` -> 100, ``<10%`` -> 5, blank -> None."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip()
    match = _PERCENT.search(text)
    if not match:
        return None
    number = float(match.group(1))
    if text.startswith("<"):
        return float(BELOW_TABLE) if number <= 10 else number / 2
    return number


def derive_buffer(activities: dict[str, object]) -> str:
    """The buffer to name in the Buffer column.

    ``activities`` maps ``CutSmart``, ``r1.1``, ``r2.1`` and ``r3.1`` to the
    percent activity written in the vendor table (for example ``"100%"``).
    CutSmart wins when it gives 100%. Otherwise the first buffer at 100% in the
    order r1.1, r2.1, r3.1. When none reaches 100% the most active one is
    returned with ``check buffer`` added, because the 10xCutSmart reaction of
    the RE sheet is then not the right one.
    """
    parsed = {name: parse_activity(activities.get(name)) for name in BUFFER_ORDER}
    for name in BUFFER_ORDER:
        if parsed[name] is not None and parsed[name] >= 100:
            return name
    known = [(parsed[name], -i, name) for i, name in enumerate(BUFFER_ORDER)
             if parsed[name] is not None]
    if not known:
        return "check buffer"
    best = max(known)[2]
    return f"{best} (check buffer)"
