"""relist.py: enzyme names and the Buffer column."""

from __future__ import annotations

import pytest
import relist


@pytest.mark.parametrize(("raw", "clean"), [
    ("BbsI-HF®", "BbsI-HF"),
    ("BsaI-HF®v2", "BsaI-HFv2"),
    ("  EcoRI-HF ", "EcoRI-HF"),
    ("Foo™", "Foo"),
    ("XhoI", "XhoI"),
    (None, ""),
])
def test_normalize_enzyme_name(raw, clean):
    assert relist.normalize_enzyme_name(raw) == clean


@pytest.mark.parametrize(("raw", "value"), [
    ("100%", 100), ("50%", 50), ("<10%", 5), (" 25 % ", 25), (None, None), ("star", None),
    (75, 75.0),
])
def test_parse_activity(raw, value):
    assert relist.parse_activity(raw) == value


def acts(r11=None, r21=None, r31=None, cut=None):
    return {"r1.1": r11, "r2.1": r21, "r3.1": r31, "CutSmart": cut}


def test_cutsmart_wins_when_it_gives_full_activity():
    assert relist.derive_buffer(acts("100%", "100%", "100%", "100%")) == "CutSmart"


def test_full_activity_buffers_are_taken_in_order():
    assert relist.derive_buffer(acts("100%", "100%", "50%", "25%")) == "r1.1"
    assert relist.derive_buffer(acts("<10%", "100%", "100%", "25%")) == "r2.1"
    assert relist.derive_buffer(acts("<10%", "50%", "100%", "25%")) == "r3.1"


def test_without_full_activity_the_best_buffer_is_flagged():
    assert relist.derive_buffer(acts("25%", "75%", "50%", "10%")) == "r2.1 (check buffer)"
    # a tie goes to the buffer that comes first in the order CutSmart, r1.1, r2.1, r3.1
    assert relist.derive_buffer(acts("50%", "50%", "<10%", "<10%")) == "r1.1 (check buffer)"


def test_no_data_at_all():
    assert relist.derive_buffer(acts()) == "check buffer"
