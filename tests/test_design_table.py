"""design_table.py: reading blocks and CSV files, numbers, --sub."""

from __future__ import annotations

import design_table as dt
import pytest
from conftest import CONFIG, HEADINGS, NAMES, hifi_row, write_csv, write_project_sheet


def test_block_is_found_and_rows_are_read(project_xlsx):
    blocks = dt.read_blocks(project_xlsx, "Projects", CONFIG)
    assert len(blocks) == 1
    block = blocks[0]
    assert (block.header, block.first, block.last) == (1, 2, 5)
    assert block.project == "P001"
    assert [r.sub for r in block.rows] == ["1", "2", "3", "3"]
    assert [r.own for r in block.rows] == [True, True, True, False]
    assert block.rows[0].values["re1"] == "EcoRI-HF"
    assert block.rows[3].values["oligo1"] == "O-7"


def test_two_blocks_are_split_by_the_blank_row(tmp_path, rows):
    other = [hifi_row(1, "q-one", "EcoRI-HF", "", "bb", 100, "O-1", "O-2", "t", 10,
                      project="P002")]
    path = write_project_sheet(tmp_path / "p.xlsx", [rows, other])
    blocks = dt.read_blocks(path, "Projects", CONFIG)
    assert [b.project for b in blocks] == ["P001", "P002"]
    assert dt.pick_block(blocks, "p002", "Plasmid Construction", "Projects").project == "P002"


def test_project_is_required_when_there_are_several_blocks(tmp_path, rows):
    other = [hifi_row(1, "q", "EcoRI-HF", project="P002")]
    path = write_project_sheet(tmp_path / "p.xlsx", [rows, other])
    blocks = dt.read_blocks(path, "Projects", CONFIG)
    with pytest.raises(dt.DesignError, match="P001, P002"):
        dt.pick_block(blocks, None, "Plasmid Construction", "Projects")
    with pytest.raises(dt.DesignError, match="not found"):
        dt.pick_block(blocks, "P009", "Plasmid Construction", "Projects")


def test_a_block_with_another_label_is_refused(tmp_path, rows):
    path = write_project_sheet(tmp_path / "p.xlsx", [rows], label="Other work")
    blocks = dt.read_blocks(path, "Projects", CONFIG)
    with pytest.raises(dt.DesignError, match="Other work"):
        dt.pick_block(blocks, "P001", "Plasmid Construction", "Projects")


def test_blank_project_column_needs_the_option(tmp_path, rows):
    for r in rows:
        r["project_no"] = None
    path = write_project_sheet(tmp_path / "p.xlsx", [rows])
    blocks = dt.read_blocks(path, "Projects", CONFIG)
    assert blocks[0].project is None
    with pytest.raises(dt.DesignError, match="--project"):
        dt.pick_block(blocks, None, "Plasmid Construction", "Projects")
    assert dt.pick_block(blocks, "P777", "Plasmid Construction", "Projects") is blocks[0]


def test_unknown_sheet_names_the_sheets(project_xlsx):
    with pytest.raises(dt.DesignError, match="Projects"):
        dt.read_blocks(project_xlsx, "Nope", CONFIG)


def test_norm_sub_and_to_number():
    assert dt.norm_sub(3.0) == "3" and dt.norm_sub("03") == "3" and dt.norm_sub(None) == ""
    assert dt.norm_sub("A") == "A"
    assert dt.to_number("01") == 1 and dt.to_number("2.5") == 2.5 and dt.to_number("4.0") == 4
    assert dt.to_number("120 ng/uL") is None and dt.to_number("-") is None
    assert dt.to_number("") is None
    assert dt.to_number("nan") is None and dt.to_number("inf") is None


def test_csv_headers_ignore_case_spaces_and_unknown_columns(tmp_path, rows):
    header = [" Sub_No ", "PLASMID_NAME", "Re1", "extra column", "project_no"]
    names = ["sub_no", "plasmid_name", "re1", "extra column", "project_no"]
    path = write_csv(tmp_path / "d.csv", rows, header=names)
    lines = path.read_text(encoding="utf-8").splitlines()
    lines[0] = ",".join(header)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    got, seen = dt.read_csv(path, CONFIG)
    assert seen == ["sub_no", "plasmid_name", "re1", "project_no"]
    assert got[0].values["re1"] == "EcoRI-HF" and got[0].values["backbone"] is None


def test_csv_utf8_bom_is_accepted(tmp_path, rows):
    path = write_csv(tmp_path / "d.csv", rows)
    path.write_bytes(b"\xef\xbb\xbf" + path.read_bytes())
    got, _ = dt.read_csv(path, CONFIG)
    assert got[0].values["sub_no"] == "1"


def test_csv_project_no_repeats_the_one_above(tmp_path, rows):
    rows[3]["project_no"] = None
    got, _ = dt.read_csv(write_csv(tmp_path / "d.csv", rows), CONFIG)
    assert got[3].values["project_no"] == "P001"
    assert got[3].sub == "3" and not got[3].own


def test_coerce_turns_text_into_numbers(tmp_path, rows):
    rows[0]["sub_no"] = "01"
    rows[0]["conc"] = "120"
    rows[1]["conc"] = "120 ng/uL"
    got, _ = dt.read_csv(write_csv(tmp_path / "d.csv", rows), CONFIG)
    dt.coerce(got, ["sub_no", "backbone_length", "insert_length"])
    assert got[0].values["sub_no"] == 1 and got[0].sub == "1"
    assert got[0].values["backbone_length"] == 5000
    assert got[0].values["conc"] == 120 and got[1].values["conc"] == "120 ng/uL"
    assert got[3].values["backbone_length"] is None  # blank stays blank


def test_coerce_names_the_line_and_column(tmp_path, rows):
    rows[1]["insert_length"] = "long"
    got, _ = dt.read_csv(write_csv(tmp_path / "d.csv", rows), CONFIG)
    with pytest.raises(dt.DesignError, match=r"line 3: 'insert_length'.*long"):
        dt.coerce(got, ["insert_length"])


def test_sub_selection_takes_continuation_rows(project_xlsx):
    block = dt.read_blocks(project_xlsx, "Projects", CONFIG)[0]
    picked, notes = dt.select_rows(block.rows, ["3"])
    assert [r.line for r in picked] == [4, 5]
    assert notes == []
    picked, _ = dt.select_rows(block.rows, ["1", "3"])
    assert [r.line for r in picked] == [2, 4, 5]


def test_sub_selection_uses_the_last_of_a_repeated_sub(tmp_path, rows):
    rows.append(hifi_row(1, "p-redo", "EcoRI-HF", "", "bb", 100, "O-9", "O-10", "t", 10))
    block = dt.read_blocks(write_project_sheet(tmp_path / "p.xlsx", [rows]), "Projects",
                           CONFIG)[0]
    picked, notes = dt.select_rows(block.rows, ["1"])
    assert [r.line for r in picked] == [6]
    assert "appears 2 times" in notes[0]


def test_unknown_sub_is_an_error(project_xlsx):
    block = dt.read_blocks(project_xlsx, "Projects", CONFIG)[0]
    with pytest.raises(dt.DesignError, match=r"Sub\. No\. 9 not found"):
        dt.select_rows(block.rows, ["9"])


def test_rows_before_any_sub_are_shared_and_left_out(tmp_path, rows):
    shared = hifi_row("", "", "", "", "", None, "O-0", "O-0b", "t0", 10)
    block = dt.read_blocks(write_project_sheet(tmp_path / "p.xlsx", [[shared, *rows]]),
                           "Projects", CONFIG)[0]
    picked, notes = dt.select_rows(block.rows, ["2"])
    assert [r.line for r in picked] == [4]
    assert "shared step" in notes[0]


def test_split_subs():
    assert dt.split_subs(["3,4", "05"]) == ["3", "4", "5"]
    assert dt.split_subs(None) == []


def test_config_has_a_name_for_every_heading():
    assert len(NAMES) == 32 and len(HEADINGS) == 33
