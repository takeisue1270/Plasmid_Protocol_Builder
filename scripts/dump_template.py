r"""Write a workbook's cells, comments, number formats and conditional formats as text.

One item per line, sheet by sheet and cell by cell:

    <sheet>!<cell>\t<kind>\t<content>

``kind`` is value, formula, numfmt or comment. A conditional format is keyed by
its range: ``<sheet>!<range>\tcf\t<type>|<formula>|<priority>``. Newlines and
tabs inside a content are written as ``\n`` and ``\t``, so the output of two
workbooks can be compared line by line.

    python scripts/dump_template.py templates/protocol.hifi.xlsx
    python scripts/dump_template.py --update      # rewrite tests/golden/*.txt

``--update`` rewrites the golden file of each bundled template. The test suite
compares a fresh dump with that file, so any change to a template shows up in
the diff of a commit.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ("protocol.hifi.xlsx", "protocol.gga.xlsx")


def _flat(value: object) -> str:
    text = str(value)
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\\", "\\\\") \
        .replace("\n", "\\n").replace("\t", "\\t")


def dump(path: str | Path) -> list[str]:
    """The dump of one workbook, as a list of lines without line ends."""
    wb = openpyxl.load_workbook(path)
    lines: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                where = f"{ws.title}!{cell.coordinate}"
                value = cell.value
                if value is not None:
                    is_formula = isinstance(value, str) and value.startswith("=")
                    kind = "formula" if is_formula else "value"
                    lines.append(f"{where}\t{kind}\t{_flat(value)}")
                if value is not None and cell.number_format != "General":
                    lines.append(f"{where}\tnumfmt\t{_flat(cell.number_format)}")
                if cell.comment is not None:
                    lines.append(f"{where}\tcomment\t{_flat(cell.comment.text)}")
        rules = []
        for rng, items in ws.conditional_formatting._cf_rules.items():
            for rule in items:
                formula = ";".join(str(f) for f in (rule.formula or []))
                text = f"{rule.type}|{formula}|{rule.priority}"
                rules.append((str(rng.sqref), rule.priority, text))
        for sqref, _, text in sorted(rules, key=lambda r: (r[0], r[1] or 0)):
            lines.append(f"{ws.title}!{sqref}\tcf\t{_flat(text)}")
    return lines


def write_golden(root: Path = ROOT) -> list[Path]:
    """Rewrite tests/golden/<stem>.txt for each bundled template."""
    out_dir = root / "tests" / "golden"
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for name in TEMPLATES:
        target = out_dir / f"{Path(name).stem}.txt"
        text = "\n".join(dump(root / "templates" / name)) + "\n"
        with open(target, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        written.append(target)
    return written


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("workbook", nargs="?", help="the .xlsx to dump (omit with --update)")
    ap.add_argument("--update", action="store_true",
                    help="rewrite tests/golden/ from the bundled templates")
    args = ap.parse_args(argv)
    if args.update:
        for path in write_golden():
            print(f"wrote {path}")
        return 0
    if not args.workbook:
        ap.error("give a workbook, or --update")
    sys.stdout.reconfigure(encoding="utf-8", newline="\n")
    for line in dump(args.workbook):
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
