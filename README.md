# Plasmid Protocol Builder

Builds bench-protocol workbooks for HiFi assembly and Golden Gate assembly (GGA)
from a design table.

## Purpose

Once a construct is designed, the bench work is arithmetic: how many digests,
how much of each reagent in the master mix, how long the extension step is, how
many plates. This tool writes that arithmetic for a batch of constructs into one
Excel workbook. You give it a design table (a CSV file, or one Project No. block
of a project workbook). It writes the rows into the `Data` sheet of a template,
and every other sheet is formulas that read `Data`. Excel computes them when the
workbook is opened.

The project record written by
[Plasmid Assembly Designer](https://github.com/takeisue1270/Plasmid_Assembly_Designer)
is read as it is. A CSV file works as well when its headings are the column names
of that record in lower case with underscores (`sub_no`, `plasmid_name`, `re1`, ...;
see [`docs/design-table.md`](docs/design-table.md)).

## What you get

| workbook | sheets |
| --- | --- |
| HiFi (`*.HiFi.TF.xlsx`) | `RElist` enzyme table · `Data` · `RE` backbone digestion · `PCR` insert PCR and DpnI · `HiFi` assembly mix · `TF` transformation |
| GGA (`*.GGA.TF.xlsx`) | `Data` · `Anneal` oligo phosphorylation and annealing · `GGA` ligation mixes · `TF` transformation |

[`docs/method.md`](docs/method.md) states what each sheet computes.

## Quick start

```bash
git clone https://github.com/takeisue1270/Plasmid_Protocol_Builder.git
cd Plasmid_Protocol_Builder
pip install -r requirements.txt
python scripts/make_protocol.py --design examples/design_table.csv --date 261011
```

This writes `output/261011.EX001.HiFi.TF.xlsx` and prints a `WARN` line for every row
whose enzyme is not in that sheet, because the `RElist` sheet that ships with the template is empty (see
below). Open the workbook in Excel. The sheets calculate on opening.

Then, in the workbook:

- Set `RE!C10` to `Yes` if the digested backbone is column-purified, or to `No` if it
  is heat-inactivated instead. The default is `Yes`.
- On the `HiFi` sheet, type the measured concentration of every fragment into
  column C (ng/µL). Until you do, the volume columns next to it show `#DIV/0!`, and
  they stay so on the unused rows below your constructs; hide those rows.

## Adding your own enzymes (RElist)

The `RE` sheet looks each `re1` and `re2` up in the `RElist` sheet. A row whose
`re1` is not there is treated as a backbone prepared by PCR, and a row whose `re1`
is there but whose `re2` is not shows `Check` for the reaction temperature. So
**the sheet ships empty and you add the enzymes you use**. From row 4 down, one enzyme per row, columns
A to D:

| column | content |
| --- | --- |
| A Enzyme | the name as it appears in the design table, matched ignoring case |
| B Buffer | `CutSmart` if the reaction in the `RE` sheet applies |
| C Reaction temp. | `37°C` |
| D Heat inactivation | `65°C for 20 minutes`, or `No` if the enzyme cannot be heat-inactivated |

The sheet reads rows 4 to 650. Write the values from your supplier's current
datasheet. An edit made to an output workbook is lost the next time one is built,
so copy `templates/protocol.hifi.xlsx`, add the rows to the copy, and pass it
with `--template`:

```bash
python scripts/make_protocol.py --design design.csv --template my_hifi_template.xlsx
```

`scripts/relist.py` has two helpers for building the table from a supplier's
list: `normalize_enzyme_name` removes registered-trademark signs from a name, and
`derive_buffer` returns `CutSmart` when the activity in CutSmart is 100%, else the
first of r1.1, r2.1 and r3.1 at 100%, else the most active buffer marked
`(check buffer)`.

## The design table

A CSV file whose headings are the column names below, or one block of a project
workbook (a header row whose column B reads `Project No.`, then one row per
construct). [`docs/design-table.md`](docs/design-table.md) gives the columns, the
rules for continuation rows, and a worked example.

| needed for | columns |
| --- | --- |
| both methods | `sub_no`, `plasmid_name`, `re1`, `backbone`, `oligo1`, `oligo2` |
| HiFi, in addition | `re2`, `backbone_length`, `template`, `insert_length` |
| optional | `project_no` (needed through `--project` when the table has none), `method` (`HiFi` or `GGA`; HiFi when absent) |

## Options

| option | meaning |
| --- | --- |
| `--design <file>` | the CSV file or the project workbook |
| `--sheet <name>` | sheet of the project workbook (default `Projects`) |
| `--project <No.>` | the Project No., needed when the table has several |
| `--method HiFi\|GGA` | the method; keeps only the rows whose `method` column says so (a row with none stays) |
| `--sub 3,4,5` | only these sub numbers, each with its continuation rows |
| `--date yymmdd` | date in the file name (default today) |
| `--template <xlsx>` | a template of your own (default the bundled one) |
| `--out-dir <dir>` | output folder (default `./output`) |
| `--config <file>` | another `config.json` (default the one beside `scripts/`) |
| `--dry-run` | print what would be written and write nothing |
| `--force` | overwrite an existing output file |

The output is `<out-dir>/YYMMDD.<project>.<HiFi|GGA>.TF.xlsx`. The script prints a
`WARN` line when something in the table will make a sheet show `Check` or leave
rows out; read them before using the workbook.

## Requirements

- Python 3.10 or later with `openpyxl` (`requirements.txt`).
- Excel to open the result. The workbooks were checked in Microsoft Excel; other
  spreadsheet programs are untested. The scripts need neither Excel nor a network
  and run on Windows, macOS and Linux. CI runs the tests on Ubuntu and Windows
  with Python 3.10 and 3.12.

## Limits

- The workbook is saved without calculated values. A preview that does not
  calculate (a file browser, a mail client) shows the formula cells empty until
  the file is opened in Excel.
- The `RE` sheet is built for a 10× CutSmart reaction: 50 µL, 10 µL plasmid, 2 µL
  of each enzyme. An enzyme that needs another buffer needs the sheet edited.
  `make_protocol.py` warns for an enzyme whose Buffer column is not `CutSmart`.
- The reagent volumes and cycling conditions in the templates are the author's
  protocol values: KOD One for the insert PCR, a FastDigest-type DpnI step, NEBridge
  kits for GGA, 25 µL of competent cells and 900 µL of recovery medium per
  sample. Check them against the datasheet of what you use, and change the cells
  that the comments in the template point to.
- A HiFi workbook covers 12 rows and a GGA workbook 48; a continuation row counts as
  a row. Rows beyond that are not in the protocol, and the script warns.
- Column purification is a single choice for the whole `RE` sheet.

## Working without the scripts

Everything the script does is copying values: paste the design rows into the
`Data` sheet of a copy of the template, starting at row 2, in the columns that
[`docs/design-table.md`](docs/design-table.md) lists. The other sheets need
nothing else.

## Contents

| path | content |
| --- | --- |
| `scripts/make_protocol.py` | the command line |
| `scripts/design_table.py` | reads a CSV file or a project-workbook block |
| `scripts/relist.py` | enzyme-name and buffer helpers for building an RElist |
| `scripts/dump_template.py` | writes a workbook's cells as text; `--update` rewrites `tests/golden/` |
| `templates/` | the HiFi and GGA templates |
| `config.json` | column map, per-method settings, the notes written into the templates |
| `examples/design_table.csv` | a synthetic example |
| `docs/` | how the sheets compute; the design-table format; what the template comments say |
| `tests/` | offline tests; no Excel, no network |

## What leaves your machine

Nothing. The scripts read and write local files only and open no network
connection.

## Scope and responsibilities

This tool turns a list of constructs into reagent volumes. It does not choose a
target, design a sequence or judge whether an experiment is permitted where you
work. Work with engineered organisms is regulated — in Japan by the Cartagena Act
— and your institution's biosafety committee, not this repository, decides what
you may do. Uses for human germline editing, for gene drives intended for release,
and for increasing the transmissibility or virulence of a pathogen are outside
the intended use of this tool.

Check every volume and condition against the datasheet of the reagent you use
before you rely on it.

## Disclaimer

This is a research tool, provided as is and without warranty of any kind, as the
[licence](LICENSE) states. It is not intended for diagnostic, therapeutic or
clinical use, and what it returns is a proposal — check it before you spend
anything on it.

Microsoft Excel, New England Biolabs (NEB) and NEBridge, Toyobo and KOD One, and
Thermo Fisher Scientific and FastDigest belong to the companies that make them.
This project is independent of all of them: not affiliated with, endorsed by, or
supported by any of them, and it names them only to say which reagents the
template protocols were written for and which program opens the workbooks.

## Citation

Citation metadata is in [`CITATION.cff`](CITATION.cff), which GitHub renders
under "Cite this repository".

## Licence

MIT. See [LICENSE](LICENSE).
