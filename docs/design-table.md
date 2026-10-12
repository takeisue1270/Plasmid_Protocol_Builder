# The design table

The input is a table with one row per construct. It can be a CSV file or one
block of a project workbook; both are read by `scripts/design_table.py`, and the
column names are those in `config.json` (`columns`).

## Columns

| column | name | HiFi | GGA | content |
| --- | --- | --- | --- | --- |
| B | `project_no` | read | | Project No.; the same on every row of the table |
| C | `sub_no` | required | required | the construct's number; empty on a continuation row |
| D | `method` | | | `HiFi` or `GGA`; HiFi when it is empty everywhere. A table with both stops until `--method` keeps the rows of one (a row with no method, and the continuation rows of a sub, stay) |
| G | `plasmid_name` | required | required | name of the product plasmid |
| H | `re1` | required | required | HiFi: the enzyme that opens the backbone, or the name of a primer when the backbone is prepared by PCR (see below); GGA: the Type IIS enzyme |
| I | `re2` | required (may be empty) | | second enzyme, or the second primer of a PCR backbone; empty for a single-enzyme digest |
| J | `backbone` | required | required | backbone plasmid or template name |
| L | `backbone_length` | required | | backbone fragment length (bp) |
| M | `oligo1` | required | required | first PCR primer, or the sense oligo (GGA) |
| N | `oligo2` | required | required | second PCR primer, or the antisense oligo (GGA) |
| O | `template` | required | | insert template |
| P | `insert_length` | required | | insert fragment length (bp) |

"Required" means the column must exist in a CSV file; a value may be empty where
the row does not need it (a single-enzyme digest has an empty `re2`). The other
names in `config.json` (`plasmid_map`, `length`, `conc`, dates and so on) are
copied to `Data` when present and are not used by the sheets. In a CSV file,
headings are matched ignoring case and surrounding spaces, unknown columns are
ignored, the separator is a comma, and the encoding is UTF-8 (a BOM is accepted;
in Excel, save as "CSV UTF-8"). The headings are the names in the second column
of the table above, not the headings of a workbook ("Project No.", "RE1"). A
file with none of those names, or in another encoding, is refused with a message.

## Continuation rows

A construct that needs a second PCR (two inserts) has a second row with an empty
`sub_no`. It belongs to the sub above it, carries its own `oligo1`, `oligo2`,
`template` and `insert_length`, and the other cells stay empty. In a CSV file
`project_no` may be empty on such a row; it repeats the one above.

## Numbers

`sub_no`, `backbone_length` and `insert_length` must be numbers when they are not
empty (`01` and `1.0` are read as 1), in a CSV file and in a workbook alike; text
such as `1.2 kb` stops the run and names the line. `length` and `conc` are numbers when they read as one
and are kept as text otherwise (`120 ng/µL`). An empty CSV cell is written as an
empty cell.

## HiFi: `re1` as an enzyme or as a primer name

The `RE` sheet looks `re1` up in the `RElist` sheet. If it is there, the backbone
is digested. If it is not, the backbone is taken to be prepared by PCR and the row
is marked `PCR` there. `make_protocol.py` warns for a name that is not in the
sheet, since a mistyped enzyme looks the same as a primer name.

## A CSV file

```csv
project_no,sub_no,method,plasmid_name,re1,re2,backbone,backbone_length,oligo1,oligo2,template,insert_length
EX001,1,HiFi,pExample-01,EcoRI-HF,BamHI-HF,pBackbone-A,5200,OL-001,OL-002,pTemplate-A,1200
EX001,2,HiFi,pExample-02,EcoRI-HF,,pBackbone-A,5200,OL-003,OL-004,pTemplate-B,900
EX001,3,HiFi,pExample-03,XhoI,BamHI-HF,pBackbone-B,4800,OL-005,OL-006,pTemplate-C,1500
,,,,,,,,OL-007,OL-008,pTemplate-D,700
```

The last line is the second insert of sub 3. `examples/design_table.csv` is the
same file.

A GGA table needs fewer columns:

```csv
project_no,sub_no,method,plasmid_name,re1,backbone,oligo1,oligo2
EX002,1,GGA,pGuide-01,BsaI,pVector,OL-101,OL-102
EX002,2,GGA,pGuide-02,BsaI,pVector,OL-103,OL-104
```

## A block of a project workbook

A block is a header row whose column B reads `Project No.` (column A names the
kind of block; only `Plasmid Construction` is accepted), then one row per
construct, up to the next blank row or header row. The columns are in the order of
the table above, B to AG. Column B holds the Project No. on every row, as a value
or as a formula that repeats the cell above. Formula cells are read from their
last saved value, so save the workbook in Excel first.

```bash
python scripts/make_protocol.py --design projects.xlsx --project EX001
python scripts/make_protocol.py --design projects.xlsx --sheet Records --project EX001 --sub 3,4
```

`--project` matches the Project No. ignoring case, exactly otherwise. A workbook's
Project No. is read in capitals (`ex001` becomes `EX001`); a CSV file keeps the
spelling of its table. Without it,
a sheet with one block uses that block, and a sheet with several stops and lists
them. The headings in columns J, M, N and O are checked, because the sheets read
those columns by position; a different heading gives a `WARN`.

## Rows without a Sub. No.

A row that has no `sub_no` and no sub above it (the first row of the table, or rows
before the first sub), and a `sub_no` used on two separate rows, each give a `WARN`.
The sheets cannot place the first kind (the HiFi sheet can show an error), and the
second is counted as two constructs.

## Choosing rows: `--sub`

`--sub 3,4,5` keeps those subs, each with its continuation rows, in the order of
the table. A sub that occurs twice takes its last occurrence. Rows before the
first sub number are shared steps and are left out.
