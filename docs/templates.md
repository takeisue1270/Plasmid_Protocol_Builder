# The templates and the notes written into them

`templates/protocol.hifi.xlsx` and `templates/protocol.gga.xlsx` are ordinary
workbooks. Copy one, edit the copy, and pass it with `--template`.

## Notes in the cells

Each cell below carries a comment that says what to change when your reagents
differ from the author's. The comments are also listed in `config.json`
(`template_notes`), which the tests compare with the templates.

| template | cell | what the comment says |
| --- | --- | --- |
| HiFi | `RElist!A3` | add enzymes below the last row, columns A to D, rows 4 to 650; edit a copy of the template |
| HiFi | `RE!H13` | the buffer and its volume are for 10× CutSmart |
| HiFi | `PCR!B28` | volumes and cycling are for KOD One |
| HiFi | `PCR!B46` | the volumes in rows 46 and 47 are for a FastDigest-type DpnI and its buffer |
| HiFi | `Data!A1` | only columns B, C, G, H, I, J, L, M, N, O and P are read |
| GGA | `Anneal!B9` | reagents and volumes are the author's |
| GGA | `GGA!B14`, `GGA!B25` | volumes are for the NEBridge kit named there |
| GGA | `Data!A1` | only columns C, G, H, J, M and N are read |
| both | `TF!B6` | 25 µL of competent cells per sample; the number is in the formula in column C, the text in column E and the instruction line in row 12 |
| both | `TF!B7` | 900 µL of medium per sample; the number is in the formula in column C, the text in column E and the instruction line in row 18 |

## RElist

Row 1 and row 2 are instructions. Row 3 is the heading: Enzyme, Buffer, Reaction
temp., Heat inactivation. Enzymes start in row 4. The `RE` sheet reads
`RElist!$A$4:$D$650` for temperature and heat inactivation and
`RElist!$A$4:$A$650` to decide whether `re1` is an enzyme. An enzyme row must have
all four cells filled and nothing to the right of column D.

## Conditional formats in `Data`

Rows 2 to 4 of `Data` in the HiFi template carry four conditional formats that change only the look of
the cells: rows without a `Sub. No.`, GGA rows (the columns a GGA row does not use),
rows where column W holds a value (yellow), and `R2:R4` when `template` (column O)
does not begin with `V`, an identifier prefix of the author's laboratory. Delete
a rule (Home, Conditional Formatting, Manage Rules) or edit the letter if it does
not suit you.

## What the build keeps and sets

`make_protocol.py` copies the template, clears `Data` below the heading, writes
the rows, and saves with the calculation set to run when the file is opened. The
author name and last-modified name of the file are empty, and the comments are
saved at a size that shows their text.

## Checking a template

```bash
python scripts/dump_template.py templates/protocol.hifi.xlsx
```

prints every value, formula, number format, comment and conditional format, one
per line. `python scripts/dump_template.py --update` rewrites `tests/golden/`;
`pytest` fails when a template and its golden file differ, so a change to a
template always appears as a change to the golden file in the same commit.
