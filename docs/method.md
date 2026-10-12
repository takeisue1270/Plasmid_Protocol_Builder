# What each sheet computes

Every sheet except `Data` is formulas. `make_protocol.py` only fills `Data`; the
rest follows. Cell addresses below are those of the bundled templates. The
volumes and conditions are the author's protocol values, so check them against
the datasheet of the reagent you use.

## Data

One row per line of the design table, from row 2, in the columns of
[`design-table.md`](design-table.md). Column B holds the Project No. on every
row, and column C (`Sub. No.`, formatted `00`) is empty on a continuation row.
The other sheets read only these columns:

| template | columns read |
| --- | --- |
| HiFi | B, C, G, H, I, J, L, M, N, O, P |
| GGA | C, G, H, J, M, N |

The other columns are copied from the table so that the output can be compared
with the design, and they are free for your own notes.

## HiFi workbook

### RElist

The enzyme table the `RE` sheet looks up: Enzyme, Buffer, Reaction temp., Heat
inactivation, from row 4 to row 650. It ships empty; see the README.

### RE: digesting the backbone

A row is a digest (`RE`) when `re1` is in `RElist`, and a backbone prepared by
PCR (`PCR`) when it is not. A row with no `Sub. No.` is blank. A row with no `re1`
shows its names but nothing in the `Prep` column and the columns after it.

- The reaction is 50 µL (`C8`): 10 µL plasmid (`C9`), 5 µL 10× CutSmart, 2 µL of
  each enzyme (`C8/25`), water to volume. A single-enzyme row has 0 µL of the
  second enzyme and more water.
- Reaction temperature: the temperature of the enzyme in `RElist`; with two
  enzymes, the common temperature, or `Check conditions` when they differ; `Check`
  when an enzyme is not in `RElist`.
- Heat inactivation: `Not needed` while column purification (`C10`) is `Yes`.
  Otherwise `Not possible` if either enzyme is `No` in `RElist`, else `Possible`;
  `Check` when an enzyme is not in `RElist`.
- A premix of water and 10× CutSmart is calculated for the number of digests plus
  two. A note appears when the water volume differs between rows (a different
  number of enzymes), because one premix then does not fit all of them.

The sheet is built for 10× CutSmart. It does not read the Buffer column; the
script warns for an enzyme whose Buffer is not `CutSmart`.

### PCR: the insert

One PCR per line that has a first primer (`oligo1`), including continuation rows.

- Template solution (1/5 dilution): 2 µL template DNA and 8 µL water, per template.
- Primer solution (100 µM to 10 µM): 2 µL primer and 18 µL water, per primer.
- PCR mix per tube: 21 µL water, 25 µL KOD One, 1.5 µL of each primer, 1 µL of the
  diluted template. The number of mixes is the number of PCRs plus two.
- Cycling: 98 °C 2 min; 40 cycles of 98 °C 10 s, 55 °C 5 s, 68 °C for the
  extension time; hold at 4 °C. The extension time is 5 s per kb of the longest
  product (`insert_length`), at least 5 s.
- DpnI: 5.5 µL 10×FD buffer and 1 µL DpnI per PCR, added to the finished PCR, 37 °C
  for 15 min then 80 °C for 5 min. These two volumes are for a FastDigest-type
  DpnI.

### HiFi: the assembly

The sheet turns the concentration you type (ng/µL, column C) and the fragment
length (bp) of each fragment into pmol/µL (ng/µL × 1000 ÷ bp ÷ 650), then into the
volume that holds a set amount. Every fragment, insert or backbone, is set to the
same amount `C3` (`G9:G20` and `G25:G36` are `=$C$3`; `C3` is 0.007 × `A3`, with
the reaction volume `A3` = 5 µL). Column `H` is the volume of a fragment that
holds `C3` pmol and column `I` is half of it. The last block uses `I`: for each
construct it gives the volumes of the inserts and the backbone, 2.5 µL of 2× HiFi
master mix (`A3/2`), water to make `A3` µL, and the incubation (50 °C, 60 min).
Continuation rows of one `Sub. No.` are the inserts of that construct. The column
headings "final 20 uL" and "final 10 uL" are the author's labels and do not match
`A3`; the sheet follows its formulas.

`B1` of the sheet says "Insert : backbone ratio = 2 : 1", and `E3` shows a
backbone amount that is half of `C3` for up to three inserts, but no volume reads
`E3`. `E3`, `G2` and `G3` are displayed totals computed from `A8`, which counts
the rows of the most common sub (blank rows count 1, 2, 3 … in the same column, so
`A8` can also pick up a coincidence when the subs do not start at 1). To use a 2 : 1
ratio, change the formulas in `G25:G36`.

### TF: the transformation

For the number of constructs: 25 µL of competent cells and 900 µL of recovery
medium per sample, one LB+Amp plate for every four samples (four-section plates),
and a list assigning each sample a plate and a section. The list numbers the rows of
the sheet by position, so a continuation row in the middle of a HiFi table leaves a
gap and the list can name a plate beyond the plate count; count the plates from
the list.

## GGA workbook

### Anneal

Phosphorylation and annealing of each oligo pair: a T4 PNK mix (water, T4 DNA
ligase reaction buffer, T4 polynucleotide kinase) for the number of samples plus
two, 1 µL of each oligo at 10 µM, and the thermal-cycler program.

### GGA

Each construct is assigned to mix 2-1 when `re1` begins with `BsaI` or `BsmBI`
(NEBridge Golden Gate Enzyme Mix, BsaI-HFv2), and to mix 2-2 otherwise
(NEBridge Ligase Master Mix and the Type IIS enzyme of that row). Reaction volume
6 µL and 0.4 µL of backbone plasmid; the annealed oligo mix at 50 nM is 0.6 µL in
mix 2-1 and 0.4 µL in mix 2-2. A note appears when mix 2-2 has more than one
enzyme, since the mix must then be split by enzyme. The 2-1 mix is built for the
BsaI-HFv2 enzyme mix and is used for `BsmBI` rows as well; change it if another
mix suits your enzyme.

### TF

As for the HiFi workbook, with the plate count taken from the number of samples
in `GGA`.
