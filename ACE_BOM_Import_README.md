# ACE_BOM_Import — BOM extraction → ACE workbook

`ACE_BOM_Import.bas` is an Excel VBA module that takes a CATIA BOM extraction
(the file produced by the `CATIA_BOM_Extractor_*` macros, e.g. `flint bom 2.xlsx`)
and brings it into the ACE cost model (`ACE_2_1_newVersion.xlsm`).

One run does two things:

1. **Separate sheet with the same format** — the complete extraction (every
   column, all formatting, all thumbnails) is copied into a new sheet of the
   ACE workbook called **`BOM Extract`**, placed right after the `Input` sheet.
2. **`Input` sheet filled** — part numbers, names, quantities, levels and the
   thumbnails are written into the `Input` sheet of the cost model.

## Column mapping

| BOM extraction        | → | `Input` sheet          |
|-----------------------|---|------------------------|
| `Level`               | → | `A` Level              |
| `Part Number`         | → | `C` Part Number        |
| `Description`         | → | `D` Name               |
| `First Level`         | → | `E` Other Reference    |
| `Thumbnail` (picture) | → | `F` Picture            |
| `Qty`                 | → | `G` Qty System         |

Data is written from row 15 downwards (row 14 is the header row).

Before writing, the macro empties all entry columns of a BOM line (see
*Cleaning* below), so nothing of the previous BOM stays behind. The columns it
cannot fill from the extraction — Lookup Key, Material Reference, process
entries — simply stay empty.

Everything else on the `Input` sheet is left untouched — the cost columns,
factory/country, material, and every formula stay exactly as they are, so the
cost model keeps working after an import.

Headers are matched by name, not by position, so the extraction columns may sit
anywhere and the header row may be any of the first 20 rows. Accepted synonyms:
`Part Number` / `PartNumber` / `Part No`, `Description` / `Name` / `Designation`,
`Qty` / `Quantity`, `Level`, `First Level` / `Top Level`.

### Newer extractions

The extractors now also write **sub-assembly rows**, a **`First Level`** column
and a **`Type`** column (Assembly / Part / Body). Nothing needs to be configured
for that:

* extra columns are simply ignored — the mapping goes by header name,
* sub-assembly and body rows are imported like any other row, so the tree
  structure of the BOM ends up in the `Input` sheet,
* the indentation the extractor puts in front of a body name (`    PartBody`)
  is removed, so no part number starts with blanks,
* `First Level` lands in `E` Other Reference, so the `Input` sheet can be
  filtered per first-level assembly too. Set `IMPORT_FIRST_LEVEL = False` at the
  top of the module to leave `E` empty instead.
* rows whose `Type` is `Assembly` are written **in bold** (columns A–G), so the
  structure rows can be told apart from the parts that carry the cost. The block
  is un-bolded first on every import, so nothing stale survives. Switch it off
  with `MARK_ASSEMBLY_ROWS = False`.

Because sub-assembly rows add lines, a BOM can now exceed the 273 prepared rows
of the `Input` sheet more easily — the macro then offers to extend the sheet by
copying the last prepared row down (see below).

## Installation

1. Open `ACE_2_1_newVersion.xlsm`.
2. `Alt` + `F11` → **File → Import File…** → select `ACE_BOM_Import.bas`.
3. Close the VBA editor and save the workbook as `.xlsm` (macro-enabled).

4. Run `CreateMacroButtons` once (`Alt` + `F8`) — it puts the two push buttons
   **[ Import BOM ]** and **[ Clean Input ]** on the `Input` sheet. Save again
   to keep them.

## Usage

`Alt` + `F8`, then pick one of:

| Macro                  | What it does                                                     |
|------------------------|------------------------------------------------------------------|
| `ImportBOM`            | Normal use — creates the `BOM Extract` sheet **and** fills `Input` |
| `ImportBOM_SheetOnly`  | Only creates the `BOM Extract` copy sheet                        |
| `ImportBOM_InputOnly`  | Only fills the `Input` sheet                                     |
| `ClearBOMImport`       | **Cleaning** — empties all entry columns of `Input` (values + pictures) |
| `FitPicturesToCells`   | Re-fits all thumbnails into their `Picture` cells                |
| `ResetInputSheet`      | Empties the **whole** `Input` sheet for a new estimate — formulas stay |
| `NewInputSheet`        | Creates a new, empty copy of the `Input` sheet                   |
| `CreateMacroButtons`   | Puts the two push buttons on the `Input` sheet (run once)        |

A file dialog asks for the extraction file (it starts in the folder of the ACE
workbook). If that file is already open in Excel it is reused and left open;
otherwise it is opened read-only and closed again straight away.

At the end a summary reports how many rows and pictures were imported and how
long it took.

## Cleaning — `ClearBOMImport`

Empties every **entry column** of a BOM line in rows 15–287 and removes all
pictures of the data area. This is what the **[ Clean Input ]** button runs, and
every import runs it first.

| | | |
|---|---|---|
| `A` Level | `B` Lookup Key | `C` Part Number |
| `D` Name | `E` Other Reference | `F` Picture |
| `G` Qty System | `P` Material Reference | `Q` Description / Comments |
| `AC` Process Reference | `AD` Process Description | `AH` Process Time |
| `AI` Direct Operator | `AN` Qty Operators for setup | `AO` Set Up Time per Lot Size |
| `AP` Other Setup Cost | | |

Cells containing a **formula are never removed** — only entered values. On the
current workbook that clears 1 620 cells; the 34 formula cells inside `AH` and
everything in the calculated columns stay in place.

Two consequences worth knowing, both identical to deleting those cells by hand:
emptying `B` (Lookup Key) and `Q` (Comments) makes the VLOOKUPs in `R` (Material
Consumption) show `#N/A`, and emptying `P` (Material Reference) does the same for
`S` / `T` — until new entries are made.

The column list is the constant `CLEAR_COLUMNS` at the top of the module; add or
remove letters there to change what the cleaning covers.

## Starting a new estimate

Two macros clear the entered data without ever touching a formula:

### `ResetInputSheet` — empty the existing sheet

Deletes every **typed-in value** in the data area of `Input` (rows 15–287) and
all pictures of that area. Cells containing a **formula are never touched** — the
macro clears constants only, so the complete cost model keeps working and simply
calculates on empty inputs. On the current workbook this clears ~2 560 entered
cells while all ~15 400 formulas stay in place.

Cleared: everything the cleaning covers (see above) plus Reference Cost, supply
option, material, manually entered material consumption / consumption units /
process category, rank & learning parameter, and the Product / subproducts /
Manufacturing labels.

**Kept on purpose** (columns `K, L, AA, AE, BC`): Company, Country, Process
Type, Lot Size and Valuation Method. These are entered values too, but they
carry the template defaults — emptying `AE` (Lot Size) would give `#DIV/0!` in
the setup cost because it is a divisor, and emptying `K` (Company) would set
every factory rate to 0. Change the list in the constant `RESET_KEEP_COLUMNS` at
the top of the module if you want a different behaviour.

A confirmation dialog appears first (the default button is "No"); the action
cannot be undone, so save a copy beforehand if in doubt.

### `NewInputSheet` — start on a fresh sheet

Duplicates the `Input` sheet with **all formulas, formatting, dropdowns and
column widths**, then empties the copy in exactly the same way. The original
`Input` sheet is not changed, so the previous estimate stays intact. The new
sheet is named `Input (new)` (`Input (new) 2`, … if it already exists).

Note: the evaluation sheets (`Overview_1`, `System`, `CBS`, `Parametric Cost`, …)
keep reading from the original `Input` sheet — the new sheet is a self-contained
working copy, e.g. for a variant or a second BOM.

A typical "new project" run is therefore: `ResetInputSheet` → `ImportBOM`.

## Behaviour details

* **Re-importable.** Every run first runs the same cleaning as
  `ClearBOMImport` (see below), so nothing of the previous BOM stays behind.
* **Pictures.** Each thumbnail is scaled into the `Picture` cell of its row —
  aspect ratio kept, centred in the cell, row height raised to at least 45 pt —
  and set to **Move and size with cells**, so it follows when you change the row
  height or column width. Imported pictures are named `BOMPIC_<row>`; that name
  is what `ClearBOMImport` and `FitPicturesToCells` use to find them again, even
  if a picture was dragged somewhere else.
  After all thumbnails are pasted the macro fits every one of them a second time,
  so a hiccup during a single paste cannot leave a picture oversized.
* **`FitPicturesToCells`** re-fits every thumbnail of the `Input` sheet into its
  cell. Run it after changing row heights or column widths, or to repair
  pictures that ended up at the wrong place.
* **Cleaning up pictures.** `ClearBOMImport` and every new import delete all
  imported pictures plus any picture sitting inside the data rows — including
  ones that landed at the wrong position. The sheet logo above the data area is
  kept.
* **More BOM rows than prepared template rows.** The `Input` sheet ships with
  prepared rows 15–287. If the extraction is longer, the macro asks whether to
  extend the sheet by copying the last prepared row (with all its formulas) down,
  or to import only as many rows as fit. Skipped rows are reported.
* **Numbers stay numbers.** The extractor writes `Level` and `Qty` as text;
  they are converted back to numeric values so `Qty System` feeds the cadence
  formula in column J.
* **Formulas in the copy sheet** are replaced by their values so that the new
  sheet never points back at the extraction workbook.
* **Sheet protection** on `Input` is removed before writing and restored
  afterwards. If the sheet is password protected, the macro stops and asks you
  to unprotect it manually.
* Screen updating, events and calculation are switched off during the import and
  restored (plus a full recalculation) at the end — also if an error occurs.

## The push buttons

`CreateMacroButtons` puts two form-control buttons on the `Input` sheet:

| Button | Macro |
|--------|-------|
| **Import BOM** | `ImportBOM` — copy sheet + fill `Input` |
| **Clean Input** | `ClearBOMImport` — the cleaning described above |

They are placed at cell `F5` (the free area next to the header block), 100 × 30 pt
each. Run the macro again at any time — it replaces the two buttons instead of
adding new ones. To move a button by hand, right-click it and drag it; to change
the default position or size, edit `BUTTON_ANCHOR_CELL`, `BUTTON_WIDTH`,
`BUTTON_HEIGHT` and `BUTTON_GAP` at the top of the module and run
`CreateMacroButtons` again.

The buttons are part of the workbook — save the file as `.xlsm` after creating
them, and they stay there for everyone who opens it.

## Tuning

Constants at the top of the module:

| Constant                | Default        | Meaning                                       |
|-------------------------|----------------|-----------------------------------------------|
| `BOM_SHEET_NAME`        | `BOM Extract`  | Name of the copy sheet                        |
| `WRITE_LEVEL`           | `True`         | `False` keeps the levels already in column A  |
| `COPY_PICTURES`         | `True`         | `False` imports data only (much faster)       |
| `MIN_PIC_ROW_HEIGHT`    | `45`           | Minimum row height for rows with a picture    |
| `INPUT_FIRST_DATA_ROW`  | `15`           | First data row of the `Input` sheet           |
| `RESET_KEEP_COLUMNS`    | `K,L,AA,AE,AH,AI,AN,AO,AP,BC` | Columns whose defaults a reset keeps |
| `CLEAR_COLUMNS`         | `A,B,C,D,E,F,G,P,Q,AC,AD,AH,AI,AN,AO,AP` | Entry columns the cleaning empties |
| `BUTTON_ANCHOR_CELL`    | `F5`            | Where `CreateMacroButtons` puts the buttons |

## Good to know about the reset

A column such as `R` (Material Consumption) or `AB` (Process Category) contains a
formula in most rows and a manually entered override in a few. The reset removes
those overrides, so the cells stay **empty** afterwards — the formula rows keep
their formulas, but a cleared override is not replaced by one. That is intended:
the macro deletes inputs, it never writes formulas.
