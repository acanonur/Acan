# ACE_BOM_Import — BOM extraction → ACE workbook

`ACE_BOM_Import.bas` is an Excel VBA module that takes a CATIA BOM extraction
(the file produced by the `CATIA_BOM_Extractor_*` macros, e.g. `flint bom 2.xlsx`)
and brings it into the ACE cost model (`ACE_2_1_newVersion.xlsm`).

One run does two things:

1. **Separate sheet with the same format** — the complete extraction (all 13
   columns, all formatting, all thumbnails) is copied into a new sheet of the
   ACE workbook called **`BOM Extract`**, placed right after the `Input` sheet.
2. **`Input` sheet filled** — part numbers, names, quantities, levels and the
   thumbnails are written into the `Input` sheet of the cost model.

## Column mapping

| BOM extraction        | → | `Input` sheet          |
|-----------------------|---|------------------------|
| `Level`               | → | `A` Level              |
| `Part Number`         | → | `C` Part Number        |
| `Description`         | → | `D` Name               |
| `Thumbnail` (picture) | → | `F` Picture            |
| `Qty`                 | → | `G` Qty System         |

Data is written from row 15 downwards (row 14 is the header row).

`B` Lookup Key and `E` Other Reference are **emptied** as well — the extraction
has nothing to fill them with, and stale keys of the previous BOM would be worse
than empty ones.

Everything else on the `Input` sheet is left untouched — the cost columns,
factory/country, material, process and every formula stay exactly as they are,
so the cost model keeps working after an import.

Headers are matched by name, not by position, so the extraction columns may sit
anywhere and the header row may be any of the first 20 rows. Accepted synonyms:
`Part Number` / `PartNumber` / `Part No`, `Description` / `Name` / `Designation`,
`Qty` / `Quantity`, `Level`.

## Installation

1. Open `ACE_2_1_newVersion.xlsm`.
2. `Alt` + `F11` → **File → Import File…** → select `ACE_BOM_Import.bas`.
3. Close the VBA editor and save the workbook as `.xlsm` (macro-enabled).

Optional: put a button on the `Input` sheet (Developer → Insert → Button) and
assign the macro `ImportBOM`.

## Usage

`Alt` + `F8`, then pick one of:

| Macro                  | What it does                                                     |
|------------------------|------------------------------------------------------------------|
| `ImportBOM`            | Normal use — creates the `BOM Extract` sheet **and** fills `Input` |
| `ImportBOM_SheetOnly`  | Only creates the `BOM Extract` copy sheet                        |
| `ImportBOM_InputOnly`  | Only fills the `Input` sheet                                     |
| `ClearBOMImport`       | Empties the part columns A–G of `Input` (values + pictures)      |
| `ResetInputSheet`      | Empties the **whole** `Input` sheet for a new estimate — formulas stay |
| `NewInputSheet`        | Creates a new, empty copy of the `Input` sheet                   |

A file dialog asks for the extraction file (it starts in the folder of the ACE
workbook). If that file is already open in Excel it is reused and left open;
otherwise it is opened read-only and closed again straight away.

At the end a summary reports how many rows and pictures were imported and how
long it took.

## Starting a new estimate

Two macros clear the entered data without ever touching a formula:

### `ResetInputSheet` — empty the existing sheet

Deletes every **typed-in value** in the data area of `Input` (rows 15–287) and
all pictures of that area. Cells containing a **formula are never touched** — the
macro clears constants only, so the complete cost model keeps working and simply
calculates on empty inputs. On the current workbook this clears ~1 200 entered
cells while all ~15 400 formulas stay in place.

Cleared: Level, Part Number, Name, Other Reference, Picture, Qty, Reference
Cost, supply option, material, material reference, comments, manually entered
material consumption / units / process category / process reference /
description, rank & learning parameter, and the Product / subproducts /
Manufacturing labels.

**Kept on purpose** (columns `K, L, AA, AE, AH, AI, AN, AO, AP, BC`): Company,
Country, Process Type, Lot Size, Process Time, Direct Operator, Qty Operators,
Set Up Time, Other Setup Cost and Valuation Method. These are entered values
too, but they carry the template defaults — emptying `AE` (Lot Size) would give
`#DIV/0!` in the setup cost, emptying `K` (Company) would set every factory rate
to 0. Change the list in the constant `RESET_KEEP_COLUMNS` at the top of the
module if you want a different behaviour.

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

* **Re-importable.** Every run first empties the part columns
  `A` Level, `B` Lookup Key, `C` Part Number, `D` Name, `E` Other Reference,
  `F` Picture and `G` Qty System (constant `CLEAR_COLUMNS`) including all
  pictures of the data area, so nothing of the previous BOM stays behind.
  `B` and `E` are cleared but not refilled — the extraction has nothing to put
  there. Cells containing a **formula are never removed**, only entered values.
  Note that emptying `B` (Lookup Key) makes the VLOOKUPs of column `R` (Material
  Consumption) show `#N/A` until new keys are entered — exactly as if you deleted
  the keys by hand. Take `B` out of `CLEAR_COLUMNS` if you want to keep them.
* **Pictures.** Each thumbnail is copied into the `Picture` cell of its row,
  scaled to fit the cell while keeping its aspect ratio, and centred. Row height
  is raised to at least 45 pt where a picture is placed. Imported pictures are
  named `BOMPIC_<row>` so `ClearBOMImport` can find them again.
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
| `CLEAR_COLUMNS`         | `A,B,C,D,E,F,G` | Part columns emptied before every import |

## Good to know about the reset

A column such as `R` (Material Consumption) or `AB` (Process Category) contains a
formula in most rows and a manually entered override in a few. The reset removes
those overrides, so the cells stay **empty** afterwards — the formula rows keep
their formulas, but a cleared override is not replaced by one. That is intended:
the macro deletes inputs, it never writes formulas.
