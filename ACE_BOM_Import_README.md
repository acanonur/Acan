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

Everything else on the `Input` sheet is left untouched — `B` Lookup Key, the
cost columns, factory/country, material, process and every formula stay exactly
as they are, so the cost model keeps working after an import.

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
| `ClearBOMImport`       | Removes imported values and pictures from `Input` again          |

A file dialog asks for the extraction file (it starts in the folder of the ACE
workbook). If that file is already open in Excel it is reused and left open;
otherwise it is opened read-only and closed again straight away.

At the end a summary reports how many rows and pictures were imported and how
long it took.

## Behaviour details

* **Re-importable.** Every run first clears the previously imported values
  (columns A, C, D, G) and all pictures in column F of the data area, so
  importing a second BOM never leaves leftovers of the first one.
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
