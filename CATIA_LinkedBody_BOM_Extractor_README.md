# CATIA_LinkedBody_BOM_Extractor — BOM from a CATPart of linked bodies

Some parts are not an assembly but a **single CATPart** whose tree is a flat
list of bodies that were pasted with link from an assembly. Every body then
carries the instance path of its source part in its name:

```
2DP000064819.1\2DP000033345.1\2DP000033352.1\2DP000033416.1\2DP000034947.1\2DP000034948.1\PartBody
2DP000064819.1\2DP000033345.1\2DP000033352.1\2DP000033416.1\2DP000033444.1\Inner-Winding
2DP000064819.1\2DP000033345.1\2DP000033352.1\2DP000033416.1\2DP000033444.1\PANEL-INTERFACE-WINDING.1
…
```

`CATIA_LinkedBody_BOM_Extractor.bas` turns that into **exactly the same sheet**
as the assembly extractor (`CATIA_BOM_Extractor_Fast_MultiBody.bas`), so the
result can be imported into the ACE cost model with `ImportBOM` unchanged:

| A | B | C | D | E | F | G | H | I | J | K | L | M |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Thumbnail | Level | Part Number | Description | Qty | Mass (kg) | Volume (m3) | Area (m2) | Length (mm) | Width (mm) | Height (mm) | Material | Density (kg/m3) |

## How the body names are read

| Column | Taken from |
|--------|------------|
| `Level` | number of part numbers in the path — the first example above gives **6** |
| `Part Number` | deepest part number, instance suffix removed: `2DP000034948.1` → `2DP000034948` |
| `Description` | the body name itself: `PartBody`, `Inner-Winding`, `PANEL-INTERFACE-WINDING` |
| `Qty` | how many identical bodies were merged into that row |

Bodies whose name has no path (normally modelled bodies) are reported on
`Level 1` with the part number of the CATPart itself.

**Merging:** bodies with the same source path, the same body name (ignoring the
trailing `.1`, `.2`, …) **and** the same mass and volume become one row with
`Qty > 1`. That is what turns `PANEL-INTERFACE-WINDING.1 … .6` into a single row
with Qty 6. Bodies that only share a name but differ in geometry stay separate
rows, because mass and volume are part of the comparison.

## Measurements

Per body, with the same routines and the same fallbacks as the assembly
extractor:

* **Mass / density** — `SPAWorkbench.Inertias`
* **Volume / area** — `SPAWorkbench.GetMeasurable` (m³ / m²)
* **Material** — material on the body, part material as fallback
* **Length / Width / Height** — bounding box (`AddNewBoundingBox`), then
  axis-aligned extremum points, then the equivalent box from the inertia;
  always sorted large → small
* **Thumbnail** — all bodies are hidden once, then each body alone is shown,
  reframed and captured

## Usage

1. Open the CATPart in CATIA.
2. `Tools → Macro → Macros…` → select `ExtractLinkedBodyBOM` → **Run**.
   (Install the file once with `Tools → Macro → Macros… → Macro libraries…`,
   or paste it into a new module of your VBA project.)
3. Excel opens with the finished BOM sheet.
4. Save it as `.xlsx`, then import it into the ACE workbook with `ImportBOM`.

The macro works on a **CATPart**. For a CATProduct use `GenerateMasterBOM` from
`CATIA_BOM_Extractor_Fast_MultiBody.bas`.

While it runs, the spec tree and the compass are hidden and every body is
hidden except the one being measured. Both are restored at the end, also when
a single body fails to measure.

## Settings

Constants at the top of the module:

| Constant | Default | Meaning |
|----------|---------|---------|
| `GROUP_IDENTICAL_BODIES` | `True` | merge equal bodies into one row with `Qty` |
| `GROUP_BY_GEOMETRY` | `True` | only merge when mass **and** volume match |
| `CAPTURE_THUMBNAILS` | `True` | `False` = no pictures, much faster |
| `INCLUDE_GEOMETRICAL_SETS` | `False` | also list surface sets (no volume/mass) |
| `WRITE_ROOT_ROW` | `False` | extra `Level 1` row for the CATPart itself |
| `WRITE_SOURCE_PATH` | `False` | full body name into an extra column N |
| `TEMP_FOLDER` | `C:\Temp` | folder for the temporary screenshots (deleted afterwards) |

## Coexistence with the other macros

All helper routines are `Private` and prefixed with `LB`, and `Sleep` is
declared `Private`, so the module can live in the same VBA project as
`CATIA_BOM_Extractor_Fast_MultiBody.bas` and the other extractors without any
"ambiguous name" error. The only public entry point is `ExtractLinkedBodyBOM`.
