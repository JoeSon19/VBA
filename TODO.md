# VBA Todo

## Goal
- Create a VBA macro to find formulas using `XLOOKUP`, `VLOOKUP`, or `SUMIFS`.
- Replace each matching formula with its current calculated value.

## Todo
- [x] Decide scope: active sheet only or entire workbook. → Both (two macros, matching existing repo pattern).
- [x] Find formula cells with `SpecialCells(xlCellTypeFormulas)`.
- [x] Detect formulas that contain `XLOOKUP(`, `VLOOKUP(`, or `SUMIFS(`.
- [x] Make the formula-name check case-insensitive.
- [x] Skip cells that are already constants.
- [x] Replace matching formulas with values using `cell.Value = cell.Value`.
- [x] Count how many cells were converted.
- [x] Show a completion message with the total converted count.
- [ ] Test `XLOOKUP` formulas and confirm they become values.
- [ ] Test `VLOOKUP` formulas and confirm they become values.
- [ ] Test `SUMIFS` formulas and confirm they become values.
- [ ] Test non-matching formulas such as `IF`, `SUM`, and `INDEX/MATCH` to confirm they are not changed.
- [x] Decide how to handle formulas returning errors. → Skipped via `IsError` check.

## Notes
- A simple detection approach is `InStr(1, cell.Formula, "XLOOKUP(", vbTextCompare)`.
- Do the same check for `VLOOKUP(` and `SUMIFS(`.
- Existing modules in this repo already use the `cell.Value = cell.Value` hardcode pattern.
