# VBA Todo

## Goal
- Create a VBA macro to find formulas that reference another worksheet/tab.
- If the formula result is a number, replace the formula with its current value.

## Todo
- [x] Decide scope: active sheet only or entire workbook. → Both (two macros, matching existing repo pattern).
- [x] Find cells with formulas using `SpecialCells(xlCellTypeFormulas)`.
- [x] Detect formulas that reference another tab/worksheet.
- [x] Check whether the evaluated result is numeric.
- [x] Convert matching formulas to values with `cell.Value = cell.Value`.
- [x] Skip formulas that return text, blanks, or errors.
- [x] Count how many cells were converted.
- [x] Show a completion message with the total converted count.
- [ ] Test with formulas that reference the same sheet to confirm they are not changed.
- [ ] Test with formulas from other tabs that return numbers to confirm they are hardcoded.

## Notes
- Likely formula pattern to detect: sheet references such as `Sheet2!A1`.
- Existing macros in this repo already use a similar hardcode-to-value pattern.
