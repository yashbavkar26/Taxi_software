# Excel Workbook Generation Rules

Follow these rules unconditionally whenever creating or modifying Excel files or generators using `openpyxl`, `xlsxwriter`, or other Python libraries:

## 1. Formula Syntax & Parentheses Validation
- **Strict Parentheses Matching**: Never write a formula string without verifying that opening and closing parentheses are strictly equal (`count('(') == count(')')`). An off-by-one parenthesis error immediately corrupts the sheet XML and triggers Excel's "We found a problem with some content..." recovery prompt.
- **Concise Logic**: Avoid deeply nested (>5 level) `IF` statements in formula strings. Use clean vectorized or arithmetic formulas (e.g. `IF(H2<=80, J2*100, J2*100 + (H2-80)*J2)`) instead of cascading manual conditions.

## 2. Modern OpenXML Future Functions (`_xlfn.`)
- **Required Prefix**: In openpyxl, functions introduced in Excel 2019 / Microsoft 365 (`FILTER`, `UNIQUE`, `SORT`, `XLOOKUP`, `XMATCH`, `SEQUENCE`, etc.) MUST be prefixed with `_xlfn.` in formula strings (e.g., `_xlfn.FILTER(...)`).
- Without `_xlfn.`, Excel treats the function as an unknown legacy token and strips the formula upon opening, prompting a file recovery warning.

## 3. Data Validation Dropdowns
- When configuring `openpyxl.worksheet.datavalidation.DataValidation(type="list", ...)`:
  - Do not pass `showDropDown=False` (which produces `showDropDown="0"` in XML and can create schema validation conflicts in desktop Excel).
  - Ensure string literal lists are enclosed in standard inner quotes: `'"Item1,Item2,Item3"'`.
