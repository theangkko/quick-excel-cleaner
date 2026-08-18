# Excel regression fixtures

Store representative workbooks here when a fixture is more stable and readable as a file than as generated Open XML code.

Recommended catalog:

- `basic.xlsx` — minimal readable workbook.
- `duplicate-styles.xlsx` — duplicate `cellXfs` with referenced and unreferenced members.
- `unused-styles.xlsx` — unreferenced styles that are safe to remove.
- `row-column-styles.xlsx` — styles applied through row and column defaults.
- `tiny-drawings.xlsx` — 1–2px objects mixed with normal drawings.
- `complex-workbook.xlsx` — multiple worksheets and combined cleanup targets.
- `feature-preservation.xlsx` — merged cells, hidden rows/columns, freeze panes, conditional formatting, and named styles.
- `macro-enabled.xlsm` — a macro-enabled package whose `xl/vbaProject.bin` must survive unchanged.

Each fixture should have a corresponding test that documents the feature it represents and the expected cleanup/preservation behavior.

Generated fixtures may remain in test code when the API surface is intentionally under test; binary fixtures are preferred for regression cases where SDK object construction itself is not the behavior being tested.
