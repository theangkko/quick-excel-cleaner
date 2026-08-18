---
name: excel-cleaner
description: Preserve Excel workbook semantics while safely analyzing and removing unused styles, duplicate styles, and tiny drawing objects.
license: MIT
metadata:
  version: "0.1.0"
---

# Excel Cleaner

Use this skill whenever changing workbook scanning, style cleanup, drawing cleanup, or Excel compatibility behavior.

## Style graph rules

Treat Cell, Row, and Column style references as one workbook-wide reference graph. A style index is only an implementation reference; it is not a semantic identity.

When deduplicating `cellXfs`:

1. Group styles by semantic/signature identity.
2. Select a canonical member for each duplicate group.
3. Prefer a canonical member that is referenced by the workbook; when multiple referenced members exist, use a deterministic lowest-index rule.
4. Remap every Cell, Row, and Column reference to the canonical style.
5. Remove only styles that are both non-canonical/unreferenced and permitted by cleanup options.
6. Rebuild indexes and validate all references after writing.

Do not solve a style failure by changing a test expectation until the workbook semantics and canonicalization rules have been checked.

## Drawing rules

Support `oneCellAnchor` and `twoCellAnchor`. Use the configured pixel threshold as a conservative cleanup policy. Never delete normal-size drawing objects merely because they are uncommon or unnamed.

## Workbook safety

Never modify the source in place. Create a backup, copy to the output path, validate the source copy before mutation, mutate, then validate the result. On validation failure, delete the generated output and surface the original error.

## XLSM safety

Preserve the VBA package payload. Tests should verify that `xl/vbaProject.bin` remains present and unchanged when a macro-enabled workbook is cleaned.

## Regression strategy

Every cleanup rule should have a focused regression test plus a complex workbook integration test. Preserve values and workbook features such as sheets, merged cells, hidden rows/columns, freeze panes, conditional formatting, named styles, and normal drawings.
