# XLSM preservation

A cleanup operation on `.xlsm` must not remove the VBA project package.

Minimum regression evidence:

- `xl/vbaProject.bin` exists before and after cleanup.
- The VBA package bytes are unchanged by the cleanup operation.
- Workbook and worksheet parts remain readable after cleanup.

The test must focus on preservation, not VBA execution.
