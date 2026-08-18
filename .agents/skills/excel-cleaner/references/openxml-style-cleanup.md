# Open XML style cleanup

`cellXfs` indexes can change after cleanup. Tests should compare the resulting `CellFormat` semantics rather than assuming a particular numeric index survives.

The cleanup algorithm must account for styles referenced by cells, rows, and columns. Duplicate groups should be canonicalized before unreferenced styles are discarded, otherwise a referenced duplicate can incorrectly become canonical while the earlier canonical candidate is removed.
