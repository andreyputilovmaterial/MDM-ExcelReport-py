# TODOs

This section captures everything that comes to mind that should ideally be updated.
Tasks listed here are not necessarily ordered by priority.

- ~~Clarify the dependency set~~ — done: `pandas` + `xlsxwriter`. `openpyxl` and
  `numpy` are not needed. Recorded in `requirements.txt` (versions intentionally
  unpinned).
- ~~Freeze the first row (header) and the first column (row names)~~ — done via
  `worksheet.freeze_panes(1, 1)` in `format_sheet_xlsxwriter.py`.
- ~~Add a `build.bat` (pinliner bundle)~~ — declined for now: no benefit here, and
  pinliner should not be copied into every submodule.
- ~~Review the unused `format_sheet_openpyxl.py` / `format_sheet_openpyxl_overview.py`~~ — removed (openpyxl dropped).
- Decide whether the "backend" plugin hook in `plugins.py` should stay a single-purpose
  hook or grow into a real plugin system.
- Add automated tests (none exist yet).
- Test against real MDD/diff JSON output and capture findings back here.

