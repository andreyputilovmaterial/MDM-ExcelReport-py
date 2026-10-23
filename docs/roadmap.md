# TODOs

This section captures everything that comes to mind that should ideally be updated.
Tasks listed here are not necessarily ordered by priority.

- Clarify the dependency set: `xlsxwriter`, `openpyxl`, or both. Currently
  `report_create.py` writes with xlsxwriter, `util_dataframe_wrapper.py` uses
  `openpyxl.utils.get_column_letter`, and `format_sheet_openpyxl*.py` are unused legacy.
- Freeze the first row (header) and the first column (row names) in the output Excel so
  they stay visible when scrolling.
- Review the unused `format_sheet_openpyxl.py` / `format_sheet_openpyxl_overview.py`
  and either wire them up or remove them.
- Decide whether the "backend" plugin hook in `plugins.py` should stay a single-purpose
  hook or grow into a real plugin system.
- Add automated tests (none exist yet).
- Add a `build.bat` (pinliner bundle) for this submodule, mirroring the sibling repos.
- Test against real MDD/diff JSON output and capture findings back here.
