# Project Context — MDM-ExcelReport-py (`mdmexcelreportpy`)

> Last updated: 2026-10-23
> Status: initial capture.
>
> **Convention:** read this file at the start of every task and update it at the end,
> so it stays the permanently updated context of the current work (see
> `.clinerules/00-project-context.md`).

## 1. What this package does

Sister copy of `mdmreportpy`: it consumes the common JSON "report" shape and writes an
Excel workbook instead of an HTML page. It is the "Excel export" half of the toolchain
and is used by the broader tool set (e.g. translation-overlay workflows).

## 2. Layout & build

- Package root: `src/lib/mdmexcelreportpy`
- Source modules: `src/lib/mdmexcelreportpy/src/`
- No static assets and no `GENERATED` usage currently (see
  `.clinerules/02-python-portability.md`).

## 3. Modules

| Module | Role |
| --- | --- |
| `src/report_create.py` | `ReportDocument` + `entry_point`. |
| `src/launcher.py` | Dispatcher (`--program report_excel`). |
| `src/util_format_mddread_inputs.py` | JSON report → per-section rows (cell sanitization). |
| `src/util_dataframe_wrapper.py` | Fast append-friendly pandas DataFrame wrapper. |
| `src/plugins.py` | "Backend" (Python) plugin hooks. |
| `src/format_sheet_xlsxwriter.py` | xlsxwriter formatting. |
| `src/format_sheet_openpyxl.py` / `_overview.py` | Legacy/unused openpyxl formatters. |

## 4. Design decisions

- **Excel is written via `xlsxwriter`** (rich-text support for diff coloring).
- **`openpyxl` is still imported** in `util_dataframe_wrapper.py` for
  `get_column_letter`; the `format_sheet_openpyxl*.py` formatters are unused legacy.
- **No JS plugins** — an Excel file has no JS runtime, so the frontend interactivity of
  `mdmreportpy` is intentionally absent. A "backend" plugin concept exists only as a
  thin hook in `plugins.py` and should be extended carefully.
- **Excel is both worse and better than HTML.** Worse for filtering/copying (properties
  and diff are inline within cells); better for a wider tool ecosystem (e.g.
  translation-overlay helper).

## 5. Findings / changes (2026-10-23)

- Restructured the package into `src/` to mirror `mdmreportpy`.
- Aligned `entry_point(*argcs, **kwargs)`, error reporting (to `sys.stderr`) and console
  colors (red errors, green "saving"/"finished") with `mdmreportpy`.
- Added `src/launcher.py`, cline rules, docs and `requirements.txt`.
- Fixed the mutable `inp={}` / `config={}` defaults in `ReportDocument.__init__`.
- Clarified dependencies: `pandas`, `xlsxwriter`, `openpyxl` (unpinned); removed the
  unused `numpy` import from `plugins.py`.
- Added `worksheet.freeze_panes(1, 1)` so the header row and first column stay visible.
- Declined pinliner/`build.bat` for this submodule for now (keep pinliner only where it
  already exists).

## 6. Roadmap

See `docs/roadmap.md`.

## 7. Open questions

- Keep or remove the unused `format_sheet_openpyxl*.py` formatters.
- Whether the "backend" plugin hook in `plugins.py` should grow into a real plugin
  system.
