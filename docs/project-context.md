# Project Context — MDM-ExcelReport-py (`mdmexcelreportpy`)

> Last updated: 2026-11-03
> (2026-11-03) Launcher smoke test: added a dependency-free ``--program done`` (prints
> ``done`` in green) to ``src/launcher.py`` plus ``tests/test_done_program.py``.
> ``report_create`` (and therefore pandas/xlsxwriter) is now imported lazily via
> ``_import_module`` so ``done`` loads without any optional dependencies.
> (2026-11-02) Packaging/docs: `requirements.txt` already lists `pandas`/`xlsxwriter`;
> documented the three launch strategies in `README.md`. No `build.bat`/`run_init` yet
> (pinliner bundling is still declined for this submodule — see §5).
> Status: `column_headers` support added (human-readable headers by default, raw field
> names via a CLI flag); tests added.
> (2026-11-02) Implemented `column_headers`: Excel now writes human-readable headers by
> default and raw column ids via `--config-column-headers fieldname` (see §4/§5).
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

## 4. Design decisions

- **Excel is written via `xlsxwriter`** (rich-text support for diff coloring).
- **No JS plugins** — an Excel file has no JS runtime, so the frontend interactivity of
  `mdmreportpy` is intentionally absent. A "backend" plugin concept exists only as a
  thin hook in `plugins.py` and should be extended carefully.
- **Excel is both worse and better than HTML.** Worse for filtering/copying (properties
  and diff are inline within cells); better for a wider tool ecosystem (e.g.
  translation-overlay helper).
- **Column headers (decision).** The header row is now configurable via
  `--config-column-headers {text|fieldname}`. `text` (default) resolves each column id
  through the merged `column_headers` (global `report_scheme.column_headers`, with
  per-section `section.column_headers` winning, falling back to the field name);
  `fieldname` uses the raw JSON field names. Header values accept the same rich-text
  (`parts`/`role`) shapes as cell values and are rendered through the existing
  `sanitize` + `write_rich_string` path, so diff coloring also applies to headers.
- **Source of truth for the data format is `mdmreportpy`.** Its
  `docs/report-json-format.md` and `docs/rich-text-format.md` are the master definition
  of the report JSON shape and the rich-text format; this package defers to them.

## 5. Findings / changes (2026-10-23)

- Restructured the package into `src/` to mirror `mdmreportpy`.
- Aligned `entry_point(*argcs, **kwargs)`, error reporting (to `sys.stderr`) and console
  colors (red errors, green "saving"/"finished") with `mdmreportpy`.
- Added `src/launcher.py`, cline rules, docs and `requirements.txt`.
- Fixed the mutable `inp={}` / `config={}` defaults in `ReportDocument.__init__`.
- Clarified dependencies: `pandas`, `xlsxwriter` (unpinned); removed the unused `numpy`
  import from `plugins.py`.
- Dropped `openpyxl` entirely: removed the dead `get_column_letter` helper and the
  unused `format_sheet_openpyxl*.py` formatters (xlsxwriter's `write_rich_string`
  already covers the diff coloring).
- Added `worksheet.freeze_panes(1, 1)` so the header row and first column stay visible.
- Declined pinliner/`build.bat` for this submodule for now (keep pinliner only where it
  already exists).

## 5b. Findings / changes (2026-11-02)

- Added `resolve_column_headers` to `util_format_mddread_inputs.py`; `prep_dataframes`
  now returns `columns` and a resolved `column_headers` mapping per section, and
  `write_to_file` writes the header row from that mapping (sanitizing rich-text values).
- Added `--config-column-headers {text|fieldname}` (default `text`) to `report_create.py`
  and plumbed it via `config['column_headers_mode']`.
- Added `tests/test_column_headers.py` (10 `unittest` cases; run with
  `python -m unittest discover -s tests`). Verified end-to-end that the `.xlsx` header
  row renders `['Item name', 'Item label', 'English']` in `text` mode and
  `['name', 'label', 'langcode-ENU']` in `fieldname` mode.

## 6. Roadmap

See `docs/roadmap.md`.

## 7. Open questions

- Whether the "backend" plugin hook in `plugins.py` should grow into a real plugin
  system.
