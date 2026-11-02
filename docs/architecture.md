# MDM-ExcelReport-py — architecture & input JSON format

Stable reference. Living working-memory notes live in `docs/project-context.md` (updated
on every agent run; not linked from user-facing docs). TODO items live in
`docs/roadmap.md`.

## 1. Where it fits

`mdmexcelreportpy` is a **reporter**: it consumes the common JSON "report" shape and
emits an Excel workbook. It is a sister copy of `mdmreportpy` (HTML) — same input,
different output.

Producers of the input JSON include `mdmreadpy`, the diff tool, and the other format
readers in the global project.

## 2. Modules (`src/`)

| Module | Role |
| --- | --- |
| `src/report_create.py` | `ReportDocument` + `entry_point`. Prepares DataFrames, writes the workbook, CLI. |
| `src/launcher.py` | Dispatcher (`--program report_excel`). |
| `src/util_format_mddread_inputs.py` | Turns the JSON report into per-section rows (sanitizing cells). |
| `src/util_dataframe_wrapper.py` | Fast append-friendly wrapper around a pandas DataFrame. |
| `src/plugins.py` | "Backend" (Python) plugin hooks (currently one translation-overlay prep plugin). |
| `src/format_sheet_xlsxwriter.py` | Applies column widths / row formatting via xlsxwriter. |

## 3. Input JSON format

Same common shape as `mdmreportpy`, whose spec is the source of truth:
`docs/report-json-format.md` (report JSON) and `docs/rich-text-format.md` (rich text).
This page only notes the Excel-specific handling. Sections map to worksheets: one sheet
per section, plus an "overview" sheet built from `report_type`, `source_file`,
timestamps and `source_file_metadata`.

`column_headers` values (global `report_scheme.column_headers` merged with per-section
`section.column_headers`, section wins) provide human-readable column titles. Excel
writes them as the header row by default; `--config-column-headers fieldname` switches
to the raw column ids. Header values accept the same rich-text (`parts`/`role`) shapes
as cell values and are rendered through the same sanitize/color path.

## 4. Excel vs HTML

No JS plugins are possible (an Excel file has no JS runtime), so the frontend
interactivity of `mdmreportpy` is absent. A Python "backend" plugin concept exists only
as a thin hook (`plugins.py`) and should be extended carefully.

Excel is both worse (properties/diff not filterable, multi-line text hard to copy) and
better (usable in a wider tool ecosystem, e.g. translation-overlay workflows).
