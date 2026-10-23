# MDM-ExcelReport-py (`mdmexcelreportpy`)

Renders the common JSON "report" format to an Excel workbook (`.xlsx`). It is a sister
copy of `mdmreportpy`, with one key difference: the output is **Excel**, not an HTML
page.

Input JSON is produced by the same readers and the diff tool that feed `mdmreportpy`
(e.g. the MDD reader `mdmreadpy`, the diff tool, and the other format readers).

## What it does

- Reads a JSON "report" file (see `docs/architecture.md` for the format).
- Flattens each section into a pandas DataFrame and writes one sheet per section (plus
  an "overview" sheet) into a single `.xlsx` file.
- Applies light coloring to diff marks (added / removed / changed) via rich text.

## Requirements

- Python 3.
- `pandas`, `xlsxwriter` (see `requirements.txt`; versions are intentionally unpinned).

## Usage

Direct (during development):

```
python src/report_create.py --inpfile path/to/report.json
```

Or through the universal launcher:

```
python src/launcher.py --program report_excel --inpfile path/to/report.json
```

Options:

- `--output-format excel` — the default and the only implemented format here.
- `--flags LIST` — comma-delimited flags that toggle behavior (e.g. the
  translation-overlay preparation plugin is enabled by a flag).

## No plugins (unlike `mdmreportpy`)

`mdmreportpy` ships a set of JavaScript plugins for interactivity (show/hide columns,
filters, expandable sections, …). None of that is possible here — an Excel file has no
JS runtime, so there is no frontend functionality to implement.

A "backend" (Python) plugin concept is *theoretically* possible (transforms applied
before the workbook is written), but it should be approached carefully and only
introduced if it is actually needed.

## Excel: worse and better than HTML

**Worse** — people love Excel, but for this "report" it does not bring benefits:

- Custom properties are shown inside a cell, so they cannot be filtered on.
- Text can be multi-line, which makes it harder to copy out of Excel (the number of
  rows is not the number of lines in the copied content).
- Diff is shown via inline colors/classes, so it cannot be used in filters.

For those purposes the Excel form is effectively useless.

**Better** — it can be used in a much broader set of tools and has a wider usage
pattern. For example, the translation-overlay helper (the sibling
`MDM-TranslationOverlayHelper`) consumes this Excel output.

## Documentation

- [Architecture & input JSON format](docs/architecture.md) — stable reference.
- [Roadmap / TODOs](docs/roadmap.md) — upcoming work.

