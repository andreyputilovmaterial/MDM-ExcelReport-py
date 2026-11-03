# Example inputs of `mdmexcelreportpy`

Each `*.mdd.json` is an example *input* (produced by `mdmreadpy` or the diff tool).
This package renders them to `.xlsx`, which is **not stored here** (binary, and
regenerated in `tests/test_fixtures.py`). All inputs are small synthetic reports —
**not real/sensitive data**.

- `a1_v1.mdd.json`, `a1_v2.mdd.json`, `b1_today.mdd.json`, `b1_tomorrow.mdd.json` —
  plain MDD reports from `mdmreadpy`.
- `report.diff.a1_-v1-v2-.mdd.json`, `report.diff.b1_-today-tomorrow-.mdd.json` —
  diff reports from the diff tool.
- `report.diff.report.diff.-a1_-b1_---v1-today---v2-tomorrow--.mdd.json` — a
  diff-on-diff report.

`mdmreportpy` keeps the matching HTML outputs for these same inputs in its own
`tests/fixtures/` folder.
