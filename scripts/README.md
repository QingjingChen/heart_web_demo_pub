# Workbook Scripts

This directory keeps only the final workbook scripts from the working folder.
Earlier one-off scoring, calibration, backup, and raw-output scripts were not
carried into the public demo repository.

The workbook these scripts operate on (`workbooks/科技伦理toolkit.xlsx`) is
mirrored from the canonical resource repository:
<https://github.com/QingjingChen/Heart>. CSV exports of every
sheet (104-row benchmark distribution, 48-row tool toolbox, 14-rubric
anchors, 140-row diagnosis table, citation verification) live under
`workbook/exports/` in that repository and are usually more convenient than
re-running these extraction scripts from scratch.

## Kept

- `extract_lineage_views.py` builds lineage and external-view additions with
  Qwen/DashScope.
- `write_lineage_xlsx.py` renders `lineage_views.json` into a companion lineage
  workbook.
- `extract_per_rubric.py` maps paper/tool evidence onto the 14 HEART rubrics.
- `write_per_rubric_xlsx.py` applies `per_rubric.json` to
  `workbooks/科技伦理toolkit.xlsx`.
- `lineage_map*.json`, `lineage_views.json`, and `per_rubric.json` are the
  curated/supporting data needed by the final scripts.

## Not Kept

The old `score_*`, `calibrate_*`, `write_*_diagnosis.py`, raw `.jsonl` model
outputs, workbook backups, logs, virtualenv files, and local credentials were
working artifacts. They are intentionally omitted from this repo.

The extraction scripts expect a local DashScope key at `~/.dashscope_key` and,
for `extract_lineage_views.py`, the survey PDF named
`A_Comprehensive_Survey_of_AI_Ethics_Benchmarks (6).pdf` at the repository root.
