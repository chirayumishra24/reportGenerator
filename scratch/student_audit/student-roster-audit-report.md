# Student Roster Audit Report

## Scope
- Files checked: `TARGET  SHEET 2024-25 NEW (1).xlsx`, `half-yearly_X.xls`, `pb1_X.xls`, `pb2_X.xls`, `CBSE RESULT 2026.xlsx`
- Goal: verify whether student names and student counts are consistent across the sheets in `sheets/`.

## Executive Summary
- The three internal exam workbooks (`half-yearly_X.xls`, `pb1_X.xls`, `pb2_X.xls`) are fully consistent by section, by enrollment number, and by student name. Total live roster size: `239` students.
- The CBSE workbook is internally consistent across `RANK`, `BEST FIVE`, and `All Subjects Report`. Each tab contains `239` student rows and the same name multiset after normalization.
- The target workbook is not aligned with the live roster. After removing non-student summary rows, it contains `235` students, versus `239` in the live exam/CBSE data.
- Target vs live roster overlap is only `3` names out of `235` target rows, and there are `0` shared enrollment numbers. This indicates the target workbook is a different batch/cohort, not just a section-label mismatch.

## Section Count Comparison
- `XA` vs `X-A`: target `47`, live exam `53`, delta `-6`.
- `X B` vs `X-B`: target `48`, live exam `54`, delta `-6`.
- `X C` vs `X-C`: target `56`, live exam `61`, delta `-5`.
- `X D` vs `X-D`: target `34`, live exam `32`, delta `2`.
- `X E` vs `X-E`: target `50`, live exam `39`, delta `11`.

## Important Findings
- `TARGET  SHEET 2024-25 NEW (1).xlsx` includes non-student summary rows in `XA`, `X B`, `X C`, and `X E`: `95-100`, `90-94`, `80-89`, `60-79`, `50-59`, `below 50`. These rows can inflate naive student counts by 24.
- Only three names overlap at all between the target workbook and the live roster: `AARAV JAIN, NAVYA JAIN, PARTH BHARDWAJ`. None of these share the same enrollment number, so they are not safe matches.
- Shared enrollment numbers between target and live exam files: `0`.
- Existing `baseline-mismatch-report.csv` has `232` rows, but the cleaned target workbook has `235` real student rows. The report is missing `3` target entries.
- The missing real students from `baseline-mismatch-report.csv` are `AARAV JAIN`, `NAVYA JAIN`, and `PARTH BHARDWAJ`.

## Conclusion
- If the live roster should come from the current academic year, use `half-yearly_X.xls` / `pb1_X.xls` / `pb2_X.xls` / `CBSE RESULT 2026.xlsx` as the valid source set.
- `TARGET  SHEET 2024-25 NEW (1).xlsx` should not be used as the baseline for current reconciliation unless you intentionally want last year's cohort.
- Any reconciliation logic should explicitly exclude non-student summary rows from the target workbook before counting or matching students.

## Generated Detail Files
- `section-count-summary.csv`
- `target-vs-exam-mismatches.csv`
- `baseline-report-missing-rows.csv`