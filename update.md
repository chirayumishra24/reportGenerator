# Update Log

## Latest Architecture Iteration
- Added a first-pass system architecture doc at `docs/system-architecture.md`.
- Rebuilt the backend entry around a modular shell:
  - `backend/src/app.js`
  - `backend/src/routes/*`
  - `backend/src/controllers/*`
  - `backend/src/services/*`
  - `backend/src/middleware/*`
- Preserved the working report logic behind a compatibility layer:
  - `backend/src/legacy/report-engine.js`
- Added versioned modular endpoints without breaking the current API surface:
  - `GET /api/v1/health`
  - `POST /api/v1/reports/parse`
  - `POST /api/v1/reports/import-persistent`
  - `GET /api/v1/reports/cumulative`
- Rebuilt the frontend entry around a modular shell:
  - `frontend/src/app/*`
  - `frontend/src/features/report-generator/*`
- Moved the current large report-generator implementation behind:
  - `frontend/src/features/report-generator/legacy/LegacyReportGeneratorApp.jsx`
- Extracted shared frontend runtime constants into:
  - `frontend/src/features/report-generator/reportGenerator.config.js`
- Replaced root entry files with thin wrappers so new work can land on stable boundaries instead of the old monoliths.
- Verification completed:
  - backend app loads successfully through the new server entry
  - frontend production build succeeds
  - Vite reports the existing Node version warning (`22.4.0` vs recommended `22.12+`)

## Latest Backend Extraction Pass
- Used the local repo skill guidance from:
  - `skills/senior-fullstack/SKILL.md`
  - `skills/nodejs-backend-patterns/SKILL.md`
  - `skills/nodejs-backend-patterns/resources/implementation-playbook.md`
- Added a shared Prisma bootstrap at:
  - `backend/src/config/database.js`
- Extracted workbook parsing and validation into:
  - `backend/src/services/parser.service.js`
- Extracted cumulative report shaping into:
  - `backend/src/services/cumulative.service.js`
- Extracted persistent import and DB-backed report logic into:
  - `backend/src/services/persistence.service.js`
- Updated the modular v1 API service/controller/routes to use the extracted services instead of calling the legacy engine directly for:
  - parse
  - persistent import
  - cumulative report
  - db status
  - student history
- Kept the current legacy app working and aligned by pointing it at the shared DB config.
- Fixed a baseline-stage detection gap so target-sheet style uploads are recognized as `BASELINE` even when only the exam/file name is available.
- Added backend tests using Node's built-in test runner:
  - `backend/test/parser.service.test.js`
  - `backend/test/cumulative.service.test.js`
- Backend verification now passes:
  - `npm test`
  - backend app boot check via `node -e "require('./server')"`

## Latest No-DB Adjustment
- Core workflow no longer auto-uses the database when it is not required.
- Added a dedicated in-memory structured import path at:
  - `POST /api/v1/reports/structured-import`
- The new structured-import API:
  - parses all uploaded structured files in memory
  - validates them without persistence
  - returns merged sheets plus a cumulative Class 10 sheet payload
- Added in-memory cumulative workbook building to:
  - `backend/src/services/cumulative.service.js`
- Updated the frontend structured upload flow to use the new non-DB endpoint by default:
  - `frontend/src/features/report-generator/legacy/LegacyReportGeneratorApp.jsx`
- Phase 1 frontend no longer auto-checks DB status on load.
- Generic file upload no longer auto-persists files to the database.
- Database-backed endpoints remain available only for features that truly need persistence:
  - cumulative DB reporting
  - db status
  - student history
- Verification after this change:
  - backend tests pass
  - frontend production build passes

## Latest Extraction Accuracy Fix
- Exam-stage detection now normalizes teacher-style file names such as:
  - `HY_10A_X-A.xls`
  - `PB1_10B_X-B.xls`
  - `PB2_10C_X-C.xls`
  - `BOARD_CLASS10_RESULT.xls`
- This fixes the validation failure where structured imports showed:
  - `Could not detect exam stage for Sheet1`
- Cumulative-sheet generation continues to use enrollment number as the primary key.
- Name differences across sheets are now treated as non-authoritative when the enrollment number matches.
- The baseline/first matched record name remains the display name when later exam sheets contain teacher-edited variants.

## Latest Rankwise Handling
- Rankwise result sheets are now supported during extraction.
- The parser now tolerates rank-oriented columns such as `Rank` while still extracting the student data rows.
- Aggregate/non-student rows in rankwise sheets are filtered out during parsing, including rows like:
  - `Average`
  - `Summary`
  - repeated header rows
  - topper/pass/fail style summary labels
- Cumulative merge logic was tightened so enrollment number is the authoritative identity when present.
- Name-based alias merging is now only used as a fallback when enrollment data is missing.

## Latest Class 9 & Target Score Fix
- Resolved issue where Class 9 Percentage and Target Percentage were not appearing in student reports.
- Root causes fixed:
  - Destructive Overwriting: Removed logic in `recalcGrandTotal` (frontend & backend) that was overwriting the `% in IX` column with current exam percentages.
  - Strict Merging: Updated `buildClass10CumulativeSheet` to extract baseline data (Class 9/Target) from any sheet where it exists, rather than only from sheets labeled "BASELINE".
  - Rigid Column Detection: Expanded `findClass9Column` and `findTarget100Column` to support more variants like "IX Percent", "Target %", "IX Marks", and "9th %".
- Improved Target detection to ignore 'IX' in fuzzy matches to avoid confusion between Class 9 baseline and Class 10 Target.
- Verified that cumulative sheets correctly merge baseline data from any part of the workbook.

## Latest Dashboard & Navigation Improvements (2026-04-22)
- **Sidebar-Based Navigation**: Replaced horizontal sheet tabs with a professional vertical sidebar for managing uploaded workbooks. This layout provides a clear overview of all sheets and separates workbook navigation from sheet content.
- **Improved Header Row Detection**: Updated both frontend and backend `detectHeaderRowIndex` with weighted scoring logic. The system now scans up to 12 rows and prioritizes "Class 9 %" and "Target" columns, ensuring baseline sheets are correctly identified even when they lack traditional subject marks.
- **Cumulative Generation Workflow**: Added a 'Combine Sheets' shortcut in the dashboard sidebar to facilitate the merging process of current session data.
- **Normalization Persistence**: Synchronized data scaling (0.xx to xx.x%) and 100% capping on Target metrics across all layers of the application.
- **Roster Alignment**: Enforced that Enrollment Number matching is the authoritative student identity in cumulative merges, ensuring exact matches even when names are inconsistently formatted across source files.

## Cumulative Sheet Architecture
- A cumulative class 10 sheet is generated by enrollment-first matching.
- Exclude "Rank", "S.No", and other index/ranking columns from cumulative data rows.
- The cumulative sheet is sorted by Section (A to E) and then Enrollment No.
- Relevant Files:
  - `frontend/src/features/report-generator/legacy/LegacyReportGeneratorApp.jsx`
  - `backend/src/services/cumulative.service.js`
  - `backend/src/services/parser.service.js`
  - `frontend/src/index.css`
  - `update.md`

## Notes for future tasks
- The cumulative sheet should always be sorted by Section then Enrollment No.
- Matching logic priority: Enrollment No -> Roll No -> Name + Section.
- Exclude "Rank", "S.No", and other index/ranking columns from cumulative data rows to prevent clutter in merged reports.
- Target percentage is strictly capped at 100%.



## Current Product Context
- The app now supports a structured Class 10 workflow:
  - one baseline sheet for `Class 9 + Target`
  - section-wise uploads for:
    - `HY`
    - `PB1`
    - `PB2`
  - one combined upload for:
    - `Board`
  - sections expected: `10A` to `10E`
- A cumulative class 10 sheet is generated by enrollment-first matching.
- Header detection now supports sample sheets like `X-A.xls`, where the top row is a title row and the second row is the actual header row.
- PostgreSQL + Prisma groundwork has been added in `backend/`.

## Recommended Next Steps
1. Phase 1:
   - make section-wise cumulative sheet fully correct
   - validate HY / PB1 / PB2 / Board percentage filling
   - keep the result limited to one cumulative view
2. Phase 2:
   - add export of only the finalized cumulative sheet
   - add cumulative sheet cleanup and formatting controls
3. Phase 3:
   - add student-level comparison reports and section-vs-section reporting
4. Phase 4:
   - add graphs and richer analytics only after cumulative accuracy is locked

## Notes for future tasks
- Use enrollment/admission number as the primary identity whenever possible.
- Treat `Target` as a baseline goal field, not as an exam.
- Keep the cumulative sheet read-only.
- `Board` is a single combined upload, so UI logic must stay null-safe for non-section-wise stages.
