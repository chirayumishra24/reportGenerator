# Report Generator System Architecture

## Purpose
This system ingests school result spreadsheets, validates their structure, builds a cumulative Class 10 progress view, persists exam history when a database is available, and exports polished Excel reports for staff use.

The current implementation is being rebuilt in-place around a modular architecture. Existing behavior is preserved through a compatibility layer while new work lands on cleaner boundaries.

## Product Scope
- Structured import for Class 9 baseline + Class 10 target data
- Section-wise exam imports for `HY`, `PB1`, `PB2`
- Combined `Board` import
- Cumulative student comparison sheet
- Student history and trend analysis
- Excel export with analysis sheets and charts
- Optional persistence through PostgreSQL + Prisma

## Architecture Overview
```text
React + Vite Frontend
  -> Express API
    -> Report Application Layer
      -> Workbook Parsing + Validation
      -> Cumulative Report Builder
      -> Export Generator
      -> Persistence Service
        -> Prisma
          -> PostgreSQL

Compatibility Layer
  -> Existing monolithic report engine reused while modules are extracted

Deployment Adapter
  -> api/server.js for Vercel-style serverless entry
```

## Main Runtime Components

### 1. Frontend App
- Stack: React 19 + Vite
- Responsibility: file upload workflow, validation feedback, cumulative dashboard, student drill-down, export triggers
- New direction: a small app shell and feature modules instead of a single `App.jsx`

### 2. API App
- Stack: Express 5
- Responsibility: HTTP routing, upload handling, validation responses, versioned API surface
- New direction: `app -> routes -> controllers -> services`

### 3. Report Domain
- Responsibility:
  - detect headers and exam stage
  - normalize student identity
  - calculate subject totals and percentages
  - generate cumulative rows and comparisons
  - build export-ready structures

### 4. Persistence Layer
- Stack: Prisma + PostgreSQL
- Responsibility:
  - store upload batches and exam sheets
  - upsert student records
  - store student performance across exam stages
  - produce cumulative timeline/history reports

### 5. Compatibility Layer
- Responsibility: keep the current report engine working while its logic is extracted into modules
- Location: `backend/src/legacy/report-engine.js`

## Request Flows

### Structured Parse Flow
1. Frontend uploads workbook files.
2. API receives the file via `multer`.
3. Parsing service reads workbook contents and detects headers.
4. Validation rules verify exam stage, student name, admission/enrollment number, and duplicate identifiers.
5. Frontend receives normalized sheets plus validation issues.

### Persistent Import Flow
1. Frontend uploads all structured exam files.
2. API parses and validates every file.
3. Persistence service creates upload batches, sheets, students, and performances.
4. API returns import summary plus refreshed cumulative report.

### Export Flow
1. Frontend sends normalized sheet payload.
2. Export engine writes raw sheets, analysis sheets, section comparison, and student comparison.
3. Server responds with an Excel workbook buffer.

## Data Model

### UploadBatch
- Groups a user import session

### ExamSheet
- Stores a parsed worksheet and its exam metadata

### Student
- Stores normalized identity for matching across uploads

### StudentPerformance
- Stores class 9 baseline, target, exam %, subject breakdown, and raw row data

## Code Layout Target
```text
backend/
  src/
    app.js
    config/
    controllers/
    middleware/
    routes/
    services/
    legacy/

frontend/
  src/
    app/
    features/
      report-generator/
        legacy/
        reportGenerator.config.js
```

## Incremental Build Strategy
### Phase 1
- Keep the cumulative Class 10 workflow correct and stable
- Preserve current UI behavior through the feature shell
- Expose new versioned backend routes without breaking existing `/api` routes

### Phase 2
- Extract parsing and validation out of the legacy engine
- Move frontend upload/dashboard state into smaller feature modules

### Phase 3
- Add tests around parsing, identity matching, and cumulative generation
- Split export generation into dedicated report builders

## Current Decisions
- Enrollment/admission number is the primary identity key.
- Name + section is only a fallback identity alias.
- `Target` is baseline planning data, not an exam.
- `Board` stays as one combined import.
- Compatibility is preserved while architecture is rebuilt.
