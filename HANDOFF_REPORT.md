# Aung Sinkha Development Project Phase 1A (EPC-01)
## Engineering Progress Tracking & Analytics Dashboard — Technical & Operational Handoff Report

**Last Updated:** 17 Aug 2026
**Version:** 2.2

---

## 1. Conceptual Overview

### 1.1 Background & Purpose
In complex Offshore & Onshore Engineering, Procurement, and Construction (EPC) projects, tracking engineering deliverables across thousands of individual documents, multiple Work Packages (WPs), and diverse engineering disciplines is critical to preventing project slippage.

The **Engineering Progress Tracking & Analytics Dashboard** was developed to replace static, manual spreadsheet reporting with an automated, high-performance, real-time web application. It ingests weekly Master Document Register (MDR) Excel cut-offs from project contractors (GCME / PJM) and transforms raw tabular records into actionable executive insights, dual variance reporting (**Rebased Line Variance** & **Non-Rebased Line**), S-Curve time-series variance analysis, granular discipline delays, procurement engineering (MR/TBE) tracking, and proactive 14-day lookahead warning alerts.

### 1.2 Target Audience & Stakeholders
- **Project Director & EPC Project Managers:** High-level visibility into overall engineering progress (Plan vs. Actual vs. Forecast/Rebaseline), dual variance indicators, Schedule Performance Indexes (SPI), and macro-level slippage.
- **Engineering Managers & Discipline Leads:** Granular tracking of document submission statuses across Process, Piping, Mechanical, Electrical, Civil/Structural, and Instrumentation disciplines.
- **Project Controls & Planning Engineers (PJM/PTTEPI):** Automated extraction and validation of weekly MDR cut-offs, S-Curve generation, and audit compliance against baseline and rebaseline milestones.
- **Document Controllers:** Real-time register monitoring, document search, and tracking of overdue reviews.

### 1.3 High-Level Architecture
- **Backend (`server.py`):** Pure Python 3.11 threaded HTTP server utilizing `openpyxl`, in-memory thread-safe caching, native Excel report generator (`/api/export/*`), RESTful API (`/api/data`, `/api/upload`, `/api/refresh`).
- **Frontend (`index.html`, `app.js`, `styles.css`):** Vanilla HTML5/CSS3/ES6 SPA with dark glassmorphism UI, 7-card executive KPI summary, Scope & Milestone filters, and powered by Chart.js for interactive data visualization.

---

## 2. Supported Excel Files & Data Model

### 2.1 Weekly File Naming Convention
Files must include a date pattern for automatic detection:

| Pattern | Example |
|:---|:---|
| `*Cut off DD-Mon-YY*` | `MM-ASK-1A-GEN01-ENG-MDR-0001_B1- Cut off 14-Aug-26.xlsx` |

`find_latest_excel()` parses all `.xlsx` filenames and automatically selects the newest date.

### 2.2 Single-Source Upload (ASK only)
The system tracks engineering progress via a single Excel file per cut-off period. (Note: Z1F tracking is explicitly out of scope for this dashboard).

| Source | File Pattern | Work Packages |
|:---|:---|:---|
| **ASK** (Aung Sinkha) | `*ASK*.xlsx` | WP01 Topside, WP01 Jacket, WP02 Topside, WP02 Jacket |

### 2.3 Input Excel Sheet Schema

| Excel Sheet | Purpose | Key Cells |
|:---|:---|:---|
| `Sum` | Overall weekly summary | `D9` = Cut-off date, `K9` = Prev cut-off, Row 14 = Rebaseline Variance, Rows 12–18 = KPIs |
| `Weekly Summary` | Discipline-level breakdown | `U7` = Cut-off date, Rows 12–42 = WP/Discipline rows, Col T = Non-Rebased Line, Col U = Rebased Line Variance |
| `Overdue Summary` | Discipline overdue matrix | `B3` = Total text, Rows 6–19 = Disciplines |
| `Overdue List` | Individual overdue documents | `J6` = Cut-off date, Rows 9+ = Documents |
| `PRO.ENGINEERING_MR TBE` | Procurement deliverables (350 MR/TBE) | Rows 12+ = Documents across 4 WPs; Col V-Y = IFA (A1), Col AH-AK = IFA (C1) |
| `Eng. S-Curve_OVERALL` | Overall S-Curve time-series | Row 41 = Dates, Row 43/45/47 = Plan/Forecast/Actual Cum %, Rows 12-16 = Snapshot |
| `Eng. S-Curve_WP01 TOPSIDE` | WP01 Topside S-Curve | Same layout as OVERALL, Rows 12-22 Snapshot |
| `Eng. S-Curve_WP01 JACKET` | WP01 Jacket S-Curve | Same layout as OVERALL, Rows 12-15 Snapshot |
| `Eng. S-Curve_WP02 TOPSIDE` | WP02 Topside S-Curve | Same layout as OVERALL, Rows 12-22 Snapshot |
| `Eng. S-Curve_WP02 JACKET` | WP02 Jacket S-Curve | Same layout as OVERALL, Rows 12-15 Snapshot |

### 2.4 Dual Variance & Cut-off Reference
- **Rebased Line Variance** (`+0.30%`): `Actual Progress - Rebaseline/Forecast Progress` (from `Weekly Summary` Col U & `Sum` Row 14).
- **Non-Rebased Line** (`-5.46%`): `Actual Progress - Original Baseline Plan Progress` (from `Weekly Summary` Col T & `Sum` Actual vs. Plan).
- **Cut-off Date**: Strictly read from Excel cell `Sum!D9` (`2026-08-14`).

---

## 3. Requirements & Compliance Matrix

| ID | Requirement | Specification | Solution | Status |
|:---:|:---|:---|:---|:---:|
| REQ-01 | Reference Date Baseline | All tracking dates: Today = cut-off D9; Forecast Date > Plan Date fallback | `compute_delay_and_lookahead()` uses `D9` date, not `datetime.date.today()` | PASS |
| REQ-02 | Empty Date Handling | Blank dates = Not Yet Issued | None/""/N/A/"-" all treated as unissued | PASS |
| REQ-03 | Type 3.1 Delay | Submit Date > Forecast Date | `(submit_dt - ref_dt).days` | PASS |
| REQ-04 | Type 3.2 Delay | Today >= Forecast Date, not submitted | `(today - ref_dt).days` | PASS |
| REQ-05 | 14-Day Lookahead | Today < Forecast <= Today+14 | Lookahead window from cut-off D9 date | PASS |
| REQ-06 | Filters | Filter by WP, Discipline, Scope, Milestone | Interactive header filters, instant re-render | PASS |
| REQ-07 | S-Curve Time-Series | Proper cumulative S-shape per WP, updates weekly | 80+ weekly points from Row 41 dates + Rows 43/45/47 cumulative % | PASS |
| REQ-08 | UI/UX | Premium dark glassmorphism, responsive | Deep navy (#030712), 7 KPI cards, Chart.js, micro-animations | PASS |
| REQ-09 | Dynamic Weekly Column | S-Curve auto-expands as new weeks added | Scans `ws.max_column` dynamically — zero hardcoded limits | PASS |
| REQ-10 | Dual Variance Metrics | Rebased Line Variance vs Non-Rebased Line | Implemented across Overall Summary, WP Summary, Weekly, and S-Curve tabs | PASS |
| REQ-11 | Procurement Deliverables | Ingest `PRO.ENGINEERING_MR TBE` (350 MR/TBE items) | Integrated into Document Status, Overdue & Look-ahead, and Exports (1,789 Total) | PASS |

---

## 4. Engineering Solutions & Implementation

### 4.1 S-Curve Time-Series Extraction (`extract_scurve` in server.py)
All 5 S-Curve sheets share identical row structure:

```
Row 41  — Ordinal Week dates (datetime), columns E+ (col 5+)
Row 42  — Plan Incremental
Row 43  — Plan Cumulative        <-- chart series
Row 44  — Forecast Incremental
Row 45  — Forecast Cumulative    <-- chart series
Row 46  — Actual Incremental
Row 47  — Actual Cumulative      <-- chart series
Rows 12-16 — Current-week summary snapshot (for table only)
```

Dynamic column scanning (no hardcoded limits):
```python
for col in range(START_COL, ws.max_column + 1):
    cell_val = ws.cell(row=DATE_ROW, column=col).value
    date_str = safe_date(cell_val)
    if date_str:
        date_cols.append((col, date_str))
```

### 4.2 Delay & Lookahead Algorithm
```python
today = datetime.date.fromisoformat(cutoff_date_str)  # from Sum D9, not system clock
lookahead_end = today + datetime.timedelta(days=14)
ref_dt = forecast_dt if forecast_dt else plan_dt

if submit_dt:
    if submit_dt > ref_dt:              # Type 3.1: Submitted Late
        delayed.append(...)
else:
    if ref_dt <= today:                 # Type 3.2: Overdue, Not Submitted
        delayed.append(...)
    elif today < ref_dt <= lookahead_end:  # Lookahead Risk Window
        lookahead.append(...)
```

### 4.3 SPI Formulas
- SPI_Weekly = Actual_This_Week / Plan_This_Week
- SPI_Cumulative = Actual_Cumulative / Plan_Cumulative
- SPI > 1.00: Ahead of schedule | SPI = 1.00: On schedule | SPI < 1.00: Behind schedule

---

## 5. Operational Guide

### 5.1 Server Startup
```powershell
cd "c:\Users\pipes\OneDrive\Documents\Google_AntiGravity\Project\Engineering_Progress"
python server.py
# Open: http://localhost:8081
```

### 5.2 Weekly Update Workflow
1. Upload new `.xlsx` file via the Upload button in the dashboard header, OR drop it into the project folder.
2. Click "Update Data" button (or the server auto-detects on next API request).
3. Dashboard refreshes: updated cut-off date, S-Curves, delay counts, lookahead.

> TIP: Always include the cut-off date in the filename (e.g., "Cut off 10-Jul-26") so `find_latest_excel()` can auto-select the correct file.

### 5.3 Audit Verification
```powershell
python -c "import server; d=server.extract_all_data(); print('Cutoff:', d['summary']['cutoff_date'], '| Delayed:', d['delay_lookahead']['delayed_count'], '| Lookahead:', d['delay_lookahead']['lookahead_count'])"
```

---

## 6. Deliverables Checklist

| File | Status | Description |
|:---|:---:|:---|
| `server.py` | DONE | Thread-safe backend, S-Curve time-series, Type 3.1/3.2 delay, dynamic column scanning |
| `app.js` | DONE | S-Curve rendering from time-series, summary table, 4-digit year dates |
| `index.html` | DONE | 5-tab SPA: Summary, Work Packages, Documents, Overdue Analysis, S-Curve |
| `styles.css` | DONE | Glassmorphism dark mode, full filename badge, responsive layout |
| `HANDOFF_REPORT.md` | DONE | This document (v2.0) |
| `handoff.md` | DONE | Agent-facing quick handoff |
| `spec.md` | DONE | System specification |
| `lesson_learn.md` | DONE | Lessons learned & known gotchas |
