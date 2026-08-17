# Handoff: Engineering Progress Tracking & Analytics Dashboard (Phase 1A EPC-01)
**Last Updated:** 17 Aug 2026 | **Version:** 2.4

---

## 1. Executive Summary
- **Objective**: Maintain and enhance the real-time Engineering Progress Tracking & Analytics Dashboard replacing static spreadsheet reporting (MDR Excel cut-offs).
- **Current Status**: Fully operational. Ingests `MM-ASK-1A-GEN01-ENG-MDR-0001_B1- Cut off 14-Aug-26.xlsx`. Supports dual variance tracking: **Rebased Line Variance** (`+0.30%`) and **Non-Rebased Line** (`-5.46%`). Ingests all 350 MR/TBE procurement engineering deliverables from `PRO.ENGINEERING_MR TBE` alongside standard engineering deliverables (Total: **1,789 deliverables**). Implements strict **PENDING FINAL APPROVAL** (`400`: `ENG 262 / MR 116 / TBE 22`) and **COMPLETE** (`398`: `ENG 301 (20.9%) / MR 53 (30.3%) / TBE 44 (25.1%)`) tracking.
- **Recipient Action**: Use `HANDOFF_REPORT.md` as the technical source of truth. Use `spec.md` for system behaviour specification. Use `lesson_learn.md` for known gotchas.

---

## 2. Architecture

| Layer | File | Description |
|:---|:---|:---|
| Backend | `server.py` | Python 3.11 threaded HTTP server, openpyxl Excel parser, in-memory cache, REST API, native Excel export generation |
| Frontend JS | `app.js` | Vanilla ES6 SPA controller, Chart.js S-Curve & bar charts, dual variance renderers, MR/TBE badge and filters, PENDING FINAL APPROVAL and COMPLETE tracking |
| Frontend HTML | `index.html` | 5-tab shell: Summary / Work Packages / Documents / Overdue / S-Curve |
| Frontend CSS | `styles.css` | Glassmorphism dark mode, HSL accent palette, responsive layout, MR/TBE badges, approval & complete status badges |

**API Endpoints:**
- `GET /api/data` — Returns full JSON payload (summary with dual variances, 1,789 documents, overdue, S-Curves, delays, lookahead)
- `POST /api/upload` — Accepts multipart `.xlsx` upload, caches data, invalidates old cache
- `GET /api/refresh` — Forces re-read from disk of latest Excel file
- `GET /api/export/documents`, `/api/export/lookahead`, `/api/export/delayed` — Native Excel export downloads

---

## 3. Key Design Decisions

### Dual Variance Reference
- **Rebased Line Variance** (`+0.30%`): `Actual Progress - Rebaseline/Forecast Progress` (from `Weekly Summary` Column U & `Sum` Row 14).
- **Non-Rebased Line** (`-5.46%`): `Actual Progress - Original Baseline Plan Progress` (from `Weekly Summary` Column T & `Sum` Actual vs. Plan).

### Procurement Engineering (MR / TBE)
- 350 deliverables extracted from `PRO.ENGINEERING_MR TBE` worksheet.
- 2-stage milestone gate tracking: Milestone 1 `IFA (A1)` and Milestone 2 `IFA (C1)`.
- Fully integrated into Document Status (Scope filter, status badges) and Overdue & Look-ahead analysis.

### Approval & Completion Tracking Chips
- **PENDING FINAL APPROVAL** (`400`): AFC submit date filled & AP Return date blank for regular docs (`ENG 262`) + IFA A1 / IFA C1 submit filled & corresponding AP Return date blank for MR/TBE (`MR 116 / TBE 22`).
- **COMPLETE** (`398`): Approved Status AP and returned from PTTEPI with scope percentage closure rates: `ENG 301 (20.9%) / MR 53 (30.3%) / TBE 44 (25.1%)`.
- **Sanitization**: Excel default `datetime.time(0,0)` objects sanitized so only authentic text return codes (`APPR`, `APPR-C`, `Issue Rev.X`) or `—` appear.

### Cut-off Date Source
Cut-off date is read from **Excel cell `Sum!D9`** (`2026-08-14`) — NOT from the filename.

---

## 4. Progress Tracker

- [x] Backend (`server.py`): Type 3.1/3.2 delay, lookahead, S-Curve time-series, dynamic column scan
- [x] Frontend (`app.js`): S-Curve from `series[]`, summary table from `summary[]`, 4-digit year dates
- [x] Single-file upload (ASK only) — Z1F is no longer required for this project
- [x] Cut-off date strictly from `Sum!D9` — not filename
- [x] Filename badge: full name displayed, no truncation
- [x] `find_latest_excel()` auto-selects newest Excel by filename date (14-Aug-2026)
- [x] Dynamic S-Curve: `ws.max_column` scanned — auto-expands with each new weekly column
- [x] Backend: Implemented native `.xlsx` generation using `openpyxl` for exporting Delayed, Look-ahead, and All Documents
- [x] Extraction: Made `extract_summary` fully dynamic via label-matching
- [x] Dual Variance Reporting: Implemented **Rebased Line Variance** (`+0.30%`) and **Non-Rebased Line** (`-5.46%`) across Overall Summary, Work Package Progress Summary, Weekly Progress, and S-Curve tabs
- [x] Procurement Engineering: Ingested 350 MR/TBE documents from `PRO.ENGINEERING_MR TBE` into Document Status and Overdue & Look-ahead tabs
- [x] Approval Status: Implemented **PENDING FINAL APPROVAL** tracking (400 items: `ENG 262 / MR 116 / TBE 22`)
- [x] Completion Status: Implemented **COMPLETE** tracking (398 items: `ENG 301 (20.9%) / MR 53 (30.3%) / TBE 44 (25.1%)`)
- [x] Return Code Sanitization: Purged `00:00:00` time artifacts from return code columns

---

## 5. File Reference

| File | Role |
|:---|:---|
| `server.py` | Backend: extract_summary (dual variance), extract_documents, extract_mr_tbe_documents, compute_delay_and_lookahead, exports |
| `app.js` | Frontend: renderSummaryTab (7 KPI cards), renderWPSummaryTable, renderWeeklyTab, renderDocumentsTab (Scope filter, MR/TBE badges, PENDING FINAL APPROVAL, COMPLETE chip), renderOverdueTab |
| `index.html` | HTML shell with 5-tab nav, modals, Scope filter, COMPLETE filter option |
| `styles.css` | CSS: glassmorphism, responsive .kpi-grid, MR/TBE badges, pending final approval and complete badges |
| `HANDOFF_REPORT.md` | Comprehensive technical handoff (v2.4) |
| `spec.md` | System functional specification |
| `lesson_learn.md` | Documented lessons learned (LL-01 to LL-16) |

---

## 6. Commands

```powershell
# Start server
python server.py

# Verify data extraction
python -c "import server; d=server.extract_all_data(); print('Total Docs:', len(d['documents']), '| Delayed:', d['delay_lookahead']['delayed_count'], '| Lookahead:', d['delay_lookahead']['lookahead_count'])"

# Open dashboard
start http://localhost:8081
```

