# System Specification — Engineering Progress Tracking Dashboard
**Project:** EPC OF WELLHEAD PLATFORMS FOR AUNG SINKHA DEVELOPMENT PROJECT PHASE 1A (EPC-01)
**Contractor:** GC Maintenance and Engineering Company Limited (GCME)
**Job No.:** SD-20-26600-01
**Document Version:** 1.0
**Last Updated:** 12 Jul 2026

---

## 1. Purpose & Scope

This specification defines the functional and technical behaviour of the Engineering Progress Tracking & Analytics Dashboard. It serves as the authoritative reference for developers, project controls engineers, and QA reviewers.

---

## 2. Functional Requirements

### 2.1 Data Ingestion
| ID | Requirement |
|:---|:---|
| F-01 | The system SHALL accept weekly MDR Excel files (`*.xlsx`) containing engineering progress data. |
| F-02 | The system SHALL support two file sources per cut-off period: ASK (contains WP01 and WP02). |
| F-03 | The system SHALL automatically detect the latest Excel file using date patterns parsed from the filename (e.g., `Cut off 10-Jul-26`). |
| F-04 | The system SHALL also accept Excel files via browser drag-and-drop or the Upload button (`/api/upload` endpoint). |
| F-05 | After upload, the system SHALL cache data in memory and serve it via `/api/data` without re-reading the file on each request. |

### 2.2 Cut-off Date
| ID | Requirement |
|:---|:---|
| F-06 | The cut-off date displayed in the dashboard header SHALL be read from Excel cell `Sum!D9`. |
| F-07 | The previous cut-off date SHALL be read from Excel cell `Sum!K9`. |
| F-08 | The cut-off date for discipline progress SHALL be read from `Weekly Summary!U7`. |
| F-09 | The cut-off date for the overdue register SHALL be read from `Overdue List!J6`. |
| F-10 | The filename date SHALL NOT override cell-level cut-off dates. |

### 2.3 Document Status Classification
| ID | Requirement |
|:---|:---|
| F-11 | A document with no submission dates (IFR, IFA, AFC all blank) SHALL be classified as `Not Yet Submitted`. |
| F-12 | A document with an IFR submission date only SHALL be classified as `IFR Submitted`. |
| F-13 | A document with an IFA submission date (regardless of IFR) SHALL be classified as `IFA Submitted`. |
| F-14 | A document with an AFC submission date SHALL be classified as `AFC Submitted`. |
| F-15 | The values `""`, `None`, `"N/A"`, and `"-"` SHALL all be treated as blank (no submission). |

### 2.4 Reference Date Hierarchy
| ID | Requirement |
|:---|:---|
| F-16 | The Reference Date for delay/lookahead calculations SHALL be: Forecast Date if present and valid; otherwise Plan Date. |
| F-17 | The Reference Date SHALL apply at each milestone level (IFR, IFA, AFC) independently. |

### 2.5 Delay Classification
| ID | Requirement |
|:---|:---|
| F-18 | **Type 3.1 (Submitted Late):** A document SHALL be classified as Type 3.1 delayed if it WAS submitted AND Submit Date > Reference Date. |
| F-19 | **Type 3.2 (Overdue, Not Submitted):** A document SHALL be classified as Type 3.2 delayed if it was NOT submitted AND Cut-off Date >= Reference Date. |
| F-20 | Delay days for Type 3.1 = max((Submit Date - Reference Date).days, 1). |
| F-21 | Delay days for Type 3.2 = max((Cut-off Date - Reference Date).days, 1). |

### 2.6 14-Day Lookahead
| ID | Requirement |
|:---|:---|
| F-22 | The lookahead window SHALL cover: Cut-off Date < Reference Date <= Cut-off Date + 14 days. |
| F-23 | A document in the lookahead window SHALL be flagged as "Slipping" if its Forecast Date > Plan Date. |
| F-24 | The lookahead SHALL report: days remaining to deadline, urgency (this_week = 7 days / next_week 8–14 days). |

### 2.7 S-Curve
| ID | Requirement |
|:---|:---|
| F-25 | The S-Curve SHALL display three cumulative lines per sheet: Plan, Forecast (dashed), Actual. |
| F-26 | The X-axis SHALL be weekly dates extracted from Row 41 of each S-Curve sheet. |
| F-27 | The Y-axis SHALL represent cumulative progress percentage (0–100%). |
| F-28 | The S-Curve SHALL support 5 views: Overall, WP01 Topside, WP01 Jacket, WP02 Topside, WP02 Jacket. |
| F-29 | The S-Curve data extraction SHALL be fully dynamic — scanning all columns up to `ws.max_column`. No hardcoded column limits. |
| F-30 | A summary table below the chart SHALL show the current-week snapshot (rows 12–16) with Plan/Forecast/Actual incremental and cumulative percentages and deviation. |

### 2.8 Filters
| ID | Requirement |
|:---|:---|
| F-31 | All document-level tabs SHALL support real-time filtering by Work Package (WP), Discipline, and Milestone. |
| F-32 | Filters SHALL apply instantly without a page reload. |
| F-33 | Filtering SHALL affect both chart renders and table rows. |

### 2.9 Dashboard Display
| ID | Requirement |
|:---|:---|
| F-34 | The header SHALL display: full reference filename (no truncation), cut-off date (DD Mon YYYY format, 4-digit year), project name. |
| F-35 | Dates throughout the dashboard SHALL use format: DD Mon YYYY (e.g., 10 Jul 2026). |
| F-36 | Delay aging badges SHALL use four severity levels: critical (>30 days), high (15–30 days), medium (8–14 days), low (1–7 days). |

---

## 3. Non-Functional Requirements

| ID | Requirement |
|:---|:---|
| NF-01 | The server SHALL run on Python 3.11+ with no external cloud dependencies. |
| NF-02 | The server SHALL use in-memory caching. Re-processing the Excel file SHALL only occur when a new file is uploaded or `/api/refresh` is called. |
| NF-03 | The server SHALL be thread-safe for concurrent browser requests. |
| NF-04 | The dashboard SHALL render within 2 seconds of loading on a standard LAN/localhost connection. |
| NF-05 | The system SHALL NOT require a database (SQL/NoSQL). All state is in-memory and sourced from Excel files. |
| NF-06 | The UI SHALL be responsive and viewable on 1920×1080 and 1366×768 resolutions. |

---

## 4. Excel Sheet Layout Reference

### 4.1 `Sum` Sheet
| Cell / Range | Content |
|:---|:---|
| `D9` | Current cut-off date |
| `K9` | Previous cut-off date |
| `D12` | Plan Last Week |
| `E12` | Plan This Week |
| `E13` | Forecast This Week |
| `D14` / `E14` | Actual Last Week / This Week |
| `D15` / `E15` | Variance Last Week / This Week |
| `D16` / `E16` | SPI Weekly Last / This |
| `D17` / `E17` | SPI Cumulative Last / This |
| `G18` | Area of Concern (text) |

### 4.2 `Eng. S-Curve_*` Sheets (all 5 sheets identical structure)
| Row | Content |
|:---|:---|
| 8 | Column headers (Phase, Weight, Plan, Forecast, Actual, Dev) |
| 9–10 | Sub-headers (Incremental, Cumulative, [A], [B], [C]) |
| 12 | WP01 TOPSIDE snapshot (cols E–K) |
| 13 | WP01 JACKET snapshot |
| 14 | WP02 TOPSIDE snapshot |
| 15 | WP02 JACKET snapshot |
| 16 | TOTAL snapshot |
| 41 | Ordinal Week dates (datetime), starting column E (col 5) |
| 42 | Plan Incremental time-series |
| 43 | **Plan Cumulative time-series** ? Chart |
| 44 | Forecast Incremental time-series |
| 45 | **Forecast Cumulative time-series** ? Chart (dashed) |
| 46 | Actual Incremental time-series |
| 47 | **Actual Cumulative time-series** ? Chart |
| 48–49 | Actual meDMS (may contain #REF!) |
| 50–51 | Deviation rows |
| 52–54 | SPI rows |

### 4.3 `Overdue List` Sheet
| Cell / Range | Content |
|:---|:---|
| `J6` | Cut-off date |
| `B9+` | Row number |
| `D9+` | Document number |
| `E9+` | Document title |
| `F9+` | Discipline |
| `G9+` | WP |

### 4.4 `Overdue Summary` Sheet
| Cell / Range | Content |
|:---|:---|
| `B3` | Total overdue text |
| `S3` | Overall progress % |
| `B6–B19` | Discipline names |
| `C6–F19` | Total / 1st Rev / IFA+AFC Issued / Pending Return |

---

## 5. API Specification

### GET /api/data
Returns the full JSON payload from the cached Excel extraction.

```json
{
  "file_name": "MM-ASK-1A-GEN01-PJM-MDR-0001_A1- Cut off 10-Jul-26.xlsx",
  "summary": {
    "cutoff_date": "2026-07-03",
    "prev_cutoff_date": "2026-06-26",
    "plan_this_week": 0.4800,
    "actual_this_week": 0.4132,
    "spi_weekly_this": 0.685,
    "spi_cumulative_this": 0.861,
    "area_of_concern": "..."
  },
  "weekly": {
    "cutoff_date": "2026-07-03",
    "disciplines": [...]
  },
  "scurves": {
    "Overall": {
      "label": "Overall",
      "series": [
        {"date": "2026-03-06", "plan_cum": 0.0, "forecast_cum": 0.0, "actual_cum": 0.0},
        ...
      ],
      "summary": [
        {"phase": "WP01 TOPSIDE", "plan_incr": 0.0743, "plan_cum": 0.5178, ...}
      ]
    }
  },
  "delay_lookahead": {
    "reference_date": "2026-07-03",
    "delayed_count": 174,
    "delayed_type1_count": 30,
    "delayed_type2_count": 144,
    "lookahead_count": 349,
    "lookahead_forecast_delay": 186,
    "delayed": [...],
    "lookahead": [...]
  }
}
```

### POST /api/upload
Accepts multipart/form-data with an `.xlsx` file. Returns:
```json
{"status": "ok", "file": "filename.xlsx", "documents": 1438}
```

### GET /api/refresh
Forces re-scan of the workspace directory and re-loads the latest Excel file. Returns same structure as `/api/data`.

---

## 6. Technology Stack

| Component | Technology | Version |
|:---|:---|:---|
| Backend runtime | Python | 3.11+ |
| Excel parsing | openpyxl | latest |
| HTTP server | http.server (stdlib) | built-in |
| Frontend framework | Vanilla HTML5/CSS3/ES6 | — |
| Charting | Chart.js | CDN v4 |
| Fonts | Google Fonts (Inter) | CDN |

---

## 7. Acceptance Criteria

The system is accepted when all the following are verified:

- [ ] Cut-off date in header matches `Sum!D9` cell value
- [ ] S-Curve shows smooth cumulative curves for all 5 WP tabs (not zigzag)
- [ ] S-Curve X-axis shows weekly dates matching Excel Row 41
- [ ] Type 3.1 delayed count matches manual count from Overdue List sheet
- [ ] Type 3.2 delayed count matches manual count from Overdue List sheet
- [ ] Lookahead window correctly covers 14 days from cut-off date
- [ ] Filename badge shows full filename without truncation
- [ ] Uploading new weekly file updates all dashboard data correctly
- [ ] Filter by WP/Discipline/Milestone works across all tabs
