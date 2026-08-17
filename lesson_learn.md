# Lessons Learned — Engineering Progress Tracking Dashboard
**Project:** EPC OF WELLHEAD PLATFORMS FOR AUNG SINKHA DEVELOPMENT PROJECT PHASE 1A (EPC-01)
**Last Updated:** 12 Jul 2026

---

## LL-01: S-Curve Shape Was Wrong — Static Table vs. Time-Series

**Problem:**
The S-Curve chart displayed a zigzag shape instead of a smooth cumulative S-curve. The X-axis showed 5 discipline names (WP01 Topside, WP01 Jacket, etc.) instead of weekly dates.

**Root Cause:**
`extract_scurve()` was reading the static summary table (rows 12–16 = current-week snapshot, 5 rows) and using discipline names as X-axis labels. This is NOT the time-series data.

**Correct Layout:**
The actual weekly time-series lives at a completely different location in the same sheet:
- Row 41 = Ordinal Week dates (datetime objects), starting from column E (col 5)
- Row 43 = Plan Cumulative %
- Row 45 = Forecast Cumulative %
- Row 47 = Actual Cumulative %

**Lesson:**
When Excel sheets contain both a snapshot summary table (rows 12–16) AND a time-series table (rows 41+), always inspect both areas before deciding which to use. The chart needed the time-series; the summary table below the chart correctly uses the snapshot.

---

## LL-02: Cut-off Date Source — Cell vs. Filename

**Problem:**
User uploaded a file named `Cut off 10-Jul-26.xlsx` but the dashboard showed Cut-off: 03 Jul 2026. The filename said 10 Jul but the cell `Sum!D9` still contained 03 Jul (not yet updated in the Excel content).

**Root Cause:**
Contractor copies the previous week's file and renames it without updating cell D9 inside the spreadsheet.

**Decision Made:**
Cut-off date is read strictly from **Excel cells** (`Sum!D9`, `Weekly Summary!U7`, `Overdue List!J6`) — NOT from the filename. This reflects the actual data the contractor has entered.

**Lesson:**
Never override cell-level data with filename-parsed data for reporting purposes. The cell value IS the data; the filename is metadata only.

---

## LL-03: Delay & Lookahead "Today" Reference — Cut-off Date, Not System Clock

**Problem:**
The delay algorithm originally used `datetime.date.today()` (actual calendar date) as the baseline for "today". This caused incorrect delay counts when the dashboard was viewed days after the cut-off date, because documents due between the cut-off and the actual viewing date would suddenly be classified as overdue.

**Fix:**
`compute_delay_and_lookahead()` now uses the cut-off date from `Sum!D9` as "today" (parsed via `datetime.date.fromisoformat(cutoff_date_str)`). This locks the delay evaluation to the reporting period.

**Lesson:**
In project controls dashboards, "today" must refer to the **reporting cut-off date**, not the system clock, to ensure consistent and reproducible results regardless of when the dashboard is viewed.

---

## LL-04: openpyxl read_only=True Does Not Return ws.max_column Reliably

**Problem:**
When using `openpyxl.load_workbook(read_only=True)`, `ws.max_column` may return `None` or an unreliable value because openpyxl does not pre-scan the full sheet in read-only mode.

**Fix:**
The production server uses `openpyxl.load_workbook(data_only=True)` (NOT read_only). This loads all cells into memory but returns reliable `ws.max_column` values, which is essential for the dynamic S-Curve column scan.

**Lesson:**
Use `read_only=True` only for quick one-off inspections. Production extraction must use `data_only=True` to get reliable dimensions and computed cell values.

---

## LL-05: Excel #REF! Errors in Cells

**Problem:**
Some cells in the S-Curve sheets return `#REF!` strings (e.g., `Actual meDMS` rows). If treated as numeric values, these cause crashes or NaN in the chart.

**Fix:**
`safe_float()` returns `None` for any non-numeric value including `#REF!`, `#N/A`, `#VALUE!`. Chart.js handles `null` values gracefully with `spanGaps: true` (for Plan/Forecast) and `spanGaps: false` (for Actual — stops line at last known value).

**Lesson:**
Always wrap Excel cell reads with a safe-cast function that returns `None` on error strings, not `0` or a float.

---

## LL-06: Filename Badge Was Truncating Long Names

**Problem:**
The reference filename badge in the header was showing `MM-ASK-1A-GEN01-PJM-MDR-0001_A1- Cut off 10...` — truncated with `...`. Long MDR filenames were not fully visible.

**Fix:**
Removed `max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap` from `.file-name-badge`. Replaced with `max-width: unset; white-space: normal; word-break: break-word`.

**Lesson:**
For reference/audit file names, always display the full name. Truncation hides critical identifiers (revision number, cut-off date) that users need for audit purposes.

---

## LL-07: grep_search Fails with Quoted SearchPath or Query

**Problem:**
When calling grep_search with extra quotes in the SearchPath parameter (e.g., `"\"c:/path\""` instead of `"c:/path"`), the tool returns no results even when matches exist.

**Lesson:**
Pass SearchPath and Query to grep_search without extra escaping. The tool handles paths natively. Use view_file directly when grep yields empty results.

---

## LL-08: Dynamic S-Curve Columns — Design Confirmed

**Behaviour:**
Each weekly update adds one new date column to the right of row 41 in every S-Curve sheet. The server's column scan automatically picks up the new column on the next upload/refresh. No code changes needed.

**Verification:**
Current file (10-Jul-26) has 75 weekly data points. Next week (17-Jul-26) file will have 76. The dashboard renders all points automatically.

**Lesson:**
Design S-Curve extraction around `ws.max_column` scanning rather than fixed column ranges. This makes the system maintenance-free for weekly updates throughout the project lifecycle.

---

## LL-09: Summary Table Rows Shifted by New Row Insertions

**Problem:**
The dashboard displayed incorrect and negative values for Actual Progress (-0.03%) and absurdly high Variance values (48.77%) when processing the `17-Jul-26` MDR file.

**Root Cause:**
The `extract_summary` function hardcoded row indices (e.g., `D14` for Actual, `E15` for Variance). However, in the `17-Jul-26` file, the contractor inserted a new row ("Rebaseline Variance") at row 14. This shifted all subsequent rows down by one. The server ended up reading "Rebaseline Variance" as Actual Progress, and "Actual Progress" as Variance.

**Fix:**
Implemented dynamic row mapping by scanning column C for specific text labels (e.g., "PLAN PROGRESS", "ACTUAL PROGRESS", "VARIANCE") and dynamically assigning row indices. Extracted `.strip()` from the cell values to account for leading/trailing whitespaces, and added checks to ignore matched labels inside the Legend section at the bottom of the table.

**Lesson:**
Never rely on hardcoded row indices for project control deliverables if the contractor is prone to modifying the template (like inserting new rows). Always use label-based row/column detection.

---

## LL-10: Robust Backend Excel Export Over Frontend CSV Generation

**Problem:**
The frontend JavaScript 'Export to CSV' function generated files that caused errors or formatting issues when opened in Excel.

**Root Cause:**
Manually constructing CSV strings in JavaScript is prone to character escaping issues, encoding problems, and layout mismatches. Excel often fails to cleanly parse CSVs containing special characters or commas within the data fields.

**Fix:**
Implemented native Excel (.xlsx) generation directly on the backend using the `openpyxl` library. Created dedicated endpoints (`/api/export/documents`, `/api/export/lookahead`, `/api/export/delayed`) that construct formatted workbooks in-memory and return them as binary downloads.

**Lesson:**
For reliable tabular data export in web applications, avoid manual CSV generation on the frontend. Use robust backend libraries (like openpyxl or pandas) to generate native Excel files. This guarantees formatting, avoids escaping bugs, and allows the inclusion of styling (like bold headers and auto-sized columns) which improves the user experience.

---

## LL-11: Dual Variance Reporting (Rebased Line vs. Non-Rebased Line)

**Context & Requirement:**
Starting with Rev B1 (Cut-off 14-Aug-26), the MDR template provides two variance metrics:
1. **Column T in Weekly Summary / Deviation [b-a] in S-Curve**: `Actual Progress - Original Plan Progress` = **Non-Rebased Line** (e.g., -5.46% Overall).
2. **Column U in Weekly Summary / Row 14 in Sum / Deviation [b-c] in S-Curve**: `Actual Progress - Rebaseline/Forecast Progress` = **Rebased Line Variance** (e.g., +0.30% Overall).

**Solution:**
- Backend (`server.py`): `extract_summary` extracts both `variance_rebased_this` and `variance_non_rebased_this`. `extract_weekly_summary` preserves both columns T (`var_actual_plan`) and U (`var_actual_forecast`). `extract_scurve` snapshot dynamically extracts all discipline rows and computes `dev_non_rebased` and `dev_rebased`.
- Frontend (`app.js`, `index.html`): Displays 7 KPI cards on Overall Summary, dual variance columns in Work Package Summary Table, Weekly Progress by Discipline Table, and S-Curve Snapshot Table.

**Lesson:**
When engineering projects undergo rebaselining, stakeholders require visibility against *both* the approved rebaseline schedule (operational variance) and the original contractual baseline (macro delay tracking). Standardizing terminology ("Rebased Line Variance" vs. "Non-Rebased Line") across all tabs prevents ambiguity.

---

## LL-12: Ingestion & Tracking of Procurement Engineering (MR / TBE) Deliverables

**Context & Problem:**
Engineering deliverables are not restricted to design reports and drawings (e.g. `CAL`, `DWG`, `SPE`, `MDR`). In EPC projects, Material Requisitions (`MR`) and Technical Bid Evaluations (`TBE`) tracked in dedicated sheets like `PRO.ENGINEERING_MR TBE` represent critical procurement path deliverables with different milestone gates (`IFA (A1)` and `IFA (C1)` rather than `IFR`/`IFA`/`AFC`). Omission of this sheet results in an incomplete register count and missed procurement delay warnings.

**Solution:**
1. Implemented `extract_mr_tbe_documents(wb)` in `server.py` to scan `PRO.ENGINEERING_MR TBE`, dynamically detecting Work Packages (`WP01 Topside`, `WP01 Jacket`, `WP02 Topside`, `WP02 Jacket`).
2. Mapped 2-stage approval milestones: Milestone 1 `IFA (A1)` (Columns V, W, X, Y, AB-AE) and Milestone 2 `IFA (C1)` (Columns AH, AI, AJ, AK, AN-AQ).
3. Integrated MR/TBE deliverables into `compute_delay_and_lookahead()` for Type 3.1, Type 3.2, and 14-day lookahead tracking against the cut-off baseline (`Sum!D9`).
4. Enhanced UI with a dedicated **Scope Filter** (`Engineering Deliverables` vs. `MR / TBE (Procurement)`), type badges (`MR`, `TBE`), and updated milestone filter dropdowns.

**Lesson:**
Always scan for dedicated procurement/vendor engineering sheets in the MDR workbook. Aligning differing milestone naming conventions (`IFA A1`/`C1` vs `IFR`/`IFA`/`AFC`) into a unified data structure ensures holistic delay and lookahead forecasting without breaking tabular views.

---

### LL-13: Scope & Logic Definition for "PENDING FINAL APPROVAL" Tracking

**Problem:**
1. Premature inclusion of deliverables without final submission as "Pending Approval" (e.g. counting items where only IFA or IFR was submitted).
2. Cluttering top-level executive dashboards (Overall Summary) with granular document approval counters, and confusion over tracking "Returned by PTTEPI".

**Solution:**
1. **Strict Logic Definition**:
   - Regular Engineering Deliverables: `is_pending_final_approval = bool(afc_submit_date) and not bool(ap_return_date)`. If `afc_submit_date` is blank, it is NOT pending final approval.
   - MR / TBE Deliverables: Milestone 1 `bool(ifa_submit_date) and not bool(ap_return_date)` OR Milestone 2 `bool(afc_submit_date) and not bool(ap2_return_date)`.
2. **Naming & Emoji Policy**:
   - Standardized term to **"PENDING FINAL APPROVAL"** across all filters, tables, and chips, with all emojis removed.
   - Removed tracking for generic "Returned by PTTEPI" as it is unnecessary for workflow bottlenecks.
3. **Dashboard Scoping**:
   - Kept **Overall Summary (Tab 1)** focused strictly on Progress, Variances, and SPI KPIs.
   - Confined **PENDING FINAL APPROVAL** tracking to **Document Status (Tab 3)** and **Overdue & Look-ahead (Tab 4)** where actionable document-level triage occurs.

**Lesson:**
Ensure status metrics reflect the true contractual stage (e.g., AFC is the final submission milestone before approval). Keep executive summary dashboards clean and focused on high-level progress while providing detailed workflow filters in document exploration tabs.

---

### LL-14: Sanitizing Excel Time Objects (`datetime.time`) in MDR Return Codes

**Problem:**
In openpyxl, empty cells in certain formatted columns (like Column `AP` in regular engineering worksheets) can evaluate to default Python `datetime.time(0, 0)` objects instead of `None`, causing return codes to display as `"00:00:00"` in web interfaces.

**Solution:**
Implemented `safe_code(val)` in `server.py` and secondary string validation in `app.js` to reject any `datetime.time`, `datetime.datetime`, `datetime.date`, `"00:00:00"`, `"0"`, `"-"`, or `"N/A"` values, ensuring only valid text return codes (e.g., `APPR`, `APPR-C`, `Issue Rev.X`) are displayed.

**Lesson:**
Always apply strict type sanitization to Excel cell values in engineering workbooks to prevent raw time/date artifacts from leaking into textual categorical fields.

---

### LL-15: Scope Sub-Breakdown for Approval Bottlenecks (ENG / MR / TBE)

**Problem:**
Users triaging pending approvals needed immediate visibility into the origin of pending deliverables (Standard Engineering vs. Procurement Material Requisitions vs. Technical Bid Evaluations) within the same high-level stat chip.

**Solution:**
Integrated dynamic sub-counter aggregation in `renderDocStatsRow` to breakdown `PENDING FINAL APPROVAL` into `ENG {count} / MR {count} / TBE {count}` dynamically respecting active table filters (e.g. Total unfiltered: `ENG 262 / MR 116 / TBE 22 = 400`).

**Lesson:**
Providing compact scope sub-breakdowns within primary workflow chips gives engineers instant operational context without needing to switch tabs or apply extra filter clicks.

---

### LL-16: "COMPLETE" Approved & Returned Tracking with Scope Percentages

**Problem:**
Tracking progress towards project document closeout requires distinct visibility of approved/returned deliverables (AP Return from PTTEPI) broken down by engineering discipline and procurement scopes (ENG / MR / TBE) alongside percentage of total scope completed.

**Solution:**
Implemented the **`COMPLETE`** stat chip in `renderDocStatsRow` and filter in `index.html` calculating:
- Total Complete: `398` of `1,789` (`22.2%`)
- ENG: `301` of `1,439` (`20.9%`)
- MR: `53` of `175` (`30.3%`)
- TBE: `44` of `175` (`25.1%`)
- Dynamic recalculation on active table filters (e.g. WP or discipline selection).

**Lesson:**
Presenting completion counts alongside scope percentage completion rates (`ENG XX (XX%) / MR XX (XX%) / TBE XX (XX%)`) provides an immediate, high-level gauge of closure velocity across engineering and procurement streams.