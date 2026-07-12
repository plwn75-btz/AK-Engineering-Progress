# Aung Sinkha Phase 1A (EPC-01) — Engineering Progress Tracking Dashboard

A high-performance, real-time web dashboard designed to track engineering deliverables across multi-discipline Work Packages (WPs) for the **Aung Sinkha Development Project Phase 1A**. It replaces manual spreadsheet manipulation by automatically parsing weekly Master Document Register (MDR) cut-off Excel sheets (`*.xlsx`) and generating executive S-Curve analytics, delay categorizations, and 14-day lookahead warnings.

---

## 🚀 Key Features

- **Dynamic S-Curve Analytics:** Automatically extracts weekly time-series data from project Excel sheets (`Eng. S-Curve_*`) to plot Plan, Actual, and Forecast cumulative progress (`Chart.js`) across 5 views: Overall, WP01 Topside, WP01 Jacket, WP02 Topside, and WP02 Jacket.
- **Precision Delay Categorization:**
  - **Type 3.1 (Submitted Late):** Tracks documents that were submitted *after* their baseline Forecast Date.
  - **Type 3.2 (Overdue & Not Submitted):** Highlights unissued documents whose Forecast Date has passed compared to the reporting Cut-off Date.
- **14-Day Proactive Lookahead:** Flags upcoming document deadlines (`Cut-off < Forecast ≤ Cut-off + 14 Days`) and highlights slipping deliverables (`Forecast > Plan`).
- **Two-Source File Upload:** Seamlessly handles multi-contractor uploads (ASK & Z1F files) via drag-and-drop or direct upload directly from the web UI.
- **Real-time Multi-dimensional Filtering:** Filter by Work Package (`WP`), Discipline, or Milestone (`IFR`, `IFA`, `AFC`) across charts and registers instantly.
- **Zero Cloud Database Dependency:** Self-contained, lightweight Python backend with thread-safe in-memory caching.

---

## 🛠️ Technology Stack

- **Backend:** Pure Python 3.11+ (`http.server`, `threading`, `openpyxl`)
- **Frontend:** Vanilla HTML5, CSS3 (Glassmorphism dark mode), and ES6 JavaScript
- **Visualization:** Chart.js v4 + Google Inter Typography

---

## 📦 Local Setup & Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/YOUR_USERNAME/engineering-progress-dashboard.git
   cd engineering-progress-dashboard
   ```

2. **Install required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Start the server:**
   ```bash
   python server.py
   ```

4. **Access the Dashboard:**
   Open your browser and navigate to: [http://localhost:8081](http://localhost:8081)

---

## ☁️ Deployment on Render

This repository is pre-configured for instant deployment on [Render.com](https://render.com).

### Option 1: Blueprint Deployment (Recommended)
1. Push this repository to your GitHub account.
2. Log in to [Render Dashboard](https://dashboard.render.com/) and click **New → Blueprint**.
3. Connect your GitHub repository.
4. Render will automatically detect the `render.yaml` blueprint file, run `pip install -r requirements.txt`, and start the application.

### Option 2: Manual Web Service Setup
1. Click **New → Web Service** on Render and select your GitHub repo.
2. Configure the following settings:
   - **Environment:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python server.py`
3. Click **Create Web Service**. Render will automatically assign and inject the `PORT` environment variable (`server.py` dynamically binds to `os.environ.get("PORT", 8081)`).

---

## 📂 Project Structure

```
├── server.py              # Thread-safe Python HTTP server & openpyxl data extraction
├── app.js                 # Client-side SPA state controller & Chart.js renderer
├── index.html             # 5-tab responsive glassmorphism UI shell
├── styles.css             # HSL curated dark-mode stylesheet & aging indicators
├── requirements.txt       # Python dependencies (openpyxl>=3.1.0)
├── render.yaml            # Render PaaS Blueprint configuration
├── HANDOFF_REPORT.md      # Comprehensive technical & operational specifications
├── spec.md                # System functional & non-functional requirements
├── lesson_learn.md        # Documented engineering lessons learned & gotchas
└── handoff.md             # Agent/developer quick reference checklist
```

---

## 📑 Documentation

- **[Technical Handoff Report](HANDOFF_REPORT.md):** Deep-dive into architecture, data model, and verification scripts.
- **[System Specifications](spec.md):** Exact functional requirements and Excel cell mapping reference (`Sum!D9`, `Weekly Summary!U7`, etc.).
- **[Lessons Learned](lesson_learn.md):** Best practices regarding Excel time-series extraction, zero-date handling, and read-only workbook limitations.
