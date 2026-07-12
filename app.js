/* ═══════════════════════════════════════════════════════════════
   Engineering Progress Dashboard — Frontend Logic (app.js)
   ═══════════════════════════════════════════════════════════════ */

let DASHBOARD_DATA = null;
let currentTab = 'summary';
let docCurrentPage = 1;
const DOCS_PER_PAGE = 50;

let weeklyChartInstance = null;
let scurveChartInstance = null;
let docDisciplineChartInstance = null;
let overdueDisciplineChartInstance = null;
let activeScurveLabel = 'Overall';

// ─── Initialization & Fetching ──────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
    fetchData();
    setupDragAndDrop();
});

async function fetchData() {
    showLoading(true);
    try {
        const response = await fetch('/api/data');
        if (!response.ok) throw new Error(`HTTP error ${response.status}`);
        const data = await response.json();
        
        if (data.error) {
            showToast('Error: ' + data.error, 'error');
            return;
        }

        DASHBOARD_DATA = data;
        renderAll();
        showToast('Data loaded successfully', 'success');
    } catch (err) {
        console.error('Failed to load data:', err);
        showToast('Failed to load dashboard data: ' + err.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function refreshData() {
    const btn = document.getElementById('btnRefresh');
    if (btn) btn.disabled = true;
    showLoading(true);
    
    try {
        const response = await fetch('/api/refresh');
        if (!response.ok) throw new Error(`HTTP error ${response.status}`);
        const data = await response.json();
        
        if (data.error) {
            showToast('Error: ' + data.error, 'error');
            return;
        }

        DASHBOARD_DATA = data;
        renderAll();
        showToast('Data refreshed from Excel file', 'success');
    } catch (err) {
        console.error('Refresh failed:', err);
        showToast('Failed to refresh data: ' + err.message, 'error');
    } finally {
        if (btn) btn.disabled = false;
        showLoading(false);
    }
}

function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (!overlay) return;
    if (show) {
        overlay.style.display = 'flex';
        document.querySelectorAll('.tab-panel').forEach(p => p.style.display = 'none');
    } else {
        overlay.style.display = 'none';
        const activePanel = document.getElementById(`tab-${currentTab}`);
        if (activePanel) activePanel.style.display = 'block';
    }
}

// ─── Rendering Master ───────────────────────────────────────────────────────

function renderAll() {
    if (!DASHBOARD_DATA) return;

    // Header info
    document.getElementById('projectName').textContent = DASHBOARD_DATA.summary?.project_name || 'AUNG SINKHA DEVELOPMENT PROJECT PHASE 1A (EPC-01)';
    const fileNameEl = document.getElementById('fileName');
    if (fileNameEl) {
        fileNameEl.textContent = DASHBOARD_DATA.file_name || 'Unknown Excel File';
        fileNameEl.title = DASHBOARD_DATA.file_name || 'Unknown Excel File';
    }
    document.getElementById('cutoffDate').textContent = formatDate(DASHBOARD_DATA.summary?.cutoff_date) || '—';

    // Populate filter dropdowns
    populateFilters();

    // Render tabs
    renderSummaryTab();
    renderWeeklyTab();
    renderDocumentsTab();
    renderOverdueTab();
    renderScurveTab();

    // Update notification badges
    updateBadges();
}

function updateBadges() {
    const overdueCount = DASHBOARD_DATA.delay_lookahead?.delayed_count || 0;
    const lookaheadCount = DASHBOARD_DATA.delay_lookahead?.lookahead_count || 0;
    const totalWarnings = overdueCount + lookaheadCount;
    
    const badge = document.getElementById('overdueBadge');
    if (badge) {
        if (totalWarnings > 0) {
            badge.textContent = totalWarnings;
            badge.style.display = 'inline-block';
        } else {
            badge.style.display = 'none';
        }
    }
}

// ─── Tab Switching ──────────────────────────────────────────────────────────

function switchTab(tabId, btnElement) {
    currentTab = tabId;
    
    // Update navigation button active state
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    if (btnElement) btnElement.classList.add('active');

    // Show panel
    document.querySelectorAll('.tab-panel').forEach(panel => {
        panel.style.display = 'none';
        panel.classList.remove('active');
    });
    
    const targetPanel = document.getElementById(`tab-${tabId}`);
    if (targetPanel) {
        targetPanel.style.display = 'block';
        targetPanel.classList.add('active');
    }

    // Trigger chart resize or lazy rendering if needed
    if (tabId === 'weekly' && weeklyChartInstance) {
        weeklyChartInstance.resize();
    }
    if (tabId === 'scurve' && scurveChartInstance) {
        scurveChartInstance.resize();
    }
}

// ─── Tab 1: Summary Tab ─────────────────────────────────────────────────────

function renderSummaryTab() {
    const s = DASHBOARD_DATA.summary || {};
    const grid = document.getElementById('kpiGrid');
    if (!grid) return;

    const planPct = formatPct(s.plan_this_week);
    const actualPct = formatPct(s.actual_this_week);
    const forecastPct = formatPct(s.forecast_this_week);
    const varPct = formatPct(s.variance_this_week);

    // Determine variance color
    let varClass = 'green';
    let varTextClass = 'positive';
    if (s.variance_this_week < -0.01) {
        varClass = 'red';
        varTextClass = 'negative';
    } else if (s.variance_this_week < 0) {
        varClass = 'amber';
        varTextClass = 'negative';
    }

    grid.innerHTML = `
        <div class="kpi-card blue">
            <div class="kpi-label">Plan Progress (This Week)</div>
            <div class="kpi-value blue">${planPct}</div>
            <div class="kpi-sub">Last Week: ${formatPct(s.plan_last_week)}</div>
        </div>
        <div class="kpi-card ${varClass}">
            <div class="kpi-label">Actual Progress (This Week)</div>
            <div class="kpi-value ${varClass}">${actualPct}</div>
            <div class="kpi-sub">Last Week: ${formatPct(s.actual_last_week)}</div>
        </div>
        <div class="kpi-card purple">
            <div class="kpi-label">Forecast Progress</div>
            <div class="kpi-value purple">${forecastPct !== '—' ? forecastPct : 'N/A'}</div>
            <div class="kpi-sub">Cut-off: ${formatDate(s.cutoff_date)}</div>
        </div>
        <div class="kpi-card ${varClass}">
            <div class="kpi-label">Progress Variance</div>
            <div class="kpi-value ${varClass}">${varPct}</div>
            <div class="kpi-sub">
                <span class="kpi-change ${varTextClass}">${s.variance_this_week >= 0 ? '▲' : '▼'} ${varPct}</span> vs Plan
            </div>
        </div>
        <div class="kpi-card green">
            <div class="kpi-label">SPI (Weekly)</div>
            <div class="kpi-value green">${s.spi_weekly_this?.toFixed(2) || '—'}</div>
            <div class="kpi-sub">Last Week: ${s.spi_weekly_last?.toFixed(2) || '—'}</div>
        </div>
        <div class="kpi-card green">
            <div class="kpi-label">SPI (Cumulative)</div>
            <div class="kpi-value green">${s.spi_cumulative_this?.toFixed(2) || '—'}</div>
            <div class="kpi-sub">Last Week: ${s.spi_cumulative_last?.toFixed(2) || '—'}</div>
        </div>
    `;

    // Render Work Package cards
    renderWPCards();

    // Render Work Package Summary Table
    renderWPSummaryTable();
}

function renderWPCards() {
    const wpGrid = document.getElementById('wpGrid');
    if (!wpGrid) return;

    const stats = DASHBOARD_DATA.doc_stats || {};
    let html = '';

    for (const [wpName, st] of Object.entries(stats)) {
        const submittedPct = st.total > 0 ? ((st.submitted / st.total) * 100).toFixed(1) + '%' : '0%';
        html += `
            <div class="wp-card">
                <div class="wp-card-title">📦 ${wpName}</div>
                <div class="wp-card-stats">
                    <div class="wp-stat">
                        <div class="wp-stat-value">${st.total}</div>
                        <div class="wp-stat-label">Total Docs</div>
                    </div>
                    <div class="wp-stat">
                        <div class="wp-stat-value var-positive">${st.submitted}</div>
                        <div class="wp-stat-label">Submitted</div>
                    </div>
                    <div class="wp-stat">
                        <div class="wp-stat-value var-warning">${st.not_submitted}</div>
                        <div class="wp-stat-label">Not Yet Submitted</div>
                    </div>
                </div>
                <div class="progress-bar">
                    <div class="progress-bar-fill green" style="width: ${submittedPct}"></div>
                </div>
            </div>
        `;
    }
    wpGrid.innerHTML = html;
}

function renderWPSummaryTable() {
    const thead = document.getElementById('wpSummaryHead');
    const tbody = document.getElementById('wpSummaryBody');
    if (!thead || !tbody) return;

    thead.innerHTML = `
        <th>Work Package / Scope</th>
        <th>Weight (L1 %)</th>
        <th>Weight (L2 %)</th>
        <th>Plan % (Last)</th>
        <th>Actual % (Last)</th>
        <th>Plan % (This Week)</th>
        <th>Forecast % (This Week)</th>
        <th>Actual % (This Week)</th>
        <th>Variance vs Plan</th>
    `;

    const disciplines = DASHBOARD_DATA.weekly_summary?.disciplines || [];
    let html = '';

    disciplines.forEach(item => {
        if (item.is_wp_header) {
            const isOverall = item.is_overall;
            const rowClass = isOverall ? 'overall-row' : 'wp-header';
            const varVal = item.var_actual_plan;
            const varFormatted = formatPct(varVal);
            let varClass = 'var-neutral';
            if (varVal > 0) varClass = 'var-positive';
            else if (varVal < -0.01) varClass = 'var-negative';
            else if (varVal < 0) varClass = 'var-warning';

            html += `
                <tr class="${rowClass}">
                    <td>${item.name}</td>
                    <td>${item.weight_l1 !== null ? (item.weight_l1 * 100).toFixed(2) + '%' : '—'}</td>
                    <td>${item.weight_l2 !== null && !isOverall ? (item.weight_l2 * 100).toFixed(2) + '%' : '—'}</td>
                    <td>${formatPct(item.last_plan)}</td>
                    <td>${formatPct(item.last_actual)}</td>
                    <td>${formatPct(item.this_plan)}</td>
                    <td>${formatPct(item.this_forecast)}</td>
                    <td>${formatPct(item.this_actual)}</td>
                    <td class="${varClass}">${varFormatted}</td>
                </tr>
            `;
        }
    });

    tbody.innerHTML = html;
}

// ─── Tab 2: Weekly Progress Tab ─────────────────────────────────────────────

function renderWeeklyTab() {
    const thead = document.getElementById('weeklyHead');
    const tbody = document.getElementById('weeklyBody');
    if (!thead || !tbody) return;

    thead.innerHTML = `
        <th>Discipline / Area</th>
        <th>Weight (L2 %)</th>
        <th>Up To Last Week (Plan)</th>
        <th>Up To Last Week (Actual)</th>
        <th>Incremental Plan</th>
        <th>Incremental Actual</th>
        <th>Up To This Week (Plan)</th>
        <th>Up To This Week (Forecast)</th>
        <th>Up To This Week (Actual)</th>
        <th>Variance (Act - Plan)</th>
    `;

    const disciplines = DASHBOARD_DATA.weekly_summary?.disciplines || [];
    let html = '';

    disciplines.forEach(item => {
        if (item.is_wp_header) {
            const isOverall = item.is_overall;
            const rowClass = isOverall ? 'overall-row' : 'wp-header';
            html += `
                <tr class="${rowClass}">
                    <td colspan="2">📦 ${item.name}</td>
                    <td>${formatPct(item.last_plan)}</td>
                    <td>${formatPct(item.last_actual)}</td>
                    <td>${formatPct(item.incr_plan)}</td>
                    <td>${formatPct(item.incr_actual)}</td>
                    <td>${formatPct(item.this_plan)}</td>
                    <td>${formatPct(item.this_forecast)}</td>
                    <td>${formatPct(item.this_actual)}</td>
                    <td class="${item.var_actual_plan < 0 ? 'var-negative' : 'var-positive'}">${formatPct(item.var_actual_plan)}</td>
                </tr>
            `;
        } else {
            const varVal = item.var_actual_plan;
            let varClass = 'var-neutral';
            if (varVal > 0) varClass = 'var-positive';
            else if (varVal < -0.01) varClass = 'var-negative';
            else if (varVal < 0) varClass = 'var-warning';

            html += `
                <tr>
                    <td style="padding-left: 1.5rem;">🔹 ${item.name}</td>
                    <td>${item.weight_l2 !== null ? (item.weight_l2 * 100).toFixed(2) + '%' : '—'}</td>
                    <td>${formatPct(item.last_plan)}</td>
                    <td>${formatPct(item.last_actual)}</td>
                    <td>${formatPct(item.incr_plan)}</td>
                    <td>${formatPct(item.incr_actual)}</td>
                    <td>${formatPct(item.this_plan)}</td>
                    <td>${formatPct(item.this_forecast)}</td>
                    <td>${formatPct(item.this_actual)}</td>
                    <td class="${varClass}">${formatPct(item.var_actual_plan)}</td>
                </tr>
            `;
        }
    });

    tbody.innerHTML = html;
    renderWeeklyChart(disciplines);
}

function renderWeeklyChart(disciplines) {
    const ctx = document.getElementById('weeklyChart')?.getContext('2d');
    if (!ctx) return;

    // Filter WP headers only (excluding Overall)
    const wpHeaders = disciplines.filter(d => d.is_wp_header && !d.is_overall);
    const labels = wpHeaders.map(d => d.name);
    const planData = wpHeaders.map(d => d.this_plan !== null ? (d.this_plan * 100) : 0);
    const actualData = wpHeaders.map(d => d.this_actual !== null ? (d.this_actual * 100) : 0);
    const forecastData = wpHeaders.map(d => d.this_forecast !== null ? (d.this_forecast * 100) : 0);

    if (weeklyChartInstance) weeklyChartInstance.destroy();

    weeklyChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Plan %',
                    data: planData,
                    backgroundColor: 'rgba(99, 102, 241, 0.7)',
                    borderColor: 'rgba(99, 102, 241, 1)',
                    borderWidth: 1,
                    borderRadius: 4
                },
                {
                    label: 'Actual %',
                    data: actualData,
                    backgroundColor: 'rgba(16, 185, 129, 0.7)',
                    borderColor: 'rgba(16, 185, 129, 1)',
                    borderWidth: 1,
                    borderRadius: 4
                },
                {
                    label: 'Forecast %',
                    data: forecastData,
                    backgroundColor: 'rgba(245, 158, 11, 0.5)',
                    borderColor: 'rgba(245, 158, 11, 1)',
                    borderWidth: 1,
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#f1f5f9', font: { family: 'Inter' } } },
                tooltip: {
                    callbacks: {
                        label: (ctx) => `${ctx.dataset.label}: ${ctx.raw?.toFixed(2)}%`
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#94a3b8', callback: val => `${val}%` }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#f1f5f9' }
                }
            }
        }
    });
}

// ─── Tab 3: Documents Tab ───────────────────────────────────────────────────

function populateFilters() {
    // Populate WP filter
    const wpSelect = document.getElementById('filterWP');
    const delayWpSelect = document.getElementById('delayFilterWP');
    const overdueWpSelect = document.getElementById('overdueFilterWP');
    if (wpSelect && DASHBOARD_DATA.wp_list) {
        wpSelect.innerHTML = '<option value="">All Work Packages</option>' + 
            DASHBOARD_DATA.wp_list.map(wp => `<option value="${wp}">${wp}</option>`).join('');
    }
    if (delayWpSelect && DASHBOARD_DATA.wp_list) {
        delayWpSelect.innerHTML = '<option value="">All Work Packages</option>' + 
            DASHBOARD_DATA.wp_list.map(wp => `<option value="${wp}">${wp}</option>`).join('');
    }
    if (overdueWpSelect && DASHBOARD_DATA.wp_list) {
        overdueWpSelect.innerHTML = '<option value="">All Work Packages</option>' + 
            DASHBOARD_DATA.wp_list.map(wp => `<option value="${wp}">${wp}</option>`).join('');
    }

    // Populate Discipline filter
    const discSelect = document.getElementById('filterDisc');
    const overdueDiscSelect = document.getElementById('overdueFilterDisc');
    if (discSelect && DASHBOARD_DATA.disciplines) {
        discSelect.innerHTML = '<option value="">All Disciplines</option>' + 
            DASHBOARD_DATA.disciplines.map(d => `<option value="${d}">${d}</option>`).join('');
    }
    if (overdueDiscSelect && DASHBOARD_DATA.disciplines) {
        overdueDiscSelect.innerHTML = '<option value="">All Disciplines</option>' + 
            DASHBOARD_DATA.disciplines.map(d => `<option value="${d}">${d}</option>`).join('');
    }
}

function filterDocuments() {
    docCurrentPage = 1;
    renderDocumentsTab();
}

function renderDocumentsTab() {
    const tbody = document.getElementById('docBody');
    if (!tbody) return;

    const wpFilter = document.getElementById('filterWP')?.value || '';
    const discFilter = document.getElementById('filterDisc')?.value || '';
    const statusFilter = document.getElementById('filterStatus')?.value || '';
    const searchFilter = document.getElementById('filterSearch')?.value?.toLowerCase().trim() || '';

    const allDocs = DASHBOARD_DATA.documents || [];

    // Filter
    const filtered = allDocs.filter(doc => {
        if (wpFilter && doc.wp !== wpFilter) return false;
        if (discFilter && doc.discipline !== discFilter) return false;
        if (statusFilter && doc.status !== statusFilter) return false;
        if (searchFilter) {
            const matchNo = doc.doc_no?.toLowerCase().includes(searchFilter);
            const matchTitle = doc.title?.toLowerCase().includes(searchFilter);
            if (!matchNo && !matchTitle) return false;
        }
        return true;
    });

    // Update stats row
    renderDocStatsRow(allDocs, filtered);

    // Render stacked bar chart by discipline
    renderDocDisciplineChart(allDocs, wpFilter);

    // Update document count info
    const countInfo = document.getElementById('docCountInfo');
    if (countInfo) {
        countInfo.textContent = `Showing ${filtered.length} of ${allDocs.length} documents`;
    }

    // Paginate
    const totalPages = Math.ceil(filtered.length / DOCS_PER_PAGE) || 1;
    if (docCurrentPage > totalPages) docCurrentPage = totalPages;

    const startIdx = (docCurrentPage - 1) * DOCS_PER_PAGE;
    const pageDocs = filtered.slice(startIdx, startIdx + DOCS_PER_PAGE);

    // Render table
    let html = '';
    pageDocs.forEach((doc, idx) => {
        const rowNo = startIdx + idx + 1;
        const statusClass = getStatusBadgeClass(doc.status);
        const varPct = doc.variance !== null ? (doc.variance * 100).toFixed(1) + '%' : '—';
        const varClass = doc.variance < 0 ? 'var-negative' : 'var-positive';

        html += `
            <tr>
                <td>${rowNo}</td>
                <td><span class="status-badge" style="background: rgba(255,255,255,0.05); color: var(--text-primary); border: none;">${doc.wp}</span></td>
                <td><b>${doc.discipline || '—'}</b></td>
                <td style="font-family: monospace; color: var(--accent-cyan);">${doc.doc_no || '—'}</td>
                <td style="max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${doc.title || ''}">${doc.title || '—'}</td>
                <td>${formatPct(doc.plan_pct)}</td>
                <td>${formatPct(doc.actual_pct)}</td>
                <td class="${varClass}">${varPct}</td>
                <td>${formatDate(doc.ifr_forecast || doc.ifr_plan)}</td>
                <td>${formatDate(doc.ifr_submit_date)}</td>
                <td>${formatDate(doc.ifa_forecast || doc.ifa_plan)}</td>
                <td>${formatDate(doc.ifa_submit_date)}</td>
                <td>${formatDate(doc.afc_forecast || doc.afc_plan)}</td>
                <td>${formatDate(doc.afc_submit_date)}</td>
                <td><span class="status-badge ${statusClass}">${doc.status || 'Unknown'}</span></td>
            </tr>
        `;
    });

    if (pageDocs.length === 0) {
        html = `<tr><td colspan="15" style="text-align: center; padding: 2rem; color: var(--text-muted);">No documents match your filter criteria.</td></tr>`;
    }

    tbody.innerHTML = html;
    renderDocPagination(totalPages);
}

function getStatusBadgeClass(status) {
    switch (status) {
        case 'Not Yet Submitted': return 'not-submitted';
        case 'IFR Submitted': return 'ifr-submitted';
        case 'IFA Submitted': return 'ifa-submitted';
        case 'AFC Submitted': return 'afc-submitted';
        default: return '';
    }
}

function renderDocStatsRow(allDocs, filteredDocs) {
    const statsContainer = document.getElementById('docStatsRow');
    if (!statsContainer) return;

    const notSub = filteredDocs.filter(d => d.status === 'Not Yet Submitted').length;
    const ifrSub = filteredDocs.filter(d => d.status === 'IFR Submitted').length;
    const ifaSub = filteredDocs.filter(d => d.status === 'IFA Submitted').length;
    const afcSub = filteredDocs.filter(d => d.status === 'AFC Submitted').length;

    statsContainer.innerHTML = `
        <div class="stat-chip info">
            <div class="stat-number">${filteredDocs.length}</div>
            <div class="stat-label">Filtered Documents</div>
        </div>
        <div class="stat-chip warning">
            <div class="stat-number">${notSub}</div>
            <div class="stat-label">Not Yet Submitted</div>
        </div>
        <div class="stat-chip" style="border-color: var(--accent-amber);">
            <div class="stat-number" style="color: var(--accent-amber);">${ifrSub}</div>
            <div class="stat-label">IFR Submitted</div>
        </div>
        <div class="stat-chip" style="border-color: var(--accent-blue);">
            <div class="stat-number" style="color: var(--accent-blue);">${ifaSub}</div>
            <div class="stat-label">IFA Submitted</div>
        </div>
        <div class="stat-chip success">
            <div class="stat-number">${afcSub}</div>
            <div class="stat-label">AFC Submitted</div>
        </div>
    `;
}

function isDocDelayedOrSlipped(doc, todayStr) {
    if (!todayStr) todayStr = DASHBOARD_DATA?.delay_lookahead?.reference_date || new Date().toISOString().split('T')[0];
    const milestones = [
        { forecast: doc.ifr_forecast, plan: doc.ifr_plan, submit: doc.ifr_submit_date },
        { forecast: doc.ifa_forecast, plan: doc.ifa_plan, submit: doc.ifa_submit_date },
        { forecast: doc.afc_forecast, plan: doc.afc_plan, submit: doc.afc_submit_date }
    ];

    for (let ms of milestones) {
        const refDate = (ms.forecast && ms.forecast !== 'N/A' && ms.forecast !== '-') ? ms.forecast : ms.plan;
        if (!refDate || refDate === 'N/A' || refDate === '-') continue;

        // Type 3.1: submit date is after Forecast Date
        if (ms.submit && ms.submit !== 'N/A' && ms.submit !== '-') {
            if (ms.submit > refDate) {
                return true;
            }
        } else {
            // Type 3.2: not submitted and today is overdue the Forecast date
            if (refDate <= todayStr) {
                return true;
            }
        }
    }
    return false;
}

function renderDocDisciplineChart(allDocs, wpFilter = '') {
    const ctx = document.getElementById('docDisciplineChart')?.getContext('2d');
    if (!ctx) return;

    // Filter documents by selected Work Package (if any)
    const docs = wpFilter ? allDocs.filter(d => d.wp === wpFilter) : allDocs;

    // Get unique disciplines
    const disciplines = [...new Set(docs.map(d => d.discipline).filter(Boolean))].sort();
    if (disciplines.length === 0) return;

    // Requirement 1: All tracking date at all dashboard shall refer from Today and Forecast date
    const todayStr = DASHBOARD_DATA?.delay_lookahead?.reference_date || new Date().toISOString().split('T')[0];

    const submittedData = [];
    const notSubData = [];
    const delayedData = [];

    disciplines.forEach(disc => {
        const discDocs = docs.filter(d => d.discipline === disc);
        let subCount = 0;
        let notSubCount = 0;
        let delayCount = 0;

        discDocs.forEach(doc => {
            if (isDocDelayedOrSlipped(doc, todayStr)) {
                delayCount++;
            } else if (doc.status === 'Not Yet Submitted') {
                notSubCount++;
            } else {
                subCount++;
            }
        });

        submittedData.push(subCount);
        notSubData.push(notSubCount);
        delayedData.push(delayCount);
    });

    if (docDisciplineChartInstance) docDisciplineChartInstance.destroy();

    docDisciplineChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: disciplines,
            datasets: [
                {
                    label: 'Delayed Document',
                    data: delayedData,
                    backgroundColor: 'rgba(239, 68, 68, 0.85)', // Crimson Red
                    borderColor: 'rgba(239, 68, 68, 1)',
                    borderWidth: 1,
                    borderRadius: 2
                },
                {
                    label: 'Not Yet Submitted (On Schedule / Future)',
                    data: notSubData,
                    backgroundColor: 'rgba(245, 158, 11, 0.8)', // Amber Orange
                    borderColor: 'rgba(245, 158, 11, 1)',
                    borderWidth: 1,
                    borderRadius: 2
                },
                {
                    label: 'Submitted (IFR/IFA/AFC)',
                    data: submittedData,
                    backgroundColor: 'rgba(16, 185, 129, 0.85)', // Emerald Green
                    borderColor: 'rgba(16, 185, 129, 1)',
                    borderWidth: 1,
                    borderRadius: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: { color: '#f1f5f9', font: { family: 'Inter', size: 13, weight: '500' }, padding: 16 }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.raw || 0;
                            const total = (submittedData[ctx.dataIndex] || 0) + (notSubData[ctx.dataIndex] || 0) + (delayedData[ctx.dataIndex] || 0);
                            const pct = total > 0 ? ((val / total) * 100).toFixed(1) + '%' : '0%';
                            return `${ctx.dataset.label}: ${val} docs (${pct})`;
                        },
                        footer: (tooltipItems) => {
                            if (!tooltipItems.length) return '';
                            const idx = tooltipItems[0].dataIndex;
                            const total = (submittedData[idx] || 0) + (notSubData[idx] || 0) + (delayedData[idx] || 0);
                            return `Total Discipline Docs: ${total}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { color: '#e2e8f0', font: { family: 'Inter', weight: '600' } }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.06)' },
                    ticks: { color: '#94a3b8', stepSize: 20 }
                }
            }
        }
    });
}

function renderDocPagination(totalPages) {
    const container = document.getElementById('docPagination');
    if (!container) return;

    if (totalPages <= 1) {
        container.innerHTML = '';
        return;
    }

    let html = `
        <button class="page-btn" onclick="changeDocPage(${docCurrentPage - 1})" ${docCurrentPage === 1 ? 'disabled' : ''}>◀ Prev</button>
        <span class="page-info">Page ${docCurrentPage} of ${totalPages}</span>
        <button class="page-btn" onclick="changeDocPage(${docCurrentPage + 1})" ${docCurrentPage === totalPages ? 'disabled' : ''}>Next ▶</button>
    `;

    container.innerHTML = html;
}

function changeDocPage(newPage) {
    docCurrentPage = newPage;
    renderDocumentsTab();
}

function exportDocumentsCSV() {
    const allDocs = DASHBOARD_DATA.documents || [];
    if (allDocs.length === 0) return;

    const headers = [
        'WP', 'Discipline', 'Document No.', 'Title', 'Class', 'Plan %', 'Actual %', 'Variance',
        'IFR Forecast', 'IFR Submit Date', 'IFA Forecast', 'IFA Submit Date', 'AFC Forecast', 'AFC Submit Date', 'Status'
    ];

    const rows = allDocs.map(d => [
        `"${d.wp || ''}"`,
        `"${d.discipline || ''}"`,
        `"${d.doc_no || ''}"`,
        `"${(d.title || '').replace(/"/g, '""')}"`,
        `"${d.class || ''}"`,
        `"${formatPct(d.plan_pct)}"`,
        `"${formatPct(d.actual_pct)}"`,
        `"${d.variance !== null ? (d.variance * 100).toFixed(1) + '%' : ''}"`,
        `"${formatDate(d.ifr_forecast || d.ifr_plan)}"`,
        `"${formatDate(d.ifr_submit_date)}"`,
        `"${formatDate(d.ifa_forecast || d.ifa_plan)}"`,
        `"${formatDate(d.ifa_submit_date)}"`,
        `"${formatDate(d.afc_forecast || d.afc_plan)}"`,
        `"${formatDate(d.afc_submit_date)}"`,
        `"${d.status || ''}"`
    ]);

    const csvContent = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.setAttribute('download', `Engineering_Documents_Export_${DASHBOARD_DATA.summary?.cutoff_date || 'latest'}.csv`);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// ─── Tab 4: Overdue & Look-ahead Tab ────────────────────────────────────────

function filterOverdueTab() {
    renderOverdueTab();
}

function filterDelayed() {
    renderOverdueTab();
}

function renderOverdueTab() {
    const dl = DASHBOARD_DATA.delay_lookahead || {};
    
    const wpFilter = document.getElementById('overdueFilterWP')?.value || '';
    const discFilter = document.getElementById('overdueFilterDisc')?.value || '';
    const msFilter = document.getElementById('overdueFilterMS')?.value || '';

    // Filter delayed and lookahead lists for stats and tables
    const allDelayed = dl.delayed || [];
    const filteredDelayed = allDelayed.filter(item => {
        if (wpFilter && item.wp !== wpFilter) return false;
        if (discFilter && item.discipline !== discFilter) return false;
        if (msFilter && item.milestone !== msFilter) return false;
        return true;
    });

    const allLookahead = dl.lookahead || [];
    const filteredLookahead = allLookahead.filter(item => {
        if (wpFilter && item.wp !== wpFilter) return false;
        if (discFilter && item.discipline !== discFilter) return false;
        if (msFilter && item.milestone !== msFilter) return false;
        return true;
    });

    // Stats row
    const statsRow = document.getElementById('overdueStatsRow');
    if (statsRow) {
        const t1Count = filteredDelayed.filter(d => d.delay_type_code === '3.1').length;
        const t2Count = filteredDelayed.filter(d => d.delay_type_code === '3.2').length;
        const lookaheadSlipping = filteredLookahead.filter(d => d.is_slipping).length;

        statsRow.innerHTML = `
            <div class="stat-chip danger" style="flex: 1.5;">
                <div class="stat-number">${filteredDelayed.length}</div>
                <div class="stat-label">Total Delayed (${t1Count} Submitted Late / ${t2Count} Not Issued)</div>
            </div>
            <div class="stat-chip warning" style="flex: 1.5;">
                <div class="stat-number">${filteredLookahead.length}</div>
                <div class="stat-label">Due Next 14 Days (${lookaheadSlipping} Forecasted to Delay)</div>
            </div>
            <div class="stat-chip info">
                <div class="stat-number">${formatDate(dl.reference_date || new Date().toISOString().split('T')[0])}</div>
                <div class="stat-label">Today Reference Date</div>
            </div>
        `;
    }

    // Render stacked bar chart on Overdue tab across filtered documents
    renderOverdueDisciplineChart(wpFilter, discFilter, msFilter);

    // Render Lookahead Table
    const lookBody = document.getElementById('lookaheadBody');
    if (lookBody) {
        let html = '';
        filteredLookahead.forEach(item => {
            const urgencyClass = item.urgency === 'this_week' ? 'urgency-this-week' : 'urgency-next-week';
            const badgeText = item.urgency === 'this_week' ? '🔴 Due This Week' : '🟡 Due Next Week';
            html += `
                <tr class="${urgencyClass}">
                    <td><span class="aging-badge ${item.urgency === 'this_week' ? 'critical' : 'medium'}">${badgeText}</span></td>
                    <td><b>${item.wp}</b></td>
                    <td>${item.discipline}</td>
                    <td style="font-family: monospace; color: var(--accent-cyan);">${item.doc_no}</td>
                    <td style="max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.title}">${item.title}</td>
                    <td><span class="status-badge ifr-submitted">${item.milestone}</span></td>
                    <td>${formatDate(item.plan_date)}</td>
                    <td>${formatDate(item.forecast_date)}</td>
                    <td><b>${item.days_remaining} days</b></td>
                    <td>${item.is_slipping ? '<span style="color: var(--accent-red); font-weight: bold;">⚠️ Slipping</span>' : '—'}</td>
                </tr>
            `;
        });
        if (filteredLookahead.length === 0) {
            html = `<tr><td colspan="10" style="text-align: center; padding: 1.5rem; color: var(--text-muted);">✅ No documents scheduled for submission within the next 14 days matching filters.</td></tr>`;
        }
        lookBody.innerHTML = html;
    }

    // Render Delayed Section
    renderDelayedSection(filteredDelayed, allDelayed.length);

    // Render Overdue Summary Table (from Overdue Summary sheet)
    renderOverdueSummaryTable();
}

function renderOverdueDisciplineChart(wpFilter = '', discFilter = '', msFilter = '') {
    const ctx = document.getElementById('overdueDisciplineChart')?.getContext('2d');
    if (!ctx) return;

    let allDocs = DASHBOARD_DATA?.documents || [];
    if (wpFilter) allDocs = allDocs.filter(d => d.wp === wpFilter);
    if (discFilter) allDocs = allDocs.filter(d => d.discipline === discFilter);

    const disciplines = [...new Set(allDocs.map(d => d.discipline).filter(Boolean))].sort();
    if (disciplines.length === 0) return;

    const cutoffDateStr = DASHBOARD_DATA?.delay_lookahead?.reference_date || new Date().toISOString().split('T')[0];

    const submittedData = [];
    const notSubData = [];
    const delayedData = [];

    disciplines.forEach(disc => {
        const discDocs = allDocs.filter(d => d.discipline === disc);
        let subCount = 0;
        let notSubCount = 0;
        let delayCount = 0;

        discDocs.forEach(doc => {
            if (isDocDelayedOrSlipped(doc, cutoffDateStr)) {
                delayCount++;
            } else if (doc.status === 'Not Yet Submitted') {
                notSubCount++;
            } else {
                subCount++;
            }
        });

        submittedData.push(subCount);
        notSubData.push(notSubCount);
        delayedData.push(delayCount);
    });

    if (overdueDisciplineChartInstance) overdueDisciplineChartInstance.destroy();

    overdueDisciplineChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: disciplines,
            datasets: [
                {
                    label: 'Delayed Document',
                    data: delayedData,
                    backgroundColor: 'rgba(239, 68, 68, 0.85)',
                    borderColor: 'rgba(239, 68, 68, 1)',
                    borderWidth: 1,
                    borderRadius: 2
                },
                {
                    label: 'Not Yet Submitted (On Schedule / Future)',
                    data: notSubData,
                    backgroundColor: 'rgba(245, 158, 11, 0.8)',
                    borderColor: 'rgba(245, 158, 11, 1)',
                    borderWidth: 1,
                    borderRadius: 2
                },
                {
                    label: 'Submitted (IFR/IFA/AFC)',
                    data: submittedData,
                    backgroundColor: 'rgba(16, 185, 129, 0.85)',
                    borderColor: 'rgba(16, 185, 129, 1)',
                    borderWidth: 1,
                    borderRadius: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: { color: '#f1f5f9', font: { family: 'Inter', size: 13, weight: '500' }, padding: 16 }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.raw || 0;
                            const total = (submittedData[ctx.dataIndex] || 0) + (notSubData[ctx.dataIndex] || 0) + (delayedData[ctx.dataIndex] || 0);
                            const pct = total > 0 ? ((val / total) * 100).toFixed(1) + '%' : '0%';
                            return `${ctx.dataset.label}: ${val} docs (${pct})`;
                        },
                        footer: (tooltipItems) => {
                            if (!tooltipItems.length) return '';
                            const idx = tooltipItems[0].dataIndex;
                            const total = (submittedData[idx] || 0) + (notSubData[idx] || 0) + (delayedData[idx] || 0);
                            return `Total Discipline Docs: ${total}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { color: '#e2e8f0', font: { family: 'Inter', weight: '600' } }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.06)' },
                    ticks: { color: '#94a3b8', stepSize: 20 }
                }
            }
        }
    });
}

function renderDelayedSection(filteredList = null, totalLen = null) {
    const dl = DASHBOARD_DATA.delay_lookahead || {};
    const delayedBody = document.getElementById('delayedBody');
    if (!delayedBody) return;

    const allDelayed = dl.delayed || [];
    const filtered = filteredList !== null ? filteredList : allDelayed;
    const total = totalLen !== null ? totalLen : allDelayed.length;

    const infoEl = document.getElementById('delayedCountInfo');
    if (infoEl) infoEl.textContent = `Showing ${filtered.length} of ${total} delayed items`;

    let html = '';
    filtered.forEach((item, idx) => {
        const typeBadge = item.delay_type_code === '3.1'
            ? '<span class="status-badge" style="background: rgba(249, 115, 22, 0.2); color: #fb923c; border: 1px solid #fb923c;">🟠 3.1: Submitted Late</span>'
            : '<span class="status-badge" style="background: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #f87171;">🔴 3.2: Overdue (Not Issued)</span>';
        
        const submitDateCell = item.submit_date
            ? `<span style="color: var(--accent-emerald); font-weight: 600;">${formatDate(item.submit_date)}</span>`
            : '<span style="color: var(--text-muted); font-style: italic;">Not Issued</span>';

        html += `
            <tr>
                <td>${idx + 1}</td>
                <td>${typeBadge}</td>
                <td><b>${item.wp}</b></td>
                <td>${item.discipline}</td>
                <td style="font-family: monospace; color: var(--accent-rose);">${item.doc_no}</td>
                <td style="max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.title}">${item.title}</td>
                <td><span class="status-badge not-submitted">${item.milestone}</span></td>
                <td>${formatDate(item.plan_date)}</td>
                <td>${formatDate(item.forecast_date)}</td>
                <td>${submitDateCell}</td>
                <td><span class="aging-badge critical">${item.delay_days} days late</span></td>
            </tr>
        `;
    });

    if (filtered.length === 0) {
        html = `<tr><td colspan="11" style="text-align: center; padding: 1.5rem; color: var(--text-muted);">✅ No delayed documents found matching filters.</td></tr>`;
    }
    delayedBody.innerHTML = html;
}

function renderOverdueSummaryTable() {
    const tbody = document.getElementById('overdueSummaryBody');
    if (!tbody) return;

    const summaryList = DASHBOARD_DATA.overdue_summary?.disciplines || [];
    let html = '';

    summaryList.forEach(disc => {
        html += `
            <tr>
                <td><b>${disc.discipline}</b></td>
                <td class="var-negative">${disc.total}</td>
                <td>${disc.first_rev}</td>
                <td>${disc.ifa_afc_issued}</td>
                <td>${disc.pending_return}</td>
            </tr>
        `;
    });

    if (summaryList.length === 0) {
        html = `<tr><td colspan="5" style="text-align: center; padding: 1.5rem; color: var(--text-muted);">No overdue summary data extracted.</td></tr>`;
    }
    tbody.innerHTML = html;
}

// ─── Tab 5: S-Curve Tab ─────────────────────────────────────────────────────

function renderScurveTab() {
    const scurves = DASHBOARD_DATA.scurves || {};
    const controls = document.getElementById('scurveControls');
    if (!controls) return;

    const labels = Object.keys(scurves);
    if (labels.length === 0) return;

    // Render controls
    controls.innerHTML = labels.map(label => `
        <button class="chart-btn ${label === activeScurveLabel ? 'active' : ''}" onclick="switchScurve('${label}', this)">
            ${label}
        </button>
    `).join('');

    renderActiveScurve();
}

function switchScurve(label, btnElement) {
    activeScurveLabel = label;
    document.querySelectorAll('#scurveControls .chart-btn').forEach(b => b.classList.remove('active'));
    if (btnElement) btnElement.classList.add('active');
    renderActiveScurve();
}

function renderActiveScurve() {
    const scurveData = DASHBOARD_DATA.scurves?.[activeScurveLabel];
    if (!scurveData) return;

    // ── Chart: weekly time-series S-Curve ──────────────────────
    const ctx = document.getElementById('scurveChart')?.getContext('2d');
    if (ctx && scurveData.series && scurveData.series.length > 0) {
        const series = scurveData.series;

        // X labels: short date strings
        const dateLabels = series.map(d => {
            try {
                const dt = new Date(d.date);
                return dt.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
            } catch { return d.date; }
        });

        const planCum     = series.map(d => d.plan_cum     != null ? +(d.plan_cum     * 100).toFixed(4) : null);
        const forecastCum = series.map(d => d.forecast_cum != null ? +(d.forecast_cum * 100).toFixed(4) : null);
        const actualCum   = series.map(d => d.actual_cum   != null ? +(d.actual_cum   * 100).toFixed(4) : null);

        if (scurveChartInstance) scurveChartInstance.destroy();

        scurveChartInstance = new Chart(ctx, {
            type: 'line',
            data: {
                labels: dateLabels,
                datasets: [
                    {
                        label: 'Plan Cumulative %',
                        data: planCum,
                        borderColor: '#6366f1',
                        backgroundColor: 'rgba(99, 102, 241, 0.08)',
                        borderWidth: 2,
                        tension: 0.35,
                        fill: false,
                        pointRadius: 2,
                        pointHoverRadius: 5,
                        spanGaps: true
                    },
                    {
                        label: 'Actual Cumulative %',
                        data: actualCum,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.08)',
                        borderWidth: 3,
                        tension: 0.35,
                        fill: false,
                        pointRadius: 2,
                        pointHoverRadius: 5,
                        spanGaps: false
                    },
                    {
                        label: 'Forecast Cumulative %',
                        data: forecastCum,
                        borderColor: '#f59e0b',
                        borderDash: [5, 5],
                        borderWidth: 2,
                        tension: 0.35,
                        fill: false,
                        pointRadius: 0,
                        pointHoverRadius: 4,
                        spanGaps: true
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 400 },
                plugins: {
                    legend: {
                        labels: { color: '#f1f5f9', font: { family: 'Inter', size: 12 } }
                    },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `${ctx.dataset.label}: ${ctx.raw != null ? ctx.raw.toFixed(2) + '%' : '—'}`
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        max: 100,
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        ticks: {
                            color: '#94a3b8',
                            callback: val => `${val}%`
                        },
                        title: { display: true, text: 'Cumulative Progress (%)', color: '#94a3b8', font: { size: 11 } }
                    },
                    x: {
                        grid: { color: 'rgba(255,255,255,0.03)' },
                        ticks: {
                            color: '#94a3b8',
                            maxTicksLimit: 20,
                            maxRotation: 45,
                            minRotation: 30
                        }
                    }
                }
            }
        });
    }

    // ── Table: current-week snapshot ───────────────────────────
    const tbody = document.getElementById('scurveBody');
    const tableData = scurveData.summary || scurveData.data || [];
    if (tbody) {
        let html = '';
        tableData.forEach(row => {
            const devVal = row.deviation;
            let devClass = 'var-neutral';
            if (devVal > 0) devClass = 'var-positive';
            else if (devVal < -0.01) devClass = 'var-negative';
            else if (devVal < 0) devClass = 'var-warning';

            html += `
                <tr>
                    <td><b>${row.phase}</b></td>
                    <td>${formatPct(row.plan_incr)}</td>
                    <td>${formatPct(row.plan_cum)}</td>
                    <td>${formatPct(row.forecast_incr)}</td>
                    <td>${formatPct(row.forecast_cum)}</td>
                    <td>${formatPct(row.actual_incr)}</td>
                    <td>${formatPct(row.actual_cum)}</td>
                    <td class="${devClass}">${formatPct(row.deviation)}</td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
    }
}


// ─── File Upload Modal & Drag-Drop ──────────────────────────────────────────

function openUploadModal() {
    document.getElementById('uploadOverlay')?.classList.add('active');
}

function closeUploadModal() {
    document.getElementById('uploadOverlay')?.classList.remove('active');
}

function setupDragAndDrop() {
    const dropzone = document.getElementById('dropzone');
    if (!dropzone) return;

    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
        }, false);
    });

    dropzone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            uploadFile(files[0]);
        }
    }, false);
}

function handleFileSelect(event) {
    const files = event.target.files;
    if (files.length > 0) {
        uploadFile(files[0]);
    }
}

async function uploadFile(file) {
    if (!file.name.endsWith('.xlsx')) {
        showToast('Only .xlsx files are allowed', 'error');
        return;
    }

    const progressDiv = document.getElementById('uploadProgress');
    const progressBar = document.getElementById('uploadProgressBar');
    const progressText = document.getElementById('uploadProgressText');

    if (progressDiv) progressDiv.style.display = 'block';
    if (progressBar) progressBar.style.width = '30%';
    if (progressText) progressText.textContent = `Uploading ${file.name}...`;

    const formData = new FormData();
    formData.append('file', file);

    try {
        if (progressBar) progressBar.style.width = '60%';
        if (progressText) progressText.textContent = `Processing and extracting data...`;

        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const errJson = await response.json().catch(() => ({}));
            throw new Error(errJson.error || `Upload failed (HTTP ${response.status})`);
        }

        const resData = await response.json();
        if (resData.error) throw new Error(resData.error);

        if (progressBar) progressBar.style.width = '100%';
        if (progressText) progressText.textContent = `Complete!`;

        showToast(resData.message || 'File uploaded and data refreshed successfully', 'success');
        
        if (resData.data) {
            DASHBOARD_DATA = resData.data;
            renderAll();
        } else {
            fetchData();
        }

        setTimeout(() => {
            closeUploadModal();
            if (progressDiv) progressDiv.style.display = 'none';
        }, 1200);

    } catch (err) {
        console.error('Upload failed:', err);
        showToast('Upload error: ' + err.message, 'error');
        if (progressText) progressText.textContent = `Error: ${err.message}`;
    }
}

// ─── Utility Functions ──────────────────────────────────────────────────────

function formatPct(val) {
    if (val === null || val === undefined || isNaN(val) || val === 'N/A') return '—';
    return (val * 100).toFixed(2) + '%';
}

function formatDate(dateStr) {
    if (!dateStr || dateStr === 'N/A' || dateStr === '-') return '—';
    try {
        const dt = new Date(dateStr);
        if (isNaN(dt.getTime()) || dt.getFullYear() < 1950) return '—';
        return dt.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    } catch {
        return dateStr;
    }
}

function showToast(message, type = 'info') {
    const existing = document.querySelector('.toast');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span>${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}</span> <span>${message}</span>`;
    document.body.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.3s';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}
