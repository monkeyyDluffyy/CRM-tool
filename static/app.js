/**
 * NexusCRM — Frontend Application
 * Handles data fetching, rendering, CRUD, charts, and live Excel sync
 */

// ═══════════════════════════════════════
// STATE
// ═══════════════════════════════════════

const API_BASE = '';
const POLL_INTERVAL = 3000; // Check for Excel changes every 3 seconds

const state = {
    currentView: 'dashboard',
    currentSheet: null,
    allData: {},
    dashboardStats: {},
    sortColumn: null,
    sortDirection: 'asc',
    searchQuery: '',
    editingRow: null,
    charts: {},
    pollTimer: null,
    lastTimestamp: 0,
};

// ═══════════════════════════════════════
// INITIALIZATION
// ═══════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    initSearch();
    initModal();
    initRefresh();
    initSidebarToggle();
    loadAllData();
    startPolling();
});

// ═══════════════════════════════════════
// DATA FETCHING
// ═══════════════════════════════════════

async function loadAllData() {
    try {
        const [dataRes, dashRes] = await Promise.all([
            fetch(`${API_BASE}/api/data`),
            fetch(`${API_BASE}/api/dashboard`)
        ]);

        const newData = await dataRes.json();
        const newDash = await dashRes.json();

        state.allData = newData;
        state.dashboardStats = newDash;

        updateBadges();
        renderCurrentView();
        updateSyncStatus(true);
    } catch (err) {
        console.error('Failed to load data:', err);
        showToast('Failed to load data from server', 'error');
        updateSyncStatus(false);
    }
}

// Silent background sync — fetches fresh data from Excel every 3 seconds
async function checkForUpdates() {
    try {
        const [dataRes, dashRes] = await Promise.all([
            fetch(`${API_BASE}/api/data`),
            fetch(`${API_BASE}/api/dashboard`)
        ]);

        const newData = await dataRes.json();
        const newDash = await dashRes.json();

        // Check if data actually changed
        const oldJson = JSON.stringify(state.allData);
        const newJson = JSON.stringify(newData);

        if (oldJson !== newJson) {
            state.allData = newData;
            state.dashboardStats = newDash;
            updateBadges();
            renderCurrentView();
            updateSyncStatus(true);
            console.log('[SYNC] Data updated from Excel');
        }
    } catch (err) {
        // Silently fail on poll errors
    }
}

function startPolling() {
    if (state.pollTimer) clearInterval(state.pollTimer);
    state.pollTimer = setInterval(checkForUpdates, POLL_INTERVAL);
}

// ═══════════════════════════════════════
// NAVIGATION
// ═══════════════════════════════════════

function initNavigation() {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const view = item.dataset.view;
            navigateTo(view);
        });
    });
}

function navigateTo(view) {
    // Update active nav
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    const navItem = document.querySelector(`.nav-item[data-view="${view}"]`);
    if (navItem) navItem.classList.add('active');

    state.currentView = view;

    const titles = {
        'dashboard': { title: 'Dashboard', subtitle: 'Overview of your CRM metrics' },
        'Customers': { title: 'Customers', subtitle: 'Manage your customer accounts' },
        'Interactions': { title: 'Interactions', subtitle: 'Track all customer interactions' },
        'Deals': { title: 'Deals', subtitle: 'Monitor your sales pipeline' },
        'Support Tickets': { title: 'Support Tickets', subtitle: 'Handle customer support requests' },
    };

    const info = titles[view] || { title: view, subtitle: '' };
    document.getElementById('page-title').textContent = info.title;
    document.getElementById('page-subtitle').textContent = info.subtitle;

    // Show/hide search box based on view
    const searchBox = document.getElementById('search-box');
    if (view === 'dashboard') {
        searchBox.style.display = 'none';
    } else {
        searchBox.style.display = 'block';
    }

    renderCurrentView();

    // Close sidebar on mobile
    document.getElementById('sidebar').classList.remove('open');
}

function renderCurrentView() {
    const dashView = document.getElementById('view-dashboard');
    const tableView = document.getElementById('view-table');

    if (state.currentView === 'dashboard') {
        dashView.classList.add('active');
        tableView.classList.remove('active');
        renderDashboard();
    } else {
        dashView.classList.remove('active');
        tableView.classList.add('active');
        state.currentSheet = state.currentView;
        renderTable();
    }
}

// ═══════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════

function renderDashboard() {
    const stats = state.dashboardStats;
    if (!stats.customers) return;

    // Animate KPIs
    animateValue('kpi-customers', stats.customers.total);
    animateValue('kpi-mrr', stats.customers.total_mrr, true);
    animateValue('kpi-pipeline', stats.deals.total_pipeline, true);
    animateValue('kpi-health', stats.customers.avg_health);
    animateValue('kpi-won', stats.deals.won);
    animateValue('kpi-tickets', stats.tickets.open);

    renderCharts();
    renderCustomerCards();
}

function animateValue(elementId, target, isCurrency = false) {
    const el = document.getElementById(elementId);
    if (!el) return;

    const duration = 1000;
    const start = 0;
    const startTime = performance.now();

    function update(currentTime) {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3); // ease-out cubic
        const current = Math.round(start + (target - start) * eased);

        if (isCurrency) {
            el.textContent = '$' + current.toLocaleString();
        } else {
            el.textContent = current.toLocaleString();
        }

        if (progress < 1) {
            requestAnimationFrame(update);
        }
    }

    requestAnimationFrame(update);
}

// ═══════════════════════════════════════
// CHARTS
// ═══════════════════════════════════════

function renderCharts() {
    const stats = state.dashboardStats;

    // Destroy existing charts
    Object.values(state.charts).forEach(c => c.destroy());
    state.charts = {};

    const chartDefaults = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'bottom',
                labels: {
                    color: '#94a3b8',
                    padding: 16,
                    usePointStyle: true,
                    pointStyleWidth: 10,
                    font: { family: 'Inter', size: 11 }
                }
            }
        }
    };

    // Tier Distribution (Doughnut)
    const tierCtx = document.getElementById('chart-tiers');
    if (tierCtx) {
        const tierData = stats.customers.tier_distribution;
        state.charts.tiers = new Chart(tierCtx, {
            type: 'doughnut',
            data: {
                labels: Object.keys(tierData),
                datasets: [{
                    data: Object.values(tierData),
                    backgroundColor: ['#8b5cf6', '#fbbf24', '#94a3b8', '#fb923c'],
                    borderWidth: 0,
                    hoverOffset: 8,
                }]
            },
            options: {
                ...chartDefaults,
                cutout: '65%',
                plugins: {
                    ...chartDefaults.plugins,
                }
            }
        });
    }

    // Deal Pipeline (Bar)
    const dealCtx = document.getElementById('chart-deals');
    if (dealCtx) {
        const stageData = stats.deals.stage_distribution;
        const stageColors = {
            'Proposal': '#818cf8', 'Negotiation': '#fbbf24', 'Demo': '#22d3ee',
            'Closed Won': '#34d399', 'Closed Lost': '#fb7185',
        };
        state.charts.deals = new Chart(dealCtx, {
            type: 'bar',
            data: {
                labels: Object.keys(stageData),
                datasets: [{
                    data: Object.values(stageData),
                    backgroundColor: Object.keys(stageData).map(k => stageColors[k] || '#6366f1'),
                    borderRadius: 8,
                    borderSkipped: false,
                    maxBarThickness: 44,
                }]
            },
            options: {
                ...chartDefaults,
                plugins: { ...chartDefaults.plugins, legend: { display: false } },
                scales: {
                    x: {
                        ticks: { color: '#64748b', font: { family: 'Inter', size: 11 } },
                        grid: { display: false },
                        border: { display: false },
                    },
                    y: {
                        beginAtZero: true,
                        ticks: { color: '#64748b', stepSize: 1, font: { family: 'Inter', size: 11 } },
                        grid: { color: 'rgba(255,255,255,0.04)' },
                        border: { display: false },
                    }
                }
            }
        });
    }

    // Revenue by Tier (Horizontal Bar)
    const revCtx = document.getElementById('chart-revenue');
    if (revCtx) {
        const mrrData = stats.customers.mrr_by_tier;
        const tierColors = { 'Platinum': '#8b5cf6', 'Gold': '#fbbf24', 'Silver': '#94a3b8', 'Bronze': '#fb923c' };
        state.charts.revenue = new Chart(revCtx, {
            type: 'bar',
            data: {
                labels: Object.keys(mrrData),
                datasets: [{
                    data: Object.values(mrrData),
                    backgroundColor: Object.keys(mrrData).map(k => tierColors[k] || '#6366f1'),
                    borderRadius: 8,
                    borderSkipped: false,
                    maxBarThickness: 32,
                }]
            },
            options: {
                ...chartDefaults,
                indexAxis: 'y',
                plugins: { ...chartDefaults.plugins, legend: { display: false } },
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: {
                            color: '#64748b',
                            callback: v => '$' + v.toLocaleString(),
                            font: { family: 'Inter', size: 11 }
                        },
                        grid: { color: 'rgba(255,255,255,0.04)' },
                        border: { display: false },
                    },
                    y: {
                        ticks: { color: '#94a3b8', font: { family: 'Inter', size: 11 } },
                        grid: { display: false },
                        border: { display: false },
                    }
                }
            }
        });
    }

    // Ticket Priority (Polar Area)
    const ticketCtx = document.getElementById('chart-tickets');
    if (ticketCtx) {
        const prioData = stats.tickets.priority_distribution;
        const prioColors = { 'High': '#fb7185', 'Medium': '#fbbf24', 'Low': '#34d399' };
        state.charts.tickets = new Chart(ticketCtx, {
            type: 'polarArea',
            data: {
                labels: Object.keys(prioData),
                datasets: [{
                    data: Object.values(prioData),
                    backgroundColor: Object.keys(prioData).map(k => {
                        const c = prioColors[k] || '#6366f1';
                        return c + '44';
                    }),
                    borderColor: Object.keys(prioData).map(k => prioColors[k] || '#6366f1'),
                    borderWidth: 2,
                }]
            },
            options: {
                ...chartDefaults,
                scales: {
                    r: {
                        ticks: { display: false },
                        grid: { color: 'rgba(255,255,255,0.06)' },
                    }
                }
            }
        });
    }
}

function renderCustomerCards() {
    const container = document.getElementById('customer-cards');
    if (!container) return;

    const customers = state.allData['Customers']?.data || [];
    container.innerHTML = '';

    customers.forEach(c => {
        const healthScore = parseFloat(c['Health Score']) || 0;
        const healthClass = healthScore >= 75 ? 'health-high' : healthScore >= 50 ? 'health-medium' : 'health-low';
        const tierClass = `tier-${(c['Customer Tier'] || '').toLowerCase()}`;
        const mrr = parseFloat(c['MRR (USD)']) || 0;
        const arr = parseFloat(c['ARR (USD)']) || 0;
        const nps = c['NPS Score'] || 'N/A';

        const card = document.createElement('div');
        card.className = 'customer-card';
        card.innerHTML = `
            <div class="customer-card-header">
                <div>
                    <div class="customer-card-name">${escapeHtml(c['Company Name'] || '')}</div>
                    <div class="customer-card-industry">${escapeHtml(c['Industry'] || '')} · ${escapeHtml(c['City'] || '')}</div>
                </div>
                <span class="customer-tier-badge ${tierClass}">${escapeHtml(c['Customer Tier'] || '')}</span>
            </div>
            <div class="customer-card-stats">
                <div class="card-stat">
                    <span class="card-stat-label">MRR</span>
                    <span class="card-stat-value">$${mrr.toLocaleString()}</span>
                </div>
                <div class="card-stat">
                    <span class="card-stat-label">ARR</span>
                    <span class="card-stat-value">$${arr.toLocaleString()}</span>
                </div>
                <div class="card-stat">
                    <span class="card-stat-label">NPS</span>
                    <span class="card-stat-value">${nps}</span>
                </div>
                <div class="card-stat">
                    <span class="card-stat-label">Health</span>
                    <span class="card-stat-value">${healthScore}%</span>
                </div>
            </div>
            <div class="health-bar">
                <div class="health-bar-fill ${healthClass}" style="width: ${healthScore}%"></div>
            </div>
        `;

        card.addEventListener('click', () => navigateTo('Customers'));
        container.appendChild(card);
    });
}

// ═══════════════════════════════════════
// DATA TABLE
// ═══════════════════════════════════════

function renderTable() {
    const sheetData = state.allData[state.currentSheet];
    if (!sheetData) return;

    const { headers, data } = sheetData;
    const filtered = filterData(data);
    const sorted = sortData(filtered, headers);

    renderTableHead(headers);
    renderTableBody(sorted, headers);
    updateRecordCount(sorted.length);
}

function renderTableHead(headers) {
    const thead = document.getElementById('table-head');
    thead.innerHTML = '';

    const tr = document.createElement('tr');
    // Actions column (Prepend)
    const actionTh = document.createElement('th');
    actionTh.textContent = 'Actions';
    actionTh.style.width = '90px';
    actionTh.className = 'sticky-action-col sticky-action-left';
    tr.appendChild(actionTh);

    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;

        const sortIcon = document.createElement('span');
        sortIcon.className = 'sort-icon';
        if (state.sortColumn === h) {
            th.classList.add('sorted');
            sortIcon.innerHTML = state.sortDirection === 'asc'
                ? '<i class="fas fa-sort-up"></i>'
                : '<i class="fas fa-sort-down"></i>';
        } else {
            sortIcon.innerHTML = '<i class="fas fa-sort"></i>';
        }
        th.appendChild(sortIcon);

        th.addEventListener('click', () => {
            if (state.sortColumn === h) {
                state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                state.sortColumn = h;
                state.sortDirection = 'asc';
            }
            renderTable();
        });

        tr.appendChild(th);
    });

    thead.appendChild(tr);
}

function renderTableBody(data, headers) {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (data.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = headers.length + 1;
        td.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-inbox"></i>
                <h3>No records found</h3>
                <p>Try adjusting your search or add a new record.</p>
            </div>
        `;
        tr.appendChild(td);
        tbody.appendChild(tr);
        return;
    }

    data.forEach((row, index) => {
        const tr = document.createElement('tr');

        // Action buttons (Prepend)
        const actionTd = document.createElement('td');
        actionTd.className = 'sticky-action-col sticky-action-left';
        actionTd.innerHTML = `
            <div class="row-actions">
                <button class="btn-edit" title="Edit" data-index="${index}">
                    <i class="fas fa-pen"></i>
                </button>
                <button class="btn-delete" title="Delete" data-index="${index}">
                    <i class="fas fa-trash-alt"></i>
                </button>
            </div>
        `;
        tr.appendChild(actionTd);

        headers.forEach(h => {
            const td = document.createElement('td');
            const val = row[h] != null ? String(row[h]) : '';
            td.innerHTML = formatCell(h, val);
            td.title = val;
            tr.appendChild(td);
        });



        // Find original index for edit/delete (in case of filtering)
        const originalIndex = state.allData[state.currentSheet].data.indexOf(row);

        actionTd.querySelector('.btn-edit').addEventListener('click', (e) => {
            e.stopPropagation();
            openEditModal(originalIndex);
        });

        actionTd.querySelector('.btn-delete').addEventListener('click', (e) => {
            e.stopPropagation();
            openDeleteModal(originalIndex);
        });

        tbody.appendChild(tr);
    });
}

function formatCell(header, value) {
    const h = header.toLowerCase();
    const v = String(value).trim();

    // ID columns
    if (h.includes(' id') && h !== 'customer id' || h === 'customer id' && v.startsWith('CUST')) {
        if (v.match(/^[A-Z]+-\d+$/)) {
            return `<span class="cell-id">${escapeHtml(v)}</span>`;
        }
    }

    // Currency
    if (h.includes('usd') || h.includes('value') || h === 'mrr (usd)' || h === 'arr (usd)') {
        const num = parseFloat(v);
        if (!isNaN(num)) {
            return `<span class="cell-currency">$${num.toLocaleString()}</span>`;
        }
    }

    // Percentage
    if (h.includes('probability') || h.includes('score') || h === 'health score') {
        const num = parseFloat(v);
        if (!isNaN(num)) {
            return `<span class="cell-percentage">${num}</span>`;
        }
    }

    // Status / Outcome
    if (h === 'status' || h === 'outcome') {
        const statusClass = getStatusClass(v);
        return `<span class="cell-status ${statusClass}"><span class="pulse-dot" style="width:5px;height:5px;animation:none;background:currentColor"></span>${escapeHtml(v)}</span>`;
    }

    // Stage
    if (h === 'stage') {
        const stageClass = getStageClass(v);
        return `<span class="cell-status ${stageClass}">${escapeHtml(v)}</span>`;
    }

    // Priority
    if (h === 'priority') {
        const prioClass = `priority-${v.toLowerCase()}`;
        return `<span class="cell-priority ${prioClass}"><i class="fas fa-flag" style="font-size:10px"></i>${escapeHtml(v)}</span>`;
    }

    // Customer Tier
    if (h === 'customer tier') {
        const tierClass = `tier-${v.toLowerCase()}`;
        return `<span class="customer-tier-badge ${tierClass}">${escapeHtml(v)}</span>`;
    }

    // Type
    if (h === 'type') {
        const icons = { 'Call': 'fa-phone', 'Email': 'fa-envelope', 'Meeting': 'fa-calendar' };
        const icon = icons[v] || 'fa-comment';
        return `<i class="fas ${icon}" style="color:var(--text-muted);margin-right:6px;font-size:11px"></i>${escapeHtml(v)}`;
    }

    return escapeHtml(v);
}

function getStatusClass(status) {
    const map = {
        'Positive': 'status-positive', 'Resolved': 'status-resolved',
        'Pending': 'status-pending', 'In Progress': 'status-in-progress',
        'Open': 'status-open', 'Neutral': 'status-neutral',
    };
    return map[status] || 'status-neutral';
}

function getStageClass(stage) {
    const map = {
        'Closed Won': 'status-closed-won', 'Closed Lost': 'status-closed-lost',
        'Negotiation': 'status-negotiation', 'Demo': 'status-demo',
        'Proposal': 'status-proposal',
    };
    return map[stage] || 'status-neutral';
}

function filterData(data) {
    const query = state.searchQuery.toLowerCase();
    if (!query) return [...data];
    return data.filter(row =>
        Object.values(row).some(v =>
            String(v).toLowerCase().includes(query)
        )
    );
}

function sortData(data, headers) {
    if (!state.sortColumn || !headers.includes(state.sortColumn)) return data;

    return [...data].sort((a, b) => {
        let va = a[state.sortColumn] ?? '';
        let vb = b[state.sortColumn] ?? '';

        // Try numeric sort
        const na = parseFloat(va);
        const nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) {
            return state.sortDirection === 'asc' ? na - nb : nb - na;
        }

        // String sort
        va = String(va).toLowerCase();
        vb = String(vb).toLowerCase();
        if (va < vb) return state.sortDirection === 'asc' ? -1 : 1;
        if (va > vb) return state.sortDirection === 'asc' ? 1 : -1;
        return 0;
    });
}

function updateRecordCount(count) {
    const el = document.getElementById('record-count');
    if (el) el.textContent = `${count} record${count !== 1 ? 's' : ''}`;
}

// ═══════════════════════════════════════
// SEARCH
// ═══════════════════════════════════════

function initSearch() {
    const globalSearch = document.getElementById('global-search');
    const tableSearch = document.getElementById('table-search-input');

    globalSearch.addEventListener('input', (e) => {
        state.searchQuery = e.target.value;
        tableSearch.value = e.target.value;
        if (state.currentView !== 'dashboard') renderTable();
    });

    tableSearch.addEventListener('input', (e) => {
        state.searchQuery = e.target.value;
        globalSearch.value = e.target.value;
        renderTable();
    });
}

// ═══════════════════════════════════════
// MODAL / CRUD
// ═══════════════════════════════════════

function initModal() {
    // Add record button
    document.getElementById('btn-add-record').addEventListener('click', () => {
        openAddModal();
    });

    // Close modals
    document.getElementById('modal-close').addEventListener('click', closeModal);
    document.getElementById('modal-cancel').addEventListener('click', closeModal);
    document.getElementById('modal-overlay').addEventListener('click', (e) => {
        if (e.target === e.currentTarget) closeModal();
    });

    document.getElementById('delete-modal-close').addEventListener('click', closeDeleteModal);
    document.getElementById('delete-cancel').addEventListener('click', closeDeleteModal);
    document.getElementById('delete-modal-overlay').addEventListener('click', (e) => {
        if (e.target === e.currentTarget) closeDeleteModal();
    });

    // Save
    document.getElementById('modal-save').addEventListener('click', handleSave);
}

function openAddModal() {
    const sheetData = state.allData[state.currentSheet];
    if (!sheetData) return;

    state.editingRow = null;
    document.getElementById('modal-title').textContent = `Add ${getSingular(state.currentSheet)}`;
    renderModalForm(sheetData.headers, {});
    document.getElementById('modal-overlay').classList.add('active');
}

function openEditModal(rowIndex) {
    const sheetData = state.allData[state.currentSheet];
    if (!sheetData) return;

    state.editingRow = rowIndex;
    const row = sheetData.data[rowIndex];

    document.getElementById('modal-title').textContent = `Edit ${getSingular(state.currentSheet)}`;
    renderModalForm(sheetData.headers, row);
    document.getElementById('modal-overlay').classList.add('active');
}

function renderModalForm(headers, rowData) {
    const body = document.getElementById('modal-body');
    const formGrid = document.createElement('div');
    formGrid.className = 'form-grid';

    headers.forEach(h => {
        const group = document.createElement('div');
        group.className = 'form-group';

        // Make description fields full width
        if (h.toLowerCase().includes('description') || h.toLowerCase().includes('notes') || h.toLowerCase().includes('action')) {
            group.classList.add('full-width');
        }

        const label = document.createElement('label');
        label.textContent = h;
        label.htmlFor = `field-${h}`;

        let input;

        // Choose input type based on header
        const hLower = h.toLowerCase();
        if (hLower.includes('description') || hLower.includes('notes') || hLower.includes('action')) {
            input = document.createElement('textarea');
        } else if (hLower === 'customer tier') {
            input = createSelect(['Platinum', 'Gold', 'Silver', 'Bronze']);
        } else if (hLower === 'stage') {
            input = createSelect(['Proposal', 'Negotiation', 'Demo', 'Closed Won', 'Closed Lost']);
        } else if (hLower === 'priority') {
            input = createSelect(['High', 'Medium', 'Low']);
        } else if (hLower === 'status') {
            input = createSelect(['Open', 'In Progress', 'Resolved']);
        } else if (hLower === 'outcome') {
            input = createSelect(['Positive', 'Pending', 'Neutral']);
        } else if (hLower === 'type') {
            input = createSelect(['Call', 'Email', 'Meeting']);
        } else if (hLower === 'category') {
            input = createSelect(['Integration', 'Billing', 'Feature Request', 'Performance', 'Other']);
        } else if (hLower.includes('date')) {
            input = document.createElement('input');
            input.type = 'date';
        } else if (hLower.includes('usd') || hLower.includes('value') || hLower.includes('score') ||
                   hLower.includes('probability') || hLower.includes('duration') || hLower === 'csat score') {
            input = document.createElement('input');
            input.type = 'number';
        } else {
            input = document.createElement('input');
            input.type = 'text';
        }

        input.id = `field-${h}`;
        input.name = h;
        input.value = rowData[h] != null ? String(rowData[h]) : '';

        group.appendChild(label);
        group.appendChild(input);
        formGrid.appendChild(group);
    });

    body.innerHTML = '';
    body.appendChild(formGrid);
}

function createSelect(options) {
    const select = document.createElement('select');
    const defaultOpt = document.createElement('option');
    defaultOpt.value = '';
    defaultOpt.textContent = 'Select...';
    select.appendChild(defaultOpt);
    options.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o;
        opt.textContent = o;
        select.appendChild(opt);
    });
    return select;
}

async function handleSave() {
    const sheetData = state.allData[state.currentSheet];
    if (!sheetData) return;

    const formData = {};
    sheetData.headers.forEach(h => {
        const input = document.getElementById(`field-${h}`);
        if (input) {
            formData[h] = input.value;
        }
    });

    try {
        let res;
        if (state.editingRow !== null) {
            // Update existing
            res = await fetch(`${API_BASE}/api/data/${state.currentSheet}/${state.editingRow}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(formData),
            });
        } else {
            // Add new
            res = await fetch(`${API_BASE}/api/data/${state.currentSheet}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(formData),
            });
        }

        const result = await res.json();
        if (result.success) {
            showToast(result.message + ' — Excel file updated', 'success');
            closeModal();
            await loadAllData();
        } else {
            showToast(result.error || 'Failed to save', 'error');
        }
    } catch (err) {
        showToast('Error saving record: ' + err.message, 'error');
    }
}

function closeModal() {
    document.getElementById('modal-overlay').classList.remove('active');
    state.editingRow = null;
}

// ─── Delete Modal ───

let deleteTargetIndex = null;

function openDeleteModal(rowIndex) {
    deleteTargetIndex = rowIndex;
    document.getElementById('delete-modal-overlay').classList.add('active');
}

function closeDeleteModal() {
    document.getElementById('delete-modal-overlay').classList.remove('active');
    deleteTargetIndex = null;
}

document.getElementById('delete-confirm')?.addEventListener('click', async () => {
    if (deleteTargetIndex === null) return;

    try {
        const res = await fetch(`${API_BASE}/api/data/${state.currentSheet}/${deleteTargetIndex}`, {
            method: 'DELETE',
        });
        const result = await res.json();
        if (result.success) {
            showToast('Record deleted — Excel file updated', 'success');
            closeDeleteModal();
            await loadAllData();
        } else {
            showToast(result.error || 'Failed to delete', 'error');
        }
    } catch (err) {
        showToast('Error deleting record: ' + err.message, 'error');
    }
});

// ═══════════════════════════════════════
// UI HELPERS
// ═══════════════════════════════════════

function initRefresh() {
    document.getElementById('btn-refresh').addEventListener('click', async () => {
        const btn = document.getElementById('btn-refresh');
        btn.classList.add('spinning');
        await loadAllData();
        showToast('Data refreshed from Excel file', 'success');
        setTimeout(() => btn.classList.remove('spinning'), 800);
    });
}

function initSidebarToggle() {
    document.getElementById('sidebar-toggle').addEventListener('click', () => {
        document.getElementById('sidebar').classList.toggle('open');
    });
}

function updateBadges() {
    const data = state.allData;
    setBadge('badge-customers', data['Customers']?.data?.length || 0);
    setBadge('badge-interactions', data['Interactions']?.data?.length || 0);
    setBadge('badge-deals', data['Deals']?.data?.length || 0);
    setBadge('badge-tickets', data['Support Tickets']?.data?.length || 0);
}

function setBadge(id, count) {
    const el = document.getElementById(id);
    if (el) el.textContent = count;
}

function updateSyncStatus(synced) {
    const dot = document.querySelector('.sync-dot');
    const label = document.querySelector('.sync-status span');
    const timeEl = document.getElementById('last-sync-time');

    if (synced) {
        dot?.classList.remove('error');
        if (label) label.textContent = 'Synced with Excel';
    } else {
        dot?.classList.add('error');
        if (label) label.textContent = 'Connection error';
    }

    if (timeEl) {
        const now = new Date();
        timeEl.textContent = `Last sync: ${now.toLocaleTimeString()}`;
    }
}

function getSingular(sheetName) {
    const map = {
        'Customers': 'Customer',
        'Interactions': 'Interaction',
        'Deals': 'Deal',
        'Support Tickets': 'Ticket',
    };
    return map[sheetName] || sheetName;
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ═══════════════════════════════════════
// TOAST SYSTEM
// ═══════════════════════════════════════

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');

    const icons = {
        success: 'fa-check',
        error: 'fa-exclamation',
        info: 'fa-info',
        warning: 'fa-triangle-exclamation',
    };

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <div class="toast-icon"><i class="fas ${icons[type]}"></i></div>
        <span class="toast-message">${escapeHtml(message)}</span>
        <button class="toast-close"><i class="fas fa-times"></i></button>
    `;

    toast.querySelector('.toast-close').addEventListener('click', () => removeToast(toast));
    container.appendChild(toast);

    setTimeout(() => removeToast(toast), 4000);
}

function removeToast(toast) {
    toast.classList.add('toast-exit');
    setTimeout(() => toast.remove(), 300);
}

// Keyboard shortcut: Escape to close modals
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        closeModal();
        closeDeleteModal();
    }
});
