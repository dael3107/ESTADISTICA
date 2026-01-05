// State
let rawData = [];
let workbook = null;
let mainChartInstance = null;
let pieChartInstance = null;
let annualChartInstance = null;
let compChartInstance = null;

// New Chart Instances for Advanced Reports
let compReportChartInstance = null;
let sellerChartInstance = null;
let productChartInstance = null;

let globalClientColIdx = -1;
let globalCNuevoIdx = -1;
let globalVendedorIdx = -1;
let globalMonthCols = [];

// Elements
const dropZone = document.getElementById('upload-section');
const fileInput = document.getElementById('file-input');
const columnMapper = document.getElementById('column-mapper');
const dashboardView = document.getElementById('dashboard-view');
const annualView = document.getElementById('annual-view');
const comparisonView = document.getElementById('comparison-view');
const historyModal = document.getElementById('history-modal');
const compareModal = document.getElementById('comparison-modal');

// Buttons
const btnAnnual = document.getElementById('annual-btn');
const btnHistory = document.getElementById('history-btn');
const btnCompare = document.getElementById('compare-btn');
const backFromAnnualSum = document.getElementById('back-from-annual-btn');
const backFromComp = document.getElementById('back-to-dash-btn');
const closeHistoryBtn = document.getElementById('close-history-btn');
const searchActionBtn = document.getElementById('search-action-btn');
const runComparisonBtn = document.getElementById('run-comparison-btn');
const closeCompModalBtn = document.getElementById('close-modal-btn');

// New Feature Elements
const btnNoSales = document.getElementById('no-sales-btn');
const btnNewClients = document.getElementById('new-clients-btn');
const noSalesModal = document.getElementById('no-sales-modal');
const closeNoSalesBtn = document.getElementById('close-no-sales-btn');
const newClientsModal = document.getElementById('new-clients-modal');
const closeNewClientsBtn = document.getElementById('close-new-clients-btn');
const backSellersBtn = document.getElementById('back-sellers-btn');

// Advanced Report Buttons
const btnOpenCompReport = document.getElementById('btn-open-comp-report');
const btnOpenSellerReport = document.getElementById('btn-open-seller-report');
const btnOpenProductReport = document.getElementById('btn-open-product-report');

// Advanced Report Modals
const modalCompReport = document.getElementById('modal-comp-report');
const modalSellerReport = document.getElementById('modal-seller-report');
const modalProductReport = document.getElementById('modal-product-report');

// Reset Button
const resetBtn = document.getElementById('reset-btn');
if (resetBtn) {
    resetBtn.addEventListener('click', () => {
        location.reload();
    });
}
const backToHomeBtn = document.getElementById('back-to-home-btn');
if (backToHomeBtn) {
    backToHomeBtn.addEventListener('click', () => {
        dashboardView.classList.add('hidden');
        columnMapper.classList.remove('hidden');
    });
}


// Drag & Drop Handlers
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length) handleFile(files[0]);
});
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

// Sheet Selector Listener (Global)
document.getElementById('sheet-trend').addEventListener('change', (e) => {
    if (workbook) parseColumns(e.target.value);
});

// File Processing
function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });
        populateSheetSelects();
        showMapper();
    };
    reader.readAsArrayBuffer(file);
}

function populateSheetSelects() {
    const selTrend = document.getElementById('sheet-trend');
    const selComp = document.getElementById('sheet-comp');
    const selSeller = document.getElementById('sheet-seller');
    const selProduct = document.getElementById('sheet-product');

    // Clear all
    selTrend.innerHTML = '';
    selComp.innerHTML = '';
    selSeller.innerHTML = '';
    selProduct.innerHTML = '';

    workbook.SheetNames.forEach(name => {
        const upper = name.toUpperCase();
        const option = document.createElement('option');
        option.value = name;
        option.text = name;

        if (upper.includes('TENDENCIA')) {
            selTrend.appendChild(option);
        } else if (upper.includes('COMPARATIVO')) {
            selComp.appendChild(option.cloneNode(true));
        } else if (upper.includes('VENDEDOR')) {
            selSeller.appendChild(option.cloneNode(true));
        } else if (upper.includes('PRODUCTO')) {
            selProduct.appendChild(option.cloneNode(true));
        }
    });

    if (selTrend.options.length > 0) {
        parseColumns(selTrend.options[0].value);
    }
}

function parseColumns(sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (rawData.length === 0) return;

    const headers = rawData[0];
    const filterSelect = document.getElementById('new-client-filter');
    const amtSelect = document.getElementById('amount-col-select');

    globalClientColIdx = -1;
    globalCNuevoIdx = -1;
    globalVendedorIdx = -1;
    globalMonthCols = [];

    filterSelect.innerHTML = '';
    amtSelect.innerHTML = '';

    const monthNames = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

    headers.forEach((h, index) => {
        if (!h) return;
        const lower = String(h).toLowerCase().trim();

        if (globalClientColIdx === -1 && (lower.includes('cliente') || lower.includes('razon') || lower.includes('social'))) {
            globalClientColIdx = index;
        }

        if (globalCNuevoIdx === -1 && (lower === 'c nuevo' || lower === 'c.nuevo')) {
            globalCNuevoIdx = index;
        }

        if (globalVendedorIdx === -1 && (lower === 'vendedor')) {
            globalVendedorIdx = index;
        }

        if (monthNames.includes(lower)) {
            const option = document.createElement('option');
            option.value = index;
            option.text = h;
            amtSelect.appendChild(option);

            globalMonthCols.push({
                name: lower,
                label: h,
                index: index,
                order: monthNames.indexOf(lower)
            });
        }
    });

    // Populate Filter (Sorted Chronologically & Exclude 'SIN VENTA')
    const uniqueCNuevo = new Set();

    // Always add TODOS first
    const todosOpt = document.createElement('option');
    todosOpt.value = 'TODOS';
    todosOpt.text = 'TODOS';
    filterSelect.appendChild(todosOpt);

    if (globalCNuevoIdx !== -1) {
        for (let i = 1; i < rawData.length; i++) {
            const val = rawData[i][globalCNuevoIdx];
            if (val) {
                const sVal = String(val).toUpperCase().trim();
                // Exclude SIN VENTA
                if (sVal !== 'TODOS' && !sVal.includes('SIN VENTA')) {
                    uniqueCNuevo.add(sVal);
                }
            }
        }
    }

    // Sort keys based on monthNames order
    const sortedKeys = Array.from(uniqueCNuevo).sort((a, b) => {
        const ia = monthNames.indexOf(a.toLowerCase());
        const ib = monthNames.indexOf(b.toLowerCase());

        if (ia !== -1 && ib !== -1) return ia - ib;
        if (ia === -1 && ib === -1) return a.localeCompare(b);
        return ia !== -1 ? -1 : 1;
    });

    sortedKeys.forEach(val => {
        const option = document.createElement('option');
        option.value = val;
        option.text = val;
        filterSelect.appendChild(option);
    });
}

function showMapper() {
    dropZone.classList.add('hidden');
    columnMapper.classList.remove('hidden');
}

// Generate Dashboard
document.getElementById('generate-btn').addEventListener('click', () => {
    const sheetName = document.getElementById('sheet-trend').value;
    const filterVal = document.getElementById('new-client-filter').value;
    const amtIdx = document.getElementById('amount-col-select').value;
    const custIdx = globalClientColIdx;

    const amountSelect = document.getElementById('amount-col-select');
    const amountLabel = amountSelect.options[amountSelect.selectedIndex].text;
    const titleEl = document.getElementById('chart-title');
    if (titleEl) titleEl.innerHTML = `Top 10 Clientes <span class="text-blue-500">| ${amountLabel.toUpperCase()}</span>`;

    processStandardData(sheetName, custIdx, amtIdx, filterVal);

    columnMapper.classList.add('hidden');
    dashboardView.classList.remove('hidden');
});

function processStandardData(sheetName, custIdx, amtIdx, filterVal) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const salesMap = {};
    let grandTotal = 0;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let customer = row[custIdx];
        let amount = row[amtIdx];
        let cNuevoVal = globalCNuevoIdx !== -1 ? String(row[globalCNuevoIdx]).toUpperCase().trim() : null;

        if (filterVal && filterVal !== 'TODOS') {
            if (cNuevoVal !== filterVal) continue;
        }
        if (!customer) continue;

        amount = parseAmount(amount);
        if (amount) {
            salesMap[customer] = (salesMap[customer] || 0) + amount;
            grandTotal += amount;
        }
    }

    const sortedData = Object.entries(salesMap)
        .map(([name, total]) => ({ name, total }))
        .sort((a, b) => b.total - a.total);

    updateKPIs(sortedData, grandTotal);
    renderCharts(sortedData);
    renderTable(sortedData, grandTotal);
}

function updateKPIs(data, total) {
    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });
    document.getElementById('kpi-total').innerText = formatter.format(total);
    document.getElementById('kpi-count').innerText = data.length;
    document.getElementById('kpi-top-client').innerText = data.length > 0 ? data[0].name : '-';
}

function renderCharts(data) {
    const top10 = data.slice(0, 10);
    const ctxMain = document.getElementById('mainChart').getContext('2d');
    const ctxPie = document.getElementById('pieChart').getContext('2d');

    if (mainChartInstance) mainChartInstance.destroy();
    if (pieChartInstance) pieChartInstance.destroy();

    mainChartInstance = new Chart(ctxMain, {
        type: 'bar',
        data: {
            labels: top10.map(d => d.name.substring(0, 15) + '...'),
            datasets: [{
                label: 'Ventas (S/)',
                data: top10.map(d => d.total),
                backgroundColor: '#3b82f6',
                borderRadius: 6
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { display: false }, tooltip: {
                    callbacks: { label: (c) => new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' }).format(c.raw) }
                }
            },
            scales: {
                y: { grid: { color: '#334155' }, ticks: { color: '#94a3b8' } },
                x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
            }
        }
    });

    const top5 = data.slice(0, 5);
    const othersTotal = data.slice(5).reduce((acc, curr) => acc + curr.total, 0);
    const pieData = [...top5, { name: 'Otros', total: othersTotal }];

    pieChartInstance = new Chart(ctxPie, {
        type: 'doughnut',
        data: {
            labels: pieData.map(d => d.name),
            datasets: [{
                data: pieData.map(d => d.total),
                backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#64748b'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false, cutout: '70%',
            plugins: { legend: { position: 'right', labels: { color: '#cbd5e1', font: { size: 10 }, boxWidth: 10 } } }
        }
    });
}

function renderTable(data, total) {
    const tbody = document.getElementById('table-body');
    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });

    tbody.innerHTML = data.map((d, index) => `
        <tr class="hover:bg-slate-800 transition-colors">
            <td class="px-6 py-4 font-medium text-slate-400">#${index + 1}</td>
            <td class="px-6 py-4 text-white">${d.name}</td>
            <td class="px-6 py-4 text-right font-medium text-emerald-400">${formatter.format(d.total)}</td>
            <td class="px-6 py-4 text-right text-slate-500">${total > 0 ? ((d.total / total) * 100).toFixed(2) : 0}%</td>
        </tr>
    `).join('');

    document.getElementById('search-table').addEventListener('keyup', (e) => {
        const term = e.target.value.toLowerCase();
        tbody.querySelectorAll('tr').forEach(r => {
            r.style.display = r.innerText.toLowerCase().includes(term) ? '' : 'none';
        });
    });
}


// --- 1. ANNUAL VIEW FUNCTIONALITY ---
btnAnnual.addEventListener('click', () => {
    columnMapper.classList.add('hidden');
    annualView.classList.remove('hidden');
    processAnnualData();
});
backFromAnnualSum.addEventListener('click', () => {
    annualView.classList.add('hidden');
    columnMapper.classList.remove('hidden');
});

function processAnnualData() {
    if (!workbook) return;
    const sheetName = document.getElementById('sheet-trend').value;
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const clientTotals = {};
    const monthlyTotals = new Array(12).fill(0);
    let grandAnnualTotal = 0;
    const filterVal = document.getElementById('new-client-filter').value;
    globalMonthCols.sort((a, b) => a.order - b.order);

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const cust = row[globalClientColIdx];
        if (!cust) continue;

        let cNuevoVal = globalCNuevoIdx !== -1 ? String(row[globalCNuevoIdx]).toUpperCase().trim() : null;
        if (filterVal && filterVal !== 'TODOS' && cNuevoVal !== filterVal) continue;

        let clientSum = 0;
        let activeMonths = 0;

        globalMonthCols.forEach((mCol) => {
            const val = parseAmount(row[mCol.index]);
            if (val > 0) {
                clientSum += val;
                monthlyTotals[mCol.order] += val;
                grandAnnualTotal += val;
                activeMonths++;
            }
        });

        if (clientSum > 0) {
            clientTotals[cust] = { name: cust, total: clientSum, activeMonths };
        }
    }
    const ranking = Object.values(clientTotals).sort((a, b) => b.total - a.total);
    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });
    document.getElementById('annual-total-kpi').innerText = formatter.format(grandAnnualTotal);
    document.getElementById('annual-top-client').innerText = ranking.length > 0 ? ranking[0].name : '-';
    document.getElementById('annual-table-body').innerHTML = ranking.map((d, index) => `
        <tr class="hover:bg-slate-800 transition-colors">
            <td class="px-6 py-4 font-medium text-slate-400">#${index + 1}</td>
            <td class="px-6 py-4 text-white">${d.name}</td>
            <td class="px-6 py-4 text-right font-medium text-emerald-400">${formatter.format(d.total)}</td>
            <td class="px-6 py-4 text-center text-slate-400">${d.activeMonths} / 12</td>
        </tr>
    `).join('');

    const ctx = document.getElementById('annualTrendChart').getContext('2d');
    if (annualChartInstance) annualChartInstance.destroy();
    annualChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
            datasets: [{ label: 'Venta', data: monthlyTotals, borderColor: '#10b981', backgroundColor: 'rgba(16, 185, 129, 0.1)', fill: true, tension: 0.4 }]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { grid: { color: '#334155' } }, x: { grid: { display: false } } } }
    });
}


// --- 2. HISTORY FUNCTIONALITY ---
btnHistory.addEventListener('click', () => {
    historyModal.classList.remove('hidden');
    populateClientDatalist();
});
closeHistoryBtn.addEventListener('click', () => { historyModal.classList.add('hidden'); });

function populateClientDatalist() {
    const sheetName = document.getElementById('sheet-trend').value;
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    const datalist = document.getElementById('client-datalist');
    datalist.innerHTML = '';
    const clients = new Set();
    for (let i = 1; i < data.length; i++) {
        if (data[i][globalClientColIdx]) clients.add(data[i][globalClientColIdx]);
    }
    clients.forEach(c => {
        const opt = document.createElement('option');
        opt.value = c;
        datalist.appendChild(opt);
    });
}

searchActionBtn.addEventListener('click', () => {
    const term = document.getElementById('client-search-input').value;
    if (!term) return;
    processHistoryData(term);
});

function processHistoryData(clientName) {
    const sheetName = document.getElementById('sheet-trend').value;
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    // Find client row
    let clientRow = null;
    for (let i = 1; i < data.length; i++) {
        const name = String(data[i][globalClientColIdx]).trim().toLowerCase();
        if (name === clientName.toLowerCase().trim()) {
            clientRow = data[i];
            break;
        }
    }

    if (!clientRow) return alert('Cliente no encontrado');

    document.getElementById('history-content').classList.remove('hidden');
    document.getElementById('hist-client-name').innerText = clientName;

    let total = 0;
    const tbody = document.getElementById('hist-table-body');
    let html = '';
    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });

    globalMonthCols.sort((a, b) => a.order - b.order).forEach(m => {
        const val = parseAmount(clientRow[m.index]);
        total += val;
        const status = val > 0 ? '<span class="text-emerald-400">Activo</span>' : '<span class="text-red-400">Sin Compra</span>';
        html += `
            <tr class="border-b border-slate-700">
                <td class="px-6 py-3 text-white">${m.label}</td>
                <td class="px-6 py-3 text-right font-mono ${val > 0 ? 'text-emerald-300' : 'text-slate-500'}">${formatter.format(val)}</td>
                <td class="px-6 py-3 text-right text-xs font-bold uppercase">${status}</td>
            </tr>
        `;
    });

    document.getElementById('hist-client-total').innerText = formatter.format(total);
    tbody.innerHTML = html;
}


// --- 3. COMPARISON FUNCTIONALITY (Month vs Month) ---
btnCompare.addEventListener('click', () => {
    compareModal.classList.remove('hidden');
    // Populate Selects
    const s1 = document.getElementById('comp-month-a');
    const s2 = document.getElementById('comp-month-b');
    s1.innerHTML = ''; s2.innerHTML = '';

    globalMonthCols.forEach(m => {
        s1.add(new Option(m.label, m.index));
        s2.add(new Option(m.label, m.index));
    });
});
closeCompModalBtn.addEventListener('click', () => compareModal.classList.add('hidden'));

runComparisonBtn.addEventListener('click', () => {
    compareModal.classList.add('hidden');
    columnMapper.classList.add('hidden');
    comparisonView.classList.remove('hidden');
    processComparisonData();
});

backFromComp.addEventListener('click', () => {
    comparisonView.classList.add('hidden');
    columnMapper.classList.remove('hidden');
});


function processComparisonData() {
    const idxA = parseInt(document.getElementById('comp-month-a').value);
    const idxB = parseInt(document.getElementById('comp-month-b').value);
    const sheetName = document.getElementById('sheet-trend').value;
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    // Labels
    const selectA = document.getElementById('comp-month-a');
    const labelA = selectA.options[selectA.selectedIndex].text;
    const selectB = document.getElementById('comp-month-b');
    const labelB = selectB.options[selectB.selectedIndex].text;

    document.getElementById('comp-label-a').innerText = `Venta ${labelA}`;
    document.getElementById('comp-label-b').innerText = `Venta ${labelB}`;

    let totalA = 0, totalB = 0;
    const clients = [];
    const churnList = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const cust = row[globalClientColIdx];
        if (!cust) continue;

        const valA = parseAmount(row[idxA]);
        const valB = parseAmount(row[idxB]);

        totalA += valA;
        totalB += valB;

        if (valA > 0 || valB > 0) {
            clients.push({ name: cust, a: valA, b: valB, diff: valB - valA });
        }

        // Churn logic: Bought in A > 0, Bought in B == 0
        if (valA > 0 && valB === 0) {
            churnList.push({ name: cust, lost: valA });
        }
    }

    // Render Summary
    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });
    document.getElementById('comp-total-a').innerText = formatter.format(totalA);
    document.getElementById('comp-total-b').innerText = formatter.format(totalB);

    const diff = totalB - totalA;
    const perc = totalA > 0 ? (diff / totalA) * 100 : 0;
    document.getElementById('comp-diff-val').innerText = (diff >= 0 ? '+' : '') + formatter.format(diff);
    const badge = document.getElementById('comp-diff-perc');
    badge.innerText = perc.toFixed(2) + '%';
    badge.className = `text-sm font-medium mb-1 px-2 py-0.5 rounded ${perc >= 0 ? 'bg-emerald-900 text-emerald-300' : 'bg-red-900 text-red-300'}`;

    // Render Chart
    clients.sort((a, b) => b.b - a.b); // Sort by current month sales
    const top10 = clients.slice(0, 10);
    const ctx = document.getElementById('compChart').getContext('2d');
    if (compChartInstance) compChartInstance.destroy();

    compChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: top10.map(c => c.name.substring(0, 10) + '...'),
            datasets: [
                { label: labelA, data: top10.map(c => c.a), backgroundColor: '#64748b' },
                { label: labelB, data: top10.map(c => c.b), backgroundColor: '#3b82f6' }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // Render Churn Table
    churnList.sort((a, b) => b.lost - a.lost);
    const churnBody = document.getElementById('churn-table-body');
    churnBody.innerHTML = churnList.map(c => `
        <tr class="border-b border-slate-700/50 hover:bg-red-900/10">
            <td class="px-6 py-3 text-slate-300">${c.name}</td>
            <td class="px-6 py-3 text-right text-slate-400">${formatter.format(c.lost)}</td>
            <td class="px-6 py-3 text-right text-red-400 font-bold">S/ 0.00</td>
            <td class="px-6 py-3 text-right text-red-500">-${formatter.format(c.lost)}</td>
        </tr>
    `).join('');
}


// --- 4. NO SALES BUTTON ---
btnNoSales.addEventListener('click', () => { noSalesModal.classList.remove('hidden'); populateNoSalesTable(); });
closeNoSalesBtn.addEventListener('click', () => noSalesModal.classList.add('hidden'));

function populateNoSalesTable() {
    if (!workbook) return;
    const sheetName = document.getElementById('sheet-trend').value;
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    const tbody = document.getElementById('no-sales-body');
    let rows = '';
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const cNuevo = row[globalCNuevoIdx] ? String(row[globalCNuevoIdx]).toUpperCase() : '';
        const vendedor = globalVendedorIdx !== -1 ? (row[globalVendedorIdx] || 'Sin Asignar') : '-';
        const cliente = row[globalClientColIdx];
        if (!cliente) continue;
        if (cNuevo.includes('SIN VENTA')) {
            rows += `
                <tr class="hover:bg-slate-800/50 transition-colors border-b border-slate-800 last:border-0">
                    <td class="px-6 py-3 font-medium text-white">${cliente}</td>
                    <td class="px-6 py-3 text-slate-300 font-bold tracking-wide">${vendedor}</td>
                    <td class="px-6 py-3 text-right"><span class="px-2 py-1 rounded bg-red-900/50 text-red-300 text-xs border border-red-900">SIN VENTA</span></td>
                </tr>
            `;
        }
    }
    if (!rows) rows = '<tr><td colspan="3" class="px-6 py-8 text-center text-slate-500">No se encontraron clientes "SIN VENTA".</td></tr>';
    tbody.innerHTML = rows;
}


// --- 5. NEW CLIENTS BUTTON (UPDATED) ---
btnNewClients.addEventListener('click', () => { newClientsModal.classList.remove('hidden'); prepareNewClientsData(); showSellersGrid(); });
closeNewClientsBtn.addEventListener('click', () => newClientsModal.classList.add('hidden'));
backSellersBtn.addEventListener('click', () => showSellersGrid());

function prepareNewClientsData() {
    if (!workbook) return;
    const sheetName = document.getElementById('sheet-trend').value;
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    const monthNames = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    newClientsDataCache = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const cNuevo = row[globalCNuevoIdx] ? String(row[globalCNuevoIdx]).toLowerCase().trim() : '';
        const isMonth = monthNames.includes(cNuevo);
        const isExcluded = cNuevo.includes('sin venta') || cNuevo === 'todos' || cNuevo === '';

        // Exclude strictly SIN VENTA here if we want clean data? 
        // User asked to hide SIN VENTA from filter. 
        // For report, better to stick to months.
        if (isMonth) {
            const vendedor = globalVendedorIdx !== -1 ? (row[globalVendedorIdx] || 'SIN ASIGNAR') : 'OFICINA';
            const cliente = row[globalClientColIdx];
            if (!cliente) continue;
            newClientsDataCache.push({ vendedor: String(vendedor).toUpperCase(), cliente: cliente, startMonth: row[globalCNuevoIdx], row: row });
        }
    }
}

function showSellersGrid() {
    document.getElementById('nc-sellers-view').classList.remove('hidden');
    document.getElementById('nc-detail-view').classList.add('hidden');
    backSellersBtn.classList.add('hidden');
    const grid = document.getElementById('nc-sellers-grid');
    grid.innerHTML = '';
    const counts = {};
    newClientsDataCache.forEach(item => { counts[item.vendedor] = (counts[item.vendedor] || 0) + 1; });

    // UI Update: Removing "??" => "??" and making name bigger (text-xl)
    Object.keys(counts).sort().forEach(seller => {
        const btn = document.createElement('div');
        btn.className = 'glass-panel p-6 rounded-xl border border-slate-700 hover:border-yellow-500 cursor-pointer transition-all flex flex-col items-center justify-center gap-4 group';
        btn.innerHTML = `
            <div class="w-16 h-16 rounded-full bg-slate-800 flex items-center justify-center text-3xl group-hover:scale-110 transition-transform shadow-lg">??</div>
            <h4 class="font-bold text-white text-xl text-center tracking-wide">${seller}</h4>
            <span class="text-xs font-bold px-3 py-1 bg-yellow-900/40 text-yellow-300 rounded-full border border-yellow-800 shadow-sm">${counts[seller]} Clientes Nuevos</span>`;
        btn.addEventListener('click', () => showSellerDetails(seller));
        grid.appendChild(btn);
    });
    if (Object.keys(counts).length === 0) grid.innerHTML = '<p class="col-span-full text-center text-slate-500">No hay clientes nuevos registrados.</p>';
}

function showSellerDetails(sellerName) {
    document.getElementById('nc-sellers-view').classList.add('hidden');
    document.getElementById('nc-detail-view').classList.remove('hidden');
    backSellersBtn.classList.remove('hidden');
    document.getElementById('nc-seller-name').innerText = sellerName;

    const clients = newClientsDataCache.filter(c => c.vendedor === sellerName);

    // Sorting by Month Order (Jan -> Dec)
    const monthOrder = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

    clients.sort((a, b) => {
        const mA = String(a.startMonth).toLowerCase().trim();
        const mB = String(b.startMonth).toLowerCase().trim();
        return monthOrder.indexOf(mA) - monthOrder.indexOf(mB);
    });

    const tbody = document.getElementById('nc-table-body');
    const formatter = new Intl.NumberFormat('es-PE', { style: 'decimal', minimumFractionDigits: 2 });

    // Header Reuse
    let headerHtml = `<th class="px-4 py-3 bg-slate-800 text-left">Cliente</th><th class="px-4 py-3 bg-slate-800 text-center">Mes Inicio</th>`;
    globalMonthCols.sort((a, b) => a.order - b.order).forEach(m => { headerHtml += `<th class="px-2 py-3 bg-slate-800 text-right min-w-[80px]">${m.label.substring(0, 3)}</th>`; });
    headerHtml += `<th class="px-4 py-3 bg-slate-800 text-right">Total Acum.</th>`;
    document.getElementById('nc-table-header').innerHTML = headerHtml;

    tbody.innerHTML = clients.map(c => {
        let rowHtml = `<td class="px-4 py-3 font-medium text-white border-b border-slate-700">${c.cliente}</td>`;

        // Month badge color
        const mName = String(c.startMonth).toUpperCase();
        rowHtml += `<td class="px-4 py-3 text-center border-b border-slate-700">
            <span class="px-2 py-1 rounded-md bg-yellow-900/20 text-yellow-400 font-bold text-xs border border-yellow-900/50">${mName}</span>
        </td>`;

        let clientTotal = 0;
        globalMonthCols.forEach(m => {
            const val = parseAmount(c.row[m.index]);
            clientTotal += val;
            const cellStyle = val > 0 ? 'text-emerald-400 font-mono' : 'text-slate-600';
            rowHtml += `<td class="px-2 py-3 text-right ${cellStyle} border-b border-slate-700">${val > 0 ? formatter.format(val) : '-'}</td>`;
        });
        rowHtml += `<td class="px-4 py-3 text-right text-white font-bold border-b border-slate-700 bg-slate-800/50">${formatter.format(clientTotal)}</td>`;
        return `<tr>${rowHtml}</tr>`;
    }).join('');
}


// -------------------------------------------------------------
// ADVANCED REPORT 1: COMPARATIVE (Year vs Year)
// -------------------------------------------------------------
btnOpenCompReport.addEventListener('click', () => {
    const sheetName = document.getElementById('sheet-comp').value;
    if (!sheetName) return alert("Selecciona una hoja comparativa.");
    modalCompReport.classList.remove('hidden');
    processComparativeReport(sheetName);
});

document.getElementById('close-comp-report-btn').addEventListener('click', () => modalCompReport.classList.add('hidden'));

function processComparativeReport(sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const yearsFound = [];
    const yearRegex = /20\d{2}/;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const label = String(row[0]).trim();
        if (yearRegex.test(label) && row.length > 2) {
            const amounts = [];
            for (let m = 1; m <= 12; m++) {
                amounts.push(parseAmount(row[m]));
            }
            yearsFound.push({ year: label, data: amounts, total: amounts.reduce((a, b) => a + b, 0) });
        }
    }

    if (yearsFound.length < 2) {
        alert("No se encontraron suficientes datos de años (2024, 2025) en la hoja.");
        return;
    }

    yearsFound.sort((a, b) => parseInt(a.year) - parseInt(b.year));

    const prev = yearsFound[yearsFound.length - 2];
    const curr = yearsFound[yearsFound.length - 1];

    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });
    document.getElementById('cr-total-prev').innerText = formatter.format(prev.total);
    document.getElementById('cr-total-curr').innerText = formatter.format(curr.total);

    const growth = prev.total > 0 ? ((curr.total - prev.total) / prev.total) * 100 : 0;
    const growthEl = document.getElementById('cr-growth');
    growthEl.innerText = growth.toFixed(2) + '%';
    growthEl.className = growth >= 0 ? 'text-2xl font-bold text-emerald-400 mt-1' : 'text-2xl font-bold text-red-400 mt-1';

    const ctx = document.getElementById('compReportChart').getContext('2d');
    if (compReportChartInstance) compReportChartInstance.destroy();

    const monthLabels = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

    compReportChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: monthLabels,
            datasets: [
                {
                    label: prev.year,
                    data: prev.data,
                    borderColor: '#94a3b8',
                    borderDash: [5, 5],
                    tension: 0.4
                },
                {
                    label: curr.year,
                    data: curr.data,
                    borderColor: '#2dd4bf',
                    backgroundColor: 'rgba(45, 212, 191, 0.2)',
                    fill: true,
                    tension: 0.4
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { text: { color: '#fff' } },
            scales: { y: { grid: { color: '#334155' } }, x: { grid: { display: false } } }
        }
    });

    const tbody = document.getElementById('cr-table-body');
    let html = '';
    for (let i = 0; i < 12; i++) {
        const valP = prev.data[i];
        const valC = curr.data[i];
        const diff = valC - valP;
        const color = diff >= 0 ? 'text-emerald-400' : 'text-red-400';

        html += `
            <tr class="border-b border-slate-700 hover:bg-slate-800/50">
                <td class="px-6 py-3 text-white">${monthLabels[i]}</td>
                <td class="px-6 py-3 text-right text-slate-400">${formatter.format(valP)}</td>
                <td class="px-6 py-3 text-right text-white font-bold">${formatter.format(valC)}</td>
                <td class="px-6 py-3 text-right ${color}">${(diff >= 0 ? '+' : '') + formatter.format(diff)}</td>
            </tr>
        `;
    }
    tbody.innerHTML = html;
}

// -------------------------------------------------------------
// ADVANCED REPORT 2: SELLER MANAGEMENT (Meta vs Real)
// -------------------------------------------------------------
let sellerDataCache = [];

btnOpenSellerReport.addEventListener('click', () => {
    const sheetName = document.getElementById('sheet-seller').value;
    if (!sheetName) return alert("Selecciona una hoja de Vendedores.");
    modalSellerReport.classList.remove('hidden');
    processSellerReport(sheetName);
});

document.getElementById('close-seller-report-btn').addEventListener('click', () => modalSellerReport.classList.add('hidden'));
document.getElementById('back-seller-list-btn').addEventListener('click', () => {
    document.getElementById('sr-list-view').classList.remove('hidden');
    document.getElementById('sr-detail-view').classList.add('hidden');
    document.getElementById('back-seller-list-btn').classList.add('hidden');
});

// UPDATED LOGIC (META COLUMN)
function processSellerReport(sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // 1. Find Header Row (Looking for "VENDEDOR" and "META" as columns)
    let headerRowIdx = -1;
    let idxName = -1, idxMeta = -1, idxJan = -1;

    for (let i = 0; i < Math.min(data.length, 20); i++) {
        const row = data[i];
        if (!row) continue;
        const lowerRow = row.map(c => String(c).toLowerCase().trim());

        if (lowerRow.includes('vendedor') && lowerRow.includes('meta')) {
            headerRowIdx = i;
            idxName = lowerRow.indexOf('vendedor');
            idxMeta = lowerRow.indexOf('meta');
            idxJan = lowerRow.findIndex(c => c.includes('enero')); // Find start of months
            break;
        }
    }

    if (headerRowIdx === -1 || idxName === -1 || idxMeta === -1) {
        alert("Error: No se encontró la estructura correcta (Columnas requeridas: VENDEDOR, META, ENERO...). Verifique su Excel.");
        return;
    }

    // 2. Parse Sellers
    sellerDataCache = [];

    // Iterate rows after header
    for (let i = headerRowIdx + 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;

        const name = row[idxName];
        if (!name) continue; // Skip empty rows

        const monthlyMetaVal = parseAmount(row[idxMeta]);

        const sales = [];
        const metas = [];

        // Determine where months start (Default to column after Meta if 'Enero' not explicitly found)
        let currentMonthIdx = idxJan !== -1 ? idxJan : (idxMeta + 1);

        for (let m = 0; m < 12; m++) {
            sales.push(parseAmount(row[currentMonthIdx + m]));
            metas.push(monthlyMetaVal); // Monthly target
        }

        sellerDataCache.push({
            name: String(name).toUpperCase(),
            sales: sales,
            meta: metas,
            totalSale: sales.reduce((a, b) => a + b, 0),
            totalMeta: metas.reduce((a, b) => a + b, 0) // Annual Meta
        });
    }

    if (sellerDataCache.length === 0) {
        alert("No se encontraron vendedores en la lista.");
        return;
    }

    // Render Grid
    const grid = document.getElementById('sr-grid');
    grid.innerHTML = '';

    sellerDataCache.forEach(s => {
        const div = document.createElement('div');
        div.className = 'glass-panel p-6 rounded-xl border border-slate-700 hover:border-orange-500 cursor-pointer transition-all flex flex-col items-center gap-2 group';
        div.innerHTML = `
            <div class="w-16 h-16 rounded-full bg-slate-800 flex items-center justify-center text-3xl group-hover:scale-110 transition-transform">??</div>
            <h4 class="font-bold text-white text-2xl">${s.name}</h4>
            <p class="text-sm text-slate-400">Venta Total: S/ ${s.totalSale.toLocaleString('es-PE')}</p>
        `;
        div.addEventListener('click', () => showSellerDetail(s));
        grid.appendChild(div);
    });

    // Show View 1
    document.getElementById('sr-list-view').classList.remove('hidden');
    document.getElementById('sr-detail-view').classList.add('hidden');
    document.getElementById('back-seller-list-btn').classList.add('hidden');
}

function showSellerDetail(seller) {
    document.getElementById('sr-list-view').classList.add('hidden');
    document.getElementById('sr-detail-view').classList.remove('hidden');
    document.getElementById('back-seller-list-btn').classList.remove('hidden');

    document.getElementById('sr-seller-name').innerText = seller.name;

    // KPIs
    const formatter = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' });
    document.getElementById('sr-meta-total').innerText = formatter.format(seller.totalMeta);
    document.getElementById('sr-sale-total').innerText = formatter.format(seller.totalSale);

    const compliance = seller.totalMeta > 0 ? (seller.totalSale / seller.totalMeta) * 100 : 0;
    document.getElementById('sr-compliance').innerText = compliance.toFixed(2) + '%';
    document.getElementById('sr-compliance').className = compliance >= 100 ? 'text-2xl font-bold text-emerald-400 mt-1' : 'text-2xl font-bold text-yellow-400 mt-1';

    // Chart
    const ctx = document.getElementById('sellerChart').getContext('2d');
    if (sellerChartInstance) sellerChartInstance.destroy();

    const monthLabels = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

    sellerChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: monthLabels,
            datasets: [
                {
                    type: 'line',
                    label: 'Meta Mensual',
                    data: seller.meta,
                    borderColor: '#f97316', // Orange
                    backgroundColor: 'rgba(249, 115, 22, 0.1)',
                    borderWidth: 2,
                    borderDash: [5, 5],
                    pointStyle: 'circle',
                    tension: 0,
                    fill: false // Don't fill area
                },
                {
                    type: 'bar',
                    label: 'Venta Real',
                    data: seller.sales,
                    backgroundColor: '#3b82f6', // Blue
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                y: { grid: { color: '#334155' }, ticks: { color: '#94a3b8' } },
                x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
            }
        }
    });

    // Table
    const tbody = document.getElementById('sr-table-body');
    let html = '';
    for (let i = 0; i < 12; i++) {
        const meta = seller.meta[i];
        const sale = seller.sales[i];
        const comp = meta > 0 ? (sale / meta) * 100 : 0;
        const color = comp >= 100 ? 'text-emerald-400' : 'text-yellow-400';

        html += `
             <tr class="border-b border-slate-700 hover:bg-slate-800/50">
                <td class="px-6 py-3 text-white">${monthLabels[i]}</td>
                <td class="px-6 py-3 text-right text-orange-300">${formatter.format(meta)}</td>
                <td class="px-6 py-3 text-right text-white font-bold">${formatter.format(sale)}</td>
                <td class="px-6 py-3 text-right ${color}">${comp.toFixed(1)}%</td>
            </tr>
        `;
    }
    tbody.innerHTML = html;
}

// -------------------------------------------------------------
// ADVANCED REPORT 3: PRODUCT ANALYSIS (Top 10)
// -------------------------------------------------------------
btnOpenProductReport.addEventListener('click', () => {
    const sheetName = document.getElementById('sheet-product').value;
    if (!sheetName) return alert("Selecciona una hoja de Productos.");
    modalProductReport.classList.remove('hidden');
    processProductReport(sheetName);
});

document.getElementById('close-product-report-btn').addEventListener('click', () => modalProductReport.classList.add('hidden'));

function processProductReport(sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const products = [];
    // Identify structure: Col 0 = Name, Cols 1... = Months
    // Skip header
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const name = row[0];
        if (!name) continue;

        let total = 0;
        // Sum roughly 1-12
        for (let m = 1; m <= 12; m++) {
            total += parseAmount(row[m]);
        }
        products.push({ name: name, total: total });
    }

    // Sort
    products.sort((a, b) => b.total - a.total);

    // Top 10 Chart
    const top10 = products.slice(0, 10);
    const ctx = document.getElementById('productChart').getContext('2d');
    if (productChartInstance) productChartInstance.destroy();

    productChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: top10.map(p => p.name.substring(0, 15) + '...'),
            datasets: [{
                label: 'Unidades / Venta',
                data: top10.map(p => p.total),
                backgroundColor: '#ec4899', // Pink
                borderRadius: 4
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            indexAxis: 'y', // Horizontal bars
            scales: { x: { grid: { color: '#334155' } }, y: { grid: { display: false } } }
        }
    });

    // Table
    const tbody = document.getElementById('pr-table-body');
    const formatter = new Intl.NumberFormat('es-PE', { style: 'decimal' });

    tbody.innerHTML = products.map((p, index) => `
        <tr class="border-b border-slate-700 hover:bg-slate-800/50">
            <td class="px-6 py-3 text-slate-400 font-mono">#${index + 1}</td>
            <td class="px-6 py-3 text-white font-medium">${p.name}</td>
            <td class="px-6 py-3 text-right text-pink-300">${formatter.format(p.total)}</td>
        </tr>
    `).join('');
}

// Utils
function parseAmount(val) {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    if (typeof val === 'string') return parseFloat(val.replace(/[^0-9.-]+/g, "")) || 0;
    return 0;
}
