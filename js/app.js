/* ============================================================
   InvDash app.js  v2.0 — redesigned UI
   ============================================================ */

/* ── PWA / Install ─────────────────────────────────────── */
let deferredPrompt = null;

const isSecureContextForPWA =
  location.protocol === 'http:' ||
  location.protocol === 'https:' ||
  location.hostname === 'localhost' ||
  location.hostname === '127.0.0.1';

if ('serviceWorker' in navigator && isSecureContextForPWA) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js')
      .catch(err => console.log('SW register skipped:', err));
  });
}

window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault();
  deferredPrompt = e;
  const b = document.getElementById('installBtn');
  if (b) b.style.display = 'flex';
});

document.addEventListener('click', async e => {
  if (e.target && e.target.closest('#installBtn') && deferredPrompt) {
    deferredPrompt.prompt();
    await deferredPrompt.userChoice;
    deferredPrompt = null;
    const b = document.getElementById('installBtn');
    if (b) b.style.display = 'none';
  }
});

/* ── State ─────────────────────────────────────────────── */
let inventoryData = [];
let stocktake = {};
let settings = { lowStockThreshold: 10, sessionName: '' };
let db;
let activeTab = 'dashboard';
const pageSize = 50;
const paginationState = { preview: 1, alerts: 1, items: 1, stocktake: 1 };

const $ = id => document.getElementById(id);

const tableBodies = {
  preview:         () => document.querySelector('#previewTable tbody'),
  dashboardAlerts: () => document.querySelector('#dashboardAlertsTable tbody'),
  alerts:          () => document.querySelector('#alertsTable tbody'),
  items:           () => document.querySelector('#itemsTable tbody'),
  stocktake:       () => document.querySelector('#stocktakeTable tbody')
};

/* ── Helpers ────────────────────────────────────────────── */
function toNumberOrNull(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'string') {
    const t = v.trim();
    if (!t || t.toUpperCase() === '#N/A') return null;
    const n = Number(t.replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function cleanText(v) {
  if (v === null || v === undefined) return '';
  const t = String(v).trim();
  return t.toUpperCase() === '#N/A' ? '' : t;
}

function formatNumber(n) { return Number(n || 0).toLocaleString(); }

function getStatus(item) {
  const qty = Number(item.QtyOnHand || 0);
  const min = item.MinLevel;
  if (qty <= 0) return 'STOCKOUT';
  if (min !== null && min !== undefined && min !== '' && qty < Number(min)) return 'LOW STOCK';
  if (qty < Number(settings.lowStockThreshold || 0)) return 'REVIEW';
  return 'OK';
}

function statusBadge(s) {
  if (s === 'STOCKOUT')  return '<span class="pill pill-danger">STOCKOUT</span>';
  if (s === 'LOW STOCK') return '<span class="pill pill-warning">LOW STOCK</span>';
  if (s === 'REVIEW')    return '<span class="pill pill-warning">REVIEW</span>';
  return '<span class="pill pill-ok">OK</span>';
}

function rowClass(s) {
  if (s === 'STOCKOUT')              return 'stockout';
  if (s === 'LOW STOCK' || s === 'REVIEW') return 'low-stock';
  return 'ok-row';
}

function normalizeItem(raw) {
  const stock = raw.StockCode ?? raw['Stock Code'] ?? raw['SAP Material'] ?? raw.Material;
  const desc = raw.Description ?? raw['Material Description'] ?? raw['Item Description'];
  if (!stock || !desc) return null;
  const onOrderHeader = Object.keys(raw).find(header => /^on[\s_]*order$/i.test(header.trim()));
  return {
    StockCode:    cleanText(stock),
    Description:  cleanText(desc),
    AlternateKey: cleanText(raw.AlternateKey ?? raw['old stkcode'] ?? raw['old stk code']),
    QtyOnHand:    toNumberOrNull(raw.QtyOnHand ?? raw['Stock on hand '] ?? raw['Stock on hand'] ?? raw.Qty ?? raw['Plant stock']) ?? 0,
    OnOrder:      toNumberOrNull(onOrderHeader ? raw[onOrderHeader] : null),
    DefaultBin:   cleanText(raw.DefaultBin ?? raw['Bin location'] ?? raw['Default Bin'] ?? raw.Bin ?? raw.BIN),
    MinLevel:     toNumberOrNull(raw.MinLevel ?? raw['minimum QTY'] ?? raw.Min ?? raw.min),
    MaxLevel:     toNumberOrNull(raw.MaxLevel ?? raw['maximum QTY'] ?? raw.Max ?? raw.max),
    UnitCost:     toNumberOrNull(raw.UnitCost ?? raw['Unit cost '] ?? raw['Unit cost'] ?? raw['unit Price']),
    TotalValue:   toNumberOrNull(raw.TotalValue ?? raw['Total value'] ?? raw['Value Unrestricted']),
    Department:   cleanText(raw.Department ?? raw['Department '] ?? raw.department),
    Machine:      cleanText(raw.Machine ?? raw['Machine ']),
    Currency:     cleanText(raw.Currency)
  };
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, m =>
    ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[m]));
}
function escapeAttr(str) { return escapeHtml(str); }
function cssEscape(str) { return String(str).replace(/"/g, '\\"'); }

/* ── Excel Export ───────────────────────────────────────── */
function downloadExcel(rows, sheetName, filename, { columnWidths, autoFilter = false, prepareSheet } = {}) {
  if (typeof XLSX === 'undefined') {
    alert('Excel export library not loaded. Open through a local server or while online once.');
    return;
  }
  const ws = XLSX.utils.json_to_sheet(rows);
  prepareSheet?.(ws);
  if (columnWidths) ws['!cols'] = columnWidths.map(wch => ({ wch }));
  if (autoFilter) ws['!autofilter'] = { ref: ws['!ref'] };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename, { cellStyles: Boolean(columnWidths) });
}

/* ── IndexedDB ──────────────────────────────────────────── */
const request = indexedDB.open('InventoryPWA_DB_MOBILE', 1);
request.onerror = e => console.log('IndexedDB error:', e.target.errorCode);
request.onupgradeneeded = e => {
  db = e.target.result;
  if (!db.objectStoreNames.contains('inventory')) db.createObjectStore('inventory', { keyPath: 'StockCode' });
  if (!db.objectStoreNames.contains('stocktake')) db.createObjectStore('stocktake', { keyPath: 'StockCode' });
  if (!db.objectStoreNames.contains('settings'))  db.createObjectStore('settings',  { keyPath: 'key' });
};
request.onsuccess = e => { db = e.target.result; loadAll(); };

function tx(name, mode = 'readonly') { return db.transaction([name], mode).objectStore(name); }

function saveInventory() {
  const s = tx('inventory', 'readwrite');
  s.clear();
  inventoryData.forEach(it => s.put(it));
}
function saveStocktake() {
  const s = tx('stocktake', 'readwrite');
  s.clear();
  Object.entries(stocktake).forEach(([StockCode, obj]) => s.put({ StockCode, ...obj }));
}
function saveSettings() {
  const s = tx('settings', 'readwrite');
  s.put({ key: 'lowStockThreshold', value: settings.lowStockThreshold });
  s.put({ key: 'sessionName', value: settings.sessionName });
}

function loadAll() {
  tx('inventory').getAll().onsuccess = e => {
    inventoryData = e.target.result || [];
    tx('stocktake').getAll().onsuccess = e2 => {
      stocktake = {};
      (e2.target.result || []).forEach(r => {
        stocktake[r.StockCode] = { CountedQty: r.CountedQty ?? null, UpdatedAt: r.UpdatedAt ?? null };
      });
      tx('settings').getAll().onsuccess = e3 => {
        (e3.target.result || []).forEach(r => { settings[r.key] = r.value; });
        if ($('lowStockThreshold')) $('lowStockThreshold').value = settings.lowStockThreshold ?? 10;
        if ($('sessionName'))       $('sessionName').value       = settings.sessionName ?? '';
        renderVisibleTab();
        updateDashboard();
      };
    };
  };
}

/* ── Navigation ─────────────────────────────────────────── */
function setActiveTab(tab) {
  activeTab = tab;

  document.querySelectorAll('.nav-item[data-tab], .mob-nav-btn[data-tab]').forEach(button => {
    const active = button.dataset.tab === tab;
    button.classList.toggle('active', active);
    if (active) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  });
  $('mainContent').scrollTop = 0;

  // Sections
  document.querySelectorAll('.tab-section').forEach(sec => {
    sec.style.display = sec.id === tab ? 'block' : 'none';
  });

  renderVisibleTab();
  closeSidebar();
}

function renderVisibleTab() {
  if (activeTab === 'upload')     renderPreviewTable();
  else if (activeTab === 'dashboard') renderDashboardAlertsTable();
  else if (activeTab === 'alerts')    renderAlertsTable();
  else if (activeTab === 'items')     renderItemsTable();
  else if (activeTab === 'stocktake') renderStocktakeTable();
  updateDashboard();
}

/* ── Mobile Sidebar ─────────────────────────────────────── */
function openSidebar() {
  $('sidebar').classList.add('open');
  $('sidebarOverlay').classList.add('show');
  $('menuToggle').setAttribute('aria-expanded', 'true');
  document.querySelector('.sidebar-nav .active')?.focus();
}
function closeSidebar() {
  if ($('sidebar').classList.contains('open')) $('menuToggle').focus();
  $('sidebar').classList.remove('open');
  $('sidebarOverlay').classList.remove('show');
  $('menuToggle').setAttribute('aria-expanded', 'false');
}

/* ── Debounce ───────────────────────────────────────────── */
function debounce(fn, delay = 180) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), delay); };
}

/* ── Render Helpers ─────────────────────────────────────── */
function matchesQuery(it, q) {
  if (!q) return true;
  return (
    (it.StockCode    || '').toLowerCase().includes(q) ||
    (it.Description  || '').toLowerCase().includes(q) ||
    (it.AlternateKey || '').toLowerCase().includes(q) ||
    (it.DefaultBin   || '').toLowerCase().includes(q) ||
    (it.Department   || '').toLowerCase().includes(q) ||
    (it.Machine      || '').toLowerCase().includes(q)
  );
}

function alertItems() {
  return inventoryData.filter(it => getStatus(it) !== 'OK');
}

function getPageItems(key, items) {
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));
  const currentPage = Math.min(Math.max(paginationState[key] || 1, 1), totalPages);
  paginationState[key] = currentPage;
  const start = (currentPage - 1) * pageSize;
  return { pageItems: items.slice(start, start + pageSize), currentPage, totalPages };
}

function ensurePager(tbody, key) {
  const table = tbody.closest('table');
  if (!table) return null;
  let pager = document.querySelector(`[data-pager="${key}"]`);
  if (!pager) {
    pager = document.createElement('div');
    pager.className = 'table-pager';
    pager.dataset.pager = key;
    table.insertAdjacentElement('afterend', pager);
  }
  return pager;
}

function updatePager(tbody, key, totalItems, currentPage, totalPages) {
  const pager = ensurePager(tbody, key);
  if (!pager) return;
  const start = totalItems === 0 ? 0 : ((currentPage - 1) * pageSize) + 1;
  const end = Math.min(currentPage * pageSize, totalItems);
  pager.innerHTML = `
    <span class="pager-info m-3" style="margin-left:10px; opacity: 0.8;">Showing ${formatNumber(start)}-${formatNumber(end)} of ${formatNumber(totalItems)}</span>
    <button type="button" class ="pager-btn" data-page-key="${key}" data-page-dir="prev" ${currentPage <= 1 ? 'disabled' : ''}>Prev</button>
    <span class="pager-page">Page ${formatNumber(currentPage)} of ${formatNumber(totalPages)}</span>
    <button type="button" style="margin-right:10px; opacity: 0.8;" class="pager-btn" data-page-key="${key}" data-page-dir="next" ${currentPage >= totalPages ? 'disabled' : ''}>Next</button>`;
}

function createRows(rowsHtmlTarget, items, rowBuilder, key) {
  rowsHtmlTarget.innerHTML = '';
  const { pageItems, currentPage, totalPages } = getPageItems(key, items);
  const frag = document.createDocumentFragment();
  pageItems.forEach(it => frag.appendChild(rowBuilder(it)));
  rowsHtmlTarget.appendChild(frag);
  updatePager(rowsHtmlTarget, key, items.length, currentPage, totalPages);
}

function createLimitedRows(rowsHtmlTarget, items, rowBuilder, limit) {
  rowsHtmlTarget.innerHTML = '';
  const frag = document.createDocumentFragment();
  items.slice(0, limit).forEach(it => frag.appendChild(rowBuilder(it)));
  rowsHtmlTarget.appendChild(frag);
}

function resetPage(key) {
  if (paginationState[key] !== undefined) paginationState[key] = 1;
}

function injectPaginationStyles() {
  if (document.getElementById('paginationStyles')) return;
  const style = document.createElement('style');
  style.id = 'paginationStyles';
  style.textContent = `
    .table-pager {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 10px;
      flex-wrap: wrap;
      margin: 12px 0 20px;
      font-size: 0.9rem;
    }
    .pager-info {
      margin-right: auto;
      opacity: 0.8;
    }
    .pager-btn {
      min-width: 72px;
      padding: 8px 12px;
      border-radius: 8px;
      cursor: pointer;
    }
    .pager-btn:disabled {
      cursor: not-allowed;
      opacity: 0.5;
    }
    @media (max-width: 640px) {
      .table-pager {
        justify-content: center;
      }
      .pager-info {
        flex-basis: 100%;
        margin-right: 0;
        text-align: center;
      }
    }`;
  document.head.appendChild(style);
}

/* ── Render Functions ───────────────────────────────────── */
function renderPreviewTable() {
  const tbody = tableBodies.preview();
  if (!tbody) return;
  const q = ($('searchInput')?.value || '').toLowerCase().trim();
  const items = inventoryData.filter(it => matchesQuery(it, q));
  createRows(tbody, items, it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td class="right mono">${it.MinLevel ?? ''}</td><td>${statusBadge(s)}</td>`;
    return tr;
  }, 'preview');
}

const healthLabels = { OK: 'Available', 'LOW STOCK': 'Low stock', STOCKOUT: 'Stockout', REVIEW: 'Review' };
const dashboardItemLimit = 10;
let dashboardHealthFilter = '';

function setDashboardHealthFilter(status) {
  dashboardHealthFilter = status;
  document.querySelectorAll('[data-health-filter]').forEach(button => {
    button.setAttribute('aria-pressed', String(button.dataset.healthFilter === status));
  });
  $('clearHealthFilter').hidden = !status;
  renderDashboardAlertsTable();
}

function renderDashboardAlertsTable() {
  const tbody = tableBodies.dashboardAlerts();
  if (!tbody) return;
  const order = { 'STOCKOUT': 0, 'LOW STOCK': 1, 'REVIEW': 2 };
  const items = dashboardHealthFilter
    ? inventoryData.filter(item => getStatus(item) === dashboardHealthFilter)
    : alertItems().sort((a, b) => order[getStatus(a)] - order[getStatus(b)]);
  const label = healthLabels[dashboardHealthFilter];
  $('dashboardTableTitle').textContent = label ? `${label} items` : 'Items to review';
  $('dashboardTableSummary').textContent = items.length
    ? `Showing ${formatNumber(Math.min(items.length, dashboardItemLimit))} of ${formatNumber(items.length)} items`
    : '0 matching items';
  createLimitedRows(tbody, items, it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="right mono">${it.MinLevel ?? ''}</td><td class="hide-mobile">${escapeHtml(it.Department || '')}</td><td>${statusBadge(s)}</td>`;
    return tr;
  }, dashboardItemLimit);

  const emptyEl = $('dashboardEmpty');
  if (emptyEl) emptyEl.style.display = items.length === 0 ? 'block' : 'none';
  $('dashboardEmptyTitle').textContent = !inventoryData.length ? 'No inventory imported'
    : label ? `No items with status ${label}` : 'No local stock alerts';
  $('dashboardEmptyMessage').textContent = inventoryData.length
    ? (label ? 'Choose another status to see more items.' : 'No imported items meet your alert rules.')
    : 'Import an Excel workbook to see local stock alerts.';
}

function renderAlertsTable() {
  const tbody = tableBodies.alerts();
  if (!tbody) return;
  const filter = $('alertFilter')?.value || '';
  const q = ($('alertSearch')?.value || '').toLowerCase().trim();
  const order = { 'STOCKOUT': 0, 'LOW STOCK': 1, 'REVIEW': 2 };
  const items = alertItems()
    .filter(it => matchesQuery(it, q))
    .filter(it => {
      const s = getStatus(it);
      if (filter === 'stockout') return s === 'STOCKOUT';
      if (filter === 'low')      return s === 'LOW STOCK';
      if (filter === 'review')   return s === 'REVIEW';
      return true;
    })
    .sort((a, b) => order[getStatus(a)] - order[getStatus(b)]);

  createRows(tbody, items, it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="right mono">${it.MinLevel ?? ''}</td><td class="hide-mobile">${escapeHtml(it.Department || '')}</td><td class="hide-mobile">${escapeHtml(it.Machine || '')}</td><td>${statusBadge(s)}</td>`;
    return tr;
  }, 'alerts');
}

function renderItemsTable() {
  const tbody = tableBodies.items();
  if (!tbody) return;
  const query = $('itemsSearch').value.trim().toLowerCase();
  const items = inventoryData.filter(item => matchesQuery(item, query))
    .sort((a, b) => String(a.StockCode).localeCompare(String(b.StockCode)));
  const empty = $('itemsEmpty');
  empty.hidden = items.length > 0;
  empty.textContent = query ? 'No items match your search.' : 'Import inventory to see items here.';
  createRows(tbody, items, it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td>${statusBadge(s)}</td><td class="center"><button data-del="${escapeAttr(it.StockCode)}">Del</button></td>`;
    return tr;
  }, 'items');
}

function renderStocktakeTable() {
  const tbody = tableBodies.stocktake();
  if (!tbody) return;
  const q = ($('stocktakeFilter')?.value || '').toLowerCase().trim();
  const items = [...inventoryData]
    .filter(it => matchesQuery(it, q))
    .sort((a, b) => String(a.StockCode).localeCompare(String(b.StockCode)));
  createRows(tbody, items, it => {
    const counted = stocktake[it.StockCode]?.CountedQty ?? null;
    const variance = counted === null ? '' : (counted - (it.QtyOnHand || 0));
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="right"><input inputmode="numeric" type="number" data-count="${escapeAttr(it.StockCode)}" value="${counted ?? ''}" /></td><td class="right mono" data-var="${escapeAttr(it.StockCode)}">${variance}</td><td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td>${statusBadge(s)}</td>`;
    return tr;
  }, 'stocktake');
}

/* ── Dashboard KPIs ─────────────────────────────────────── */
function updateDashboard() {
  $('totalItems').innerText    = formatNumber(inventoryData.length);
  $('totalQty').innerText      = formatNumber(inventoryData.reduce((s, it) => s + (it.QtyOnHand || 0), 0));
  const counts = { OK: 0, 'LOW STOCK': 0, STOCKOUT: 0, REVIEW: 0 };
  for (const item of inventoryData) counts[getStatus(item)]++;
  $('stockoutCount').innerText = formatNumber(counts.STOCKOUT);
  $('lowStockCount').innerText = formatNumber(counts['LOW STOCK']);
  const total = inventoryData.length;
  $('healthSummary').textContent = total
    ? `${formatNumber(total - counts.OK)} of ${formatNumber(total)} items need attention.`
    : 'Import inventory to see stock health.';
  document.querySelectorAll('[data-health-count]').forEach(label => {
    label.textContent = formatNumber(counts[label.dataset.healthCount]);
  });
  $('healthTotal').textContent = formatNumber(total);
  const ring = $('healthRing');
  const percentage = count => `${total ? (count / total) * 100 : 0}%`;
  ring.style.setProperty('--available-end', percentage(counts.OK));
  ring.style.setProperty('--low-end', percentage(counts.OK + counts['LOW STOCK']));
  ring.style.setProperty('--out-end', percentage(total - counts.REVIEW));
  ring.classList.toggle('has-inventory', total > 0);
  $('inventoryValue').innerText = formatNumber(inventoryData.reduce((s, it) =>
    s + (it.TotalValue || ((it.UnitCost || 0) * (it.QtyOnHand || 0))), 0));
  $('stockoutValue').innerText = formatNumber(inventoryData
    .filter(it => getStatus(it) === 'STOCKOUT')
    .reduce((s, it) => s + (it.UnitCost || 0), 0));
  $('countedLines').innerText  = formatNumber(Object.keys(stocktake).length);

  const byCode = new Map(inventoryData.map(it => [it.StockCode, it]));
  $('varianceLines').innerText = formatNumber(
    Object.keys(stocktake).filter(code => {
      const it = byCode.get(code);
      const counted = stocktake[code]?.CountedQty;
      return counted !== null && counted !== undefined && (counted - (it?.QtyOnHand || 0)) !== 0;
    }).length
  );

  // Alert badge
  const alertCount = total - counts.OK;
  const badge = $('alertsBadge');
  if (badge) {
    badge.textContent = alertCount;
    badge.style.display = alertCount > 0 ? 'inline-block' : 'none';
  }
}

/* ── Export Functions ───────────────────────────────────── */
function exportVarianceExcel() {
  if (inventoryData.length === 0) return alert('No inventory loaded.');
  const byCode = new Map(inventoryData.map(it => [it.StockCode, it]));
  const rows = [];
  Object.entries(stocktake).forEach(([code, obj]) => {
    const it = byCode.get(code);
    if (!it || obj.CountedQty === null || obj.CountedQty === undefined) return;
    const variance = obj.CountedQty - (it.QtyOnHand || 0);
    if (variance === 0) return;
    rows.push({
      Session: settings.sessionName || '',
      SAP_Material: it.StockCode,
      Material_Description: it.Description,
      Bin_location: it.DefaultBin || '',
      Department: it.Department || '',
      Machine: it.Machine || '',
      SystemQty: it.QtyOnHand || 0,
      CountedQty: obj.CountedQty,
      Variance: variance,
      Status: getStatus(it),
      UpdatedAt: obj.UpdatedAt || ''
    });
  });
  if (rows.length === 0) return alert('No variances found yet.');
  downloadExcel(rows, 'Variances', 'Stocktake_Variances.xlsx');
}

function exportFullCountExcel() {
  if (inventoryData.length === 0) return alert('No inventory loaded.');
  const rows = inventoryData.map(it => {
    const counted = stocktake[it.StockCode]?.CountedQty ?? null;
    return {
      Session: settings.sessionName || '',
      SAP_Material: it.StockCode,
      Material_Description: it.Description,
      Bin_location: it.DefaultBin || '',
      Department: it.Department || '',
      Machine: it.Machine || '',
      Stock_on_hand: it.QtyOnHand || 0,
      minimum_QTY: it.MinLevel ?? '',
      maximum_QTY: it.MaxLevel ?? '',
      Status: getStatus(it),
      CountedQty: counted,
      Variance: counted === null ? null : (counted - (it.QtyOnHand || 0))
    };
  });
  downloadExcel(rows, 'FullCount', 'Stocktake_FullCount.xlsx');
}

function exportCycleCountSheet() {
  if (inventoryData.length === 0) return alert('No inventory loaded.');
  const rows = inventoryData.map(it => ({
    SAP_Material: it.StockCode,
    Material_Description: it.Description,
    Bin_location: it.DefaultBin || '',
    Department: it.Department || '',
    Machine: it.Machine || '',
    Stock_on_hand: it.QtyOnHand || 0,
    CountedQty: '',
    minimum_QTY: it.MinLevel ?? '',
    maximum_QTY: it.MaxLevel ?? '',
    Status: getStatus(it)
  }));
  downloadExcel(rows, 'CycleCount', 'CycleCountSheet.xlsx');
}

function exportAlertsExcel() {
  const rows = alertItems().map(it => ({
    SAP_Material: it.StockCode,
    Material_Description: it.Description,
    Bin_location: it.DefaultBin || '',
    Department: it.Department || '',
    Machine: it.Machine || '',
    Stock_on_hand: it.QtyOnHand || 0,
    minimum_QTY: it.MinLevel ?? '',
    maximum_QTY: it.MaxLevel ?? '',
    'On order': it.OnOrder ?? '',
    'To buy': '',
    supplier: '',
    slot: '',
    Status: getStatus(it)
  }));
  if (rows.length === 0) return alert('No stock alerts found.');
  if (rows.every(item => item['On order'] === '') &&
      !confirm('This import has no On order quantities for the stock alerts. Select Cancel to re-upload the updated report, or OK to export with blank On order and To buy cells.')) {
    setActiveTab('upload');
    return;
  }
  downloadExcel(rows, 'StockAlerts', 'Stock_Alerts_Report.xlsx', {
    columnWidths: [16, 42, 18, 18, 24, 17, 16, 16, 14, 12, 20, 14, 15],
    autoFilter: true,
    prepareSheet(sheet) {
      rows.forEach((item, index) => {
        const row = index + 2;
        const max = item.maximum_QTY;
        const onOrder = item['On order'];
        const value = Number.isFinite(max) && Number.isFinite(onOrder)
          ? Math.max(0, max - onOrder - item.Stock_on_hand)
          : '';
        sheet[`J${row}`] = {
          t: typeof value === 'number' ? 'n' : 's',
          v: value,
          f: `IF(COUNT(F${row},H${row}:I${row})=3,MAX(0,H${row}-I${row}-F${row}),"")`
        };
      });
    }
  });
}

/* ── Event Delegation ───────────────────────────────────── */
// Delete item
document.addEventListener('click', e => {
  const delBtn = e.target.closest('button[data-del]');
  if (!delBtn) return;
  const code = delBtn.getAttribute('data-del');
  if (!confirm(`Delete item ${code}?`)) return;
  inventoryData = inventoryData.filter(x => x.StockCode !== code);
  delete stocktake[code];
  saveInventory();
  saveStocktake();
  renderVisibleTab();
});

// Pagination buttons
document.addEventListener('click', e => {
  const btn = e.target.closest('button[data-page-key]');
  if (!btn) return;
  const key = btn.getAttribute('data-page-key');
  const dir = btn.getAttribute('data-page-dir');
  paginationState[key] = Math.max(1, (paginationState[key] || 1) + (dir === 'next' ? 1 : -1));
  renderVisibleTab();
});

// Stocktake input
document.addEventListener('input', e => {
  const input = e.target.closest('input[data-count]');
  if (!input) return;
  const code = input.getAttribute('data-count');
  const countedQty = toNumberOrNull(input.value);
  if (countedQty === null) delete stocktake[code];
  else stocktake[code] = { CountedQty: countedQty, UpdatedAt: new Date().toISOString() };
  saveStocktake();
  const it = inventoryData.find(x => x.StockCode === code);
  const varCell = document.querySelector(`[data-var="${cssEscape(code)}"]`);
  if (varCell) varCell.textContent = countedQty === null ? '' : (countedQty - (it?.QtyOnHand || 0));
  updateDashboard();
});

/* ── Wire up all event listeners ────────────────────────── */
// Navigation
document.querySelectorAll('[data-tab]').forEach(btn => {
  btn.addEventListener('click', () => setActiveTab(btn.dataset.tab));
});

document.querySelectorAll('[data-health-filter]').forEach(button => {
  button.addEventListener('click', () => {
    const status = button.dataset.healthFilter;
    setDashboardHealthFilter(dashboardHealthFilter === status ? '' : status);
  });
});
$('clearHealthFilter').addEventListener('click', () => {
  const activeButton = document.querySelector('[data-health-filter][aria-pressed="true"]');
  setDashboardHealthFilter('');
  activeButton?.focus();
});

$('inventorySearchForm').addEventListener('submit', event => {
  event.preventDefault();
  $('itemsSearch').value = $('globalSearch').value;
  resetPage('items');
  setActiveTab('items');
  $('itemsSearch').focus();
});
$('itemsSearch').addEventListener('input', debounce(() => {
  resetPage('items');
  renderItemsTable();
}));

// Mobile sidebar toggle
document.getElementById('menuToggle')?.addEventListener('click', openSidebar);
document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && $('sidebar').classList.contains('open')) closeSidebar();
});

// Filter chips on alerts page
document.querySelectorAll('.filter-chip').forEach(chip => {
  chip.addEventListener('click', () => {
    document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
    chip.classList.add('active');
    const filterInput = $('alertFilter');
    if (filterInput) filterInput.value = chip.dataset.val;
    resetPage('alerts');
    renderAlertsTable();
  });
});

// Upload
function readInventorySheet(workbook) {
  const mr01 = workbook.SheetNames.find(name => name.trim().toUpperCase() === 'MR01');
  const data = workbook.SheetNames.find(name => name.trim().toUpperCase() === 'DATA');
  const readRows = name => XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '' });

  if (!mr01 && data && data !== workbook.SheetNames[0]) {
    const rows = readRows(data);
    const first = rows[0];
    if (first && Object.hasOwn(first, 'Material') && Object.hasOwn(first, 'Plant stock') &&
        Object.hasOwn(first, 'Material Description')) return { name: data, rows };
  }

  const name = mr01 || workbook.SheetNames[0];
  return { name, rows: readRows(name) };
}

function importInventoryFile(file) {
  if (!file) return;
  if (typeof XLSX === 'undefined') {
    alert('Excel library not loaded. Please run through a local server.');
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const { name, rows } = readInventorySheet(workbook);
      const normalized = rows.map(normalizeItem).filter(Boolean);
      if (!normalized.length) {
        alert('No inventory items found. Check that the sheet has a SAP code column and at least one item.');
        return;
      }
      const existing = { ...stocktake };
      inventoryData = normalized;
      stocktake = {};
      inventoryData.forEach(it => { if (existing[it.StockCode]) stocktake[it.StockCode] = existing[it.StockCode]; });
      saveInventory();
      saveStocktake();
      renderVisibleTab();
      const withOnOrder = normalized.filter(item => item.OnOrder !== null).length;
      alert(`Imported ${formatNumber(normalized.length)} items from ${name}. On order quantities found for ${formatNumber(withOnOrder)} items.`);
    } catch (err) {
      console.error(err);
      alert('Import failed. Check your SAP export format.');
    }
  };
  reader.onerror = () => alert('The workbook could not be read. Try selecting the file again.');
  reader.readAsArrayBuffer(file);
}

$('excelUpload').addEventListener('change', event => {
  importInventoryFile(event.target.files[0]);
  event.target.value = '';
});

// Drag & drop on upload zone
const dropZone = $('dropZone');
if (dropZone) {
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    importInventoryFile(e.dataTransfer.files[0]);
  });
}

$('loadSampleBtn')?.addEventListener('click', async () => {
  if ((inventoryData.length || Object.keys(stocktake).length) &&
      !confirm('Replace the inventory and stocktake on this device with sample data?')) return;
  try {
    const response = await fetch('./sample-data.json');
    if (!response.ok) throw new Error(`Sample request failed (${response.status})`);
    const raw = await response.json();
    inventoryData = raw.map(normalizeItem).filter(Boolean);
  } catch (error) {
    console.error(error);
    alert('Sample data could not be loaded. Please try again.');
    return;
  }
  stocktake = {};
  saveInventory();
  saveStocktake();
  setActiveTab('dashboard');
});

$('searchInput')?.addEventListener('input', debounce(() => { resetPage('preview'); renderPreviewTable(); }));
$('stocktakeFilter')?.addEventListener('input', debounce(() => { resetPage('stocktake'); renderStocktakeTable(); }));
$('alertSearch')?.addEventListener('input', debounce(() => { resetPage('alerts'); renderAlertsTable(); }));

$('lowStockThreshold')?.addEventListener('input', () => {
  settings.lowStockThreshold = Number($('lowStockThreshold').value || 0);
  saveSettings();
  renderVisibleTab();
});

$('clearAllBtn')?.addEventListener('click', () => {
  if (!confirm('Clear ALL inventory + stocktake data on this device?')) return;
  inventoryData = [];
  stocktake = {};
  settings.sessionName = '';
  saveInventory();
  saveStocktake();
  saveSettings();
  renderVisibleTab();
});

$('exportInventoryBtn')?.addEventListener('click', () => {
  if (inventoryData.length === 0) return alert('No inventory data to export.');
  downloadExcel(inventoryData, 'Inventory', 'InventoryData_CurrentFormat.xlsx');
});

$('exportAlertsBtn')?.addEventListener('click', exportAlertsExcel);
$('exportReorderTriggersBtn')?.addEventListener('click', exportAlertsExcel);
$('goToAlertsBtn')?.addEventListener('click', () => setActiveTab('alerts'));

$('addItemBtn')?.addEventListener('click', () => {
  const newItem = normalizeItem({
    StockCode: $('newStockCode').value,
    Description: $('newDescription').value,
    'old stkcode': $('newAlternateKey').value,
    'Stock on hand': $('newQtyOnHand').value,
    'Bin location': $('newDefaultBin').value,
    'minimum QTY': $('newMinLevel').value,
    'maximum QTY': $('newMaxLevel').value
  });
  if (!newItem) return alert('SAP Material and Description are required.');
  const idx = inventoryData.findIndex(x => x.StockCode === newItem.StockCode);
  if (idx >= 0) inventoryData[idx] = { ...inventoryData[idx], ...newItem };
  else inventoryData.push(newItem);
  saveInventory();
  renderVisibleTab();
  ['newStockCode','newDescription','newAlternateKey','newDefaultBin','newMinLevel','newMaxLevel'].forEach(id => { const el = $(id); if (el) el.value = ''; });
  $('newQtyOnHand').value = '0';
});

$('saveSessionNameBtn')?.addEventListener('click', () => {
  settings.sessionName = $('sessionName').value.trim();
  saveSettings();
  alert('Session name saved.');
});

$('markAllUncountedBtn')?.addEventListener('click', () => {
  if (!confirm('Reset ALL counted quantities to blank?')) return;
  stocktake = {};
  saveStocktake();
  renderStocktakeTable();
  updateDashboard();
});

$('exportVarianceBtn')?.addEventListener('click', exportVarianceExcel);
$('exportFullCountBtn')?.addEventListener('click', exportFullCountExcel);
$('exportCycleCountBtn')?.addEventListener('click', exportCycleCountSheet);

/* ── Clock ──────────────────────────────────────────────── */
function updateClock() {
  const el = $('topbarTime');
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}
updateClock();
setInterval(updateClock, 10000);

/* ── Inject sidebar overlay ─────────────────────────────── */
const overlay = document.createElement('div');
overlay.className = 'sidebar-overlay';
overlay.id = 'sidebarOverlay';
document.body.appendChild(overlay);
overlay.addEventListener('click', closeSidebar);

/* ── Boot ───────────────────────────────────────────────── */
injectPaginationStyles();
setActiveTab(activeTab);
