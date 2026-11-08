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
const buyPriorities = new Map();
const priorityStoragePrefix = 'buyPriority:';
const priorityRank = { High: 0, Normal: 1, Low: 2 };
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
  return {
    StockCode:    cleanText(stock),
    Description:  cleanText(desc),
    AlternateKey: cleanText(raw.AlternateKey ?? raw['old stkcode'] ?? raw['old stk code']),
    PartNo:       cleanText(raw.PartNo ?? raw['part no']),
    QtyOnHand:    toNumberOrNull(raw.QtyOnHand ?? raw['Stock on hand '] ?? raw['Stock on hand'] ?? raw.Qty ?? raw['Plant stock']) ?? 0,
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
function downloadExcel(rows, sheetName, filename, { columnWidths, autoFilter = false, extraSheet } = {}) {
  if (typeof XLSX === 'undefined') {
    alert('Excel export library not loaded. Open through a local server or while online once.');
    return;
  }
  const wb = XLSX.utils.book_new();
  function appendSheet(data, name, widths, filter) {
    const sheet = XLSX.utils.json_to_sheet(data);
    if (widths) sheet['!cols'] = widths.map(wch => ({ wch }));
    if (filter) sheet['!autofilter'] = { ref: sheet['!ref'] };
    XLSX.utils.book_append_sheet(wb, sheet, name);
  }
  appendSheet(rows, sheetName, columnWidths, autoFilter);
  if (extraSheet) appendSheet(extraSheet.rows, extraSheet.name, extraSheet.columnWidths, true);
  XLSX.writeFile(wb, filename, { cellStyles: Boolean(columnWidths || extraSheet?.columnWidths) });
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
    // Drop the retired field from inventory saved by older releases.
    inventoryData = (e.target.result || []).map(({ OnOrder, ...item }) => item);
    tx('stocktake').getAll().onsuccess = e2 => {
      stocktake = {};
      (e2.target.result || []).forEach(r => {
        stocktake[r.StockCode] = { CountedQty: r.CountedQty ?? null, UpdatedAt: r.UpdatedAt ?? null };
      });
      tx('settings').getAll().onsuccess = e3 => {
        (e3.target.result || []).forEach(record => {
          if (record.key.startsWith(priorityStoragePrefix)) buyPriorities.set(record.value.id, record.value);
          else settings[record.key] = record.value;
        });
        $('saveBuyPriority').disabled = false;
        updatePriorityCount();
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
  else if (activeTab === 'priorities') renderBuyPriorities();
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
    const code = escapeAttr(it.StockCode);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td>
      <td class="right mono">${escapeHtml(it.MinLevel ?? '—')}</td><td class="right mono">${escapeHtml(it.MaxLevel ?? '—')}</td>
      <td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td>${statusBadge(s)}</td>
      <td><div class="btn-group item-actions"><button type="button" class="btn-ghost" data-add-priority="${code}">Add to Buy Priorities</button><button type="button" data-del="${code}">Del</button></div></td>`;
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
const materialExportWidths = [16, 18, 42, 18, 18, 18, 24];

function materialExportFields(item) {
  return {
    Material: item?.StockCode ?? '',
    'old stk code': item?.AlternateKey ?? '',
    'Material Description': item?.Description ?? '',
    'part no': item?.PartNo ?? '',
    BIN: item?.DefaultBin ?? '',
    department: item?.Department ?? '',
    Machine: item?.Machine ?? ''
  };
}

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
      ...materialExportFields(it),
      Session: settings.sessionName || '',
      SystemQty: it.QtyOnHand || 0,
      CountedQty: obj.CountedQty,
      Variance: variance,
      Status: getStatus(it),
      UpdatedAt: obj.UpdatedAt || ''
    });
  });
  if (rows.length === 0) return alert('No variances found yet.');
  downloadExcel(rows, 'Variances', 'Stocktake_Variances.xlsx', { columnWidths: materialExportWidths });
}

function exportFullCountExcel() {
  if (inventoryData.length === 0) return alert('No inventory loaded.');
  const rows = inventoryData.map(it => {
    const counted = stocktake[it.StockCode]?.CountedQty ?? null;
    return {
      ...materialExportFields(it),
      Session: settings.sessionName || '',
      Stock_on_hand: it.QtyOnHand || 0,
      minimum_QTY: it.MinLevel ?? '',
      maximum_QTY: it.MaxLevel ?? '',
      Status: getStatus(it),
      CountedQty: counted,
      Variance: counted === null ? null : (counted - (it.QtyOnHand || 0))
    };
  });
  downloadExcel(rows, 'FullCount', 'Stocktake_FullCount.xlsx', { columnWidths: materialExportWidths });
}

function exportCycleCountSheet() {
  if (inventoryData.length === 0) return alert('No inventory loaded.');
  const rows = inventoryData.map(it => ({
    ...materialExportFields(it),
    Stock_on_hand: it.QtyOnHand || 0,
    CountedQty: '',
    minimum_QTY: it.MinLevel ?? '',
    maximum_QTY: it.MaxLevel ?? '',
    Status: getStatus(it)
  }));
  downloadExcel(rows, 'CycleCount', 'CycleCountSheet.xlsx', { columnWidths: materialExportWidths });
}

function sortedBuyPriorities() {
  return [...buyPriorities.values()].sort((a, b) =>
    Number(a.done) - Number(b.done) || priorityRank[a.level] - priorityRank[b.level] || a.createdAt - b.createdAt);
}

function matchingPriorityItems(priority) {
  const reference = priority.reference.toLowerCase();
  const field = priority.type === 'bin' ? 'DefaultBin' : 'StockCode';
  return inventoryData.filter(item => String(item[field] ?? '').trim().toLowerCase() === reference);
}

function updatePriorityCount() {
  const count = [...buyPriorities.values()].filter(priority => !priority.done).length;
  $('prioritiesBadge').textContent = formatNumber(count);
  $('prioritiesBadge').hidden = count === 0;
  $('openPriorityCount').textContent = formatNumber(count);
}

function renderBuyPriorities() {
  const priorities = sortedBuyPriorities();
  $('buyPriorityEmpty').hidden = priorities.length > 0;
  $('buyPriorityTable').tBodies[0].innerHTML = priorities.map(priority => {
    const items = matchingPriorityItems(priority);
    const match = items.length === 1 ? items[0].Description : items.length > 1
      ? `${formatNumber(items.length)} matching materials` : 'No inventory match';
    const type = priority.type === 'bin' ? 'Bin' : 'Material';
    const badge = priority.level === 'High' ? 'pill-danger' : priority.level === 'Normal' ? 'pill-warning' : 'pill-ok';
    return `<tr>
      <td><span class="pill ${badge}">${escapeHtml(priority.level)}</span></td>
      <td>${type} ${escapeHtml(priority.reference)}</td>
      <td>${escapeHtml(match)}</td><td>${escapeHtml(priority.note)}</td>
      <td>${priority.done ? 'Done' : 'Open'}</td>
      <td><button type="button" class="btn-ghost" data-priority-toggle="${escapeAttr(priority.id)}">${priority.done ? 'Reopen' : 'Mark done'}</button></td>
    </tr>`;
  }).join('');
}

function saveBuyPriority(priority) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('settings', 'readwrite');
    transaction.oncomplete = () => {
      buyPriorities.set(priority.id, priority);
      updatePriorityCount();
      resolve();
    };
    transaction.onabort = () => reject(transaction.error);
    transaction.objectStore('settings').put({ key: priorityStoragePrefix + priority.id, value: priority });
  });
}

function buildBuyPrioritySheet() {
  const rows = sortedBuyPriorities().filter(priority => !priority.done).flatMap(priority => {
    const items = matchingPriorityItems(priority);
    return (items.length ? items : [null]).map(item => ({
      ...materialExportFields(item ?? (priority.type === 'bin'
        ? { DefaultBin: priority.reference }
        : { StockCode: priority.reference })),
      Priority: priority.level,
      'Look up by': priority.type === 'bin' ? 'Bin location' : 'Material code',
      Reference: priority.reference,
      Stock_on_hand: item?.QtyOnHand ?? '',
      'Follow-up note': priority.note,
      'Inventory match': item ? 'Matched' : 'No match'
    }));
  });
  return { name: 'Buy Priorities', rows, columnWidths: [...materialExportWidths, 12, 18, 18, 17, 42, 18] };
}

function exportBuyPrioritiesExcel() {
  const { name, rows, columnWidths } = buildBuyPrioritySheet();
  if (!rows.length) return alert('No open buy priorities to export.');
  downloadExcel(rows, name, 'Buy_Priorities.xlsx', { columnWidths, autoFilter: true });
}

function exportAlertsExcel() {
  const rows = alertItems().map(it => ({
    ...materialExportFields(it),
    Stock_on_hand: it.QtyOnHand || 0,
    minimum_QTY: it.MinLevel ?? '',
    maximum_QTY: it.MaxLevel ?? '',
    supplier: '',
    slot: '',
    Status: getStatus(it)
  }));
  const prioritySheet = buildBuyPrioritySheet();
  if (rows.length === 0 && prioritySheet.rows.length === 0) return alert('No stock alerts or open buy priorities found.');
  const filename = 'Stock_Alerts_Report.xlsx';
  if (!rows.length) return downloadExcel(prioritySheet.rows, prioritySheet.name, filename, { columnWidths: prioritySheet.columnWidths, autoFilter: true });
  downloadExcel(rows, 'StockAlerts', filename, {
    columnWidths: [...materialExportWidths, 17, 16, 16, 20, 14, 15],
    autoFilter: true,
    extraSheet: prioritySheet.rows.length ? prioritySheet : undefined
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
      alert(`Imported ${formatNumber(normalized.length)} items from ${name}.`);
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

function exportInventoryExcel() {
  if (inventoryData.length === 0) return alert('No inventory data to export.');
  const rows = inventoryData.map(item => {
    const { StockCode, AlternateKey, Description, PartNo, DefaultBin, Department, Machine, ...stockFields } = item;
    return { ...materialExportFields(item), ...stockFields };
  });
  downloadExcel(rows, 'Inventory', 'InventoryData_CurrentFormat.xlsx', { columnWidths: materialExportWidths });
}

$('exportInventoryBtn')?.addEventListener('click', exportInventoryExcel);

$('exportAlertsBtn')?.addEventListener('click', exportAlertsExcel);
$('exportReorderTriggersBtn')?.addEventListener('click', exportAlertsExcel);
$('exportPrioritiesBtn').addEventListener('click', exportBuyPrioritiesExcel);
$('goToAlertsBtn')?.addEventListener('click', () => setActiveTab('alerts'));

$('itemsTable').addEventListener('click', event => {
  const button = event.target.closest('[data-add-priority]');
  if (!button) return;
  $('priorityType').value = 'material';
  $('priorityReference').value = button.dataset.addPriority;
  setActiveTab('priorities');
  $('priorityLevel').focus();
});

$('buyPriorityForm').addEventListener('submit', async event => {
  event.preventDefault();
  const reference = $('priorityReference').value.trim();
  if (!reference) { $('priorityReference').focus(); return; }
  const button = $('saveBuyPriority');
  button.disabled = true;
  try {
    await saveBuyPriority({
      id: crypto.randomUUID(), type: $('priorityType').value, reference,
      level: $('priorityLevel').value, note: $('priorityNote').value.trim(),
      done: false, createdAt: Date.now()
    });
    $('priorityReference').value = '';
    $('priorityNote').value = '';
    $('priorityFeedback').textContent = 'Priority saved.';
    renderBuyPriorities();
    if (activeTab === 'priorities') $('priorityReference').focus();
  } catch (error) {
    $('priorityFeedback').textContent = 'Could not save the priority. Please try again.';
    console.error(error);
  } finally {
    button.disabled = false;
  }
});

$('buyPriorityTable').addEventListener('click', async event => {
  const button = event.target.closest('[data-priority-toggle]');
  if (!button) return;
  const priority = buyPriorities.get(button.dataset.priorityToggle);
  if (!priority) return;
  button.disabled = true;
  try {
    await saveBuyPriority({ ...priority, done: !priority.done });
    renderBuyPriorities();
    if (activeTab === 'priorities') $('buyPriorityTable').querySelector(`[data-priority-toggle="${priority.id}"]`)?.focus();
  } catch (error) {
    $('priorityFeedback').textContent = 'Could not update the priority. Please try again.';
    button.disabled = false;
    console.error(error);
  }
});

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
