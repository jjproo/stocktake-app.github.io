(() => {
  const $wl = id => document.getElementById(id);
  const state = { receipts: null, demand: null, settings: [], date: '', lead: 60 };
  const storageKey = 'invdash-stockout-watchlist-v1';
  const snapshotKey = 'invdash-consumption-snapshot-v1';
  const specifications = {
    receipts: ['Material', 'Date', 'Quantity'],
    demand: ['Material', 'Date', 'Quantity'],
    settings: ['Material', 'Department', 'Bin', 'Criticality', 'Lead days', 'Safety stock', 'Critical minimum', 'Machine']
  };
  const model = globalThis.StockoutModel;
  const pageSize = 75;
  let baseline = [], historyDays = 1, filtered = [], page = 0, loaded = true, error = '';
  let pending = false;
  const number = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });
  const format = value => value === null ? '—' : number.format(value);
  const key = value => String(value ?? '').trim();
  const esc = value => escapeHtml(value ?? '—');

  try { Object.assign(state, JSON.parse(localStorage.getItem(storageKey) || '{}')); }
  catch { error = 'Saved watchlist inputs could not be read.'; }
  function persist() {
    try { localStorage.setItem(storageKey, JSON.stringify(state)); }
    catch { error = 'Changes are available in this session. The browser could not save them.'; }
  }

  function localToday() {
    return new Intl.DateTimeFormat('sv-SE', {timeZone:'Africa/Kampala',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  }
  $wl('wlDate').value = state.date || localToday();
  $wl('wlLead').value = state.lead;

  function index(rows) {
    const map = new Map();
    for (const row of rows || []) {
      if (!map.has(row.material)) map.set(row.material, []);
      map.get(row.material).push(row);
    }
    return map;
  }

  function build() {
    const current = new Map(inventoryData.map(item => [key(item.StockCode),item]));
    const receipts = index(state.receipts), demand = index(state.demand);
    const settings = new Map(state.settings.map(row => [row.material,row]));
    const source = new Map(baseline.map(([material,description,unit,mrp,stock,,rop,cap,safety,net,months,unitFlag,bin,department,machine]) =>
      [material,{material,description,unit,mrp,stock,rop,cap,safety,net,months,unitFlag,bin,department,machine}]));
    for (const [material, item] of current) {
      if (!source.has(material)) source.set(material,{material, description:item.Description, unit:'', mrp:'', stock:item.QtyOnHand,
        rop:item.MinLevel || 0, cap:item.MaxLevel || 0, safety:0, net:0, months:0});
    }
    const options = {today:model.day($wl('wlDate').value), defaultLead:Number($wl('wlLead').value),
      receiptsLoaded:state.receipts !== null, requirementsLoaded:state.demand !== null, historyDays};
    const rows = [];
    for (const [material, item] of source) {
      const stockItem = current.get(material);
      const input = settings.get(material) || {};
      const department = key(input.department || stockItem?.Department || item.department).toLowerCase();
      const merged = {
        ...item,
        description: stockItem?.Description || item.description,
        department: department ? department.replace(/(^|\s)\S/g, s => s.toUpperCase()) : 'Unassigned',
        bin: input.bin || stockItem?.DefaultBin || item.bin || '',
        machine: key(input.machine || stockItem?.Machine || item.machine) || 'Unassigned',
        stock: stockItem?.QtyOnHand ?? item.stock,
        stockSource: stockItem ? 'InvDash' : 'Reorder workbook'
      };
      rows.push({...merged,...model.assess(merged,receipts.get(material) || [],demand.get(material) || [],input,options)});
    }
    return rows;
  }

  function renderPage() {
    const last = Math.max(0,Math.ceil(filtered.length/pageSize)-1);
    page = Math.min(page,last);
    const start = page*pageSize;
    const visible = filtered.slice(start,start+pageSize);
    let department, html = '';
    for (const row of visible) {
      if (department !== row.department) {
        department = row.department;
        html += `<tr class="wl-group"><th colspan="13" scope="rowgroup">${esc(department)}</th></tr>`;
      }
      const date = row.shortage === null ? '—' : model.iso(row.shortage);
      html += `<tr><td class="mono">${esc(row.material)}</td>
        <td>${esc(row.description)}<br><small>${esc(row.mrp || 'MRP unassigned')} · ${esc(row.unit || 'UoM unassigned')}</small></td>
        <td>${esc(row.bin || '—')}</td><td>${esc(row.machine)}</td>
        <td class="${row.priority === 0 || row.critical === 'Line stop' ? 'wl-risk' : ''}">${esc(row.critical)}</td>
        <td>${format(row.stock)}<br><small>${esc(row.stockSource)}</small></td><td>${format(row.net)}</td><td><strong>${format(row.reorder)}</strong><br><small>${esc(row.reorderBasis)}</small></td>
        <td class="${row.priority === 0 ? 'wl-risk' : ''}">${date}</td><td>${row.shortage === null ? '—' : format(row.shortQty)}</td>
        <td>${row.nextReceipt === null ? '—' : model.iso(row.nextReceipt)}</td><td>${format(row.replenishment)}</td>
        <td><strong>${esc(row.status)}</strong>${row.nextAction === row.status ? '' : `<br><small>${esc(row.nextAction)}</small>`}</td></tr>`;
    }
    $wl('wlRows').innerHTML = html || '<tr><td colspan="13">No materials match the filters.</td></tr>';
    $wl('wlSummary').textContent = filtered.length ? `Showing ${number.format(start+1)}–${number.format(start+visible.length)} of ${number.format(filtered.length)} materials` : '0 materials';
    $wl('wlNext').disabled = page >= last;
    $wl('wlPrev').disabled = page === 0;
  }

  function render() {
    if (!loaded) {
      $wl('wlSummary').textContent = error || 'Loading consumption history…';
      return;
    }
    const date = model.day($wl('wlDate').value), lead = Number($wl('wlLead').value);
    if (!Number.isFinite(date) || !Number.isInteger(lead) || lead < 1 || lead > 730) {
      filtered = [];
      renderPage();
      $wl('wlSummary').textContent = 'Enter a snapshot date and lead time from 1 to 730 days.';
      return;
    }
    const rows = build();
    const previous = $wl('wlDepartment').value;
    const departments = Array.from(new Set(rows.map(row => row.department))).sort((a,b) => a.localeCompare(b));
    $wl('wlDepartment').replaceChildren(new Option('All departments',''),...departments.map(name => new Option(name,name)));
    if (departments.includes(previous)) $wl('wlDepartment').value = previous;
    const search = key($wl('wlSearch').value).toLowerCase();
    const mode = $wl('wlFilter').value;
    const department = $wl('wlDepartment').value;
    const machineSelect = $wl('wlMachine');
    const priorMachine = machineSelect.value;
    const machines = Array.from(new Set(rows.filter(row => !department || row.department === department).map(row => row.machine))).sort((a,b) => a.localeCompare(b));
    machineSelect.replaceChildren(new Option('All machines',''),...machines.map(name => new Option(name,name)));
    if (machines.includes(priorMachine)) machineSelect.value = priorMachine;
    const machine = machineSelect.value;
    filtered = rows.filter(row =>
      (!department || row.department === department) && (!machine || row.machine === machine) &&
      (!search || [row.material,row.description,row.bin].some(value => key(value).toLowerCase().includes(search))) &&
      (!mode || (mode === 'risk' && row.shortage !== null) || (mode === 'critical' && row.critical === 'Line stop') || (mode === 'review' && row.priority < 4)));
    filtered.sort((a,b) => a.department.localeCompare(b.department)
      || Number(b.critical === 'Line stop') - Number(a.critical === 'Line stop')
      || a.priority - b.priority || (a.shortage ?? Infinity) - (b.shortage ?? Infinity) || a.material.localeCompare(b.material));
    const baseCount = rows.filter(row => row.stockSource === 'Reorder workbook').length;
    const missingDepartments = rows.filter(row => row.department === 'Unassigned').length;
    const known = new Set(rows.map(row => row.material));
    const unmatched = new Set([...(state.receipts || []), ...(state.demand || []), ...state.settings]
      .map(row => row.material).filter(material => !known.has(material))).size;
    $wl('wlImportState').textContent = [error,
      baseline.length ? "" : "Import your consumption snapshot to calculate consumption-based reorder points.",
      state.receipts === null ? 'Dated receipts missing. Projections assume no incoming supply and suggested quantities are withheld.' : `${number.format(state.receipts.length)} dated receipt rows loaded.`,
      state.demand === null ? 'Dated requirements missing. PD coverage is unverified.' : `${number.format(state.demand.length)} requirements loaded.`,
      baseCount ? `${number.format(baseCount)} materials still use workbook stock.` : '',
      missingDepartments ? `${number.format(missingDepartments)} materials have no department.` : '',
      unmatched ? `${number.format(unmatched)} imported material codes match neither the baseline nor InvDash.` : ''
    ].filter(Boolean).join(' ');
    renderPage();
  }

  function readRows(workbook, kind) {
    const sheet = workbook.Sheets[kind] || workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',blankrows:false});
    const headers = (matrix.shift() || []).map(value => key(value).toLowerCase());
    const expected = specifications[kind];
    const required = kind === 'settings' ? ['Material'] : expected;
    if (!required.every(name => headers.includes(name.toLowerCase()))) {
      throw new Error(`Required columns are ${required.join(', ')}.`);
    }
    const seen = new Set();
    return matrix.map((cells,i) => {
      const get = name => cells[headers.indexOf(name.toLowerCase())];
      const fail = message => { throw new Error(`Row ${i+2} · ${message}`); };
      const numeric = (name, min, max, integer = false) => {
        const value = get(name);
        if (!key(value)) return undefined;
        const n = Number(value);
        if (!Number.isFinite(n) || n < min || n > max || (integer && !Number.isInteger(n))) fail(`Check ${name}.`);
        return n;
      };
      const material = key(get('Material'));
      if (!material) fail('Material is blank.');
      if (kind === 'settings') {
        if (seen.has(material)) fail('Duplicate material in department and planning inputs.');
        seen.add(material);
        const critical = key(get('Criticality'));
        if (critical && !['Line stop','Standard'].includes(critical)) fail('Criticality must be Line stop or Standard.');
        return {material, department:key(get('Department')), bin:key(get('Bin')), machine:key(get('Machine')), critical,
          lead:numeric('Lead days',1,730,true), safety:numeric('Safety stock',0,Infinity), minimum:numeric('Critical minimum',0,Infinity)};
      }
      const inputDate = get('Date');
      let date = model.day(inputDate);
      if (typeof inputDate === 'number') {
        const p = XLSX.SSF.parse_date_code(inputDate,{date1904:workbook.Workbook?.WBProps?.date1904});
        if (p) date = model.day(`${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`);
      }
      if (!Number.isFinite(date)) fail('Use a valid date in YYYY-MM-DD format.');
      const quantity = numeric('Quantity',Number.MIN_VALUE,Infinity);
      if (quantity === undefined) fail('Quantity is blank.');
      return {material,date,quantity};
    });
  }

  for (const [id,kind] of [['wlReceipts','receipts'],['wlDemand','demand'],['wlSettings','settings']]) {
    $wl(id).addEventListener('change', async e => {
      const file = e.target.files[0];
      if (!file || pending) return;
      pending = true;
      try {
        if (typeof XLSX === 'undefined') throw new Error('The spreadsheet reader is unavailable. Open InvDash online and retry.');
        const workbook = XLSX.read(await file.arrayBuffer(),{type:'array',raw:true});
        state[kind] = readRows(workbook,kind);
        error = '';
        persist();
      } catch (problem) { error = `${file.name}: ${problem.message}`; }
      finally { pending = false; e.target.value = ''; page=0; render(); }
    });
  }

  function saveFile(name, data, type) {
    const url = URL.createObjectURL(new Blob([data],{type}));
    const link = document.createElement('a');
    link.href = url;
    link.download = name;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url),3000);
  }
  $wl('wlTemplates').addEventListener('click', () => {
    if (typeof XLSX === 'undefined') { error='Open InvDash online to download the Excel templates.'; render(); return; }
    const book = XLSX.utils.book_new();
    for (const [name,fields] of Object.entries(specifications)) XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([fields]),name);
    const settings = new Map(state.settings.map(row => [row.material,row]));
    const rows = loaded ? build().map(row => {
      const input = settings.get(row.material) || {};
      return [row.material,row.department === 'Unassigned' ? '' : row.department,row.bin,row.critical === 'Unclassified' ? '' : row.critical,input.lead,input.safety,input.minimum,row.machine === 'Unassigned' ? '' : row.machine];
    }) : [];
    XLSX.utils.sheet_add_aoa(book.Sheets.settings,rows,{origin:'A2'});
    XLSX.writeFile(book,'InvDash stockout import templates.xlsx');
  });
  $wl('wlExport').addEventListener('click', () => {
    const safe = value => {
      const raw = String(value ?? '');
      return '"' + (typeof value === 'string' && /^[=+@-]/.test(raw) ? "'" + raw : raw).replaceAll('"','""') + '"';
    };
    const heading = ['Material','Description','Bin location','Machine','Department','MRP','Criticality','Stock','Stock source','Net consumption','History days','Current reorder point','Suggested reorder point','Reorder basis','Lead days','Shortage date','First gap','Next receipt','Suggested qty','Status','Action'];
    const data = filtered.map(r=>[r.material,r.description,r.bin,r.machine,r.department,r.mrp,r.critical,r.stock,r.stockSource,r.net,historyDays,r.rop,r.reorder,r.reorderBasis,r.lead,
      model.iso(r.shortage),r.shortage === null ? '' : r.shortQty,model.iso(r.nextReceipt),r.replenishment,r.status,r.nextAction]);
    saveFile('InvDash stockout watchlist.csv',[heading,...data].map(row=>row.map(safe).join(',')).join('\r\n'),'text/csv;charset=utf-8');
  });
  for (const id of ['wlDepartment','wlMachine','wlFilter','wlDate','wlLead']) $wl(id).addEventListener('change', () => {
    if (id === 'wlDate' || id === 'wlLead') {
      state.date = $wl('wlDate').value;
      state.lead = Number($wl('wlLead').value);
      persist();
    }
    page=0;
    render();
  });
  $wl('wlSearch').addEventListener('input', debounce(() => {page=0;render();}));
  $wl('wlPrev').addEventListener('click', () => { page--; renderPage(); });
  $wl('wlNext').addEventListener('click', () => { page++; renderPage(); });
  globalThis.renderStockoutWatchlist = render;
  function loadSnapshot(data) {
    const first = model.day(data?.start), last = model.day(data?.end);
    if (!Number.isInteger(data?.days) || !Number.isFinite(first) || data.days !== last - first + 1 ||
        !Array.isArray(data.rows) || !data.rows.length || data.rows.length > 50000) {
      throw new Error('Choose a valid InvDash consumption snapshot.');
    }
    const materials = new Set();
    for (const row of data.rows) {
      if (!Array.isArray(row) || ![14,15].includes(row.length) || typeof row[0] !== 'string' || !row[0].trim() ||
          materials.has(row[0]) || ![4,5,6,7,8,9,10].every(i => Number.isFinite(row[i])) ||
          (row.length === 15 && typeof row[14] !== 'string') || ![1,2,3,11,12,13].every(i => typeof row[i] === 'string')) {
        throw new Error('The snapshot contains invalid or duplicate material rows.');
      }
      materials.add(row[0]);
    }
    baseline = data.rows;
    historyDays = data.days;
    $wl('wlHistory').textContent = `SAP consumption from ${data.start} through ${data.end} (${number.format(data.days)} days).`;
  }
  try {
    const saved = localStorage.getItem(snapshotKey);
    if (saved) loadSnapshot(JSON.parse(saved));
  } catch { error = 'The saved consumption snapshot could not be loaded. Import it again.'; }
  $wl('wlBaseline').addEventListener('change', async event => {
    const file = event.target.files[0];
    if (!file) return;
    try {
      if (file.size > 5_000_000) throw new Error('The consumption snapshot is too large for browser storage.');
      const data = JSON.parse(await file.text());
      loadSnapshot(data);
      error = '';
      try { localStorage.setItem(snapshotKey, JSON.stringify(data)); }
      catch { error = 'The snapshot is loaded for this session. Browser storage is full or unavailable.'; }
    } catch (problem) { error = `${file.name}: ${problem.message}`; }
    finally { event.target.value = ''; page = 0; render(); }
  });
  render();
})();
