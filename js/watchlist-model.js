(function (root) {
  const dayMs = 86400000;
  const epsilon = 1e-8;
  const day = value => {
    const text = value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return NaN;
    const parsed = Date.parse(text + 'T00:00:00Z');
    return Number.isFinite(parsed) && new Date(parsed).toISOString().slice(0, 10) === text ? parsed / dayMs : NaN;
  };
  const iso = value => Number.isFinite(value) ? new Date(value * dayMs).toISOString().slice(0, 10) : '';
  const roundUp = (value, unit) => {
    const scale = unit === 'EA' ? 1 : 100;
    return Math.ceil(Math.max(0, value) * scale - epsilon) / scale;
  };

  function assess(item, receipts, requirements, settings, options) {
    const { today, defaultLead, receiptsLoaded, requirementsLoaded, historyDays } = options;
    const { unit, mrp, net, months, rop, cap, stock, safety, unitFlag } = item;
    const lead = settings.lead ?? defaultLead;
    const critical = settings.critical || 'Unclassified';
    const minimum = settings.minimum ?? 0;
    const rate = unitFlag || mrp !== 'V1' || months < 2 ? 0 : Math.max(0, net / historyDays);
    const reorder = Math.max(minimum, unitFlag || net <= 0
      ? rop : roundUp(net / historyDays * lead + (settings.safety ?? safety), unit));
    const scheduled = receipts.filter(row => row.date >= today);
    const overdue = receipts.some(row => row.date < today);
    const demands = requirements.map(row => ({ ...row, date: Math.max(today, row.date) }));
    const horizon = demands.reduce((last,row) => Math.max(last,row.date),today + lead - 1);
    const inbound = new Map();
    const outbound = new Map();
    for (const row of scheduled) inbound.set(row.date, (inbound.get(row.date) || 0) + row.quantity);
    for (const row of demands) outbound.set(row.date, (outbound.get(row.date) || 0) + row.quantity);
    const eventDays = [...new Set([today, horizon, ...inbound.keys(), ...outbound.keys()])].sort((a,b) => a-b);
    const nextReceipt = scheduled.length ? Math.min(...scheduled.map(row => row.date)) : null;
    let shortage = null, shortQty = 0, worst = 0, balance = stock, previous = today - 1;
    for (const date of eventDays) {
      if (date > horizon) break;
      let beforeReceipt = balance - rate * Math.max(0, date - previous - 1);
      const firstGap = rate && balance >= -epsilon ? previous + Math.floor((Math.max(0,balance) + epsilon) / rate) + 1 : null;
      if (firstGap !== null && firstGap < date && shortage === null) {
        shortage = firstGap;
        shortQty = rate * (firstGap - previous) - balance;
      }
      worst = Math.max(worst, -beforeReceipt);
      // Enter exact V1 demand as an override of the daily run rate on that date.
      balance = beforeReceipt + (inbound.get(date) || 0) - Math.max(rate, outbound.get(date) || 0);
      if (balance < -epsilon && shortage === null) {
        shortage = date;
        shortQty = -balance;
      }
      worst = Math.max(worst, -balance);
      previous = date;
    }
    const leadReceipts = scheduled.filter(row => row.date <= today + lead).reduce((sum,row) => sum + row.quantity, 0);
    const position = stock + leadReceipts;
    const refillMinimum = roundUp(minimum - position, unit);
    const replenishment = mrp === 'PD'
      ? Math.max(roundUp(worst, unit), refillMinimum)
      : Math.max(roundUp(worst, unit), position <= reorder
        ? roundUp(Math.max(cap, reorder) - position, unit) : 0);
    let status = 'Covered in projection';
    let priority = 4;
    if (shortage !== null) { status = 'Projected shortage'; priority = 0; }
    else if (minimum > stock && critical === 'Line stop') { status = 'Below critical minimum'; priority = 1; }
    else if (overdue) { status = 'Check overdue receipt'; priority = 2; }
    else if (mrp === 'PD' && !requirementsLoaded) { status = 'Import PD requirements'; priority = 3; }
    else if (unitFlag) { status = 'Check unit of measure'; priority = 3; }
    else if (mrp !== 'PD' && months < 2) { status = 'Review spare minimum'; priority = 3; }
    else if (mrp !== 'PD' && position <= reorder && replenishment > 0) { status = 'Reorder review'; priority = 2; }
    const nextAction = shortage !== null
      ? (nextReceipt === null ? 'Check transfer, substitute or expedited buy' : 'Check receipt timing; expedite or transfer')
      : priority < 4 ? status : 'Monitor';
    const reorderBasis = unitFlag || net <= 0 ? "Existing SAP point / minimum" : "Lead-time usage + safety stock";
    return { critical, lead, reorder, reorderBasis, shortage, shortQty: Math.max(0,shortQty), nextReceipt,
      replenishment: (!receiptsLoaded || (mrp === 'PD' && !requirementsLoaded) || unitFlag) ? null : replenishment,
      status, priority, nextAction, overdue };
  }

  root.StockoutModel = { day, iso, assess };
})(globalThis);

