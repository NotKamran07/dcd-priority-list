// === INITIALISATION, THEME & UI HELPERS ===
document.getElementById('date-badge').textContent =
  new Date().toLocaleDateString('en-GB', {weekday:'long', day:'numeric', month:'long', year:'numeric'});

function escAttr(s) {
  return String(s).split('&').join('&amp;').split('"').join('&quot;').split('<').join('&lt;');
}

function showToast(msg) {
  const t = document.createElement('div');
  t.textContent = msg;
  t.style.cssText = 'position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--surface2);color:var(--text);border:1px solid var(--accent2);border-radius:6px;padding:0.6rem 1.1rem;font-size:0.75rem;z-index:100001;box-shadow:0 4px 16px rgba(0,0,0,0.4);';
  document.body.appendChild(t);
  setTimeout(() => { if (t.parentNode) t.parentNode.removeChild(t); }, 5000);
}

function toggleTheme() {
  const isLight = document.body.classList.toggle('light-mode');
  document.getElementById('theme-label').textContent = isLight ? '☀️ Light' : '🌙 Dark';
  localStorage.setItem('dcd-theme', isLight ? 'light' : 'dark');
}

// Restore saved theme preference
(function() {
  const saved = localStorage.getItem('dcd-theme');
  if (saved === 'light') {
    document.body.classList.add('light-mode');
    document.addEventListener('DOMContentLoaded', () => {
      const lbl = document.getElementById('theme-label');
      if (lbl) lbl.textContent = '☀️ Light';
    });
  }
})();

// === PRIORITY LIST LOADING, INDEXING & FILTERING ===
let loadAllRunning = false, loadAllPending = false;
async function loadAll(quiet) {
  if (loadAllRunning) { loadAllPending = true; return; }   // already loading: run once more when it finishes
  loadAllRunning = true;
  try { await loadAllNow(quiet); }
  finally {
    loadAllRunning = false;
    if (loadAllPending) { loadAllPending = false; loadAll(true); }
  }
}

async function loadAllNow(quiet) {
  // Load file upload timestamp
  try {
    const snap = await db.ref('filesUploadedAt').once('value');
    filesUploadedAt = snap.val() || 0;
  } catch(e) { filesUploadedAt = 0; }
  // Load custom parts added via the web app
  try {
    const cpSnap = await CUSTOM_PARTS_REF.once('value');
    customParts = cpSnap.val() || {};
  } catch(e) { customParts = {}; }
  // Load the DCD parts list from Firebase (seeded from the built-in list on first run)
  try {
    const dsnap = await DCD_PARTS_REF.once('value');
    let dlist = dsnap.val();
    if (!dlist) {
      dlist = {};
      DCD_DEFAULTS.forEach(p => { dlist[safeKey(p)] = p; });
      await DCD_PARTS_REF.set(dlist);
    }
    DCD_PARTS.clear();
    Object.values(dlist).forEach(p => DCD_PARTS.add(p));
  } catch(e) { console.error('DCD list load failed, using built-in list', e); }
  try {
    const migRef = db.ref('migrations/dcdRemoveTemporary1');
    if (!(await migRef.once('value')).val()) {
      DCD_REMOVE_ONCE.forEach(p => { DCD_PARTS.delete(p); DCD_PARTS_REF.child(safeKey(p)).remove(); });
      await migRef.set(Date.now());
    }
  } catch(e) { console.error('DCD clean-up skipped', e); }
  // Load the app's own grade list (filled once from the built-in starting list)
  try {
    const gsnap = await GRADES_REF.once('value');
    let g = gsnap.val();
    if (!g) {
      g = {};
      Object.keys(GRADE_SEED).forEach(p => { g[safeKey(p)] = { partId: p, grade: GRADE_SEED[p] }; });
      Object.values(customParts).forEach(cp => { if (cp.partId && cp.grade) g[safeKey(cp.partId)] = { partId: cp.partId, grade: cp.grade }; });
      await GRADES_REF.set(g);
    }
    partGrades = {};
    Object.values(g).forEach(e => { if (e && e.partId && e.grade) partGrades[e.partId] = e.grade; });
  } catch(e) { console.error('Grade list load failed, using built-in list', e); partGrades = Object.assign({}, GRADE_SEED); }
  try {
    const fixRef = db.ref('migrations/gradeFixes1');
    if (!(await fixRef.once('value')).val()) {
      for (const [p, g] of Object.entries(GRADE_FIXES_ONCE)) {
        partGrades[p] = g;
        await GRADES_REF.child(safeKey(p)).set({ partId: p, grade: g });
      }
      await fixRef.set(Date.now());
      // also grade any of this user's earlier ungraded history for those parts (retry while history loads)
      [3000, 8000].forEach(ms => setTimeout(() => Object.entries(GRADE_FIXES_ONCE).forEach(([p, g]) => applyGradeToHistory(p, g)), ms));
    }
  } catch(e) { console.error('Grade corrections skipped', e); }
  // Load the department part lists (filled once from the built-in starting lists)
  try {
    const dsnap2 = await DEPT_PARTS_REF.once('value');
    let dv = dsnap2.val();
    if (!dv) {
      dv = {};
      ['PC', 'PCB', 'Printers'].forEach(d => {
        dv[d] = {};
        (DEPT_SEED[d] || []).forEach(p => { dv[d][safeKey(p)] = p; });
      });
      (DEPT_SEED.DCDExtra || []).forEach(p => { DCD_PARTS.add(p); DCD_PARTS_REF.child(safeKey(p)).set(p); });
      await DEPT_PARTS_REF.set(dv);
    }
    ['PC', 'PCB', 'Printers'].forEach(d => {
      DEPT_PARTS[d].clear();
      Object.values(dv[d] || {}).forEach(p => DEPT_PARTS[d].add(p));
    });
  } catch(e) { console.error('Department lists load failed', e); }
  await applyDcdMoveOnce();

  if (!quiet) {   // a background refresh leaves the screen as it is until the new data is ready
    document.getElementById('loading-bar').style.display = 'flex';
    document.getElementById('stats-bar').style.display = 'none';
    document.getElementById('actions-bar').style.display = 'none';
    document.getElementById('table-wrap').style.display = 'none';
  }
  document.getElementById('error-msg').style.display = 'none';

  try {
    document.getElementById('loading-msg').textContent = 'Loading PowerBI data (Both PowerBi Complete.xlsx)...';
    const pbiRows = await loadXLSX('Both PowerBi Complete.xlsx');

    document.getElementById('loading-msg').textContent = 'Loading backorder report (New Repair Centre Back Order Report v1.csv)...';
    const boRows = await loadCSV('New Repair Centre Back Order Report v1.csv');

    document.getElementById('loading-msg').textContent = 'Building priority list...';

    // Grades come from the app's own list (Firebase), not from the PowerBI file
    window.gradeMap = Object.assign({}, partGrades);
    // Custom parts only fill in grades the main list doesn't have
    Object.values(customParts).forEach(cp => {
      if (!cp.partId) return;
      if (cp.grade && !window.gradeMap[cp.partId]) window.gradeMap[cp.partId] = cp.grade;
      if (cp.isDCD) DCD_PARTS.add(cp.partId);
    });
    const gradeMap = window.gradeMap;

    // Every part we know about (all departments) so Add Part can find anything, plus any new parts
    const idx = buildPartIndex(pbiRows, boRows);
    window.allPartIds = idx.all;
    window.partTeams = idx.teams;
    try { autoAssigned = (await db.ref('autoAssigned').once('value')).val() || {}; } catch (e) { autoAssigned = {}; }
    try { noDeptParts = new Set(Object.values((await NO_DEPT_REF.once('value')).val() || {})); } catch (e) { noDeptParts = new Set(); }
    try { await autoAssignFromTeams(idx.active, idx.teams); } catch (e) { console.error('Auto-assign skipped', e); }
    const np = computeNewParts(idx.active);
    window.newParts = np.list;
    window.newReasons = np.reasons;
    updateNewBtn();

    const pbiMap = {};
    pbiRows.forEach(r => {
      const id = (r['Part ID'] || '').toString().trim();
      if (!id || !activeParts().has(id)) return;
      pbiMap[id] = { repReq: Number(r['Rep Req']) || 0, totalBad: Number(r['Total Bad']) || 0, team: (r['Repair Team'] || '').toString().trim() };
    });

    const boMap = {};
    const boZeroQtyParts = new Set(); // parts with qty=0 in backorder report
    applyKeywordHiding(boRows);

    boRows.forEach(r => {
      const id = (r['Part ID'] || '').toString().trim();
      if (!id || !activeParts().has(id)) return;
      if (excludedByNote.has(id)) return; // Skip excluded parts
      if (!boMap[id]) boMap[id] = { count: 0, totalQty: 0 };
      boMap[id].count++;
      const qty = Number(r['Qty']) || 0;
      boMap[id].totalQty = Math.max(boMap[id].totalQty, qty); // MAX across rows, not sum (Qty repeats per row)
      if (qty === 0) boZeroQtyParts.add(id); // flag parts with 0 qty
    });

    const boIds = new Set(Object.keys(boMap));
    const rank1 = [], rank2 = [];

    Object.keys(boMap).forEach(id => {
      if (excludedByNote.has(id)) return; // Skip excluded parts
      const bo = boMap[id], pbi = pbiMap[id] || { repReq: 0, totalBad: 0, team: '' };
      const cls = gradeMap[id] || '';
      // Backorder report Qty is the primary source of truth for stock quantity.
      // Only fall back to PowerBI Total Bad if the part isn't on the backorder report at all.
      let totalQty = bo.totalQty;
      rank1.push({ Priority: 1, 'Part ID': id, 'Repair Team': pbi.team, 'Rep Req': pbi.repReq, 'Open Orders': bo.count, 'Total Qty': totalQty, 'Part Class': cls, Value: GRADE_VALUE[cls] ?? '' });
    });
    rank1.sort((a, b) => b['Open Orders'] - a['Open Orders']);

    Object.keys(pbiMap).forEach(id => {
      if (boIds.has(id)) return;
      if (excludedByNote.has(id)) return; // hidden by backorder keyword filter
      const pbi = pbiMap[id];
      if (pbi.repReq === 0) return;
      const cls = gradeMap[id] || '';
      rank2.push({ Priority: 2, 'Part ID': id, 'Repair Team': pbi.team, 'Rep Req': pbi.repReq, 'Open Orders': '', 'Total Qty': pbi.totalBad, 'Part Class': cls, Value: GRADE_VALUE[cls] ?? '' });
    });
    rank2.sort((a, b) => b['Rep Req'] - a['Rep Req']);

    outputRows = [...rank1, ...rank2];

    document.getElementById('stat-r1').textContent = rank1.length;
    document.getElementById('stat-r2').textContent = rank2.length;
    document.getElementById('stat-total').textContent = outputRows.length;
    document.getElementById('loading-bar').style.display = 'none';
    document.getElementById('stats-bar').style.display = 'flex';
    document.getElementById('tab-bar').style.display = 'flex';
    document.getElementById('actions-bar').style.display = 'flex';
    document.getElementById('table-wrap').style.display = 'block';
    updatePriorityTotals();
    document.getElementById('last-updated').style.display = 'block';
    document.getElementById('last-updated').textContent = 'Loaded at ' + new Date().toLocaleTimeString('en-GB');
    // Set list updated bar
    const listUpdatedBar = document.getElementById('list-updated-bar');
    const listUpdatedText = document.getElementById('list-updated-text');
    if (listUpdatedBar && listUpdatedText) {
      listUpdatedBar.style.display = 'flex';
      listUpdatedText.textContent = 'Priority list last updated: ' + new Date().toLocaleString('en-GB', { weekday:'short', day:'numeric', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
    }
    document.getElementById('live-status').style.display = 'block';
    sortCol = null;
    renderTable();
    if (currentTab !== 'main') {   // a refresh must not switch you back to the priority list
      document.getElementById('actions-bar').style.display = 'none';
      document.getElementById('table-wrap').style.display = 'none';
      document.getElementById('zero-section').style.display = 'none';
      if (currentTab === 'done') renderCompleted();
      else if (currentTab === 'history') renderHistory();
    }
    dataLoaded = true;
    // Now that outputRows is populated, save any existing ticks/BERs to history
    saveToHistory();
    const manageOverlay = document.getElementById('manage-overlay');
    if (manageOverlay && manageOverlay.style.display === 'flex') renderManageList();
  } catch(err) {
    document.getElementById('loading-bar').style.display = 'none';
    document.getElementById('error-msg').style.display = 'block';
    document.getElementById('error-msg').textContent = '⚠ ' + err.message;
  }
}

// A real part ID starts with letters/digits then a hyphen (spaces are allowed inside). Filters out stray dates, headers and text blobs in the reports
function looksLikePartId(id) { return id.length <= 60 && /^[A-Za-z0-9]+-[A-Za-z0-9\/.+ -]+$/.test(id); }

// Every known part, the PowerBI team for each, and the "new parts" that need a department or a grade
function buildPartIndex(pbiRows, boRows) {
  const all = new Set(), teams = {}, active = new Set();
  pbiRows.forEach(r => {
    const id = (r['Part ID'] || '').toString().trim();
    if (!looksLikePartId(id)) return;
    all.add(id);
    if (r['Repair Team'] && !teams[id]) teams[id] = String(r['Repair Team']).trim();
    if (Number(r['Rep Req']) > 0) active.add(id);
  });
  boRows.forEach(r => {
    const id = (r['Part ID'] || '').toString().trim();
    if (!looksLikePartId(id)) return;
    all.add(id); active.add(id);
  });
  Object.keys(partGrades).forEach(p => all.add(p));
  Object.values(DEPT_PARTS).forEach(set => set.forEach(p => all.add(p)));
  return { all, teams, active };
}

// Hide a part from the priority list when EVERY back order call for it says the work is done:
//  - Note contains "done", "completed" or "in progress", or
//  - Status is AT DEPLOYMENT (unless the call says "no stock" anywhere, e.g. "no stock at deployment")
function applyKeywordHiding(boRows) {
  excludedByNote.clear();
  const EXCLUDED_KEYWORDS = ['done', 'completed', 'in progress'];
  const partNotes = {};
  boRows.forEach(r => {
    const partId = r['Part ID'];
    const note = (r['Note'] || '').toLowerCase().trim();
    const status = (r['Status'] || '').toLowerCase().trim();
    if (!partId) return;
    if (!partNotes[partId]) partNotes[partId] = { total: 0, excluded: 0, reasons: new Set() };
    partNotes[partId].total++;
    const matchedKw = note ? EXCLUDED_KEYWORDS.find(kw => note.includes(kw)) : null;
    const noStock = /\bno stock\b/.test((r['AllText'] || '') + ' ' + note + ' ' + status);
    const isDeployExcluded = status === 'at deployment' && !noStock;
    if (matchedKw) partNotes[partId].reasons.add('Note: "' + matchedKw + '"');
    if (isDeployExcluded) partNotes[partId].reasons.add('Status: AT DEPLOYMENT');
    if (matchedKw || isDeployExcluded) partNotes[partId].excluded++;
  });
  hiddenParts = [];
  Object.keys(partNotes).forEach(partId => {
    const p = partNotes[partId];
    if (p.total > 0 && p.excluded >= p.total) {
      excludedByNote.add(partId);
      if (activeParts().has(partId)) hiddenParts.push({ id: partId, calls: p.total, reasons: [...p.reasons] });
    }
  });
  hiddenParts.sort((a, b) => a.id.localeCompare(b.id));
  updateHiddenBtn();
}

function partOnAnyList(id) { return Object.values(DEPT_PARTS).some(set => set.has(id)); }

// A part that is active in the reports, on no department list, and whose PowerBI team IS a real
// department gets put on that department's list (once). The admin can still change it afterwards.
async function autoAssignFromTeams(active, teams) {
  const todo = [];
  active.forEach(id => {
    const d = TEAM_TO_DEPT[(teams[id] || '').trim().toUpperCase()];
    if (!d || partOnAnyList(id) || autoAssigned[safeKey(id)] || noDeptParts.has(id)) return;
    todo.push([id, d]);
  });
  if (!todo.length) return 0;
  const updates = {};
  todo.forEach(([id, d]) => {
    const k = safeKey(id);
    DEPT_PARTS[d].add(id);                       // apply on this screen straight away
    autoAssigned[k] = d;
    updates[(d === 'DCD' ? 'dcdParts/' : 'deptParts/' + d + '/') + k] = id;
    updates['autoAssigned/' + k] = d;
  });
  await db.ref().update(updates);                // shared: every screen picks it up
  logChange('Auto-assigned from PowerBI team', '', todo.length + ' part(s): '
    + todo.slice(0, 6).map(([id, d]) => id + ' → ' + DEPT_LABELS[d]).join(', ') + (todo.length > 6 ? ' …' : ''));
  return todo.length;
}

// Parts that still need a person to decide: active in the reports and (no department that PowerBI can
// supply, or no grade). Returns the list plus why each one is on it.
function computeNewParts(active) {
  const gm = window.gradeMap || {};
  const reasons = {};
  active.forEach(id => {
    const needs = [];
    if (!partOnAnyList(id) && !autoAssigned[safeKey(id)] && !noDeptParts.has(id)) needs.push('needs department');
    if (!gm[id]) needs.push('needs grade');
    if (needs.length) reasons[id] = needs.join(' + ');
  });
  return { list: Object.keys(reasons).sort(), reasons };
}

function updateNewBtn() {
  const b = document.getElementById('new-btn');
  if (!b) return;
  const n = (window.newParts || []).length;
  b.style.display = n ? 'inline-block' : 'none';
  b.textContent = '🆕 ' + n + ' new';
}

function openNewParts() {
  openManageModal();
  document.getElementById('manage-search').value = '!';
  renderManageList();
}

function updateHiddenBtn() {
  const b = document.getElementById('hidden-btn');
  if (!b) return;
  b.style.display = hiddenParts.length ? 'inline-block' : 'none';
  b.textContent = '👁 ' + hiddenParts.length + ' hidden';
}

// === PRIORITY TABLE RENDERING, SORTING & PRINTING ===
function updatePriorityTotals() {
  let rep = 0, stock = 0;
  outputRows.forEach(r => {
    if (Number(r['Total Qty']) === 0 || r['Total Qty'] === '') return;
    const eff = getEffectiveRow(r);
    rep += Number(eff['Rep Req']) || 0;
    stock += Number(eff['Total Qty']) || 0;
  });
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  set('stat-rep-total', rep);
  set('stat-stock-total', stock);
  set('tab-count-main', rep);   // Priority List badge = total rep req
}

function renderTable() {
  updatePriorityTotals();
  const search = document.getElementById('search-input').value.toLowerCase();
  let rows = outputRows.filter(r => {
    if (Number(r['Total Qty']) === 0 || r['Total Qty'] === '') return false;
    const eff = getEffectiveRow(r);
    // Only hide when FULLY done (open orders AND rep req both at 0)
    if (!showDone && eff._fullyDone) return false;
    if (deptFilter !== 'all' && (r['Repair Team'] || '').trim() !== deptFilter) return false;
    return !search || r['Part ID'].toLowerCase().includes(search) || (r['Repair Team'] || '').toLowerCase().includes(search);
  });
  if (sortCol) {
    rows = [...rows].sort((a, b) => {
      let av = a[sortCol], bv = b[sortCol];
      if (av === '' || av == null) av = sortAsc ? Infinity : -Infinity;
      if (bv === '' || bv == null) bv = sortAsc ? Infinity : -Infinity;
      if (typeof av === 'number' && typeof bv === 'number') return sortAsc ? av - bv : bv - av;
      return sortAsc ? String(av).localeCompare(String(bv)) : String(bv).localeCompare(String(av));
    });
  }
  const tbody = document.getElementById('table-body');
  tbody.innerHTML = '';
  rows.forEach(row => {
    const eff = getEffectiveRow(row);
    const tickCount = getTickCount(row['Part ID']);
    const isDone = tickCount > 0;
    const isFullyDone = eff._fullyDone || false;

    // Only promote to rank 2 when open orders fully reach 0
    let effPriority = eff.Priority;
    if (isFullyDone && eff.Priority === 1 && eff['Open Orders'] === 0) effPriority = 2;

    const tr = document.createElement('tr');
    tr.className = isFullyDone ? 'completed' : (effPriority === 1 ? 'rank1' : 'rank2');
    const cls = eff['Part Class'] || '';
    const valStr = typeof eff.Value === 'number' ? `£${eff.Value.toFixed(2)}` : '';
    const partId = eff['Part ID'];

    // Show effective badge based on open orders
    const displayPriority = (eff.Priority === 1 && isDone && eff['Open Orders'] === 0) ? 2 : eff.Priority;
    const openOrdersDisplay = eff.Priority === 1
      ? (isDone ? eff['Open Orders'] : (eff['Open Orders'] !== '' ? eff['Open Orders'] : ''))
      : '';

    const rowTickCount = getTickCount(partId);
    const berCount = getBerCount(partId);
    const origTotalQty = typeof row['Total Qty'] === 'number' ? row['Total Qty'] : 0;
    // Count all users' ticks for this part
    const allTicksForPart = Object.entries(allUserTicks).reduce((sum, [uid, userTicks]) => {
      if (!userTicks) return sum;
      const k = safeKey(partId);
      const t = userTicks[k];
      return sum + (t && typeof t === 'object' ? (t.count || 0) : 0);
    }, 0);
    const usedQty = allTicksForPart + berCount;
    const plusDisabled = origTotalQty <= 0 || usedQty >= origTotalQty;
    // BER + button in Completed Today is never limited by stock - it's a record of completed work
    const berPlusDisabled = false;
    // Build user tick indicators
    const userIndicators = Object.entries(allUserTicks).map(([uid, userTicks]) => {
      if (!userTicks) return '';
      const k = safeKey(partId);
      const t = userTicks[k];
      const count = t && typeof t === 'object' ? (t.count || 0) : 0;
      if (count === 0) return '';
      const u = USERS[uid];
      return u ? `<span class="user-tick-dot" style="background:${u.color};">${u.initials}:${count}</span>` : '';
    }).join('');
    tr.innerHTML = `
      <td class="tick-cell">
        <div class="tick-counter">
          <button class="tick-minus" onclick="removeTick('${partId}')" ${rowTickCount === 0 ? 'disabled' : ''}>−</button>
          <span class="tick-num ${rowTickCount > 0 ? 'ticked' : ''}">${rowTickCount}</span>
          <button class="tick-plus" onclick="addTick('${partId}')" ${plusDisabled ? 'disabled' : ''}>+</button>
        </div>
        ${userIndicators ? `<div class="user-tick-indicators">${userIndicators}</div>` : ''}
      </td>
      <td class="ber-cell">
        <div class="ber-counter">
          <button class="ber-btn ber-minus" onclick="removeBer('${partId}')" ${berCount === 0 ? 'disabled' : ''}>−</button>
          <span class="ber-num ${berCount > 0 ? 'berred' : ''}">${berCount}</span>
          <button class="ber-btn" onclick="addBer('${partId}')" ${berPlusDisabled ? 'disabled' : ''}>+</button>
        </div>
      </td>
      <td><span class="badge badge-${displayPriority}">${displayPriority}</span></td>
      <td class="part-id">${partId}</td>
      <td>${eff['Repair Team'] || ''}</td>
      <td class="num-cell">${eff['Rep Req'] !== '' ? eff['Rep Req'] : ''}</td>
      <td class="num-cell">${openOrdersDisplay !== '' ? openOrdersDisplay : ''}</td>
      <td class="num-cell">${eff['Total Qty'] !== '' ? eff['Total Qty'] : ''}${eff._manualAdj ? '<span style="font-size:0.6rem;color:var(--accent);margin-left:3px;">(-' + eff._manualAdj + ')</span>' : ''}</td>
      <td class="num-cell grade-${cls}">${cls}</td>
      <td class="num-cell">${valStr}</td>
      <td class="num-cell"><button data-pid="${partId}" onclick="openAdjustModal(this.dataset.pid)" style="font-family:'Syne',sans-serif;font-weight:700;font-size:0.7rem;padding:3px 8px;background:var(--surface2);border:1px solid var(--border);color:var(--text);cursor:pointer;border-radius:2px;">⚙</button></td>`;
    tbody.appendChild(tr);
  });
  document.getElementById('row-count').textContent = `${rows.length} parts shown`;

  // Zero qty section
  const zeroRows = outputRows.filter(r => {
    const eff = getEffectiveRow(r);
    if (deptFilter !== 'all' && (r['Repair Team'] || '').trim() !== deptFilter) return false;
    return (Number(r['Total Qty']) === 0 || r['Total Qty'] === '') && !eff._fullyDone;
  });
  const zeroSection = document.getElementById('zero-section');
  const zeroBody = document.getElementById('zero-body');
  document.getElementById('zero-count').textContent = `(${zeroRows.length})`;
  zeroBody.innerHTML = '';
  if (zeroRows.length > 0 && outputRows.length > 0 && currentTab === 'main') {
    zeroSection.style.display = 'block';
    zeroRows.forEach(row => {
      const tr = document.createElement('tr');
      const cls = row['Part Class'] || '';
      const valStr = typeof row.Value === 'number' ? `£${row.Value.toFixed(2)}` : '';
      tr.innerHTML = `
        <td><span class="zero-badge">${row.Priority}</span></td>
        <td class="part-id">${row['Part ID']}</td>
        <td>${row['Repair Team'] || ''}</td>
        <td class="num-cell">${row['Rep Req'] !== '' ? row['Rep Req'] : ''}</td>
        <td class="num-cell">—</td>
        <td class="num-cell grade-${cls}">${cls}</td>
        <td class="num-cell">${valStr}</td>`;
      zeroBody.appendChild(tr);
    });
  } else {
    zeroSection.style.display = 'none';
  }
}

function setDept(dept, btn) {
  deptFilter = dept;
  document.querySelectorAll('.dept-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderTable();
}

function sortBy(col) {
  sortAsc = sortCol === col ? !sortAsc : false;
  sortCol = col;
  renderTable();
}

function toggleDone() {
  showDone = !showDone;
  document.getElementById('show-done-btn').textContent = showDone ? 'Hide Completed' : 'Show Completed';
  renderTable();
}

function resetTicks() {
  if (confirm('Reset all ticks for today? This will clear completed items for everyone.')) {
    TICKS_REF.remove();
  }
}

function togglePrintMenu(e) {
  e.stopPropagation();
  const menu = document.getElementById('print-menu');
  menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
}
function closePrintMenu() {
  document.getElementById('print-menu').style.display = 'none';
}
// Close print menu when clicking elsewhere
document.addEventListener('click', () => closePrintMenu());

function printList(dept) {
  // Set print header
  const deptLabel = dept === 'all' ? 'All Departments' : dept;
  const dateStr = new Date().toLocaleDateString('en-GB', {weekday:'long', day:'numeric', month:'long', year:'numeric'});
  document.getElementById('print-header').textContent = `Priority Work List — ${deptLabel} — ${dateStr}`;

  // Temporarily filter table to selected dept
  const prevFilter = deptFilter;
  deptFilter = dept === 'all' ? 'all' : dept;
  renderTable();

  // Print
  window.print();

  // Restore previous filter
  deptFilter = prevFilter;
  renderTable();
}

// === STOCK ADJUSTMENTS & TICK / BER COUNTERS ===
function openAdjustModal(partId) {
  adjustModalPartId = partId;
  const row = outputRows.find(r => r['Part ID'] === partId);
  if (!row) return;
  const reportQty = typeof row['Total Qty'] === 'number' ? row['Total Qty'] : 0;
  const currentAdj = stockAdjustments[safeKey(partId)] || 0;
  const effQty = Math.max(0, reportQty - currentAdj);
  document.getElementById('adjust-part-id').textContent = partId;
  document.getElementById('adjust-report-qty').textContent = reportQty;
  document.getElementById('adjust-current-qty2').textContent = effQty;
  const adjEl = document.getElementById('adjust-diff');
  if (currentAdj === 0) {
    adjEl.textContent = 'None';
    adjEl.style.color = 'var(--muted)';
  } else {
    adjEl.textContent = (currentAdj > 0 ? '-' : '+') + Math.abs(currentAdj);
    adjEl.style.color = currentAdj > 0 ? 'var(--accent)' : 'var(--green)';
  }
  document.getElementById('adjust-qty-input').value = 1;
  document.getElementById('adjust-overlay').style.display = 'flex';
}

function closeAdjustModal() {
  document.getElementById('adjust-overlay').style.display = 'none';
  adjustModalPartId = null;
}

function resetAdjust() {
  if (!adjustModalPartId) return;
  STOCK_ADJ_REF.child(safeKey(adjustModalPartId)).remove();
  closeAdjustModal();
}

function confirmAdjust(action) {
  if (!adjustModalPartId) return;
  const key = safeKey(adjustModalPartId);
  const qty = parseInt(document.getElementById('adjust-qty-input').value) || 1;
  const current = stockAdjustments[key] || 0;
  if (action === 'add') {
    STOCK_ADJ_REF.child(key).set(current - qty);
  } else {
    STOCK_ADJ_REF.child(key).set(current + qty);
  }
  closeAdjustModal();
}

function addTick(partId) {
  const key = safeKey(partId);
  const current = (ticks[key] && typeof ticks[key] === 'object') ? (ticks[key].count || 0) : 0;
  TICKS_REF.child(key).transaction(t => nextEntry(t, 1));
}

function removeTick(partId) {
  const key = safeKey(partId);
  const current = (ticks[key] && typeof ticks[key] === 'object') ? (ticks[key].count || 0) : 0;
  TICKS_REF.child(key).transaction(t => nextEntry(t, -1));
}

// Raw tick count — all ticks for today (for display/history)
function getTickCount(partId) {
  const key = safeKey(partId);
  const tick = ticks[key];
  if (!tick) return 0;
  if (typeof tick === 'object') return tick.count || 0;
  return tick ? 1 : 0;
}

// Raw BER count — all BERs for today (for display/history)
function getBerCount(partId) {
  const key = safeKey(partId);
  const ber = bers[key];
  if (!ber) return 0;
  if (typeof ber === 'object') return ber.count || 0;
  return 0;
}

// Filtered BER count — only post-upload BERs (for quantity calculations)
function getFilteredBerCount(partId) {
  return entryPostCount(bers[safeKey(partId)]);
}

function addBer(partId) {
  const key = safeKey(partId);
  const currentBer = getBerCount(partId);
  // Allow BER even for parts not in outputRows (e.g. manually added in Completed Today)
  BER_REF.child(key).transaction(b => nextEntry(b, 1));
}

function removeBer(partId) {
  const key = safeKey(partId);
  const current = getBerCount(partId);
  if (current <= 0) return;
  BER_REF.child(key).transaction(b => nextEntry(b, -1));
}

function getAllUserTickCount(partId) {
  // Sum ticks across ALL users for Open Orders calculation
  const key = safeKey(partId);
  let total = 0;
  Object.values(allUserTicks).forEach(userTicks => {
    if (!userTicks) return;
    const t = userTicks[key];
    total += t && typeof t === 'object' ? (t.count || 0) : 0;
  });
  return total;
}

// Each tick/BER entry is {count, pre, lastUpdated}. 'pre' = how many of 'count' were already
// done when the latest PowerBI/backorder upload happened, so only newer work affects quantities.
function entryPostCount(t) {
  if (!t) return 0;
  if (typeof t !== 'object') return filesUploadedAt > 0 ? 0 : 1;   // legacy true/false tick
  if (filesUploadedAt > 0 && (!t.lastUpdated || t.lastUpdated <= filesUploadedAt)) return 0;
  return Math.max(0, (t.count || 0) - (t.pre || 0));
}

// Next stored value for an entry after adding (delta > 0) or removing (delta < 0) work. null = delete.
function nextEntry(old, delta) {
  const isObj = old && typeof old === 'object';
  const oldCount = isObj ? (old.count || 0) : (old ? 1 : 0);
  const wasPre = filesUploadedAt > 0 && (!isObj || !old.lastUpdated || old.lastUpdated <= filesUploadedAt);
  let pre = wasPre ? oldCount : (isObj ? (old.pre || 0) : 0);
  const count = Math.max(0, oldCount + delta);
  if (pre > count) pre = count;
  if (count <= 0) return null;
  return { count, pre, lastUpdated: Date.now() };
}

// Ticks from all users, ignoring anything done before the latest PowerBI/backorder upload
function getFilteredAllUserTickCount(partId) {
  const key = safeKey(partId);
  let total = 0;
  Object.values(allUserTicks).forEach(userTicks => {
    if (userTicks) total += entryPostCount(userTicks[key]);
  });
  return total;
}

function getEffectiveRow(row) {
  const myTickCount = getTickCount(row['Part ID']);               // raw, for display
  const rawAllTickCount = getAllUserTickCount(row['Part ID']);    // raw, for display
  const rawBerCount = getBerCount(row['Part ID']);                // raw, for display
  // Quantity maths ignores work done before the latest PowerBI/backorder upload
  const allTickCount = getFilteredAllUserTickCount(row['Part ID']);
  const berCount = getFilteredBerCount(row['Part ID']);
  const manualAdj = stockAdjustments[safeKey(row['Part ID'])] || 0; // manual stock adjustment

  const origOpenOrders = typeof row['Open Orders'] === 'number' ? row['Open Orders'] : 0;
  const origTotalQty = typeof row['Total Qty'] === 'number' ? row['Total Qty'] : 0;
  const origRepReq = typeof row['Rep Req'] === 'number' ? row['Rep Req'] : 0;

  // Total Qty always subtracts ticks + BERs + manual adjustment
  const newTotalQty = Math.max(0, origTotalQty - allTickCount - berCount - manualAdj);

  // Rep Req and Open Orders only subtract ticks first (no manual adj yet)
  const newRepReq = Math.max(0, origRepReq - allTickCount);
  const newOpenOrders = row.Priority === 1 ? Math.max(0, origOpenOrders - allTickCount) : origOpenOrders;

  // Total Qty after ticks only (before BER and manual adj) — used as threshold
  const qtyAfterTicks = Math.max(0, origTotalQty - allTickCount);

  // Rep Req: follows manual adj only if it has caught up to Total Qty (after ticks)
  // BER reduces Rep Req only if Rep Req === Total Qty (strict 1:1)
  const repReqEqualsTotalQty = origRepReq === origTotalQty;
  const effRepReq = repReqEqualsTotalQty
    ? Math.max(0, newRepReq - berCount - manualAdj)   // exact match — follow everything
    : (newRepReq <= qtyAfterTicks - manualAdj
        ? Math.max(0, newRepReq - manualAdj)           // caught up after adj — follow adj only
        : newRepReq);                                   // still higher — don't adjust

  // Open Orders: only reduce with ticks (already done in newOpenOrders)
  // BER and manual adj only affect Open Orders if Open Orders === Total Qty (strict 1:1)
  const openOrdersEqualsTotalQty = row.Priority === 1 && origOpenOrders === origTotalQty;
  const effOpenOrders = row.Priority === 1
    ? (openOrdersEqualsTotalQty
        ? Math.max(0, newOpenOrders - berCount - manualAdj)  // exact match — follow everything
        : newOpenOrders)                                       // different — ticks only
    : newOpenOrders;

  const fullyDone = row.Priority === 1 ? effOpenOrders === 0 : effRepReq === 0;

  return { ...row, 'Open Orders': effOpenOrders, 'Total Qty': newTotalQty, 'Rep Req': effRepReq, _tickCount: myTickCount, _allTickCount: rawAllTickCount, _berCount: rawBerCount, _manualAdj: manualAdj, _fullyDone: fullyDone };
}

function updateDoneCount() {
  // Total individual ticks across all parts
  const totalTicks = Object.values(ticks).reduce((sum, v) => sum + (typeof v === 'object' ? (v.count || 0) : (v ? 1 : 0)), 0);
  // Total BERs — only current user's BERs
  const totalBers = Object.values(bers).reduce((sum, v) => sum + (typeof v === 'object' ? (v.count || 0) : 0), 0);
  document.getElementById('stat-done').textContent = totalTicks;
  document.getElementById('stat-ber-today').textContent = totalBers;
  updatePriorityTotals();
  // Completed tab badge: completed count, with BERs shown separately
  document.getElementById('tab-count-done').innerHTML = totalBers > 0 ? (totalTicks + ' <span style="opacity:0.8;">+ ' + totalBers + ' BER</span>') : String(totalTicks);
  // Refresh completed tab if active
  if (currentTab === 'done') renderCompleted();
}

// === COMPLETED TODAY & ADD PART MODAL ===
function switchTab(tab, btn) {
  currentTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');

  document.getElementById('actions-bar').style.display = 'none';
  document.getElementById('table-wrap').style.display = 'none';
  document.getElementById('zero-section').style.display = 'none';
  document.getElementById('completed-wrap').style.display = 'none';
  document.getElementById('history-wrap').style.display = 'none';

  if (tab === 'main') {
    document.getElementById('actions-bar').style.display = 'flex';
    document.getElementById('table-wrap').style.display = 'block';
    renderTable();
  } else if (tab === 'done') {
    document.getElementById('completed-wrap').style.display = 'block';
    renderCompleted();
  } else if (tab === 'history') {
    document.getElementById('history-wrap').style.display = 'block';
    // If historyData is empty, force a fresh read from Firebase
    if (Object.keys(historyData).length === 0 && HISTORY_REF) {
      HISTORY_REF.once('value', snap => {
        historyData = snap.val() || {};
        renderHistory();
      });
    } else {
      renderHistory();
    }
  }
}

function renderCompleted() {
  const tbody = document.getElementById('completed-body');
  tbody.innerHTML = '';
  // Parts from outputRows that have ticks/BERs
  const doneRows = outputRows.filter(r => getTickCount(r['Part ID']) > 0 || getBerCount(r['Part ID']) > 0);
  const includedPartIds = new Set(doneRows.map(r => r['Part ID']));

  // Also include manually added parts from today's history that aren't in outputRows
  if (historyData && historyData[TODAY]) {
    Object.values(historyData[TODAY]).forEach(entry => {
      if (!entry || !entry.partId) return;
      if (includedPartIds.has(entry.partId)) return; // already in list
      if ((entry.count || 0) === 0 && (entry.berCount || 0) === 0) return;
      // Add a synthetic row for this manual part
      doneRows.push({
        Priority: 2,
        'Part ID': entry.partId,
        'Repair Team': 'Manual',
        'Rep Req': 0,
        'Open Orders': '',
        'Total Qty': 0,
        'Part Class': entry.cls || '',
        Value: entry.val || 0,
        _manual: true
      });
      includedPartIds.add(entry.partId);
    });
  }

  // Final safety net: any part with a live tick/BER count not yet included (uses manualPartsMap for correct ID)
  Object.keys(manualPartsMap).forEach(key => {
    const m = manualPartsMap[key];
    if (includedPartIds.has(m.partId)) return;
    const tickCount = getTickCount(m.partId);
    const berCount = getBerCount(m.partId);
    if (tickCount === 0 && berCount === 0) return;
    doneRows.push({
      Priority: 2, 'Part ID': m.partId, 'Repair Team': 'Manual',
      'Rep Req': 0, 'Open Orders': '', 'Total Qty': 0,
      'Part Class': m.cls || '', Value: m.val || 0, _manual: true
    });
    includedPartIds.add(m.partId);
  });

  if (doneRows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" class="empty-done">No parts completed yet today</td></tr>';
    document.getElementById('earnings-bar').style.display = 'none';
    return;
  }

  let totalEarnings = 0;
  let totalTicks = 0;

  doneRows.forEach(row => {
    const tickCount = getTickCount(row['Part ID']);
    const berCount = getBerCount(row['Part ID']);
    const cls = row['Part Class'] || '';
    const val = typeof row.Value === 'number' ? row.Value : 0;
    const berVal = 0.10;
    const rowEarnings = (val * tickCount) + (berVal * berCount);
    totalEarnings += rowEarnings;
    totalTicks += tickCount;

    const valStr = val ? `£${val.toFixed(2)}` : '';
    const rowEarningsStr = rowEarnings ? `£${rowEarnings.toFixed(2)}` : '—';
    const partId = row['Part ID'];
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="badge badge-${row.Priority}">${row.Priority}</span></td>
      <td class="part-id">${partId}</td>
      <td>${row['Repair Team'] || ''}</td>
      <td class="num-cell">
        <div class="tick-counter">
          <button class="tick-minus" onclick="removeTick('${partId}')" ${tickCount === 0 ? 'disabled' : ''}>−</button>
          <span class="tick-num ${tickCount > 0 ? 'ticked' : ''}">${tickCount}</span>
          <button class="tick-plus" onclick="addTick('${partId}')">+</button>
        </div>
      </td>
      <td class="num-cell">
        <div class="ber-counter">
          <button class="ber-btn ber-minus" onclick="removeBer('${partId}')" ${berCount === 0 ? 'disabled' : ''}>−</button>
          <span class="ber-num ${berCount > 0 ? 'berred' : ''}">${berCount}</span>
          <button class="ber-btn" onclick="addBer('${partId}')">+</button>
        </div>
      </td>
      <td class="num-cell">${row['Open Orders'] !== '' ? row['Open Orders'] : '—'}</td>
      <td class="num-cell">${row['Total Qty'] !== '' ? row['Total Qty'] : '—'}</td>
      <td class="num-cell grade-${cls}">${cls}</td>
      <td class="num-cell">${valStr}</td>
      <td class="num-cell" style="color:var(--green);font-weight:600;">${rowEarningsStr}</td>`;
    tbody.appendChild(tr);
  });

  // Update earnings bar
  const bar = document.getElementById('earnings-bar');
  bar.style.display = 'flex';
  document.getElementById('earnings-total').textContent = `£${totalEarnings.toFixed(2)}`;
  document.getElementById('earnings-parts').textContent = doneRows.length;
  document.getElementById('earnings-ticks').textContent = totalTicks;
}

function openModal() {
  document.getElementById('modal-overlay').classList.add('open');
  document.getElementById('part-search-input').value = '';
  document.getElementById('part-dropdown').style.display = 'none';
  document.getElementById('modal-selected').style.display = 'none';
  document.getElementById('modal-qty-wrap').style.display = 'none';
  document.getElementById('modal-qty-input').value = 1;
  document.getElementById('modal-confirm-tick').style.display = 'none';
  document.getElementById('modal-confirm-ber').style.display = 'none';
  document.getElementById('part-dropdown').innerHTML = '';
  selectedPartId = null;
  setTimeout(() => document.getElementById('part-search-input').focus(), 100);
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('open');
  selectedPartId = null;
}

function searchParts(query) {
  const dropdown = document.getElementById('part-dropdown');
  if (!query || query.length < 1) { dropdown.style.display = 'none'; return; }
  const q = query.toUpperCase();
  // Search through ALL parts in gradeMap PLUS any ungraded DCD_PARTS
  const allParts = new Set([...(window.allPartIds || []), ...Object.keys(window.gradeMap || {}), ...DCD_PARTS]);
  const matches = [...allParts].filter(p => p.includes(q)).slice(0, 30);
  if (matches.length === 0) { dropdown.style.display = 'none'; return; }
  dropdown.innerHTML = '';
  matches.forEach(partId => {
    const cls = (window.gradeMap || {})[partId] || '';
    const val = GRADE_VALUE[cls] ? `£${GRADE_VALUE[cls].toFixed(2)}` : 'No grade';
    const div = document.createElement('div');
    div.className = 'part-option';
    div.innerHTML = `<span>${partId}</span><span class="part-option-grade">${cls ? cls + ' — ' + val : 'No grade'}</span>`;
    div.onclick = () => selectPart(partId, cls);
    dropdown.appendChild(div);
  });
  dropdown.style.display = 'block';
}

function selectPart(partId, cls) {
  selectedPartId = partId;
  const val = GRADE_VALUE[cls];
  document.getElementById('sel-part-id').textContent = partId;
  document.getElementById('sel-part-info').textContent = cls ? `Class: ${cls} — Value: £${val.toFixed(2)}` : 'No grade assigned';
  document.getElementById('modal-selected').style.display = 'flex';
  document.getElementById('modal-qty-wrap').style.display = 'block';
  document.getElementById('modal-qty-input').value = 1;
  document.getElementById('modal-confirm-tick').style.display = 'inline-block';
  document.getElementById('modal-confirm-ber').style.display = 'inline-block';
  document.getElementById('part-dropdown').style.display = 'none';
  document.getElementById('part-search-input').value = partId;
}

function confirmAddPart(type) {
  if (!selectedPartId) return;
  const cls = (window.gradeMap || {})[selectedPartId] || '';
  const val = typeof GRADE_VALUE[cls] === 'number' ? GRADE_VALUE[cls] : 0;
  const key = safeKey(selectedPartId);
  const qty = Math.max(1, Math.min(999, parseInt(document.getElementById('modal-qty-input').value, 10) || 1));

  // Remember this part permanently (survives outputRows being rebuilt by renderTable)
  manualPartsMap[key] = { partId: selectedPartId, cls, val };

  // Ensure part is in outputRows so saveToHistory can pick it up
  if (!outputRows.find(r => r['Part ID'] === selectedPartId)) {
    outputRows.push({
      Priority: 2, 'Part ID': selectedPartId, 'Repair Team': 'Manual',
      'Rep Req': 0, 'Open Orders': '', 'Total Qty': 0,
      'Part Class': cls, Value: val, _manual: true
    });
  }

  if (type === 'tick') {
    // Only write to TICKS_REF — saveToHistory will handle history automatically
    const currentTick = getTickCount(selectedPartId);
    TICKS_REF.child(key).transaction(t => nextEntry(t, qty));
  } else if (type === 'ber') {
    // Only write to BER_REF — saveToHistory will handle history automatically
    const currentBer = getBerCount(selectedPartId);
    BER_REF.child(key).transaction(b => nextEntry(b, qty));
  }

  closeModal();
  // Immediately re-render the current tab so the part appears without needing to switch tabs
  if (currentTab === 'done') renderCompleted();
  else if (currentTab === 'history') renderHistory();
}

// Close modal on overlay click
document.addEventListener('click', e => {
  if (e.target.id === 'modal-overlay') closeModal();
});

// === HISTORY LOG & KPI CSV IMPORTER ===
function updateHistoryCount() {
  // Only count repairs for the current month
  const currentMonth = TODAY.slice(0, 7); // YYYY-MM
  let monthTotal = 0;
  Object.keys(historyData).forEach(date => {
    if (date === '_meta') return;
    if (!date.startsWith(currentMonth)) return;
    const day = historyData[date];
    if (!day) return;
    Object.values(day).forEach(entry => { monthTotal += entry.count || 0; });
  });
  document.getElementById('tab-count-history').textContent = monthTotal;
}

function saveToHistory() {
  if (!HISTORY_REF || outputRows.length === 0) return;
  const dayRef = HISTORY_REF.child(TODAY);

  // Use transactions so we never accidentally overwrite a manually decremented value
  // Only update if tick count has INCREASED or entry doesn't exist yet
  Object.keys(ticks).forEach(key => {
    const tick = ticks[key];
    const count = typeof tick === 'object' ? (tick.count || 0) : 0;
    if (count === 0) {
      // Remove from history if ticks go to 0
      dayRef.child(key).remove();
      return;
    }
    const row = outputRows.find(r => safeKey(r['Part ID']) === key);
    const manual = manualPartsMap[key];
    const partId = row ? row['Part ID'] : (manual ? manual.partId : key);
    const cls = row ? (row['Part Class'] || '') : (manual ? manual.cls : '');
    const val = row ? (typeof row.Value === 'number' ? row.Value : 0) : (manual ? manual.val : 0);
    const berCount = getBerCount(partId);
    const berEarned = berCount * 0.10;
    dayRef.child(key).transaction(existing => {
      // Only set if not exists, or count has increased
      if (!existing) {
        return { partId, count, cls, val, earned: (val * count) + berEarned, berCount, berEarned, type: 'tick' };
      }
      // Always sync to current tick count (handles both increases and decreases)
      existing.count = count;
      existing.berCount = berCount;
      existing.berEarned = berEarned;
      existing.earned = (val * count) + berEarned;
      return existing;
    });
  });

  // Save BER-only parts
  Object.keys(bers).forEach(key => {
    if (ticks[key]) return;
    const ber = bers[key];
    const berCount = typeof ber === 'object' ? (ber.count || 0) : 0;
    if (berCount === 0) {
      dayRef.child('ber_' + key).remove();
      return;
    }
    const row = outputRows.find(r => safeKey(r['Part ID']) === key);
    const manual = manualPartsMap[key];
    const partId = row ? row['Part ID'] : (manual ? manual.partId : key);
    const cls = row ? (row['Part Class'] || '') : (manual ? manual.cls : '');
    dayRef.child('ber_' + key).transaction(existing => {
      if (!existing) {
        return { partId, count: 0, cls, val: 0, earned: berCount * 0.10, berCount, berEarned: berCount * 0.10, type: 'ber' };
      }
      existing.berCount = berCount;
      existing.berEarned = berCount * 0.10;
      existing.earned = (existing.val || 0) * (existing.count || 0) + existing.berEarned;
      return existing;
    });
  });

  // Cleanup: remove today's history entries that no longer have any ticks or BERs
  if (historyData && historyData[TODAY]) {
    Object.keys(historyData[TODAY]).forEach(histKey => {
      const isBerEntry = histKey.startsWith('ber_');
      const tickKey = isBerEntry ? histKey.slice(4) : histKey;
      const hasTick = ticks[tickKey] && (typeof ticks[tickKey] === 'object' ? ticks[tickKey].count > 0 : ticks[tickKey]);
      const hasBer = bers[tickKey] && (typeof bers[tickKey] === 'object' ? bers[tickKey].count > 0 : false);
      if (isBerEntry && !hasBer) {
        dayRef.child(histKey).remove();
      } else if (!isBerEntry && !hasTick && !hasBer) {
        dayRef.child(histKey).remove();
      }
    });
  }
}

function decrementHistoryEntry(date, key, currentCount, currentBer) {
  if (!key) return;
  const newCount = currentCount - 1;
  // Update history
  if (newCount <= 0 && currentBer <= 0) {
    HISTORY_REF.child(date).child(key).remove();
  } else {
    HISTORY_REF.child(date).child(key).transaction(entry => {
      if (!entry) return entry;
      entry.count = Math.max(0, (entry.count || 0) - 1);
      entry.earned = ((entry.val || 0) * entry.count) + (entry.berEarned || 0);
      return entry;
    });
  }
  // Sync to ticks/BER refs if today
  if (date === TODAY && !key.startsWith('ber_')) {
    TICKS_REF.child(key).transaction(t => nextEntry(t, -1));
  }
}

function incrementHistoryEntry(date, key) {
  if (!key) return;
  HISTORY_REF.child(date).child(key).transaction(entry => {
    if (!entry) return entry;
    entry.count = (entry.count || 0) + 1;
    entry.earned = ((entry.val || 0) * entry.count) + (entry.berEarned || 0);
    return entry;
  });
  // Sync to ticks ref if today
  if (date === TODAY && !key.startsWith('ber_')) {
    TICKS_REF.child(key).transaction(t => nextEntry(t, 1));
  }
}

function removeBerFromHistory(date, key, currentCount, currentBer) {
  if (!key) return;
  const newBer = currentBer - 1;
  if (newBer <= 0 && currentCount <= 0) {
    HISTORY_REF.child(date).child(key).remove();
  } else {
    HISTORY_REF.child(date).child(key).transaction(entry => {
      if (!entry) return entry;
      entry.berCount = Math.max(0, (entry.berCount || 0) - 1);
      entry.berEarned = entry.berCount * 0.10;
      entry.earned = ((entry.val || 0) * (entry.count || 0)) + entry.berEarned;
      return entry;
    });
  }
  // Sync to BER ref if today
  if (date === TODAY) {
    const berKey = key.startsWith('ber_') ? key.slice(4) : key;
    BER_REF.child(berKey).transaction(b => nextEntry(b, -1));
  }
}

function addBerToHistory(date, key) {
  if (!key) return;
  HISTORY_REF.child(date).child(key).transaction(entry => {
    if (!entry) return entry;
    entry.berCount = (entry.berCount || 0) + 1;
    entry.berEarned = entry.berCount * 0.10;
    entry.earned = ((entry.val || 0) * (entry.count || 0)) + entry.berEarned;
    return entry;
  });
  // Sync to BER ref if today
  if (date === TODAY) {
    const berKey = key.startsWith('ber_') ? key.slice(4) : key;
    BER_REF.child(berKey).transaction(b => nextEntry(b, 1));
  }
}

function switchHistorySub(sub, btn) {
  document.querySelectorAll('.hist-sub-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('hist-history-section').style.display = sub === 'history' ? 'block' : 'none';
  document.getElementById('hist-import-section').style.display = sub === 'import' ? 'block' : 'none';
  if (sub === 'history') renderHistory();
}

function handleImportFile(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try { parseImportCSV(e.target.result); }
    catch(err) { showImportError('Error reading file: ' + err.message); }
  };
  reader.readAsText(file, 'utf-8');
}

function parseImportCSV(text) {
  text = text.replace(/^\uFEFF/, '');
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) { showImportError('CSV must have a header row and at least one data row.'); return; }

  const headers = lines[0].split(',').map(h => h.trim().toLowerCase().replace(/^"|"$/g, ''));
  const isKpiFormat = headers.includes('unit_status') && headers.includes('createdtime');

  let dateIdx, partIdx, qtyIdx, statusIdx;
  if (isKpiFormat) {
    dateIdx   = headers.indexOf('createdtime');
    partIdx   = headers.indexOf('part_id');
    qtyIdx    = headers.indexOf('quantity');
    statusIdx = headers.indexOf('unit_status');
  } else {
    dateIdx   = headers.findIndex(h => h === 'date');
    partIdx   = headers.findIndex(h => h.includes('part'));
    qtyIdx    = headers.findIndex(h => h.includes('repair') || h === 'repairs' || h === 'count');
    statusIdx = -1;
  }

  if (dateIdx === -1 || partIdx === -1) {
    showImportError('Could not find required columns. Expected "part_id" and "CreatedTime" (KPI format) or "date" and "part_id" (custom format).');
    return;
  }

  const grouped = {};
  const errors = [];
  const seenTags = new Set();

  lines.slice(1).forEach((line, i) => {
    if (!line.trim()) return;
    const cols = line.split(',').map(c => c.trim().replace(/^"|"$/g, ''));
    const rawDate = (cols[dateIdx] || '').trim();
    const partId  = (cols[partIdx] || '').trim();
    const qty     = parseInt(cols[qtyIdx]) || 0;
    const status  = statusIdx !== -1 ? (cols[statusIdx] || 'COMP').trim().toUpperCase() : 'COMP';
    if (!partId) return;

    let isoDate = '';
    const ddmmyyyy = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    const yyyymmdd = /^(\d{4})-(\d{2})-(\d{2})$/;
    if (ddmmyyyy.test(rawDate)) {
      const m = rawDate.match(ddmmyyyy);
      isoDate = m[3] + '-' + m[2] + '-' + m[1];
    } else if (yyyymmdd.test(rawDate)) {
      isoDate = rawDate;
    } else {
      errors.push('Row ' + (i+2) + ': unrecognised date "' + rawDate + '"');
      return;
    }

    // Treat NFF same as COMP
    const normStatus = (status === 'NFF') ? 'COMP' : status;
    // Use repair_tag to deduplicate — each unique tag = 1 job
    const repairTag = (cols[headers.indexOf('repair_tag')] || '').trim();
    const dedupKey = repairTag ? ('tag_' + repairTag) : (isoDate + '__' + partId + '__' + normStatus + '__' + i);
    if (seenTags.has(dedupKey)) return; // skip duplicate repair tags
    seenTags.add(dedupKey);
    // Skip weekends
    const dayOfWeek = new Date(isoDate + 'T12:00:00').getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) return; // 0=Sunday, 6=Saturday

    const key = isoDate + '__' + partId + '__' + normStatus;
    if (!grouped[key]) grouped[key] = { date: isoDate, partId, status: normStatus, qty: 0 };
    grouped[key].qty += 1; // each unique repair tag = 1 job
  });

  importData = Object.values(grouped).map(g => {
    const isComp = g.status === 'COMP';
    const cls = (window.gradeMap || {})[g.partId] || '';
    const val = typeof GRADE_VALUE[cls] === 'number' ? GRADE_VALUE[cls] : 0;
    const repairs = isComp ? g.qty : 0;
    const ber = isComp ? 0 : g.qty;
    const earned = (val * repairs) + (ber * 0.10);
    return { date: g.date, partId: g.partId, repairs, ber, cls, val, earned };
  }).filter(r => r.repairs > 0 || r.ber > 0)
    .sort((a, b) => a.date.localeCompare(b.date) || a.partId.localeCompare(b.partId));

  if (importData.length === 0) {
    showImportError('No valid data rows found. ' + (errors.length ? errors[0] : ''));
    return;
  }
  document.getElementById('import-error').style.display = 'none';
  renderImportPreview(errors);
}

function renderImportPreview(errors) {
  const preview = document.getElementById('import-preview');
  const tbody = document.getElementById('import-preview-body');
  preview.style.display = 'block';
  document.getElementById('import-preview-title').textContent = 'Preview — ' + importData.length + ' entries';
  document.getElementById('import-preview-sub').textContent = errors.length ? errors.length + ' rows skipped due to errors' : 'All rows valid';
  tbody.innerHTML = '';
  importData.forEach((r, i) => {
    const bg = i % 2 === 0 ? 'var(--surface)' : 'var(--surface2)';
    const d = new Date(r.date + 'T12:00:00');
    const dateLabel = d.toLocaleDateString('en-GB', { weekday:'short', day:'numeric', month:'short', year:'numeric' });
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';font-size:0.7rem;color:var(--muted);">' + dateLabel + '</td>' +
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';font-weight:600;">' + r.partId + '</td>' +
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;color:var(--green);font-weight:700;">' + (r.repairs || '—') + '</td>' +
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;color:#c00000;font-weight:700;">' + (r.ber || '—') + '</td>' +
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;" class="grade-' + r.cls + '">' + (r.cls || '—') + '</td>' +
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;color:var(--green);">' + (r.earned ? '£' + r.earned.toFixed(2) : '—') + '</td>' +
      '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;font-size:0.65rem;color:var(--green);">✓ Ready</td>';
    tbody.appendChild(tr);
  });
}

function confirmImport() {
  if (!importData.length || !currentUser) return;

  // Group importData by date+partId to get definitive counts from this file
  // Use SET (not ADD) so re-importing the same file doesn't duplicate
  importData.forEach(r => {
    const key = safeKey(r.partId);
    const histKey = r.repairs > 0 ? key : ('ber_' + key);
    HISTORY_REF.child(r.date).child(histKey).transaction(entry => {
      if (!entry) {
        // New entry - set from import
        return { partId: r.partId, count: r.repairs, cls: r.cls, val: r.val, earned: r.earned, berCount: r.ber, berEarned: r.ber * 0.10, type: r.repairs > 0 ? 'tick' : 'ber' };
      }
      // Existing entry — only update if import count is HIGHER (prevents duplicating on re-import)
      const newCount = Math.max(entry.count || 0, r.repairs);
      const newBer = Math.max(entry.berCount || 0, r.ber);
      const newBerEarned = newBer * 0.10;
      const val = entry.val || r.val;
      return {
        ...entry,
        count: newCount,
        berCount: newBer,
        berEarned: newBerEarned,
        earned: (val * newCount) + newBerEarned
      };
    });
  });

  const importedCount = importData.length;
  clearImportPreview();
  // Reset month filter to all so imported data is visible
  activeMonthFilter = 'all';
  alert('✓ ' + importedCount + ' entries imported successfully!');
  switchHistorySub('history', document.getElementById('sub-history'));
  // Force re-render after Firebase transactions settle
  setTimeout(() => {
    if (HISTORY_REF) {
      HISTORY_REF.once('value', snap => {
        historyData = snap.val() || {};
        renderHistory();
        updateHistoryCount();
      });
    }
  }, 1000);
}

function clearImportPreview() {
  importData = [];
  document.getElementById('import-preview').style.display = 'none';
  document.getElementById('import-file-input').value = '';
}

function showImportError(msg) {
  const el = document.getElementById('import-error');
  el.innerHTML = '⚠ ' + msg;
  el.style.display = 'block';
}

// === COMMISSION TRACKER & MONTHLY TARGET CALCULATIONS ===
function setMonthFilter(month, btn) {
  activeMonthFilter = month;
  activeGradeFilter = 'all'; // reset grade filter on month change
  document.querySelectorAll('.month-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderHistory();
}

function buildMonthButtons() {
  const bar = document.getElementById('month-filter-bar');
  if (!bar) return;
  const months = new Set();
  Object.keys(historyData).forEach(date => {
    if (date === '_meta') return;
    months.add(date.slice(0, 7));
  });
  let btns = '<button class="month-btn all-btn' + (activeMonthFilter === 'all' ? ' active' : '') + '" data-month="all" onclick="setMonthFilter(this.dataset.month, this)">All Time</button>';
  [...months].sort().reverse().forEach(m => {
    const d = new Date(m + '-01');
    const label = d.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' });
    const isActive = activeMonthFilter === m;
    btns += '<button class="month-btn' + (isActive ? ' active' : '') + '" data-month="' + m + '" onclick="setMonthFilter(this.dataset.month, this)">' + label + '</button>';
  });
  bar.innerHTML = btns;
}

function updateMonthSummary(filteredEntries, monthKey, passedDays) {
  const card = document.getElementById('month-summary-card');
  if (!card) return;
  if (monthKey === 'all' || filteredEntries.length === 0) {
    card.style.display = 'none';
    const bd = document.getElementById('ms-grade-breakdown');
    if (bd) bd.style.display = 'none';
    return;
  }
  card.style.display = 'grid';
  const repairs = filteredEntries.reduce((s, e) => s + (e.count || 0), 0);
  const earned = filteredEntries.reduce((s, e) => s + (e.earned || 0), 0);
  // Use passedDays (adjusted) if provided, otherwise count from entries
  const rawDays = new Set(filteredEntries.filter(e => (e.count||0)>0||(e.berCount||0)>0).map(e=>e.date)).size;
  const workDays = passedDays !== undefined ? passedDays : rawDays;
  const avg = workDays > 0 ? repairs / workDays : 0;
  const commission = calcCommission(avg, monthKey);
  const d = new Date(monthKey + '-01');
  const monthLabel = d.toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
  document.getElementById('ms-month').textContent = monthLabel;
  document.getElementById('ms-repairs').textContent = repairs;
  document.getElementById('ms-days').textContent = workDays % 1 === 0 ? workDays : workDays.toFixed(1);
  document.getElementById('ms-avg').textContent = avg.toFixed(1);
  document.getElementById('ms-earned').textContent = '£' + earned.toFixed(2);
  document.getElementById('ms-commission-badge').textContent = commission.label;
  document.getElementById('ms-commission-badge').style.background = commission.color;
  document.getElementById('ms-commission').textContent = commission.amount > 0 ? '+£' + commission.amount : '£0';
  document.getElementById('ms-commission').style.color = commission.amount > 0 ? 'var(--green)' : 'var(--muted)';
  document.getElementById('ms-total').textContent = '£' + (earned + commission.amount).toFixed(2);

  // Grade breakdown
  const gradeOrder = ['X', 'S', 'C', 'B', 'A', 'M', 'BER', '?'];
  const gradeColors = { X: '#1a7a3a', S: '#4caf50', C: '#f0d040', B: '#b08800', A: '#e0a020', M: '#e07020', BER: '#c00000', '?': '#666666' };
  const gradeValues = { X: 3.00, S: 1.25, C: 0.75, B: 0.50, A: 0.25, M: 0.10, BER: 0.10 };
  const gradeCounts = {};
  filteredEntries.forEach(function(e) {
    const cls = (e.cls || '').toUpperCase();
    const count = e.count || 0;
    if (count === 0) return;
    if (cls) {
      gradeCounts[cls] = (gradeCounts[cls] || 0) + count;
    } else {
      gradeCounts['?'] = (gradeCounts['?'] || 0) + count; // ungraded
    }
  });
  const berTotal = filteredEntries.reduce(function(s, e) { return s + (e.berCount || 0); }, 0);
  if (berTotal > 0) gradeCounts['BER'] = berTotal;

  const breakdown = document.getElementById('ms-grade-breakdown');
  const barsEl = document.getElementById('ms-grade-bars');
  if (!breakdown || !barsEl) return;
  const hasGrades = Object.keys(gradeCounts).length > 0;
  breakdown.style.display = hasGrades ? 'block' : 'none';
  if (hasGrades) {
    const maxCount = Math.max.apply(null, Object.values(gradeCounts));
    let barsHTML = '';
    gradeOrder.filter(function(g) { return gradeCounts[g] > 0; }).forEach(function(g) {
      const count = gradeCounts[g];
      const barH = Math.max(20, Math.round((count / maxCount) * 80));
      const gradeEarned = (count * (gradeValues[g] || 0)).toFixed(2);
      const color = gradeColors[g];
      const isGradeActive = activeGradeFilter === g;
      const barOpacity = activeGradeFilter === 'all' ? '0.85' : (isGradeActive ? '1' : '0.25');
      const barOutline = isGradeActive ? ';outline:2px solid ' + color + ';outline-offset:2px;' : '';
      barsHTML += '<div data-grade="' + g + '" onclick="filterByGrade(this.dataset.grade)" style="display:flex;flex-direction:column;align-items:center;gap:4px;min-width:48px;cursor:pointer;transition:opacity 0.2s;" title="Filter by grade ' + g + '">';
      barsHTML += '<div style="font-size:0.65rem;color:var(--muted);">£' + gradeEarned + '</div>';
      barsHTML += '<div style="font-size:0.75rem;font-weight:700;color:' + color + ';">' + count + '</div>';
      barsHTML += '<div style="width:40px;height:' + barH + 'px;background:' + color + ';border-radius:3px 3px 0 0;opacity:' + barOpacity + barOutline + '"></div>';
      barsHTML += '<div style="font-family:Syne,sans-serif;font-weight:800;font-size:0.75rem;color:' + color + (isGradeActive ? ';text-decoration:underline' : '') + ';">' + g + '</div>';
      barsHTML += '</div>';
    });
    barsEl.innerHTML = barsHTML;
  }
}

function calcCommission(avgPerDay, monthKey) {
  // New targets from July 2026 onwards (11/13/15), old targets before (18/21/24)
  const m = monthKey || TODAY.slice(0, 7);
  const isNewTargets = m >= '2026-07';
  if (isNewTargets) {
    if (avgPerDay >= 15) return { label: '15/day ✓', amount: 75, color: '#2ecc71' };
    if (avgPerDay >= 13) return { label: '13/day ✓', amount: 50, color: '#f0a040' };
    if (avgPerDay >= 11) return { label: '11/day ✓', amount: 25, color: '#f0c040' };
  } else {
    if (avgPerDay >= 24) return { label: '24/day ✓', amount: 75, color: '#2ecc71' };
    if (avgPerDay >= 21) return { label: '21/day ✓', amount: 50, color: '#f0a040' };
    if (avgPerDay >= 18) return { label: '18/day ✓', amount: 25, color: '#f0c040' };
  }
  return { label: 'Not reached', amount: 0, color: 'var(--border)' };
}

function updateTargetTracker(totalRepairs, adjustedDays) {
  const tracker = document.getElementById('target-tracker');

  // Hide on All Time view
  if (activeMonthFilter === 'all' || adjustedDays <= 0 || totalRepairs <= 0) {
    tracker.style.display = 'none';
    return;
  }
  tracker.style.display = 'block';
  // Also ensure grade breakdown is visible
  const bd = document.getElementById('ms-grade-breakdown');
  if (bd) bd.style.display = 'block';

  const currentAvg = totalRepairs / adjustedDays;
  // New targets from July 2026 onwards (11/day, 13/day, 15/day)
  // Old targets before July 2026 (18/day, 21/day, 24/day)
  const filterMonth = activeMonthFilter === 'all' ? TODAY.slice(0, 7) : activeMonthFilter;
  const isNewTargets = filterMonth >= '2026-07';

  const targets = isNewTargets
    ? [
        { id: '18', target: 11, bonus: 25 },
        { id: '21', target: 13, bonus: 50 },
        { id: '24', target: 15, bonus: 75 }
      ]
    : [
        { id: '18', target: 18, bonus: 25 },
        { id: '21', target: 21, bonus: 50 },
        { id: '24', target: 24, bonus: 75 }
      ];

  // Update label text to match correct targets
  if (isNewTargets) {
    document.getElementById('label-18').innerHTML = '11/day <span style="color:#f0c040;font-weight:700;">+£25</span>';
    document.getElementById('label-21').innerHTML = '13/day <span style="color:#f0a040;font-weight:700;">+£50</span>';
    document.getElementById('label-24').innerHTML = '15/day <span style="color:#2ecc71;font-weight:700;">+£75</span>';
  } else {
    document.getElementById('label-18').innerHTML = '18/day <span style="color:#f0c040;font-weight:700;">+£25</span>';
    document.getElementById('label-21').innerHTML = '21/day <span style="color:#f0a040;font-weight:700;">+£50</span>';
    document.getElementById('label-24').innerHTML = '24/day <span style="color:#2ecc71;font-weight:700;">+£75</span>';
  }

  // Check if this is the current month or a past month
  const currentMonth = TODAY.slice(0, 7);
  const isPastMonth = activeMonthFilter < currentMonth;

  const now = new Date();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const workingDaysInMonth = Math.round(daysInMonth * 22 / 30);
  const remainingDays = Math.max(0, workingDaysInMonth - adjustedDays);

  targets.forEach(t => {
    // Bar fills based on actual avg vs target — capped at 100%
    const pct = Math.min(100, (currentAvg / t.target) * 100);
    const bar = document.getElementById('bar-' + t.id);
    const pctEl = document.getElementById('pct-' + t.id);
    const statusEl = document.getElementById('status-' + t.id);

    bar.style.width = pct.toFixed(1) + '%';
    pctEl.textContent = pct.toFixed(0) + '%';

    if (currentAvg >= t.target) {
      statusEl.textContent = isPastMonth ? '✓ Achieved' : '✓ On track';
      statusEl.style.color = 'var(--green)';
    } else if (isPastMonth) {
      // Past month — show how far short they were
      statusEl.textContent = (t.target - currentAvg).toFixed(1) + ' short';
      statusEl.style.color = 'var(--accent)';
    } else {
      // Current month — show extra needed per day
      const totalNeeded = t.target * workingDaysInMonth;
      const extraNeeded = Math.max(0, totalNeeded - totalRepairs);
      const extraPerDay = remainingDays > 0 ? (extraNeeded / remainingDays).toFixed(1) : '∞';
      statusEl.textContent = '+' + extraPerDay + '/day';
      statusEl.style.color = 'var(--accent)';
    }
  });
}

function formatHistoryDate(dateStr) {
  if (!dateStr) return dateStr;
  const d = new Date(dateStr + 'T12:00:00');
  const day = d.toLocaleDateString('en-GB', { weekday: 'short' });
  const formatted = d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  const isToday = dateStr === TODAY;
  return day + ', ' + formatted + (isToday ? ' · Today' : '');
}

function getDateRowStyle(dateStr) {
  const d = new Date(dateStr + 'T12:00:00');
  const isWeekend = d.getDay() === 0 || d.getDay() === 6;
  const isToday = dateStr === TODAY;
  if (isToday) return 'color: var(--green);';
  if (isWeekend) return 'color: var(--accent);';
  return 'color: var(--muted);';
}

function filterByGrade(grade) {
  activeGradeFilter = (activeGradeFilter === grade) ? 'all' : grade;
  renderHistory();
  if (activeGradeFilter === '?') {
    const sec = document.getElementById('ungraded-section');
    if (sec && sec.style.display !== 'none') setTimeout(() => sec.scrollIntoView({ behavior: 'smooth', block: 'center' }), 50);
  }
}

function renderHistory() {
  const tbody = document.getElementById('history-body');
  tbody.innerHTML = '';

  const allEntries = [];
  Object.keys(historyData).sort().reverse().forEach(date => {
    if (date === '_meta') return;
    if (activeMonthFilter !== 'all' && !date.startsWith(activeMonthFilter)) return;
    const day = historyData[date];
    if (!day) return;
    Object.keys(day).forEach(key => {
      const entry = day[key];
      // Apply grade filter if active
      if (activeGradeFilter !== 'all') {
        if (activeGradeFilter === 'BER') {
          if (!entry.berCount || entry.berCount === 0) return;
        } else if (activeGradeFilter === '?') {
          if ((entry.cls || '').toUpperCase() !== '') return; // show only ungraded
        } else {
          if ((entry.cls || '').toUpperCase() !== activeGradeFilter) return;
        }
      }
      allEntries.push({ date, ...entry, isBerOnly: entry.type === 'ber', _key: key });
    });
  });

  // Show grade filter indicator
  const gradeFilterEl = document.getElementById('grade-filter-indicator');
  if (gradeFilterEl) {
    if (activeGradeFilter !== 'all') {
      gradeFilterEl.textContent = 'Filtering by Grade ' + activeGradeFilter + ' — click bar again to clear';
      gradeFilterEl.style.display = 'block';
    } else {
      gradeFilterEl.style.display = 'none';
    }
  }

  // Build month filter buttons
  buildMonthButtons();

  if (allEntries.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="empty-done">No history yet — complete some parts to start tracking</td></tr>`;
    document.getElementById('hist-total-parts').textContent = '0';
    document.getElementById('hist-total-repairs').textContent = '0';
    document.getElementById('hist-total-earned').textContent = '£0.00';
    return;
  }

  let totalRepairs = 0, totalEarned = 0;

  allEntries.forEach((e, i) => {
    const bg = i % 2 === 0 ? 'var(--surface)' : 'var(--surface2)';
    const earned = e.earned || (e.val * e.count) || 0;
    const entryKey = e._key || '';
    // Only count as repair if not a BER-only entry
    if (!e.isBerOnly) totalRepairs += e.count || 0;
    totalEarned += earned;
    const cls = e.cls || '';
    const berCount = e.berCount || 0;
    const tickLabel = e.count > 0 ? `<span style="color:var(--green);font-weight:700;">${e.count}</span>` : '—';
    const berLabel = berCount > 0 ? `<span style="color:#c00000;font-weight:700;">${berCount} BER</span>` : '—';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};font-size:0.7rem;${getDateRowStyle(e.date)}">${formatHistoryDate(e.date)}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};font-weight:600;">${e.partId || ''}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};text-align:center;">${tickLabel}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};text-align:center;">${berLabel}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};text-align:center;" class="grade-${cls}">${cls}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};text-align:center;">${e.val ? '£' + Number(e.val).toFixed(2) : '—'}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};text-align:center;color:var(--green);font-weight:600;">${earned ? '£' + earned.toFixed(2) : '—'}</td>
      <td style="padding:0.4rem 0.8rem;border-bottom:1px solid rgba(255,255,255,0.03);background:${bg};text-align:center;">
        <div style="display:flex;flex-direction:column;gap:3px;align-items:center;">
          <div style="display:flex;flex-direction:column;gap:4px;align-items:center;">
          <div style="display:flex;align-items:center;gap:3px;">
            <button onclick="decrementHistoryEntry('${e.date}','${entryKey}',${e.count},${berCount})" style="font-family:'Syne',sans-serif;font-weight:700;font-size:0.65rem;width:20px;height:20px;border-radius:2px;border:1px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;line-height:1;" ${e.count === 0 ? 'disabled' : ''}>−</button>
            <span style="font-size:0.75rem;font-weight:700;color:var(--green);min-width:16px;text-align:center;">${e.count}</span>
            <button onclick="incrementHistoryEntry('${e.date}','${entryKey}')" style="font-family:'Syne',sans-serif;font-weight:700;font-size:0.65rem;width:20px;height:20px;border-radius:2px;border:1px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;line-height:1;">+</button>
            <span style="font-size:0.6rem;color:var(--muted);">rep</span>
          </div>
          <div style="display:flex;align-items:center;gap:3px;">
            <button onclick="removeBerFromHistory('${e.date}','${entryKey}',${e.count},${berCount})" style="font-family:'Syne',sans-serif;font-weight:700;font-size:0.65rem;width:20px;height:20px;border-radius:2px;border:1px solid #c00000;background:transparent;color:#c00000;cursor:pointer;line-height:1;" ${berCount === 0 ? 'disabled' : ''}>−</button>
            <span style="font-size:0.75rem;font-weight:700;color:#c00000;min-width:16px;text-align:center;">${berCount}</span>
            <button onclick="addBerToHistory('${e.date}','${entryKey}')" style="font-family:'Syne',sans-serif;font-weight:700;font-size:0.65rem;width:20px;height:20px;border-radius:2px;border:1px solid #c00000;background:#c00000;color:white;cursor:pointer;line-height:1;">+</button>
          </div>
        </div>
      </td>`;
    tbody.appendChild(tr);
  });

  // Only count days where at least one entry has actual work (repairs or BERs > 0)
  const daysWithWork = new Set();
  Object.keys(historyData).forEach(date => {
    if (date === '_meta') return; // skip metadata
    const day = historyData[date];
    if (!day) return;
    const hasWork = Object.values(day).some(entry =>
      (entry.count || 0) > 0 || (entry.berCount || 0) > 0
    );
    if (hasWork) daysWithWork.add(date);
  });
  baseDaysWorked = daysWithWork.size;

  // Load per-month adjustments from Firebase metadata (merge, don't overwrite)
  const metaAdj = (historyData['_meta'] && historyData['_meta'].adj) ? historyData['_meta'].adj : {};
  if (typeof daysAdjustment !== 'object') daysAdjustment = {};
  // Only merge from Firebase - don't reset local changes
  Object.keys(metaAdj).forEach(k => { if (daysAdjustment[k] === undefined) daysAdjustment[k] = metaAdj[k]; });

  // For a specific month, count only that month's worked days
  let adjustedDays;
  if (activeMonthFilter !== 'all') {
    const monthWorkedDays = new Set(
      allEntries.filter(e => (e.count||0) > 0 || (e.berCount||0) > 0).map(e => e.date)
    ).size;
    monthBaseDays[activeMonthFilter] = monthWorkedDays;
    const monthAdj = daysAdjustment[activeMonthFilter] || 0;
    adjustedDays = Math.max(monthWorkedDays > 0 ? 0.5 : 0, monthWorkedDays + monthAdj);
  } else {
    const allAdj = daysAdjustment['alltime'] || 0;
    adjustedDays = Math.max(baseDaysWorked > 0 ? 0.5 : 0, baseDaysWorked + allAdj);
  }
  const avgRepairs = adjustedDays > 0 ? (totalRepairs / adjustedDays).toFixed(1) : '0';
  document.getElementById('hist-total-parts').textContent = allEntries.length;
  document.getElementById('hist-total-repairs').textContent = totalRepairs;
  document.getElementById('hist-total-earned').textContent = `£${totalEarned.toFixed(2)}`;
  // Show diff indicator if days have been adjusted
  const adjStr = adjustedDays % 1 === 0 ? adjustedDays : adjustedDays.toFixed(1);
  const _mk = activeMonthFilter === 'all' ? 'alltime' : activeMonthFilter;
  const currentAdj = (typeof daysAdjustment === 'object') ? (daysAdjustment[_mk] || 0) : 0;
  let diffStr = '';
  if (currentAdj !== 0) {
    const sign = currentAdj > 0 ? '+' : '';
    const diffVal = currentAdj % 1 === 0 ? currentAdj : currentAdj.toFixed(1);
    diffStr = ` <span style="font-size:0.7rem;color:${currentAdj > 0 ? 'var(--accent)' : 'var(--green)'};">(${sign}${diffVal})</span>`;
  }
  document.getElementById('hist-days-worked').textContent = adjStr;
  const adjInd = document.getElementById('days-adj-indicator');
  if (adjInd) {
    // Strip HTML tags for the indicator, just show (+0.5) or (-0.5)
    const adjMatch = diffStr.match(/[+-][0-9.]+/);
    adjInd.textContent = adjMatch ? adjMatch[0] + ' day adj.' : '';
    adjInd.style.color = currentAdj > 0 ? 'var(--green)' : currentAdj < 0 ? 'var(--accent)' : '';
  }
  document.getElementById('hist-avg-repairs').textContent = avgRepairs;
  updateMonthSummary(allEntries, activeMonthFilter, activeMonthFilter !== 'all' ? adjustedDays : undefined);
  updateTargetTracker(totalRepairs, adjustedDays);

  // Ungraded section — show parts with no grade, regardless of month/grade filter
  const knownGradesSet = new Set(['X','S','C','B','A','M']);
  const ungradedEntries = [];
  Object.keys(historyData).sort().reverse().forEach(date => {
    if (date === '_meta') return;
    if (activeMonthFilter !== 'all' && !date.startsWith(activeMonthFilter)) return;
    const day = historyData[date];
    if (!day) return;
    Object.keys(day).forEach(key => {
      const e = day[key];
      const cls = (e.cls || '').toUpperCase();
      if (!cls || !knownGradesSet.has(cls)) {
        if ((e.count || 0) > 0) ungradedEntries.push({ date, ...e, _key: key });
      }
    });
  });

  const ungradedSection = document.getElementById('ungraded-section');
  const ungradedTbody = document.getElementById('ungraded-tbody');
  if (ungradedSection && ungradedTbody) {
    if (ungradedEntries.length > 0 && (activeGradeFilter === 'all' || activeGradeFilter === '?')) {
      ungradedSection.style.display = 'block';
      document.getElementById('ungraded-count').textContent = ungradedEntries.length + ' parts';
      ungradedTbody.innerHTML = '';
      ungradedEntries.forEach((e, i) => {
        const bg = i % 2 === 0 ? 'var(--surface)' : 'var(--surface2)';
        const d = new Date(e.date + 'T12:00:00');
        const dateLabel = d.toLocaleDateString('en-GB', { weekday:'short', day:'numeric', month:'short', year:'numeric' });
        const tr = document.createElement('tr');
        tr.innerHTML = '<td style="padding:0.4rem 0.8rem;background:' + bg + ';font-size:0.7rem;color:var(--muted);">' + dateLabel + '</td>'
          + '<td style="padding:0.4rem 0.8rem;background:' + bg + ';font-weight:600;">' + (e.partId || '') + '</td>'
          + '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;color:var(--green);font-weight:700;">' + (e.count || 0) + '</td>'
          + '<td style="padding:0.4rem 0.8rem;background:' + bg + ';text-align:center;">' + (isAdmin() ? '<select data-pid="' + escAttr(e.partId || '') + '" onchange="gradeFromUngraded(this)" style="background:var(--surface2);color:var(--text);border:1px solid var(--border);border-radius:3px;font-size:0.7rem;padding:2px 4px;"><option value="">Set grade…</option><option value="X">X</option><option value="S">S</option><option value="C">C</option><option value="B">B</option><option value="A">A</option><option value="M">M</option></select>' : '<span style="color:var(--muted);font-size:0.65rem;">ask admin</span>') + '</td>';
        ungradedTbody.appendChild(tr);
      });
    } else {
      ungradedSection.style.display = 'none';
    }
  }
}

function adjustDaysWorked(delta) {
  // Per-month adjustments stored separately e.g. _meta/adj/2026-03
  const monthKey = activeMonthFilter === 'all' ? 'alltime' : activeMonthFilter;
  const currentAdj = typeof daysAdjustment === 'object' ? (daysAdjustment[monthKey] || 0) : (activeMonthFilter === 'all' ? (daysAdjustment || 0) : 0);
  const base = activeMonthFilter === 'all' ? baseDaysWorked : (monthBaseDays[monthKey] || baseDaysWorked);
  const minDays = base > 0 ? 0.5 : 0;
  const newAdj = Math.max(-(base - minDays), currentAdj + delta);

  if (typeof daysAdjustment !== 'object') daysAdjustment = {};
  daysAdjustment[monthKey] = newAdj; // only update this month

  const adjustedDays = Math.max(0.5, base + newAdj);
  // Save per-month adjustment to Firebase
  if (HISTORY_REF) HISTORY_REF.child('_meta/adj/' + monthKey).set(newAdj);
  updateDaysDisplay(adjustedDays);
}

function updateDaysDisplay(adjustedDays) {
  const totalRepairs = parseInt(document.getElementById('hist-total-repairs').textContent) || 0;
  // Also update month summary card with new adjusted days
  if (activeMonthFilter !== 'all' && historyData && historyData[activeMonthFilter]) {
    const allEntries = [];
    const day = historyData[activeMonthFilter];
    if (day) Object.keys(day).forEach(k => allEntries.push({ date: activeMonthFilter, ...day[k], _key: k }));
    updateMonthSummary(allEntries, activeMonthFilter, adjustedDays);
  }
  const baseStr = baseDaysWorked % 1 === 0 ? baseDaysWorked : baseDaysWorked.toFixed(1);
  const adjStr = adjustedDays % 1 === 0 ? adjustedDays : adjustedDays.toFixed(1);
  const _mk = activeMonthFilter === 'all' ? 'alltime' : activeMonthFilter;
  const currentAdj = (typeof daysAdjustment === 'object') ? (daysAdjustment[_mk] || 0) : 0;
  let diffStr = '';
  if (currentAdj !== 0) {
    const sign = currentAdj > 0 ? '+' : '';
    const diffVal = currentAdj % 1 === 0 ? currentAdj : currentAdj.toFixed(1);
    diffStr = ` <span style="font-size:0.7rem;color:${currentAdj > 0 ? 'var(--accent)' : 'var(--green)'};">(${sign}${diffVal})</span>`;
  }
  document.getElementById('hist-days-worked').textContent = adjStr;
  const adjInd = document.getElementById('days-adj-indicator');
  if (adjInd) {
    // Strip HTML tags for the indicator, just show (+0.5) or (-0.5)
    const adjMatch = diffStr.match(/[+-][0-9.]+/);
    adjInd.textContent = adjMatch ? adjMatch[0] + ' day adj.' : '';
    adjInd.style.color = currentAdj > 0 ? 'var(--green)' : currentAdj < 0 ? 'var(--accent)' : '';
  }
  const avg = adjustedDays > 0 ? (totalRepairs / adjustedDays).toFixed(1) : '0';
  document.getElementById('hist-avg-repairs').textContent = avg;
  updateTargetTracker(totalRepairs, adjustedDays);
}

function recalculateHistory() {
  if (!HISTORY_REF) return;
  if (!confirm('Recalculate all historical earnings using the current part gradings? This will update all values in your history.')) return;

  const gradeMap = window.gradeMap || {};
  let updateCount = 0;

  Object.keys(historyData).forEach(date => {
    const day = historyData[date];
    if (!day) return;
    Object.keys(day).forEach(key => {
      const entry = day[key];
      if (!entry || !entry.partId) return;

      const cls = gradeMap[entry.partId] || entry.cls || '';
      const val = typeof GRADE_VALUE[cls] === 'number' ? GRADE_VALUE[cls] : 0;
      const count = entry.count || 0;
      const berCount = entry.berCount || 0;
      const berEarned = berCount * 0.10;
      const earned = (val * count) + berEarned;

      // Only update if values have changed
      if (cls !== entry.cls || val !== entry.val || earned !== entry.earned) {
        HISTORY_REF.child(date).child(key).update({ cls, val, earned, berEarned });
        updateCount++;
      }
    });
  });

  setTimeout(() => {
    alert('✓ Recalculation complete — ' + updateCount + ' entries updated with current grades and values.');
  }, 500);
}

function clearHistory() {
  if (confirm('Clear all history permanently? This cannot be undone.')) {
    HISTORY_REF.remove();
  }
}

// === MANAGE PARTS, HIDDEN PARTS & NEW PART MODALS ===
function openManageModal() {
  document.getElementById('manage-help').textContent = isAdmin()
    ? "Tick a department to put a part on that department's list (a part can be on several). Changing a grade also grades any of your ungraded history for that part. Type ? to list parts with no grade, or ! for active parts that need a department (RC MCD or no team in PowerBI) or a grade. Parts with a real PowerBI department are placed automatically. Tick None if no department should show a part (type ~ to see those). Changes sync to everyone."
    : 'Look up any part here. To change a grade or department, press Request change and the admin will approve it. Type ? to list parts with no grade.';
  document.getElementById('manage-search').value = '';
  document.getElementById('manage-overlay').style.display = 'flex';
  renderManageList();
  setTimeout(() => document.getElementById('manage-search').focus(), 100);
}

function closeManageModal() {
  document.getElementById('manage-overlay').style.display = 'none';
}

function renderManageList() {
  const q = (document.getElementById('manage-search').value || '').trim().toUpperCase();
  const admin = isAdmin();
  const customKeys = new Set(Object.keys(customParts));
  const depts = ['DCD', 'PC', 'PCB', 'Printers'];
  const labels = { DCD: 'DCD', PC: 'PC', PCB: 'PCB', Printers: 'P&amp;P' };
  let pool;
  if (q === '!') {            // "!" lists new parts: in the reports but on no department list, or with no grade
    pool = new Set(window.newParts || []);
  } else if (q === '~') {     // "~" lists parts marked None (no department shows them)
    pool = new Set(noDeptParts);
  } else if (q === '?') {            // "?" lists every part with no grade yet
    pool = new Set([...(window.allPartIds || [])].filter(p => !(window.gradeMap || {})[p]));
  } else if (q) {             // searching covers every part in every department
    pool = new Set([...(window.allPartIds || []), ...Object.keys(window.gradeMap || {})]);
  } else {
    pool = new Set([...DCD_PARTS, ...DEPT_PARTS.PC, ...DEPT_PARTS.PCB, ...DEPT_PARTS.Printers,
      ...Object.values(customParts).map(cp => cp.partId).filter(Boolean)]);
  }
  let list = [...pool].sort();
  if (q && q !== '?' && q !== '!' && q !== '~') list = list.filter(p => p.includes(q));
  const total = list.length;
  list = list.slice(0, 80);
  const cell = 'padding:0.35rem 0.5rem;border-bottom:1px solid var(--border);';
  const selStyle = 'background:var(--surface2);color:var(--text);border:1px solid var(--border);border-radius:3px;font-size:0.72rem;padding:2px 4px;';
  const th = 'padding:0.35rem 0.5rem;';
  let h = '<table style="width:100%;border-collapse:collapse;font-size:0.72rem;"><thead><tr style="color:var(--muted);text-transform:uppercase;font-size:0.6rem;letter-spacing:0.06em;">'
    + '<th style="text-align:left;' + th + '">Part ID</th><th style="' + th + '">Grade</th>'
    + depts.map(d => '<th style="' + th + '">' + labels[d] + '</th>').join('') + '<th style="' + th + '" title="No department will show this part">None</th><th style="' + th + '"></th></tr></thead><tbody>';
  list.forEach(p => {
    const gr = (window.gradeMap || {})[p] || '';
    const pa = escAttr(p);
    const opts = ['', 'X', 'S', 'C', 'B', 'A', 'M'].map(g =>
      '<option value="' + g + '"' + (g === gr ? ' selected' : '') + '>' + (g || '—') + '</option>').join('');
    const gradeCell = admin
      ? '<select data-pid="' + pa + '" onchange="manageGradeChange(this)" style="' + selStyle + '">' + opts + '</select>'
      : '<strong>' + (gr || '—') + '</strong>';
    const boxes = depts.map(d => '<td style="' + cell + 'text-align:center;"><input type="checkbox" data-pid="' + pa + '" data-dept="' + d
      + '" onchange="manageDeptChange(this)"' + (DEPT_PARTS[d].has(p) ? ' checked' : '') + (admin ? '' : ' disabled') + '/></td>').join('');
    const noneBox = '<td style="' + cell + 'text-align:center;"><input type="checkbox" data-pid="' + pa + '" onchange="manageNoDeptChange(this)"' + (noDeptParts.has(p) ? ' checked' : '') + (admin ? '' : ' disabled') + '/></td>';
    const lastCell = admin
      ? (customKeys.has(safeKey(p))
          ? '<button data-pid="' + pa + '" onclick="manageDeleteCustom(this)" style="font-size:0.65rem;padding:2px 8px;cursor:pointer;border-radius:3px;background:transparent;color:var(--accent);border:1px solid var(--accent);">Delete custom</button>'
          : '')
      : '<button data-pid="' + pa + '" onclick="requestChangeFor(this)" style="font-size:0.65rem;padding:2px 8px;cursor:pointer;border-radius:3px;background:transparent;color:var(--accent2);border:1px solid var(--accent2);">Request change</button>';
    h += '<tr>'
      + '<td style="' + cell + 'font-family:IBM Plex Mono,monospace;">' + pa
        + ((window.partTeams || {})[p] ? '<div style="font-size:0.58rem;color:var(--muted);font-family:inherit;">' + escAttr(window.partTeams[p]) + '</div>' : '')
        + (q === '!' && (window.newReasons || {})[p] ? '<div style="font-size:0.58rem;color:var(--accent);font-family:inherit;font-weight:700;">' + escAttr(window.newReasons[p]) + '</div>' : '') + '</td>'
      + '<td style="' + cell + 'text-align:center;">' + gradeCell + '</td>'
      + boxes + noneBox
      + '<td style="' + cell + 'text-align:right;">' + lastCell + '</td></tr>';
  });
  h += '</tbody></table>';
  if (!list.length) h = '<div style="padding:1rem;color:var(--muted);font-size:0.75rem;">No matching parts.</div>';
  document.getElementById('manage-list').innerHTML = h;
  document.getElementById('manage-count').textContent = total > 80 ? ('Showing 80 of ' + total + ' — type to narrow down') : (total + ' part' + (total === 1 ? '' : 's'));
}

function manageGradeChange(sel) { setPartGrade(sel.dataset.pid, sel.value); }
function manageDeptChange(cb) { setPartDept(cb.dataset.pid, cb.dataset.dept, cb.checked); }
function manageNoDeptChange(cb) {
  setPartNoDept(cb.dataset.pid, cb.checked);
  renderManageList();
}

function openHiddenModal() {
  const cell = 'padding:0.4rem 0.6rem;border-bottom:1px solid var(--border);';
  let h = '<table style="width:100%;border-collapse:collapse;font-size:0.72rem;"><thead><tr style="color:var(--muted);text-transform:uppercase;font-size:0.6rem;letter-spacing:0.06em;">'
    + '<th style="text-align:left;padding:0.4rem 0.6rem;">Part ID</th><th style="padding:0.4rem 0.6rem;">Calls</th><th style="text-align:left;padding:0.4rem 0.6rem;">Why hidden</th></tr></thead><tbody>';
  hiddenParts.forEach(p => {
    h += '<tr><td style="' + cell + 'font-family:IBM Plex Mono,monospace;">' + escAttr(p.id) + '</td>'
      + '<td style="' + cell + 'text-align:center;">' + p.calls + '</td>'
      + '<td style="' + cell + '">' + escAttr(p.reasons.join(', ')) + '</td></tr>';
  });
  h += '</tbody></table>';
  if (!hiddenParts.length) h = '<div style="padding:1rem;color:var(--muted);font-size:0.75rem;">Nothing is hidden by the keyword filter.</div>';
  document.getElementById('hidden-list').innerHTML = h;
  document.getElementById('hidden-overlay').style.display = 'flex';
}

function closeHiddenModal() {
  document.getElementById('hidden-overlay').style.display = 'none';
}

function openNewPartModal() {
  if (!requireAdmin()) return;
  document.getElementById('newpart-id-input').value = '';
  document.getElementById('newpart-grade-input').value = '';
  document.getElementById('newpart-dcd-input').checked = false;
  document.getElementById('newpart-error').style.display = 'none';
  document.getElementById('newpart-overlay').style.display = 'flex';
  setTimeout(() => document.getElementById('newpart-id-input').focus(), 100);
}

function closeNewPartModal() {
  document.getElementById('newpart-overlay').style.display = 'none';
}

// === PUBLIC READ-ONLY EARNINGS VIEW ===
let publicEarningsRequestId = 0;
let publicEarningsUserId = null;
let publicEarningsEntries = [];
let publicEarningsPeriod = 'current';

function initEarningsViewer() {
  const select = document.getElementById('earnings-user-select');
  if (!select) return;
  Object.entries(USERS)
    .sort(([, a], [, b]) => a.name.localeCompare(b.name))
    .forEach(([id, user]) => {
      const option = document.createElement('option');
      option.value = id;
      option.textContent = user.name;
      select.appendChild(option);
    });
  select.value = '';
}

async function loadPublicEarnings(userId) {
  const requestId = ++publicEarningsRequestId;
  const status = document.getElementById('public-earnings-status');
  const content = document.getElementById('public-earnings-content');
  const tbody = document.getElementById('public-earnings-body');
  content.style.display = 'none';
  tbody.replaceChildren();
  publicEarningsUserId = null;
  publicEarningsEntries = [];
  if (!userId || !USERS[userId]) {
    status.textContent = '';
    return;
  }

  status.textContent = 'Loading ' + USERS[userId].name + "'s earnings…";
  try {
    const snapshot = await db.ref('users/' + userId + '/history').once('value');
    if (requestId !== publicEarningsRequestId) return;
    const data = snapshot.val() || {};
    publicEarningsEntries = [];
    Object.keys(data).sort().reverse().forEach(date => {
      if (date === '_meta' || !data[date]) return;
      Object.values(data[date]).forEach(entry => {
        if (!entry || typeof entry !== 'object') return;
        const count = Number(entry.count) || 0;
        const berCount = Number(entry.berCount) || 0;
        const earned = Number(entry.earned) || (Number(entry.val) * count) || 0;
        publicEarningsEntries.push({ date, partId: entry.partId || '', count, berCount, cls: entry.cls || '', earned, isBerOnly: entry.type === 'ber' });
      });
    });

    publicEarningsUserId = userId;
    publicEarningsPeriod = 'current';
    const tabs = document.querySelectorAll('#public-earnings-period-tabs .hist-sub-btn');
    tabs.forEach((tab, index) => tab.classList.toggle('active', index === 0));
    const monthSelect = document.getElementById('public-earnings-month');
    monthSelect.replaceChildren();
    const previousMonths = [...new Set(publicEarningsEntries.map(entry => entry.date.slice(0, 7)))]
      .filter(month => month < TODAY.slice(0, 7))
      .sort()
      .reverse();
    previousMonths.forEach(month => {
      const option = document.createElement('option');
      option.value = month;
      option.textContent = new Date(month + '-01T12:00:00').toLocaleDateString('en-GB', { month: 'long', year: 'numeric' });
      monthSelect.appendChild(option);
    });
    monthSelect.value = previousMonths[0] || '';
    content.style.display = 'block';
    renderPublicEarnings();
  } catch (error) {
    if (requestId !== publicEarningsRequestId) return;
    status.textContent = 'Unable to load earnings. Please try again later.';
    console.error('Unable to load public earnings for ' + userId, error);
  }
}

function setPublicEarningsPeriod(period, button) {
  publicEarningsPeriod = period;
  document.querySelectorAll('#public-earnings-period-tabs .hist-sub-btn').forEach(tab => tab.classList.remove('active'));
  button.classList.add('active');
  document.getElementById('public-earnings-month').style.display = period === 'previous' ? '' : 'none';
  renderPublicEarnings();
}

function renderPublicEarnings() {
  if (!publicEarningsUserId) return;
  const month = publicEarningsPeriod === 'current'
    ? TODAY.slice(0, 7)
    : document.getElementById('public-earnings-month').value;
  const entries = month ? publicEarningsEntries.filter(entry => entry.date.startsWith(month)) : [];
  const tbody = document.getElementById('public-earnings-body');
  tbody.replaceChildren();

  let repairs = 0;
  let earned = 0;
  entries.forEach((entry, index) => {
    if (!entry.isBerOnly) repairs += entry.count;
    earned += entry.earned;
    const row = document.createElement('tr');
    const date = document.createElement('td');
    date.textContent = formatHistoryDate(entry.date);
    const partId = document.createElement('td');
    partId.textContent = entry.partId;
    const repairCount = document.createElement('td');
    repairCount.textContent = entry.count ? String(entry.count) : '—';
    const berCount = document.createElement('td');
    berCount.textContent = entry.berCount ? entry.berCount + ' BER' : '—';
    const grade = document.createElement('td');
    grade.textContent = entry.cls;
    const earnedCell = document.createElement('td');
    earnedCell.textContent = entry.earned ? '£' + entry.earned.toFixed(2) : '—';
    [date, partId, repairCount, berCount, grade, earnedCell].forEach((cell, cellIndex) => {
      cell.style.background = index % 2 === 0 ? 'var(--surface)' : 'var(--surface2)';
      if (cellIndex >= 2) cell.style.textAlign = 'center';
      row.appendChild(cell);
    });
    tbody.appendChild(row);
  });

  document.getElementById('public-earnings-parts').textContent = String(entries.length);
  document.getElementById('public-earnings-repairs').textContent = String(repairs);
  document.getElementById('public-earnings-total').textContent = '£' + earned.toFixed(2);
  const monthLabel = month
    ? new Date(month + '-01T12:00:00').toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })
    : 'previous months';
  document.getElementById('public-earnings-status').textContent = entries.length
    ? USERS[publicEarningsUserId].name + "'s earnings for " + monthLabel + '.'
    : 'No earnings recorded for ' + monthLabel + '.';
}

// === APP BOOTSTRAP ===
// Sign in from a personal link (?user=Name), or show users for the selected department.
function initApp() {
  initEarningsViewer();
  initLoginFromLink();
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initApp);
else initApp();
