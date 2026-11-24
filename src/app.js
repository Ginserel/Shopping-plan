/* Shopping Plan – browser UI. Needs XLSX (SheetJS), SPCore, SPXlsx and SPSample loaded first. */
(function () {
  'use strict';
  const SP = window.SPCore;
  const DAY = SP.DAY;
  const app = document.getElementById('app');
  const fileInput = document.getElementById('file-input');

  /* ----------------------------- storage ---------------------------- */
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* private mode */ } },
  };
  const clone = (o) => JSON.parse(JSON.stringify(o));

  function loadCfg() {
    const v3 = store.get('sp.cfg.v3', null);
    if (v3) return Object.assign(clone(SP.DEFAULT_CONFIG), v3);
    // Older saved rules: keep them, but kitting and decant are now left out by default.
    const v2 = store.get('sp.cfg.v2', {});
    const keep = clone(v2);
    delete keep.includeKitting; delete keep.skipKinds; delete keep.xdockKinds;
    return Object.assign(clone(SP.DEFAULT_CONFIG), keep);
  }
  function loadAssign() {
    const a = store.get('sp.assign.v1', {});
    const cutoff = Date.now() - 14 * 86400000;
    for (const id of Object.keys(a)) if (!a[id] || !a[id].k || a[id].at < cutoff) delete a[id];
    return a;
  }

  /* ------------------------------ state ----------------------------- */
  const state = {
    picking: null, pickingName: '', pickingSample: true,
    transport: null, transportName: '', transportSample: true,
    cfg: loadCfg(),
    overrides: store.get('sp.overrides.v1', {}),
    assign: loadAssign(), // orderId -> { k: trailer key, at: time added }
    comments: store.get('sp.comments.v1', {}), // trailer key -> text for the admin
    undo: [],
    now: null, nowFromFile: true,
    tab: 'plan', day: null, customer: '',
    showOther: false, showEmpty: false, showPast: false, listEmpty: true,
    message: null, result: null, list: null,
  };
  const saveCfg = () => store.set('sp.cfg.v3', state.cfg);
  const saveOverrides = () => store.set('sp.overrides.v1', state.overrides);
  const saveAssign = () => store.set('sp.assign.v1', state.assign);
  const saveComments = () => store.set('sp.comments.v1', state.comments);

  function loadSample() {
    state.picking = SP.parsePicking(window.SPSample.picking());
    state.transport = SP.parseTransport(window.SPSample.transport());
    state.pickingName = 'Example picking status';
    state.transportName = 'Example transport plan';
    state.pickingSample = state.transportSample = true;
    state.nowFromFile = true;
    state.day = null;
  }

  /* ---------------------------- helpers ----------------------------- */
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const pct = (f) => (f > 0 && f < 0.005 ? '<1%' : `${Math.round(f * 100)}%`);
  const pal = (n) => (n >= 10 ? Math.round(n) : n >= 1 ? n.toFixed(1).replace(/\.0$/, '') : n > 0 ? '<1' : '0');
  const t = SP.fmtTime;
  const dayOf = (abs) => Math.floor(abs / DAY);
  function when(abs, refDay) {
    if (abs == null) return '–';
    return dayOf(abs) === refDay ? t(abs) : `${SP.fmtDay(abs).slice(0, 3)} ${t(abs)}`;
  }
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
  const laneName = (s) => `${s.to} ${t(s.delivery)}${s.kind ? ' ' + s.kind : ''}`;
  const custOk = (c) => !state.customer || c === state.customer;
  const STATE_CLASS = { picked: 'ok', picking: 'under', released: '', release: 'info' };
  function statusClass(text) {
    if (!text) return '';
    if (/^picked/i.test(text)) return 'ok';
    if (/progress/i.test(text)) return 'under';
    if (/^to release/i.test(text)) return 'info';
    return '';
  }
  function toast(msg, undo) {
    document.querySelectorAll('.toast').forEach((x) => x.remove());
    const el = document.createElement('div');
    el.className = 'toast';
    el.setAttribute('role', 'status');
    el.innerHTML = `<span>${esc(msg)}</span>${undo ? '<button type="button" data-act="undo">Undo</button>' : ''}`;
    document.body.appendChild(el);
    el.addEventListener('click', (e) => { if (e.target.closest('[data-act=undo]')) { doUndo(); el.remove(); } });
    setTimeout(() => el.remove(), undo ? 5000 : 2200);
  }

  /* ------------------------------ files ----------------------------- */
  function readFile(file) {
    return new Promise((resolve) => {
      const fr = new FileReader();
      fr.onload = () => {
        try {
          const wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: false });
          let best = null;
          for (const name of wb.SheetNames) {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
            const d = SP.detect(rows);
            if (d.type && (!best || rows.length > best.n)) best = { n: rows.length, d };
          }
          resolve(best ? best.d : { type: null });
        } catch (e) {
          resolve({ type: null, error: e.message });
        }
      };
      fr.onerror = () => resolve({ type: null, error: 'could not read the file' });
      fr.readAsArrayBuffer(file);
    });
  }
  async function loadFiles(files) {
    const problems = [];
    for (const f of files) {
      const d = await readFile(f);
      if (d.type === 'picking') {
        state.picking = d.data; state.pickingName = f.name; state.pickingSample = false; state.nowFromFile = true; state.day = null;
      } else if (d.type === 'transport') {
        state.transport = d.data; state.transportName = f.name; state.transportSample = false; state.day = null;
      } else {
        problems.push(`"${f.name}" ${d.error ? '– ' + d.error : 'was not recognised. A picking status needs Delivery Date and Order ID columns; a transport plan needs Deliver To and Delivery Time columns.'}`);
      }
    }
    state.message = problems.length ? { kind: 'err', text: problems.join(' ') } : null;
    if (!problems.length) toast(`Loaded ${plural(files.length, 'file')}`);
    render();
  }
  fileInput.addEventListener('change', () => { if (fileInput.files.length) loadFiles([...fileInput.files]); fileInput.value = ''; });
  let dragDepth = 0;
  const isFileDrag = (e) => e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files');
  window.addEventListener('dragenter', (e) => { if (!isFileDrag(e)) return; e.preventDefault(); dragDepth++; document.body.classList.add('drag'); });
  window.addEventListener('dragleave', (e) => { if (!isFileDrag(e)) return; if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('drag'); } });
  window.addEventListener('dragover', (e) => { if (isFileDrag(e)) e.preventDefault(); });
  window.addEventListener('drop', (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault(); dragDepth = 0; document.body.classList.remove('drag');
    if (e.dataTransfer.files.length) loadFiles([...e.dataTransfer.files]);
  });

  /* ----------------------------- compute ---------------------------- */
  const assignMap = () => Object.fromEntries(Object.entries(state.assign).map(([id, a]) => [id, a.k]));
  function compute() {
    if (state.nowFromFile || state.now == null) state.now = state.picking.asOf != null ? state.picking.asOf : Math.floor(Date.now() / 60000) - new Date().getTimezoneOffset();
    const fixed = assignMap();
    state.result = SP.plan(state.picking, state.transport, state.cfg, { now: state.now, overrides: state.overrides, fixed });
    state.list = SP.buildList(state.result, fixed);
    const days = dayOptions();
    if (state.day == null || !days.includes(state.day)) state.day = defaultDay(days);
  }
  function dayOptions() {
    const r = state.result;
    const set = new Set(r.orders.map((o) => dayOf(o.delivery)));
    set.add(dayOf(state.now)); set.add(dayOf(state.now) + 1);
    const withTrailers = new Set(r.board.map((s) => dayOf(s.delivery)));
    return [...set].filter((d) => withTrailers.has(d)).sort((a, b) => a - b);
  }
  function defaultDay(days) {
    const r = state.result;
    const today = dayOf(state.now);
    const open = (d) => r.orders.filter((o) => dayOf(o.delivery) === d && o.remaining > 0).length;
    const next = days.filter((d) => d >= today).find((d) => open(d) > 0);
    return next != null ? next : days.includes(today + 1) ? today + 1 : days[0];
  }

  /* ----------------------- choosing orders ------------------------- */
  function snapshot() {
    state.undo.push(JSON.stringify(state.assign));
    if (state.undo.length > 30) state.undo.shift();
  }
  function doUndo() {
    if (!state.undo.length) return;
    state.assign = JSON.parse(state.undo.pop());
    saveAssign();
    rerender();
  }
  function assign(ids, key, msg) {
    if (!ids.length) return;
    snapshot();
    const at = Date.now();
    for (const id of ids) {
      if (key) state.assign[id] = { k: key, at };
      else delete state.assign[id];
    }
    saveAssign();
    rerender();
    if (msg) toast(msg, true);
  }
  const orderById = () => new Map(state.result.orders.map((o) => [o.id, o]));
  const rowByKey = () => new Map(state.list.rows.map((r) => [r.key, r]));

  /* ------------------------------ views ----------------------------- */
  function viewTop(r) {
    const p = state.picking, tr = state.transport;
    const used = r.used.filter((u) => u.day === state.day || u.day === state.day + 1);
    const sample = state.pickingSample || state.transportSample;
    const ign = r.ignored || { kitting: 0, other: 0 };
    const ignTxt = [ign.kitting ? `${ign.kitting} kitting` : '', ign.other ? `${ign.other} decant` : ''].filter(Boolean).join(' and ');
    const borrowed = used.filter((u) => !u.exact);
    return `
      <header class="topbar">
        <div class="brand"><h1>Shopping Plan</h1><span class="sub">Pick the right orders for every trailer</span></div>
        <span class="spacer"></span>
        <button class="btn primary" data-act="load">Load files</button>
        ${sample ? '' : '<button class="btn ghost" data-act="sample">Show example</button>'}
      </header>
      ${sample ? `<div class="banner"><strong>Example data.</strong> These orders and trailers are made up. Load the Picking Status export and the transport plan (or drop both files on this page) to plan for real.</div>` : ''}
      ${state.message ? `<div class="banner ${state.message.kind}">${esc(state.message.text)}</div>` : ''}
      <section class="sources dropzone">
        <div class="source ${state.pickingSample ? 'is-sample' : ''}">
          <span class="label">Picking status</span>
          <span class="name">${esc(state.pickingName)}</span>
          <span class="small muted">${r.orders.length} orders${ignTxt ? ` (${ignTxt} left out)` : ''}${p.asOf != null ? ` · exported ${SP.fmtDayTime(p.asOf)}` : ''}</span>
        </div>
        <div class="source ${state.transportSample ? 'is-sample' : ''}">
          <span class="label">Transport plan</span>
          <span class="name">${esc(state.transportName)}</span>
          <span class="small muted">${tr.days.length ? `${tr.days.length} days, ${SP.fmtDay(tr.days[0].day * DAY)} – ${SP.fmtDay(tr.days[tr.days.length - 1].day * DAY)}` : 'no days found'}</span>
        </div>
      </section>
      ${borrowed.length ? `<div class="banner warn">No sheet in the transport plan for ${borrowed.map((u) => `${SP.fmtDay(u.day * DAY)} (times taken from ${SP.fmtDay(u.source * DAY)})`).join(', ')}. Load a newer transport plan when it's ready.</div>` : ''}
    `;
  }

  function viewControls(r) {
    const days = dayOptions();
    const customers = [...new Set(r.board.filter((s) => dayOf(s.delivery) === state.day).map((s) => s.customer).concat(r.orders.map((o) => o.customer).filter(Boolean)))].sort();
    return `
      <section class="controls">
        <label class="field"><span class="label">Delivery day</span>
          <select id="day">${days.map((d) => `<option value="${d}" ${d === state.day ? 'selected' : ''}>${SP.fmtDay(d * DAY)}${d === dayOf(state.now) ? ' (today)' : d === dayOf(state.now) + 1 ? ' (tomorrow)' : ''}</option>`).join('')}</select></label>
        <label class="field"><span class="label">Customer</span>
          <select id="customer"><option value="">All customers</option>${customers.map((c) => `<option ${c === state.customer ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
        <label class="field"><span class="label">Time now</span>
          <input type="datetime-local" id="now" value="${SP.absToLocalInput(state.now)}"></label>
        ${state.nowFromFile ? '' : '<button class="btn ghost" data-act="now-file">Use export time</button>'}
      </section>`;
  }

  function dayRows() {
    return state.list.rows.filter((row) => dayOf(row.slot.delivery) === state.day);
  }
  function visibleRows() {
    return dayRows().filter((row) => {
      const s = row.slot;
      if (!custOk(s.customer)) return false;
      if (row.chosen.length) return true;
      if (s.past && !state.showPast) return false;
      if (s.role !== 'ours' && !state.showOther) return false;
      if (!row.suggested.length && !state.showEmpty) return false;
      return true;
    });
  }
  function dayPool() {
    return state.list.pool.filter((x) => {
      const o = x.order;
      if (!custOk(o.customer)) return false;
      const sugDay = x.suggestion ? dayOf(x.suggestion.slot.delivery) : null;
      return dayOf(o.delivery) <= state.day && (sugDay == null || sugDay <= state.day) || sugDay === state.day;
    });
  }

  function viewKpis() {
    const rows = dayRows();
    const onList = rows.reduce((n, r) => n + r.chosen.length, 0);
    const trailers = rows.filter((r) => r.chosen.length).length;
    const pool = dayPool();
    const toPlace = pool.filter((x) => !x.order.collection).length;
    const sugg = visibleRows().reduce((n, r) => n + r.suggested.length, 0);
    const attention = rows.filter((r) => r.chosen.length && r.warnings.some((w) => w.kind !== 'info')).length;
    const toRelease = rows.reduce((n, r) => n + r.toRelease.length, 0);
    return `
      <section class="kpis">
        <div class="kpi attn"><span class="label">On the shopping list</span><span class="v">${onList}</span><span class="small muted">${plural(onList, 'order')} on ${plural(trailers, 'trailer')} · ${toRelease} to release</span></div>
        <div class="kpi ${toPlace ? 'warn' : ''}"><span class="label">Not on the list yet</span><span class="v">${toPlace}</span><span class="small muted">orders due ${SP.fmtDay(state.day * DAY)} or earlier</span></div>
        <div class="kpi"><span class="label">Suggestions waiting</span><span class="v">${sugg}</span><span class="small muted">add or ignore them on the Plan</span></div>
        <div class="kpi ${attention ? 'bad' : ''}"><span class="label">Trailers to check</span><span class="v">${attention}</span><span class="small muted">over capacity, late orders, pick-by passed</span></div>
      </section>`;
  }

  function viewTabs() {
    const rows = dayRows();
    const listed = rows.filter((r) => r.chosen.length).length;
    const ex = state.result.exceptions.filter((e) => e.kind !== 'collection' && e.kind !== 'estimate').length;
    const tabs = [['plan', 'Plan', null], ['list', 'Shopping list', listed], ['checks', 'Checks', ex], ['orders', 'Orders', state.result.orders.length], ['rules', 'Rules', null]];
    return `<nav class="tabs" role="tablist">${tabs.map(([k, label, n]) => `<button class="tab" role="tab" data-tab="${k}" aria-selected="${state.tab === k}">${label}${n != null ? `<span class="count">${n}</span>` : ''}</button>`).join('')}</nav>`;
  }

  /* ------------------------------ plan ------------------------------ */
  function orderLine(o, ref) {
    return `due ${when(o.delivery, ref)} · ${o.remaining}/${o.total} picks · ${pct(o.frac)}`;
  }
  const SUG_SHORT = { home: 'due then', shop: 'fills trailer', noslot: 'no trailer at its time', roll: 'missed trailer' };

  function laneOptions(row, ref) {
    // Orders that could go on this trailer: same customer first, nearest due time first.
    const s = row.slot;
    const pool = state.list.pool.filter((x) => Math.abs(dayOf(x.order.delivery) - state.day) <= 1);
    const near = (x) => Math.abs(x.order.delivery - s.delivery);
    const same = pool.filter((x) => x.order.customer === s.customer).sort((a, b) => near(a) - near(b));
    const other = pool.filter((x) => x.order.customer !== s.customer).sort((a, b) => a.order.delivery - b.order.delivery);
    const opt = (x) => `<option value="${esc(x.order.id)}">${esc(when(x.order.delivery, ref))} · ${esc(x.order.seller || x.order.id)}${x.order.customer !== s.customer ? ` · ${esc(x.order.customer || x.order.company)}` : ''} · ${x.order.remaining} picks</option>`;
    return `<option value="">+ Add an order…</option>${same.length ? `<optgroup label="${esc(s.customer)}">${same.map(opt).join('')}</optgroup>` : ''}${other.length ? `<optgroup label="Other customers">${other.map(opt).join('')}</optgroup>` : ''}`;
  }

  function lane(row, ref) {
    const s = row.slot, cfg = state.result.cfg;
    const max = Math.max(1, row.load + row.sugLoad);
    const scale = (f) => `${(Math.min(f, max) / max) * 100}%`;
    const over = row.load > cfg.maxFill + 1e-9;
    const tags = [];
    if (s.role === 'xdock') tags.push('<span class="chip">X-Dock</span>');
    if (s.role === 'other') tags.push(`<span class="chip">From ${esc(s.from)}</span>`);
    if (s.emer) tags.push('<span class="chip low">EMER</span>');
    if (s.past) tags.push('<span class="chip">Gone</span>');
    else if (s.status) tags.push(`<span class="chip">${esc(s.status)}</span>`);
    if (s.templated) tags.push('<span class="chip" title="Times copied from another day of the transport plan">Expected</span>');
    const chosen = row.chosen.map((o) => {
      const st = SP.orderState(o);
      const notes = row.notes.get(o.id) || [];
      return `<li class="pick" draggable="true" data-order="${esc(o.id)}">
        <span class="dot ${STATE_CLASS[st]}" title="${esc(SP.STATE_TEXT[st])}"></span>
        <span class="who"><span class="ref">${esc(o.seller || o.id)}</span>
          <span class="small muted mono">${esc(orderLine(o, ref))} · wave ${esc(o.wave)}</span>
          ${notes.map((n) => `<span class="chip ${n.kind === 'late' ? 'low' : n.kind === 'warn' ? 'under' : ''}">${esc(n.text)}</span>`).join('')}
          ${o.comments ? `<span class="small note">${esc(o.comments)}</span>` : ''}</span>
        <span class="st small">${esc(SP.STATE_TEXT[st])}</span>
        <button class="x" data-unassign="${esc(o.id)}" aria-label="Take ${esc(o.seller)} off this trailer" title="Take off the list">×</button>
      </li>`;
    }).join('');
    const sugs = row.suggested.map(({ order: o, reason }) => `<button class="sug sug-${reason.kind}" data-add="${esc(o.id)}" data-to="${esc(row.key)}" title="${esc(reason.text)}">
        <span class="plus" aria-hidden="true">+</span><span class="ref">${esc(o.seller || o.id)}</span>
        <span class="small muted">${esc(reason.kind === 'home' ? `due ${when(o.delivery, ref)}` : reason.kind === 'shop' ? `due ${when(o.delivery, ref)}, ${SUG_SHORT.shop}` : SUG_SHORT[reason.kind])} · ${pct(o.frac)}</span></button>`).join('');
    const warn = row.warnings.map((w) => `<span class="chip ${w.kind === 'over' ? 'over' : w.kind === 'late' ? 'low' : ''}">${esc(w.text)}</span>`).join('');
    const vclass = !row.chosen.length ? 'v-empty' : over ? 'v-over' : row.load >= cfg.targetFill ? 'v-ok' : row.load < cfg.lowFill ? 'v-low' : 'v-under';
    return `<article class="lane ${vclass} role-${s.role} ${s.past ? 'is-past' : ''}" data-drop="${esc(row.key)}" aria-label="${esc(laneName(s))}">
      <div class="lane-time">
        <span class="label">Pick by</span>
        <span class="big">${s.pickBy != null ? when(s.pickBy, ref) : '–'}</span>
        <span class="small muted">load ${t(s.paperwork)}</span>
        <span class="small muted">deliver ${t(s.delivery)}</span>
      </div>
      <div class="lane-main">
        <header>
          <h3>${esc(s.to)}</h3><span class="kind">${esc(s.kind || '')}</span>${tags.join('')}
          <span class="spacer"></span>
          ${row.status ? `<span class="chip ${statusClass(row.status)}">${esc(row.status)}</span>` : ''}
        </header>
        <div class="fill-row" title="Solid: on the list. Striped: suggested.">
          <div class="bar ${over ? 'over' : ''}"><span class="own" style="width:${scale(row.load)}"></span><span class="shop" style="width:${scale(row.sugLoad)}"></span>${max > 1 ? `<i class="target" style="left:${(1 / max) * 100}%" title="One trailer"></i>` : `<i class="target" style="left:${cfg.targetFill * 100}%"></i>`}</div>
          <span class="pct">${pct(row.load)}</span>
          ${row.sugLoad > 0.005 ? `<span class="small muted">+${pct(row.sugLoad)} suggested</span>` : ''}
        </div>
        ${warn ? `<div class="warns">${warn}</div>` : ''}
        ${chosen ? `<ul class="picks">${chosen}</ul>` : ''}
        ${sugs ? `<div class="sugs"><span class="label">Suggested</span>${sugs}${row.suggested.length > 1 ? `<button class="btn small" data-add-all="${esc(row.key)}">Add all ${row.suggested.length}</button>` : ''}</div>` : ''}
        <div class="lane-foot">
          <select data-add-to="${esc(row.key)}" aria-label="Add an order to ${esc(laneName(s))}">${laneOptions(row, ref)}</select>
          <input type="text" data-comment="${esc(row.key)}" value="${esc(state.comments[row.key] || '')}" placeholder="Comment for the admin" aria-label="Comment for ${esc(laneName(s))}">
        </div>
      </div>
    </article>`;
  }

  function poolItem(x, ref) {
    const o = x.order, sug = x.suggestion;
    const st = SP.orderState(o);
    const targets = dayRows().filter((r) => r.slot.customer === o.customer && r.key !== (sug && sug.key) && !r.slot.past);
    const opts = targets.length ? targets : dayRows().filter((r) => r.slot.role === 'ours' && !r.slot.past);
    return `<li class="pool-item" draggable="true" data-order="${esc(o.id)}">
      <div class="pool-head"><span class="dot ${STATE_CLASS[st]}" title="${esc(SP.STATE_TEXT[st])}"></span><span class="ref">${esc(o.seller || o.id)}</span><span class="small muted">${esc(o.customer || o.company)}</span></div>
      <div class="small muted mono">${esc(orderLine(o, ref))}</div>
      ${o.collection ? '<span class="chip under">Customer collection</span>' : ''}
      ${x.missing ? '<span class="chip low">Its trailer is no longer in the transport plan</span>' : ''}
      <div class="pool-actions">
        ${sug ? `<button class="btn small go" data-add="${esc(o.id)}" data-to="${esc(sug.key)}" title="${esc(sug.reason.text)}">→ ${esc(laneName(sug.slot))}${dayOf(sug.slot.delivery) !== state.day ? ` (${SP.fmtDay(sug.slot.delivery).slice(0, 3)})` : ''}</button>` : ''}
        ${opts.length ? `<select data-move="${esc(o.id)}" aria-label="Put ${esc(o.seller)} on a trailer"><option value="">${sug ? 'Other trailer…' : 'Put on trailer…'}</option>${opts.map((r) => `<option value="${esc(r.key)}">${esc(when(r.slot.pickBy != null ? r.slot.pickBy : r.slot.delivery, ref))} pick · ${esc(laneName(r.slot))}</option>`).join('')}</select>` : ''}
      </div>
    </li>`;
  }

  function viewPlan() {
    const ref = state.day;
    const rows = visibleRows();
    const pool = dayPool();
    const sugCount = rows.reduce((n, r) => n + r.suggested.length, 0);
    const hidden = dayRows().filter((r) => custOk(r.slot.customer)).length - rows.length;
    const groups = [
      ['Suggested for a trailer today', pool.filter((x) => x.suggestion && dayOf(x.suggestion.slot.delivery) === state.day)],
      ['Suggested for another day', pool.filter((x) => x.suggestion && dayOf(x.suggestion.slot.delivery) !== state.day)],
      ['No trailer found', pool.filter((x) => !x.suggestion)],
    ].filter(([, l]) => l.length);
    return `
      <div class="plan-tools">
        <label class="check"><input type="checkbox" id="show-other" ${state.showOther ? 'checked' : ''}> X-Dock and other sites' trailers</label>
        <label class="check"><input type="checkbox" id="show-empty" ${state.showEmpty ? 'checked' : ''}> Trailers with nothing to pick</label>
        <label class="check"><input type="checkbox" id="show-past" ${state.showPast ? 'checked' : ''}> Trailers already gone</label>
        <span class="spacer"></span>
        <button class="btn" data-act="undo" ${state.undo.length ? '' : 'disabled'}>Undo</button>
        <button class="btn" data-act="clear-day" id="clear-btn">Clear this day</button>
        <button class="btn primary" data-act="add-all" ${sugCount ? '' : 'disabled'}>Add all suggestions (${sugCount})</button>
      </div>
      <p class="muted small hint">Nothing goes on the shopping list until you add it. Suggestions use the delivery times in the transport plan, and pull later orders forward to fill trailers that would leave part-empty. Drag orders between trailers, or use the buttons.${hidden > 0 ? ` ${plural(hidden, 'trailer')} hidden by the filters above.` : ''}</p>
      <div class="plan-grid">
        <section class="lanes" aria-label="Trailers">${rows.length ? rows.map((r) => lane(r, ref)).join('') : '<div class="empty-state">No trailers to show for this day. Tick the boxes above to see the whole transport plan.</div>'}</section>
        <aside class="pool" data-drop="pool" aria-label="Orders not on the list">
          <div class="pool-title"><h2>Not on the list yet</h2><span class="count">${pool.length}</span></div>
          <p class="small muted">Drop an order here to take it off its trailer.</p>
          ${groups.length ? groups.map(([title, list]) => `<h3 class="pool-group">${title} <span class="muted small">${list.length}</span></h3><ul class="pool-list">${list.map((x) => poolItem(x, ref)).join('')}</ul>`).join('') : '<p class="empty-pool">Every order for this day is on the list.</p>'}
        </aside>
      </div>`;
  }

  /* -------------------------- shopping list ------------------------- */
  function listRows() {
    return dayRows().filter((r) => custOk(r.slot.customer) && (r.chosen.length || (state.listEmpty && !r.slot.past)));
  }
  function viewList() {
    const ref = state.day;
    const rows = listRows();
    const table = SP.pickPlanTable(rows, state.comments);
    const total = rows.reduce((n, r) => n + r.chosen.length, 0);
    const body = table.map(({ row, cells }) => {
      const s = row.slot;
      return `<tr class="${row.chosen.length ? '' : 'quiet'} ${s.role === 'xdock' ? 'xd' : ''}">
        <td><strong>${esc(cells[0])}</strong></td><td>${esc(cells[1])}</td>
        <td class="t mono">${row.chosen.length || s.role === 'ours' ? esc(when(s.pickBy, ref)) : ''}</td><td class="mono">${esc(cells[3])}</td><td class="mono">${esc(cells[4])}</td>
        <td>${row.chosen.map((o) => `<div class="li"><span class="dot ${STATE_CLASS[SP.orderState(o)]}"></span>${esc(o.seller || o.id)}</div>`).join('')}</td>
        <td>${row.status ? `<span class="chip ${statusClass(row.status)}">${esc(row.status)}</span>` : ''}</td>
        <td><input type="text" class="cmt" data-comment="${esc(row.key)}" value="${esc(state.comments[row.key] || '')}" aria-label="Comment"><span class="print-only">${esc(state.comments[row.key] || '')}</span></td>
        <td class="mono small">${row.waves.map(esc).join('<br>')}</td>
        <td class="num mono">${esc(cells[9])}</td><td class="num mono">${esc(cells[10])}</td>
      </tr>`;
    }).join('');
    return `
      <div class="panel-head">
        <div><h2 class="print-title">Pick Plan – ${SP.fmtDay(state.day * DAY)}</h2>
          <p class="muted small">${plural(total, 'order')} on ${plural(rows.filter((r) => r.chosen.length).length, 'trailer')}. Same columns as the hand-made pick plan, plus the waves the admin still has to release.</p></div>
        <div class="btns">
          <label class="check"><input type="checkbox" id="list-empty" ${state.listEmpty ? 'checked' : ''}> Include trailers with no orders</label>
          <button class="btn" data-act="copy">Copy table</button>
          <button class="btn" data-act="print">Print</button>
          <button class="btn primary" data-act="xlsx">Download Excel</button>
        </div>
      </div>
      ${rows.length ? `<div class="table-box"><table class="pickplan">
        <thead><tr>${SP.PICK_PLAN_HEAD.map((h, i) => `<th class="${i >= 9 ? 'num' : ''}">${esc(h)}</th>`).join('')}</tr></thead>
        <tbody>${body}</tbody></table></div>` : '<div class="empty-state">Nothing on the shopping list for this day yet. Add orders on the Plan tab.</div>'}`;
  }

  function exportXlsx() {
    const rows = listRows();
    const table = SP.pickPlanTable(rows, state.comments);
    const dayTxt = SP.fmtDay(state.day * DAY);
    const iso = SP.dayToISO(state.day);
    const made = `Made ${SP.fmtDayTime(state.now)} from ${state.pickingName} and ${state.transportName}`;
    const head = SP.PICK_PLAN_HEAD.map((h) => ({ v: h, s: 'head' }));
    const lines = table.map(({ row, cells }) => {
      const xd = row.slot.role === 'xdock' && !row.chosen.length;
      const base = xd ? 'xdock' : 'body';
      const st = !cells[6] ? base : /^picked/i.test(cells[6]) ? 'picked' : /progress/i.test(cells[6]) ? 'picking' : /^to release/i.test(cells[6]) ? 'release' : 'released';
      return cells.map((v, i) => ({ v, s: i === 6 ? st : xd ? 'xdock' : i === 0 ? 'bold' : i >= 2 && i <= 4 ? (i === 2 ? 'time' : 'body') : base }));
    });
    const heights = [24, 16, 8, 30].concat(table.map(({ row }) => Math.max(18, 15 * Math.max(1, row.chosen.length, row.waves.length))));
    const sheet1 = {
      name: 'Pick Plan', cols: [16, 24, 10, 10, 10, 34, 18, 30, 16, 10, 10], freezeRows: 4,
      rows: [[{ v: `Pick Plan – ${dayTxt}`, s: 'title' }], [{ v: made, s: 'sub' }], [], head, ...lines], heights,
      autoFilter: `A4:K${4 + lines.length}`,
    };
    const ohead = ['Picked By', 'Trailer', 'Seller Order ID', 'Order ID', 'Wave ID', 'Customer', 'Due', 'Status', 'Task Release', 'Picks left', 'Total picks', 'Note'];
    const olines = [];
    for (const row of rows) for (const o of row.chosen) {
      const st = SP.orderState(o);
      olines.push([
        { v: t(row.slot.pickBy), s: 'time' }, { v: laneName(row.slot), s: 'body' }, { v: o.seller, s: 'bold' }, { v: o.id.replace(/^0+/, ''), s: 'body' },
        { v: o.wave, s: 'body' }, { v: o.customer || o.company, s: 'body' }, { v: SP.fmtDayTime(o.delivery), s: 'body' },
        { v: SP.STATE_TEXT[st], s: st === 'picked' ? 'picked' : st === 'picking' ? 'picking' : st === 'release' ? 'release' : 'released' },
        { v: o.release, s: 'body' }, { v: o.remaining, s: 'body' }, { v: o.total, s: 'body' },
        { v: (row.notes.get(o.id) || []).map((n) => n.text).join('; '), s: 'body' },
      ]);
    }
    const sheet2 = { name: 'Orders', cols: [10, 30, 30, 12, 12, 18, 16, 16, 14, 10, 10, 36], freezeRows: 1, rows: [ohead.map((h) => ({ v: h, s: 'head' })), ...olines], autoFilter: `A1:L${1 + olines.length}` };
    const bytes = window.SPXlsx.build([sheet1, sheet2], XLSX.CFB);
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Pick Plan ${iso.slice(8, 10)}-${iso.slice(5, 7)}-${iso.slice(0, 4)}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    toast('Excel file downloaded');
  }

  function copyList() {
    const table = SP.pickPlanTable(listRows(), state.comments);
    const clean = (v) => String(v == null ? '' : v).replace(/\n/g, ', ').replace(/\t/g, ' ');
    const text = [SP.PICK_PLAN_HEAD.join('\t'), ...table.map(({ cells }) => cells.map(clean).join('\t'))].join('\n');
    const fallback = () => {
      const panel = document.getElementById('panel');
      const ta = document.createElement('textarea');
      ta.className = 'copy-fallback'; ta.value = text; ta.readOnly = true;
      panel.prepend(ta); ta.focus(); ta.select();
      toast('Press Ctrl+C to copy');
    };
    try { navigator.clipboard.writeText(text).then(() => toast('Copied – paste into Excel or Teams'), fallback); } catch (e) { fallback(); }
  }

  /* ------------------------------ checks ---------------------------- */
  const EX_GROUPS = [
    ['decide', 'Need a size', 'Set the trailer share for these orders on the Orders tab.'],
    ['over', 'Suggested trailers over capacity', 'More orders are due than fit on one trailer. Ask transport for another trailer or split them.'],
    ['big', 'Orders bigger than one trailer', ''],
    ['noslot', 'No trailer at the due time', 'The transport plan has no trailer for this customer at this time.'],
    ['missed', 'Missed trailers', 'The trailer has gone but picks are still open.'],
    ['xdock', 'Could share with X-Dock', 'Our trailer is under target and an X-Dock trailer goes to the same customer around the same time.'],
    ['free', 'Trailers maybe not needed', 'Every order for this trailer is suggested for an earlier one. Check with transport before cancelling.'],
    ['customer', 'Unknown customers', 'Add a pattern for them under Rules › Customers.'],
    ['collection', 'Customer collections', 'Not planned on our trailers.'],
    ['estimate', 'Sizes estimated', 'Fully picked orders lose their area in the export, so they are sized as Wide.'],
  ];
  function viewChecks(r) {
    const ref = state.day;
    const listed = dayRows().filter((row) => row.chosen.length && row.warnings.length && custOk(row.slot.customer));
    const mine = listed.length ? `<section class="ex-group"><h3>On your shopping list <span class="muted small">${listed.length}</span></h3>
      <ul class="ex-list">${listed.map((row) => `<li><span class="what">${esc(laneName(row.slot))}</span><span>${row.warnings.map((w) => esc(w.text)).join(' · ')}</span></li>`).join('')}</ul></section>` : '';
    const parts = EX_GROUPS.map(([kind, title, help]) => {
      const items = r.exceptions.filter((e) => e.kind === kind && custOk(e.order ? e.order.customer : e.slot.customer));
      if (!items.length) return '';
      return `<section class="ex-group"><h3>${title} <span class="muted small">${items.length}</span></h3>${help ? `<p class="muted small">${help}</p>` : ''}
        <ul class="ex-list">${items.map((e) => e.order
          ? `<li><span class="what">${e.orders.length > 1 ? `${e.orders.length} orders on wave ${esc(e.order.wave)}` : esc(e.order.seller || e.order.id)}</span><span class="muted small mono">${esc(e.order.customer || e.order.company)} · due ${when(e.order.delivery, ref)}${e.orders.length > 1 ? '' : ` · order ${esc(e.order.id.replace(/^0+/, ''))}`}</span><span>${esc(e.text)}</span></li>`
          : `<li><span class="what">${esc(e.slot.customer)} ${when(e.slot.delivery, ref)} ${esc(e.slot.kind)}</span><span>${esc(e.text)}</span></li>`).join('')}</ul></section>`;
    }).join('');
    return (mine + parts) || `<div class="empty-state">Nothing to check.</div>`;
  }

  /* ------------------------------ orders ---------------------------- */
  function viewOrders(r) {
    const ref = state.day;
    const rows = rowByKey();
    const byId = new Map(r.board.map((s) => [s.id, s]));
    const list = r.orders.filter((o) => custOk(o.customer)).sort((a, b) => a.delivery - b.delivery || (a.customer || '').localeCompare(b.customer || ''));
    return `<p class="muted small">Every order in the picking status (kitting and decant left out). Set a trailer share by hand when the estimate is wrong, for example Landpower engines that can't be stacked.</p>
      <div class="table-box"><table>
      <thead><tr><th>Due</th><th>Customer</th><th>Order</th><th>Status</th><th>Picks left</th><th class="num">Pallets</th><th class="num">Trailer share</th><th>Set %</th><th>Shopping list</th></tr></thead>
      <tbody>${list.map((o) => {
        const a = state.assign[o.id];
        const row = a && rows.get(a.k);
        const sug = !row && o.slot && byId.get(o.slot);
        const ov = state.overrides[o.id] || {};
        const st = SP.orderState(o);
        const on = row ? `<span class="chip info">${esc(laneName(row.slot))}</span> <button class="btn small ghost" data-unassign="${esc(o.id)}">Remove</button>`
          : o.collection ? '<span class="chip">Collection</span>'
          : sug && !sug.past ? `<button class="btn small" data-add="${esc(o.id)}" data-to="${esc(sug.key)}">Add to ${esc(laneName(sug))}</button>`
          : '<span class="chip low">No trailer</span>';
        return `<tr>
          <td class="mono">${when(o.delivery, ref)}</td>
          <td>${esc(o.customer || o.company)}${o.decide ? ' <span class="chip under">Size?</span>' : ''}</td>
          <td><span class="ref">${esc(o.seller || o.id)}</span><div class="small muted mono">Order ${esc(o.id.replace(/^0+/, ''))} · Wave ${esc(o.wave)}</div>${o.comments ? `<div class="small">${esc(o.comments)}</div>` : ''}</td>
          <td class="small"><span class="dot ${STATE_CLASS[st]}"></span> ${esc(SP.STATE_TEXT[st])}<div class="muted">${esc(o.status)}</div></td>
          <td><span class="mono">${o.remaining} / ${o.total}</span>${areaChips(o)}</td>
          <td class="num mono">${pal(o.pallets)}</td>
          <td class="num mono" title="${esc(o.basis)}">${pct(o.frac)}</td>
          <td><input type="number" min="0" max="400" step="5" data-pct="${esc(o.id)}" value="${ov.pct != null ? esc(ov.pct) : ''}" placeholder="auto" aria-label="Trailer share for ${esc(o.seller)}"></td>
          <td>${on}</td></tr>`;
      }).join('')}</tbody></table></div>`;
  }
  function areaChips(o) {
    const e = Object.entries(o.areas);
    if (!e.length) return '<span class="muted small">all picked</span>';
    return `<span class="areas">${e.map(([a, n]) => `<span class="area">${esc(a)} ${n}</span>`).join('')}</span>`;
  }

  /* ------------------------------ rules ----------------------------- */
  function viewRules() {
    const c = state.cfg;
    const numField = (id, label, val, step, help) => `<label class="field"><span class="label">${label}</span><input type="number" step="${step}" data-cfg="${id}" value="${val}">${help ? `<span class="small muted">${help}</span>` : ''}</label>`;
    const chk = (id, label) => `<label class="check"><input type="checkbox" data-cfg-bool="${id}" ${c[id] ? 'checked' : ''}> ${label}</label>`;
    const txt = (id, label, help) => `<label class="field"><span class="label">${label}</span><input type="text" data-cfg-text="${id}" value="${esc(c[id] || '')}"><span class="small muted">${help}</span></label>`;
    const days = state.transport.days;
    return `<div class="rules">
      <section class="card"><h3>Times and suggestions</h3>
        <p>"Picked by" is the transport plan's paperwork time ("Loaded by") minus the time below. Suggestions top up a trailer that already has orders until it reaches the target, without going over the maximum.</p>
        <div class="grid2">
          ${numField('pickLeadMinutes', 'Pick by (min before loading)', c.pickLeadMinutes, 15)}
          ${numField('targetFill%', 'Target fill %', Math.round(c.targetFill * 100), 5)}
          ${numField('maxFill%', 'Maximum fill %', Math.round(c.maxFill * 100), 5)}
          ${numField('lowFill%', 'Flag red below %', Math.round(c.lowFill * 100), 5)}
          ${numField('horizonHours', 'Look ahead (hours)', c.horizonHours, 1)}
          ${numField('leadMinutes', 'Stop suggesting (min before loading)', c.leadMinutes, 15)}
        </div>
        ${chk('sameDayOnly', 'Only pull orders delivering the same day')}
        ${chk('protectSources', "Don't take orders off trailers that are already full enough")}
        ${chk('includeKitting', 'Include kitting orders (Cabs kits, Landpower kits, T4)')}
        ${chk('useXdock', 'Suggest X-Dock trailers as if they were ours')}
        ${txt('ignoreOrders', 'Leave out orders matching', 'Order type or seller order ID, e.g. decant')}
      </section>
      <section class="card"><h3>Pallets by picking area</h3>
        <p>1 pick is 1 pallet, except box areas. A mixed order adds up its share of a trailer from each area. Untick "Count" for columns that repeat picks counted elsewhere (P&D is already in VNA).</p>
        <div class="table-box"><table><thead><tr><th>Area</th><th>Column pattern</th><th class="num">Picks / pallet</th><th class="num">Pallets / trailer</th><th>Count</th></tr></thead><tbody>
        ${c.areas.map((a, i) => `<tr><td>${esc(a.name)}</td><td><input type="text" data-area="${i}" data-k="match" value="${esc(a.match)}"></td><td class="num"><input type="number" min="1" data-area="${i}" data-k="perPallet" value="${a.perPallet}"></td><td class="num"><input type="number" min="1" data-area="${i}" data-k="perTrailer" value="${a.perTrailer}"></td><td><input type="checkbox" data-area="${i}" data-k="count" ${a.count === false ? '' : 'checked'} aria-label="Count ${esc(a.name)}"></td></tr>`).join('')}
        </tbody></table></div>
        ${numField('defaultPerTrailer', 'Pallets per trailer, any other area', c.defaultPerTrailer, 1)}
      </section>
      <section class="card"><h3>Order rules</h3>
        <p>For orders that don't follow the area numbers. The first rule that matches wins. Use "Pallets / trailer" for big items (16 shovels fill a trailer), or "Orders / pallet" for kits. "Decide" asks you to set the share by hand.</p>
        <div class="table-box"><table><thead><tr><th>Name</th><th>Look in</th><th>Contains</th><th class="num">Pallets / trailer</th><th class="num">Orders / pallet</th><th>Decide</th><th></th></tr></thead><tbody>
        ${c.rules.map((ru, i) => `<tr><td><input type="text" data-rule="${i}" data-k="name" value="${esc(ru.name)}"></td>
          <td><select data-rule="${i}" data-k="field">${['seller', 'company', 'type', 'comments'].map((f) => `<option value="${f}" ${ru.field === f ? 'selected' : ''}>${{ seller: 'Seller order ID', company: 'Company', type: 'Order type', comments: 'Comments' }[f]}</option>`).join('')}</select></td>
          <td><input type="text" data-rule="${i}" data-k="match" value="${esc(ru.match)}"></td>
          <td class="num"><input type="number" min="0" data-rule="${i}" data-k="cap" value="${ru.cap || ''}" placeholder="–"></td>
          <td class="num"><input type="number" min="0" data-rule="${i}" data-k="perPallet" value="${ru.perPallet || ''}" placeholder="–"></td>
          <td><input type="checkbox" data-rule="${i}" data-k="decide" ${ru.decide ? 'checked' : ''}></td>
          <td><button class="btn ghost danger" data-del-rule="${i}" aria-label="Remove rule">Remove</button></td></tr>`).join('')}
        </tbody></table></div>
        <div><button class="btn" data-act="add-rule">Add rule</button></div>
      </section>
      <section class="card"><h3>Customers</h3>
        <p>Links the company name in the picking status to the "Deliver To" name in the transport plan. The first pattern that matches wins, so put specific names first.</p>
        <div class="table-box"><table><thead><tr><th>Customer</th><th>Pattern</th><th></th></tr></thead><tbody>
        ${c.customers.map(([n, p], i) => `<tr><td><input type="text" data-cust="${i}" data-k="0" value="${esc(n)}"></td><td><input type="text" data-cust="${i}" data-k="1" value="${esc(p)}"></td><td><button class="btn ghost danger" data-del-cust="${i}" aria-label="Remove customer">Remove</button></td></tr>`).join('')}
        </tbody></table></div>
        <div><button class="btn" data-act="add-cust">Add customer</button></div>
      </section>
      <section class="card"><h3>Transport plan</h3>
        ${txt('ourSite', 'Our trailers come from', 'Pattern for "Delivery / Collection From". Other sites’ trailers are shown but not suggested.')}
        ${txt('skipKinds', 'Ignore these delivery types', 'Not customer trailers, e.g. transfers and decant')}
        ${txt('noShopKinds', 'Don’t suggest extra orders for', 'These keep their own orders but take nothing extra')}
        ${txt('xdockKinds', 'X-Dock trailers', 'Delivery types run by the X-Dock department')}
        ${txt('collectionMatch', 'Customer collections', 'Matched against seller order ID and comments')}
        <label class="field"><span class="label">Transport day to use</span>
          <select data-cfg-text="templateDay"><option value="">Automatic (same date, else latest Mon–Thu; Fridays use a Friday)</option>
          ${days.map((d) => `<option value="${SP.dayToISO(d.day)}" ${c.templateDay === SP.dayToISO(d.day) ? 'selected' : ''}>${SP.fmtDay(d.day * DAY)} ${esc(d.title)}</option>`).join('')}</select></label>
      </section>
      <section class="card"><h3>Reset</h3>
        <p>Rules, the shopping list and comments are saved in this browser only.</p>
        <div><button class="btn danger" data-act="reset" id="reset-btn">Reset rules to default</button></div>
      </section>
    </div>`;
  }

  /* ----------------------------- render ----------------------------- */
  function render() {
    compute();
    const r = state.result;
    const views = { plan: viewPlan, list: viewList, checks: viewChecks, orders: viewOrders, rules: viewRules };
    app.innerHTML = viewTop(r) + viewControls(r) + viewKpis() + viewTabs() + `<main class="panel tab-${state.tab}" id="panel">${views[state.tab](r)}</main>`;
  }
  function rerender() {
    const y = window.scrollY;
    const pool = document.querySelector('.pool');
    const py = pool ? pool.scrollTop : 0;
    render();
    window.scrollTo(0, y);
    const p2 = document.querySelector('.pool');
    if (p2) p2.scrollTop = py;
  }

  /* ----------------------------- events ----------------------------- */
  let armed = null;
  function arm(b, label, armedLabel) {
    if (armed === b.id) return true;
    armed = b.id; b.textContent = armedLabel;
    setTimeout(() => { if (armed === b.id) armed = null; if (b.isConnected) b.textContent = label; }, 4000);
    return false;
  }

  app.addEventListener('click', (e) => {
    const b = e.target.closest('[data-act],[data-tab],[data-del-rule],[data-del-cust],[data-add],[data-unassign],[data-add-all]');
    if (!b) return;
    const d = b.dataset;
    if (d.tab) { state.tab = d.tab; render(); return; }
    if (d.add) {
      const o = orderById().get(d.add), row = rowByKey().get(d.to);
      assign([d.add], d.to, `${o ? o.seller : 'Order'} → ${row ? laneName(row.slot) : 'trailer'}`);
      return;
    }
    if (d.unassign) { const o = orderById().get(d.unassign); assign([d.unassign], null, `${o ? o.seller : 'Order'} taken off the list`); return; }
    if (d.addAll) {
      const row = rowByKey().get(d.addAll);
      assign(row.suggested.map((x) => x.order.id), row.key, `${plural(row.suggested.length, 'order')} → ${laneName(row.slot)}`);
      return;
    }
    if (d.delRule != null) { state.cfg.rules.splice(+d.delRule, 1); saveCfg(); rerender(); return; }
    if (d.delCust != null) { state.cfg.customers.splice(+d.delCust, 1); saveCfg(); rerender(); return; }
    const act = d.act;
    if (act === 'load') fileInput.click();
    else if (act === 'sample') { loadSample(); state.message = null; render(); }
    else if (act === 'now-file') { state.nowFromFile = true; render(); }
    else if (act === 'undo') doUndo();
    else if (act === 'add-all') {
      snapshot();
      const at = Date.now();
      let n = 0;
      for (const row of visibleRows()) for (const x of row.suggested) { state.assign[x.order.id] = { k: row.key, at }; n++; }
      saveAssign(); rerender(); toast(`${plural(n, 'order')} added to the list`, true);
    } else if (act === 'clear-day') {
      if (!arm(b, 'Clear this day', 'Click again to clear')) return;
      armed = null;
      const keys = new Set(dayRows().filter((r) => custOk(r.slot.customer)).map((r) => r.key));
      const ids = Object.keys(state.assign).filter((id) => keys.has(state.assign[id].k));
      assign(ids, null, `${plural(ids.length, 'order')} taken off the list`);
    } else if (act === 'copy') copyList();
    else if (act === 'print') window.print();
    else if (act === 'xlsx') {
      try { exportXlsx(); } catch (err) { state.message = { kind: 'err', text: `Could not make the Excel file: ${err.message}` }; render(); }
    } else if (act === 'add-rule') { state.cfg.rules.push({ name: 'New rule', field: 'seller', match: '', cap: 35 }); saveCfg(); rerender(); }
    else if (act === 'add-cust') { state.cfg.customers.push(['New customer', '']); saveCfg(); rerender(); }
    else if (act === 'reset') {
      if (!arm(b, 'Reset rules to default', 'Click again to reset')) return;
      armed = null; state.cfg = clone(SP.DEFAULT_CONFIG); saveCfg(); toast('Rules reset'); render();
    }
  });

  app.addEventListener('change', (e) => {
    const el = e.target;
    const d = el.dataset;
    if (d.addTo) { if (el.value) { const o = orderById().get(el.value); assign([el.value], d.addTo, `${o ? o.seller : 'Order'} added`); } return; }
    if (d.move) { if (el.value) { const row = rowByKey().get(el.value); assign([d.move], el.value, `Moved to ${row ? laneName(row.slot) : 'trailer'}`); } return; }
    if (d.comment != null) {
      if (el.value.trim()) state.comments[d.comment] = el.value.trim(); else delete state.comments[d.comment];
      saveComments();
      return; // no re-render, so typing isn't interrupted
    }
    if (el.id === 'now') { const v = SP.localInputToAbs(el.value); if (v != null) { state.now = v; state.nowFromFile = false; } }
    else if (el.id === 'day') state.day = +el.value;
    else if (el.id === 'customer') state.customer = el.value;
    else if (el.id === 'show-other') state.showOther = el.checked;
    else if (el.id === 'show-empty') state.showEmpty = el.checked;
    else if (el.id === 'show-past') state.showPast = el.checked;
    else if (el.id === 'list-empty') state.listEmpty = el.checked;
    else if (d.pct) {
      const o = state.overrides[d.pct] || {};
      if (el.value === '') delete o.pct; else o.pct = Math.max(0, +el.value);
      state.overrides[d.pct] = o; saveOverrides();
    } else if (d.cfg) {
      const v = +el.value;
      if (!isFinite(v)) return;
      if (d.cfg.endsWith('%')) state.cfg[d.cfg.slice(0, -1)] = v / 100; else state.cfg[d.cfg] = v;
      saveCfg();
    } else if (d.cfgBool) { state.cfg[d.cfgBool] = el.checked; saveCfg(); }
    else if (d.cfgText) { state.cfg[d.cfgText] = el.value || (d.cfgText === 'templateDay' ? null : ''); saveCfg(); }
    else if (d.area != null) { const a = state.cfg.areas[+d.area]; a[d.k] = d.k === 'count' ? el.checked : d.k === 'match' ? el.value : Math.max(1, +el.value || 1); saveCfg(); }
    else if (d.rule != null) { const ru = state.cfg.rules[+d.rule]; ru[d.k] = d.k === 'decide' ? el.checked : (d.k === 'cap' || d.k === 'perPallet') ? (+el.value || 0) : el.value; saveCfg(); }
    else if (d.cust != null) { state.cfg.customers[+d.cust][+d.k] = el.value; saveCfg(); }
    else return;
    rerender();
  });

  /* drag orders between trailers and the pool */
  app.addEventListener('dragstart', (e) => {
    const el = e.target.closest && e.target.closest('[data-order]');
    if (!el) return;
    e.dataTransfer.setData('application/x-order', el.dataset.order);
    e.dataTransfer.effectAllowed = 'move';
    el.classList.add('dragging');
  });
  app.addEventListener('dragend', (e) => {
    const el = e.target.closest && e.target.closest('[data-order]');
    if (el) el.classList.remove('dragging');
    app.querySelectorAll('.drop-over').forEach((x) => x.classList.remove('drop-over'));
  });
  const isOrderDrag = (e) => e.dataTransfer && [...(e.dataTransfer.types || [])].includes('application/x-order');
  app.addEventListener('dragover', (e) => {
    const zone = e.target.closest && e.target.closest('[data-drop]');
    if (!zone || !isOrderDrag(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    app.querySelectorAll('.drop-over').forEach((x) => { if (x !== zone) x.classList.remove('drop-over'); });
    zone.classList.add('drop-over');
  });
  app.addEventListener('drop', (e) => {
    const zone = e.target.closest && e.target.closest('[data-drop]');
    if (!zone || !isOrderDrag(e)) return;
    e.preventDefault();
    e.stopPropagation();
    const id = e.dataTransfer.getData('application/x-order');
    const key = zone.dataset.drop;
    const cur = state.assign[id] && state.assign[id].k;
    if (key === 'pool') { if (cur) assign([id], null, 'Taken off the list'); }
    else if (cur !== key) { const row = rowByKey().get(key); assign([id], key, `Moved to ${row ? laneName(row.slot) : 'trailer'}`); }
    app.querySelectorAll('.drop-over').forEach((x) => x.classList.remove('drop-over'));
  });

  /* ------------------------------ boot ------------------------------ */
  loadSample();
  render();
})();
