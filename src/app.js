/* Shopping Plan – browser UI. Needs XLSX (SheetJS), SPCore and SPSample loaded first. */
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

  /* ------------------------------ state ----------------------------- */
  const state = {
    picking: null, pickingName: '', pickingSample: true,
    transport: null, transportName: '', transportSample: true,
    cfg: Object.assign(clone(SP.DEFAULT_CONFIG), store.get('sp.cfg.v2', {})),
    overrides: store.get('sp.overrides.v1', {}),
    done: store.get('sp.done.v1', {}),
    now: null, nowFromFile: true,
    tab: 'shopping', customer: '', showEmpty: false, showPast: false,
    message: null, result: null,
  };
  const saveCfg = () => store.set('sp.cfg.v2', state.cfg);
  const saveOverrides = () => store.set('sp.overrides.v1', state.overrides);
  const saveDone = () => store.set('sp.done.v1', state.done);

  function loadSample() {
    state.picking = SP.parsePicking(window.SPSample.picking());
    state.transport = SP.parseTransport(window.SPSample.transport());
    state.pickingName = 'Example picking status';
    state.transportName = 'Example transport plan';
    state.pickingSample = state.transportSample = true;
    state.nowFromFile = true;
  }

  /* ---------------------------- helpers ----------------------------- */
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const pct = (f) => (f > 0 && f < 0.005 ? '<1%' : `${Math.round(f * 100)}%`);
  const pal = (n) => (n >= 10 ? Math.round(n) : n >= 1 ? n.toFixed(1).replace(/\.0$/, '') : n > 0 ? '<1' : '0');
  const t = SP.fmtTime;
  function when(abs, refDay) {
    if (abs == null) return '–';
    return Math.floor(abs / DAY) === refDay ? t(abs) : `${SP.fmtDay(abs).slice(0, 3)} ${t(abs)}`;
  }
  const VERDICT = { ok: 'Full enough', under: 'Under target', low: 'Low fill', over: 'Over capacity', empty: 'No orders', past: 'Gone' };
  const VCLASS = { ok: 'ok', under: 'under', low: 'low', over: 'over', empty: '', past: '' };
  const MOVE = { shop: ['Pulled forward', 'info'], noslot: ['No trailer at due time', 'under'], roll: ['Missed trailer', 'low'] };
  function areaChips(o) {
    const e = Object.entries(o.areas);
    if (!e.length) return '<span class="muted small">all picked</span>';
    return `<span class="areas">${e.map(([a, n]) => `<span class="area">${esc(a)} ${n}</span>`).join('')}</span>`;
  }
  function action(o, slot) {
    if (o.remaining <= 0) return `Already picked – stage for the ${t(slot.delivery)} trailer`;
    if (/not released/i.test(o.release)) return `Release wave ${esc(o.wave)} and pick`;
    return `Prioritise remaining picks`;
  }
  function toast(msg) {
    const el = document.createElement('div');
    el.className = 'toast';
    el.textContent = msg;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 2200);
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
        state.picking = d.data; state.pickingName = f.name; state.pickingSample = false; state.nowFromFile = true;
      } else if (d.type === 'transport') {
        state.transport = d.data; state.transportName = f.name; state.transportSample = false;
      } else {
        problems.push(`"${f.name}" ${d.error ? '– ' + d.error : 'was not recognised. A picking status needs Delivery Date and Order ID columns; a transport plan needs Deliver To and Delivery Time columns.'}`);
      }
    }
    state.message = problems.length ? { kind: 'err', text: problems.join(' ') } : null;
    if (!problems.length) toast(`Loaded ${files.length} file${files.length > 1 ? 's' : ''}`);
    render();
  }
  fileInput.addEventListener('change', () => { if (fileInput.files.length) loadFiles([...fileInput.files]); fileInput.value = ''; });
  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => { e.preventDefault(); dragDepth++; document.body.classList.add('drag'); });
  window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('drag'); } });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    e.preventDefault(); dragDepth = 0; document.body.classList.remove('drag');
    if (e.dataTransfer && e.dataTransfer.files.length) loadFiles([...e.dataTransfer.files]);
  });

  /* ----------------------------- compute ---------------------------- */
  function compute() {
    if (state.nowFromFile || state.now == null) state.now = state.picking.asOf != null ? state.picking.asOf : Math.floor(Date.now() / 60000) - new Date().getTimezoneOffset();
    state.result = SP.plan(state.picking, state.transport, state.cfg, { now: state.now, overrides: state.overrides });
  }

  /* ------------------------------ views ----------------------------- */
  function viewTop(r) {
    const p = state.picking, tr = state.transport;
    const used = r.used.map((u) => `${SP.fmtDay(u.day * DAY)} ← ${u.exact ? 'same day' : 'template ' + SP.fmtDay(u.source * DAY)}`).join(' · ');
    const sample = state.pickingSample || state.transportSample;
    return `
      <header class="topbar">
        <div class="brand"><h1>Shopping Plan</h1><span class="sub">Fill every trailer before it leaves</span></div>
        <span class="spacer"></span>
        <button class="btn primary" data-act="load">Load files</button>
        ${sample ? '' : '<button class="btn ghost" data-act="sample">Show example</button>'}
      </header>
      ${sample ? `<div class="banner"><strong>Example data.</strong> These orders and trailers are made up. Load your Picking Status export and transport plan (or drop both files on this page) to plan for real.</div>` : ''}
      ${state.message ? `<div class="banner ${state.message.kind}">${esc(state.message.text)}</div>` : ''}
      <section class="sources dropzone">
        <div class="source ${state.pickingSample ? 'is-sample' : ''}">
          <span class="label">Picking status</span>
          <span class="name">${esc(state.pickingName)}</span>
          <span class="small muted">${p.orders.length} orders${p.asOf != null ? ` · exported ${SP.fmtDayTime(p.asOf)}` : ''}</span>
        </div>
        <div class="source ${state.transportSample ? 'is-sample' : ''}">
          <span class="label">Transport plan</span>
          <span class="name">${esc(state.transportName)}</span>
          <span class="small muted">${tr.slots.length} rows · trailers for ${esc(used || 'no matching day')}</span>
        </div>
      </section>
      ${r.used.some((u) => !u.exact) ? `<div class="banner warn">No sheet in the transport plan for ${r.used.filter((u) => !u.exact).map((u) => `${SP.fmtDay(u.day * DAY)} (using ${SP.fmtDay(u.source * DAY)})`).join(', ')}. The schedule repeats every day except Friday, so trailer times are taken from another day. Load today's plan to also match the wave numbers written on it.</div>` : ''}
      ${r.unconfirmedAreas.length ? `<div class="banner warn">Trailer capacity for ${r.unconfirmedAreas.map(esc).join(', ')} is a guess. Check it under Rules.</div>` : ''}
    `;
  }

  function viewControls(r) {
    const customers = [...new Set(r.slots.map((s) => s.customer).concat(r.orders.map((o) => o.customer).filter(Boolean)))].sort();
    return `
      <section class="controls">
        <label class="field"><span class="label">Plan from</span>
          <input type="datetime-local" id="now" value="${SP.absToLocalInput(state.now)}"></label>
        ${state.nowFromFile ? '' : '<button class="btn ghost" data-act="now-file">Use export time</button>'}
        <label class="field"><span class="label">Customer</span>
          <select id="customer"><option value="">All customers</option>${customers.map((c) => `<option ${c === state.customer ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></label>
        <label class="check"><input type="checkbox" id="show-empty" ${state.showEmpty ? 'checked' : ''}> Show trailers with no orders</label>
        <label class="check"><input type="checkbox" id="show-past" ${state.showPast ? 'checked' : ''}> Show trailers already gone</label>
      </section>`;
  }

  function viewKpis(r) {
    const s = r.summary;
    const lines = SP.groupShopping(r.shopping).length;
    const decide = r.exceptions.filter((e) => ['decide', 'over', 'big', 'noslot', 'missed', 'customer'].includes(e.kind)).length;
    return `
      <section class="kpis">
        <div class="kpi attn"><span class="label">Lines to pick early</span><span class="v">${lines}</span><span class="small muted">${s.moves} order${s.moves === 1 ? '' : 's'}, to fill trailers that leave sooner</span></div>
        <div class="kpi ${s.under ? 'warn' : ''}"><span class="label">Trailers under ${pct(r.cfg.targetFill)}</span><span class="v">${s.under} <small>was ${s.underBefore}</small></span><span class="small muted">of ${s.trailers - s.empty} running</span></div>
        <div class="kpi attn"><span class="label">Trailers maybe not needed</span><span class="v">${s.freed}</span><span class="small muted">all their orders go earlier</span></div>
        <div class="kpi ${decide ? 'bad' : ''}"><span class="label">Need a decision</span><span class="v">${decide}</span><span class="small muted">see Exceptions</span></div>
      </section>`;
  }

  function viewTabs(r) {
    const ex = r.exceptions.filter((e) => e.kind !== 'collection' && e.kind !== 'estimate').length;
    const tabs = [['shopping', 'Shopping list', SP.groupShopping(r.shopping).length], ['trailers', 'Trailers', null], ['exceptions', 'Exceptions', ex], ['orders', 'Orders', r.orders.length], ['rules', 'Rules', null]];
    return `<nav class="tabs" role="tablist">${tabs.map(([k, label, n]) => `<button class="tab" role="tab" data-tab="${k}" aria-selected="${state.tab === k}">${label}${n != null ? `<span class="count">${n}</span>` : ''}</button>`).join('')}</nav>`;
  }

  const custOk = (c) => !state.customer || c === state.customer;
  const planDay = () => Math.floor(state.now / DAY) + (state.now % DAY >= 12 * 60 ? 1 : 0);

  function viewShopping(r) {
    const rows = SP.groupShopping(r.shopping.filter((m) => custOk(m.order.customer)));
    if (!rows.length) return `<div class="empty-state">Nothing to shop. Every running trailer is full enough, or nothing later can go on it.</div>`;
    const ref = planDay();
    let lastKey = null;
    const body = rows.map((g) => {
      const o = g.order, s = g.slot, many = g.orders.length > 1;
      const key = s.paperwork != null ? s.paperwork : s.delivery;
      let group = '';
      if (key !== lastKey) {
        lastKey = key;
        group = `<tr class="group"><td colspan="8">Paperwork by ${when(key, ref)} – ${esc(s.customer)} ${t(s.delivery)} trailer, now ${pct(s.load)} full (≈ ${pal(s.pallets)} pallets)</td></tr>`;
      }
      const [mLabel, mClass] = MOVE[g.type] || ['Moved', ''];
      const ids = g.orders.map((x) => x.id);
      const done = ids.every((id) => state.done[id]);
      const areas = many ? '' : areaChips(o);
      const name = many
        ? `<span class="ref">${g.orders.length} ${esc(o.type.replace(/^\d+\s*-\s*/, '') || 'orders')} orders</span><div class="small muted mono">Wave ${esc(o.wave)} · ${esc(g.orders[0].seller)} … ${esc(g.orders[g.orders.length - 1].seller)}</div>`
        : `<span class="ref">${esc(o.seller || o.id)}</span><div class="small muted mono">Order ${esc(o.id.replace(/^0+/, ''))} · Wave ${esc(o.wave)}</div>`;
      return `${group}<tr class="${done ? 'done' : ''}">
        <td><input type="checkbox" data-done="${esc(ids.join(','))}" ${done ? 'checked' : ''} aria-label="Mark as sent to picking"></td>
        <td><span class="t">${when(key, ref)}</span><div class="small muted">paperwork</div></td>
        <td><strong>${esc(s.customer)} ${t(s.delivery)}</strong><div class="small muted">${esc(s.kind)}${s.despatch != null ? ` · desp ${t(s.despatch)}` : ''}</div></td>
        <td>${name}</td>
        <td><span class="mono">${when(o.delivery, ref)}</span><div><span class="chip ${mClass}">${mLabel}</span></div></td>
        <td><span class="mono">${g.remaining} / ${g.total}</span>${areas}</td>
        <td class="num mono">+${pct(g.frac)}<div class="small muted">≈ ${pal(g.pallets)} pal</div></td>
        <td class="action">${action(o, s)}</td>
      </tr>`;
    }).join('');
    return `
      <div class="panel-head">
        <p class="muted small">Ordered by when the trailer's paperwork is due. Orders on the same wave are one line. Tick a line once picking have it.</p>
        <button class="btn" data-act="copy">Copy for Excel / Teams</button>
      </div>
      <div class="table-box"><table>
        <thead><tr><th>Sent</th><th>Need by</th><th>Put on trailer</th><th>Order</th><th>Was due</th><th>Picks left</th><th class="num">Trailer</th><th>Action</th></tr></thead>
        <tbody>${body}</tbody></table></div>`;
  }

  function trailerCard(s, r, ref) {
    const target = r.cfg.targetFill;
    const own = Math.min(s.ownLoad, 1), shop = Math.min(Math.max(0, s.load - s.ownLoad), Math.max(0, 1 - own));
    const verdict = s.lostAll ? 'Maybe not needed' : VERDICT[s.verdict];
    const vclass = s.lostAll ? 'info' : VCLASS[s.verdict];
    const lines = [];
    const byWave = new Map();
    for (const o of s.orders.slice().sort((a, b) => a.delivery - b.delivery)) {
      const k = o.rule && /kit/i.test(o.rule) ? `${o.wave}|${!!o.moveType}|${o.delivery}` : o.id;
      if (byWave.has(k)) { const l = byWave.get(k); l.n++; l.frac += o.frac; l.rem += o.remaining; l.tot += o.total; continue; }
      const l = { o, n: 1, frac: o.frac, rem: o.remaining, tot: o.total };
      byWave.set(k, l); lines.push(l);
    }
    const lis = lines.map(({ o, n, frac, rem, tot }) => {
      const moved = o.moveType ? ` <span class="muted">(due ${when(o.delivery, ref)})</span>` : '';
      const who = n > 1 ? `${n} kit orders, wave ${esc(o.wave)}` : esc(o.seller || o.id);
      return `<li class="${o.moveType ? 'in' : ''}"><span class="who">${who}${moved}</span><span class="pct">${rem}/${tot} · ${pct(frac)}</span></li>`;
    }).join('');
    const gone = r.moves.filter((m) => m.from === s.id && m.order.slot !== s.id);
    const goneTxt = gone.length ? `<p class="small muted">${gone.length} order${gone.length > 1 ? 's' : ''} moved to earlier trailers.</p>` : '';
    return `<article class="trailer v-${s.verdict} ${s.lostAll ? 'freed' : ''}">
      <header><span class="t">${when(s.delivery, ref)}</span><span class="kind">${esc(s.kind || 'Trailer')}</span><span class="chip ${vclass}">${verdict}</span></header>
      <div class="fill-row"><div class="bar ${s.load > r.cfg.maxFill ? 'over' : ''}" title="Solid: own orders · striped: shopped in"><span class="own" style="width:${own * 100}%"></span><span class="shop" style="width:${shop * 100}%"></span><i class="target" style="left:${target * 100}%"></i></div><span class="pct">${pct(s.load)}</span></div>
      <dl><div><dt>Paperwork</dt><dd>${t(s.paperwork)}</dd></div><div><dt>Despatch</dt><dd>${t(s.despatch)}</dd></div><div><dt>Pallets</dt><dd>≈ ${pal(s.pallets)}</dd></div><div><dt>Picks left</dt><dd>${s.picksLeft}</dd></div>${s.trailer ? `<div><dt>Trailer</dt><dd>${esc(s.trailer)}</dd></div>` : ''}</dl>
      ${lis ? `<ul>${lis}</ul>` : ''}${goneTxt}
    </article>`;
  }

  function viewTrailers(r) {
    const ref = planDay();
    const show = (s) => (state.showPast || !s.past) && (state.showEmpty || s.orders.length || s.lostAll) && custOk(s.customer);
    const groups = new Map();
    r.slots.filter(show).forEach((s) => { if (!groups.has(s.customer)) groups.set(s.customer, []); groups.get(s.customer).push(s); });
    if (!groups.size) return `<div class="empty-state">No trailers to show. Tick "Show trailers with no orders" to see the whole transport plan.</div>`;
    const order = [...groups.keys()].sort((a, b) => groups.get(a)[0].delivery - groups.get(b)[0].delivery);
    return `<p class="muted small">Solid bar: the trailer's own orders. Striped: orders shopped in from later. The line marks the ${pct(r.cfg.targetFill)} target.</p>` +
      order.map((c) => {
        const list = groups.get(c);
        const running = list.filter((s) => s.orders.length);
        return `<section class="cust"><div class="cust-head"><h2>${esc(c)}</h2><span class="muted small">${running.length} trailer${running.length === 1 ? '' : 's'} with orders</span></div>
          <div class="board">${list.map((s) => trailerCard(s, r, ref)).join('')}</div></section>`;
      }).join('');
  }

  const EX_GROUPS = [
    ['decide', 'Need a decision', 'Set the trailer share for these orders on the Orders tab.'],
    ['over', 'Trailers over capacity', 'More orders than fit on one trailer. Ask transport for another trailer or use the unused slot named.'],
    ['big', 'Orders bigger than one trailer', ''],
    ['noslot', 'No trailer at the due time', 'The transport plan has no trailer for this customer at this time.'],
    ['missed', 'Missed trailers', 'The trailer has gone but picks are still open.'],
    ['xdock', 'Could share with X-Dock', 'Our trailer is under target and an X-Dock trailer goes to the same customer around the same time.'],
    ['free', 'Trailers maybe not needed', 'All orders moved to earlier trailers. Check with transport before cancelling.'],
    ['customer', 'Unknown customers', 'Add a pattern for them under Rules › Customers.'],
    ['collection', 'Customer collections', 'Not planned on our trailers.'],
    ['estimate', 'Sizes estimated', 'Fully picked orders lose their area in the export, so they are sized as Wide.'],
  ];
  function viewExceptions(r) {
    const ref = planDay();
    const parts = EX_GROUPS.map(([kind, title, help]) => {
      const items = r.exceptions.filter((e) => e.kind === kind && custOk(e.order ? e.order.customer : e.slot.customer));
      if (!items.length) return '';
      return `<section class="ex-group"><h3>${title} <span class="muted small">${items.length}</span></h3>${help ? `<p class="muted small">${help}</p>` : ''}
        <ul class="ex-list">${items.map((e) => e.order
          ? `<li><span class="what">${e.orders.length > 1 ? `${e.orders.length} orders on wave ${esc(e.order.wave)}` : esc(e.order.seller || e.order.id)}</span><span class="muted small mono">${esc(e.order.customer || e.order.company)} · due ${when(e.order.delivery, ref)}${e.orders.length > 1 ? '' : ` · order ${esc(e.order.id.replace(/^0+/, ''))}`}</span><span>${esc(e.text)}</span></li>`
          : `<li><span class="what">${esc(e.slot.customer)} ${when(e.slot.delivery, ref)} ${esc(e.slot.kind)}</span><span>${esc(e.text)}</span></li>`).join('')}</ul></section>`;
    }).join('');
    return parts || `<div class="empty-state">No exceptions.</div>`;
  }

  function viewOrders(r) {
    const ref = planDay();
    const slotById = new Map(r.slots.map((s) => [s.id, s]));
    const rows = r.orders.filter((o) => custOk(o.customer)).sort((a, b) => a.delivery - b.delivery || (a.customer || '').localeCompare(b.customer || ''));
    return `<p class="muted small">Set a trailer share by hand (for example Landpower engines that can't be stacked, or rubber-track machines). Tick "Keep" to stop an order being moved.</p>
      <div class="table-box"><table>
      <thead><tr><th>Due</th><th>Customer</th><th>Order</th><th>Status</th><th>Picks left</th><th class="num">Pallets</th><th class="num">Trailer share</th><th>Set %</th><th>Keep</th><th>Planned on</th></tr></thead>
      <tbody>${rows.map((o) => {
        const s = o.slot && slotById.get(o.slot);
        const ov = state.overrides[o.id] || {};
        const planned = o.collection ? '<span class="chip">Collection</span>' : s ? `${esc(s.customer)} ${when(s.delivery, ref)} <span class="muted small">${esc(s.kind)}</span>${o.moveType ? ` <span class="chip ${MOVE[o.moveType][1]}">${MOVE[o.moveType][0]}</span>` : ''}` : '<span class="chip low">No trailer</span>';
        return `<tr>
          <td class="mono">${when(o.delivery, ref)}</td>
          <td>${esc(o.customer || o.company)}${o.decide ? ' <span class="chip under">Decide</span>' : ''}</td>
          <td><span class="ref">${esc(o.seller || o.id)}</span><div class="small muted mono">Order ${esc(o.id.replace(/^0+/, ''))} · Wave ${esc(o.wave)}</div>${o.comments ? `<div class="small">${esc(o.comments)}</div>` : ''}</td>
          <td class="small">${esc(o.status)}<div class="muted">${esc(o.release)}</div></td>
          <td><span class="mono">${o.remaining} / ${o.total}</span>${areaChips(o)}</td>
          <td class="num mono">${pal(o.pallets)}</td>
          <td class="num mono" title="${esc(o.basis)}">${pct(o.frac)}</td>
          <td><input type="number" min="0" max="400" step="5" data-pct="${esc(o.id)}" value="${ov.pct != null ? esc(ov.pct) : ''}" placeholder="auto" aria-label="Trailer share for ${esc(o.seller)}"></td>
          <td><input type="checkbox" data-lock="${esc(o.id)}" ${ov.lock ? 'checked' : ''} aria-label="Keep ${esc(o.seller)} on its own trailer"></td>
          <td>${planned}</td></tr>`;
      }).join('')}</tbody></table></div>`;
  }

  function viewRules(r) {
    const c = state.cfg;
    const numField = (id, label, val, step, help) => `<label class="field"><span class="label">${label}</span><input type="number" step="${step}" data-cfg="${id}" value="${val}">${help ? `<span class="small muted">${help}</span>` : ''}</label>`;
    const chk = (id, label) => `<label class="check"><input type="checkbox" data-cfg-bool="${id}" ${c[id] ? 'checked' : ''}> ${label}</label>`;
    const txt = (id, label, help) => `<label class="field"><span class="label">${label}</span><input type="text" data-cfg-text="${id}" value="${esc(c[id] || '')}"><span class="small muted">${help}</span></label>`;
    const days = state.transport.days;
    return `<div class="rules">
      <section class="card"><h3>Trailer fill</h3>
        <p>Shopping tops up a trailer that already has orders until it reaches the target, without going over the maximum.</p>
        <div class="grid2">
          ${numField('targetFill%', 'Target fill %', Math.round(c.targetFill * 100), 5)}
          ${numField('maxFill%', 'Maximum fill %', Math.round(c.maxFill * 100), 5)}
          ${numField('lowFill%', 'Flag red below %', Math.round(c.lowFill * 100), 5)}
          ${numField('horizonHours', 'Look ahead (hours)', c.horizonHours, 1)}
          ${numField('leadMinutes', 'Stop shopping (min before paperwork)', c.leadMinutes, 15)}
        </div>
        ${chk('sameDayOnly', 'Only pull orders delivering the same day')}
        ${chk('protectSources', "Don't take orders off trailers that are already full enough")}
        ${chk('includeKitting', 'Include kitting orders (Cabs kits, Landpower kits, T4)')}
        ${chk('useXdock', 'Plan our orders onto X-Dock trailers as if they were ours')}
      </section>
      <section class="card"><h3>Pallets by picking area</h3>
        <p>1 pick is 1 pallet, except box areas. A mixed order adds up its share of a trailer from each area. Untick "Count" for columns that repeat picks counted elsewhere (P&D is already in VNA).</p>
        <div class="table-box"><table><thead><tr><th>Area</th><th>Column pattern</th><th class="num">Picks / pallet</th><th class="num">Pallets / trailer</th><th>Count</th></tr></thead><tbody>
        ${c.areas.map((a, i) => `<tr><td>${esc(a.name)}</td><td><input type="text" data-area="${i}" data-k="match" value="${esc(a.match)}"></td><td class="num"><input type="number" min="1" data-area="${i}" data-k="perPallet" value="${a.perPallet}"></td><td class="num"><input type="number" min="1" data-area="${i}" data-k="perTrailer" value="${a.perTrailer}"></td><td><input type="checkbox" data-area="${i}" data-k="count" ${a.count === false ? '' : 'checked'} aria-label="Count ${esc(a.name)}"></td></tr>`).join('')}
        </tbody></table></div>
        ${numField('defaultPerTrailer', 'Pallets per trailer, any other area', c.defaultPerTrailer, 1)}
      </section>
      <section class="card"><h3>Order rules</h3>
        <p>For orders that don't follow the area numbers. The first rule that matches wins. Use "Pallets / trailer" for big items (16 shovels fill a trailer), or "Orders / pallet" for kits (16 small boxes on a pallet). "Decide" asks the admin to set the share by hand.</p>
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
        ${txt('ourSite', 'Our trailers come from', 'Pattern for "Delivery / Collection From"')}
        ${txt('skipKinds', 'Ignore these delivery types', 'Not customer trailers, e.g. transfers')}
        ${txt('noShopKinds', 'Don\u2019t shop into', 'These keep their own orders but take nothing extra')}
        ${txt('xdockKinds', 'X-Dock trailers', 'Delivery types run by the X-Dock department')}
        ${txt('collectionMatch', 'Customer collections', 'Matched against seller order ID and comments')}
        <label class="field"><span class="label">Transport day to use</span>
          <select data-cfg-text="templateDay"><option value="">Automatic (same date, else same weekday)</option>
          ${days.map((d) => `<option value="${SP.dayToISO(d.day)}" ${c.templateDay === SP.dayToISO(d.day) ? 'selected' : ''}>${SP.fmtDay(d.day * DAY)} ${esc(d.title)}</option>`).join('')}</select></label>
      </section>
      <section class="card"><h3>Reset</h3>
        <p>Rules and hand-set trailer shares are saved in this browser only.</p>
        <div><button class="btn danger" data-act="reset" id="reset-btn">Reset rules to default</button></div>
      </section>
    </div>`;
  }

  /* ----------------------------- render ----------------------------- */
  function render() {
    compute();
    const r = state.result;
    const views = { shopping: viewShopping, trailers: viewTrailers, exceptions: viewExceptions, orders: viewOrders, rules: viewRules };
    app.innerHTML = viewTop(r) + viewControls(r) + viewKpis(r) + viewTabs(r) + `<main class="panel" id="panel">${views[state.tab](r)}</main>`;
  }

  /* ----------------------------- events ----------------------------- */
  function copyShopping() {
    const r = state.result;
    const head = ['Need by', 'Customer', 'Trailer delivery', 'Delivery type', 'Seller order', 'Orders', 'Wave', 'Was due', 'Picks left', 'Total picks', 'Pallets', 'Trailer %', 'Action'];
    const lines = SP.groupShopping(r.shopping.filter((m) => custOk(m.order.customer))).map((g) => {
      const o = g.order, s = g.slot, many = g.orders.length > 1;
      return [SP.fmtDayTime(s.paperwork != null ? s.paperwork : s.delivery), s.customer, SP.fmtDayTime(s.delivery), s.kind, many ? `${g.orders.length} orders` : o.seller, many ? g.orders.map((x) => x.id).join(' ') : o.id, o.wave, SP.fmtDayTime(o.delivery), g.remaining, g.total, Math.round(g.pallets * 10) / 10, Math.round(g.frac * 100), action(o, s).replace(/<[^>]+>/g, '')].join('\t');
    });
    const text = [head.join('\t'), ...lines].join('\n');
    const fallback = () => {
      const panel = document.getElementById('panel');
      const ta = document.createElement('textarea');
      ta.className = 'copy-fallback'; ta.value = text; ta.readOnly = true;
      panel.prepend(ta); ta.focus(); ta.select();
      toast('Press Ctrl+C to copy');
    };
    try {
      navigator.clipboard.writeText(text).then(() => toast('Copied – paste into Excel or Teams'), fallback);
    } catch (e) { fallback(); }
  }

  let resetArmed = false;
  app.addEventListener('click', (e) => {
    const b = e.target.closest('[data-act],[data-tab],[data-del-rule],[data-del-cust]');
    if (!b) return;
    if (b.dataset.tab) { state.tab = b.dataset.tab; render(); return; }
    if (b.dataset.delRule != null) { state.cfg.rules.splice(+b.dataset.delRule, 1); saveCfg(); render(); return; }
    if (b.dataset.delCust != null) { state.cfg.customers.splice(+b.dataset.delCust, 1); saveCfg(); render(); return; }
    const act = b.dataset.act;
    if (act === 'load') fileInput.click();
    else if (act === 'sample') { loadSample(); state.message = null; render(); }
    else if (act === 'now-file') { state.nowFromFile = true; render(); }
    else if (act === 'copy') copyShopping();
    else if (act === 'add-rule') { state.cfg.rules.push({ name: 'New rule', field: 'seller', match: '', cap: 35 }); saveCfg(); render(); }
    else if (act === 'add-cust') { state.cfg.customers.push(['New customer', '']); saveCfg(); render(); }
    else if (act === 'reset') {
      if (!resetArmed) { resetArmed = true; b.textContent = 'Click again to reset'; setTimeout(() => { resetArmed = false; if (b.isConnected) b.textContent = 'Reset rules to default'; }, 4000); return; }
      resetArmed = false; state.cfg = clone(SP.DEFAULT_CONFIG); saveCfg(); toast('Rules reset'); render();
    }
  });

  app.addEventListener('change', (e) => {
    const el = e.target;
    const d = el.dataset;
    if (el.id === 'now') { const v = SP.localInputToAbs(el.value); if (v != null) { state.now = v; state.nowFromFile = false; } }
    else if (el.id === 'customer') state.customer = el.value;
    else if (el.id === 'show-empty') state.showEmpty = el.checked;
    else if (el.id === 'show-past') state.showPast = el.checked;
    else if (d.done) { for (const id of d.done.split(',')) { if (el.checked) state.done[id] = Date.now(); else delete state.done[id]; } saveDone(); }
    else if (d.pct) {
      const o = state.overrides[d.pct] || {};
      if (el.value === '') delete o.pct; else o.pct = Math.max(0, +el.value);
      state.overrides[d.pct] = o; saveOverrides();
    } else if (d.lock) {
      const o = state.overrides[d.lock] || {}; o.lock = el.checked; state.overrides[d.lock] = o; saveOverrides();
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
    const y = window.scrollY;
    render();
    window.scrollTo(0, y);
  });

  /* ------------------------------ boot ------------------------------ */
  loadSample();
  render();
})();
