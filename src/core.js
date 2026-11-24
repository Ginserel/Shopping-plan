/*
 * Shopping Plan – core engine.
 *
 * Reads the WMS "Picking Status" export and the daily transport plan (both as
 * arrays of rows, as SheetJS gives them), links orders to trailer slots and
 * works out which later orders should be picked early ("shopped") to fill
 * trailers that would otherwise leave half empty.
 *
 * No DOM code here: the same file runs in the browser and in Node (CLI/tests).
 * All times are "absolute minutes": days since 1970-01-01 * 1440 + minute of day.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SPCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DAY = 1440;
  const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  /* ------------------------------------------------------------------ */
  /* Default rules. Everything here is editable in the app's Rules tab. */
  /* ------------------------------------------------------------------ */
  const DEFAULT_CONFIG = {
    // Per picking area: picks that make one pallet, and pallets that fill a trailer.
    // First pattern that matches the export's column name wins. count:false = not counted.
    areas: [
      { name: 'Wide', match: '^wide', perPallet: 1, perTrailer: 35 },
      { name: 'VNA 1-5', match: '^vna', perPallet: 1, perTrailer: 35 },
      { name: 'Bulk (internal)', match: '^bulk ?int', perPallet: 1, perTrailer: 14 },
      { name: 'Bulk external', match: '^bulk ?ext', perPallet: 1, perTrailer: 8 },
      { name: 'Mothersons (boxes)', match: '^m(o)?th', perPallet: 18, perTrailer: 35 },
      { name: 'Mezzanine (boxes)', match: '^mezz', perPallet: 18, perTrailer: 35 },
      { name: 'P&D (already in VNA)', match: '^p ?& ?d', perPallet: 1, perTrailer: 35, count: false },
      { name: 'Kits', match: 'kit$|^t4$', perPallet: 64, perTrailer: 35 },
    ],
    defaultPerTrailer: 35, // pallets per trailer for anything not listed
    // Order-level rules, matched against seller order ID / company / comments / order type.
    // First match wins. cap: pallets per trailer for the whole order (1 pick = 1 pallet).
    // perPallet: whole orders per pallet (kits are one small box each).
    // decide: the admin sets the trailer share by hand.
    rules: [
      { name: 'Cabs kits', field: 'type', match: 'cabs kit', perPallet: 16 },
      { name: 'Landpower kits', field: 'type', match: 'landpower.?kit|lp.?kit', perPallet: 16 },
      { name: 'T4 kits', field: 'type', match: 't4', perPallet: 16 },
      { name: 'Shovels', field: 'seller', match: 'shovel', cap: 16 },
      { name: 'Skips', field: 'seller', match: 'skip', cap: 5 },
      { name: 'Landpower (decide)', field: 'company', match: 'landpower', decide: true },
    ],
    targetFill: 0.85, // try to get every trailer to at least this
    maxFill: 1.0, // never shop a trailer above this
    lowFill: 0.5, // below this a trailer is flagged red
    horizonHours: 16, // only pull orders due within this many hours after the trailer
    sameDayOnly: true, // only pull orders delivering on the same day as the trailer
    leadMinutes: 60, // stop shopping into a trailer this long before its paperwork time
    protectSources: true, // don't take orders off a trailer that is already full enough
    includeKitting: false, // kitting is planned separately
    ignoreOrders: 'decant', // order type / seller order ID matching this is left out (decant)
    pickLeadMinutes: 120, // "Picked by" = paperwork ("Loaded by") minus this
    xdockKinds: '^\\s*x-?dock(?!\\s*[/+])', // X-Dock department trailers ("X-Dock", "XDOCK", not "XDOCK / BLOWOUT"): not ours
    useXdock: false, // true = plan our orders onto X-Dock trailers like our own
    ourSite: 'emg|maersk', // "Delivery / Collection From" values that are our trailers
    skipKinds: 'transfer|collection|emg to cover|decant', // delivery types that are not our customer trailers
    noShopKinds: 'live tip', // delivery types that keep their own orders but take nothing extra
    collectionMatch: 'coll|customer to collect', // orders the customer collects
    templateDay: null, // force a transport day (YYYY-MM-DD) as template; null = automatic
    customers: [
      // canonical name, regex tested against order company and transport "Deliver To"
      ['Landpower', 'landpower|lp kit'],
      ['Power Systems', 'power'],
      ['Compact', 'compact'],
      ['Loadall', 'loadall'],
      ['Backhoe', 'backhoe|\\bbhl\\b'],
      ['Earthmovers', 'earthmover'],
      ['Heavy', 'heavy'],
      ['Cabs', 'cabs'],
      ['Rayne', 'rayne'],
      ['Service', 'service|world parts'],
      ['Attachments', 'attachment'],
      ['Transmissions', 'transmission'],
      ['Gates', 'gates'],
      ['HBU', '\\bhbu\\b'],
      ['Rocester', 'rocester'],
    ],
  };

  /* ---------------------------- helpers ----------------------------- */
  const clean = (v) => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim());
  const lc = (v) => clean(v).toLowerCase();
  const num = (v) => {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const n = parseFloat(clean(v).replace(/[^0-9.\-]/g, ''));
    return isFinite(n) ? n : 0;
  };
  const stripZeros = (v) => clean(v).replace(/^0+(?=\d)/, '');
  function rx(src) {
    try { return new RegExp(src, 'i'); } catch (e) { return /$^/; }
  }

  const dayFromYMD = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 86400000);
  const dayToISO = (day) => new Date(day * 86400000).toISOString().slice(0, 10);
  const isoToDay = (s) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || '');
    return m ? dayFromYMD(+m[1], +m[2], +m[3]) : null;
  };
  const weekday = (day) => WEEKDAYS[(((day + 4) % 7) + 7) % 7];

  function parseDay(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v > 20000 ? Math.floor(v) - 25569 : null;
    if (v instanceof Date) return dayFromYMD(v.getFullYear(), v.getMonth() + 1, v.getDate());
    const s = clean(v);
    let m = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})/.exec(s);
    if (m) {
      let y = +m[3];
      if (y < 100) y += 2000;
      return dayFromYMD(y, +m[2], +m[1]);
    }
    m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
    return m ? dayFromYMD(+m[1], +m[2], +m[3]) : null;
  }

  function parseTime(v) {
    if (v == null || v === '' || typeof v === 'boolean') return null;
    if (typeof v === 'number') {
      if (v < 0) return null;
      return Math.round((v - Math.floor(v)) * DAY) % DAY;
    }
    if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
    const m = /(\d{1,2}):(\d{2})/.exec(clean(v));
    if (!m || +m[1] > 24 || +m[2] > 59) return null;
    return ((+m[1] * 60 + +m[2]) % DAY);
  }

  function parseDateTime(v) {
    const day = parseDay(v);
    if (day == null) return null;
    let t = 0;
    if (typeof v === 'number') t = parseTime(v) || 0;
    else if (!(v instanceof Date)) t = parseTime(clean(v).replace(/^\S+/, '')) || 0;
    else t = parseTime(v) || 0;
    return day * DAY + t;
  }

  // A delivery at 00:00 counts as 23:59 on the day before (the site's convention).
  const midnight = (abs) => (abs != null && abs % DAY === 0 ? abs - 1 : abs);

  // "13 h, 28 min left" / "1 days, 1 h, 28 min left" / "2 h, 5 min late"
  function parseLate(v) {
    const s = lc(v);
    if (!s) return null;
    const g = (re) => +((re.exec(s) || [])[1] || 0);
    const mins = g(/(\d+)\s*day/) * DAY + g(/(\d+)\s*h\b/) * 60 + g(/(\d+)\s*min/);
    if (/late|overdue|ago/.test(s)) return -mins;
    if (/left/.test(s)) return mins;
    return null;
  }

  function headerMap(row) {
    const map = {};
    (row || []).forEach((c, i) => {
      const k = lc(c);
      if (k && !(k in map)) map[k] = i;
    });
    return map;
  }
  function col(map, ...names) {
    for (const n of names) if (n in map) return map[n];
    for (const n of names) for (const h in map) if (h.startsWith(n)) return map[h];
    return -1;
  }
  const at = (row, i) => (i >= 0 ? row[i] : null);

  /* ------------------------- picking status ------------------------- */
  function parsePicking(rows) {
    const orders = [];
    const seen = new Set();
    const asOfSamples = [];
    let section = 'Orders';
    let h = null;
    let areaCols = [];

    rows.forEach((row) => {
      row = row || [];
      const cells = row.map(clean);
      const filled = cells.filter(Boolean);
      const lower = cells.map((c) => c.toLowerCase());

      if (lower.includes('delivery date') && lower.includes('order id')) {
        const m = headerMap(row);
        h = {
          del: col(m, 'delivery date'), id: col(m, 'order id'), wave: col(m, 'wave id'),
          seller: col(m, 'seller order id', 'seller order'), type: col(m, 'order type'),
          company: col(m, 'company'), status: col(m, 'status'), release: col(m, 'task release'),
          comments: col(m, 'comments'), late: col(m, 'late'), total: col(m, 'total picks'),
          rem: col(m, 'rmn. picks', 'rmn picks', 'remaining picks'), pct: col(m, 'picked%', 'picked %'),
        };
        const start = h.pct >= 0 ? h.pct + 1 : Math.max(h.total, h.rem) + 1;
        areaCols = [];
        for (let i = start; i < row.length; i++) if (clean(row[i])) areaCols.push({ i, name: clean(row[i]) });
        return;
      }
      if (filled.length === 1) {
        if (/^summary/i.test(filled[0])) { h = null; return; }
        if (/orders$/i.test(filled[0]) && !/^total/i.test(filled[0])) { section = filled[0]; return; }
      }
      if (!h) return;
      const id = clean(at(row, h.id));
      const delivery = midnight(parseDateTime(at(row, h.del)));
      if (delivery == null || !/^\d{3,}$/.test(id) || seen.has(id)) return;
      seen.add(id);

      const areas = {};
      let areaSum = 0;
      areaCols.forEach((a) => {
        const n = num(row[a.i]);
        if (n > 0) { areas[a.name] = n; areaSum += n; }
      });
      const total = num(at(row, h.total));
      const remaining = h.rem >= 0 ? num(at(row, h.rem)) : areaSum;
      const left = parseLate(at(row, h.late));
      if (left != null) asOfSamples.push(delivery - left);

      orders.push({
        id, delivery, section,
        wave: stripZeros(at(row, h.wave)),
        seller: clean(at(row, h.seller)),
        type: clean(at(row, h.type)),
        company: clean(at(row, h.company)),
        status: clean(at(row, h.status)),
        release: clean(at(row, h.release)),
        comments: clean(at(row, h.comments)),
        total: Math.max(total, remaining),
        remaining, areas, areaSum,
      });
    });

    asOfSamples.sort((a, b) => a - b);
    const asOf = asOfSamples.length ? asOfSamples[Math.floor(asOfSamples.length / 2)] : null;
    return { orders, asOf, areaNames: [...new Set(orders.flatMap((o) => Object.keys(o.areas)))] };
  }

  /* ------------------------- transport plan ------------------------- */
  function parseTransport(rows) {
    const slots = [];
    const days = new Map(); // day -> title
    let h = null;
    let blockDay = null;
    let section = 'main';
    let started = false;

    rows.forEach((row, r) => {
      row = row || [];
      const lower = row.map(lc);
      if (lower.includes('deliver to') && lower.some((c) => c.startsWith('delivery time'))) {
        const m = headerMap(row);
        h = {
          date: col(m, 'date'), title: col(m, 'hyperlink'), trl: col(m, 'trl / load check', 'trl'),
          pallets: col(m, 'pallet count', 'pallet'), trailer: col(m, 'trailer no', 'trailer'), bay: col(m, 'bay'),
          from: col(m, 'delivery / collection from', 'delivery /'), to: col(m, 'deliver to'),
          kind: col(m, 'delivery type'), loadType: col(m, 'load type'), split: col(m, 'split'), emer: col(m, 'emer'),
          paperwork: col(m, 'paperwork required by', 'paperwork'), despatch: col(m, 'despatch time'),
          delivery: col(m, 'delivery time'), team: col(m, 'loading team'), status: col(m, 'loading status'),
          completed: col(m, 'loading completed'), despatched: col(m, 'despatched'),
        };
        return;
      }
      if (!h) return;
      const day = parseDay(at(row, h.date));
      if (day == null) return;
      if (day !== blockDay) { blockDay = day; started = false; section = 'main'; }

      const title = clean(at(row, h.title));
      const trl = clean(at(row, h.trl));
      const to = clean(at(row, h.to));
      const from = clean(at(row, h.from));
      if (!to && !from && !trl && title) {
        if (/additional|rolled/i.test(title)) section = 'additional';
        else { section = 'main'; days.set(day, title); }
        return;
      }
      if (!days.has(day)) days.set(day, '');
      if (/line not used/i.test(trl) || (!to && !from)) return;

      const p = parseTime(at(row, h.paperwork));
      const s = parseTime(at(row, h.despatch));
      const d = parseTime(at(row, h.delivery));
      const anchor = p != null ? p : s != null ? s : d;
      let paperwork = null, despatch = null, delivery = null;
      if (anchor != null) {
        let aDay = day;
        if (section === 'main') {
          // The day's sheet opens with late-evening slots belonging to the night before.
          if (!started && anchor >= 18 * 60) aDay = day - 1;
          else if (anchor < 18 * 60) started = true;
        }
        const a = aDay * DAY + anchor;
        const fwd = (t) => (t == null ? null : a + ((t - anchor + DAY) % DAY));
        paperwork = fwd(p); despatch = fwd(s); delivery = midnight(fwd(d));
      }
      const statusText = clean(at(row, h.status));
      slots.push({
        id: 'R' + (r + 1), row: r + 1, day, section, title, trl,
        waves: (trl.match(/\d{5,7}/g) || []).map(stripZeros),
        pallets: clean(at(row, h.pallets)), trailer: clean(at(row, h.trailer)), bay: clean(at(row, h.bay)),
        from, to, kind: clean(at(row, h.kind)), loadType: clean(at(row, h.loadType)),
        emer: at(row, h.emer) === true || /^true|yes$/i.test(clean(at(row, h.emer))),
        paperwork, despatch, delivery,
        team: clean(at(row, h.team)), status: statusText,
        noOrder: /no order/i.test(trl) || /no order/i.test(statusText),
        gone: /despatch/i.test(statusText) || !!clean(at(row, h.despatched)),
      });
    });
    return { slots, days: [...days.entries()].map(([day, title]) => ({ day, title })).sort((a, b) => a.day - b.day) };
  }

  /* ----------------------------- sizing ----------------------------- */
  function areaRule(area, cfg) {
    for (const a of cfg.areas) if (rx(a.match).test(area)) return a;
    return null;
  }

  // Pallets and share of one trailer this order needs (frac 1 = a full trailer).
  function sizeOrder(o, cfg, override) {
    const out = { frac: 0, pallets: 0, basis: '', rule: null, decide: false, estimated: false };
    const per = cfg.defaultPerTrailer || 35;
    for (const r of cfg.rules) {
      const field = r.field === 'company' ? o.company : r.field === 'comments' ? o.comments : r.field === 'type' ? o.type : o.seller;
      if (!r.match || !rx(r.match).test(field)) continue;
      out.rule = r.name;
      if (r.decide) out.decide = true;
      if (+r.perPallet > 0) {
        out.pallets = 1 / +r.perPallet;
        out.frac = out.pallets / per;
        out.basis = `1 order of ${r.perPallet} per pallet (${r.name})`;
      } else if (+r.cap > 0) {
        out.pallets = o.total;
        out.frac = o.total / +r.cap;
        out.basis = `${o.total} pallets ÷ ${r.cap} per trailer (${r.name})`;
      }
      break;
    }
    if (!out.basis) {
      let frac = 0, pallets = 0, counted = 0;
      const parts = [];
      for (const [area, n] of Object.entries(o.areas)) {
        const a = areaRule(area, cfg);
        if (a && a.count === false) continue;
        const pp = a ? Math.max(1, +a.perPallet || 1) : 1;
        const pt = a ? Math.max(0.1, +a.perTrailer || per) : per;
        pallets += n / pp;
        frac += n / pp / pt;
        counted += n;
        parts.push(pp > 1 ? `${n} ${area} ÷ ${pp}/pallet` : `${n} ${area}`);
      }
      const picked = Math.max(0, o.total - o.areaSum);
      if (picked > 0) {
        // Already-picked picks lose their area in the export: assume the same mix.
        const fPer = counted > 0 ? frac / counted : 1 / per;
        const pPer = counted > 0 ? pallets / counted : 1;
        frac += picked * fPer;
        pallets += picked * pPer;
        parts.push(`${picked} picked`);
        if (counted === 0) out.estimated = true;
      }
      out.frac = frac;
      out.pallets = pallets;
      out.basis = parts.join(' + ');
    }
    if (override && override.pct != null && override.pct !== '') {
      out.frac = Math.max(0, +override.pct) / 100;
      out.pallets = out.frac * per;
      out.basis = `set by hand to ${override.pct}%`;
      out.manual = true;
      out.decide = false;
    }
    return out;
  }

  /* --------------------------- customers ---------------------------- */
  function customerOf(text, cfg) {
    const t = clean(text);
    if (!t) return null;
    for (const [name, pat] of cfg.customers) if (rx(pat).test(t)) return name;
    return null;
  }

  /* ---------------------------- planning ---------------------------- */
  function chooseTemplateDay(transport, day, cfg) {
    const forced = isoToDay(cfg.templateDay);
    if (forced != null && transport.days.some((d) => d.day === forced)) return forced;
    if (transport.days.some((d) => d.day === day)) return day;
    // Mon–Thu share one schedule, so the most recent Mon–Thu sheet is the best guess.
    // Fridays and weekends only use their own weekday.
    const wd = weekday(day);
    const normal = (d) => ['Mon', 'Tue', 'Wed', 'Thu'].includes(weekday(d.day));
    let pool = ['Mon', 'Tue', 'Wed', 'Thu'].includes(wd) ? transport.days.filter(normal) : transport.days.filter((d) => weekday(d.day) === wd);
    if (!pool.length) pool = transport.days;
    const before = pool.filter((d) => d.day < day);
    const list = before.length ? before : pool;
    return list.length ? list[list.length - 1].day : null;
  }

  function slotsForDays(transport, dayList, cfg) {
    const out = [];
    const used = [];
    for (const day of dayList) {
      const src = chooseTemplateDay(transport, day, cfg);
      if (src == null) continue;
      const exact = src === day;
      used.push({ day, source: src, exact });
      const shift = (day - src) * DAY;
      for (const s of transport.slots) {
        if (s.day !== src) continue;
        if (s.delivery == null) continue;
        if (!exact && s.section !== 'main') continue; // one-off extras don't repeat
        const mv = (t) => (t == null ? null : t + shift);
        out.push({
          ...s,
          id: s.id + '@' + dayToISO(day),
          day,
          paperwork: mv(s.paperwork), despatch: mv(s.despatch), delivery: mv(s.delivery),
          templated: !exact,
          sourceDay: src,
          waves: exact ? s.waves : [],
          gone: exact ? s.gone : false,
          status: exact ? s.status : '',
          pallets: exact ? s.pallets : '',
          trailer: exact ? s.trailer : '',
        });
      }
    }
    // A stable name for each trailer, so choices survive loading a newer export.
    const seen = new Map();
    out.sort((a, b) => a.delivery - b.delivery || a.row - b.row);
    for (const s of out) {
      const base = `${dayToISO(Math.floor(s.delivery / DAY))} ${fmtTime(s.delivery)} ${clean(s.to)} ${clean(s.kind)}`;
      const n = (seen.get(base) || 0) + 1;
      seen.set(base, n);
      s.key = n > 1 ? `${base} #${n}` : base;
    }
    return { slots: out, used };
  }

  function plan(picking, transport, cfg, opts) {
    cfg = Object.assign({}, DEFAULT_CONFIG, cfg || {});
    opts = opts || {};
    const overrides = opts.overrides || {};
    const now = opts.now != null ? opts.now : picking.asOf != null ? picking.asOf : 0;
    const siteRx = rx(cfg.ourSite);
    const skipRx = rx(cfg.skipKinds);
    const noShopRx = cfg.noShopKinds ? rx(cfg.noShopKinds) : /$^/;
    const collRx = rx(cfg.collectionMatch);
    const xdRx = cfg.xdockKinds ? rx(cfg.xdockKinds) : /$^/;

    /* orders */
    const ignoreRx = cfg.ignoreOrders ? rx(cfg.ignoreOrders) : null;
    const ignored = { kitting: 0, other: 0 };
    const orders = picking.orders
      .filter((o) => {
        if (!cfg.includeKitting && /kit/i.test(o.section)) { ignored.kitting++; return false; }
        if (ignoreRx && (ignoreRx.test(o.type) || ignoreRx.test(o.seller))) { ignored.other++; return false; }
        return true;
      })
      .map((o) => {
        const size = sizeOrder(o, cfg, overrides[o.id]);
        return {
          ...o, ...size,
          customer: customerOf(o.company, cfg),
          collection: collRx.test(o.seller) || collRx.test(o.comments),
          locked: !!(overrides[o.id] && overrides[o.id].lock),
          home: null, slot: null, moveType: null, note: '',
        };
      });

    /* trailer slots for every day that has orders (and today) */
    // Day + 1 too: its first rows are the 00:00 trailers, which count as 23:59 the day before.
    const daySet = new Set();
    for (const o of orders) { const d = Math.floor(o.delivery / DAY); daySet.add(d); daySet.add(d + 1); }
    daySet.add(Math.floor(now / DAY));
    daySet.add(Math.floor(now / DAY) + 1);
    const { slots: rawSlots, used } = slotsForDays(transport, [...daySet].sort((a, b) => a - b), cfg);
    const mapped = rawSlots.map((s) => ({
      ...s,
      customer: customerOf(s.to, cfg),
      ours: !s.from || siteRx.test(s.from),
      xdock: xdRx.test(s.kind),
      pickBy: s.paperwork != null ? s.paperwork - (cfg.pickLeadMinutes || 0) : null,
    }));
    const isTrailer = (s) => s.customer && s.delivery != null && !skipRx.test(s.kind) && !/external|transfer/i.test(s.loadType);
    const candidates = mapped.filter((s) => s.ours && isTrailer(s));
    // X-Dock trailers belong to another department: kept aside to suggest sharing.
    const xdock = candidates.filter((s) => s.xdock && !cfg.useXdock && !s.gone && s.delivery > now);
    const slots = candidates
      .filter((s) => cfg.useXdock || !s.xdock)
      .map((s) => ({
        ...s,
        orders: [],
        load: 0,
        past: s.gone || s.delivery <= now,
        shopOpen: !s.gone && s.paperwork != null ? s.paperwork - cfg.leadMinutes > now : !s.gone && s.delivery - 120 - cfg.leadMinutes > now,
        noShop: noShopRx.test(s.kind),
      }))
      .sort((a, b) => a.delivery - b.delivery || a.row - b.row);

    const slotById = new Map(slots.map((s) => [s.id, s]));
    const byCustomer = new Map();
    slots.forEach((s) => {
      if (!byCustomer.has(s.customer)) byCustomer.set(s.customer, []);
      byCustomer.get(s.customer).push(s);
    });

    const put = (o, s) => {
      if (o.slot) {
        const old = slotById.get(o.slot);
        old.orders = old.orders.filter((x) => x !== o);
        old.load -= o.frac;
      }
      o.slot = s ? s.id : null;
      if (s) { s.orders.push(o); s.load += o.frac; }
    };
    const leastLoaded = (list) => list.reduce((best, s) => (!best || s.load < best.load ? s : best), null);

    /* 0. the planner's own choices are fixed; suggestions work around them */
    const fixed = opts.fixed || {};
    const slotByKey = new Map(slots.map((s) => [s.key, s]));
    for (const o of orders) {
      const k = fixed[o.id];
      if (!k) continue;
      o.fixed = k;
      o.locked = true;
      const s = slotByKey.get(k);
      if (s) { o.home = s.id; put(o, s); }
    }

    /* 1. home trailer: wave written on the plan, else same customer + delivery time */
    const homeless = [];
    const sorted = [...orders].sort((a, b) => b.frac - a.frac);
    for (const o of sorted) {
      if (o.fixed || o.collection || !o.customer) continue;
      let s = leastLoaded(slots.filter((x) => o.wave && x.waves.includes(o.wave)));
      if (!s) s = leastLoaded((byCustomer.get(o.customer) || []).filter((x) => x.delivery === o.delivery));
      if (s) { o.home = s.id; put(o, s); } else homeless.push(o);
    }

    const moves = [];
    const move = (o, s, type, note) => {
      moves.push({ order: o, from: o.slot, to: s.id, type, note });
      o.moveType = type;
      o.note = note;
      put(o, s);
    };

    /* 2. orders with no trailer at their time: latest open trailer before, else next one */
    for (const o of homeless.sort((a, b) => a.delivery - b.delivery)) {
      const list = (byCustomer.get(o.customer) || []).filter((x) => !x.past && !x.noShop);
      const before = list.filter((x) => x.delivery <= o.delivery && x.shopOpen && o.delivery - x.delivery <= cfg.horizonHours * 60);
      const s = before.length ? before[before.length - 1] : list.find((x) => x.delivery > o.delivery);
      if (s) move(o, s, 'noslot', `No ${o.customer} trailer at ${fmtTime(o.delivery)}`);
    }

    /* 3. missed trailers: home trailer gone/past but picks still open -> roll to next */
    for (const o of orders) {
      const s = o.slot && slotById.get(o.slot);
      if (!s || !s.past || o.remaining <= 0 || o.fixed) continue;
      const next = (byCustomer.get(o.customer) || []).find((x) => !x.past && x.delivery > s.delivery && !x.noShop);
      if (next) move(o, next, 'roll', `Missed ${fmtTime(s.delivery)} trailer`);
      else o.note = `Missed ${fmtTime(s.delivery)} trailer – no later trailer in plan`;
    }

    /* 4. shopping: fill under-filled trailers with later orders, earliest due first.
          Only trailers that are already running (have orders) get shopped into:
          an empty slot is a trailer transport doesn't need to send. */
    for (const [, list] of byCustomer) {
      for (const s of list) {
        // s.orders is live: a trailer that already gave its orders away stays empty.
        if (!s.orders.length || !s.shopOpen || s.noShop || s.load >= cfg.targetFill) continue;
        const pool = orders
          .filter((o) => {
            if (o.customer !== s.customer || o.locked || o.collection || !o.slot) return false;
            if (o.frac <= 0 || o.frac > cfg.maxFill) return false;
            const cur = slotById.get(o.slot);
            if (cur === s || cur.past) return false;
            if (cur.delivery <= s.delivery) return false;
            if (o.delivery - s.delivery > cfg.horizonHours * 60) return false;
            if (cfg.sameDayOnly && Math.floor(o.delivery / DAY) !== Math.floor(s.delivery / DAY)) return false;
            return true;
          })
          .sort((a, b) => a.delivery - b.delivery || b.frac - a.frac);
        for (const o of pool) {
          if (s.load >= cfg.targetFill) break;
          if (s.load + o.frac > cfg.maxFill + 1e-9) continue;
          const cur = slotById.get(o.slot);
          if (cfg.protectSources && cur.load >= cfg.targetFill && cur.load - o.frac < cfg.targetFill) continue;
          move(o, s, 'shop', `Pulled forward from ${fmtTime(cur.delivery)}`);
        }
      }
    }

    /* 5. slot verdicts */
    for (const s of slots) {
      const own = s.orders.filter((o) => o.home === s.id && !o.moveType);
      s.ownLoad = own.reduce((t, o) => t + o.frac, 0);
      s.shopLoad = s.load - s.ownLoad;
      s.picksLeft = s.orders.reduce((t, o) => t + o.remaining, 0);
      s.picks = s.orders.reduce((t, o) => t + o.total, 0);
      s.pallets = s.orders.reduce((t, o) => t + o.pallets, 0);
      s.trailersNeeded = Math.max(1, Math.ceil(s.load / cfg.maxFill - 1e-9));
      if (s.past) s.verdict = 'past';
      else if (!s.orders.length) s.verdict = 'empty';
      else if (s.load > cfg.maxFill + 1e-9) s.verdict = 'over';
      else if (s.load < cfg.lowFill) s.verdict = 'low';
      else if (s.load < cfg.targetFill) s.verdict = 'under';
      else s.verdict = 'ok';
      s.lostAll = s.verdict === 'empty' && moves.some((m) => m.from === s.id);
    }

    /* 6. exceptions */
    const ex = [];
    const xdockNear = (customer, abs, hours) => xdock
      .filter((x) => x.customer === customer && Math.abs(x.delivery - abs) <= hours * 60)
      .sort((a, b) => Math.abs(a.delivery - abs) - Math.abs(b.delivery - abs));
    const xdText = (list) => list.length ? ` X-Dock trailer at ${list.slice(0, 2).map((x) => fmtTime(x.delivery)).join(' / ')} – ask X-Dock if it can take it.` : '';
    for (const o of orders) {
      if (o.fixed) { if (o.decide) ex.push({ kind: 'decide', order: o, text: `${o.rule}: set how much of a trailer this takes` }); continue; }
      if (o.collection) ex.push({ kind: 'collection', order: o, text: 'Customer collection – not on our trailers' });
      else if (!o.customer) ex.push({ kind: 'customer', order: o, text: `Unknown customer "${o.company}" – add it under Rules › Customers` });
      else if (!o.slot) ex.push({ kind: 'noslot', order: o, text: `No ${o.customer} trailer in the transport plan for ${fmtDay(o.delivery)}.${xdText(xdockNear(o.customer, o.delivery, 3))}` });
      else if (o.moveType === 'noslot') ex.push({ kind: 'noslot', order: o, text: `${o.note} – planned on ${fmtTime(slotById.get(o.slot).delivery)}` });
      if (o.note && /Missed/.test(o.note)) ex.push({ kind: 'missed', order: o, text: o.note });
      if (o.decide) ex.push({ kind: 'decide', order: o, text: `${o.rule}: set how much of a trailer this takes` });
      if (o.frac > cfg.maxFill + 1e-9) ex.push({ kind: 'big', order: o, text: `Order alone needs ${Math.ceil(o.frac / cfg.maxFill)} trailers` });
      if (o.estimated) ex.push({ kind: 'estimate', order: o, text: 'Fully picked – area unknown, sized as 1 pick = 1 pallet' });
    }
    for (const s of slots) {
      if (s.verdict === 'over') {
        const spare = (byCustomer.get(s.customer) || []).filter((x) => x !== s && !x.past && x.shopOpen && !x.orders.length && Math.floor(x.delivery / DAY) === Math.floor(s.delivery / DAY) && x.delivery < s.delivery);
        const hint = spare.length ? ` Unused earlier slot: ${spare.map((x) => fmtTime(x.delivery) + ' ' + x.kind).join(', ')}.` : '';
        ex.push({ kind: 'over', slot: s, text: `Over capacity: ${Math.round(s.load * 100)}% – needs ${s.trailersNeeded} trailers.${hint}` });
      }
      if ((s.verdict === 'under' || s.verdict === 'low') && !s.past) {
        const near = xdockNear(s.customer, s.delivery, 2);
        if (near.length) ex.push({ kind: 'xdock', slot: s, text: `${Math.round(s.load * 100)}% full.${xdText(near)}` });
      }
      if (s.verdict === 'empty' && s.lostAll) ex.push({ kind: 'free', slot: s, text: 'All orders moved to an earlier trailer – check with transport if this trailer is still needed' });
    }
    const usedAreaCaps = new Set();
    for (const o of orders) for (const a of Object.keys(o.areas)) {
      const rule = areaRule(a, cfg);
      if (rule && rule.confirm) usedAreaCaps.add(rule.name);
    }

    // Many orders with the same problem (e.g. 40 cabs kits on one wave) become one line.
    const merged = [];
    const byKey = new Map();
    for (const e of ex) {
      if (!e.order) { merged.push(e); continue; }
      const key = [e.kind, e.order.customer, e.order.delivery, e.text].join('|');
      if (byKey.has(key)) { byKey.get(key).orders.push(e.order); continue; }
      const item = { ...e, orders: [e.order] };
      byKey.set(key, item);
      merged.push(item);
    }
    ex.length = 0;
    ex.push(...merged);

    const upcoming = slots.filter((s) => !s.past);
    const shopping = moves
      .filter((m) => m.order.slot === m.to) // last move wins
      .filter((m, i, arr) => arr.findIndex((x) => x.order === m.order) === i)
      .map((m) => ({ ...m, slot: slotById.get(m.to), fromSlot: m.from ? slotById.get(m.from) : null }))
      .sort((a, b) => (a.slot.paperwork || a.slot.delivery) - (b.slot.paperwork || b.slot.delivery));

    /* every customer trailer in the plan, for the planner's board */
    const board = mapped.filter(isTrailer).map((s) => {
      const planned = slotById.get(s.id);
      if (planned) return Object.assign(planned, { role: 'ours' });
      return { ...s, role: s.xdock ? 'xdock' : 'other', orders: [], load: 0, past: s.gone || s.delivery <= now, verdict: 'empty' };
    }).sort((a, b) => (a.paperwork != null ? a.paperwork : a.delivery) - (b.paperwork != null ? b.paperwork : b.delivery) || a.delivery - b.delivery || a.row - b.row);

    return {
      now, cfg, orders, slots, board, ignored, xdock, used, moves, shopping, exceptions: ex,
      unconfirmedAreas: [...usedAreaCaps],
      summary: {
        orders: orders.length,
        trailers: upcoming.length,
        underBefore: upcoming.filter((s) => s.ownLoad + 1e-9 < cfg.targetFill && s.orders.some((o) => o.home === s.id)).length,
        under: upcoming.filter((s) => s.verdict === 'under' || s.verdict === 'low').length,
        moves: shopping.length,
        freed: upcoming.filter((s) => s.lostAll).length,
        empty: upcoming.filter((s) => s.verdict === 'empty').length,
        noSlot: ex.filter((e) => e.kind === 'noslot').reduce((t, e) => t + e.orders.length, 0),
        decide: ex.filter((e) => e.kind === 'decide').length,
        over: upcoming.filter((s) => s.verdict === 'over').length,
      },
    };
  }

  /* ------------------------ the planner's list ----------------------- */
  // Where an order is in picking: picked / picking / released / release (still to release).
  function orderState(o) {
    if (o.remaining <= 0 && o.total > 0) return 'picked';
    if (o.remaining < o.total || /progress/i.test(o.status)) return 'picking';
    if (/released/i.test(o.release) && !/not/i.test(o.release)) return 'released';
    return 'release';
  }
  const STATE_TEXT = { picked: 'Picked', picking: 'Pick in progress', released: 'Released', release: 'To release' };

  function dur(mins) {
    mins = Math.round(Math.abs(mins));
    const h = Math.floor(mins / 60), m = mins % 60;
    return h && m ? `${h}h ${m}m` : h ? `${h}h` : `${m}m`;
  }

  function reasonFor(o, s) {
    if (o.moveType === 'shop') return { kind: 'shop', text: `Pull forward from ${fmtTime(o.delivery)} to fill this trailer` };
    if (o.moveType === 'noslot') return { kind: 'noslot', text: o.note };
    if (o.moveType === 'roll') return { kind: 'roll', text: o.note };
    if (o.wave && s.waves && s.waves.includes(o.wave)) return { kind: 'home', text: `Wave ${o.wave} is written on this trailer` };
    return { kind: 'home', text: `Due ${fmtTime(o.delivery)}` };
  }

  // assign: { orderId: trailer key } – the planner's choices. Returns one row per trailer
  // (chosen orders, suggestions, fill, status, warnings) and the orders not on the list.
  function buildList(r, assign) {
    assign = assign || {};
    const cfg = r.cfg;
    const rows = new Map();
    for (const s of r.board) rows.set(s.key, { key: s.key, slot: s, chosen: [], suggested: [], load: 0, sugLoad: 0, warnings: [] });
    const byId = new Map(r.board.map((s) => [s.id, s]));
    const pool = [];
    for (const o of r.orders) {
      const k = assign[o.id];
      const row = k && rows.get(k);
      if (row) { row.chosen.push(o); row.load += o.frac; continue; }
      const s = o.slot && byId.get(o.slot);
      const sug = s && !s.past ? { slot: s, key: s.key, reason: reasonFor(o, s) } : null;
      pool.push({ order: o, suggestion: sug, missing: k || null });
      if (sug) {
        const rr = rows.get(s.key);
        rr.suggested.push({ order: o, reason: sug.reason });
        rr.sugLoad += o.frac;
      }
    }
    for (const row of rows.values()) {
      const s = row.slot;
      row.chosen.sort((a, b) => a.delivery - b.delivery || a.seller.localeCompare(b.seller));
      row.suggested.sort((a, b) => a.order.delivery - b.order.delivery);
      row.notes = new Map();
      for (const o of row.chosen) {
        const n = [];
        if (o.delivery < s.delivery - 1) n.push({ kind: 'late', text: `Due ${fmtTime(o.delivery)} – arrives ${dur(s.delivery - o.delivery)} late` });
        else if (o.delivery > s.delivery + 1) n.push({ kind: 'early', text: `Due ${Math.floor(o.delivery / DAY) === Math.floor(s.delivery / DAY) ? fmtTime(o.delivery) : fmtDayTime(o.delivery)} – going early` });
        if (o.collection) n.push({ kind: 'warn', text: 'Customer collection' });
        else if (o.customer !== s.customer) n.push({ kind: 'info', text: `For ${o.customer || o.company}` });
        if (o.decide) n.push({ kind: 'warn', text: 'Size is a guess – set it on the Orders tab' });
        row.notes.set(o.id, n);
      }
      const states = row.chosen.map(orderState);
      row.picks = row.chosen.reduce((t, o) => t + o.total, 0);
      row.picksLeft = row.chosen.reduce((t, o) => t + o.remaining, 0);
      row.pallets = row.chosen.reduce((t, o) => t + o.pallets, 0);
      row.toRelease = row.chosen.filter((o) => orderState(o) === 'release');
      row.waves = [...new Set(row.toRelease.map((o) => o.wave).filter(Boolean))];
      row.status = !states.length ? ''
        : states.every((x) => x === 'picked') ? STATE_TEXT.picked
        : states.some((x) => x === 'picked' || x === 'picking') ? STATE_TEXT.picking
        : states.every((x) => x === 'released') ? STATE_TEXT.released
        : `${STATE_TEXT.release}${row.toRelease.length < states.length ? ` (${row.toRelease.length} of ${states.length})` : ''}`;
      if (row.load > cfg.maxFill + 1e-9) row.warnings.push({ kind: 'over', text: `Over capacity – about ${Math.ceil(row.load / cfg.maxFill - 1e-9)} trailers' worth` });
      if (row.chosen.some((o) => row.notes.get(o.id).some((n) => n.kind === 'late'))) row.warnings.push({ kind: 'late', text: 'Has orders that will arrive late' });
      if (s.pickBy != null && s.pickBy <= r.now && !s.past && states.some((x) => x !== 'picked')) row.warnings.push({ kind: 'late', text: 'Pick-by time has passed' });
      if (s.role !== 'ours' && row.chosen.length) row.warnings.push({ kind: 'info', text: s.role === 'xdock' ? 'X-Dock trailer – check with X-Dock' : `Trailer from ${s.from}` });
    }
    pool.sort((a, b) => a.order.delivery - b.order.delivery || (a.order.customer || '~').localeCompare(b.order.customer || '~'));
    return { rows: [...rows.values()], pool };
  }

  // The admin's pick plan, in the same columns as the hand-made sheet plus what the admin needs
  // to release. rows: buildList(...).rows already filtered and in order. comments: { key: text }.
  const PICK_PLAN_HEAD = ['Deliver To', 'Delivery Type', 'Picked By', 'Loaded By', 'Delivery Time', 'Orders to pick', 'Status', 'Comments', 'Waves to release', 'Picks left', 'Trailer fill'];
  function pickPlanTable(rows, comments) {
    comments = comments || {};
    return rows.map((row) => {
      const s = row.slot;
      const orders = row.chosen.map((o) => o.seller || o.id.replace(/^0+/, ''));
      const tag = s.role === 'xdock' && !/x-?dock/i.test(s.kind || '') ? ' (X-Dock)' : s.role === 'other' && s.from ? ` (from ${s.from})` : '';
      return {
        row,
        cells: [
          s.to, (s.kind || '') + tag,
          s.pickBy != null && s.role === 'ours' || row.chosen.length ? fmtTime(s.pickBy) : '',
          fmtTime(s.paperwork), fmtTime(s.delivery),
          orders.join('\n'), row.status, comments[row.key] || '',
          row.waves.join('\n'), row.chosen.length ? row.picksLeft : '',
          row.chosen.length ? `${Math.round(row.load * 100)}%` : '',
        ],
      };
    });
  }

  // Shopping lines for the same wave going to the same trailer are picked together.
  function groupShopping(shopping) {
    const groups = [];
    const byKey = new Map();
    for (const m of shopping) {
      const key = [m.order.wave || m.order.id, m.to, m.type, m.order.delivery].join('|');
      if (byKey.has(key)) {
        const g = byKey.get(key);
        g.orders.push(m.order);
        g.frac += m.order.frac; g.pallets += m.order.pallets; g.remaining += m.order.remaining; g.total += m.order.total;
        continue;
      }
      const g = { ...m, orders: [m.order], frac: m.order.frac, pallets: m.order.pallets, remaining: m.order.remaining, total: m.order.total };
      byKey.set(key, g);
      groups.push(g);
    }
    return groups;
  }

  /* ---------------------------- formatting -------------------------- */
  function pad(n) { return String(n).padStart(2, '0'); }
  function fmtTime(abs) {
    if (abs == null) return '–';
    const m = ((abs % DAY) + DAY) % DAY;
    return pad(Math.floor(m / 60)) + ':' + pad(m % 60);
  }
  function fmtDay(abs) {
    const day = Math.floor(abs / DAY);
    const d = new Date(day * 86400000);
    return `${weekday(day)} ${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}`;
  }
  function fmtDayTime(abs) { return abs == null ? '–' : `${fmtDay(abs)} ${fmtTime(abs)}`; }
  function absToLocalInput(abs) { return `${dayToISO(Math.floor(abs / DAY))}T${fmtTime(abs)}`; }
  function localInputToAbs(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(s || '');
    return m ? dayFromYMD(+m[1], +m[2], +m[3]) * DAY + +m[4] * 60 + +m[5] : null;
  }

  // Decide which parsed workbook is which.
  function detect(rows) {
    const p = parsePicking(rows);
    if (p.orders.length) return { type: 'picking', data: p };
    const t = parseTransport(rows);
    if (t.slots.length) return { type: 'transport', data: t };
    return { type: null };
  }

  return {
    DAY, DEFAULT_CONFIG, parsePicking, parseTransport, plan, sizeOrder, customerOf, detect, groupShopping,
    buildList, orderState, STATE_TEXT, dur, pickPlanTable, PICK_PLAN_HEAD,
    fmtTime, fmtDay, fmtDayTime, absToLocalInput, localInputToAbs, dayToISO, isoToDay, weekday,
    _internal: { parseDay, parseTime, parseDateTime, parseLate, chooseTemplateDay },
  };
});
