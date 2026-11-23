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
    // How many picks from each area fill one trailer. First match wins.
    areas: [
      { name: 'Wide', match: '^wide', cap: 35 },
      { name: 'VNA 1-5', match: '^vna', cap: 35 },
      { name: 'Bulk (internal)', match: '^bulk ?int', cap: 14 },
      { name: 'Bulk external', match: '^bulk ?ext', cap: 8 },
      { name: 'Mothersons', match: '^m(o)?th', cap: 200, confirm: true },
      { name: 'Mezzanine', match: '^mezz', cap: 100, confirm: true },
      { name: 'P&D', match: '^p ?& ?d', cap: 35, confirm: true },
      { name: 'Kits (Cabs / LP / T4)', match: 'kit$|^t4$', cap: 35, confirm: true },
    ],
    defaultCap: 35,
    // Order-level rules, matched against seller order ID / company / comments.
    // cap: picks per trailer for the whole order. decide: admin sets the fill by hand.
    rules: [
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
    includeKitting: false,
    ourSite: 'emg|maersk', // "Delivery / Collection From" values that are our trailers
    skipKinds: 'transfer|collection|emg to cover', // delivery types that are not our customer trailers
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
      const delivery = parseDateTime(at(row, h.del));
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
        paperwork = fwd(p); despatch = fwd(s); delivery = fwd(d);
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
  function areaCap(area, cfg) {
    for (const a of cfg.areas) if (rx(a.match).test(area)) return Math.max(0.1, +a.cap || cfg.defaultCap);
    return cfg.defaultCap;
  }

  // Share of one trailer this order needs (1 = a full trailer).
  function sizeOrder(o, cfg, override) {
    const out = { frac: 0, basis: '', rule: null, decide: false, estimated: false };
    for (const r of cfg.rules) {
      const field = r.field === 'company' ? o.company : r.field === 'comments' ? o.comments : o.seller;
      if (!rx(r.match).test(field)) continue;
      out.rule = r.name;
      if (r.decide) out.decide = true;
      if (+r.cap > 0) {
        out.frac = o.total / +r.cap;
        out.basis = `${o.total} picks ÷ ${r.cap} per trailer (${r.name})`;
      }
      break;
    }
    if (!out.basis) {
      let frac = 0;
      const parts = [];
      for (const [area, n] of Object.entries(o.areas)) {
        const cap = areaCap(area, cfg);
        frac += n / cap;
        parts.push(`${n} ${area}/${cap}`);
      }
      const picked = Math.max(0, o.total - o.areaSum);
      if (picked > 0) {
        // Already-picked pallets lose their area in the export: assume the same mix.
        const perPick = o.areaSum > 0 ? frac / o.areaSum : 1 / cfg.defaultCap;
        frac += picked * perPick;
        parts.push(`${picked} picked`);
        if (o.areaSum === 0) out.estimated = true;
      }
      out.frac = frac;
      out.basis = parts.join(' + ');
    }
    if (override && override.pct != null && override.pct !== '') {
      out.frac = Math.max(0, +override.pct) / 100;
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
    const sameWd = transport.days.filter((d) => weekday(d.day) === weekday(day));
    const pool = sameWd.length ? sameWd : transport.days;
    return pool.length ? pool[pool.length - 1].day : null;
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

    /* orders */
    const orders = picking.orders
      .filter((o) => cfg.includeKitting || !/kit/i.test(o.section))
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
    const daySet = new Set(orders.map((o) => Math.floor(o.delivery / DAY)));
    daySet.add(Math.floor(now / DAY));
    const { slots: rawSlots, used } = slotsForDays(transport, [...daySet].sort((a, b) => a - b), cfg);
    const slots = rawSlots
      .map((s) => ({
        ...s,
        customer: customerOf(s.to, cfg),
        ours: !s.from || siteRx.test(s.from),
      }))
      .filter((s) => s.ours && s.customer && s.delivery != null && !skipRx.test(s.kind) && !/external|transfer/i.test(s.loadType))
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

    /* 1. home trailer: wave written on the plan, else same customer + delivery time */
    const homeless = [];
    const sorted = [...orders].sort((a, b) => b.frac - a.frac);
    for (const o of sorted) {
      if (o.collection || !o.customer) continue;
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
      if (!s || !s.past || o.remaining <= 0) continue;
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
    for (const o of orders) {
      if (o.collection) ex.push({ kind: 'collection', order: o, text: 'Customer collection – not on our trailers' });
      else if (!o.customer) ex.push({ kind: 'customer', order: o, text: `Unknown customer "${o.company}" – add it under Rules › Customers` });
      else if (!o.slot) ex.push({ kind: 'noslot', order: o, text: `No ${o.customer} trailer in the transport plan for ${fmtDay(o.delivery)}` });
      else if (o.moveType === 'noslot') ex.push({ kind: 'noslot', order: o, text: `${o.note} – planned on ${fmtTime(slotById.get(o.slot).delivery)}` });
      if (o.note && /Missed/.test(o.note)) ex.push({ kind: 'missed', order: o, text: o.note });
      if (o.decide) ex.push({ kind: 'decide', order: o, text: `${o.rule}: set how much of a trailer this takes` });
      if (o.frac > cfg.maxFill + 1e-9) ex.push({ kind: 'big', order: o, text: `Order alone needs ${Math.ceil(o.frac / cfg.maxFill)} trailers` });
      if (o.estimated) ex.push({ kind: 'estimate', order: o, text: 'Fully picked – area unknown, sized as Wide' });
    }
    for (const s of slots) {
      if (s.verdict === 'over') {
        const spare = (byCustomer.get(s.customer) || []).filter((x) => x !== s && !x.past && x.shopOpen && !x.orders.length && Math.floor(x.delivery / DAY) === Math.floor(s.delivery / DAY) && x.delivery < s.delivery);
        const hint = spare.length ? ` Unused earlier slot: ${spare.map((x) => fmtTime(x.delivery) + ' ' + x.kind).join(', ')}.` : '';
        ex.push({ kind: 'over', slot: s, text: `Over capacity: ${Math.round(s.load * 100)}% – needs ${s.trailersNeeded} trailers.${hint}` });
      }
      if (s.verdict === 'empty' && s.lostAll) ex.push({ kind: 'free', slot: s, text: 'All orders moved to an earlier trailer – check with transport if this trailer is still needed' });
    }
    const usedAreaCaps = new Set();
    for (const o of orders) for (const a of Object.keys(o.areas)) {
      const rule = cfg.areas.find((r) => rx(r.match).test(a));
      if (rule && rule.confirm) usedAreaCaps.add(rule.name);
    }

    const upcoming = slots.filter((s) => !s.past);
    const shopping = moves
      .filter((m) => m.order.slot === m.to) // last move wins
      .filter((m, i, arr) => arr.findIndex((x) => x.order === m.order) === i)
      .map((m) => ({ ...m, slot: slotById.get(m.to), fromSlot: m.from ? slotById.get(m.from) : null }))
      .sort((a, b) => (a.slot.paperwork || a.slot.delivery) - (b.slot.paperwork || b.slot.delivery));

    return {
      now, cfg, orders, slots, used, moves, shopping, exceptions: ex,
      unconfirmedAreas: [...usedAreaCaps],
      summary: {
        orders: orders.length,
        trailers: upcoming.length,
        underBefore: upcoming.filter((s) => s.ownLoad + 1e-9 < cfg.targetFill && s.orders.some((o) => o.home === s.id)).length,
        under: upcoming.filter((s) => s.verdict === 'under' || s.verdict === 'low').length,
        moves: shopping.length,
        freed: upcoming.filter((s) => s.lostAll).length,
        empty: upcoming.filter((s) => s.verdict === 'empty').length,
        noSlot: ex.filter((e) => e.kind === 'noslot').length,
        decide: ex.filter((e) => e.kind === 'decide').length,
        over: upcoming.filter((s) => s.verdict === 'over').length,
      },
    };
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
    DAY, DEFAULT_CONFIG, parsePicking, parseTransport, plan, sizeOrder, customerOf, detect,
    fmtTime, fmtDay, fmtDayTime, absToLocalInput, localInputToAbs, dayToISO, isoToDay, weekday,
    _internal: { parseDay, parseTime, parseDateTime, parseLate, chooseTemplateDay },
  };
});
