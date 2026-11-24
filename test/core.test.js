'use strict';
const test = require('node:test');
const assert = require('node:assert');
const SP = require('../src/core.js');
const S = require('../src/sample.js');

const picking = SP.parsePicking(S.picking());
const transport = SP.parseTransport(S.transport());
const run = (cfg, opts) => SP.plan(picking, transport, cfg || {}, opts || {});
const bySeller = (p, ref) => p.orders.find((o) => o.seller === ref);
const slotOf = (p, o) => p.slots.find((s) => s.id === o.slot);

test('reads the picking status export', () => {
  assert.strictEqual(picking.orders.length, 37);
  assert.strictEqual(SP.fmtDayTime(picking.asOf), 'Mon 23/11 16:30');
  const o = picking.orders.find((x) => x.seller === 'CP PNT 241126');
  assert.strictEqual(o.wave, '220122');
  assert.deepStrictEqual(o.areas, { VNA3: 1, VNA4: 1 });
  assert.strictEqual(o.total, 20);
});

test('reads the transport plan and puts late-evening rows on the night before', () => {
  assert.deepStrictEqual(transport.days.map((d) => SP.dayToISO(d.day)), ['2026-11-23', '2026-11-24']);
  const lp = transport.slots.find((s) => s.day === SP.isoToDay('2026-11-24') && s.to === 'Landpower' && /live tip/i.test(s.kind));
  assert.strictEqual(SP.fmtDayTime(lp.paperwork), 'Mon 23/11 23:00');
  assert.strictEqual(SP.fmtDayTime(lp.delivery), 'Tue 24/11 01:00');
  const cabs = transport.slots.find((s) => s.day === SP.isoToDay('2026-11-24') && s.to === 'Cabs' && s.kind === 'OTM');
  assert.strictEqual(SP.fmtDayTime(cabs.delivery), 'Tue 24/11 23:59'); // 00:00 counts as 23:59 the day before
});

test('sizes orders by area, with shovel and skip rules', () => {
  const p = run();
  assert.strictEqual(Math.round(bySeller(p, 'Shovel-1-241126').frac * 100), 100); // 16 shovels fill a trailer
  assert.strictEqual(Math.round(bySeller(p, 'Skips-241126').frac * 100), 20); // 5 skips per trailer
  assert.strictEqual(Math.round(bySeller(p, 'RAYNE-241126 CARRIAGES').frac * 100), 57); // 20 Wide / 35
  assert.ok(bySeller(p, 'LP-241126-TIGER').decide, 'Landpower needs a decision');
});

test('fills the under-filled 07:30 Compact trailer from the 18:00 one', () => {
  const p = run();
  for (const ref of ['CP PNT 241126', 'CP SO 241126-2']) {
    const o = bySeller(p, ref);
    assert.strictEqual(o.moveType, 'shop');
    assert.strictEqual(SP.fmtTime(slotOf(p, o).delivery), '07:30');
  }
  const late = p.slots.find((s) => s.customer === 'Compact' && SP.fmtDayTime(s.delivery) === 'Tue 24/11 18:00' && /MH SO/.test(s.kind));
  assert.ok(late.lostAll, '18:00 Compact trailer flagged as maybe not needed');
});

test('never shops a trailer over the maximum', () => {
  const p = run();
  for (const m of p.shopping.filter((x) => x.type === 'shop')) assert.ok(m.slot.load <= p.cfg.maxFill + 1e-9);
});

test('does not move locked orders', () => {
  const ref = bySeller(run(), 'CP PNT 241126');
  const p = run({}, { overrides: { [ref.id]: { lock: true } } });
  assert.strictEqual(bySeller(p, 'CP PNT 241126').moveType, null);
});

test('hand-set trailer share replaces the estimate', () => {
  const lp = bySeller(run(), 'LP-241126-TIGER');
  const p = run({}, { overrides: { [lp.id]: { pct: 80 } } });
  assert.strictEqual(bySeller(p, 'LP-241126-TIGER').frac, 0.8);
  assert.ok(!bySeller(p, 'LP-241126-TIGER').decide);
});

test('orders with no trailer at their time go on the latest earlier trailer', () => {
  const o = bySeller(run(), 'LDL-TANKS-241126-2'); // due 16:00, no Loadall trailer at 16:00
  assert.strictEqual(o.moveType, 'noslot');
});

test('collections stay off our trailers', () => {
  const p = run();
  const o = bySeller(p, 'HARNESS REWORK-COLL');
  assert.ok(o.collection);
  assert.strictEqual(o.slot, null);
});

test('uses the latest Mon-Thu sheet as a template when the date is missing', () => {
  const p = run({ templateDay: null }, { now: SP.localInputToAbs('2026-11-30T16:30') });
  assert.ok(p.used.some((u) => SP.dayToISO(u.day) === '2026-11-30' && !u.exact && SP.dayToISO(u.source) === '2026-11-24'));
});

test('boxes: Mothersons and Mezzanine are 18 picks per pallet', () => {
  const p = run();
  const o = bySeller(p, 'LDL-QA-241126-SEE NOTES'); // 4 Mothersons boxes
  assert.ok(Math.abs(o.pallets - 4 / 18) < 1e-9);
});

test('P&D picks are not counted twice', () => {
  const o = bySeller(run(), 'PS-241126-7AM'); // 12 picks in Wide/VNA + 2 in P&D
  assert.strictEqual(o.pallets, 12);
});

test('kitting is left out by default', () => {
  const p = run();
  assert.ok(!p.orders.some((o) => /kit/i.test(o.section)));
  assert.strictEqual(p.ignored.kitting, 14);
});

test('decant orders and decant trailers are left out', () => {
  const rows = S.picking();
  const h = rows.findIndex((r) => r[0] === 'Delivery Date');
  rows[h + 1][6] = '4 - Decant';
  const p = SP.plan(SP.parsePicking(rows), transport, {}, {});
  assert.strictEqual(p.ignored.other, 1);
  assert.ok(!p.board.some((s) => /decant/i.test(s.kind)));
});

test('kits: 16 orders to a pallet, and kit lines group by wave', () => {
  const p = run({ includeKitting: true });
  const kits = p.orders.filter((o) => o.rule === 'Cabs kits');
  assert.strictEqual(kits.length, 12);
  assert.ok(Math.abs(kits[0].pallets - 1 / 16) < 1e-9);
  const lines = SP.groupShopping(p.shopping).filter((g) => g.order.rule === 'Cabs kits');
  assert.strictEqual(lines.length, 1);
  assert.strictEqual(lines[0].orders.length, 12);
  assert.strictEqual(bySeller(p, 'F-DEC160009-02-1810').rule, 'Landpower kits');
});

test('X-Dock trailers are not ours, but are suggested for sharing', () => {
  const p = run();
  assert.ok(!p.slots.some((s) => /x-dock/i.test(s.kind)));
  assert.ok(p.exceptions.some((e) => e.kind === 'xdock' && e.slot.customer === 'Heavy'));
  const shared = run({ useXdock: true });
  assert.ok(shared.slots.some((s) => /x-dock/i.test(s.kind)));
});

test('Friday only uses a Friday; other days use any Mon-Thu', () => {
  const pick = SP._internal.chooseTemplateDay;
  const t = { days: [{ day: SP.isoToDay('2026-11-23') }, { day: SP.isoToDay('2026-11-24') }, { day: SP.isoToDay('2026-11-20') }] };
  assert.strictEqual(SP.dayToISO(pick(t, SP.isoToDay('2026-11-26'), {})), '2026-11-24'); // Thu <- latest Mon-Thu
  assert.strictEqual(SP.dayToISO(pick(t, SP.isoToDay('2026-11-27'), {})), '2026-11-20'); // Fri <- Fri
});

test('an order due at 00:00 counts as 23:59 the day before', () => {
  const rows = S.picking();
  const h = rows.findIndex((r) => r[0] === 'Delivery Date');
  rows[h + 1][0] = '25/11/2026 00:00';
  const o = SP.parsePicking(rows).orders[0];
  assert.strictEqual(SP.fmtDayTime(o.delivery), 'Tue 24/11 23:59');
});

test('the board lists every customer trailer, X-Dock included, with pick-by 2h before paperwork', () => {
  const p = run();
  const xd = p.board.find((s) => s.customer === 'Heavy' && s.kind === 'X-Dock');
  assert.strictEqual(xd.role, 'xdock');
  const cp = p.board.find((s) => s.customer === 'Compact' && /MH SO/.test(s.kind) && SP.fmtTime(s.delivery) === '07:30');
  assert.strictEqual(cp.role, 'ours');
  assert.strictEqual(SP.fmtTime(cp.pickBy), SP.fmtTime(cp.paperwork - 120));
  assert.strictEqual(new Set(p.board.map((s) => s.key)).size, p.board.length, 'trailer keys are unique');
});

test('suggestions: every planned order is suggested on its trailer until chosen', () => {
  const p = run();
  const L = SP.buildList(p, {});
  const o = bySeller(p, 'CP PNT 241126');
  const row = L.rows.find((r) => r.suggested.some((x) => x.order === o));
  assert.strictEqual(row.slot.id, o.slot);
  assert.strictEqual(row.suggested.find((x) => x.order === o).reason.kind, 'shop');
  assert.ok(L.pool.some((x) => x.order === o));
});

test('chosen orders stay where the planner put them and suggestions work around them', () => {
  const p0 = run();
  const o = bySeller(p0, 'CP PNT 241126');
  const late = p0.board.find((s) => s.customer === 'Compact' && SP.fmtDayTime(s.delivery) === 'Tue 24/11 18:00' && /MH SO/.test(s.kind));
  const assign = { [o.id]: late.key };
  const p = run({}, { fixed: assign });
  const L = SP.buildList(p, assign);
  const row = L.rows.find((r) => r.key === late.key);
  assert.ok(row.chosen.some((x) => x.id === o.id));
  assert.ok(!L.pool.some((x) => x.order.id === o.id));
  assert.ok(!L.rows.some((r) => r.suggested.some((x) => x.order.id === o.id)), 'not suggested anywhere else');
  assert.strictEqual(bySeller(p, 'CP PNT 241126').slot, late.id);
});

test('list rows flag late orders and give a status the admin can act on', () => {
  const p0 = run();
  const o = bySeller(p0, 'PS-241126-12PM'); // due 12:00, not released
  const after = p0.board.find((s) => s.customer === 'Power Systems' && s.delivery > o.delivery && s.role === 'ours');
  const assign = { [o.id]: after.key };
  const L = SP.buildList(run({}, { fixed: assign }), assign);
  const row = L.rows.find((r) => r.key === after.key);
  assert.ok(row.notes.get(o.id).some((n) => n.kind === 'late'));
  assert.ok(/^To release/.test(row.status));
  assert.deepStrictEqual(row.waves, [o.wave]);
});

test('pick plan table and Excel export use the hand-made sheet columns', () => {
  const XLSX = require('../vendor/xlsx.full.min.js');
  const W = require('../src/xlsx-out.js');
  const p0 = run();
  const L0 = SP.buildList(p0, {});
  const row0 = L0.rows.find((r) => r.suggested.length && r.slot.role === 'ours');
  const assign = Object.fromEntries(row0.suggested.map((x) => [x.order.id, row0.key]));
  const L = SP.buildList(run({}, { fixed: assign }), assign);
  const row = L.rows.find((r) => r.key === row0.key);
  const [line] = SP.pickPlanTable([row], { [row.key]: 'Bay 12' });
  assert.deepStrictEqual(SP.PICK_PLAN_HEAD.slice(0, 8), ['Deliver To', 'Delivery Type', 'Picked By', 'Loaded By', 'Delivery Time', 'Orders to pick', 'Status', 'Comments']);
  assert.strictEqual(line.cells[2], SP.fmtTime(row.slot.paperwork - 120));
  assert.strictEqual(line.cells[5].split('\n').length, row.chosen.length);
  assert.strictEqual(line.cells[7], 'Bay 12');
  const bytes = W.build([{ name: 'Pick Plan', rows: [SP.PICK_PLAN_HEAD.map((v) => ({ v, s: 'head' })), line.cells.map((v) => ({ v, s: 'body' }))] }], XLSX.CFB);
  const back = XLSX.utils.sheet_to_json(XLSX.read(bytes).Sheets['Pick Plan'], { header: 1 });
  assert.strictEqual(back[1][5], line.cells[5]);
});
