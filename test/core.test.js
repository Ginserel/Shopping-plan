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
  assert.strictEqual(picking.orders.length, 23);
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
  const cabs = transport.slots.find((s) => s.day === SP.isoToDay('2026-11-24') && s.to === 'Cabs');
  assert.strictEqual(SP.fmtDayTime(cabs.delivery), 'Wed 25/11 00:00');
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

test('uses the same weekday as a template when the date is missing', () => {
  const p = run({ templateDay: null }, { now: SP.localInputToAbs('2026-11-30T16:30') });
  assert.ok(p.used.some((u) => SP.dayToISO(u.day) === '2026-11-30' && !u.exact && SP.dayToISO(u.source) === '2026-11-23'));
});
