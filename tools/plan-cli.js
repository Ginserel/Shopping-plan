#!/usr/bin/env node
/*
 * Command-line shopping plan.
 *   node tools/plan-cli.js <picking-status.xls> <transport-plan.xlsx> [--now "2026-11-23T16:30"] [--json]
 * Prints the shopping list and trailer fills. --json prints the full plan for automations.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const SP = require(path.join(__dirname, '..', 'src', 'core.js'));

function readRows(file) {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: false });
  // Use the sheet that parses best.
  let best = null;
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const d = SP.detect(rows);
    if (d.type && (!best || rows.length > best.rows.length)) best = { rows, d };
  }
  return best ? best.d : { type: null };
}

const args = process.argv.slice(2);
const files = args.filter((a, i) => !a.startsWith('--') && !(args[i - 1] || '').startsWith('--now'));
const nowArg = args.includes('--now') ? args[args.indexOf('--now') + 1] : null;
let picking = null, transport = null;
for (const f of files) {
  const d = readRows(f);
  if (d.type === 'picking') picking = d.data;
  else if (d.type === 'transport') transport = d.data;
  else console.error(`Could not recognise ${f}`);
}
if (!picking || !transport) {
  console.error('Need one picking status export and one transport plan.');
  process.exit(1);
}
const p = SP.plan(picking, transport, {}, { now: nowArg ? SP.localInputToAbs(nowArg) : undefined });

if (args.includes('--json')) {
  const slim = {
    now: SP.fmtDayTime(p.now),
    summary: p.summary,
    shopping: p.shopping.map((m) => ({
      type: m.type, order: m.order.id, wave: m.order.wave, seller: m.order.seller, customer: m.order.customer,
      due: SP.fmtDayTime(m.order.delivery), trailer: SP.fmtDayTime(m.slot.delivery), kind: m.slot.kind,
      paperwork: SP.fmtDayTime(m.slot.paperwork), trailerPct: Math.round(m.order.frac * 100), note: m.note,
    })),
    trailers: p.slots.filter((s) => !s.past).map((s) => ({
      customer: s.customer, delivery: SP.fmtDayTime(s.delivery), kind: s.kind, fillPct: Math.round(s.load * 100),
      verdict: s.verdict, orders: s.orders.map((o) => o.id),
    })),
    exceptions: p.exceptions.map((e) => ({ kind: e.kind, order: e.order && e.order.id, text: e.text })),
  };
  console.log(JSON.stringify(slim, null, 2));
  process.exit(0);
}

const pct = (x) => `${Math.round(x * 100)}%`.padStart(5);
console.log(`Plan as of ${SP.fmtDayTime(p.now)}`);
p.used.forEach((u) => console.log(`  Trailers for ${SP.fmtDay(u.day * SP.DAY)} from transport day ${SP.fmtDay(u.source * SP.DAY)}${u.exact ? '' : ' (template)'}`));
console.log('\nSHOPPING LIST');
for (const m of p.shopping) {
  const o = m.order;
  console.log(`  by ${SP.fmtDayTime(m.slot.paperwork)}  ${o.customer.padEnd(13)} ${SP.fmtTime(m.slot.delivery)} ${m.slot.kind.padEnd(16)} <- ${o.id} wave ${o.wave} ${o.seller.padEnd(28)} due ${SP.fmtTime(o.delivery)} ${pct(o.frac)}  [${m.type}] ${o.remaining}/${o.total} picks left`);
}
console.log('\nTRAILERS');
for (const s of p.slots.filter((x) => !x.past)) {
  console.log(`  ${SP.fmtDayTime(s.delivery)} ${s.customer.padEnd(13)} ${s.kind.padEnd(22)} ${pct(s.load)} ${s.verdict.padEnd(6)} ${s.orders.map((o) => o.seller || o.id).join(', ')}`);
}
console.log('\nEXCEPTIONS');
for (const e of p.exceptions) console.log(`  [${e.kind}] ${e.order ? e.order.id + ' ' + e.order.seller : SP.fmtDayTime(e.slot.delivery) + ' ' + e.slot.customer}: ${e.text}`);
console.log('\nSUMMARY', p.summary);
