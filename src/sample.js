/*
 * Example data, shaped exactly like the two real exports so it goes through the
 * same parser. All order numbers, waves and references are made up.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SPSample = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const AREAS = ['Wide', 'Mthr sons', 'VNA1', 'VNA2', 'VNA3', 'VNA4', 'VNA5', 'BULKINT', 'MEZZ', 'BULKEXT', 'P&D', 'CABSKIT', 'LPKIT', 'T4'];
  const ASOF_H = 16, ASOF_M = 30; // "now" in the example: Mon 23/11/2026 16:30

  function late(dd, hh, mm) {
    // minutes from 23/11 16:30 to dd/11 hh:mm
    const left = (dd - 23) * 1440 + hh * 60 + mm - (ASOF_H * 60 + ASOF_M);
    const d = Math.floor(left / 1440), h = Math.floor((left % 1440) / 60), m = left % 60;
    return (d ? `${d} days, ` : '') + `${h} h, ${m} min left`;
  }

  // [time, orderId, wave, seller ref, company, status, release, comments, total, remaining, areas{}]
  const ORDERS = [
    ['06:00', '0000501101', '0000220101', 'Skips-241126', 'Backhoe Loaders', 'Allocated', 'Not Released', '', 1, 1, { BULKEXT: 1 }],
    ['06:00', '0000501102', '0000220102', 'BHL-PRIO-BAY12-231126', 'Backhoe Loaders', 'Allocated', 'Not Released', '', 9, 9, { Wide: 6, VNA4: 3 }],
    ['06:00', '0000501103', '0000220103', 'LDL-ENG-241126-01', 'Loadall', 'Allocated', 'Not Released', '', 5, 5, { Wide: 3, 'Mthr sons': 2 }],
    ['06:00', '0000501104', '0000220104', 'LDL-QA-241126-SEE NOTES', 'Loadall', 'Allocated', 'Not Released', 'Label parts NOT FOR PRODUCTION', 4, 4, { 'Mthr sons': 4 }],
    ['07:00', '0000501105', '0000220105', 'PS-241126-7AM', 'Power Systems', 'Allocated', 'Not Released', '', 12, 12, { Wide: 6, VNA2: 3, VNA5: 3 }],
    ['07:00', '0000501106', '0000220106', 'HP-241126-01', 'Heavy Products', 'Allocated', 'Not Released', '', 14, 14, { Wide: 1, 'Mthr sons': 10, VNA2: 1, BULKINT: 2 }],
    ['07:30', '0000501107', '0000220107', 'CP SO 241126-1', 'Compact Products', 'In Progress', 'Released', '', 2, 1, { 'Mthr sons': 1 }],
    ['07:30', '0000501108', '0000220108', 'CP SO 241126-3', 'Compact Products', 'Allocated', 'Released', '', 2, 2, { VNA4: 2 }],
    ['08:00', '0000501109', '0000220109', 'RAYNE-241126 CARRIAGES', 'Rayne Engineering', 'Allocated', 'Not Released', '', 20, 20, { Wide: 20 }],
    ['10:00', '0000501110', '0000220110', 'PS-241126-10AM', 'Power Systems', 'Allocated', 'Not Released', '', 34, 34, { Wide: 20, VNA2: 6, VNA4: 4, VNA5: 4 }],
    ['10:00', '0000501111', '0000220111', 'HARNESS REWORK-COLL', 'Harness Supplier', 'Allocated', 'Not Released', 'CUSTOMER TO COLLECT', 6, 6, { 'Mthr sons': 6 }],
    ['11:00', '0000501112', '0000220112', 'LDL-241126-11AM', 'Loadall', 'Allocated', 'Not Released', '11AM BUS STOP REQUIREMENT', 16, 16, { Wide: 16 }],
    ['11:00', '0000501113', '0000220113', 'EM-241126-TIGER', 'Earthmovers', 'In Progress', 'Released', '', 20, 19, { Wide: 14, VNA2: 4, VNA5: 1 }],
    ['11:00', '0000501114', '0000220114', 'EM-241126-11:00', 'Earthmovers', 'Allocated', 'Released', '', 22, 22, { Wide: 10, 'Mthr sons': 6, VNA2: 3, VNA4: 1, VNA5: 1, BULKINT: 1 }],
    ['12:00', '0000501115', '0000220115', 'RAYNE-241126 STABS', 'Rayne Engineering', 'In Progress', 'Released', '', 10, 1, { VNA3: 1 }],
    ['12:00', '0000501116', '0000220116', 'PS-241126-12PM', 'Power Systems', 'Allocated', 'Not Released', '', 13, 13, { Wide: 3, VNA2: 2, VNA3: 2, BULKINT: 6 }],
    ['13:00', '0000501117', '0000220117', 'LDL-TANKS-241126-1', 'Loadall', 'Allocated', 'Not Released', '', 19, 19, { Wide: 2, BULKINT: 17 }],
    ['13:00', '0000501118', '0000220118', 'Shovel-1-241126', 'Backhoe Loaders', 'Allocated', 'Not Released', '', 16, 16, { Wide: 16 }],
    ['14:00', '0000501119', '0000220119', 'PS-241126-2PM', 'Power Systems', 'Allocated', 'Not Released', '', 10, 10, { Wide: 1, VNA2: 2, VNA3: 2, VNA4: 2, BULKINT: 3 }],
    ['16:00', '0000501120', '0000220120', 'LDL-TANKS-241126-2', 'Loadall', 'Allocated', 'Not Released', '', 9, 9, { BULKINT: 9 }],
    ['18:00', '0000501121', '0000220121', 'LP-241126-TIGER', 'Landpower', 'Allocated', 'Not Released', '', 11, 11, { Wide: 6, VNA3: 5 }],
    ['18:00', '0000501122', '0000220122', 'CP PNT 241126', 'Compact Products', 'In Progress', 'Released', '', 20, 2, { VNA3: 1, VNA4: 1 }],
    ['18:00', '0000501123', '0000220123', 'CP SO 241126-2', 'Compact Products', 'Allocated', 'Not Released', '', 1, 1, { VNA4: 1 }],
  ];

  function picking() {
    const head = ['Delivery Date', 'Order ID', 'Wave ID', null, 'Seller Order ID', null, 'Order Type', 'Company', null, 'Status', 'Task Release', 'Comments', 'LATE', 'Total Picks', 'Rmn. Picks', 'Picked%', ...AREAS];
    const rows = [
      [null, null, null, null, null, null, null, null, null, null, 'EXAMPLE Picking Status'],
      ['Standard and Decant Orders'],
      ['DelDate: 24/11/2026'],
      head,
    ];
    let tot = 0, rem = 0;
    for (const [t, id, wave, seller, company, status, release, comments, total, remaining, areas] of ORDERS) {
      const [hh, mm] = t.split(':').map(Number);
      const pct = Math.round(((total - remaining) / total) * 100) + ' %';
      rows.push([`24/11/2026 ${t}`, id, wave, null, seller, null, '0 - Standard', company, null, status, release, comments, late(24, hh, mm), total, remaining, pct, ...AREAS.map((a) => areas[a] || null)]);
      tot += total; rem += remaining;
    }
    rows.push([null, 'Total Orders:', ORDERS.length, null, null, null, null, null, null, null, null, null, null, tot, rem]);
    return rows;
  }

  // [paperwork, despatch, delivery, deliver to, delivery type, wave refs]
  const SLOTS = [
    ['23:00', '23:45', '01:00', 'Landpower', 'X-Dock LIVE TIP ONLY', ''],
    ['01:00', '02:20', '03:00', 'Power Systems', 'X-Dock', ''],
    ['04:00', '05:00', '06:00', 'Loadall', 'ENGINE / OTM', ''],
    ['04:00', '05:00', '06:00', 'Backhoe', 'OTM', ''],
    ['05:00', '06:10', '07:00', 'Heavy', 'Blowout', ''],
    ['05:00', '06:20', '07:00', 'Power Systems', 'OTM', ''],
    ['05:30', '06:15', '07:30', 'Compact', 'MH SO Orders', ''],
    ['06:00', '06:40', '08:00', 'Rayne', 'OTM', ''],
    ['08:00', '09:20', '10:00', 'Power Systems', 'Raw Blocks & Heads', ''],
    ['09:00', '09:45', '11:00', 'Earthmovers', 'TIGER', ''],
    ['09:00', '10:00', '11:00', 'Loadall', 'Kingspan', ''],
    ['10:00', '10:40', '12:00', 'Rayne', 'OTM', ''],
    ['10:00', '11:20', '12:00', 'Power Systems', 'OTM', ''],
    ['11:00', '12:00', '13:00', 'Backhoe', 'Shovels', ''],
    ['11:00', '12:00', '13:00', 'Loadall', 'Tanks', ''],
    ['12:00', '13:20', '14:00', 'Power Systems', 'OTM', ''],
    ['13:00', '13:45', '15:00', 'Compact', 'Engines', ''],
    ['14:00', '15:00', '16:00', 'Rocester', 'Blowout', ''],
    ['16:00', '16:45', '18:00', 'Landpower', 'ACO & Kits', ''],
    ['16:00', '16:45', '18:00', 'Compact', 'MH SO Orders', ''],
    ['16:00', '17:00', '18:00', 'Loadall', 'ACO', ''],
    ['20:00', '21:10', '22:00', 'ATL Site', 'Transfers', ''],
    ['22:00', '23:10', '00:00', 'Cabs', 'OTM', ''],
  ];

  function transport() {
    const head = ['Date', 'Hyperlink', 'TRL / Load Check', 'Pallet \nCount', 'Trailer No', 'Bay', 'Delivery / Collection From', 'Deliver To', 'Delivery Type', 'Load Type', 'SPLIT', 'EMER', 'Paperwork Required By', 'Despatch Time', 'Delivery Time', 'Loading Team', 'Loading Status', 'RUN NUMBER'];
    const rows = [head];
    const day = (d, title) => {
      rows.push([d, title]);
      for (const [p, s, t, to, kind, trl] of SLOTS) {
        const loadType = /transfer/i.test(kind) ? 'Transfer' : 'Standard';
        rows.push([d, null, trl, null, null, null, 'Maersk EMG', to, kind, loadType, false, null, p, s, t, null, null, null]);
      }
      rows.push([d, 'Additional / Rolled Orders']);
      rows.push([d, null, 'LINE NOT USED']);
    };
    day('23/11/2026', 'Monday 23rd November 2026');
    day('24/11/2026', 'Tuesday 24th November 2026');
    return rows;
  }

  return { picking, transport, label: 'Example: picking status Mon 23/11 16:30, trailers for Tue 24/11' };
});
