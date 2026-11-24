# Shopping Plan

The FLM's daily pick plan, built from the **WMS Picking Status** export and the **transport plan**. It lists every trailer for the day with its pick-by, loading and delivery times, suggests which orders go on each one, and lets the FLM choose. The finished list goes to the admin as an Excel file in the same layout as the hand-made "Pick Plan" sheet. The admin uses it to release orders to the pickers.

Kitting and decant are left out. Both can be switched back on under Rules.

## Use it

1. Open `dist/shopping-plan.html` in a browser. It is one file and works offline, so it can live on a shared drive.
2. Click **Load files**, or drop both exports on the page: the Picking Status `.xls` and the transport plan `.xlsx`. The app works out which file is which.
3. Pick the **delivery day**. It starts on the next day that has orders.
4. On the **Plan** tab, each trailer is a row in pick-by order. Its suggested orders sit under it as dashed buttons:
   - click a suggestion to put it on the trailer, or **Add all** for that trailer, or **Add all suggestions** for the whole day;
   - **+ Add an order…** puts any other order on the trailer (same customer first, nearest due time first);
   - **×** takes an order off. Orders can also be dragged between trailers, or back to the side panel;
   - **Not on the list yet** (the side panel) holds every order still to place, each with a one-click button for its suggested trailer;
   - **Undo** reverses the last change;
   - each trailer has a comment box for the admin.
5. On the **Shopping list** tab, check the result and use **Download Excel**, **Copy table** or **Print**.

Nothing goes on the list until the FLM adds it. Choices and comments are saved in the browser, so loading a newer picking export later in the day keeps the list. The status updates from the new export (To release → Released → Pick in progress → Picked).

| Tab | What it's for |
| --- | --- |
| Plan | Trailers for the day with chosen orders, suggestions and fill bars. Side panel of orders not on the list yet. |
| Shopping list | The admin's pick plan: Deliver To, Delivery Type, Picked By, Loaded By, Delivery Time, Orders to pick, Status, Comments, plus waves to release, picks left and trailer fill. |
| Checks | Trailers on the list that are over capacity, have late orders or have passed pick-by. Also orders with no trailer, X-Dock sharing ideas, collections and unknown customers. |
| Orders | Every order, with a button to put it on its suggested trailer. Set a trailer share by hand when the estimate is wrong. |
| Rules | Pick-by lead time, fill target, pallets per area, order rules, customer names. Saved in the browser. |

**The Excel file** has two sheets. **Pick Plan** has the trailer rows in the hand-made layout, with coloured status, wrapped order lists, filters and a frozen header. **Orders** has one row per order with order ID, wave, customer, due time and release status, for the admin to release from.

**Warnings on a trailer:** *over capacity*, *arrives late* (an order due before the trailer delivers), *going early* (pulled forward), *for another customer* (a shared or OTM trailer), *customer collection*, and *pick-by time has passed*.

Nothing is uploaded anywhere. Files are read in the browser.

## How it decides

**Pallets and trailer share of an order.** 1 pick = 1 pallet, except box areas. For each area: pallets ÷ pallets that fill a trailer. Picks already done are assumed to have the same area mix.

| Area / order | Picks per pallet | Pallets per trailer |
| --- | --- | --- |
| Wide, VNA 1-5 | 1 | 35 |
| Bulk internal | 1 | 14 |
| Bulk external | 1 | 8 (range 5–12) |
| Mothersons, Mezzanine (boxes) | 18 (16–20) | 35 |
| P&D | not counted – already in VNA | |
| Shovels (seller ref contains "shovel") | 1 | 16 |
| Skips (seller ref contains "skip") | 1 | 5 |
| Cabs, Landpower and T4 kits | 16 orders per pallet (1 small box each) | 35 |
| Landpower | estimated, flagged for the admin to set | |

**Home trailer.** An order goes on the trailer whose wave number is written in the transport plan's "TRL / Load Check" column. Otherwise it goes on the trailer for the same customer and delivery time. If there is no trailer at that time, it goes on the latest trailer before it.

**Picked by.** Paperwork time ("Loaded by" on the pick plan) minus 2 hours, editable under Rules.

**Which trailers are listed.** Every customer trailer in the transport plan. Transfers, collections and decant trailers are left out. X-Dock trailers and trailers from other sites (e.g. ATL Marchington) are shown greyed. They are never suggested, but orders can be put on them by hand.

**Suggestions.** Every order is suggested for its home trailer. Orders already on the list stay where they were put, and the suggestions for everything else work around them. On top of that:

**Pulling forward.** For each customer, trailers are worked from the earliest. A trailer that already has orders, is under the target (85%), and is more than an hour from its paperwork time takes later orders, earliest due first. It never goes over 100%. By default an order is only pulled:

- from the same delivery day,
- from at most 16 hours later,
- off a trailer that stays at target afterwards. Orders are not taken off trailers that are already full enough.

A trailer whose orders have all moved is flagged **Maybe not needed**. It is never refilled.

**Transport plan days.** Mon–Thu share one schedule. If the plan has no sheet for a date, the most recent Mon–Thu sheet is used (a Friday only uses a Friday). Waves and loading status are ignored for a borrowed day, and its trailers are marked "Expected". The next day's sheet is always read too, because its first rows are the 00:00 trailers, which belong to the day before.

**Midnight.** A delivery at 00:00 counts as 23:59 on the day before.

**X-Dock.** X-Dock trailers belong to the X-Dock department, so orders aren't planned onto them. When one of our trailers stays under target, or an order has no trailer, and an X-Dock trailer goes to the same customer around that time, it's listed under **Could share with X-Dock**. A Rules switch plans onto X-Dock trailers like our own.

## Command line (for automation)

```bash
node tools/plan-cli.js "Picking_Status.xls" "transport.xlsx"
```

```bash
node tools/plan-cli.js "Picking_Status.xls" "transport.xlsx" --now 2026-11-23T16:30 --json
```

## Develop

```bash
npm test
```

```bash
npm run build
```

`src/core.js` holds the parsing, planning and list logic and has no DOM code. `src/xlsx-out.js` writes the styled Excel file. `src/app.js` is the UI. `src/sample.js` is made-up example data in the same layout as the real exports. Real exports are git-ignored because this repository is public.

## Ideas for next steps

- Pull the picking status and transport plan straight from the WMS and the shared transport file on a schedule.
- Learn real pallets per trailer by customer and load type from the transport plan's "Pallet Count" history.
- Handle split loads (one trailer for two customers).
- Estimate picking time per area, so shopping stops when there isn't enough time left to pick.
