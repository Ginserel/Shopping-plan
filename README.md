# Shopping Plan

A shopping list for the picking admin. It reads the **WMS Picking Status** export and the **transport plan**, links every order to its trailer, and finds trailers that would leave half empty. For each one it lists the later orders for the same customer that should be picked early ("shopped") to fill it. It also flags trailers that end up with no orders, so transport can decide whether they still need to run.

Example: the Compact 07:30 trailer has 4 picks on it. The 18:00 Compact orders are pulled forward to fill it, and the 18:00 trailer is flagged as "maybe not needed".

## Use it

1. Open `dist/shopping-plan.html` in a browser. It is one file and works offline, so it can live on a shared drive.
2. Click **Load files**, or drop both exports on the page: the Picking Status `.xls` and the transport plan `.xlsx`. The app works out which file is which.
3. Work down the **Shopping list** tab. It is sorted by when each trailer's paperwork is due. Tick orders as you send them to picking. **Copy for Excel / Teams** copies the list.

Other tabs:

| Tab | What it's for |
| --- | --- |
| Trailers | Every trailer per customer with a fill bar. Solid = own orders, striped = shopped in, line = target. |
| Exceptions | Over capacity, no trailer at the due time, missed trailers, Landpower decisions, collections. |
| Orders | Every order. Set a trailer share by hand (Landpower engines, rubber-track machines) or tick **Keep** to stop it moving. |
| Rules | Picks per trailer by area, order rules, customer names, fill target. Saved in the browser. |

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

**Shopping.** For each customer, trailers are worked from the earliest. A trailer that already has orders, is under the target (85%), and is more than an hour from its paperwork time takes later orders, earliest due first. It never goes over 100%. By default an order is only pulled:

- from the same delivery day,
- from at most 16 hours later,
- off a trailer that stays at target afterwards. Orders are not taken off trailers that are already full enough.

A trailer whose orders have all moved is flagged **Maybe not needed**. It is never refilled.

**Transport plan days.** The schedule repeats every day except Friday. If the plan has no sheet for the date, the same weekday in the file is used. Otherwise any Mon–Thu sheet is used, but a Friday only ever uses a Friday. Waves and loading status are ignored for a borrowed day.

**Midnight.** A delivery at 00:00 counts as 23:59 on the day before.

**X-Dock.** X-Dock trailers belong to the X-Dock department, so orders aren't planned onto them. When one of our trailers stays under target, or an order has no trailer, and an X-Dock trailer goes to the same customer around that time, it's listed under **Could share with X-Dock**. A Rules switch plans onto X-Dock trailers like our own.

**Grouping.** Orders on the same wave that go to the same trailer are one line (e.g. 38 Cabs kits).

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

`src/core.js` holds the parsing and planning logic and has no DOM code. `src/app.js` is the UI. `src/sample.js` is made-up example data in the same layout as the real exports. Real exports are git-ignored because this repository is public.

## Ideas for next steps

- Pull the picking status and transport plan straight from the WMS and the shared transport file on a schedule.
- Learn real pallets per trailer by customer and load type from the transport plan's "Pallet Count" history.
- Handle split loads (one trailer for two customers).
- Estimate picking time per area, so shopping stops when there isn't enough time left to pick.
