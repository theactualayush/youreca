# youreca
## FF Meeting Calculator (`index.html`)

A web version of `FF_Meeting_Calculator_OptimizeTest.xlsx` for pricing US rates futures off the FOMC calendar without Excel freezing at release time.

- Open `index.html` in a browser (it loads `engine.js` and `seed.js` from the same folder). There is no server or build step.
- **Paste new FOMC dates** into the bar at the top. They become calendar B ("Announced") and every structure is repriced in a few milliseconds.
- **Calendar A vs B:** compare the current calendar (or any 2027 scenario, Case 1–8 as in the workbook) with the new one: Fair A, Fair B, change, and edge against live prices.
- **Scenario matrix** prices every structure under every calendar scenario before the announcement.
- **Structures:** ZQ structures, FF 1M flys, SR3 flys, SR3–SR1, SR3–ZQ, SR3–ZQ 4-leg and meeting neutrals, plus outright ZQ/SR1/SR3 with Case 1 / Case 2.
- **Market data** tab: paste the Raw Data tables (ZQ_Out, SR1_Out, SR3_Out, SR1ZQ, ZQ_MS, ZQ_1MS, SR3_3MS), SOFR/EFFR fixings and holidays. `seed.js` holds the workbook's data as an example.

`engine.js` is a line-by-line port of the workbook's `SR3-SR1` sheet and Graphs formulas. `node tests/verify-against-excel.js` checks it against 524 values cached in the workbook (implied cuts, SR3/SR1/ZQ fair values and cases, Graphs structures); all agree to 1e-8.
