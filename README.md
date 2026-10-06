# youreca
## FOMC Date Repricer (`index.html`)

A single-file web page that reprices US rates futures strategies (ZQ Fed Funds, SR1 1M SOFR, SR3 3M SOFR: outrights, spreads, flies, condors and custom baskets) when the FOMC meeting calendar changes.

- Open `index.html` in any browser. It has no server, build step or dependencies, and all calculation runs locally and instantly.
- **Schedule A** holds the current meeting dates and the expected move at each one (you can imply these from ZQ or SR1 market prices).
- **Schedule B** is where you paste the newly announced dates. Every strategy is repriced right away, and the page shows fair A, fair B, the change, and the edge against market prices.
- Inputs are saved in the browser's local storage.
