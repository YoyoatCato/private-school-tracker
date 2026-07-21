# Mini-sweep backend (Cloudflare Worker)

`worker.js` is the on-type mini-sweep: when you type a school name into a "Needs a name"
card, the dashboard sends the name to this backend, which returns the school's location,
street address (from NCES), school type, and the **2026 FFIEC MSA/MD median family income**
(via the open US Census geocoder + an embedded copy of the FFIEC 2026 table). No API keys.

The dashboard already works without it — naming a card just moves it to the queue. The
backend only adds the automatic enrichment.

## Deploy (≈5 minutes, free)

1. Create a free Cloudflare account at https://dash.cloudflare.com (Workers & Pages).
2. **Workers & Pages → Create → Create Worker.** Give it a name (e.g. `school-tracker-sweep`), Deploy.
3. Click **Edit code**, delete the sample, paste the entire contents of `worker.js`, then **Deploy**.
4. Copy the Worker URL (looks like `https://school-tracker-sweep.<you>.workers.dev`).
5. In `index.html`, find the line near the top of the `<script>`:
   ```js
   const BACKEND_URL = "";
   ```
   Paste your Worker URL between the quotes and commit `index.html`.

That's it — typing a name now runs the mini-sweep and fills location/address/type/income.

### Prefer AWS Lambda / Google Cloud Functions / a VPS instead?
The logic in `worker.js` is plain JavaScript using only `fetch`. To run it elsewhere,
wrap the `fetch(request)` handler in your platform's HTTP entry point and keep the same
`POST {name, state?, town?}` → JSON contract. Tell me which platform and I'll adapt it.

## What the mini-sweep returns
`POST {"name": "...", "state": "CA", "town": ""}` →
`{ "type", "state", "region", "town", "full_address", "tract_mfi", "data_year", "mfi_area" }`
(only the fields it can confirm; blanks are left for you to fill — it never guesses an
ambiguous NCES match, and income comes straight from the FFIEC 2026 table).

## Notes & limits
- **NCES** is the 2023–24 private-school survey: brand-new schools won't be in it yet, and
  common names with several matches are skipped (never guessed). Address is tagged *unverified*.
- **Location web-search fallback:** if a card has no state, the Worker does a best-effort
  DuckDuckGo lookup to find one. That's the least reliable step; confirm it on review.
- **Income** is the 2026 FFIEC *Estimated MSA/MD/non-MSA Median Family Income* (the figure in
  your FFIEC geomap screenshot), matched by the address's Metropolitan Division/MSA.
