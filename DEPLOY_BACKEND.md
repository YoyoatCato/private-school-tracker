# Backend (Cloudflare Worker) — mini-sweep, add-by-link, and learning

`worker.js` powers four things the static dashboard can't do on its own (a GitHub
Pages site has no server and can't call NCES/Census/news sites from the browser):

| Mode (POST JSON) | What it does | Needs |
|---|---|---|
| `{name, state?, town?}` | Enrich a named school: location, NCES address, type, 2026 FFIEC income | nothing |
| `{url}` | **Add-by-link**: classify + fully enrich a pasted news article | nothing |
| `{action:"learn", name?/phrase?/domain?}` | Commit the rejection to `data/learned_filters.json` so the daily sweep learns | `GITHUB_TOKEN` + `GITHUB_REPO` |
| `{action:"analyze", reasons:[...]}` | Ask an LLM to suggest filter rules from your rejection reasons | `LLM_API_KEY` |

The dashboard works without any of this — you just lose the auto-fill/learning
extras. Reject-with-reason, the recycle bin, and on-device auto-hide of rejected
schools all work with **no backend at all** (they use your browser's storage).

## Deploy (≈5 min, free)

1. Create a free Cloudflare account → **Workers & Pages → Create → Create Worker** → Deploy.
2. **Edit code**, paste all of `worker.js`, **Deploy**. Copy the URL (`https://…workers.dev`).
3. In `index.html`, set `const BACKEND_URL = "https://…workers.dev";` and commit.

That turns on the mini-sweep, add-by-link, and (once secrets are set) learning.

## Optional secrets (Worker → Settings → Variables and Secrets)

- **Feed the daily sweep** (write learned filters back to the repo):
  - `GITHUB_REPO` = `YoyoatCato/private-school-tracker`
  - `GITHUB_TOKEN` = a fine-grained GitHub token with **Contents: Read and write** on this repo only.
  - When set, each rejection appends to `data/learned_filters.json`; the sweep reads it
    and skips matching names/phrases/domains next run.
- **AI analysis of rejection reasons:**
  - `LLM_API_KEY` = your OpenAI key (default). Optional `LLM_URL`, `LLM_MODEL` (default `gpt-4o-mini`).
  - Any OpenAI-compatible endpoint works via `LLM_URL`.

## `data/learned_filters.json` shape (created automatically)
```json
{ "muted_names": ["pine bluff high school"], "muted_phrases": ["open house"], "muted_domains": ["example.com"] }
```
You can also hand-edit/commit this file directly — the sweep reads it on every run.

## Other hosts (AWS Lambda / GCP / VPS)
`worker.js` is plain JS using only `fetch`. Wrap the `fetch(request, env)` handler in your
platform's HTTP entry point and keep the same POST-JSON contract. Read secrets from your
platform's env instead of Cloudflare's `env`. Tell me the platform and I'll adapt it.

## Notes & limits
- **NCES** is the 2023–24 survey: brand-new schools won't be in it; ambiguous multi-matches are skipped (never guessed). Address is tagged *unverified*.
- **Add-by-link** classifies opening vs closure and flags `non_event` (open house etc.) so you can reject fast.
- **Learning is per-device** in the browser (localStorage) *plus* server-side once `GITHUB_TOKEN` is set. Clearing the recycle bin does not un-learn; restore an item to un-mute it.
