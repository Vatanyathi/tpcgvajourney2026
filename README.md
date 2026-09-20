# TPC Growth & Value-Add Journey — v1.02c deployable build

This is a synced copy of everything built through v1.02c in the Claude
session — the Track A/B content rewrite matching your Thai reference
document, the redesigned growth-stage selector, the manager-question
bolding/styling fixes, the two-line stage-table labels, and the
Username1-13 test roster (plus Punjamaporn/Charinee as P&O Admins) —
wired to a real backend so anyone with the link can use it, no Claude
account needed.

## How this works

Your app already talks to storage through four simple operations:
get / set / list / delete. That didn't change. What changed is *where*
those calls go:

- **Before**: `window.storage` (only exists inside a Claude.ai chat)
- **Now**: `src/storageShim.js` intercepts those same calls and sends them
  to `/api/kv`, a small serverless function that stores everything in
  **Vercel KV** (a real, persistent database). `src/GVAJourney.jsx` carries
  your app's real content and design, plus the login/password hardening
  described further down — it's not a byte-for-byte copy of the Claude
  artifact anymore, on purpose.

## Deploy — updating your existing site

You already have a repo (`Vatanyathi/gva2026form`) and a Vercel project
connected to it. To push this v1.02 update:

```bash
git clone https://github.com/Vatanyathi/gva2026form.git
cd gva2026form
# copy every file from this package into the repo, overwriting what's there
cp -r /path/to/this/package/. .
git add -A
git commit -m "v1.02: Track A/B content rewrite, redesigned stage selector, text/UX fixes"
git push
```

Pushing triggers an automatic Vercel redeploy. Your existing Redis
connection and environment variables carry over — nothing to reconfigure.

## Deploying to a brand new project instead (~10 minutes)

If you ever need to set this up somewhere fresh:

1. **Push this folder to a GitHub repo.**

2. **Import it into Vercel**: [vercel.com/new](https://vercel.com/new) →
   pick this repo → Vercel auto-detects Vite, no config needed → click
   Deploy once (it'll build fine even before the database exists).

3. **Add a Redis database**: Vercel's own "Vercel KV" product is
   deprecated, so use the current path instead — in your project:
   **Storage** tab → **Browse Marketplace** → search **Redis** → pick an
   Upstash-backed Redis integration → create and **connect** it to this
   project. Vercel/Upstash will inject the REST URL and token as
   environment variables automatically — `api/kv.js` checks a couple of
   likely names for these, so it should work without you renaming
   anything, but if it errors on first use, check Project → Settings →
   Environment Variables for the exact names it added and match them in
   `api/kv.js` if they differ.

4. **(Recommended) Add the shared secret** so random internet traffic
   can't hit your API: Project → **Settings** → **Environment Variables**
   → add both:
   - `API_SECRET` = some random string
   - `VITE_API_SECRET` = the *same* string

   (This isn't real per-user security — see the security note below —
   but it stops drive-by bots from finding and hitting the endpoint.)

5. **Redeploy** (Deployments tab → ⋯ → Redeploy) so the new env vars take
   effect.

6. Open the deployed URL. That's it — it's live.

## Deploy via CLI instead (if you'd rather skip GitHub)

```bash
npm install -g vercel
cd tpc-gva-journey
vercel          # first deploy, follow the prompts
vercel --prod
```

Then add the Redis integration and env vars from the Vercel dashboard
(step 3–4 above) — the Marketplace integration isn't something the CLI
sets up for you.

## Local development

```bash
npm install
npm run dev
```

Note: `/api/kv.js` only runs on Vercel's own infrastructure (or via
`vercel dev`, which proxies it locally) — plain `vite dev` alone won't
serve the API route. Use `vercel dev` instead of `npm run dev` if you
need the backend working locally too.

## Before real employees start using this — please read

Since you're now deploying for real staff, I closed the biggest gap from
the first version:

**Passwords are no longer plaintext.** Login now goes through
`/api/login`, which checks credentials server-side and hashes with
scrypt (Node's built-in crypto — no extra dependency). The browser never
receives anyone's password or hash, only a yes/no plus that one person's
own record. Existing plaintext passwords (from the seed data or an Excel
import before this update) upgrade to a real hash automatically the
first time that person logs in successfully — no separate migration
step needed. Setting or changing a password (one at a time in People, or
in bulk via Excel import) now goes through `/api/set-passwords`, which
hashes server-side before it's ever stored.

**Still open, worth knowing:**

1. **No session/role check on the API yet.** `/api/kv` and
   `/api/set-passwords` trust the single shared secret every browser
   sends, not who's actually asking. The app's own role checks (Master
   Admin vs. P&O vs. Staff) still only run in the browser — a technically
   determined person could bypass them and call the API directly. The
   real fix is a signed session token issued at login and checked on
   every subsequent write, so the server — not just the browser — knows
   who's asking and what they're allowed to do.
2. **The roster (now with hashes, not plaintext) still loads into the
   browser before anyone logs in**, because the login page's site
   picker and several other things currently expect it. That's a much
   smaller exposure than before, but not zero.

Neither blocks going live with real people today. I'd treat #1 as the
next real priority once this is live and being used, since it's the gap
between "the UI won't let you" and "the server won't let you" — the
second one is the one that actually holds under a determined attempt.
Happy to build that next whenever you're ready.

## Getting this into your existing GitHub repo + Vercel project

You already have:
- Repo: `github.com/Vatanyathi/gva2026form`
- Vercel project: `vercel.com/pn-o/gva2026form` (presumably imported from
  that repo already — if not, do that first: Vercel → Add New → Project
  → import `Vatanyathi/gva2026form`)

To get this code in:

```bash
git clone https://github.com/Vatanyathi/gva2026form.git
cd gva2026form
# copy every file from this package into the repo, overwriting anything
# that's there (replace /path/to/this/package with wherever you unzipped it)
cp -r /path/to/this/package/. .
git add -A
git commit -m "Deployable build: real backend + hashed passwords"
git push
```

Pushing to the branch Vercel watches (usually `main`) triggers an
automatic deploy. Then:

1. **Storage** tab in the Vercel project → **Browse Marketplace** →
   search **Redis** → add an Upstash-backed integration → **Connect** it
   to this project (this is the current path — Vercel's own "KV" product
   is deprecated).
2. **Settings → Environment Variables** → add `API_SECRET` and
   `VITE_API_SECRET` with the same random value in both, if you want the
   shared-secret protection described above.
3. **Deployments** tab → redeploy so the env vars take effect.
4. Open the Vercel URL — sign in with any account from your roster.

