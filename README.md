# Beach & Bliss Mirissa — Revenue Manager

Production version, connected to your real Supabase project (`BnB Revenue Manager`).

## 1. First-time setup

```bash
npm install
cp .env.example .env
```

`.env` is already pre-filled with your project's URL and anon key. Never commit `.env`
to a public repo — it's already in `.gitignore`-worthy territory even though the anon
key is safe for client use (protected by Row Level Security).

## 2. Run it locally

```bash
npm run dev
```

Open the printed local URL (usually `http://localhost:5173`). Test the full flow:
add an income entry, check it appears in Supabase (Table Editor → transactions),
unlock admin with PIN `0105`, check reports.

## 3. Deploy it so staff can use it on their phones

Easiest option: **Vercel** (free tier is enough).

```bash
npm install -g vercel
vercel
```

Follow the prompts. When it asks about environment variables, add
`VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` with the values from your `.env`.

Once deployed, open the Vercel URL on a phone and use "Add to Home Screen" —
it installs like a native app (this is a PWA).

Alternative: Netlify works the same way (`netlify deploy`).

## 4. What's already wired up

- All data (transactions, rooms, categories, budgets, admin PIN) reads and writes
  directly to your Supabase project — no local/offline storage involved.
- Edits to saved entries are logged to `transaction_history` and shown to everyone
  as an "Edited" badge with full change history.
- Booking.com income automatically accrues an 18% commission, deducted from net
  figures on the dashboard and reports.
- Staff/admin split is enforced by a PIN gate in the UI. See the security note below
  — this is not yet enforced by the database itself.

## Important: current security level

Row Level Security is enabled on every table, but the policies currently allow
**any request carrying the anon key** to read and write everything. That means:

- The staff/admin separation you see in the app (staff can't see totals/reports)
  is enforced by the *front-end code only*. A technically inclined person could
  bypass it by calling the Supabase API directly with the same anon key embedded
  in the app.
- This is an appropriate trust level for a small, trusted team testing the app,
  but **not sufficient before wider staff rollout**.

### Recommended next step: real Supabase Auth

To make the split genuinely server-enforced:

1. Enable **Email** (or magic link) auth in Supabase → Authentication.
2. Create one account per staff member, tagged with a `role` (`staff` / `admin`)
   in a `profiles` table linked to `auth.users`.
3. Replace the current RLS policies (`using (true)`) with policies that check
   `auth.jwt() ->> 'role'` — e.g. only `admin` role can `select` from
   `transactions` where you want amounts hidden, or use a Postgres view that
   omits the `amount` column for staff.
4. Replace the in-app PIN screen with real sign-in (Supabase Auth UI or a
   simple email/password form).

Ask Claude to build this next when you're ready — it's a well-scoped follow-up
once you've confirmed the day-to-day workflow works for your team.
