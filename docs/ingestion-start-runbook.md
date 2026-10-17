# Make new tickets appear — start runbook

Date: 2026-10-17
Dad report: "No new tickets came in the app.. just checked and we got a few in last few days"
Verified prod 2026-10-17: only 2 complaints, last background run 2026-09-28, cron schedules=[],
`email_ingestion_enabled=false`, `email_ack_enabled=false`,
`external_notifications_enabled=true` (was false in docs — restore to false first).

Goal: new Dunkin complaint emails become tickets Dad can see, with no surprise sends.
Order matters. Do not turn ingestion + ack on together.

## Correction to earlier plan

The existing one-shot "Ingest one recent Dunkin complaint" button in Pilot
Controls does NOT ingest new mail. It is hardcoded to the Aug 1 pilot case
CCC11122413 / Slow Service / 350909-DD. That case is already in prod as
SR-2026-0001. It cannot bring in last week's tickets.

For new tickets the path is delta discovery:
`POST /api/admin/email/discover` — requires `email_ingestion_enabled=true`,
uses Graph delta cursor, calls the normal `ingestEmailMessage` path.
There is currently no UI button for it. Call it from the browser console
or curl with your Access session, or ask for a "Check mail now" button
to be added.

## Step 1a — Restore safe defaults (2 min, in app)

In StoreResolve > Pilot Controls (OWNER only):
- Rollout mode: `FAMILY_PILOT`
- External notifications: OFF (currently ON in prod — flip back)
- Email ingestion: OFF (leave off for now)
- Automatic acknowledgment: OFF (leave off)

Why: diagnostics refuse to run when external is ON, and we do not want
SMS firing while we test ingestion. Ack stays OFF until tickets flow
and one pilot ack is verified.

## Step 1b — Read-only Graph check (1 min, in browser)

While logged into StoreResolve in the same browser, open:

`https://store-resolve.moogzie.workers.dev/api/admin/email/readiness`

Paste back the exact JSON.

Expected pass:
`{"ok":true,"ready":true,"check":"INBOX_FOLDER_METADATA"}`

This only refreshes the token and reads Inbox folder metadata (id, name,
counts). It reads no message bodies, creates no rows, sends nothing.

If it fails, paste the exact `error` code:
- `MS_GRAPH_NOT_CONFIGURED` — Graph secrets missing / incomplete
- `MS_GRAPH_OAUTH_401` / `MS_GRAPH_OAUTH_FAILED` — refresh token expired
  or revoked. Fix is `pnpm oauth:microsoft` device flow (Manav runs it,
  secrets go straight to Wrangler, never in chat).
- `MS_GRAPH_API_403_SCOPE` — Mail.Read / Mail.Send scope missing
- anything else — paste it, do not proceed to Step 2

Do NOT run the pilot diagnostic endpoints yet. They are locked to the
Aug 1 pilot and require the same safe defaults.

## Step 2 — Controlled preview + selective import (only after Step 1b passes)

Use **Check Mail Now** in Pilot Controls (Phase 1, 2026-10-17). Do NOT use
raw `POST /api/admin/email/discover` for recovering missing mail: raw
discover ingests everything in the delta without preview or selection.

1. In Pilot Controls, confirm External OFF, Ack OFF. Ingestion flag is NOT
   required for Check Mail Now import (explicit selection, not broad sweep).
2. In Check Mail Now, set the window (default Oct 10–17, 2026) and press
   **Preview mail**. Preview is read-only: no complaints, no cursor advance,
   no sends. It shows masked IDs, received date, sender (Dunkin shown fully,
   others masked), subject, detected case ID / store, and classification:
   COMPLAINT_CANDIDATE, FOLLOW_UP_LIKELY, NEEDS_REVIEW, NOT_COMPLAINT,
   ALREADY_INGESTED. Forwards/replies with no preview match show NEEDS_REVIEW,
   never silently ignored.
3. Select only the messages to import. Press **Import selected** and confirm.
   Import uses BACKFILL mode for past windows (no overdue deadlines on
   recovered mail). Each selected id is re-validated against a fresh preview
   listing for the same window. Per-message results are shown: status +
   complaintId or errorCode. Failures do not block other messages.
4. Check in app: Complaints list should now show the new cases. Open one:
   - source should be `MICROSOFT_GRAPH`
   - status will be `ROUTING_REVIEW` or `MANAGER_NOTIFIED`
   - `acknowledgement_status` should stay `DISABLED` (ack is off)
   - no SMS / email should have been sent
5. If any import fails or a case lands in Review Queue, stop and paste the
   result. Review items mean identity could not be resolved or the owner
   selected an ambiguous message — a human resolves them, the system never
   invents a case ID.
6. After verifying, leave ingestion ON only if you want Dad to keep
   getting new tickets via manual checks. For automatic pickup, see Step 3.

First-window note: preview is bounded by the date window you set (max 31
days). It never pulls the whole historical Inbox. The delta cursor is NOT
touched by preview or selective import, so a later scoped delta job starts
from its stored cursor, not from your preview window.

## Step 3 — Automatic pickup (Phase 3, build, do not re-enable legacy cron)

Do NOT re-add `crons = ["*/5 * * * *"]` to `wrangler.toml`. That tick
loads unbounded state per run.

Check Mail Now already exists (Phase 1). Phase 3 is a separate scoped
scheduled discovery job: Graph delta only, narrow D1 queries, durable
locks, bounded work per run, persistent retries, cursor checkpoints, and
volume/auth/cursor safeguards. It must NOT auto-enable ack or SMS.

Decide after Step 2 works: manual Check Mail Now for now, or build the
scoped job after the manual pilot passes + explicit approval.

## Step 4 — Ack last (separate approval)

Keep `email_ack_enabled=false` until:
- tickets are flowing (Step 2 verified)
- mailbox identity confirmed (`/me` = expected mailbox)
- manual ack button built (see `docs/auto-respond-design.md`)
- one pilot ack verified in Sent Items + D1

Then one pilot ack, Manav approval, then live.

## What NOT to do

- Do not flip ingestion + ack + external all on at once.
- Do not run the Aug 1 pilot ingest expecting new tickets.
- Do not re-enable the legacy cron.
- Do not touch SR-2026-0001 (real) or SR-2026-0002 (drill) in this process.
- Do not paste Graph secrets / refresh tokens in chat.
