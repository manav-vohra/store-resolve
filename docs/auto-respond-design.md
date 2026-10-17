# Auto-respond to Dunkin — design (NOT enabled)

Status: design only. No prod change. No send enabled. No toggle flipped.
Date: 2026-10-15
Goal from Manav: when a Dunkin complaint comes in, the system auto-responds that we are working on it, and Dad can see / act on complaints from his phone.

## What "auto-respond" means here

Reply once, in the same Outlook conversation, to the sender of the original
complaint email that created the case. It does NOT mean:

- a new email to a separate Dunkin address
- an SMS or push to Dad
- a reply to the guest / customer
- a promise, timeline, outcome, or admission

Current template in `worker/ingestion.ts` (`neutralAcknowledgment`):

> We have received this complaint and are reviewing it with the appropriate location. We will follow up as needed.

That wording already matches "we are working on it" without promising a
result. Keep it. Do not add store name, case details, customer PII, or a
deadline unless Manav approves exact new wording.

## What already exists (verified in code)

Trigger today:
- `worker/operations.ts` `runScheduledOperations()` calls `pollEmail()`
- `worker/ingestion.ts` `pollEmail()` only runs if
  `emailIngestionEnabled=true` AND Graph provider is ready. For each message
  in the lookback window it calls `ingestEmailMessage()`, then if a complaint
  was created and `emailAckEnabled=true`, it calls `acknowledgeComplaint()`.

Ack path (`worker/ingestion.ts` `acknowledgeComplaint()`):
- Returns `DISABLED` if `emailAckEnabled=false`.
- Finds the original message row (`gmail_messages`, `is_follow_up=0`) for the
  complaint.
- Idempotency key `email-ack:{complaintId}` in `email_acknowledgments`.
  `PENDING -> IN_FLIGHT -> SENT / FAILED`. Second call returns
  `ALREADY_HANDLED` if the row is not `PENDING`.
- Send is `MicrosoftGraphProvider.sendAcknowledgment()` in
  `worker/microsoft-graph.ts`: `POST /me/messages/{id}/reply` with
  `{ comment: body }`. Graph returns 202, no sent-message ID.
- On success: `complaints.dunkin_acknowledged_at`, `acknowledgement_status`
  = `SENT`, `acknowledgment_body` = template, `gmail_messages`
  `acknowledgment_status` = `SENT`, `DUNKIN_ACKNOWLEDGED` complaint event,
  `ACKNOWLEDGMENT_SENT` integration event.
- On failure: `email_acknowledgments` = `FAILED` + `last_error_code`,
  `complaints.acknowledgement_status` = `FAILED`,
  `ACKNOWLEDGMENT_FAILED` integration event.

Gating that already exists:
- Separate D1 switches: `email_ingestion_enabled`,
  `email_ack_enabled`. Both default off. Pilot Controls shows them as
  independent toggles: "Email ingestion" and "Automatic acknowledgment".
- Pilot one-shot ingest (`ingestPilotEmail`) refuses to run if either
  switch is on.
- `worker/response-workflow.ts` `sendResponseAction()` never sends in
  Phase C. It throws `SEND_DISABLED` or `SEND_NOT_WIRED`. That is a
  different path (manual response actions) and is NOT the auto-ack path.

What is NOT true today:
- Manual / Simulator complaints set `acknowledgement_status=SENT` and a
  template body internally in `src/lib/workflow.ts` (`createComplaint`,
  `acknowledged=true` default). That is an internal record, not proof an
  email was sent. Only the Graph reply path above sends a real email.
- `IDENTITY_UNRESOLVED` mail goes to Review Queue. No complaint is created,
  no ack is sent. Good. Ack must stay limited to cases with a trustworthy
  case ID or conversation link.

## Why nothing sends today (three blockers)

1. Both switches are off in prod D1 (`email_ingestion_enabled=false`,
   `email_ack_enabled=false`, `external_notifications_enabled=false`).
2. Graph provider is NOT READY. Secrets required (`docs/INTEGRATIONS.md`):
   `MS_CLIENT_ID`, `MS_REFRESH_TOKEN`, `MS_MAILBOX_ADDRESS`, `MS_TENANT`,
   delegated scopes `offline_access`, `Mail.Read`, `Mail.Send`, authority
   `consumers`, Device Code Flow via `pnpm oauth:microsoft`. No send can
   happen until readiness + mailbox identity are verified.
3. Scheduled trigger is disabled by design. `wrangler.toml` has no
   `[triggers]` block (commented out, with a note not to re-add the legacy
   tick). `runScheduledOperations` only runs via the `scheduled()` handler.
   There is no manual "send ack for this complaint" endpoint. So even if
   the toggles were on, the ack path would not fire until a trigger exists.

Do NOT fix blocker 3 by re-enabling the monolithic `*/5` cron. That tick
loads unbounded state per run (standing correction 2). The ack trigger
must be targeted.

## Proposed behavior (to approve before building)

### Trigger — targeted, not cron sweep

Option A (recommended first): manual + queue, no cron.
- Add an admin-only endpoint `POST /api/complaints/:id/acknowledge`
  that runs `acknowledgeComplaint()` for that one complaint only.
- Dad / Manav sees an "Acknowledge to Dunkin" button on a complaint
  detail only when: complaint came from Graph ingestion, ack status is
  `DISABLED` / `PENDING` / `FAILED` (with retry), and `emailAckEnabled`
  is on. The button shows the exact template before sending and
  requires one confirm tap.
- Later, if volume justifies it, add a scoped scheduled job that only
  queries `email_acknowledgments WHERE status='PENDING'` + the matching
  complaint rows, with a row cap and a D1 read projection. Never the
  full `loadState()` sweep.

Option B (later only): scoped scheduled ack worker.
- Same targeted query, same cap, same idempotency. Still needs the
  manual button as the fallback and the pilot proof first.

Do not auto-ack on Simulator / manual complaints. Those have no source
Outlook message to reply to. `acknowledgeComplaint()` already returns
`FAILED` when no source row exists. Keep that guard.

### Recipient

Reply to the sender of the source message, in-thread, via Graph
`/reply`. Do not construct a To address from the complaint body.
Do not CC anyone unless explicitly approved. In the pilot selector the
approved sender is `customerservice@dunkinbrands.com`, but the code
must reply to the actual source sender of the ingested message, not a
hardcoded address.

Open question for Manav (blocks live enable): confirm the production
mailbox identity (`MS_MAILBOX_ADDRESS`) and that `/me` for the stored
refresh token is that mailbox. If the token belongs to the wrong
account, replies would come from the wrong mailbox. Verify with a
read-only readiness check before any send.

### Idempotency and exactly-once intent

Keep `email-ack:{complaintId}` + `PENDING -> IN_FLIGHT -> SENT` as the
single-send guard. Fix these gaps before live enable:

1. `FAILED` is terminal today. A failed ack can never retry because the
   next call sees non-`PENDING` and returns `ALREADY_HANDLED`. Add an
   explicit admin retry that moves `FAILED -> PENDING` only after the
   failure code is reviewed, with attempt count preserved.
2. `IN_FLIGHT` can stick if the Worker crashes mid-send. Add a stale
   `IN_FLIGHT` reclaim: only an admin action may move a row older than
   e.g. 15 minutes from `IN_FLIGHT -> PENDING`, after checking Graph /
   Sent Items for whether the reply actually went out. Never auto-retry
   a send. `docs/INTEGRATIONS.md` already states: a send is never
   automatically retried.
3. Graph `reply` returns 202 with no ID. Treat 202 as SENT only after
   the D1 batch commits. If the batch fails after Graph accepted, the
   row stays `IN_FLIGHT` and the admin reclaim path above applies.
   Never send a second reply to "make sure".

### Template and content rules

- Fixed template, no free-text ack from the phone.
- No customer name, phone, email, complaint text, or PII in the reply.
- No timeline promise ("within 24 hours"), no outcome promise.
- Keep the current neutral wording unless Manav approves a new exact
  string. The new string must be stored in code + D1
  `acknowledgment_body` so the audit shows what was actually sent.

### Audit and visibility for Dad

Every ack must be visible on the complaint detail, same as today:
- `DUNKIN_ACKNOWLEDGED` event with timestamp + provider
- `acknowledgement_status`: `DISABLED` / `PENDING` / `IN_FLIGHT` /
  `SENT` / `FAILED`
- `email_acknowledgments` row: status, attempt_count, last_error_code,
  sent_at, provider_message_id (null for Graph reply, expected)
- `integration_events`: `ACKNOWLEDGMENT_SENT` / `ACKNOWLEDGMENT_FAILED`
- Dad sees status in the detail view. He does not need to see the
  Graph message ID (there is none) or the raw mailbox address.

### Kill switches and rollout gates

All must be true for any live send:
- `email_ingestion_enabled=true` (only if ingestion is the source)
- `email_ack_enabled=true`
- Graph provider ready + mailbox identity verified
- Manual ack button (Option A) built and tested, OR scoped ack job
  built and tested
- One controlled pilot ack to a known test thread, verified in Sent
  Items + D1, approved by Manav

Any one of these turns sends off immediately:
- `email_ack_enabled=false`
- Graph provider not ready
- Manual button removed / endpoint disabled

Pilot Controls already exposes the two toggles. Do not add a third
hidden switch. The existing switches are the kill switches.

## Rollout plan (no step auto-advances)

0. Manav confirms mailbox identity + expected sender + exact template.
   (Blocks everything live.)
1. Build manual ack endpoint + detail button + FAILED retry + stale
   IN_FLIGHT reclaim. Tests: idempotency, no double-send, FAILED retry,
   no ack for manual / review items, no ack when switch off.
2. Graph readiness check in staging / pilot: token refresh, Inbox
   metadata only, confirm `/me` mailbox matches expected.
3. One controlled pilot ack to a known test thread. Verify: reply
   appears in thread, D1 shows SENT, no duplicate, audit rows complete.
   Manav approves the pilot result in words.
4. Enable for real ingestion only after step 3 approval. Dad sees ack
   status on each case. Monitor first 5 live acks manually.
5. Only if volume justifies it: add scoped ack job (Option B). Never
   re-enable the legacy full-state cron.

## What this design deliberately does NOT change

- Dad seeing all complaints (OWNER sees all, verified). No change.
- Manager vs owner action split (`ACKNOWLEDGE` etc. require assigned
  manager; `CLOSE` / `REOPEN` require owner). No change in this design.
- Drill `SR-2026-0002` deadlines / assignment. Refresh needs separate
  explicit approval.
- Push test kill switch. Still OFF. No real push in this work.
- `SR-2026-0001` real complaint. Untouched.
- Migration 0007. Still deferred, not applied.

## Risks

- Wrong mailbox / wrong sender: reply goes out under the wrong identity.
  Mitigated by mailbox verification gate (step 0 + 2).
- Duplicate reply: Graph 202 with no ID + crash window. Mitigated by
  idempotency + no auto-retry + admin reclaim only.
- Ack to a non-complaint / mis-routed message: mitigated by ack only
  after a complaint is created from a trustworthy case ID / conversation
  link. Review items never ack.
- Template over-promises: mitigated by fixed neutral wording, no
  free-text ack.
- Cron regression: mitigated by refusing the monolithic tick; manual
  button first, scoped job only later.
