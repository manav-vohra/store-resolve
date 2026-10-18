STORE RESOLVE — BRIEF FOR REVIEW (2026-10-17)

SYSTEM
- StoreResolve: Dunkin franchise complaint ops. Prod: https://store-resolve.moogzie.workers.dev
- Stack: Cloudflare Workers + D1 (store-resolve), Cloudflare Access auth, Microsoft Graph mail (Outlook), SignalWire SMS (trial), Web Push PWA Stage A deployed (c1a4e65e).
- Users in scope: Dad only. Dad = id dad, OWNER, ajvohra1@yahoo.com, iPhone PWA installed via Add to Home Screen, push subscription exists (created 2026-10-15, Apple endpoint, not revoked, last_used null = no push sent yet). Manav = id father, OWNER, manavvohra2021@gmail.com. Other owners/managers parked, out of scope.
- Goal: Dad sees new Dunkin complaints in app, can act on them from phone, system auto-responds to Dunkin that we are working on it.

CURRENT PROD STATE (verified 2026-10-17)
- Complaints: only 2.
  - SR-2026-0001, external CCC11122413, 2026-08-01, source MICROSOFT_GRAPH, status ROUTING_REVIEW, store-350909, acknowledgement_status DISABLED. Real case, untouched.
  - SR-2026-0002, external TEST-DRILL-001, 2026-10-02, source MANUAL, status MANAGER_NOTIFIED, assigned to dad, acknowledgement_status DISABLED, ack deadline 2026-10-02 (overdue). Drill case, do not touch without approval.
- gmail_messages: only 2 rows, both old. background_job_runs: last 2026-09-28. Cron schedules = []. Last Graph ingestion: 2026-09-26 pilot.
- Dad report 2026-10-17: "No new tickets came in the app.. just checked and we got a few in last few days." No example provided yet: no date / subject / case ID for those recent Dunkin emails.
- Config verified via GET /api/bootstrap after Step B:
  mode FAMILY_PILOT, externalNotificationsEnabled=false, emailIngestionEnabled=true, emailAckEnabled=false, managerAckDeadlineMinutes 30, managerResolutionTargetHours 24, escalationIntervalMinutes 60, signalWireReconcileAfterMinutes 10, emailLookbackDays 30.
  Note: externalNotificationsEnabled was true in prod before Step B (was false in docs, unknown who flipped). Flipped back to false in Step B.
- Push-test kill switch OFF. No real push sent. No real email ack sent.

DIAGNOSIS (verified)
- Not a UI bug. Dad as OWNER sees all complaints in code (API publicState + App visibleComplaints). Cause is ingestion + scheduler off.
- Ingestion was OFF (email_ingestion_enabled=false), scheduler OFF (schedules=[]), so pollEmail never ran and no new mail was read.

WHAT WAS DONE TODAY (2026-10-17)
- Step 1 DONE: Graph readiness GET /api/admin/email/readiness via Manav Access login returned {"ok":true,"ready":true,"check":"INBOX_FOLDER_METADATA"}. Token refresh works, mailbox reachable. Read-only, no bodies, no rows, no sends.
- Step 2 PARTIAL: Step B done via Pilot Controls UI toggles (browser cannot PUT/POST): external OFF, ingestion ON, ack OFF. Verified via bootstrap. Step C discover NOT run. No pilot-ingest used, no sends.
- Docs written: docs/auto-respond-design.md (ack design, NOT enabled), docs/ingestion-start-runbook.md.

WHY DISCOVER IS BLOCKED
- Browser agent can GET + click UI toggles only. No fetch / JS console / PUT / POST. No Discover button in UI.
- Pilot Controls "Ingest one recent Dunkin complaint" button is hardcoded to Aug 1 pilot case CCC11122413 / Slow Service / 350909-DD. That case is already SR-2026-0001. It cannot ingest new mail. Must NOT use for new tickets.
- Discover path: POST /api/admin/email/discover {}, requires emailIngestionEnabled=true, uses Graph delta cursor, max 10 pages, calls ingestEmailMessage. First run with no stored cursor returns current Inbox set then deltaLink cursor. Duplicate guard: DUPLICATE -> skipped.

ACK PATH (exists, OFF)
- worker/ingestion.ts acknowledgeComplaint(): requires emailAckEnabled=true. Finds original gmail_messages row (is_follow_up=0). Idempotency key email-ack:{complaintId} in email_acknowledgments, PENDING -> IN_FLIGHT -> SENT / FAILED. Send = MicrosoftGraphProvider.sendAcknowledgment(): POST /me/messages/{id}/reply {comment: fixed neutral template}. Graph returns 202, no sent ID.
- Template (fixed, in code): "We have received this complaint and are reviewing it with the appropriate location. We will follow up as needed." No PII, no timeline, no outcome promise.
- Gaps: FAILED is terminal (next call returns ALREADY_HANDLED, no retry). IN_FLIGHT can stick if Worker crashes. No auto-retry on send (by design). Ack only for Graph-ingested complaints with source message. Review / IDENTITY_UNRESOLVED mail never acks. Manual / Simulator complaints show acknowledgement_status SENT internally, but that is NOT a real email. Only Graph reply is a real send.
- Phase C response-workflow sendResponseAction() never sends (throws SEND_DISABLED / SEND_NOT_WIRED). Different path, not the auto-ack path.

DECISIONS ALREADY MADE (do not relitigate without reason)
- Dad-only pilot. Push primary, email fallback, SMS optional only, never required.
- Manual ack button first, targeted. No monolithic cron. Do NOT re-enable legacy */5 cron (caused quota incident, loads unbounded state per run).
- No Cloudflare Access bypass, no service token, no IP allowlist, no disabling Access. Access stays the auth boundary.
- No silent deadline / assignment changes to SR-2026-0002. No prod deploys / sends / toggles without Manav explicit approval.
- No inventing case IDs. Review items need human. No PII in push payloads.

OPEN ITEMS / QUESTIONS
1. How to run Discover once now: user runs one-line fetch in own browser console on app origin, vs build a "Check mail now" button in Pilot Controls + deploy, vs other route?
2. After Discover works, how to make ingestion automatic without legacy cron: scoped scheduled job (delta only, row cap, D1 read projection) vs manual Discover presses?
3. Ack: fix FAILED retry + stale IN_FLIGHT reclaim + build manual ack endpoint/button before any live ack? Design in docs/auto-respond-design.md.
4. Mailbox identity: how to confirm Graph /me == expected mailbox before any ack send, to avoid replying from wrong mailbox?
5. First delta run volume: with no stored cursor, first Discover may ingest a batch from whole Inbox, not just last week. How to bound / preview / avoid bulk?
6. externalNotificationsEnabled was true in prod before Step B. Implications? SMS trial restriction: SignalWire trial only allows messaging to trial-paired number; Dad sends fail "[10000] To must send to a verified caller id". Push is primary, SMS not required.
7. Dad drill SR-2026-0002 is stale / overdue. Refresh deadlines (with approval) vs close / leave? It clutters Needs Attention if real tickets start flowing.
8. Push: one authorized test push still pending (kill switch OFF). When to do it relative to Discover?
