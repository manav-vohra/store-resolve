# StoreResolve — PWA & Multi-Channel Notifications: Engineering Audit

Date: 2026-10-11. Auditor: Muse. Authority: actual repo state (`main` @ `f3ccf99`,
deployed as Worker `0a6019bb`) + live production D1. No code changed.

## Verified production state (checked 2026-10-11, not assumed)

- `dad`: phone `+19086251648`, `sms_enabled=1`, `complaint_notifications_enabled=1`, role `OWNER`.
- `SR-2026-0002` (`TEST-DRILL-001`): status `MANAGER_NOTIFIED`, assigned to `dad`,
  `ack_deadline` **2026-10-02T07:48Z — 9 days overdue**. The staged drill case is
  stale; its deadlines no longer reflect a fresh drill walkthrough.
- `notifications` table: 4 DELIVERED, 7 FAILED, 1 PENDING, 9 SUPPRESSED.
  The 1 PENDING is the drill SMS alert that can never deliver under the SignalWire trial.

## 0. Stack facts that constrain everything

- **Frontend:** React 19 + Vite, single `src/App.tsx` (1930 lines), **state-based views**
  (`view`/`selectedId` useState) — `react-router-dom` is installed but **unused**. No
  URL-addressable routes today. API layer (`src/lib/api.ts`) uses `credentials:
  'same-origin'`.
- **Backend:** Hono on Cloudflare Workers, D1 via Drizzle. Assets served from
  `./dist` with `run_worker_first = ["/api/*"]`.
- **Auth:** Cloudflare Access only. `worker/auth.ts` maps the
  `Cf-Access-Authenticated-User-Email` header to the `users` table. There is no
  app-level session, token, or password. `DEV_AUTH_USER_ID` escape hatch exists
  for local dev.
- **Cron:** intentionally disabled in `wrangler.toml` (`[triggers]` commented out;
  "legacy tick loads unbounded application state"). `scheduled()` handler exists
  in `worker/index.ts:1070` and calls `runScheduledOperations`.
- **PWA:** none. `index.html` is bare (no manifest link, no
  `apple-mobile-web-app-capable`, no icons, no theme beyond `#f8faf9`). No service
  worker, no webmanifest, no push code anywhere.

## 1. Current state: what exists vs. what is missing

### Exists and reusable

| Capability | Location | Notes |
|---|---|---|
| Durable notification outbox | `notifications` D1 table (0001) | PENDING → SENT/DELIVERED/FAILED/SUPPRESSED + `failure_reason`, `provider_message_id`, timestamps. This is the outbox — keep it. |
| Dispatcher with suppression | `dispatchEligible()` in `worker/index.ts:775` | Applies per-recipient gates (notifications enabled, pilot mode, recipient ready, provider ready) then sends. Channel-aware extension point. |
| Provider abstraction | `NotificationProvider` in `worker/providers.ts:3` | `ready` + `send(notification, destination)`. `SignalWireSmsProvider` implements it; `MockNotificationProvider` for tests. **This is the seam** for Push/Email adapters. |
| Delivery receipts | `notification_callbacks` table + `POST /api/signalwire/status` (`index.ts:114`, `worker/callbacks.ts`) | Provider acceptance vs. delivery already distinguished for SMS. |
| Stale-status reconciliation | `reconcileStaleSignalWire` in `worker/operations.ts` | Polls provider for `SENT`-but-unconfirmed messages. Pattern reusable per channel. |
| SLA / escalation engine | `processDeadlines` (`src/lib/workflow.ts:323`), `recordEscalations` (`worker/operations.ts`), `escalation_events` table | Ack/resolution overdue flags + escalation records. Runs inside the operations tick. |
| Ops tick with locking | `runScheduledOperations` (`worker/operations.ts:99`) | D1-backed `background_job_locks` + `background_job_runs` audit. Just needs cron re-enabled. |
| Graph mail with send scope | `worker/microsoft-graph.ts`, `GRAPH_SCOPES` includes `Mail.Send` | Only `sendAcknowledgment` (reply) is implemented today; general send is a small addition on the same auth. |
| Inbox-only ingestion | Graph delta query scoped to `/me/mailFolders/inbox` (`microsoft-graph.ts:568`) | Operational mail sent from the same mailbox lands in Sent Items → **cannot be re-ingested as complaints**. Safe. |
| Multi-org columns | `organizations` table; `organization_id` on `users`/`stores` | Single org (`org-1`) in practice; the column is the commercial seam. |
| Test harness | `/api/admin/test-notifications`, rate-limit table, Pilot Controls UI | Pattern to copy for push testing. |

### Exists but is a second, divergent path (converge, don't duplicate)

- `worker/notifications.ts` — Phase-C `NotificationService` with **log-only**
  transport, still wired into `worker/ingestion.ts:809` as the intake hook.
  The D1 outbox is the real delivery system. The intake hook should enqueue into
  the outbox (one line change in intent) rather than living as a parallel
  "notification system."

### Missing (greenfield)

- PWA shell: manifest, service worker, icons, iOS meta tags.
- Web Push: VAPID keys, `push_subscriptions` storage, subscription endpoints,
  payload encryption, `WebPushProvider`.
- `channel` on the `Notification` type is `'SMS' | 'IN_APP'` but `IN_APP` is
  **never used** — no notification center, no unread count, no `CASE_VIEWED`
  tracking. The "notification panel" in the UI is a per-complaint people list,
  not a notification center.
- Per-user channel preferences (only `sms_enabled` / `complaint_notifications_enabled` exist).
- Per-channel delivery attempts (one `notifications` row = one channel today;
  no child table for multi-channel fan-out of one logical notification).
- Deep links: no URL routes at all (state-based views).
- Operational email send (`/me/sendMail`).

## 2. Reuse vs. refactor — file by file

**Reuse as-is:**
- `worker/providers.ts` — keep `NotificationProvider`; add `WebPushProvider`,
  `GraphEmailProvider` beside `SignalWireSmsProvider`. Keep SignalWire untouched.
- `notifications` table + `dispatchEligible()` — extend, don't replace. Add
  `'PUSH' | 'EMAIL'` to the channel domain; make the dispatcher iterate the
  recipient's enabled channels per policy.
- `notification_callbacks`, `reconcileStaleSignalWire` pattern — copy for push
  (push has no delivery receipts; treat provider-accepted POST as SENT and rely
  on subscription-expiry cleanup instead).
- `worker/operations.ts` tick + locking — re-enable via cron; scope it (below).
- `worker/microsoft-graph.ts` — add `sendMail(to, subject, body)` next to
  `sendAcknowledgment`; same token flow. **Name the new outbound class
  carefully:** `worker/email-provider.ts` already exports an `EmailProvider`
  *ingestion* interface — use `OperationalEmailProvider` or similar to avoid collision.
- `escalation_events`, `complaint_events` — add `CASE_VIEWED` event type for
  authenticated-open tracking; escalation rows already drive "who owes what."

**Change:**
- `worker/ingestion.ts:809` — route the intake notification hook into the D1
  outbox instead of the log-only `NotificationService` (retire the parallel path).
- `src/App.tsx` navigation — add **hash routing** (`#/complaints/SR-2026-0002`).
  Cheapest deep-link-compatible option; no react-router adoption needed, works
  behind Access, works from `notificationclick`.
- `src/lib/types.ts` — extend `Notification.channel`; add `PushSubscription`,
  `NotificationPreference`, `NotificationDelivery` types.
- `wrangler.toml` — re-add `[triggers]` cron, but only after scoping the tick
  (see §6 M6).

**New:**
- `public/manifest.webmanifest`, icons, iOS meta tags in `index.html`.
- `public/sw.js` (or `src/sw.ts` built to `dist/`) — push + notificationclick.
  **Must be served from the origin root** (`/sw.js` scope) and is behind
  Cloudflare Access like everything else (see §5).
- Migration: `push_subscriptions`, `notification_deliveries`,
  `notification_preferences` (all with `organization_id`).
- Endpoints: `GET /api/push/public-key`, `POST /api/push/subscribe`,
  `POST /api/push/unsubscribe`, `GET /api/notifications` (center),
  `POST /api/notifications/:id/read`.
- `worker/webpush.ts` — VAPID signing (SubtleCrypto ECDSA) + `aes128gcm`
  payload encryption. No Node `web-push` package (Workers-compatible
  implementation required; ~150 lines, well-trodden).

**Do not touch:** the SignalWire provider, the Access auth model, the complaint
state machine, the review-queue/ingestion pipeline.

## 3. Recommended architecture

```
Domain event (complaint created, ack overdue, resolution submitted…)
  → Notification policy (per-user prefs + event→channel matrix + fatigue rules)
  → Recipient selection (assigned manager, owner, backup)
  → notifications outbox (one logical row per recipient×event)
  → notification_deliveries (one row per channel attempt: PUSH/EMAIL/SMS)
  → Channel adapters (WebPushProvider / OperationalEmailProvider / SignalWireSmsProvider)
  → Delivery tracking (provider acceptance; push expiry cleanup; SMS receipts)
  → complaint_events: CASE_VIEWED (authenticated open), ACKNOWLEDGED, …
  → SLA tick: missed ack → reminder → owner escalation (new outbox rows)
```

Design decisions:

- **One outbox, child delivery rows.** `notifications` stays the logical
  "who must be told about what." New `notification_deliveries` records each
  channel attempt with its own status/provider/timestamps. This is what lets the
  UI distinguish *attempted → accepted by provider → viewed → acknowledged*.
- **Push is primary for immediate alerts; email is fallback; SMS is
  escalation/opt-in.** Matrix (default, user-overridable):
  - New complaint assigned → PUSH immediately. If no push subscription or push
    fails → EMAIL within ~5 min. SMS only if user opted in or severity critical.
  - Ack-deadline reminder → PUSH; still unacked at deadline → EMAIL + SMS to
    manager, PUSH to owner.
  - Resolution submitted → PUSH to owner (approval needed).
  - Never send all three channels for the same event unless escalation demands it.
- **Fatigue rules in policy, not in each adapter:** one PUSH per
  (recipient, complaint, event-type); reminders with backoff; suppress EMAIL if
  the PUSH was followed by CASE_VIEWED within 15 min.
- **Kill switches per channel** in `settings`
  (`push_enabled`, `email_notifications_enabled`, existing
  `external_notifications_enabled` for SMS) + per-user prefs. SMS keeps its
  existing FAMILY_PILOT gating.
- **Dispatch timing without cron for the hot path:** enqueue in the request that
  creates the event, then `ctx.waitUntil(dispatchNew(...))` — push POSTs are
  fast and this makes alerts feel instant. Cron remains for SLA/escalation sweeps
  and reconciliation only.
- **Cost:** per push = 1 D1 read (subscriptions) + N fetches. D1 free tier is
  5M reads/day; this design is orders of magnitude under it.

## 4. User experience

**Install (once):** Owner/manager opens StoreResolve on their phone while signed
in via Cloudflare Access → banner/button "Install StoreResolve" (iOS: Share →
Add to Home Screen; the app detects `display-mode: standalone`) → app launches
chromeless with the StoreResolve icon.

**Enable notifications:** In-app Settings → "Enable notifications" →
  `Notification.requestPermission()` → `pushManager.subscribe()` with the VAPID
  public key → `POST /api/push/subscribe` (Access-authenticated, so the
  subscription is bound to the logged-in user). Show per-channel toggles
  (Push / Email / SMS) with plain-language descriptions.

**Receive:** Real phone notification — StoreResolve name + icon, title like
  "New complaint needs acknowledgment," body like "Florence · Drive-through wait
  · TEST-DRILL-001" (no customer PII on the lock screen). Sound/vibration follow
  device settings.

**Act:** Tap → opens the PWA to `#/complaints/SR-2026-0002` → Access session
  resumes (or Access login if expired) → complaint detail opens → an
  authenticated open writes `CASE_VIEWED`. The existing workflow UI
  (acknowledge → investigate → findings → resolve → owner close) is unchanged.

**Ongoing:** In-app notification center (bell + unread count) lists alerts with
  read state; tapping marks read + opens the case. Outstanding-task count on the
  dashboard derives from unacknowledged assignments, not from notification rows.

## 5. Technical constraints

- **iOS push requirements:** iOS 16.4+, installed to Home Screen (not in Safari),
  explicit permission. Works with the app closed (via APNs). **No badge counts
  on iOS** (`setAppBadge` unsupported) — unread count lives in-app only.
  No custom sounds — default sound/vibration per device settings.
- **Cloudflare Access × PWA:** the SW script, manifest, and API are all behind
  Access — fine, since install/subscribe happen in an authenticated session.
  Two things to verify on a real device (do not assume): (a) the Home Screen web
  app shares the Access session cookie (iOS generally shares cookies between
  Safari and web apps, but confirm on Dad's phone); (b) cold start from a
  notification tap after session expiry lands on the Access login and then
  returns to the deep link — configure Access session lifetime sensibly and make
  the app recover gracefully.
- **Web Push in Workers:** no Node `web-push` lib; implement VAPID JWT
  (ES256 via SubtleCrypto) + `aes128gcm` encryption manually. Standard recipe,
  no new infrastructure.
- **Push has no delivery receipts.** Treat HTTP 201/200 from the push endpoint
  as SENT; treat 404/410 as expired → delete the subscription (standard).
  "Delivered" for push therefore means *accepted by the push service*, and the
  audit trail must say so honestly. True read-tracking comes from `CASE_VIEWED`.
- **Privacy:** lock-screen title/body must never contain customer names,
  complaint text, or PII. Case IDs and store names only.
- **VAPID keys:** generate once (`npx web-push generate-vapid-keys` locally),
  private key → `wrangler secret put VAPID_PRIVATE_KEY`; public key served from
  `/api/push/public-key`. One keypair is fine for the pilot; per-organization
  keys later if needed.
- **The `loadState` problem:** bootstrap and the ops tick load the *entire DB*
  into memory per call. Fine for 6 stores; it is the documented reason cron is
  off. Before multi-org scale, scope dispatch to targeted SQL (pending
  notifications + their recipients), not full-state loads.

## 6. Implementation plan

| # | Milestone | Work | Effort | Acceptance criteria |
|---|---|---|---|---|
| M1 | PWA shell | manifest, icons (192/512 + maskable, apple-touch-icon), iOS meta tags, `display: standalone`, SW registered with push + notificationclick handlers | 0.5–1 d | Installable to iPhone Home Screen; launches chromeless; SW active |
| M2 | Push plumbing | VAPID keys; `push_subscriptions` migration; subscribe/unsubscribe/public-key endpoints; `WebPushProvider`; `'PUSH'` channel in outbox + dispatcher | 1–2 d | Test push from Pilot Controls arrives on a subscribed iPhone; 410 cleanup works |
| M3 | Deep links + center | Hash routing (`#/complaints/:id`); `CASE_VIEWED` event; `GET /api/notifications` + bell/unread UI; mark-read on open | 1–2 d | Tap push → drill case opens authenticated; unread count correct |
| M4 | Channel policy | `notification_preferences`; event→channel matrix; fatigue/dedupe rules; per-channel kill switches; converge ingestion hook into outbox | 2 d | Matrix test: new complaint → push only; unacked at deadline → push+email, then SMS escalation; no triple-sends |
| M5 | Operational email | `sendMail` on Graph client (`OperationalEmailProvider`); fallback path in dispatcher | 0.5 d | Fallback email received; inbox ingestion unaffected (verified: ingestion reads inbox only) |
| M6 | Cron + escalation wiring | Re-add scoped cron (5 min); tick runs deadlines → escalations → dispatch → reconcile (email polling stays manual until scoped) | 1–2 d | Staged overdue case produces reminder → owner escalation via enabled channels, all in `escalation_events` + outbox |

Total: roughly **1.5–2.5 weeks** for one engineer (Codex), sequentially testable.
Explicit non-goals for this phase: inbound SMS commands, customer-facing
messaging, 10DLC (still tracked separately for SMS), multi-org UI.

## 7. First pilot — push to Dad's iPhone (minimum slice)

1. M1 + M2 only, plus hash deep-link for `#/complaints/:id` (half of M3).
2. Dad: iOS 16.4+, Add to Home Screen while signed in, Settings → Enable
   notifications, grant permission.
3. From Pilot Controls: "Send test push" (mirrors the SMS test button, same
   allowlist + rate limit) → Dad's phone buzzes → tap → drill case opens.
4. Then enqueue a real drill alert for SR-2026-0002 via the outbox on the PUSH
   channel and walk the drill.
5. **Pre-step:** SR-2026-0002's `ack_deadline` is 9 days past. Either reset its
   drill deadlines or stage a fresh drill case so the walkthrough (and any
   escalation demo) behaves like a live complaint.

## 8. Commercial readiness (decide now, build later)

- Put `organization_id` on `push_subscriptions`, `notification_deliveries`,
  and `notification_preferences` now. No multi-org UI, no tenant switching —
  just the column and a `WHERE organization_id = ?` habit.
- One VAPID keypair today; document that per-organization keypairs are the
  future option (separate push identity per franchise).
- Scope the dispatcher off `loadState` before adding org #2 (targeted SQL for
  pending notifications + recipients).
- Keep channel policy data-driven (`settings` JSON), not hardcoded, so a future
  franchise can have different escalation chains without code changes.
- Rate limits per channel per recipient (the SMS test rate-limit table is the
  template).
- Do not build: org management screens, cross-org analytics, white-labeling,
  or a second notification pipeline. The outbox + adapters design already
  supports all of it later.
