# Stage A Implementation Brief — Installable PWA + Working Web Push

This is the implementation brief for **Stage A** of the approved plan in
`docs/pwa-notifications-audit.md` (§6 milestones M1–M2 + hash deep links from M3).
Read the audit doc first; it is the architecture authority. Build only what this
brief specifies. Do not start Stage B or C.

## Mission

Make StoreResolve installable as a PWA and able to deliver a Web Push
notification to a subscribed device, with a tap opening the correct complaint
behind the existing Cloudflare Access auth. **All real-recipient push sends stay
disabled behind an explicit kill switch until the operator approves.**

## Non-negotiable safety gates (stop, do not work around)

1. **No real pushes without approval.** Implement `POST /api/admin/push/test`
   gated by (a) admin-only, (b) the existing test-recipient allowlist pattern,
   (c) 5-minute per-recipient rate limit (reuse `test_notification_rate_limits`
   semantics), AND (d) a new settings kill switch `push_test_sends_enabled`
   defaulting to `'false'`. When off → HTTP 409 with a clear message. **Leave it
   off. Never flip it.**
2. **Do not touch the drill case.** No changes to SR-2026-0002, its deadlines,
   assignments, or any complaint. No overdue notifications triggered.
3. **Do not change rollout settings** (`NOTIFICATION_MODE`,
   `external_notifications_enabled`, or any `settings` row you didn't create).
4. **Migrations:** inspect `HEAD` and the *deployed* D1 schema first
   (`wrangler d1 migrations list store-resolve --remote`). Migration
   `drizzle/0007_phase_c_operations.sql` exists in-repo but is **not approved for
   production — do not apply it**. Author one new additive migration for this
   stage only; apply it remotely **before** deploying code that needs it.
5. **Do not enable cron** or any scheduled external sends.
6. **Do not weaken API authorization.** Push endpoints use the existing
   Access-authenticated user (`c.get('user')`); subscriptions bind to that user.
7. **Do not modify** the SignalWire provider, the ingestion pipeline, the review
   queue, the complaint state machine, or `worker/notifications.ts` usage.
8. Real-device testing (Dad's iPhone) requires the operator's explicit approval —
   stop before that step and report readiness.

## Build list

### A1. PWA shell
- `public/manifest.webmanifest`: name "StoreResolve", `display: standalone`,
  `start_url: "/"`, theme/background colors matching the app (`#f8faf9`),
  icons 192px + 512px (maskable). Generate simple branded icons (no external
  assets; SVG→PNG via a script or hand-authored PNG).
- `index.html`: manifest link, `apple-mobile-web-app-capable`,
  `apple-mobile-web-app-status-bar-style`, `apple-mobile-web-app-title`,
  `apple-touch-icon` (180px), `theme-color`.
- Register the service worker in `src/main.tsx` (guard with
  `'serviceWorker' in navigator`).

### A2. Service worker (`public/sw.js`, served from origin root)
- `push` event: parse payload `{title, body, url, badge}`; `showNotification`
  with privacy-safe content (see below); feature-detected badge update:
  `navigator.setAppBadge(count)` in try/catch (iOS 16.4+ supports it; do not
  depend on it — it is a supplement, not the accountability mechanism).
- `notificationclick`: `clients.openWindow(url)` or focus existing client, then
  navigate. Close the notification.
- Keep it dependency-free and small.

### A3. Hash deep links
- Frontend: support `#/complaints/:id` — on load and on `hashchange`, parse the
  hash and open the complaint detail view (`navigate('detail', id)`).
  Validate the id against loaded state; unknown id → dashboard + toast.
- Push payloads use `url: "/#/complaints/SR-2026-0002"` style links.

### A4. Push subscriptions (D1)
- New migration `drizzle/0008_push_subscriptions.sql`:
  `push_subscriptions(id TEXT PK, organization_id TEXT NOT NULL, user_id TEXT NOT
  NULL REFERENCES users(id), endpoint TEXT NOT NULL UNIQUE, p256dh_key TEXT NOT
  NULL, auth_key TEXT NOT NULL, device_label TEXT, created_at TEXT NOT NULL,
  last_used_at TEXT, revoked_at TEXT)`.
- Endpoints (all under existing auth, same-origin):
  - `GET /api/push/public-key` → `{ publicKey }` from `VAPID_PUBLIC_KEY` var.
  - `POST /api/push/subscribe` `{endpoint, keys:{p256dh, auth}, deviceLabel?}`:
    validate endpoint is `https:` URL, keys are valid base64url; upsert on
    `(user_id, endpoint)` scoped to the authenticated user; set `revoked_at=NULL`.
  - `POST /api/push/unsubscribe` `{endpoint}`: set `revoked_at`, only for the
    authenticated user's own rows.

### A5. Web Push delivery adapter
- **First, evaluate** a standards-compliant, Cloudflare Workers-compatible Web
  Push library (VAPID signing + `aes128gcm` content encryption). Prefer a tested
  dependency with a clear security/maintenance profile over hand-rolled crypto.
  Document the decision (chosen lib + why, or why hand-rolled with SubtleCrypto
  was necessary) in `docs/stage-a-notes.md`.
- Implement `WebPushProvider` in `worker/providers.ts` (or a new
  `worker/webpush.ts`) following the existing `NotificationProvider` shape where
  it fits; the push path needs endpoint + keys, so a
  `sendToSubscription(subscription, {title, body, url, badge})` method is
  acceptable — do **not** reshape the SignalWire interface to accommodate push.
- VAPID: public key → add `VAPID_PUBLIC_KEY = "BEaf_RWjGAb19TP5EoFUiEU5Sautp58BH5RLzMSbZwApds4Ob9sudVD4WkOUhQD6pMugRNms93Sm9Zknt2cjCAk"`
  to `wrangler.toml [vars]`; private key is already set as the
  `VAPID_PRIVATE_KEY` secret (do not touch). JWT `sub` claim: use
  `https://store-resolve.moogzie.workers.dev`.
- Delivery semantics: HTTP 201 from the push service = accepted (record SENT
  with providerMessageId unset/endpoint ref); HTTP 404/410 = expired → mark
  subscription revoked. Never claim device delivery — record "accepted by push
  service."

### A6. Permission + settings UI
- Settings view: "Notifications" section. States: unsupported / not installed
  (iOS: show "Add to Home Screen" instructions when not in standalone
  display-mode) / permission prompt button / subscribed (show device label +
  disable button).
- Flow: `Notification.requestPermission()` → `pushManager.subscribe({userVisibleOnly:
  true, applicationServerKey})` → `POST /api/push/subscribe`. Surface errors
  plainly (denied permission, no SW, etc.).

### A7. Pilot Controls test push
- Mirror the existing SMS test UI: recipient select (same allowlist), confirm
  checkbox, "Send test push" → `POST /api/admin/push/test`. Show resulting
  delivery state. The endpoint enforces the §(1) gates above. Test payload is
  privacy-safe and clearly labeled TEST.

### A8. Privacy rule (hard)
Notification title/body must never contain customer names, complaint text, or
PII. Allowed: app name, store name, case/external ID, action needed
(e.g. "StoreResolve: New complaint needs acknowledgment — Florence ·
TEST-DRILL-001").

## Engineering expectations

- pnpm (`pnpm-lock.yaml` is canonical). Scripts: `pnpm typecheck`, `pnpm lint`,
  `pnpm test`, `pnpm build`.
- Tests (vitest): subscription upsert/revoke scoping (user A cannot touch user
  B's rows), public-key endpoint, gated test endpoint (kill-switch off → 409;
  rate limit; non-admin → 403), deep-link hash parsing, push payload privacy
  (no PII fields in title/body), 404/410 → revoked.
- Commits: one clear commit per chunk (A1…A8-ish), messages like
  `stage-a: pwa shell (manifest, icons, sw registration)`.
- Deploy with `wrangler deploy` (CLOUDFLARE_API_TOKEN is in env); confirm the
  deployed version ID and smoke-test `GET /api/health` + `/api/push/public-key`.
- Write `docs/stage-a-notes.md`: what was built, the Web Push library decision,
  and the **real-device setup checklist** (iOS version, install steps, permission
  flow, what the operator must do on Dad's iPhone, and the exact approval
  needed to enable the test send).

## Done when

`pnpm typecheck && pnpm lint && pnpm test` green; deployed to production;
`/api/push/public-key` returns the VAPID key; Pilot Controls shows the (disabled)
test-push UI; `docs/stage-a-notes.md` exists with the device checklist; no
real notification sent; no complaint or setting modified; migration 0007
untouched. Report: commits, test results, deployment version, blockers.
