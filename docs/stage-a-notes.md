# Stage A implementation notes — 2026-10-11

Installable PWA + working Web Push plumbing. All real-recipient sends disabled.

## What was built

**Commit 0078905 — PWA shell**
- `public/manifest.webmanifest` (standalone, theme `#117a59`, maskable icon)
- `public/icons/`: 192, 512, 512-maskable, apple-touch-icon (green/orange "SR" brand)
- `index.html`: iOS standalone meta, manifest + icon links
- `src/main.tsx`: service worker registration (failure never breaks the shell)
- `public/sw.js`: `push` handler (notification display + feature-detected
  `setAppBadge`/`clearAppBadge`), `notificationclick` (focus-or-open deep link),
  `message` handler for badge refresh from the page
- `src/App.tsx` + `src/lib/deep-links.ts`: `#/complaints/:id` deep links, parsed
  on load and `hashchange`, synced from `navigate()`, role-scoped

**Commit 1f98700 — push backend**
- `drizzle/0008_push_subscriptions.sql`: `push_subscriptions` table,
  organization-scoped, endpoint UNIQUE, revocation via `revoked_at`
- `worker/webpush.ts`: RFC 8291 (aes128gcm) + RFC 8292 (VAPID) sender on
  Workers SubtleCrypto — no new dependencies (see library evaluation)
- `worker/providers.ts`: `WebPushProvider` adapter beside `SignalWireSmsProvider`;
  the `NotificationProvider` interface is untouched
- Endpoints (all behind existing Access auth):
  - `GET /api/push/public-key`
  - `POST /api/push/subscribe` (upsert per user, 409 on cross-user endpoint claim)
  - `POST /api/push/unsubscribe` (revokes caller's rows only)
  - `POST /api/admin/push/test` (admin-only, test-recipient allowlist,
    `push_test_sends_enabled` kill switch, 5-min rate limit, 404/410 → revoke)
- `wrangler.toml`: `VAPID_PUBLIC_KEY` var (private key is a Worker secret)

**Commit eb99988 — push UI + client**
- `src/lib/push.ts`: permission/subscribe/unsubscribe/badge client helpers
- `PushNotificationSettings` card in settings: status states, iOS "Add to Home
  Screen" guidance, enable/disable, privacy note
- Pilot Controls: gated "Send one confirmed test push" section
- `tests/push-notifications.test.ts`: 14 tests

## Library evaluation (correction #4)

No maintained Workers-native Web Push library exists with meaningful adoption
(`web-push` on npm is Node-only: node:https + Node crypto). Implemented the RFC
constructions directly on the platform's audited SubtleCrypto primitives
(ECDSA P-256, ECDH, HKDF, AES-GCM, HMAC). Not custom cryptography — only the
RFC framing is hand-written.

**Runtime finding:** WebCrypto ECDSA `sign()` returns raw r||s (64 bytes), not
DER. The sender handles raw natively with a DER fallback for non-conforming
runtimes. Verified by an independent client-side decrypt round-trip test.

## Verification

- `tsc -b`: clean (also added `worker-configuration.d.ts` to
  `tsconfig.app.json` — the first test to import the worker app needed `Env`)
- `eslint`: clean
- `vitest`: 181 passed, 1 skipped (14 new push tests, incl. full
  encrypt→decrypt round-trip, 410→revoke, kill switch, scoping, privacy)
- `vite build` (app + worker): clean; manifest/sw.js/icons present in dist

## Deployment

- Migration: `0008_push_subscriptions.sql` applied to remote D1. `0007` was
  temporarily moved aside during apply and restored — it remains pending and
  unapplied, as required.
- Pushed to GitHub: `f3ccf99..eb99988`
- `wrangler deploy`: version `2110560e-b46e-49ab-9f12-715f626dd4ed`
- Smoke: production URLs return Access 302s as expected (Access intact);
  endpoint behavior covered by integration tests against the real schema.

## Safety gates — confirmed

- `push_test_sends_enabled` kill switch: off (no settings row). Test endpoint
  409s until explicitly enabled.
- No real push sent. No complaint, deadline, assignment, or rollout setting
  changed. `SR-2026-0002` untouched. Cron untouched. SignalWire, ingestion,
  review queue, workflow untouched. `worker/notifications.ts` untouched.
- Push payloads carry no customer PII (case ID/store only); enforced by
  `buildTestPushPayload` + unit test.

## Codex environment

Codex could not run in this image: `codex-code-mode-host` is missing from
`/opt/hatch-image/bin`, so every repo tool call failed (4 attempts, incl. a
fallback flag that only fixed trivial commands). Implemented directly instead.
Recorded in `~/TOOLS.md`.

## Crypto review fixes — 2026-10-11 (post-deploy audit)

An independent review found two issues in `worker/webpush.ts`; both fixed,
verified, and redeployed (version `c1a4e65e-d3ad-43bc-bf61-a92f37391c7f`).

1. **RFC 8291 salt handling (real bug).** The per-message salt was written into
   the header but never mixed into key derivation, and the HKDF steps misused
   `deriveBits` with an empty salt. `derivePushKeys()` now mirrors the RFC 8291
   §3.4 pseudocode step-for-step with explicit HMAC-SHA-256 operations:
   `PRK = HMAC-SHA-256(salt, IKM)` uses the message salt. The old round-trip
   test passed only because the test decoder duplicated the sender's wrong
   derivation.
2. **VAPID signature format.** Already handled: WebCrypto returns raw P1363
   `r||s` (64 bytes); `derToRaw()` is a fallback for non-conforming runtimes
   only. Confirmed under workerd.

**Verification added:**
- RFC 8291 Appendix A test vector: all intermediate values (`ecdh_secret`,
  `prk_key`, `ikm`, `prk`, `cek`, `nonce`) and the full 86-octet-header +
  ciphertext message reproduced byte-for-byte.
- VAPID JWT verified independently with the public key (claims + signature).
- Full push crypto run under `workerd` (real Workers runtime): 7/7 PASS.
- Full suite: 184 passed, 1 skipped. Both builds clean.

## Migration status and safety procedure

- `0007_phase_c_operations.sql` is **intentionally deferred** — not approved,
  not applied. It remains the only pending migration on remote D1.
- `0008_push_subscriptions.sql` was applied remotely on 2026-10-11 via the
  migrations system (0007 temporarily moved aside during apply, then restored).
- **Safe procedure for future migrations:** before `wrangler d1 migrations
  apply --remote`, run `wrangler d1 migrations list --remote` and confirm every
  pending migration is approved. If an unapproved migration (like 0007) is
  pending, move it out of `drizzle/` temporarily, apply, then restore it —
  never apply it. `d1 execute --file` bypasses migration bookkeeping and must
  not be used for schema changes.

## Next: real-device gate (Dad's iPhone)

1. On iPhone Safari, open the production URL and sign in via Cloudflare Access.
2. Share → Add to Home Screen → open StoreResolve from the Home Screen.
3. Settings → Push notifications → Enable → allow.
4. Verify the card shows "Enabled on this device".
5. (After explicit approval) enable the kill switch and send one test push
   from Pilot Controls to Dad; tap the notification → complaint deep link.
