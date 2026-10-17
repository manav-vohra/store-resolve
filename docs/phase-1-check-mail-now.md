# Phase 1 — Controlled real-mail discovery (Check Mail Now)

Date: 2026-10-17
Status: Implemented + corrective patch (independent review findings), deployed 2026-10-17. No real import, no sends.

## Corrective patch (independent review, 2026-10-17)

1. **Graph pagination security.** `MicrosoftGraphProvider.request()` now validates
   every URL before attaching the bearer token (`validateGraphRequestUrl()`):
   relative paths must be `/me/…`, absolute URLs must be
   `https://graph.microsoft.com/v1.0/me/…` (no port override, no userinfo, no
   fragment, no traversal, no encoded traversal). External / malformed /
   out-of-scope URLs are rejected with `MS_GRAPH_INVALID_URL`. Bearer tokens
   are never sent to unvalidated URLs and never logged.
2. **Inbox-only preview.** `listInboxPreviewPage()` now uses
   `/me/mailFolders/inbox/messages`, never `/me/messages`. Continuation pages
   remain Inbox-scoped: callers pass an opaque `pageToken` ($skiptoken), the
   server rebuilds the Inbox URL with the same window filter. The API returns
   `nextPageToken` only, never a full Graph URL. `?nextLink=` is rejected with
   400. Sent Items / other folders can never become imports.
3. **Delta filtering.** Graph message delta only supports an initial
   `receivedDateTime ge` (or `gt`) filter, not a combined `ge + lt` expression.
   `listInboxDeltaPage()` now sends only `ge`. The upper bound is enforced
   locally in `runMailboxDiscovery()` (already did via `inWindow`).
4. **Metadata-only delta.** Delta pages are metadata-only. `runMailboxDiscovery()`
   now calls `fetchFullMessage(id)` for each in-window message before intake.
   `DeltaDiscoveryClient` requires `fetchFullMessage()`; the live client uses
   `provider.getMessage()`. Metadata-only records are never interpreted as
   complete complaint bodies.
5. **Preview pagination revalidation.** `importSelectedPreviewMessages()`
   previously checked only the first 100 preview records. It now walks all
   preview pages (up to 10 pages / 500 messages, early exit once all wanted
   ids are found). If the window still has more pages after 10 and some wanted
   ids were not found, it refuses with `WINDOW_TOO_LARGE_FOR_REVALIDATION`
   instead of incorrectly rejecting valid selections.

Tests: `tests/phase1-corrective.test.ts` (10 tests) covers all five fixes.
Full suite: 205 passed, 1 skipped.

## Goal

Recover Dad's missing recent complaints (Oct 10–17, 2026) with a preview-first,
selection-only flow. No broad ingestion, no cursor advance on preview, no sends.

## What was wrong (fixed in this phase)

1. **Delta requested bodies.** `listInboxDeltaPage()` selected `body`,
   `bodyPreview`, `internetMessageHeaders` on the initial delta. Fixed: initial
   delta is metadata-only (id, conversationId, internetMessageId, receivedDateTime,
   subject, from, sender, toRecipients, ccRecipients, parentFolderId). Full bodies
   are fetched per-message only when that message needs processing, via `getMessage()`.

2. **Delta had no date filter.** Initial delta had no `receivedDateTime` filter,
   so the first run could process substantially more historical mail than intended.
   Fixed: `listInboxDeltaPage(link?, { receivedStart, receivedEnd })` sends
   `$filter` on the initial call, and `runMailboxDiscovery()` enforces the window
   client-side (Graph may ignore `$filter` on delta). Out-of-window mail is skipped,
   not ingested, not a failure.

3. **Cursor advanced past failures.** `runMailboxDiscovery()` persisted the new
   `deltaLink` even when individual messages in the page failed. Failed messages
   would never be re-discovered. Fixed: if any message fails, the cursor is NOT
   advanced. The previous cursor is kept, `lastErrorCode = DISCOVERY_PARTIAL_FAILURES`,
   `failedIds` are returned explicitly, and the next run re-processes from before
   the failed batch (successes dedupe via `DUPLICATE`). Same rule for max-pages
   exhaustion without a deltaLink.

4. **Deterministic gate ignored unusual forwards/replies.** `extractComplaint()`
   returned `isComplaint=false` for a forward/reply whose outer text was short
   ("FYI", "See below") and whose complaint lived in the quoted original, and
   `ingestEmailMessage()` then returned `IGNORED`, losing the message.
   Fixed:
   - `extractComplaint()` now reports `isForwardOrReply` and `hasCaseIdSignal`,
     and a forward/reply carrying a case ID is a complaint candidate
     (`FORWARD_REPLY_WITH_CASE_ID`).
   - Preview classification flags Dunkin mail, forwards/replies, and case-id
     signals with no preview match as `NEEDS_REVIEW`, never `NOT_COMPLAINT`.
   - `ingestEmailMessage()` accepts `allowOwnerSelectedReview=true` (preview
     import only). When the deterministic gate says NOT a complaint but the
     owner explicitly selected the message, it routes to `REVIEW_REQUIRED`
     with reason `OWNER_SELECTED_NON_CANDIDATE` instead of `IGNORED`.
     Automatic discovery does NOT set this flag; it keeps `IGNORED`.

5. **Disabled model transport masqueraded as ABSTAINED.** When
   `mail_interpretation_enabled=true` but the only transport is
   `DisabledModelTransport`, the pipeline recorded an `ABSTAINED` disagreement,
   implying a model ran. Fixed: `MODEL_TRANSPORT_DISABLED` now returns
   `skipped=true` with no disagreements and no interpretation record. The
   pipeline proceeds deterministically. `ABSTAINED` is reserved for a live
   transport that ran but returned no usable output. Interpretation stays OFF
   by default; no live transport exists.

## New surfaces

### `GET /api/admin/email/preview?start=&end=&top=&nextLink=`

- Owner-only, `FAMILY_PILOT` only.
- Read-only: Graph `/me/messages` (NOT delta), metadata + `bodyPreview` only,
  no full bodies, no headers, no D1 writes, no cursor, no sends.
- Window: ISO dates, start < end, max 31 days. Default Oct 10–17, 2026.
- Returns `items[]` with `maskedId`, `senderDisplay` (Dunkin shown fully,
  others masked), `subject` (truncated 200), `detectedCaseId`,
  `detectedStoreNumber`, `isForwardOrReply`, `classification`,
  `classificationReason`, `linkedComplaintId`, `alreadyIngested`.
- `graphMessageId` is in the payload for import; UI displays only `maskedId`.
- Deduplication uses batched D1 queries (no N+1): existing Graph ids, existing
  case ids (lowercase), conversation links to existing complaints.

### `POST /api/admin/email/preview-import`

- Owner-only. Requires `FAMILY_PILOT`, external OFF, ack OFF. Ingestion flag
  NOT required (explicit selection, not broad sweep).
- Body: `{ messageIds: string[1..50], receivedStart?, receivedEnd?, ingestionMode?: 'LIVE'|'BACKFILL' }`. Default `BACKFILL` for past windows.
- Each id is re-validated against a fresh preview listing for the same window.
  Ids outside the window fail with `NOT_IN_PREVIEW_WINDOW`, not imported.
- Each message is fetched fully via `getMessage()` then ingested via
  `ingestEmailMessage()` with `allowOwnerSelectedReview=true`. Per-message
  results are returned (`status`, `complaintId?`, `errorCode?`). Failures do
  not block other messages. No delta cursor is touched.
- Audited as `PREVIEW_IMPORT`.

### UI: `Check Mail Now` in Pilot Controls

- Component: `src/components/CheckMailNow.tsx`, rendered in Pilot Controls
  (OWNER/ADMIN only).
- Date inputs (default Oct 10–17, 2026), Preview button, table with checkboxes,
  masked IDs, classification labels, `Fwd/Re` and `linked` markers, Load more
  (Graph `nextLink`), Import selected with `window.confirm()` showing the count
  and the no-sends guarantee, and per-message import results.
- Import uses BACKFILL for past windows so overdue deadlines are not applied
  to recovered mail.

### `POST /api/admin/email/discover` (legacy path, kept)

- Now accepts optional `{ receivedStart, receivedEnd }` and forwards them to
  `runMailboxDiscovery()`. When omitted, it uses the stored cursor (incremental).
- Result now includes `failedIds[]` and `reason` (`DISCOVERY_PARTIAL_FAILURES`
  when the cursor was not advanced).
- This path is NOT the recommended way to recover missing mail. Use Check Mail
  Now preview + selective import.

## Safety invariants

- Preview never writes, never advances cursor, never sends.
- Import is explicit selection only. No auto-import from preview.
- No model calls. Interpretation OFF. No live transport.
- No ack: ack switch stays OFF. No SMS: external stays OFF.
- No full-state loads on preview/import. Preview uses one Graph page + three
  batched D1 queries. Import fetches one full message per selected id.
- No PII beyond what the owner already sees: preview shows subject + sender
  display (owner view only). Customer PII in bodies is never shown in preview;
  full bodies are fetched only for explicitly selected imports.

## What still needs approval

- Production deploy of this code (not done).
- Running preview against real prod mail (read-only, but uses the real Graph token).
- Importing any real message (creates complaints / review items).
- Phase 2 (manual ack), Phase 3 (scoped scheduled job), Phase 4 (push test).

## Tests

- `tests/mail-preview.test.ts` (new): masking, window validation, classification
  (candidate / needs-review / already-ingested / follow-up), forward/reply
  detection, disabled transport = skipped.
- `tests/mail-discovery.test.ts` (updated): failures do NOT advance cursor,
  date window enforced client-side.
- `tests/mail-interpretation.test.ts` (updated): live transport failure =
  ABSTAINED, disabled transport = skipped, no interpretation record on disabled.
- Full suite: 195 passed, 1 skipped. `tsc -b` clean, eslint clean, both vite
  builds clean.
