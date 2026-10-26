/**
 * Dad-only new-complaint push notifications.
 *
 * When a new complaint is created in FAMILY_PILOT, Dad gets a push on his
 * iPhone automatically. This is the primary alert channel for the pilot —
 * SMS and email stay off unless explicitly approved.
 *
 * Design notes:
 * - Dad-only: only user_id='dad' is ever targeted. No other owners/managers.
 * - Never breaks ingestion: all errors are caught; a push failure never
 *   fails complaint creation.
 * - Audited: every attempt is recorded in the notifications outbox table
 *   (channel='PUSH') with SENT/FAILED/SKIPPED status.
 * - Idempotent: callers pass the complaint id; a complaint already
 *   push-notified (SENT record exists) is skipped.
 */
import { WebPushProvider } from './providers'
import { buildDraftReadyPushPayload, buildNewComplaintPushPayload } from './webpush'
import type { D1Database } from './d1'

interface DadPushEnv {
  VAPID_PUBLIC_KEY?: string
  VAPID_PRIVATE_KEY?: string
  PUBLIC_BASE_URL?: string
}

const DAD_USER_ID = 'dad'

function webPush(env: DadPushEnv): WebPushProvider {
  return new WebPushProvider({
    publicKey: env.VAPID_PUBLIC_KEY,
    privateKey: env.VAPID_PRIVATE_KEY,
    subject: env.PUBLIC_BASE_URL ?? 'https://store-resolve.moogzie.workers.dev',
  })
}

export async function notifyDadNewComplaint(
  db: D1Database,
  env: DadPushEnv,
  complaint: {
    id: string
    externalCaseId: string
    category?: string
    storeNumber?: string
    ingestionMode?: 'LIVE' | 'BACKFILL' | 'TEST'
  },
): Promise<{ sent: boolean; reason?: string }> {
  const now = new Date().toISOString()
  try {
    // Gate 0: LIVE complaints only. Historical BACKFILL imports must never
    // trigger a "new complaint" push — Dad already knows about old mail.
    if (complaint.ingestionMode !== 'LIVE')
      return { sent: false, reason: 'NOT_LIVE_INGESTION' }

    // Gate 0b: dad_push_enabled kill switch, enforced HERE (not just by
    // callers) so no path — including POST /api/complaints — can bypass it.
    // Default OFF: deploying code never authorizes real pushes by itself.
    const pushSwitch = await db
      .prepare(`SELECT value FROM settings WHERE key='dad_push_enabled'`)
      .first<{ value: string }>()
    if (pushSwitch?.value !== 'true') return { sent: false, reason: 'PUSH_DISABLED' }

    // Gate 1: FAMILY_PILOT only. The settings key is 'notification_mode'
    // (see worker/d1.ts loadState); 'mode' does not exist.
    const mode = await db
      .prepare(`SELECT value FROM settings WHERE key='notification_mode'`)
      .first<{ value: string }>()
    if (mode?.value !== 'FAMILY_PILOT') return { sent: false, reason: 'NOT_FAMILY_PILOT' }

    // Gate 2: idempotency — skip if we already sent for this complaint.
    const existing = await db
      .prepare(
        `SELECT id FROM notifications WHERE complaint_id=? AND recipient_user_id=? AND channel='PUSH' AND event_type='COMPLAINT_RECEIVED' AND status='SENT'`,
      )
      .bind(complaint.id, DAD_USER_ID)
      .first<{ id: string }>()
    if (existing) return { sent: false, reason: 'ALREADY_SENT' }

    // Gate 3: Dad must have an active push subscription.
    const sub = await db
      .prepare(
        `SELECT endpoint, p256dh_key, auth_key FROM push_subscriptions WHERE user_id=? AND revoked_at IS NULL ORDER BY last_used_at DESC LIMIT 1`,
      )
      .bind(DAD_USER_ID)
      .first<{ endpoint: string; p256dh_key: string; auth_key: string }>()
    if (!sub) {
      await record(db, complaint.id, 'SKIPPED', 'NO_PUSH_SUBSCRIPTION', now)
      return { sent: false, reason: 'NO_PUSH_SUBSCRIPTION' }
    }

    // Gate 4: Web Push must be configured.
    const push = webPush(env)
    if (!push.ready) {
      await record(db, complaint.id, 'SKIPPED', 'WEBPUSH_NOT_READY', now)
      return { sent: false, reason: 'WEBPUSH_NOT_READY' }
    }

    try {
      const result = await push.sendToSubscription(
        { endpoint: sub.endpoint, p256dhKey: sub.p256dh_key, authKey: sub.auth_key },
        buildNewComplaintPushPayload(complaint),
      )
      if (result.expired) {
        await db
          .prepare(`UPDATE push_subscriptions SET revoked_at=? WHERE endpoint=? AND revoked_at IS NULL`)
          .bind(now, sub.endpoint)
          .run()
        await record(db, complaint.id, 'FAILED', 'SUBSCRIPTION_EXPIRED', now)
        return { sent: false, reason: 'SUBSCRIPTION_EXPIRED' }
      }
      await db
        .prepare(`UPDATE push_subscriptions SET last_used_at=? WHERE endpoint=?`)
        .bind(now, sub.endpoint)
        .run()
      await record(db, complaint.id, 'SENT', null, now)
      return { sent: true }
    } catch (error) {
      const msg = (error as Error).message
      // Distinguish definite failures from uncertain outcomes.
      // Uncertain (timeout/5xx/network after submit): the push may have
      // been accepted. Record as NEEDS_RECONCILIATION — never auto-retry,
      // a human verifies before any resend to avoid double-notifying Dad.
      const uncertain = /timeout|timed out|5\d\d|network|econn|socket|abort/i.test(msg)
      const status = uncertain ? 'NEEDS_RECONCILIATION' : 'FAILED'
      await record(db, complaint.id, status, msg.slice(0, 500), now)
      return { sent: false, reason: uncertain ? 'SEND_UNCERTAIN' : 'SEND_FAILED' }
    }
  } catch (error) {
    // Never break the caller. Record the failure if we can, then swallow.
    try {
      await record(db, complaint.id, 'FAILED', `DAD_PUSH_ERROR: ${(error as Error).message.slice(0, 300)}`, now)
    } catch {
      // Even the audit write failed; nothing more we can do.
    }
    return { sent: false, reason: 'INTERNAL_ERROR' }
  }
}

async function record(
  db: D1Database,
  complaintId: string,
  status: 'SENT' | 'FAILED' | 'SKIPPED' | 'NEEDS_RECONCILIATION',
  reason: string | null,
  now: string,
): Promise<void> {
  await db
    .prepare(
      `INSERT INTO notifications(id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,failed_at,failure_reason) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    )
    .bind(
      crypto.randomUUID(),
      complaintId,
      'COMPLAINT_RECEIVED',
      DAD_USER_ID,
      'PUSH',
      'New Dunkin complaint — tap to open in StoreResolve',
      status,
      'WEBPUSH',
      null,
      now,
      status === 'SENT' ? now : null,
      status === 'FAILED' ? now : null,
      reason,
    )
    .run()
}

/**
 * Retry failed Dad pushes with durable, idempotent recovery.
 *
 * - Finds FAILED push notifications for Dad (COMPLAINT_RECEIVED, channel PUSH).
 * - Atomically reserves each one (status FAILED -> RETRYING) so concurrent
 *   workers never send duplicates. The reservation is a single conditional
 *   UPDATE; only the worker whose UPDATE matched proceeds.
 * - Caps retries at 3 attempts per notification (tracked via failure_reason
 *   prefix or a retry_count — here via re-queue with attempt marker).
 * - On success: SENT. On failure: back to FAILED with updated reason.
 * - Never throws; returns per-notification outcomes for the caller to log.
 *
 * Note on delivery semantics: a provider acceptance (Apple returns 2xx)
 * is recorded as SENT, but that is not proof the iPhone displayed it.
 * The outbox is the source of truth for attempts, not for display.
 */
/**
 * Retry failed Dad pushes with durable, idempotent recovery.
 *
 * - Finds FAILED push notifications for Dad (COMPLAINT_RECEIVED, channel PUSH).
 * - Skips non-LIVE complaints (historical mail never gets retried pushes).
 * - Atomic reservation: UPDATE ... WHERE status='FAILED' and require
 *   exactly one affected row (via result.meta.changes) before sending.
 *   Concurrent workers cannot double-send.
 * - Cumulative retry limit: 3 attempts max per notification, tracked via
 *   RETRY_n prefix in failure_reason. Exponential backoff: 5min, 30min, 2h.
 * - Uncertain provider outcomes (timeout/5xx after submit) go to
 *   NEEDS_RECONCILIATION, never auto-resent — a human verifies before retry.
 * - Never throws; returns per-notification outcomes.
 *
 * Delivery semantics: provider acceptance (2xx) is recorded as SENT, but
 * that is not proof the iPhone displayed it. The outbox tracks attempts;
 * display is confirmed only by the recipient.
 */
const MAX_RETRY_ATTEMPTS = 3
/** Exported for tests: must be exactly [300_000, 1_800_000, 7_200_000]. */
export const RETRY_BACKOFF_MS = [5 * 60_000, 30 * 60_000, 2 * 60 * 60_000]

export async function retryFailedDadPushes(
  db: D1Database,
  env: DadPushEnv,
  maxRetries = 5,
): Promise<{ retried: number; sent: number; failed: number; deferred: number }> {
  const out = { retried: 0, sent: 0, failed: 0, deferred: 0 }
  const now = new Date().toISOString()
  const nowMs = Date.now()

  // Gate: same pilot + kill switch as the send path.
  const mode = await db
    .prepare(`SELECT value FROM settings WHERE key='notification_mode'`)
    .first<{ value: string }>()
  if (mode?.value !== 'FAMILY_PILOT') return out
  const pushEnabled = await db
    .prepare(`SELECT value FROM settings WHERE key='dad_push_enabled'`)
    .first<{ value: string }>()
  if (pushEnabled?.value !== 'true') return out

  const push = new WebPushProvider({
    publicKey: env.VAPID_PUBLIC_KEY,
    privateKey: env.VAPID_PRIVATE_KEY,
    subject: env.PUBLIC_BASE_URL ?? 'https://store-resolve.moogzie.workers.dev',
  })
  if (!push.ready) return out

  const failed = await db
    .prepare(
      `SELECT n.id, n.complaint_id, n.failure_reason, n.failed_at, c.ingestion_mode
       FROM notifications n LEFT JOIN complaints c ON c.id = n.complaint_id
       WHERE n.recipient_user_id=? AND n.channel='PUSH' AND n.event_type='COMPLAINT_RECEIVED'
       AND n.status='FAILED' ORDER BY n.created_at LIMIT ?`,
    )
    .bind(DAD_USER_ID, maxRetries)
    .all<{
      id: string
      complaint_id: string
      failure_reason: string | null
      failed_at: string | null
      ingestion_mode: string | null
    }>()
  for (const row of failed.results) {
    // Never retry non-LIVE (historical BACKFILL/TEST).
    if (row.ingestion_mode !== 'LIVE') continue
    // Cumulative retry limit + backoff.
    const attemptMatch = /^RETRY_(\d+)/.exec(row.failure_reason ?? '')
    const attempts = attemptMatch ? parseInt(attemptMatch[1], 10) : 1
    if (attempts >= MAX_RETRY_ATTEMPTS) continue
    const backoff = RETRY_BACKOFF_MS[Math.min(attempts - 1, RETRY_BACKOFF_MS.length - 1)]
    const failedAtMs = row.failed_at ? Date.parse(row.failed_at) : 0
    if (Number.isFinite(failedAtMs) && nowMs - failedAtMs < backoff) {
      out.deferred += 1
      continue
    }
    // Atomic reservation: exactly one row must flip, or another worker won.
    const reserved = await db
      .prepare(`UPDATE notifications SET status='RETRYING' WHERE id=? AND status='FAILED'`)
      .bind(row.id)
      .run()
    const changes = (reserved as unknown as { meta?: { changes?: number } }).meta?.changes ?? 0
    if (!reserved.success || changes !== 1) continue
    out.retried += 1
    const nextAttempt = attempts + 1
    try {
      const sub = await db
        .prepare(
          `SELECT endpoint, p256dh_key, auth_key FROM push_subscriptions WHERE user_id=? AND revoked_at IS NULL ORDER BY last_used_at DESC LIMIT 1`,
        )
        .bind(DAD_USER_ID)
        .first<{ endpoint: string; p256dh_key: string; auth_key: string }>()
      if (!sub) {
        await db
          .prepare(`UPDATE notifications SET status='FAILED', failure_reason=?, failed_at=? WHERE id=?`)
          .bind(`RETRY_${nextAttempt}:NO_SUBSCRIPTION`, now, row.id)
          .run()
        out.failed += 1
        continue
      }
      const complaint = await db
        .prepare(`SELECT external_case_id, category, store_id FROM complaints WHERE id=?`)
        .bind(row.complaint_id)
        .first<{ external_case_id: string; category: string; store_id: string | null }>()
      const store = complaint?.store_id
        ? await db
            .prepare(`SELECT dunkin_store_number FROM stores WHERE id=?`)
            .bind(complaint.store_id)
            .first<{ dunkin_store_number: string }>()
        : null
      const { buildNewComplaintPushPayload } = await import('./webpush')
      const result = await push.sendToSubscription(
        { endpoint: sub.endpoint, p256dhKey: sub.p256dh_key, authKey: sub.auth_key },
        buildNewComplaintPushPayload({
          id: row.complaint_id,
          externalCaseId: complaint?.external_case_id ?? row.complaint_id,
          category: complaint?.category,
          storeNumber: store?.dunkin_store_number,
        }),
      )
      if (result.expired) {
        await db
          .prepare(`UPDATE push_subscriptions SET revoked_at=? WHERE endpoint=? AND revoked_at IS NULL`)
          .bind(now, sub.endpoint)
          .run()
        await db
          .prepare(`UPDATE notifications SET status='FAILED', failure_reason=?, failed_at=? WHERE id=?`)
          .bind(`RETRY_${nextAttempt}:SUBSCRIPTION_EXPIRED`, now, row.id)
          .run()
        out.failed += 1
      } else {
        await db
          .prepare(`UPDATE notifications SET status='SENT', sent_at=?, failure_reason=NULL WHERE id=?`)
          .bind(now, row.id)
          .run()
        out.sent += 1
      }
    } catch (error) {
      const msg = (error as Error).message
      // Uncertain outcomes (timeout/5xx after submit): do NOT auto-resend.
      // Mark for human reconciliation — resending could double-notify.
      const uncertain = /timeout|timed out|5\d\d|network|econn|socket/i.test(msg)
      if (uncertain) {
        await db
          .prepare(
            `UPDATE notifications SET status='NEEDS_RECONCILIATION', failure_reason=?, failed_at=? WHERE id=?`,
          )
          .bind(`RETRY_${nextAttempt}:UNCERTAIN:${msg.slice(0, 200)}`, now, row.id)
          .run()
      } else {
        await db
          .prepare(`UPDATE notifications SET status='FAILED', failure_reason=?, failed_at=? WHERE id=?`)
          .bind(`RETRY_${nextAttempt}:${msg.slice(0, 200)}`, now, row.id)
          .run()
      }
      out.failed += 1
    }
  }
  return out
}

/**
 * Dad-only draft-ready push notification.
 *
 * Fires once per response draft, when a reply is prepared for Dad's review.
 * Edits never re-notify: dedup is by (event_type='DRAFT_READY', draft id
 * stored in provider_message_id). A SENT record for the draft blocks any
 * second push, even if the draft is edited, resubmitted, or re-approved.
 *
 * Same gates as new-complaint pushes: dad_push_enabled kill switch,
 * FAMILY_PILOT mode, active subscription, Web Push configured. Never
 * breaks the caller — all failures are audited and swallowed.
 */
export async function notifyDadDraftReady(
  db: D1Database,
  env: DadPushEnv,
  draft: { id: string; complaintId: string; externalCaseId: string },
): Promise<{ sent: boolean; reason?: string }> {
  const now = new Date().toISOString()
  try {
    const pushSwitch = await db
      .prepare(`SELECT value FROM settings WHERE key='dad_push_enabled'`)
      .first<{ value: string }>()
    if (pushSwitch?.value !== 'true') return { sent: false, reason: 'PUSH_DISABLED' }

    const mode = await db
      .prepare(`SELECT value FROM settings WHERE key='notification_mode'`)
      .first<{ value: string }>()
    if (mode?.value !== 'FAMILY_PILOT') return { sent: false, reason: 'NOT_FAMILY_PILOT' }

    // Idempotency: one push per draft, ever. Edits do not re-notify.
    const existing = await db
      .prepare(
        `SELECT id FROM notifications WHERE event_type='DRAFT_READY' AND provider_message_id=? AND channel='PUSH' AND status='SENT'`,
      )
      .bind(draft.id)
      .first<{ id: string }>()
    if (existing) return { sent: false, reason: 'ALREADY_SENT' }

    const sub = await db
      .prepare(
        `SELECT endpoint, p256dh_key, auth_key FROM push_subscriptions WHERE user_id=? AND revoked_at IS NULL ORDER BY last_used_at DESC LIMIT 1`,
      )
      .bind(DAD_USER_ID)
      .first<{ endpoint: string; p256dh_key: string; auth_key: string }>()
    if (!sub) {
      await recordDraft(db, draft, 'SKIPPED', 'NO_PUSH_SUBSCRIPTION', now)
      return { sent: false, reason: 'NO_PUSH_SUBSCRIPTION' }
    }

    const push = webPush(env)
    if (!push.ready) {
      await recordDraft(db, draft, 'SKIPPED', 'WEBPUSH_NOT_READY', now)
      return { sent: false, reason: 'WEBPUSH_NOT_READY' }
    }

    try {
      const result = await push.sendToSubscription(
        { endpoint: sub.endpoint, p256dhKey: sub.p256dh_key, authKey: sub.auth_key },
        buildDraftReadyPushPayload(draft),
      )
      if (result.expired) {
        await db
          .prepare(`UPDATE push_subscriptions SET revoked_at=? WHERE endpoint=? AND revoked_at IS NULL`)
          .bind(now, sub.endpoint)
          .run()
        await recordDraft(db, draft, 'FAILED', 'SUBSCRIPTION_EXPIRED', now)
        return { sent: false, reason: 'SUBSCRIPTION_EXPIRED' }
      }
      await db
        .prepare(`UPDATE push_subscriptions SET last_used_at=? WHERE endpoint=?`)
        .bind(now, sub.endpoint)
        .run()
      await recordDraft(db, draft, 'SENT', null, now)
      return { sent: true }
    } catch (error) {
      const msg = (error as Error).message
      const uncertain = /timeout|timed out|5\d\d|network|econn|socket|abort/i.test(msg)
      const status = uncertain ? 'NEEDS_RECONCILIATION' : 'FAILED'
      await recordDraft(db, draft, status, msg.slice(0, 500), now)
      return { sent: false, reason: uncertain ? 'SEND_UNCERTAIN' : 'SEND_FAILED' }
    }
  } catch (error) {
    try {
      await recordDraft(db, draft, 'FAILED', `DAD_PUSH_ERROR: ${(error as Error).message.slice(0, 300)}`, now)
    } catch {
      // Even the audit write failed; nothing more we can do.
    }
    return { sent: false, reason: 'INTERNAL_ERROR' }
  }
}

async function recordDraft(
  db: D1Database,
  draft: { id: string; complaintId: string; externalCaseId: string },
  status: 'SENT' | 'FAILED' | 'SKIPPED' | 'NEEDS_RECONCILIATION',
  reason: string | null,
  now: string,
): Promise<void> {
  await db
    .prepare(
      `INSERT INTO notifications(id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,failed_at,failure_reason) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    )
    .bind(
      crypto.randomUUID(),
      draft.complaintId,
      'DRAFT_READY',
      DAD_USER_ID,
      'PUSH',
      `Reply ready for review — Case ${draft.externalCaseId}. Tap to review and approve.`,
      status,
      'WEBPUSH',
      draft.id,
      now,
      status === 'SENT' ? now : null,
      status === 'FAILED' ? now : null,
      reason,
    )
    .run()
}
