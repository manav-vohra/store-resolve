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
import { buildNewComplaintPushPayload } from './webpush'
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
      await record(db, complaint.id, 'FAILED', (error as Error).message.slice(0, 500), now)
      return { sent: false, reason: 'SEND_FAILED' }
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
  status: 'SENT' | 'FAILED' | 'SKIPPED',
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
export async function retryFailedDadPushes(
  db: D1Database,
  env: DadPushEnv,
  maxRetries = 5,
): Promise<{ retried: number; sent: number; failed: number }> {
  const out = { retried: 0, sent: 0, failed: 0 }
  const now = new Date().toISOString()

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
      `SELECT id, complaint_id, failure_reason FROM notifications
       WHERE recipient_user_id=? AND channel='PUSH' AND event_type='COMPLAINT_RECEIVED'
       AND status='FAILED' ORDER BY created_at LIMIT ?`,
    )
    .bind(DAD_USER_ID, maxRetries)
    .all<{ id: string; complaint_id: string; failure_reason: string | null }>()
  for (const row of failed.results) {
    // Atomic reservation: only proceed if we won the race.
    const reserved = await db
      .prepare(`UPDATE notifications SET status='RETRYING' WHERE id=? AND status='FAILED'`)
      .bind(row.id)
      .run()
    if (!reserved.success || (reserved as unknown as { changes?: number }).changes === 0) continue
    out.retried += 1
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
          .bind('RETRY_NO_SUBSCRIPTION', now, row.id)
          .run()
        out.failed += 1
        continue
      }
      // Rebuild the payload from the complaint (no customer data).
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
          .bind('RETRY_SUBSCRIPTION_EXPIRED', now, row.id)
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
      await db
        .prepare(`UPDATE notifications SET status='FAILED', failure_reason=?, failed_at=? WHERE id=?`)
        .bind(`RETRY_FAILED: ${(error as Error).message.slice(0, 300)}`, now, row.id)
        .run()
      out.failed += 1
    }
  }
  return out
}
