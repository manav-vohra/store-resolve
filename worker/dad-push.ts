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
