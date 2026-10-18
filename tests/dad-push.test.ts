/**
 * Regression tests for Dad-only push gating (worker/dad-push.ts).
 *
 * Proves:
 * 1. The FAMILY_PILOT gate reads the real production settings key
 *    ('notification_mode'), not a nonexistent 'mode' key.
 * 2. BACKFILL imports never trigger a push (only LIVE).
 * 3. Gating fails closed (no push) when config is missing.
 */
import { describe, expect, it } from 'vitest'
import { notifyDadNewComplaint } from '../worker/dad-push'
import type { D1Database } from '../worker/d1'

// Minimal fake D1 serving canned rows.
function fakeDb(opts: {
  mode?: string
  subscription?: { endpoint: string; p256dh_key: string; auth_key: string } | null
  alreadySent?: boolean
}) {
  const state = { notified: [] as string[] }
  const db = {
    prepare(sql: string) {
      const first = async () => {
        if (sql.includes('notification_mode')) return opts.mode ? { value: opts.mode } : null
        if (sql.includes('push_subscriptions')) return opts.subscription ?? null
        if (sql.includes("channel='PUSH'")) return opts.alreadySent ? { id: 'x' } : null
        return null
      }
      const q = {
        bind() {
          return { first, run: q.run }
        },
        first,
        run: async () => {
          if (sql.includes('INSERT INTO notifications')) state.notified.push(sql)
          return { success: true }
        },
      }
      return q
    },
    _state: state,
  }
  return db as unknown as D1Database & { _state: typeof state }
}

const env = {
  VAPID_PUBLIC_KEY: 'test-public',
  VAPID_PRIVATE_KEY: 'test-private',
  PUBLIC_BASE_URL: 'https://example.test',
}

const liveComplaint = {
  id: 'SR-2026-9999',
  externalCaseId: 'CCC99999999',
  category: 'Order Accuracy',
  storeNumber: '350909',
  ingestionMode: 'LIVE' as const,
}

describe('notifyDadNewComplaint gating', () => {
  it('never pushes for BACKFILL ingestion', async () => {
    const db = fakeDb({
      mode: 'FAMILY_PILOT',
      subscription: { endpoint: 'https://push.test/e', p256dh_key: 'p', auth_key: 'a' },
    })
    const result = await notifyDadNewComplaint(db, env, {
      ...liveComplaint,
      ingestionMode: 'BACKFILL',
    })
    expect(result.sent).toBe(false)
    expect(result.reason).toBe('NOT_LIVE_INGESTION')
  })

  it('fails closed when the mode setting is absent (pins notification_mode key)', async () => {
    // Production stores the mode under 'notification_mode'. If the code
    // queried a wrong key, the fake returns null and the gate must fail
    // closed rather than open.
    const db = fakeDb({
      subscription: { endpoint: 'https://push.test/e', p256dh_key: 'p', auth_key: 'a' },
    })
    const result = await notifyDadNewComplaint(db, env, liveComplaint)
    expect(result.sent).toBe(false)
    expect(result.reason).toBe('NOT_FAMILY_PILOT')
  })

  it('fails closed when not in FAMILY_PILOT', async () => {
    const db = fakeDb({
      mode: 'MOCK',
      subscription: { endpoint: 'https://push.test/e', p256dh_key: 'p', auth_key: 'a' },
    })
    const result = await notifyDadNewComplaint(db, env, liveComplaint)
    expect(result.sent).toBe(false)
    expect(result.reason).toBe('NOT_FAMILY_PILOT')
  })

  it('skips when Dad has no active subscription, audited as SKIPPED', async () => {
    const db = fakeDb({ mode: 'FAMILY_PILOT', subscription: null })
    const result = await notifyDadNewComplaint(db, env, liveComplaint)
    expect(result.sent).toBe(false)
    expect(result.reason).toBe('NO_PUSH_SUBSCRIPTION')
    expect((db as unknown as { _state: { notified: string[] } })._state.notified.length).toBe(1)
  })

  it('is idempotent per complaint', async () => {
    const db = fakeDb({
      mode: 'FAMILY_PILOT',
      subscription: { endpoint: 'https://push.test/e', p256dh_key: 'p', auth_key: 'a' },
      alreadySent: true,
    })
    const result = await notifyDadNewComplaint(db, env, liveComplaint)
    expect(result.sent).toBe(false)
    expect(result.reason).toBe('ALREADY_SENT')
  })
})
