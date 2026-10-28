// Draft-ready push audit fix (2026-10-27).
//
// Regression: the Oct 27 force re-notify reached Apple's push service but
// left NO audit row, because recordDraft stored the bare draft id in
// provider_message_id and migration 0002 puts a UNIQUE index on that
// column — the second attempt's INSERT failed. These tests run against
// the REAL SQLite schema (migrations 0001-0006 + 0008), so the UNIQUE
// index is genuinely enforced, and against a MOCKED push provider, so
// no test ever touches a real push service.
//
// Proves, for notifyDadDraftReady:
// 1. Normal notification: the attempt row exists as PENDING *before* the
//    push call, then settles to SENT — one row, per-attempt key.
// 2. Explicit re-notification (force): a second push sends AND a second
//    audit row is written (the old code failed here with INTERNAL_ERROR).
// 3. Duplicate prevention: a confirmed SENT blocks a non-force push with
//    no new row, for both legacy bare-draft-id rows and new per-attempt rows.
// 4. Failure handling: throw / uncertain / expired all settle the PENDING
//    row in place (no second row, no untracked send).
// 5. A SKIPPED attempt never blocks a later successful send.
import { readFile } from 'node:fs/promises'
import { DatabaseSync, type SQLInputValue, type StatementSync } from 'node:sqlite'
import { fileURLToPath } from 'node:url'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import { notifyDadDraftReady } from '../worker/dad-push'

// Controllable fake push provider. The real WebPushProvider is replaced
// entirely: no keys are used, no network happens, no real push is sent.
const pushControl = vi.hoisted(() => ({
  calls: 0,
  behavior: 'ok' as 'ok' | 'expired' | 'throw' | 'throw-uncertain',
  onSend: null as null | (() => Promise<void> | void),
}))

vi.mock('../worker/providers', () => ({
  WebPushProvider: class {
    constructor(
      private readonly config: { publicKey?: string; privateKey?: string; subject?: string },
    ) {}
    get ready() {
      return Boolean(this.config.publicKey && this.config.privateKey && this.config.subject)
    }
    async sendToSubscription(sub: { endpoint: string }, _payload: unknown) {
      pushControl.calls += 1
      if (pushControl.onSend) await pushControl.onSend()
      if (pushControl.behavior === 'expired') return { accepted: false, expired: true }
      if (pushControl.behavior === 'throw') throw new Error('push service rejected')
      if (pushControl.behavior === 'throw-uncertain') throw new Error('network timeout')
      return { accepted: true, expired: false }
    }
  },
}))

class SqliteStatement implements D1PreparedStatement {
  private values: SQLInputValue[] = []
  constructor(private readonly statement: StatementSync) {}
  bind(...values: unknown[]) {
    const normalized: SQLInputValue[] = []
    for (const value of values) {
      if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'bigint') {
        normalized.push(value)
        continue
      }
      throw new TypeError('Unsupported SQLite test binding')
    }
    this.values = normalized
    return this
  }
  async all<T>(): Promise<D1Result<T>> {
    return { success: true, results: this.statement.all(...this.values) as T[] }
  }
  async first<T>(): Promise<T | null> {
    return (this.statement.get(...this.values) as T | undefined) ?? null
  }
  async run(): Promise<D1Result> {
    const result = this.statement.run(...this.values)
    return { success: true, results: [], meta: { changes: Number(result.changes) } }
  }
}

class SqliteD1 implements D1Database {
  constructor(private readonly database: DatabaseSync) {}
  prepare(sql: string) {
    return new SqliteStatement(this.database.prepare(sql))
  }
  async batch<T>(statements: D1PreparedStatement[]): Promise<D1Result<T>[]> {
    this.database.exec('BEGIN')
    try {
      const results = []
      for (const statement of statements) results.push((await statement.run()) as D1Result<T>)
      this.database.exec('COMMIT')
      return results
    } catch (error) {
      this.database.exec('ROLLBACK')
      throw error
    }
  }
  async exec(sql: string): Promise<D1Result> {
    this.database.exec(sql)
    return { success: true, results: [] }
  }
}

const env = {
  VAPID_PUBLIC_KEY: 'test-public',
  VAPID_PRIVATE_KEY: 'test-private',
  PUBLIC_BASE_URL: 'https://example.test',
}

const draft = { id: 'draft-1', complaintId: 'SR-TEST-001', externalCaseId: 'CCC00000001' }

type NotificationRow = {
  id: string
  status: string
  provider_message_id: string
  sent_at: string | null
  failed_at: string | null
  failure_reason: string | null
}

describe('notifyDadDraftReady audit', () => {
  let sqlite: DatabaseSync
  let db: D1Database

  const draftReadyRows = async () =>
    (
      await db
        .prepare(`SELECT id,status,provider_message_id,sent_at,failed_at,failure_reason FROM notifications WHERE event_type='DRAFT_READY' ORDER BY created_at`)
        .all<NotificationRow>()
    ).results

  const seedLegacySent = async () => {
    await db
      .prepare(
        `INSERT INTO notifications(id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,failed_at,failure_reason) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)`,
      )
      .bind(
        'legacy-row',
        draft.complaintId,
        'DRAFT_READY',
        'dad',
        'PUSH',
        'Reply ready for review — Case CCC00000001. Tap to review and approve.',
        'SENT',
        'WEBPUSH',
        draft.id,
        '2026-10-26T03:34:59.591Z',
        '2026-10-26T03:34:59.591Z',
        null,
        null,
      )
      .run()
  }

  beforeEach(async () => {
    sqlite = new DatabaseSync(':memory:')
    db = new SqliteD1(sqlite)
    for (const name of [
      '0001_initial.sql',
      '0002_persistence_and_auth.sql',
      '0003_pilot_admin_recipient.sql',
      '0004_v1_operations.sql',
      '0005_microsoft_graph_provider.sql',
      '0006_intelligent_mail_foundation.sql',
      '0008_push_subscriptions.sql',
    ]) {
      const path = fileURLToPath(new URL(`../drizzle/${name}`, import.meta.url))
      await db.exec(await readFile(path, 'utf8'))
    }
    await db.exec(`INSERT INTO organizations(id,name,created_at,updated_at) VALUES('org-test','Test Org','2026-01-01T00:00:00Z','2026-01-01T00:00:00Z')`)
    await db.exec(
      `INSERT INTO users(id,organization_id,name,email,phone,role,timezone,created_at,updated_at) VALUES('dad','org-test','Dad','dad@example.test','+10000000000','OWNER','America/New_York','2026-01-01T00:00:00Z','2026-01-01T00:00:00Z')`,
    )
    await db.exec(
      `INSERT INTO complaints(id,external_case_id,subject,complaint_text,category,severity,status,routing_reason,routing_confidence,received_at,ack_deadline,created_at,updated_at) VALUES('SR-TEST-001','CCC00000001','Test complaint','Test body','Service','LOW','MANAGER_NOTIFIED','test','HIGH','2026-10-19T00:00:00Z','2026-10-21T00:00:00Z','2026-10-19T00:00:00Z','2026-10-19T00:00:00Z')`,
    )
    await db.exec(`INSERT INTO settings(key,value,updated_at) VALUES('dad_push_enabled','true','2026-01-01T00:00:00Z') ON CONFLICT(key) DO UPDATE SET value='true'`)
    await db.exec(`UPDATE settings SET value='FAMILY_PILOT' WHERE key='notification_mode'`)
    await db.exec(
      `INSERT INTO push_subscriptions(id,organization_id,user_id,endpoint,p256dh_key,auth_key,created_at,last_used_at,revoked_at) VALUES('sub-1','org-test','dad','https://push.test/dad-1','p','a','2026-01-01T00:00:00Z',NULL,NULL)`,
    )
    pushControl.calls = 0
    pushControl.behavior = 'ok'
    pushControl.onSend = null
  })

  afterEach(() => sqlite.close())

  it('normal notification: PENDING row exists before the push call, then settles to SENT', async () => {
    let pendingSeenDuringSend = false
    pushControl.onSend = async () => {
      const rows = await draftReadyRows()
      pendingSeenDuringSend =
        rows.length === 1 &&
        rows[0]?.status === 'PENDING' &&
        (rows[0]?.provider_message_id ?? '').startsWith(`${draft.id}:`)
    }
    const result = await notifyDadDraftReady(db, env, draft)
    expect(result).toEqual({ sent: true })
    expect(pushControl.calls).toBe(1)
    expect(pendingSeenDuringSend).toBe(true)
    const rows = await draftReadyRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.status).toBe('SENT')
    expect(rows[0]?.sent_at).not.toBeNull()
    expect(rows[0]?.provider_message_id).not.toBe(draft.id)
    expect(rows[0]?.provider_message_id.startsWith(`${draft.id}:`)).toBe(true)
  })

  it('explicit re-notification (force) sends a second push and writes a second audit row', async () => {
    await seedLegacySent()
    const result = await notifyDadDraftReady(db, env, draft, { force: true })
    expect(result).toEqual({ sent: true })
    expect(pushControl.calls).toBe(1)
    const rows = await draftReadyRows()
    expect(rows.length).toBe(2)
    expect(rows.filter((r) => r.status === 'SENT').length).toBe(2)
    // Legacy row untouched; new row carries a distinct per-attempt key.
    expect(rows[0]?.provider_message_id).toBe(draft.id)
    expect(rows[1]?.provider_message_id.startsWith(`${draft.id}:`)).toBe(true)
    expect(rows[1]?.provider_message_id).not.toBe(draft.id)
  })

  it('duplicate prevention: non-force after a legacy SENT does not push and writes no new row', async () => {
    await seedLegacySent()
    const result = await notifyDadDraftReady(db, env, draft)
    expect(result).toEqual({ sent: false, reason: 'ALREADY_SENT' })
    expect(pushControl.calls).toBe(0)
    expect((await draftReadyRows()).length).toBe(1)
  })

  it('duplicate prevention: non-force after a new-format SENT also does not push', async () => {
    const first = await notifyDadDraftReady(db, env, draft)
    expect(first).toEqual({ sent: true })
    const second = await notifyDadDraftReady(db, env, draft)
    expect(second).toEqual({ sent: false, reason: 'ALREADY_SENT' })
    expect(pushControl.calls).toBe(1)
    expect((await draftReadyRows()).length).toBe(1)
  })

  it('failure handling: push throw settles the PENDING row to FAILED with the reason', async () => {
    pushControl.behavior = 'throw'
    const result = await notifyDadDraftReady(db, env, draft)
    expect(result).toEqual({ sent: false, reason: 'SEND_FAILED' })
    expect(pushControl.calls).toBe(1)
    const rows = await draftReadyRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.status).toBe('FAILED')
    expect(rows[0]?.failed_at).not.toBeNull()
    expect(rows[0]?.failure_reason).toContain('push service rejected')
  })

  it('failure handling: uncertain errors settle to NEEDS_RECONCILIATION, never auto-retried', async () => {
    pushControl.behavior = 'throw-uncertain'
    const result = await notifyDadDraftReady(db, env, draft)
    expect(result).toEqual({ sent: false, reason: 'SEND_UNCERTAIN' })
    const rows = await draftReadyRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.status).toBe('NEEDS_RECONCILIATION')
  })

  it('failure handling: expired subscription revokes and settles FAILED', async () => {
    pushControl.behavior = 'expired'
    const result = await notifyDadDraftReady(db, env, draft)
    expect(result).toEqual({ sent: false, reason: 'SUBSCRIPTION_EXPIRED' })
    const rows = await draftReadyRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.status).toBe('FAILED')
    expect(rows[0]?.failure_reason).toBe('SUBSCRIPTION_EXPIRED')
    const sub = await db
      .prepare(`SELECT revoked_at FROM push_subscriptions WHERE id='sub-1'`)
      .first<{ revoked_at: string | null }>()
    expect(sub?.revoked_at).not.toBeNull()
  })

  it('a SKIPPED attempt never blocks a later successful send', async () => {
    await db.exec(`DELETE FROM push_subscriptions WHERE id='sub-1'`)
    const skipped = await notifyDadDraftReady(db, env, draft)
    expect(skipped).toEqual({ sent: false, reason: 'NO_PUSH_SUBSCRIPTION' })
    expect(pushControl.calls).toBe(0)
    await db.exec(
      `INSERT INTO push_subscriptions(id,organization_id,user_id,endpoint,p256dh_key,auth_key,created_at,last_used_at,revoked_at) VALUES('sub-2','org-test','dad','https://push.test/dad-2','p','a','2026-01-01T00:00:00Z',NULL,NULL)`,
    )
    const sent = await notifyDadDraftReady(db, env, draft)
    expect(sent).toEqual({ sent: true })
    const rows = await draftReadyRows()
    expect(rows.length).toBe(2)
    expect(rows.filter((r) => r.status === 'SKIPPED').length).toBe(1)
    expect(rows.filter((r) => r.status === 'SENT').length).toBe(1)
  })
})
