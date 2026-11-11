// Pending-draft reminder (2026-11-11).
//
// Manav: a reply draft that sits unreviewed after one push should not go
// silent — the app must remind Dad. Two parts:
//  - notifyDadDraftReminder: one capped re-nudge per draft, audited as
//    DRAFT_REMINDER, same PENDING-before-send pattern as notifyDadDraftReady.
//  - remindStaleDrafts: the tick sweep that finds drafts due for that one
//    reminder (exactly one DRAFT_READY SENT, at least 7 days old).
//
// Same test setup as dad-push-draft-ready.test.ts: the REAL SQLite schema
// (migrations 0001-0006 + 0008 + 0009 — 0009 is required because 0006's
// CHECK constraints reject DRAFT rows) and a MOCKED push provider, so no
// test ever touches a real push service.
//
// Proves:
// 1. Reminder: PENDING row exists before the push call, then settles SENT.
// 2. One reminder per draft, ever — a second call is ALREADY_SENT.
// 3. Failure settles the PENDING row in place.
// 4. Sweep: a due draft (one SENT >= 7 days old) is reminded exactly once.
// 5. Sweep: a draft with TWO SENTs (initial + manual re-notify, the
//    CCC11703791 case) never gets an automatic third push.
// 6. Sweep: fresh, already-reminded, and never-notified drafts are skipped.
// 7. Sweep: non-DRAFT actions are not considered; push disabled skips.
import { readFile } from 'node:fs/promises'
import { DatabaseSync, type SQLInputValue, type StatementSync } from 'node:sqlite'
import { fileURLToPath } from 'node:url'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import { notifyDadDraftReminder, remindStaleDrafts } from '../worker/dad-push'

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
      if (
        value === null ||
        typeof value === 'string' ||
        typeof value === 'number' ||
        typeof value === 'bigint'
      ) {
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

// Fixed "now" for the sweep: 2026-11-11 12:00 UTC.
const NOW_MS = Date.parse('2026-11-11T12:00:00Z')
const isoDaysAgo = (days: number) => new Date(NOW_MS - days * 24 * 60 * 60 * 1000).toISOString()

const draftA = { id: 'draft-a', complaintId: 'SR-TEST-001', externalCaseId: 'CCC00000001' }
const draftB = { id: 'draft-b', complaintId: 'SR-TEST-002', externalCaseId: 'CCC00000002' }
const draftC = { id: 'draft-c', complaintId: 'SR-TEST-003', externalCaseId: 'CCC00000003' }

type NotificationRow = {
  id: string
  event_type: string
  status: string
  provider_message_id: string
  sent_at: string | null
  failed_at: string | null
  failure_reason: string | null
}

describe('pending-draft reminder', () => {
  let sqlite: DatabaseSync
  let db: D1Database

  const reminderRows = async () =>
    (
      await db
        .prepare(
          `SELECT id,event_type,status,provider_message_id,sent_at,failed_at,failure_reason FROM notifications WHERE event_type='DRAFT_REMINDER' ORDER BY created_at`,
        )
        .all<NotificationRow>()
    ).results

  const seedDraft = async (
    id: string,
    complaintId: string,
    state = 'DRAFT',
    createdAt = '2026-10-26T03:14:48Z',
  ) => {
    await db
      .prepare(
        `INSERT INTO response_actions(id,complaint_id,source_message_id,response_kind,policy_version,policy_state,content_kind,content_ref,template_version,model_name,model_snapshot,prompt_version,approved_by,approved_at,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
      )
      .bind(
        id,
        complaintId,
        null,
        'RESOLUTION',
        'response-policy.v1',
        state,
        'draft-body',
        'test body',
        null,
        null,
        null,
        null,
        null,
        null,
        createdAt,
        createdAt,
      )
      .run()
  }

  const seedNotification = async (
    eventType: string,
    draftId: string,
    complaintId: string,
    status: 'SENT' | 'FAILED' | 'SKIPPED',
    sentAt: string,
    providerKey?: string,
  ) => {
    await db
      .prepare(
        `INSERT INTO notifications(id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,failed_at,failure_reason) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)`,
      )
      .bind(
        crypto.randomUUID(),
        complaintId,
        eventType,
        'dad',
        'PUSH',
        `${eventType} — Case ${complaintId}`,
        status,
        'WEBPUSH',
        providerKey ?? `${draftId}:${crypto.randomUUID()}`,
        sentAt,
        status === 'SENT' ? sentAt : null,
        status === 'FAILED' ? sentAt : null,
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
      '0009_response_actions_workflow_states.sql',
    ]) {
      const path = fileURLToPath(new URL(`../drizzle/${name}`, import.meta.url))
      await db.exec(await readFile(path, 'utf8'))
    }
    await db.exec(
      `INSERT INTO organizations(id,name,created_at,updated_at) VALUES('org-test','Test Org','2026-01-01T00:00:00Z','2026-01-01T00:00:00Z')`,
    )
    await db.exec(
      `INSERT INTO users(id,organization_id,name,email,phone,role,timezone,created_at,updated_at) VALUES('dad','org-test','Dad','dad@example.test','+10000000000','OWNER','America/New_York','2026-01-01T00:00:00Z','2026-01-01T00:00:00Z')`,
    )
    for (const [complaintId, caseId] of [
      ['SR-TEST-001', 'CCC00000001'],
      ['SR-TEST-002', 'CCC00000002'],
      ['SR-TEST-003', 'CCC00000003'],
    ]) {
      await db.exec(
        `INSERT INTO complaints(id,external_case_id,subject,complaint_text,category,severity,status,routing_reason,routing_confidence,received_at,ack_deadline,created_at,updated_at) VALUES('${complaintId}','${caseId}','Test complaint','Test body','Service','LOW','MANAGER_NOTIFIED','test','HIGH','2026-10-19T00:00:00Z','2026-10-21T00:00:00Z','2026-10-19T00:00:00Z','2026-10-19T00:00:00Z')`,
      )
    }
    await db.exec(
      `INSERT INTO settings(key,value,updated_at) VALUES('dad_push_enabled','true','2026-01-01T00:00:00Z') ON CONFLICT(key) DO UPDATE SET value='true'`,
    )
    await db.exec(`UPDATE settings SET value='FAMILY_PILOT' WHERE key='notification_mode'`)
    await db.exec(
      `INSERT INTO push_subscriptions(id,organization_id,user_id,endpoint,p256dh_key,auth_key,created_at,last_used_at,revoked_at) VALUES('sub-1','org-test','dad','https://push.test/dad-1','p','a','2026-01-01T00:00:00Z',NULL,NULL)`,
    )
    pushControl.calls = 0
    pushControl.behavior = 'ok'
    pushControl.onSend = null
  })

  afterEach(() => sqlite.close())

  it('reminder: PENDING row exists before the push call, then settles to SENT', async () => {
    let pendingSeenDuringSend = false
    pushControl.onSend = async () => {
      const rows = await reminderRows()
      pendingSeenDuringSend =
        rows.length === 1 &&
        rows[0]?.status === 'PENDING' &&
        (rows[0]?.provider_message_id ?? '').startsWith(`${draftA.id}:`)
    }
    const result = await notifyDadDraftReminder(db, env, draftA)
    expect(result).toEqual({ sent: true })
    expect(pushControl.calls).toBe(1)
    expect(pendingSeenDuringSend).toBe(true)
    const rows = await reminderRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.event_type).toBe('DRAFT_REMINDER')
    expect(rows[0]?.status).toBe('SENT')
    expect(rows[0]?.sent_at).not.toBeNull()
  })

  it('one reminder per draft: a second call is ALREADY_SENT with no new row', async () => {
    expect(await notifyDadDraftReminder(db, env, draftA)).toEqual({ sent: true })
    expect(await notifyDadDraftReminder(db, env, draftA)).toEqual({
      sent: false,
      reason: 'ALREADY_SENT',
    })
    expect(pushControl.calls).toBe(1)
    expect((await reminderRows()).length).toBe(1)
  })

  it('failure handling: push throw settles the PENDING row to FAILED', async () => {
    pushControl.behavior = 'throw'
    const result = await notifyDadDraftReminder(db, env, draftA)
    expect(result).toEqual({ sent: false, reason: 'SEND_FAILED' })
    const rows = await reminderRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.status).toBe('FAILED')
    expect(rows[0]?.failed_at).not.toBeNull()
    expect(rows[0]?.failure_reason).toContain('push service rejected')
  })

  it('sweep: due draft (one SENT 8 days old) gets exactly one reminder', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(8))
    const result = await remindStaleDrafts(db, env, { nowMs: NOW_MS })
    expect(result).toEqual({ checked: 1, reminded: 1, skipped: 0 })
    expect(pushControl.calls).toBe(1)
    const rows = await reminderRows()
    expect(rows.length).toBe(1)
    expect(rows[0]?.status).toBe('SENT')
    // A second sweep the same day does not send again.
    expect(await remindStaleDrafts(db, env, { nowMs: NOW_MS })).toEqual({
      checked: 1,
      reminded: 0,
      skipped: 1,
    })
    expect(pushControl.calls).toBe(1)
  })

  it('sweep: one SENT exactly 7 days old is due (boundary inclusive)', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(7))
    const result = await remindStaleDrafts(db, env, { nowMs: NOW_MS })
    expect(result).toEqual({ checked: 1, reminded: 1, skipped: 0 })
  })

  it('sweep: two SENTs (initial + manual re-notify) never gets an automatic third push', async () => {
    // The CCC11703791 shape: legacy bare-id SENT, then a force re-notify.
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification(
      'DRAFT_READY',
      draftA.id,
      draftA.complaintId,
      'SENT',
      isoDaysAgo(16),
      draftA.id,
    )
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(14))
    const result = await remindStaleDrafts(db, env, { nowMs: NOW_MS })
    expect(result).toEqual({ checked: 1, reminded: 0, skipped: 1 })
    expect(pushControl.calls).toBe(0)
    expect((await reminderRows()).length).toBe(0)
  })

  it('sweep: fresh draft (SENT 2 days old) is skipped', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(2))
    expect(await remindStaleDrafts(db, env, { nowMs: NOW_MS })).toEqual({
      checked: 1,
      reminded: 0,
      skipped: 1,
    })
    expect(pushControl.calls).toBe(0)
  })

  it('sweep: already-reminded draft is skipped', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(8))
    await seedNotification('DRAFT_REMINDER', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(1))
    expect(await remindStaleDrafts(db, env, { nowMs: NOW_MS })).toEqual({
      checked: 1,
      reminded: 0,
      skipped: 1,
    })
    expect(pushControl.calls).toBe(0)
  })

  it('sweep: never-notified draft (zero SENT) is skipped', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    expect(await remindStaleDrafts(db, env, { nowMs: NOW_MS })).toEqual({
      checked: 1,
      reminded: 0,
      skipped: 1,
    })
    expect(pushControl.calls).toBe(0)
  })

  it('sweep: non-DRAFT actions are not considered', async () => {
    await seedDraft(draftA.id, draftA.complaintId, 'PENDING_APPROVAL')
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(8))
    expect(await remindStaleDrafts(db, env, { nowMs: NOW_MS })).toEqual({
      checked: 0,
      reminded: 0,
      skipped: 0,
    })
    expect(pushControl.calls).toBe(0)
  })

  it('sweep: push disabled means a due draft is skipped with no send', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(8))
    await db.exec(`UPDATE settings SET value='false' WHERE key='dad_push_enabled'`)
    expect(await remindStaleDrafts(db, env, { nowMs: NOW_MS })).toEqual({
      checked: 1,
      reminded: 0,
      skipped: 1,
    })
    expect(pushControl.calls).toBe(0)
  })

  it('sweep: mixed drafts report correct counts', async () => {
    await seedDraft(draftA.id, draftA.complaintId)
    await seedNotification('DRAFT_READY', draftA.id, draftA.complaintId, 'SENT', isoDaysAgo(8))
    await seedDraft(draftB.id, draftB.complaintId)
    await seedNotification('DRAFT_READY', draftB.id, draftB.complaintId, 'SENT', isoDaysAgo(2))
    await seedDraft(draftC.id, draftC.complaintId)
    await seedNotification(
      'DRAFT_READY',
      draftC.id,
      draftC.complaintId,
      'SENT',
      isoDaysAgo(16),
      draftC.id,
    )
    await seedNotification('DRAFT_READY', draftC.id, draftC.complaintId, 'SENT', isoDaysAgo(14))
    const result = await remindStaleDrafts(db, env, { nowMs: NOW_MS })
    expect(result).toEqual({ checked: 3, reminded: 1, skipped: 2 })
    expect(pushControl.calls).toBe(1)
  })
})
