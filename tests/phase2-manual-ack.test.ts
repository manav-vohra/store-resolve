import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { readFile } from 'node:fs/promises'
import { DatabaseSync, type SQLInputValue, type StatementSync } from 'node:sqlite'
import { fileURLToPath } from 'node:url'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import { acknowledgeComplaint, reclaimStaleAck } from '../worker/ingestion'
import type { EmailProvider } from '../worker/email-provider'
import { EmailProviderError } from '../worker/email-provider'

function baseConfig(ackEnabled: boolean) {
  return {
    mode: 'FAMILY_PILOT',
    externalNotificationsEnabled: false,
    emailIngestionEnabled: false,
    emailAckEnabled: ackEnabled,
    managerAckDeadlineMinutes: 30,
    managerResolutionTargetHours: 24,
    escalationIntervalMinutes: 60,
    signalWireReconcileAfterMinutes: 10,
    emailLookbackDays: 30,
  } as never
}

function fakeProvider(opts: { sendImpl?: () => Promise<string | undefined> } = {}): EmailProvider {
  return {
    ready: true,
    verifyConnection: async () => {},
    listMessageIds: async () => [],
    getMessage: async () => { throw new Error('not implemented') },
    sendAcknowledgment: opts.sendImpl ?? (async () => 'prov-1'),
  }
}

describe('Phase 2 manual ack', () => {
  let sqlite: DatabaseSync
  let db: D1Database

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
        if (value instanceof Uint8Array) {
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
    ]) {
      const path = fileURLToPath(new URL(`../drizzle/${name}`, import.meta.url))
      await db.exec(await readFile(path, 'utf8'))
    }
  })

  afterEach(() => sqlite.close())

  async function seedComplaint() {
    const now = new Date().toISOString()
    await db.prepare(`INSERT INTO complaints (id, external_case_id, store_id, source, subject, complaint_text, category, severity, status, acknowledgement_status, routing_reason, routing_confidence, received_at, ack_deadline, created_at, updated_at) VALUES ('c1','CCC1','store-1','MICROSOFT_GRAPH','Test','Test complaint','SERVICE','LOW','ROUTING_REVIEW','DISABLED','test','HIGH',?,?,?,?)`).bind(now, now, now, now).run()
    await db.prepare(`INSERT INTO gmail_messages (gmail_message_id, complaint_id, gmail_thread_id, sender, recipients, subject, internal_date, is_follow_up, acknowledgment_status, processing_status, first_seen_at, updated_at) VALUES ('gm1','c1','th1','customerservice@dunkinbrands.com','ops@example.invalid','DBI Case # CCC1',?,0,'DISABLED','PROCESSED',?,?)`).bind(now, now, now).run()
  }

  it('manual send works when global ack switch is OFF (explicit owner action)', async () => {
    await seedComplaint()
    const result = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false), {
      manual: true,
      mailboxIdentityVerified: true,
    })
    expect(result).toBe('SENT')
  })

  it('manual send refuses without mailbox verification', async () => {
    await seedComplaint()
    const result = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false), {
      manual: true,
      mailboxIdentityVerified: false,
    })
    expect(result).toBe('MAILBOX_NOT_VERIFIED')
  })

  it('automatic path still requires global ack switch', async () => {
    await seedComplaint()
    const result = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false))
    expect(result).toBe('DISABLED')
  })

  it('FAILED is retryable via manual path, terminal via automatic path', async () => {
    await seedComplaint()
    const failing = fakeProvider({
      sendImpl: async () => { throw new EmailProviderError('MS_GRAPH_API_ERROR', 'bad request', 400) },
    })
    const first = await acknowledgeComplaint(db, failing, 'c1', baseConfig(true))
    expect(first).toBe('FAILED')
    const autoRetry = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(true))
    expect(autoRetry).toBe('ALREADY_HANDLED')
    const manualRetry = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false), {
      manual: true,
      allowRetryFromFailed: true,
      mailboxIdentityVerified: true,
    })
    expect(manualRetry).toBe('SENT')
  })

  it('ambiguous send (5xx) stays IN_FLIGHT, does not mark FAILED, does not auto-retry', async () => {
    await seedComplaint()
    const ambiguous = fakeProvider({
      sendImpl: async () => { throw new EmailProviderError('MS_GRAPH_API_ERROR', 'server error', 500) },
    })
    const result = await acknowledgeComplaint(db, ambiguous, 'c1', baseConfig(true))
    expect(result).toBe('AMBIGUOUS')
    const row = await db.prepare(`SELECT status,last_error_code FROM email_acknowledgments WHERE complaint_id='c1'`).first<{ status: string; last_error_code: string }>()
    expect(row?.status).toBe('IN_FLIGHT')
    expect(row?.last_error_code).toContain('AMBIGUOUS')
    const second = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(true))
    expect(second).toBe('ALREADY_HANDLED')
  })

  it('stale IN_FLIGHT with sentConfirmed=false stays in reconciliation (no auto-reset)', async () => {
    await seedComplaint()
    const ambiguous = fakeProvider({
      sendImpl: async () => { throw new EmailProviderError('MS_GRAPH_API_ERROR', 'server error', 500) },
    })
    await acknowledgeComplaint(db, ambiguous, 'c1', baseConfig(true))
    const old = new Date(Date.now() - 20 * 60_000).toISOString()
    await db.prepare(`UPDATE email_acknowledgments SET updated_at=? WHERE complaint_id='c1'`).bind(old).run()
    const result = await reclaimStaleAck(db, 'c1', { checkedSentItems: true, sentConfirmed: false })
    expect(result).toBe('NEEDS_FURTHER_RECONCILIATION')
    const row = await db.prepare(`SELECT status FROM email_acknowledgments WHERE complaint_id='c1'`).first<{ status: string }>()
    expect(row?.status).toBe('IN_FLIGHT')
  })

  it('stale IN_FLIGHT forceReset with note resets to PENDING (audited override)', async () => {
    await seedComplaint()
    const ambiguous = fakeProvider({
      sendImpl: async () => { throw new EmailProviderError('MS_GRAPH_API_ERROR', 'server error', 500) },
    })
    await acknowledgeComplaint(db, ambiguous, 'c1', baseConfig(true))
    const old = new Date(Date.now() - 20 * 60_000).toISOString()
    await db.prepare(`UPDATE email_acknowledgments SET updated_at=? WHERE complaint_id='c1'`).bind(old).run()
    const toPending = await reclaimStaleAck(db, 'c1', {
      checkedSentItems: true,
      sentConfirmed: false,
      forceReset: true,
      reconciliationNote: 'Checked Sent Items, thread, and waited 24h — no reply found. Conclusive.',
    })
    expect(toPending).toBe('RECLAIMED_TO_PENDING')
    const retry = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false), {
      manual: true,
      allowRetryFromFailed: true,
      mailboxIdentityVerified: true,
    })
    expect(retry).toBe('SENT')
  })

  it('ack refuses wrong recipient (forward from family member, not Dunkin)', async () => {
    const now = new Date().toISOString()
    await db.prepare(`INSERT INTO complaints (id, external_case_id, store_id, source, subject, complaint_text, category, severity, status, acknowledgement_status, routing_reason, routing_confidence, received_at, ack_deadline, created_at, updated_at) VALUES ('c3','CCC3','store-1','MICROSOFT_GRAPH','Test','Test','SERVICE','LOW','ROUTING_REVIEW','DISABLED','test','HIGH',?,?,?,?)`).bind(now, now, now, now).run()
    await db.prepare(`INSERT INTO gmail_messages (gmail_message_id, complaint_id, gmail_thread_id, sender, recipients, subject, internal_date, is_follow_up, acknowledgment_status, processing_status, first_seen_at, updated_at) VALUES ('gm3','c3','th3','Dad <ajvohra1@yahoo.com>','ops@example.invalid','Fwd: DBI Case # CCC3',?,0,'DISABLED','PROCESSED',?,?)`).bind(now, now, now).run()
    const result = await acknowledgeComplaint(db, fakeProvider(), 'c3', baseConfig(false), {
      manual: true,
      mailboxIdentityVerified: true,
    })
    expect(result).toBe('WRONG_RECIPIENT')
  })

  it('reclaim refuses without Sent Items check', async () => {
    await seedComplaint()
    const result = await reclaimStaleAck(db, 'c1', { checkedSentItems: false, sentConfirmed: false })
    expect(result).toBe('NEEDS_SENT_ITEMS_CHECK')
  })

  it('second SENT attempt is ALREADY_HANDLED (no double-send)', async () => {
    await seedComplaint()
    await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false), {
      manual: true,
      mailboxIdentityVerified: true,
    })
    const second = await acknowledgeComplaint(db, fakeProvider(), 'c1', baseConfig(false), {
      manual: true,
      mailboxIdentityVerified: true,
    })
    expect(second).toBe('ALREADY_HANDLED')
  })

  it('NO_SOURCE when complaint has no source Outlook message', async () => {
    const now = new Date().toISOString()
    await db.prepare(`INSERT INTO complaints (id, external_case_id, store_id, source, subject, complaint_text, category, severity, status, acknowledgement_status, routing_reason, routing_confidence, received_at, ack_deadline, created_at, updated_at) VALUES ('c2','CCC2','store-1','MICROSOFT_GRAPH','Test','Test','SERVICE','LOW','ROUTING_REVIEW','DISABLED','test','HIGH',?,?,?,?)`).bind(now, now, now, now).run()
    const result = await acknowledgeComplaint(db, fakeProvider(), 'c2', baseConfig(false), {
      manual: true,
      mailboxIdentityVerified: true,
    })
    expect(result).toBe('NO_SOURCE')
  })
})
