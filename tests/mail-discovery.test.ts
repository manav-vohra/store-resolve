// Phase C: Graph delta discovery — no live Graph calls; fake client + fake D1.
import { describe, expect, it } from 'vitest'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import type { NormalizedEmailMessage } from '../worker/email-provider'
import {
  FakeDeltaDiscoveryClient,
  runMailboxDiscovery,
  type MailboxDiscoveryOptions,
} from '../worker/mail-discovery'
import type { GraphMessage, InboxDeltaPage } from '../worker/microsoft-graph'

const graphMessage = (id: string, subject = `Subject ${id}`): GraphMessage => ({
  id,
  conversationId: `conv-${id}`,
  receivedDateTime: '2026-09-28T10:00:00.000Z',
  subject,
  from: { emailAddress: { address: 'guest@example.com' } },
})

class FakeStatement implements D1PreparedStatement {
  values: unknown[] = []
  constructor(
    readonly sql: string,
    private readonly db: FakeDb,
  ) {}
  bind(...values: unknown[]) {
    this.values = values
    return this
  }
  async all<T>(): Promise<D1Result<T>> {
    return { results: [] as T[], success: true }
  }
  async first<T>(): Promise<T | null> {
    return this.db.first(this.sql, this.values) as T | null
  }
  async run(): Promise<D1Result> {
    return this.db.run(this.sql, this.values)
  }
}

class FakeDb implements D1Database {
  cursors = new Map<string, Record<string, unknown>>()
  seenSources = new Set<string>()
  cursorWrites = 0
  prepare(sql: string): D1PreparedStatement {
    return new FakeStatement(sql, this)
  }
  async batch<T>(statements: D1PreparedStatement[]): Promise<D1Result<T>[]> {
    const out: D1Result<T>[] = []
    for (const s of statements) out.push((await (s as FakeStatement).run()) as D1Result<T>)
    return out
  }
  async exec(): Promise<unknown> {
    return null
  }
  first(sql: string, values: unknown[]): unknown {
    if (sql.includes('FROM mail_discovery_cursors WHERE provider=?')) {
      const row = [...this.cursors.values()].find(
        (c) =>
          c.provider === values[0] && c.mailbox_key === values[1] && c.scope === values[2],
      )
      return row ? { cursor_value: row.cursor_value } : null
    }
    if (sql.includes('FROM mail_discovery_cursors WHERE id=?')) {
      const row = this.cursors.get(String(values[0]))
      return row
        ? { cursor_value: row.cursor_value, high_water_received_at: row.high_water_received_at }
        : null
    }
    return null
  }
  run(sql: string, values: unknown[]): D1Result {
    if (sql.includes('INSERT INTO mail_discovery_cursors')) {
      this.cursorWrites += 1
      const id = String(values[0])
      const prev = this.cursors.get(id)
      this.cursors.set(id, {
        id,
        provider: values[1],
        mailbox_key: values[2],
        scope: values[3],
        cursor_value: values[4] ?? prev?.cursor_value ?? null,
        high_water_received_at: values[5] ?? prev?.high_water_received_at ?? null,
        last_started_at: values[6] ?? prev?.last_started_at ?? null,
        last_succeeded_at: values[7] ?? prev?.last_succeeded_at ?? null,
        last_error_code: values[8] ?? null,
      })
      return { results: [], success: true, meta: { changes: 1 } }
    }
    if (sql.includes('INSERT OR IGNORE INTO mail_source_messages')) {
      // UNIQUE(provider, mailbox_key, provider_message_id) -> values[1..3]
      const key = `${values[1]}|${values[2]}|${values[3]}`
      if (this.seenSources.has(key)) return { results: [], success: true, meta: { changes: 0 } }
      this.seenSources.add(key)
      return { results: [], success: true, meta: { changes: 1 } }
    }
    return { results: [], success: true, meta: { changes: 1 } }
  }
}

const baseOptions = (db: FakeDb, pages: InboxDeltaPage[], extra?: Partial<MailboxDiscoveryOptions>) => {
  const ingested: NormalizedEmailMessage[] = []
  const seen = new Set<string>()
  const options: MailboxDiscoveryOptions = {
    provider: 'MICROSOFT_GRAPH',
    mailboxKey: 'production',
    ingestionEnabled: true,
    client: new FakeDeltaDiscoveryClient(pages),
    // Models the real intake path: it owns source persistence and reports
    // re-delivered messages as DUPLICATE.
    ingest: async (_d, message) => {
      if (seen.has(message.id)) return { status: 'DUPLICATE' as const }
      seen.add(message.id)
      ingested.push(message)
      return { status: 'IGNORED' as const }
    },
    ...extra,
  }
  return { db, options, ingested }
}

describe('runMailboxDiscovery', () => {
  it('refuses to run when ingestion is disabled and never touches the client', async () => {
    const db = new FakeDb()
    const client = new FakeDeltaDiscoveryClient([{ messages: [graphMessage('m1')] }])
    const result = await runMailboxDiscovery(db, {
      provider: 'MICROSOFT_GRAPH',
      mailboxKey: 'production',
      ingestionEnabled: false,
      client,
    })
    expect(result.ran).toBe(false)
    expect(result.reason).toBe('INGESTION_DISABLED')
    expect(client.calls).toHaveLength(0)
    expect(db.cursorWrites).toBe(0)
  })

  it('pages the delta feed, ingests new messages, and persists the deltaLink cursor', async () => {
    const db = new FakeDb()
    const { options, ingested } = baseOptions(db, [
      { messages: [graphMessage('m1'), graphMessage('m2')], nextLink: 'next-1' },
      { messages: [graphMessage('m3')], deltaLink: 'delta-1' },
    ])
    const result = await runMailboxDiscovery(db, options)
    expect(result.ran).toBe(true)
    expect(result.discovered).toBe(3)
    expect(result.ingested).toBe(3)
    expect(result.failures).toBe(0)
    expect(result.cursorUpdated).toBe(true)
    expect(ingested).toHaveLength(3)
    // Discovery never persists sources itself: the intake path owns that.
    expect(db.seenSources.size).toBe(0)
    const cursor = db.cursors.get('MICROSOFT_GRAPH:production:inbox-delta')
    expect(cursor?.cursor_value).toBe('delta-1')
    expect(cursor?.high_water_received_at).toBe('2026-09-28T10:00:00.000Z')
    expect(cursor?.last_succeeded_at).toBeTruthy()
  })

  it('derives skipped counts from the intake DUPLICATE status (no pre-persist)', async () => {
    const db = new FakeDb()
    db.cursors.set('MICROSOFT_GRAPH:production:inbox-delta', {
      id: 'MICROSOFT_GRAPH:production:inbox-delta',
      provider: 'MICROSOFT_GRAPH',
      mailbox_key: 'production',
      scope: 'inbox-delta',
      cursor_value: 'delta-0',
    })
    const client = new FakeDeltaDiscoveryClient([
      { messages: [graphMessage('m1'), graphMessage('m1'), graphMessage('m2')], deltaLink: 'delta-2' },
    ])
    const { options, ingested } = baseOptions(db, [], { client })
    const result = await runMailboxDiscovery(db, options)
    expect(client.calls[0]).toBe('delta-0')
    expect(result.discovered).toBe(2)
    expect(result.skipped).toBe(1)
    expect(result.ingested).toBe(2)
    expect(ingested).toHaveLength(2)
  })

  it('skips tombstoned (@removed) messages without ingesting', async () => {
    const db = new FakeDb()
    const removed = { id: 'gone', '@removed': { reason: 'deleted' } } as unknown as GraphMessage
    const { options, ingested } = baseOptions(db, [
      { messages: [removed, graphMessage('m1')], deltaLink: 'delta-3' },
    ])
    const result = await runMailboxDiscovery(db, options)
    expect(result.skipped).toBe(1)
    expect(result.discovered).toBe(1)
    expect(ingested).toHaveLength(1)
  })

  it('counts ingest failures without aborting the run and does NOT advance the cursor past them', async () => {
    const db = new FakeDb()
    const { options } = baseOptions(db, [{ messages: [graphMessage('m1'), graphMessage('m2')], deltaLink: 'd' }], {
      ingest: async (_d, message) => {
        if (message.id === 'm1') throw new Error('boom')
        return { status: 'IGNORED' }
      },
    })
    const result = await runMailboxDiscovery(db, options)
    expect(result.discovered).toBe(2)
    expect(result.ingested).toBe(1)
    expect(result.failures).toBe(1)
    expect(result.failedIds).toEqual(['m1'])
    expect(result.reason).toBe('DISCOVERY_PARTIAL_FAILURES')
    // Durable cursor rule: never advance past failures; next run re-processes
    // from the previous cursor (successes dedupe via DUPLICATE).
    expect(result.cursorUpdated).toBe(false)
    const cursor = db.cursors.get('MICROSOFT_GRAPH:production:inbox-delta')
    expect(cursor?.cursor_value).not.toBe('d')
  })

  it('enforces the date window client-side even if Graph ignores $filter', async () => {
    const db = new FakeDb()
    const old = graphMessage('old')
    old.receivedDateTime = '2026-09-01T10:00:00.000Z'
    const inWindow = graphMessage('new')
    inWindow.receivedDateTime = '2026-10-12T10:00:00.000Z'
    const { options, ingested } = baseOptions(
      db,
      [{ messages: [old, inWindow], deltaLink: 'd2' }],
      {
        receivedStart: '2026-10-10T00:00:00.000Z',
        receivedEnd: '2026-10-17T23:59:59.999Z',
      },
    )
    const result = await runMailboxDiscovery(db, options)
    expect(result.discovered).toBe(1)
    expect(result.skipped).toBe(1)
    expect(ingested.map((m) => m.id)).toEqual(['new'])
    expect(result.cursorUpdated).toBe(true)
  })

  it('records the error code when the client fails', async () => {
    const db = new FakeDb()
    const client = new FakeDeltaDiscoveryClient([])
    const result = await runMailboxDiscovery(db, {
      provider: 'MICROSOFT_GRAPH',
      mailboxKey: 'production',
      ingestionEnabled: true,
      client,
    })
    expect(result.ran).toBe(true)
    expect(result.reason).toBe('DISCOVERY_FAILED')
    expect(result.errorCode).toBe('FAKE_DELTA_EXHAUSTED')
    const cursor = db.cursors.get('MICROSOFT_GRAPH:production:inbox-delta')
    expect(cursor?.last_error_code).toBe('FAKE_DELTA_EXHAUSTED')
  })
})
