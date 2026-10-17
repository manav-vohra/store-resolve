// Phase C — Microsoft Graph incremental (delta-query) mailbox discovery.
//
// Wires the mailbox into the existing intake path (ingestEmailMessage) without
// changing any safety posture:
//   - runs only when email ingestion is explicitly enabled (flag default false);
//   - discovery only *discovers*: source persistence, identity, triage, and
//     provisional-intake rules all live in ingestEmailMessage and are unchanged.
//     Discovery never pre-persists sources (a pre-persisted row under a different
//     id would make the intake path misread every message as DUPLICATE);
//   - discovered/skipped counts are derived from the intake result status;
//   - cursors persist in mail_discovery_cursors (existing table, no migration);
//   - no live Graph calls in tests: inject a fake DeltaDiscoveryClient.
//
// Delta protocol recap: the first call hits .../messages/delta; Graph pages
// results with @odata.nextLink and ends the sync with @odata.deltaLink. The
// deltaLink is the cursor for the *next* run and yields only changes since.

import type { D1Database } from './d1'
import { EmailProviderError } from './email-provider'
import type { NormalizedEmailMessage } from './email-provider'
import { ingestEmailMessage } from './ingestion'
import {
  normalizeGraphMessage,
  type GraphMessage,
  type InboxDeltaPage,
  type MicrosoftGraphProvider,
} from './microsoft-graph'

/** Minimal delta-page client so tests never touch the network. */
export interface DeltaDiscoveryClient {
  fetchDeltaPage(
    link?: string,
    options?: { receivedStart?: string; receivedEnd?: string },
  ): Promise<InboxDeltaPage>
}

/** Live client: thin wrapper over the real Graph provider. */
export class GraphDeltaDiscoveryClient implements DeltaDiscoveryClient {
  constructor(private readonly provider: MicrosoftGraphProvider) {}
  fetchDeltaPage(
    link?: string,
    options?: { receivedStart?: string; receivedEnd?: string },
  ): Promise<InboxDeltaPage> {
    return this.provider.listInboxDeltaPage(link, options)
  }
}

/** Fake client for tests: replays canned pages. */
export class FakeDeltaDiscoveryClient implements DeltaDiscoveryClient {
  public calls: Array<string | undefined> = []
  constructor(private readonly pages: InboxDeltaPage[]) {}
  async fetchDeltaPage(link?: string): Promise<InboxDeltaPage> {
    this.calls.push(link)
    const page = this.pages[this.calls.length - 1]
    if (!page) throw new EmailProviderError('FAKE_DELTA_EXHAUSTED', 'no more fake pages')
    return page
  }
}

export const DISCOVERY_SCOPE_INBOX_DELTA = 'inbox-delta'
const MAX_PAGES_DEFAULT = 10

export interface MailboxDiscoveryOptions {
  provider: 'MICROSOFT_GRAPH'
  mailboxKey: string
  scope?: string
  /** Mirrors the email_ingestion_enabled safety flag. Discovery never runs when false. */
  ingestionEnabled: boolean
  client: DeltaDiscoveryClient
  maxPages?: number
  /** Optional receivedDateTime window for the initial delta (bounds historical mail). */
  receivedStart?: string
  receivedEnd?: string
  /** Defaults to the real intake path. Tests inject a spy. */
  ingest?: (
    db: D1Database,
    message: NormalizedEmailMessage,
  ) => Promise<{ status: string; complaintId?: string }>
}

export interface MailboxDiscoveryResult {
  ran: boolean
  reason?: 'INGESTION_DISABLED' | 'DISCOVERY_FAILED' | 'DISCOVERY_PARTIAL_FAILURES'
  discovered: number
  ingested: number
  skipped: number
  failures: number
  /** Graph message ids that failed this run. Never silently dropped: the cursor does NOT advance past them. */
  failedIds: string[]
  cursorUpdated: boolean
  errorCode?: string
}

const cursorId = (provider: string, mailboxKey: string, scope: string) =>
  `${provider}:${mailboxKey}:${scope}`

async function readCursor(
  db: D1Database,
  provider: string,
  mailboxKey: string,
  scope: string,
): Promise<string | null> {
  const row = await db
    .prepare(
      'SELECT cursor_value FROM mail_discovery_cursors WHERE provider=? AND mailbox_key=? AND scope=?',
    )
    .bind(provider, mailboxKey, scope)
    .first<{ cursor_value: string | null }>()
  return row?.cursor_value ?? null
}

async function writeCursor(
  db: D1Database,
  provider: string,
  mailboxKey: string,
  scope: string,
  patch: {
    cursorValue?: string | null
    highWaterReceivedAt?: string | null
    lastStartedAt?: string | null
    lastSucceededAt?: string | null
    lastErrorCode?: string | null
  },
): Promise<void> {
  const now = new Date().toISOString()
  const current = await db
    .prepare(
      'SELECT cursor_value,high_water_received_at FROM mail_discovery_cursors WHERE id=?',
    )
    .bind(cursorId(provider, mailboxKey, scope))
    .first<{ cursor_value: string | null; high_water_received_at: string | null }>()
  const cursorValue = patch.cursorValue !== undefined ? patch.cursorValue : current?.cursor_value ?? null
  const highWater =
    patch.highWaterReceivedAt !== undefined
      ? patch.highWaterReceivedAt
      : current?.high_water_received_at ?? null
  await db
    .prepare(
      `INSERT INTO mail_discovery_cursors(
         id,provider,mailbox_key,scope,cursor_value,high_water_received_at,
         last_started_at,last_succeeded_at,last_error_code,created_at,updated_at
       ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
       ON CONFLICT(id) DO UPDATE SET
         cursor_value=excluded.cursor_value,
         high_water_received_at=excluded.high_water_received_at,
         last_started_at=COALESCE(excluded.last_started_at,mail_discovery_cursors.last_started_at),
         last_succeeded_at=COALESCE(excluded.last_succeeded_at,mail_discovery_cursors.last_succeeded_at),
         last_error_code=excluded.last_error_code,
         updated_at=excluded.updated_at`,
    )
    .bind(
      cursorId(provider, mailboxKey, scope),
      provider,
      mailboxKey,
      scope,
      cursorValue,
      highWater,
      patch.lastStartedAt ?? null,
      patch.lastSucceededAt ?? null,
      patch.lastErrorCode ?? null,
      now,
      now,
    )
    .run()
}

const isRemoved = (message: GraphMessage): boolean =>
  '@removed' in message &&
  typeof (message as { '@removed'?: unknown })['@removed'] === 'object'

export async function runMailboxDiscovery(
  db: D1Database,
  options: MailboxDiscoveryOptions,
): Promise<MailboxDiscoveryResult> {
  const empty: MailboxDiscoveryResult = {
    ran: false,
    discovered: 0,
    ingested: 0,
    skipped: 0,
    failures: 0,
    failedIds: [],
    cursorUpdated: false,
  }
  if (!options.ingestionEnabled) return { ...empty, reason: 'INGESTION_DISABLED' }

  const provider = options.provider
  const mailboxKey = options.mailboxKey
  const scope = options.scope ?? DISCOVERY_SCOPE_INBOX_DELTA
  const ingest = options.ingest ?? ((database, message) => ingestEmailMessage(database, message))
  const now = new Date().toISOString()
  await writeCursor(db, provider, mailboxKey, scope, { lastStartedAt: now })

  const result: MailboxDiscoveryResult = {
    ran: true,
    discovered: 0,
    ingested: 0,
    skipped: 0,
    failures: 0,
    failedIds: [],
    cursorUpdated: false,
  }
  const initialCursor = (await readCursor(db, provider, mailboxKey, scope)) ?? undefined
  let link: string | undefined = initialCursor
  let highWater: string | null = null
  const maxPages = options.maxPages ?? MAX_PAGES_DEFAULT
  const windowStart = options.receivedStart ? Date.parse(options.receivedStart) : NaN
  const windowEnd = options.receivedEnd ? Date.parse(options.receivedEnd) : NaN
  const inWindow = (receivedAt: string) => {
    const t = Date.parse(receivedAt)
    if (!Number.isFinite(t)) return false
    if (Number.isFinite(windowStart) && t < windowStart) return false
    if (Number.isFinite(windowEnd) && t >= windowEnd) return false
    return true
  }

  try {
    for (let page = 0; page < maxPages; page += 1) {
      let delta: InboxDeltaPage
      try {
        // Only the initial call (no link) carries the date window. Continuation
        // links already encode Graph's paging state; passing a filter with
        // them would be invalid.
        delta = await options.client.fetchDeltaPage(
          link,
          link ? undefined : { receivedStart: options.receivedStart, receivedEnd: options.receivedEnd },
        )
      } catch (fetchError) {
        // Cursor expired / invalid (Graph 410 or invalid delta token): reset to
        // a fresh sync once, with the date window, instead of silently skipping.
        const isCursorError =
          fetchError instanceof EmailProviderError &&
          (fetchError.code.includes('410') ||
            fetchError.code.includes('INVALID') ||
            fetchError.message.includes('410') ||
            fetchError.message.toLowerCase().includes('delta'))
        if (isCursorError && link && link !== initialCursor) {
          // link was a continuation we already persisted; fall through to the
          // outer catch which leaves the previous cursor intact.
          throw fetchError
        }
        if (isCursorError && link) {
          await writeCursor(db, provider, mailboxKey, scope, {
            cursorValue: null,
            lastErrorCode: 'DELTA_CURSOR_EXPIRED',
          })
          link = undefined
          delta = await options.client.fetchDeltaPage(undefined, {
            receivedStart: options.receivedStart,
            receivedEnd: options.receivedEnd,
          })
        } else {
          throw fetchError
        }
      }
      for (const graphMessage of delta.messages) {
        if (isRemoved(graphMessage)) {
          result.skipped += 1
          continue
        }
        let normalized: NormalizedEmailMessage
        try {
          normalized = normalizeGraphMessage(graphMessage)
        } catch {
          result.failures += 1
          if (graphMessage.id) result.failedIds.push(graphMessage.id)
          continue
        }
        // Enforce the date window client-side even if Graph ignored $filter.
        // Out-of-window mail is skipped, not ingested, not a failure.
        if (!inWindow(normalized.internalDate)) {
          result.skipped += 1
          continue
        }
        if (!highWater || normalized.internalDate > highWater) highWater = normalized.internalDate
        // The intake path owns source persistence (INSERT OR IGNORE +
        // lease/resume logic). Discovery only hands the message over and
        // derives its counts from the intake status: a DUPLICATE means the
        // message was already seen, anything else means it was ingested.
        try {
          const outcome = await ingest(db, normalized)
          if (outcome.status === 'DUPLICATE') {
            result.skipped += 1
          } else {
            result.discovered += 1
            result.ingested += 1
          }
        } catch {
          result.discovered += 1
          result.failures += 1
          result.failedIds.push(normalized.id)
        }
      }
      if (delta.deltaLink) {
        // Durable cursor rule: NEVER advance past failures. If any message in
        // this run failed, keep the previous cursor so the next run re-processes
        // from before the failed batch (successes dedupe via DUPLICATE). Failed
        // ids are returned explicitly; nothing is silently skipped.
        if (result.failures > 0) {
          await writeCursor(db, provider, mailboxKey, scope, {
            lastErrorCode: 'DISCOVERY_PARTIAL_FAILURES',
          })
          return {
            ...result,
            reason: 'DISCOVERY_PARTIAL_FAILURES',
            errorCode: 'DISCOVERY_PARTIAL_FAILURES',
            cursorUpdated: false,
          }
        }
        link = delta.deltaLink
        await writeCursor(db, provider, mailboxKey, scope, {
          cursorValue: delta.deltaLink,
          highWaterReceivedAt: highWater,
          lastSucceededAt: new Date().toISOString(),
          lastErrorCode: null,
        })
        result.cursorUpdated = true
        return result
      }
      if (!delta.nextLink) break
      link = delta.nextLink
    }
    // Ran out of pages without a deltaLink: only persist the nextLink when
    // there were no failures, so a failed batch is re-processed from the
    // previous cursor instead of being skipped.
    if (result.failures > 0) {
      await writeCursor(db, provider, mailboxKey, scope, {
        lastErrorCode: 'DISCOVERY_PARTIAL_FAILURES',
      })
      return {
        ...result,
        reason: 'DISCOVERY_PARTIAL_FAILURES',
        errorCode: 'DISCOVERY_PARTIAL_FAILURES',
        cursorUpdated: false,
      }
    }
    if (link && link !== initialCursor) {
      await writeCursor(db, provider, mailboxKey, scope, { cursorValue: link })
      result.cursorUpdated = true
    }
    return result
  } catch (error) {
    const code = error instanceof EmailProviderError ? error.code : 'DISCOVERY_FAILED'
    await writeCursor(db, provider, mailboxKey, scope, { lastErrorCode: code })
    return { ...result, reason: 'DISCOVERY_FAILED', errorCode: code }
  }
}
