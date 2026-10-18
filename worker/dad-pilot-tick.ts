/**
 * Dad-only pilot scheduled discovery tick.
 *
 * A narrowly scoped replacement for the legacy operations cron. Runs
 * incremental Microsoft Graph mailbox discovery on a schedule:
 *
 *   Graph delta (durable cursor) -> ingest new messages as LIVE complaints
 *   -> Dad-only push per new complaint -> deep link to the case in his PWA.
 *
 * Safety properties:
 * - Separate from the legacy runScheduledOperations cron (which stays disabled).
 * - Bounded: max 3 delta pages per run, 25 messages per run. The tick never
 *   loads full application state; it uses targeted queries + the existing
 *   discovery cursor table.
 * - Durable cursor: mail_discovery_cursors persists the delta link; the cursor
 *   never advances past failures (see runMailboxDiscovery).
 * - Dedup: messages already ingested (by graphMessageId) are skipped; a
 *   complaint is created exactly once per message (UNIQUE on external_case_id
 *   + idempotent ingest).
 * - Push: Dad-only (user_id='dad'), LIVE complaints only, durably recorded
 *   in the notifications outbox with SENT/FAILED/SKIPPED; safe to retry.
 *   Provider acceptance (!= iPhone display) is recorded as SENT — the outbox
 *   is the source of truth, not a delivery receipt.
 * - Privacy: push payload contains only case ID, category, store number —
 *   never customer name, email, phone, or complaint text (lock screen safe).
 * - Kill switches: the tick no-ops unless emailIngestionEnabled is true AND
 *   mode is FAMILY_PILOT. External SMS and auto-ack stay off independently.
 */
import { runMailboxDiscovery, DISCOVERY_SCOPE_INBOX_DELTA } from './mail-discovery'
import { MicrosoftGraphProvider } from './microsoft-graph'
import { ingestEmailMessage } from './ingestion'
import type { D1Database } from './d1'

export interface DadPilotTickEnv {
  DB: D1Database
  MS_CLIENT_ID?: string
  MS_CLIENT_SECRET?: string
  MS_TENANT_ID?: string
  MS_REFRESH_TOKEN?: string
  MS_MAILBOX_ADDRESS?: string
  VAPID_PUBLIC_KEY?: string
  VAPID_PRIVATE_KEY?: string
  PUBLIC_BASE_URL?: string
}

export interface DadPilotTickResult {
  ran: boolean
  reason?: string
  discovered: number
  ingested: number
  skipped: number
  failures: number
  cursorUpdated: boolean
}

/** Max delta pages per tick — bounds Graph + D1 work. */
const TICK_MAX_PAGES = 3
/** Max messages ingested per tick — bounds complaint creation + push. */
const TICK_MAX_MESSAGES = 25

export async function runDadPilotDiscoveryTick(
  env: DadPilotTickEnv,
): Promise<DadPilotTickResult> {
  const db = env.DB
  const empty: DadPilotTickResult = {
    ran: false,
    discovered: 0,
    ingested: 0,
    skipped: 0,
    failures: 0,
    cursorUpdated: false,
  }

  // Gate 1: FAMILY_PILOT only.
  const mode = await db
    .prepare(`SELECT value FROM settings WHERE key='notification_mode'`)
    .first<{ value: string }>()
  if (mode?.value !== 'FAMILY_PILOT') return { ...empty, reason: 'NOT_FAMILY_PILOT' }

  // Gate 2: scheduled ingestion must be explicitly enabled. This is separate
  // from the manual preview-import path and from external notifications.
  const ingestion = await db
    .prepare(`SELECT value FROM settings WHERE key='email_ingestion_enabled'`)
    .first<{ value: string }>()
  if (ingestion?.value !== 'true') return { ...empty, reason: 'INGESTION_DISABLED' }

  // Gate 3: Graph must be configured.
  const graph = new MicrosoftGraphProvider({
    clientId: env.MS_CLIENT_ID ?? '',
    clientSecret: env.MS_CLIENT_SECRET ?? '',
    tenant: env.MS_TENANT_ID ?? '',
    refreshToken: env.MS_REFRESH_TOKEN ?? '',
    mailboxAddress: env.MS_MAILBOX_ADDRESS ?? '',
  })
  if (!graph.ready) return { ...empty, reason: 'GRAPH_NOT_CONFIGURED' }

  let ingestedCount = 0
  const result = await runMailboxDiscovery(db, {
    provider: 'MICROSOFT_GRAPH',
    mailboxKey: env.MS_MAILBOX_ADDRESS ?? 'production',
    scope: DISCOVERY_SCOPE_INBOX_DELTA,
    ingestionEnabled: true,
    client: graph as unknown as import('./mail-discovery').DeltaDiscoveryClient,
    maxPages: TICK_MAX_PAGES,
    ingest: async (database, message) => {
      // Bound messages per tick.
      if (ingestedCount >= TICK_MAX_MESSAGES)
        return { status: 'SKIPPED_TICK_BOUND', complaintId: undefined }
      const outcome = await ingestEmailMessage(database, message, {
        ingestionMode: 'LIVE',
        pushEnv: {
          VAPID_PUBLIC_KEY: env.VAPID_PUBLIC_KEY,
          VAPID_PRIVATE_KEY: env.VAPID_PRIVATE_KEY,
          PUBLIC_BASE_URL: env.PUBLIC_BASE_URL,
        },
      })
      if (outcome.complaintId) ingestedCount += 1
      return outcome
    },
  })

  return {
    ran: result.ran,
    reason: result.reason,
    discovered: result.discovered,
    ingested: result.ingested,
    skipped: result.skipped,
    failures: result.failures,
    cursorUpdated: result.cursorUpdated,
  }
}
