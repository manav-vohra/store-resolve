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
 * - Bounded: max 3 delta pages / 25 messages per run. Budget exhaustion
 *   throws DiscoveryBudgetExhausted: the loop stops immediately and the
 *   cursor does NOT advance past unprocessed messages. Next run resumes.
 * - Durable cursor: mail_discovery_cursors persists the delta link; the cursor
 *   never advances past failures.
 * - First-run safety: uses the SAME cursor key as the manual discovery path
 *   ('production') and requires an initial date boundary so historical mail
 *   is never ingested as LIVE on first activation.
 * - Dedup: messages already ingested are skipped; one complaint per message.
 * - Push: Dad-only, LIVE only, gated by dad_push_enabled kill switch
 *   (default OFF). Durably recorded; failed pushes are retryable via the
 *   outbox (see worker/dad-push.ts). Provider acceptance != device display.
 * - Privacy: push payload has case ID/category/store only — no customer data.
 * - Kill switches: no-ops unless FAMILY_PILOT + email_ingestion_enabled.
 *   SMS and auto-ack stay off independently.
 */
import {
  runMailboxDiscovery,
  DISCOVERY_SCOPE_INBOX_DELTA,
  GraphDeltaDiscoveryClient,
  DiscoveryBudgetExhausted,
} from './mail-discovery'
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
/** Max messages processed per tick (including duplicates/skips). */
const TICK_MAX_MESSAGES = 25
/**
 * Cursor key shared with the manual discovery path so the scheduled job
 * resumes from the same position instead of starting over.
 */
const TICK_MAILBOX_KEY = 'production'
/**
 * First-run safety: without an initial date boundary, the first delta sync
 * could pull historical mail and ingest it as LIVE. Require an explicit
 * start date; the tick refuses to run without one until a cursor exists.
 */
const TICK_INITIAL_RECEIVED_START = '2026-10-18T00:00:00.000Z'

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

  // Gate 2: scheduled ingestion must be explicitly enabled.
  const ingestion = await db
    .prepare(`SELECT value FROM settings WHERE key='email_ingestion_enabled'`)
    .first<{ value: string }>()
  if (ingestion?.value !== 'true') return { ...empty, reason: 'INGESTION_DISABLED' }

  // Gate 3: Dad push must be explicitly enabled (separate kill switch).
  // Deploying code must not authorize real pushes by itself.
  const pushEnabled = await db
    .prepare(`SELECT value FROM settings WHERE key='dad_push_enabled'`)
    .first<{ value: string }>()
  const dadPushOn = pushEnabled?.value === 'true'

  // Gate 4: Graph must be configured.
  const provider = new MicrosoftGraphProvider({
    clientId: env.MS_CLIENT_ID ?? '',
    clientSecret: env.MS_CLIENT_SECRET ?? '',
    tenant: env.MS_TENANT_ID ?? '',
    refreshToken: env.MS_REFRESH_TOKEN ?? '',
    mailboxAddress: env.MS_MAILBOX_ADDRESS ?? '',
  })
  if (!provider.ready) return { ...empty, reason: 'GRAPH_NOT_CONFIGURED' }

  // Gate 5: first-run safety. If no cursor exists yet, require the initial
  // date boundary so we never ingest historical mail as LIVE.
  const cursor = await db
    .prepare(`SELECT cursor_value FROM mail_discovery_cursors WHERE provider=? AND mailbox_key=? AND scope=?`)
    .bind('MICROSOFT_GRAPH', TICK_MAILBOX_KEY, DISCOVERY_SCOPE_INBOX_DELTA)
    .first<{ cursor_value: string | null }>()
  const hasCursor = Boolean(cursor?.cursor_value)
  const receivedStart = hasCursor ? undefined : TICK_INITIAL_RECEIVED_START

  let messagesSeen = 0
  const result = await runMailboxDiscovery(db, {
    provider: 'MICROSOFT_GRAPH',
    mailboxKey: TICK_MAILBOX_KEY,
    scope: DISCOVERY_SCOPE_INBOX_DELTA,
    ingestionEnabled: true,
    client: new GraphDeltaDiscoveryClient(provider),
    maxPages: TICK_MAX_PAGES,
    receivedStart,
    ingest: async (database, message) => {
      // Genuine processing budget: count every message handed to ingest
      // (including duplicates). On exhaustion, throw to stop the loop
      // immediately — the cursor does not advance past unprocessed mail.
      messagesSeen += 1
      if (messagesSeen > TICK_MAX_MESSAGES) throw new DiscoveryBudgetExhausted()
      return ingestEmailMessage(database, message, {
        ingestionMode: 'LIVE',
        pushEnv: dadPushOn
          ? {
              VAPID_PUBLIC_KEY: env.VAPID_PUBLIC_KEY,
              VAPID_PRIVATE_KEY: env.VAPID_PRIVATE_KEY,
              PUBLIC_BASE_URL: env.PUBLIC_BASE_URL,
            }
          : undefined,
      })
    },
  })

  // Retry any failed Dad pushes from this run or previous runs.
  // Durable, idempotent, never throws.
  if (dadPushOn) {
    try {
      const { retryFailedDadPushes } = await import('./dad-push')
      await retryFailedDadPushes(db, {
        VAPID_PUBLIC_KEY: env.VAPID_PUBLIC_KEY,
        VAPID_PRIVATE_KEY: env.VAPID_PRIVATE_KEY,
        PUBLIC_BASE_URL: env.PUBLIC_BASE_URL,
      })
    } catch {
      // Retry failures are audited; the tick result still reports discovery.
    }
  }

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
