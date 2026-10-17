// Phase 1 — Controlled real-mail preview + selective import.
//
// Preview is READ-ONLY: it lists Graph Inbox messages for a date window via
// /me/messages (never delta), never creates complaints, never advances the
// ingestion cursor, never sends notifications or acknowledgments, and never
// writes to D1. Classification is deterministic only; no model calls.
//
// Import is EXPLICIT selection only: the owner picks message ids from a
// preview window, the server re-validates that each id is in that window
// (fresh listing), then ingests each via ingestEmailMessage with
// allowOwnerSelectedReview=true, so ambiguous forwards/replies become
// REVIEW_REQUIRED instead of being silently IGNORED.

import type { D1Database } from './d1'
import type { EmailProvider } from './email-provider'
import { ingestEmailMessage, extractComplaint } from './ingestion'
import type { GraphMessage, MicrosoftGraphProvider } from './microsoft-graph'

export type MailPreviewClassification =
  | 'COMPLAINT_CANDIDATE'
  | 'FOLLOW_UP_LIKELY'
  | 'NEEDS_REVIEW'
  | 'NOT_COMPLAINT'
  | 'ALREADY_INGESTED'

export interface MailPreviewItem {
  /** Full Graph message id (needed for import). UI must display only maskedId. */
  graphMessageId: string
  maskedId: string
  conversationId: string
  maskedConversationId: string
  receivedAt: string
  sender: string
  /** Sender display: corporate Dunkin address shown fully, others masked. */
  senderDisplay: string
  subject: string
  detectedCaseId: string | null
  detectedStoreNumber: string | null
  isForwardOrReply: boolean
  classification: MailPreviewClassification
  alreadyIngested: boolean
  /** Why this classification was assigned (deterministic rules). */
  classificationReason: string
  /** Existing complaint id if this message links to one (follow-up / duplicate). */
  linkedComplaintId: string | null
}

export const DUNKIN_SENDER = 'customerservice@dunkinbrands.com'

export function maskGraphIdentifier(id: string): string {
  if (id.length <= 8) return '****'
  return `${id.slice(0, 4)}…${id.slice(-4)}`
}

export function maskSenderAddress(sender: string): string {
  const lower = sender.trim().toLowerCase()
  if (lower === DUNKIN_SENDER) return DUNKIN_SENDER
  const at = lower.indexOf('@')
  if (at <= 0) return '***'
  const local = lower.slice(0, at)
  const domain = lower.slice(at + 1)
  if (local.length <= 2) return `**@${domain}`
  return `${local.slice(0, 2)}***@${domain}`
}

export interface PreviewWindow {
  receivedStart: string
  receivedEnd: string
}

/** Validates a preview window: ISO dates, start < end, max 31 days. */
export function validatePreviewWindow(start: string, end: string): { ok: true } | { ok: false; error: string } {
  const s = Date.parse(start)
  const e = Date.parse(end)
  if (!Number.isFinite(s) || !Number.isFinite(e)) return { ok: false, error: 'INVALID_WINDOW' }
  if (s >= e) return { ok: false, error: 'INVALID_WINDOW' }
  if (e - s > 31 * 86_400_000) return { ok: false, error: 'WINDOW_TOO_LARGE' }
  return { ok: true }
}

export const DEFAULT_PREVIEW_WINDOW: PreviewWindow = {
  receivedStart: '2026-10-10T00:00:00.000Z',
  receivedEnd: '2026-10-17T23:59:59.999Z',
}

/**
 * Classifies one preview message from metadata + bodyPreview only.
 * Never fetches full bodies, never writes, never calls a model.
 */
export function classifyPreviewMessage(
  message: GraphMessage,
  alreadyIngestedIds: Set<string>,
  alreadyIngestedCaseIds: Map<string, string>,
  linkedByConversation: Map<string, string>,
): Omit<MailPreviewItem, 'graphMessageId' | 'maskedId' | 'conversationId' | 'maskedConversationId' | 'receivedAt' | 'sender' | 'senderDisplay' | 'subject'> {
  const sender = message.from?.emailAddress?.address?.trim() ?? ''
  const subject = message.subject ?? '(no subject)'
  const bodyPreview = message.bodyPreview ?? ''
  const fakeMessage = {
    id: message.id ?? '',
    threadId: message.conversationId ?? '',
    internalDate: message.receivedDateTime ?? '',
    subject,
    sender,
    recipients: '',
    textBody: bodyPreview,
    htmlBody: undefined,
    messageIdHeader: message.internetMessageId ?? null,
    inReplyTo: null,
    references: null,
  }
  const extraction = extractComplaint(fakeMessage as never)

  // ALREADY_INGESTED takes precedence: do not re-import.
  if (message.id && alreadyIngestedIds.has(message.id)) {
    return {
      detectedCaseId: extraction.externalCaseId ?? null,
      detectedStoreNumber: extraction.storeNumber ?? null,
      isForwardOrReply: extraction.isForwardOrReply,
      classification: 'ALREADY_INGESTED',
      alreadyIngested: true,
      classificationReason: 'Graph message id already in mail_source_messages / gmail_messages',
      linkedComplaintId: extraction.externalCaseId
        ? (alreadyIngestedCaseIds.get(extraction.externalCaseId.toLowerCase()) ?? null)
        : null,
    }
  }
  // Case ID already known -> follow-up or duplicate, not a new case.
  if (extraction.externalCaseId) {
    const linked = alreadyIngestedCaseIds.get(extraction.externalCaseId.toLowerCase())
    if (linked) {
      return {
        detectedCaseId: extraction.externalCaseId,
        detectedStoreNumber: extraction.storeNumber ?? null,
        isForwardOrReply: extraction.isForwardOrReply,
        classification: 'FOLLOW_UP_LIKELY',
        alreadyIngested: false,
        classificationReason: 'Case ID matches an existing complaint',
        linkedComplaintId: linked,
      }
    }
  }
  // Conversation link to an existing complaint -> follow-up likely.
  if (message.conversationId && linkedByConversation.has(message.conversationId)) {
    return {
      detectedCaseId: extraction.externalCaseId ?? null,
      detectedStoreNumber: extraction.storeNumber ?? null,
      isForwardOrReply: extraction.isForwardOrReply,
      classification: 'FOLLOW_UP_LIKELY',
      alreadyIngested: false,
      classificationReason: 'Conversation id links to an existing complaint',
      linkedComplaintId: linkedByConversation.get(message.conversationId) ?? null,
    }
  }
  // Deterministic complaint candidate.
  if (extraction.isComplaint) {
    return {
      detectedCaseId: extraction.externalCaseId ?? null,
      detectedStoreNumber: extraction.storeNumber ?? null,
      isForwardOrReply: extraction.isForwardOrReply,
      classification: 'COMPLAINT_CANDIDATE',
      alreadyIngested: false,
      classificationReason: `Deterministic rules: ${extraction.extractionRules.join(', ') || 'complaint signal'}`,
      linkedComplaintId: null,
    }
  }
  // Ambiguous: Dunkin sender, or forward/reply, or case-id signal without full
  // complaint match in preview (bodyPreview truncated) -> human review, never silent ignore.
  const isDunkin = sender.toLowerCase() === DUNKIN_SENDER
  if (isDunkin || extraction.isForwardOrReply || extraction.hasCaseIdSignal) {
    return {
      detectedCaseId: extraction.externalCaseId ?? null,
      detectedStoreNumber: extraction.storeNumber ?? null,
      isForwardOrReply: extraction.isForwardOrReply,
      classification: 'NEEDS_REVIEW',
      alreadyIngested: false,
      classificationReason: isDunkin
        ? 'Dunkin sender with no deterministic complaint match in preview; full body may hold the complaint'
        : extraction.isForwardOrReply
          ? 'Forward/reply with no deterministic complaint match in preview; quoted original may hold the complaint'
          : 'Case-id signal with no deterministic complaint match in preview',
      linkedComplaintId: null,
    }
  }
  return {
    detectedCaseId: null,
    detectedStoreNumber: null,
    isForwardOrReply: extraction.isForwardOrReply,
    classification: 'NOT_COMPLAINT',
    alreadyIngested: false,
    classificationReason: 'No complaint, case-id, Dunkin-sender, or forward/reply signal in subject/preview',
    linkedComplaintId: null,
  }
}

/** Loads the D1 sets needed for preview classification (single queries, no N+1). */
export async function loadPreviewDedupSets(db: D1Database, messageIds: string[], caseIds: string[], conversationIds: string[]) {
  const alreadyIngestedIds = new Set<string>()
  if (messageIds.length > 0) {
    const placeholders = messageIds.map(() => '?').join(',')
    const rows = await db
      .prepare(`SELECT id FROM mail_source_messages WHERE id IN (${placeholders})`)
      .bind(...messageIds)
      .all<{ id: string }>()
    for (const row of rows.results) alreadyIngestedIds.add(row.id)
    const legacy = await db
      .prepare(`SELECT gmail_message_id AS id FROM gmail_messages WHERE gmail_message_id IN (${placeholders})`)
      .bind(...messageIds)
      .all<{ id: string }>()
    for (const row of legacy.results) alreadyIngestedIds.add(row.id)
  }
  const alreadyIngestedCaseIds = new Map<string, string>()
  if (caseIds.length > 0) {
    const placeholders = caseIds.map(() => '?').join(',')
    const lowerCaseIds = caseIds.map((c) => c.toLowerCase())
    const rows = await db
      .prepare(`SELECT id, external_case_id FROM complaints WHERE lower(external_case_id) IN (${placeholders})`)
      .bind(...lowerCaseIds)
      .all<{ id: string; external_case_id: string }>()
    for (const row of rows.results) alreadyIngestedCaseIds.set(row.external_case_id.toLowerCase(), row.id)
  }
  const linkedByConversation = new Map<string, string>()
  if (conversationIds.length > 0) {
    const placeholders = conversationIds.map(() => '?').join(',')
    const rows = await db
      .prepare(
        `SELECT ms.conversation_id AS conversation_id, cs.complaint_id AS complaint_id
         FROM mail_source_messages ms JOIN complaint_sources cs ON cs.source_message_id = ms.id
         WHERE ms.conversation_id IN (${placeholders})
         ORDER BY cs.linked_at LIMIT 100`,
      )
      .bind(...conversationIds)
      .all<{ conversation_id: string; complaint_id: string }>()
    for (const row of rows.results) {
      if (!linkedByConversation.has(row.conversation_id)) linkedByConversation.set(row.conversation_id, row.complaint_id)
    }
  }
  return { alreadyIngestedIds, alreadyIngestedCaseIds, linkedByConversation }
}

/** Runs a preview: Graph listing + D1 dedup + deterministic classification. No writes. */
export async function runMailPreview(
  db: D1Database,
  provider: MicrosoftGraphProvider,
  window: PreviewWindow,
  top = 50,
  nextLink?: string,
): Promise<{ items: MailPreviewItem[]; nextLink?: string }> {
  const page = await provider.listInboxPreviewPage({
    receivedStart: window.receivedStart,
    receivedEnd: window.receivedEnd,
    top,
    nextLink,
  })
  const messages = page.messages.filter((m) => m.id && m.receivedDateTime)
  // First pass: extract case ids for batch D1 lookup.
  const caseIds: string[] = []
  for (const m of messages) {
    const extraction = extractComplaint({
      id: m.id ?? '',
      threadId: m.conversationId ?? '',
      internalDate: m.receivedDateTime ?? '',
      subject: m.subject ?? '',
      sender: m.from?.emailAddress?.address ?? '',
      recipients: '',
      textBody: m.bodyPreview ?? '',
      htmlBody: undefined,
      messageIdHeader: m.internetMessageId ?? null,
      inReplyTo: null,
      references: null,
    } as never)
    if (extraction.externalCaseId) caseIds.push(extraction.externalCaseId)
  }
  const { alreadyIngestedIds, alreadyIngestedCaseIds, linkedByConversation } = await loadPreviewDedupSets(
    db,
    messages.map((m) => m.id!),
    caseIds,
    messages.map((m) => m.conversationId).filter((c): c is string => Boolean(c)),
  )
  const items: MailPreviewItem[] = messages.map((m) => {
    const sender = m.from?.emailAddress?.address?.trim() ?? ''
    const classification = classifyPreviewMessage(m, alreadyIngestedIds, alreadyIngestedCaseIds, linkedByConversation)
    return {
      graphMessageId: m.id!,
      maskedId: maskGraphIdentifier(m.id!),
      conversationId: m.conversationId ?? '',
      maskedConversationId: m.conversationId ? maskGraphIdentifier(m.conversationId) : '',
      receivedAt: m.receivedDateTime!,
      sender,
      senderDisplay: maskSenderAddress(sender),
      subject: (m.subject ?? '(no subject)').replace(/[\r\n]+/g, ' ').trim().slice(0, 200),
      ...classification,
    }
  })
  return { items, nextLink: page.nextLink }
}

/** Result for one imported message. Never silently drops a failure. */
export interface PreviewImportResult {
  graphMessageId: string
  maskedId: string
  status: string
  complaintId?: string
  errorCode?: string
}

/**
 * Imports explicitly selected messages. Each id is re-validated against a
 * fresh preview listing for the same window (prevents forged ids outside the
 * window). Per-message failures are returned, not swallowed; other messages
 * still import. No cursor is advanced (preview uses /me/messages, not delta).
 */
export async function importSelectedPreviewMessages(
  db: D1Database,
  provider: MicrosoftGraphProvider,
  window: PreviewWindow,
  messageIds: string[],
  ingestionMode: 'LIVE' | 'BACKFILL' = 'BACKFILL',
): Promise<{ results: PreviewImportResult[]; imported: number; failed: number }> {
  const windowCheck = validatePreviewWindow(window.receivedStart, window.receivedEnd)
  if (!windowCheck.ok) throw new Error(windowCheck.error)
  const uniqueIds = Array.from(new Set(messageIds)).slice(0, 50)
  if (uniqueIds.length === 0) return { results: [], imported: 0, failed: 0 }

  // Re-validate: fresh preview listing for the window (up to 100). If a
  // selected id is not in the window, it is a failure, not an import.
  const preview = await runMailPreview(db, provider, window, 100)
  const inWindow = new Set(preview.items.map((i) => i.graphMessageId))
  const results: PreviewImportResult[] = []
  let imported = 0
  let failed = 0
  for (const id of uniqueIds) {
    const masked = maskGraphIdentifier(id)
    if (!inWindow.has(id)) {
      results.push({ graphMessageId: id, maskedId: masked, status: 'FAILED', errorCode: 'NOT_IN_PREVIEW_WINDOW' })
      failed += 1
      continue
    }
    try {
      const full = await (provider as unknown as { getMessage: (id: string) => Promise<Parameters<typeof ingestEmailMessage>[1]> }).getMessage(id)
      const outcome = await ingestEmailMessage(db, full, {
        ingestionMode,
        allowOwnerSelectedReview: true,
      })
      if (outcome.status === 'FAILED_PARSING' || outcome.status === 'FAILED_PERSISTENCE') {
        results.push({ graphMessageId: id, maskedId: masked, status: outcome.status, errorCode: outcome.status })
        failed += 1
      } else {
        results.push({ graphMessageId: id, maskedId: masked, status: outcome.status, complaintId: outcome.complaintId })
        imported += 1
      }
    } catch (error) {
      results.push({
        graphMessageId: id,
        maskedId: masked,
        status: 'FAILED',
        errorCode: error instanceof Error ? error.message.slice(0, 120) : 'IMPORT_FAILED',
      })
      failed += 1
    }
  }
  return { results, imported, failed }
}

// Re-export for callers that need the provider type without importing Graph directly.
export type { EmailProvider }
