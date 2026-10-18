import type { AppConfig, Complaint, Severity } from '../src/lib/types'
import { createComplaint } from '../src/lib/workflow'
import type { D1Database } from './d1'
import { loadState, persistState } from './d1'
import type { EmailProvider, NormalizedEmailMessage } from './email-provider'
import { EmailProviderError } from './email-provider'
import {
  D1MailFoundationRepository,
  type CanonicalEventCommit,
  type IdentityResolutionSignals,
  type IngestionMode,
} from './mail-foundation'
import {
  DisabledModelTransport,
  OpenAiMailInterpreter,
  buildInterpretationInput,
  interpretWithConfig,
  loadMailInterpretationConfig,
  recordInterpretation,
  type InterpretationStoreResolution,
  type MailInterpretationConfig,
  type MailInterpretationResult,
  type MailInterpreter,
} from './mail-interpretation'
import {
  isApprovedPilotMessageMetadata,
  maskGraphIdentifier,
  pilotBodyDiagnosticCandidates,
  pilotMessageSelector,
  type PilotUniquenessMetadata,
  type MicrosoftGraphProvider,
} from './microsoft-graph'
import {
  LogNotificationTransport,
  NotificationService,
  criticalEscalationEvent,
  loadNotificationConfig,
} from './notifications'

export type EmailProcessingStatus =
  | 'PROCESSING'
  | 'IN_PROGRESS'
  | 'PROCESSED'
  | 'IGNORED'
  | 'ROUTING_REVIEW'
  | 'REVIEW_REQUIRED'
  | 'DUPLICATE'
  | 'FOLLOW_UP'
  | 'FAILED_PARSING'
  | 'FAILED_PERSISTENCE'

export interface ComplaintExtraction {
  isComplaint: boolean
  /** True when subject/body indicates a forward or reply (Fw/Fwd/Re, quoted original, etc.). */
  isForwardOrReply: boolean
  /** True when a case/reference ID signal is present (even if the outer text is short). */
  hasCaseIdSignal: boolean
  externalCaseId?: string
  storeNumber?: string
  locationHint?: string
  customerName?: string
  customerEmail?: string
  customerPhone?: string
  category: string
  severity: Severity
  occurrenceAt?: string
  details: string
  /** Which deterministic branches fired, in evaluation order (Phase B evidence). */
  extractionRules: string[]
}

/** Forward/reply markers in subject or body. Exported for preview classification. */
export function isForwardOrReplyMessage(subject: string, textBody: string): boolean {
  const subjectTrim = subject.trim().toLowerCase()
  if (/^(fw|fwd|re|ré)\s*:/.test(subjectTrim)) return true
  if (/^(fw|fwd|re)\s*\[.+\]/.test(subjectTrim)) return true
  const body = textBody.slice(0, 2000).toLowerCase()
  return (
    body.includes('begin forwarded message') ||
    body.includes('---original message---') ||
    body.includes('-----original message-----') ||
    body.includes('original message:') ||
    body.includes('forwarded message') ||
    (body.includes('on ') && body.includes('wrote:') && body.includes('from:'))
  )
}

const compact = (value: string) => value.replace(/\s+/g, ' ').trim()
const normalizeAlias = (value: string) =>
  compact(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
const first = (value: string, pattern: RegExp) => value.match(pattern)?.[1]?.trim()

export function extractComplaint(message: NormalizedEmailMessage): ComplaintExtraction {
  const combined = `${message.subject}\n${message.textBody}`
  const rules: string[] = []
  const match = (text: string, pattern: RegExp, rule: string) => {
    const value = first(text, pattern)
    if (value) rules.push(rule)
    return value
  }
  const test = (pattern: RegExp, rule: string) => {
    const hit = pattern.test(combined)
    if (hit) rules.push(rule)
    return hit
  }
  const isForwardOrReply = isForwardOrReplyMessage(message.subject, message.textBody)
  if (isForwardOrReply) rules.push('FORWARD_REPLY_MARKER')
  const hasCaseIdSignal = /\bcase\s*[:#]\s*[A-Z0-9][A-Z0-9-]{1,}/i.test(combined)
  if (hasCaseIdSignal) rules.push('HAS_CASE_ID_SIGNAL')
  // Forward/reply carrying a case ID is a complaint candidate even when the
  // outer text is just "FYI" / "See below" — the quoted original may hold the
  // complaint. Rule is pushed unconditionally (not in the || chain) so it is
  // recorded even when IS_CASE_REFERENCE already fired.
  if (isForwardOrReply && hasCaseIdSignal) rules.push('FORWARD_REPLY_WITH_CASE_ID')
  const isComplaint =
    test(
      /\b(complaint|guest concern|customer concern|customer issue|case id|reference id|foreign object|food poison|allerg|injury|not happy|unhappy|disappointed|disgusting)\b/i,
      'IS_COMPLAINT_KEYWORD',
    ) ||
    // "Case: DD-2024" / "Case #XYZ" style references are complaint identity
    // signals even without the word "complaint" (recall matters: a missed
    // complaint vanishes, an over-triaged one gets human review).
    test(/\bcase\s*[:#]\s*[A-Z0-9][A-Z0-9-]{1,}/i, 'IS_CASE_REFERENCE') ||
    (isForwardOrReply && hasCaseIdSignal)
  let category: string
  if (test(/clean|sanit|bathroom|dirty/i, 'CATEGORY_CLEANLINESS')) category = 'Cleanliness'
  else if (test(/staff|employee|service|rude|wait/i, 'CATEGORY_SERVICE')) category = 'Service'
  else if (test(/food|drink|coffee|order|product/i, 'CATEGORY_PRODUCT')) category = 'Product quality'
  else {
    rules.push('CATEGORY_OTHER')
    category = 'Other'
  }
  let severity: Severity
  if (test(/injur|hospital|allerg|threat|violence|fire/i, 'SEVERITY_CRITICAL')) severity = 'CRITICAL'
  else if (test(/health|safety|contamin|foreign object/i, 'SEVERITY_HIGH')) severity = 'HIGH'
  else if (test(/refund|repeat|multiple|escalat/i, 'SEVERITY_MEDIUM')) severity = 'MEDIUM'
  else {
    rules.push('SEVERITY_LOW')
    severity = 'LOW'
  }
  return {
    isComplaint,
    isForwardOrReply,
    hasCaseIdSignal,
    externalCaseId:
      match(
        combined,
        /(?:^|\n)\s*(?:complaint\s+reference|case|complaint|reference|ref)\s*(?:id|number|no\.?|#)\s*[:#-]?\s*([A-Z0-9][A-Z0-9-]{2,50})/im,
        'CASE_ID_LINE_ANCHORED',
      ) ??
      match(combined, /\b(?:case|reference)\s*[:#-]\s*([A-Z0-9][A-Z0-9-]{2,50})/i, 'CASE_ID_INLINE') ??
      match(combined, /\bDBI\s+Case\s*#\s*\(\s*([A-Z0-9][A-Z0-9-]{2,50})\s*\)/i, 'CASE_ID_DBI_PARENS'),
    storeNumber:
      match(combined, /\bPC\s*(?:number|no\.?|#)?\s*[:#.-]?\s*(\d{3,8})\b/i, 'STORE_NUMBER_PC') ??
      match(
        combined,
        /\bprofit\s*center\s*(?:number|no\.?|#)?\s*[:#.-]?\s*(\d{3,8})\b/i,
        'STORE_NUMBER_PROFIT_CENTER',
      ) ??
      match(combined, /\b(?:store|location)\s*(?:number|no\.?|#)?\s*[:#-]?\s*(\d{3,8})\b/i, 'STORE_NUMBER_LABELED') ??
      match(combined, /\b(\d{3,8})-DD\b/i, 'STORE_NUMBER_DD_SUFFIX'),
    locationHint: match(combined, /\b(?:store address|location|address)\s*:\s*([^\n]{4,160})/i, 'LOCATION_HINT'),
    customerName: match(combined, /\b(?:customer|guest)\s*name\s*:\s*([^\n]{2,100})/i, 'CUSTOMER_NAME'),
    customerEmail: match(combined, /\b(?:customer|guest)\s*email\s*:\s*([^\s<>]+@[^\s<>]+)/i, 'CUSTOMER_EMAIL'),
    customerPhone: match(combined, /\b(?:customer|guest)\s*phone\s*:\s*([+()\d .-]{7,25})/i, 'CUSTOMER_PHONE'),
    occurrenceAt: match(
      combined,
      /\b(?:incident|occurrence)\s*(?:date|time)?\s*:\s*([^\n]{4,80})/i,
      'OCCURRENCE_AT',
    ),
    category,
    severity,
    details: message.textBody.trim().slice(0, 20_000),
    extractionRules: rules,
  }
}

async function resolveStore(
  db: D1Database,
  extraction: ComplaintExtraction,
): Promise<{ storeNumber?: string; reason: InterpretationStoreResolution['reason'] }> {
  if (extraction.storeNumber) {
    const exact = await db
      .prepare('SELECT dunkin_store_number FROM stores WHERE dunkin_store_number=? AND active=1')
      .bind(extraction.storeNumber)
      .first<{ dunkin_store_number: string }>()
    if (exact) return { storeNumber: exact.dunkin_store_number, reason: 'EXACT_STORE_NUMBER' }
  }
  if (extraction.locationHint) {
    const alias = await db
      .prepare(
        'SELECT s.dunkin_store_number FROM store_aliases a JOIN stores s ON s.id=a.store_id WHERE a.alias_normalized=? AND s.active=1',
      )
      .bind(normalizeAlias(extraction.locationHint))
      .first<{ dunkin_store_number: string }>()
    if (alias) return { storeNumber: alias.dunkin_store_number, reason: 'EXACT_ALIAS' }
  }
  return { reason: 'NO_DETERMINISTIC_STORE_MATCH' }
}

async function recordIntegrationEvent(
  db: D1Database,
  eventType: string,
  entityId: string,
  outcome: string,
  detailCode?: string,
  metadata?: Record<string, unknown>,
) {
  await db
    .prepare(
      'INSERT INTO integration_events(id,integration,event_type,entity_id,outcome,detail_code,metadata,created_at) VALUES(?,?,?,?,?,?,?,?)',
    )
    .bind(
      crypto.randomUUID(),
      'MICROSOFT_GRAPH',
      eventType,
      entityId,
      outcome,
      detailCode ?? null,
      metadata ? JSON.stringify(metadata) : null,
      new Date().toISOString(),
    )
    .run()
}

const requirePilotSafety = (config: AppConfig) => {
  if (config.mode !== 'FAMILY_PILOT')
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_MODE_REQUIRED',
      'Controlled mail ingestion requires FAMILY_PILOT mode',
    )
  if (config.externalNotificationsEnabled)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_SMS_MUST_BE_OFF',
      'External notifications must remain disabled',
    )
  if (config.emailAckEnabled)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_ACK_MUST_BE_OFF',
      'Email acknowledgments must remain disabled',
    )
  if (config.emailIngestionEnabled)
    throw new EmailProviderError(
      'MS_GRAPH_BROAD_INGESTION_MUST_BE_OFF',
      'Scheduled email ingestion must remain disabled during the controlled test',
    )
}

const approvedPairIdentityMatches = (
  candidate: PilotUniquenessMetadata,
  message: NormalizedEmailMessage,
) =>
  message.id === candidate.id &&
  message.threadId === candidate.conversationId &&
  message.internalDate === new Date(candidate.receivedDateTime).toISOString() &&
  Boolean(message.messageIdHeader) &&
  (message.sender.trim().toLowerCase() === pilotMessageSelector.senderAddress ||
    message.sender.toLowerCase().endsWith(`<${pilotMessageSelector.senderAddress}>`))

export async function ingestApprovedPilotCasePair(
  db: D1Database,
  emailProvider: MicrosoftGraphProvider,
  config: AppConfig,
): Promise<{
  accessedCount: 2
  complaintId: string
  initialStatus: EmailProcessingStatus
  followUpStatus: EmailProcessingStatus
  initialMessageId: string
  followUpMessageId: string
  caseId: string
  rawStoreId: string
  normalizedStoreId: string
}> {
  requirePilotSafety(config)
  if (!emailProvider.ready)
    throw new EmailProviderError('MS_GRAPH_NOT_CONFIGURED', 'Microsoft Graph is not configured')

  await emailProvider.verifyConnection()
  const priorDiagnostic = await db
    .prepare(
      `SELECT detail_code,metadata FROM integration_events WHERE integration='MICROSOFT_GRAPH' AND event_type='PILOT_BODY_DIAGNOSTIC' ORDER BY created_at DESC LIMIT 1`,
    )
    .first<{ detail_code: string; metadata: string }>()
  let diagnosticMetadata: {
    classification?: string
    candidateA?: { messageId?: string }
    candidateB?: { messageId?: string }
  } = {}
  try {
    diagnosticMetadata = priorDiagnostic?.metadata ? JSON.parse(priorDiagnostic.metadata) : {}
  } catch {
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_DIAGNOSTIC_INVALID',
      'The approved body diagnostic record was invalid',
    )
  }
  if (
    priorDiagnostic?.detail_code !== 'FOLLOW_UP' ||
    diagnosticMetadata.classification !== 'FOLLOW_UP' ||
    diagnosticMetadata.candidateA?.messageId !== pilotBodyDiagnosticCandidates[0].maskedId ||
    diagnosticMetadata.candidateB?.messageId !== pilotBodyDiagnosticCandidates[1].maskedId
  )
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_DIAGNOSTIC_REQUIRED',
      'The approved two-message FOLLOW_UP diagnostic was not available',
    )

  const candidates = (await emailProvider.findPilotBodyDiagnosticCandidates()).sort(
    (left, right) => Date.parse(left.receivedDateTime) - Date.parse(right.receivedDateTime),
  )
  const selectionMatches =
    candidates.length === pilotBodyDiagnosticCandidates.length &&
    candidates.every(
      (candidate, index) =>
        maskGraphIdentifier(candidate.id) === pilotBodyDiagnosticCandidates[index].maskedId &&
        candidate.receivedDateTime === pilotBodyDiagnosticCandidates[index].receivedDateTime,
    )
  if (!selectionMatches)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_PAIR_SELECTION_CHANGED',
      'The approved two-message selection changed',
    )

  const existingComplaint = await db
    .prepare('SELECT id FROM complaints WHERE lower(external_case_id)=lower(?)')
    .bind(pilotMessageSelector.caseId)
    .first<{ id: string }>()
  const existingSources = await db
    .prepare('SELECT gmail_message_id FROM gmail_messages WHERE gmail_message_id IN (?,?)')
    .bind(candidates[0].id, candidates[1].id)
    .all<{ gmail_message_id: string }>()
  if (existingComplaint || existingSources.results.length)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_CASE_ALREADY_PRESENT',
      'The approved case or one of its source messages was already persisted',
    )

  const initialMessage = await emailProvider.getPilotBodyDiagnosticMessage(candidates[0].id)
  const followUpMessage = await emailProvider.getPilotBodyDiagnosticMessage(candidates[1].id)
  if (
    !approvedPairIdentityMatches(candidates[0], initialMessage) ||
    !approvedPairIdentityMatches(candidates[1], followUpMessage)
  )
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_PAIR_IDENTITY_MISMATCH',
      'An approved Microsoft Graph message no longer matched its diagnosed identity',
    )
  const initialExtraction = extractComplaint(initialMessage)
  const followUpExtraction = extractComplaint(followUpMessage)
  const normalizedInitialStore = initialExtraction.storeNumber?.replace(/\D/g, '')
  const normalizedFollowUpStore = followUpExtraction.storeNumber?.replace(/\D/g, '')
  if (
    !initialExtraction.isComplaint ||
    !followUpExtraction.isComplaint ||
    initialExtraction.externalCaseId !== pilotMessageSelector.caseId ||
    followUpExtraction.externalCaseId !== pilotMessageSelector.caseId ||
    normalizedInitialStore !== '350909' ||
    normalizedFollowUpStore !== '350909'
  )
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_PAIR_PARSE_MISMATCH',
      'The approved messages no longer parsed as the diagnosed case and store',
    )

  const attemptId = crypto.randomUUID()
  const startedAt = new Date().toISOString()
  const claimed = await db
    .prepare(
      `UPDATE background_job_locks SET locked_until='9999-12-30T23:59:59.999Z',owner_run_id=?,updated_at=? WHERE job_name='EMAIL_PILOT' AND locked_until='9999-12-31T23:59:59.999Z'`,
    )
    .bind(attemptId, startedAt)
    .run()
  if (Number(claimed.meta?.changes ?? 0) !== 1)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_LOCK_RESET_FAILED',
      'The previous one-shot lock was not available for the approved reset',
    )
  await recordIntegrationEvent(db, 'PILOT_INGESTION_LOCK_RESET', attemptId, 'SUCCESS')
  await recordIntegrationEvent(db, 'PILOT_CASE_PAIR_INGESTION_STARTED', attemptId, 'SUCCESS')

  const initial = await ingestEmailMessage(db, initialMessage)
  if (initial.status !== 'ROUTING_REVIEW' || !initial.complaintId)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_INITIAL_RESULT_MISMATCH',
      'The initial source did not create the expected routing-review complaint',
    )
  await db
    .prepare(
      `UPDATE gmail_messages SET acknowledgment_status='DISABLED',updated_at=? WHERE gmail_message_id=?`,
    )
    .bind(new Date().toISOString(), initialMessage.id)
    .run()
  const followUp = await ingestEmailMessage(db, followUpMessage)
  if (followUp.status !== 'FOLLOW_UP' || followUp.complaintId !== initial.complaintId)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_FOLLOW_UP_RESULT_MISMATCH',
      'The later source did not attach as a follow-up to the initial complaint',
    )
  await recordIntegrationEvent(
    db,
    'PILOT_CASE_PAIR_INGESTION_COMPLETED',
    initial.complaintId,
    'SUCCESS',
    'FOLLOW_UP',
    {
      attemptId,
      caseId: pilotMessageSelector.caseId,
      rawStoreId: pilotMessageSelector.storeToken,
      normalizedStoreId: '350909',
      initialMessageId: maskGraphIdentifier(initialMessage.id),
      followUpMessageId: maskGraphIdentifier(followUpMessage.id),
    },
  )
  return {
    accessedCount: 2,
    complaintId: initial.complaintId,
    initialStatus: initial.status,
    followUpStatus: followUp.status,
    initialMessageId: maskGraphIdentifier(initialMessage.id),
    followUpMessageId: maskGraphIdentifier(followUpMessage.id),
    caseId: pilotMessageSelector.caseId,
    rawStoreId: pilotMessageSelector.storeToken,
    normalizedStoreId: '350909',
  }
}

export async function ingestSinglePilotComplaint(
  db: D1Database,
  emailProvider: MicrosoftGraphProvider,
  config: AppConfig,
): Promise<{
  accessed: boolean
  status?: EmailProcessingStatus
  complaintId?: string
  messageId?: string
  conversationId?: string
  matchCount?: 0 | 1 | '2+'
  inspectedCount?: number
}> {
  requirePilotSafety(config)
  if (!emailProvider.ready)
    throw new EmailProviderError('MS_GRAPH_NOT_CONFIGURED', 'Microsoft Graph is not configured')

  await emailProvider.verifyConnection()
  const selection = await emailProvider.findPilotComplaintCandidates()
  const matchCount: 0 | 1 | '2+' =
    selection.hasMore || selection.candidates.length > 1
      ? '2+'
      : selection.candidates.length === 1
        ? 1
        : 0
  const inspectedCount = selection.inspectedCandidates.length
  if (matchCount !== 1) return { accessed: false, matchCount, inspectedCount }

  const selected = selection.candidates[0]
  if (!selected)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_INVALID_SELECTION',
      'Microsoft Graph pilot selection was not uniquely available',
    )
  const attemptId = crypto.randomUUID()
  const startedAt = new Date().toISOString()
  const claimed = await db
    .prepare(
      `UPDATE background_job_locks SET locked_until='9999-12-30T23:59:59.999Z',owner_run_id=?,updated_at=? WHERE job_name='EMAIL_PILOT' AND locked_until='9999-12-31T23:59:59.999Z'`,
    )
    .bind(attemptId, startedAt)
    .run()
  if (Number(claimed.meta?.changes ?? 0) !== 1)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_LOCK_RESET_FAILED',
      'The previous one-shot lock was not available for the approved reset',
    )
  await recordIntegrationEvent(db, 'PILOT_INGESTION_LOCK_RESET', attemptId, 'SUCCESS')
  await recordIntegrationEvent(db, 'PILOT_INGESTION_ATTEMPT_STARTED', attemptId, 'SUCCESS')

  const message = await emailProvider.getMessage(selected.id)
  const senderMatches =
    message.sender.trim().toLowerCase() === pilotMessageSelector.senderAddress ||
    message.sender.toLowerCase().endsWith(`<${pilotMessageSelector.senderAddress}>`)
  const receivedAt = Date.parse(message.internalDate)
  const approvedMetadata = {
    senderMatched: senderMatches,
    receivedWindowMatched:
      Number.isFinite(receivedAt) &&
      receivedAt >= Date.parse(pilotMessageSelector.receivedStart) &&
      receivedAt <= Date.parse(pilotMessageSelector.receivedEnd),
    caseIdMatched: message.subject
      .toLowerCase()
      .includes(pilotMessageSelector.caseId.toLowerCase()),
    subjectPhraseMatched: message.subject
      .toLowerCase()
      .includes(pilotMessageSelector.subjectPhrase.toLowerCase()),
    storeTokenMatched: message.subject
      .toLowerCase()
      .includes(pilotMessageSelector.storeToken.toLowerCase()),
  }
  if (
    message.id !== selected.id ||
    message.threadId !== selected.conversationId ||
    Date.parse(message.internalDate) !== Date.parse(selected.receivedDateTime) ||
    !isApprovedPilotMessageMetadata(approvedMetadata)
  )
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_IDENTITY_MISMATCH',
      'Microsoft Graph pilot message identity did not match the selected record',
    )
  const extraction = extractComplaint(message)
  if (!extraction.isComplaint)
    throw new EmailProviderError(
      'MS_GRAPH_PILOT_NOT_COMPLAINT',
      'The selected message did not satisfy the controlled complaint checks',
    )

  await recordIntegrationEvent(db, 'PILOT_MESSAGE_SELECTED', message.id, 'SUCCESS')
  const result = await ingestEmailMessage(db, message)
  await recordIntegrationEvent(
    db,
    'PILOT_INGESTION_COMPLETED',
    message.id,
    'SUCCESS',
    result.status,
  )
  return {
    accessed: true,
    status: result.status,
    complaintId: result.complaintId,
    messageId: message.id,
    conversationId: message.threadId,
    matchCount,
    inspectedCount,
  }
}

// --- Milestone 1 intake seam -------------------------------------------------
// Single intake path through the intelligent-mail foundation. Transport identity is the
// Graph source message (UNIQUE(provider, mailbox_key, provider_message_id)); business
// identity is the extracted Dunkin case ID, which always outranks conversation hints.
// A message with no resolvable identity NEVER fabricates a business case ID: it becomes
// a reviewable provisional intake (mail_review_items, OPEN) for a human to resolve.
// Stage 3b (structured interpretation) exists but is disabled by default and its model
// transport is stubbed: deterministic stages only, no model calls, no live API spend.

const GRAPH_MAILBOX_KEY = 'production'
const INTAKE_LEASE_MS = 5 * 60_000
const NORMALIZATION_VERSION = 'deterministic.v1'
const MAIL_EVENT_VERSION = 'mail-event.v1'

export interface IngestEmailMessageOptions {
  ingestionMode?: IngestionMode
  /** Overrides the interpreter instance (tests inject a fake; default uses the stubbed transport). */
  interpreter?: MailInterpreter
  /** Overrides the settings-table config (tests inject enabled configs directly). */
  interpretationConfig?: MailInterpretationConfig
  /**
   * Overrides the notification service (tests inject a recording transport).
   * Default is built from the settings table and is inert unless
   * external_notifications_enabled='true'.
   */
  notificationService?: NotificationService
  /** Skips the settings-table notification config load when a service is injected. */
  skipNotificationConfigLoad?: boolean
  /**
   * VAPID env for Dad-only new-complaint push (Worker request context).
   * When present, a push is sent to Dad on successful complaint creation.
   * When absent (tests, non-request contexts), no push is attempted.
   */
  pushEnv?: { VAPID_PUBLIC_KEY?: string; VAPID_PRIVATE_KEY?: string; PUBLIC_BASE_URL?: string }
  /**
   * Owner-selected import (preview flow) only. When true and the deterministic
   * gate says NOT a complaint, route to REVIEW_REQUIRED instead of IGNORED, so
   * a genuine forward/reply the gate missed is preserved for human review.
   * Automatic discovery must NOT set this: it keeps IGNORED for non-candidates.
   */
  allowOwnerSelectedReview?: boolean
}

export async function ingestEmailMessage(
  db: D1Database,
  message: NormalizedEmailMessage,
  options: IngestEmailMessageOptions = {},
): Promise<{ status: EmailProcessingStatus; complaintId?: string }> {
  const now = new Date().toISOString()
  const ingestionMode = options.ingestionMode ?? 'LIVE'
  const foundation = new D1MailFoundationRepository(db)
  const interpConfig =
    options.interpretationConfig ?? (await loadMailInterpretationConfig(db))

  // Stage 1 — immutable source persistence. INSERT OR IGNORE on the transport identity:
  // a duplicate delivery can never create a second source row.
  const inserted = await foundation.persistDiscovered({
    id: message.id,
    provider: 'MICROSOFT_GRAPH',
    mailboxKey: GRAPH_MAILBOX_KEY,
    providerMessageId: message.id,
    internetMessageId: message.messageIdHeader,
    conversationId: message.threadId,
    direction: 'INBOUND',
    receivedAt: message.internalDate,
    discoveredAt: now,
    subject: message.subject,
    senderAddress: message.sender,
    recipients: message.recipients
      .split(/[,;]/)
      .map((recipient) => recipient.trim())
      .filter(Boolean),
    ingestionMode,
  })
  if (!inserted) {
    const row = await db
      .prepare('SELECT processing_state,lease_expires_at FROM mail_source_messages WHERE id=?')
      .bind(message.id)
      .first<{ processing_state: string; lease_expires_at: string | null }>()
    const leaseActive =
      row?.processing_state === 'PROCESSING' &&
      row.lease_expires_at !== null &&
      row.lease_expires_at > now
    if (leaseActive) return { status: 'IN_PROGRESS' }
    const resumable =
      row?.processing_state === 'DISCOVERED' ||
      row?.processing_state === 'RETRY_WAIT' ||
      row?.processing_state === 'PROCESSING'
    if (!resumable) {
      const link = await db
        .prepare('SELECT complaint_id FROM complaint_sources WHERE source_message_id=? LIMIT 1')
        .bind(message.id)
        .first<{ complaint_id: string | null }>()
      const legacy = link
        ? null
        : await db
            .prepare('SELECT complaint_id FROM gmail_messages WHERE gmail_message_id=?')
            .bind(message.id)
            .first<{ complaint_id: string | null }>()
      return {
        status: 'DUPLICATE',
        complaintId: link?.complaint_id ?? legacy?.complaint_id ?? undefined,
      }
    }
    // A previous attempt never finished and holds no live lease: resume through the
    // lease acquisition below.
  }

  // Legacy message row: the acknowledgment lookup and pilot observability read this table.
  const priorMessage = await db
    .prepare('SELECT gmail_message_id FROM gmail_messages WHERE gmail_message_id=?')
    .bind(message.id)
    .first<{ gmail_message_id: string }>()
  if (!priorMessage)
    await db
      .prepare(
        `INSERT INTO gmail_messages(gmail_message_id,gmail_thread_id,internal_date,sender,recipients,subject,message_id_header,in_reply_to,references_header,processing_status,first_seen_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,'PROCESSING',?,?)`,
      )
      .bind(
        message.id,
        message.threadId,
        message.internalDate,
        message.sender,
        message.recipients,
        message.subject,
        message.messageIdHeader ?? null,
        message.inReplyTo ?? null,
        message.references ?? null,
        now,
        now,
      )
      .run()
  else
    await db
      .prepare(
        `UPDATE gmail_messages SET processing_status='PROCESSING',processing_detail=NULL,updated_at=? WHERE gmail_message_id=?`,
      )
      .bind(now, message.id)
      .run()

  // Stage 2 — processing lease. Expired leases are resumable; a live lease means another
  // worker owns this message right now.
  const lease = await foundation.acquireProcessingLease(
    message.id,
    'mail-intake',
    now,
    INTAKE_LEASE_MS,
    { normalizationVersion: NORMALIZATION_VERSION },
  )
  if (!lease) return { status: 'IN_PROGRESS' }

  const commitEvent = (
    event: Omit<
      CanonicalEventCommit,
      'eventId' | 'sourceId' | 'processingRunId' | 'eventVersion' | 'ingestionMode' | 'validatedAt'
    >,
  ) =>
    foundation.commitCanonicalEvent({
      ...event,
      eventId: crypto.randomUUID(),
      sourceId: message.id,
      processingRunId: lease.runId,
      eventVersion: MAIL_EVENT_VERSION,
      ingestionMode,
      validatedAt: now,
    })

  try {
    // Stage 3 — deterministic normalization (no model calls in this stage).
    const extraction = extractComplaint(message)
    if (!extraction.isComplaint) {
      // Owner-selected import (preview flow) preserves ambiguous mail for
      // human review instead of silently ignoring it. Automatic discovery
      // keeps the IGNORED path: non-candidates never become review items
      // without an explicit owner selection.
      if (options.allowOwnerSelectedReview) {
        await db
          .prepare(
            `UPDATE gmail_messages SET processing_status='REVIEW_REQUIRED',processing_detail='OWNER_SELECTED_NON_CANDIDATE',processed_at=?,updated_at=? WHERE gmail_message_id=?`,
          )
          .bind(now, now, message.id)
          .run()
        await recordIntegrationEvent(
          db,
          'MESSAGE_REVIEW_REQUIRED',
          message.id,
          'REVIEW_REQUIRED',
          'OWNER_SELECTED_NON_CANDIDATE',
        )
        await commitEvent({
          eventType: 'REVIEW_REQUIRED',
          payload: {
            reason: 'OWNER_SELECTED_NON_CANDIDATE',
            isForwardOrReply: extraction.isForwardOrReply,
            hasCaseIdSignal: extraction.hasCaseIdSignal,
            extractionRules: extraction.extractionRules,
          },
          occurredAt: message.internalDate,
          review: {
            id: crypto.randomUUID(),
            reasonCode: 'OWNER_SELECTED_NON_CANDIDATE',
            disagreementFlags: ['FORWARD_REPLY_OR_AMBIGUOUS'],
          },
        })
        await foundation.finishProcessing(lease, 'REVIEW_REQUIRED', now)
        return { status: 'REVIEW_REQUIRED' }
      }
      await db
        .prepare(
          `UPDATE gmail_messages SET processing_status='IGNORED',processing_detail='NOT_COMPLAINT',processed_at=?,updated_at=? WHERE gmail_message_id=?`,
        )
        .bind(now, now, message.id)
        .run()
      await recordIntegrationEvent(db, 'MESSAGE_IGNORED', message.id, 'IGNORED', 'NOT_COMPLAINT')
      await commitEvent({
        eventType: 'NON_ACTIONABLE',
        payload: { reason: 'NOT_COMPLAINT' },
        occurredAt: message.internalDate,
      })
      await foundation.finishProcessing(lease, 'NON_ACTIONABLE', now)
      return { status: 'IGNORED' }
    }

    // Stage 3b — structured interpretation (shadow / review-only). Skipped entirely
    // unless mail_interpretation_enabled='true'. Architectural cost/privacy boundary:
    // only deterministic complaint candidates reach the model — non-complaint mail
    // returned IGNORED above and can never get here (interpretWithConfig refuses it
    // too, as a second layer). The interpreter can only produce records and review
    // signals: never complaints, merges, mappings, or sends.
    // Model-level failures become ABSTAINED results; the pipeline continues
    // deterministically either way.
    let interpretation: MailInterpretationResult | undefined
    let storeRoute: { storeNumber?: string; reason: InterpretationStoreResolution['reason'] } | undefined
    if (interpConfig.enabled && extraction.isComplaint) {
      storeRoute = await resolveStore(db, extraction)
      const interpretationInput = buildInterpretationInput(
        message,
        extraction,
        storeRoute,
        interpConfig,
        GRAPH_MAILBOX_KEY,
      )
      const interpreter =
        options.interpreter ?? new OpenAiMailInterpreter(interpConfig, new DisabledModelTransport())
      interpretation = await interpretWithConfig(interpreter, interpretationInput, interpConfig)
      if (!interpretation.skipped)
        await recordInterpretation(db, lease.runId, interpretationInput, interpretation)
    }

    // Stage 4 — deterministic reconciliation. Case ID outranks conversation hints, always.
    // Interpretation signals ride along as review-only signals; they can never create
    // or override an identity (see resolveComplaintIdentity).
    const identitySignals: IdentityResolutionSignals = {
      modelSuggestedCaseId: interpretation?.output?.externalCaseId ?? null,
      disagreementCodes: interpretation?.disagreements.map((disagreement) => disagreement.code) ?? [],
    }
    const identity = await foundation.resolveComplaintIdentity(
      extraction.externalCaseId,
      'MICROSOFT_GRAPH',
      GRAPH_MAILBOX_KEY,
      message.threadId,
      identitySignals,
    )

    if (identity.basis === 'EXACT_CASE_ID' || identity.basis === 'CONVERSATION_HINT') {
      const complaintId = identity.complaintId
      const row = await db
        .prepare('SELECT follow_ups FROM complaints WHERE id=?')
        .bind(complaintId)
        .first<{ follow_ups: string }>()
      const followUps = row?.follow_ups ? (JSON.parse(row.follow_ups) as unknown[]) : []
      followUps.push({
        receivedAt: message.internalDate,
        text: extraction.details,
        emailMessageId: message.id,
      })
      await db.batch([
        db
          .prepare('UPDATE complaints SET follow_ups=?,updated_at=? WHERE id=?')
          .bind(JSON.stringify(followUps), now, complaintId),
        db
          .prepare(
            `UPDATE gmail_messages SET complaint_id=?,processing_status='FOLLOW_UP',is_follow_up=1,acknowledgment_status='NOT_APPLICABLE',processed_at=?,updated_at=? WHERE gmail_message_id=?`,
          )
          .bind(complaintId, now, now, message.id),
        db
          .prepare(
            `INSERT INTO complaint_events(id,complaint_id,event_type,actor,timestamp,metadata) VALUES(?,?,'FOLLOW_UP_RECEIVED','microsoft-graph',?,?)`,
          )
          .bind(
            crypto.randomUUID(),
            complaintId,
            now,
            JSON.stringify({ emailMessageId: message.id, provider: 'MICROSOFT_GRAPH' }),
          ),
      ])
      await recordIntegrationEvent(db, 'FOLLOW_UP_INGESTED', message.id, 'SUCCESS')
      await commitEvent({
        eventType: 'FOLLOW_UP',
        complaintId,
        externalCaseId: extraction.externalCaseId,
        payload: { linkageBasis: identity.basis },
        occurredAt: message.internalDate,
        complaintSource: {
          id: crypto.randomUUID(),
          role: 'FOLLOW_UP',
          linkageBasis: identity.basis,
          materialUpdate: true,
        },
      })
      await foundation.finishProcessing(lease, 'COMPLETED', now)
      return { status: 'FOLLOW_UP', complaintId }
    }

    if (identity.basis === 'NEW_CASE_ID' && extraction.externalCaseId) {
      // The case ID came out of the message itself: a real business identity, never fabricated.
      const caseId = extraction.externalCaseId
      const route = storeRoute ?? (await resolveStore(db, extraction))
      const state = await loadState(db)
      const result = createComplaint(
        state,
        {
          externalCaseId: caseId,
          storeNumber: route.storeNumber ?? 'UNROUTED',
          subject: message.subject,
          complaintText: extraction.details || '(No complaint body supplied)',
          category: extraction.category,
          severity: extraction.severity,
        },
        message.internalDate,
        {
          source: 'MICROSOFT_GRAPH',
          actor: 'microsoft-graph',
          acknowledged: false,
          ingestionMode,
        },
      )
      Object.assign(result.complaint, {
        source: 'MICROSOFT_GRAPH',
        gmailMessageId: message.id,
        gmailThreadId: message.threadId,
        sourceSender: message.sender,
        customerName: extraction.customerName,
        customerEmail: extraction.customerEmail,
        customerPhone: extraction.customerPhone,
        occurrenceAt: extraction.occurrenceAt,
        acknowledgementStatus: 'DISABLED',
        routingReason: route.reason,
      } satisfies Partial<Complaint>)
      await persistState(db, result.state)
      // Phase C: critical-escalation notification hook. Inert by default
      // (external_notifications_enabled=false -> notify() is a no-op that never
      // touches a transport). Notifications must never break ingestion.
      if (extraction.severity === 'CRITICAL') {
        try {
          const notifier =
            options.notificationService ??
            new NotificationService(
              options.skipNotificationConfigLoad
                ? { enabled: false, targets: [] }
                : await loadNotificationConfig(db),
              [new LogNotificationTransport()],
            )
          await notifier.notify(
            criticalEscalationEvent({
              complaintId: result.complaint.id,
              severity: extraction.severity,
              summary: `${extraction.category}: ${message.subject}`,
              metadata: {
                externalCaseId: caseId,
                storeNumber: route.storeNumber ?? null,
              },
            }),
          )
        } catch {
          // Notification failures are logged by the transport; ingestion continues.
        }
      }
      const status: EmailProcessingStatus = result.complaint.storeId
        ? 'PROCESSED'
        : 'ROUTING_REVIEW'
      await db
        .prepare(
          `UPDATE gmail_messages SET complaint_id=?,processing_status=?,processing_detail=?,acknowledgment_status='PENDING',processed_at=?,updated_at=? WHERE gmail_message_id=?`,
        )
        .bind(result.complaint.id, status, route.reason, now, now, message.id)
        .run()
      await recordIntegrationEvent(db, 'MESSAGE_INGESTED', message.id, 'SUCCESS', status)
      await commitEvent({
        eventType: 'NEW_CASE',
        complaintId: result.complaint.id,
        externalCaseId: caseId,
        canonicalStoreNumber: route.storeNumber,
        payload: {
          linkageBasis: 'NEW_CASE_ID',
          routingReason: route.reason,
          category: extraction.category,
          severity: extraction.severity,
        },
        occurredAt: message.internalDate,
        complaintSource: {
          id: crypto.randomUUID(),
          role: 'INITIAL',
          linkageBasis: 'NEW_CASE_ID',
          materialUpdate: true,
        },
      })
      await foundation.finishProcessing(lease, 'COMPLETED', now)
      // Dad-only pilot: push Dad a notification for the new complaint.
      // Fire-and-forget — never breaks ingestion. Only runs when the
      // caller passes push env (Worker request context).
      if (options.pushEnv) {
        try {
          const { notifyDadNewComplaint } = await import('./dad-push')
          await notifyDadNewComplaint(db, options.pushEnv, {
            id: result.complaint.id,
            externalCaseId: caseId,
            category: extraction.category,
            storeNumber: route.storeNumber ?? undefined,
          })
        } catch {
          // Push failures are audited inside notifyDadNewComplaint; ingestion continues.
        }
      }
      return { status, complaintId: result.complaint.id }
    }

    // No trustworthy external case ID and no conversation link: reviewable provisional
    // intake. No complaint is created and no business identity is invented; a human
    // resolves the identity from the review item. A model-suggested case ID, if any,
    // rides along as an explicitly unverified hint — never as an identity.
    const reviewSignals = identity.reviewSignals
    await db
      .prepare(
        `UPDATE gmail_messages SET processing_status='REVIEW_REQUIRED',processing_detail='IDENTITY_UNRESOLVED',processed_at=?,updated_at=? WHERE gmail_message_id=?`,
      )
      .bind(now, now, message.id)
      .run()
    await recordIntegrationEvent(
      db,
      'MESSAGE_INGESTION_REVIEW',
      message.id,
      'REVIEW',
      'IDENTITY_UNRESOLVED',
    )
    await commitEvent({
      eventType: 'REVIEW_REQUIRED',
      payload: {
        reason: 'IDENTITY_UNRESOLVED',
        subject: message.subject,
        sender: message.sender,
        ...(reviewSignals?.suggestedCaseId
          ? {
              modelSuggestedCaseId: reviewSignals.suggestedCaseId,
              modelSuggestedCaseIdNote: 'UNVERIFIED_MODEL_SUGGESTION_NOT_IDENTITY',
            }
          : {}),
      },
      occurredAt: message.internalDate,
      review: {
        id: crypto.randomUUID(),
        reasonCode: 'IDENTITY_UNRESOLVED',
        disagreementFlags: reviewSignals?.disagreementCodes ?? [],
      },
    })
    await foundation.finishProcessing(lease, 'REVIEW_REQUIRED', now)
    return { status: 'REVIEW_REQUIRED' }
  } catch (error) {
    // Store the message (truncated), not just the name — "Error" alone is
    // undiagnosable. Oct 17: all 6 imports failed with detail "Error",
    // hiding the real cause (null store managerId in createComplaint).
    const detail =
      error instanceof Error
        ? `${error.name}: ${error.message}`.slice(0, 500)
        : 'UNKNOWN'
    await db
      .prepare(
        `UPDATE gmail_messages SET processing_status='FAILED_PERSISTENCE',processing_detail=?,updated_at=? WHERE gmail_message_id=?`,
      )
      .bind(detail, now, message.id)
      .run()
    await recordIntegrationEvent(
      db,
      'MESSAGE_INGESTION_FAILED',
      message.id,
      'FAILED',
      'FAILED_PERSISTENCE',
    )
    await foundation.finishProcessing(lease, 'RETRY_WAIT', now, detail)
    throw error
  }
}

export const neutralAcknowledgment =
  'We have received this complaint and are reviewing it with the appropriate location. We will follow up as needed.'

export type AcknowledgeResult =
  | 'DISABLED'
  | 'SENT'
  | 'ALREADY_HANDLED'
  | 'FAILED'
  | 'AMBIGUOUS'
  | 'IN_FLIGHT_STALE'
  | 'MAILBOX_NOT_VERIFIED'
  | 'NO_SOURCE'
  | 'WRONG_RECIPIENT'
  | 'FORWARD_NOT_ACKABLE'

export interface AcknowledgeOptions {
  /** Manual pilot: explicit owner action. Does NOT require emailAckEnabled. */
  manual?: boolean
  /** Allow retry from FAILED (manual only). Automatic path keeps FAILED terminal. */
  allowRetryFromFailed?: boolean
  /** Caller must have verified mailbox identity for manual sends. */
  mailboxIdentityVerified?: boolean
}

const STALE_IN_FLIGHT_MS = 15 * 60_000

/** True when a send error is ambiguous (Graph may have accepted). Never auto-retry these. */
function isAmbiguousSendError(error: unknown): boolean {
  if (error instanceof EmailProviderError) {
    // 429 / 5xx: Graph may have accepted before failing. Timeout / network: same.
    if (error.status === 429) return true
    if (error.status !== undefined && error.status >= 500) return true
    if (error.code.includes('TIMEOUT') || error.code.includes('TEMPORARY')) return true
    if (error.code.includes('RETRY_EXHAUSTED')) return true
    return false
  }
  const message = error instanceof Error ? error.message.toLowerCase() : ''
  return (
    message.includes('timeout') ||
    message.includes('timed out') ||
    message.includes('network') ||
    message.includes('fetch failed') ||
    message.includes('aborted') ||
    message.includes('econnreset') ||
    message.includes('socket')
  )
}

export async function acknowledgeComplaint(
  db: D1Database,
  emailProvider: EmailProvider,
  complaintId: string,
  config: AppConfig,
  options: AcknowledgeOptions = {},
): Promise<AcknowledgeResult> {
  // Manual pilot: explicit owner action does NOT require the global automatic
  // switch, but DOES require mailbox identity verification. Automatic path
  // requires the switch and never retries FAILED.
  if (options.manual) {
    if (!options.mailboxIdentityVerified) return 'MAILBOX_NOT_VERIFIED'
  } else if (!config.emailAckEnabled) {
    return 'DISABLED'
  }
  const source = await db
    .prepare(
      `SELECT gm.*,c.acknowledgement_status FROM gmail_messages gm JOIN complaints c ON c.id=gm.complaint_id WHERE gm.complaint_id=? AND gm.is_follow_up=0 ORDER BY gm.first_seen_at LIMIT 1`,
    )
    .bind(complaintId)
    .first<Record<string, unknown>>()
  if (!source) return 'NO_SOURCE'
  // Recipient safety: Only ever ack the verified corporate source message.
  // Graph /reply replies to the sender of the source message. If the source
  // is a forward from a family member (e.g. Dad forwarding a Dunkin case)
  // or a customer, /reply would go to that person, NOT to Dunkin — a
  // wrong-recipient send. Refuse to send; import the original Dunkin
  // message instead.
  const senderRaw = String(source.sender ?? '')
  const senderEmail = (senderRaw.match(/<([^>]+)>/)?.[1] ?? senderRaw.trim()).toLowerCase()
  const senderDomain = senderEmail.split('@')[1] ?? ''
  const personalDomains = ['yahoo.com', 'gmail.com', 'aol.com', 'outlook.com', 'hotmail.com', 'icloud.com', 'me.com', 'mac.com']
  const isPersonal = personalDomains.includes(senderDomain)
  // Dunkin corporate: @dunkinbrands.com, OR display name contains Dunkin
  // and email is NOT a personal domain (covers test fixtures + any future
  // Dunkin sender alias). Family members (yahoo/gmail) and customers
  // (personal domains) are never Dunkin corporate, even if they put
  // "Dunkin" in their display name.
  const isDunkinCorporate =
    senderEmail.endsWith('@dunkinbrands.com') ||
    (!isPersonal && senderRaw.toLowerCase().includes('dunkin'))
  if (!isDunkinCorporate) return 'WRONG_RECIPIENT'
  const subjectRaw = String(source.subject ?? '').trim()
  if (/^(fwd|fw|re)\s*:/i.test(subjectRaw)) return 'FORWARD_NOT_ACKABLE'
  const idempotencyKey = `email-ack:${complaintId}`
  const now = new Date().toISOString()
  await db
    .prepare(
      `INSERT OR IGNORE INTO email_acknowledgments(id,complaint_id,gmail_thread_id,source_gmail_message_id,idempotency_key,status,created_at,updated_at) VALUES(?,?,?,?,?,'PENDING',?,?)`,
    )
    .bind(
      crypto.randomUUID(),
      complaintId,
      source.gmail_thread_id,
      source.gmail_message_id,
      idempotencyKey,
      now,
      now,
    )
    .run()
  const acknowledgment = await db
    .prepare('SELECT id,status,updated_at,attempt_count,last_error_code FROM email_acknowledgments WHERE complaint_id=?')
    .bind(complaintId)
    .first<{ id: string; status: string; updated_at: string; attempt_count: number; last_error_code: string | null }>()
  if (!acknowledgment) return 'ALREADY_HANDLED'
  // SENT: never send again.
  if (acknowledgment.status === 'SENT') return 'ALREADY_HANDLED'
  // IN_FLIGHT: fresh = another worker owns it; stale = caller must reclaim manually.
  if (acknowledgment.status === 'IN_FLIGHT') {
    const ageMs = Date.parse(now) - Date.parse(acknowledgment.updated_at)
    if (ageMs < STALE_IN_FLIGHT_MS) return 'ALREADY_HANDLED'
    return 'IN_FLIGHT_STALE'
  }
  // FAILED: automatic path keeps it terminal. Manual path may retry with explicit approval.
  if (acknowledgment.status === 'FAILED') {
    if (!options.manual || !options.allowRetryFromFailed) return 'ALREADY_HANDLED'
    // Reset to PENDING, preserving attempt_count (it increments on next IN_FLIGHT).
    await db
      .prepare(`UPDATE email_acknowledgments SET status='PENDING',last_error_code=NULL,updated_at=? WHERE id=? AND status='FAILED'`)
      .bind(now, acknowledgment.id)
      .run()
    acknowledgment.status = 'PENDING'
  }
  if (acknowledgment.status !== 'PENDING') return 'ALREADY_HANDLED'
  await db
    .prepare(
      `UPDATE email_acknowledgments SET status='IN_FLIGHT',attempt_count=attempt_count+1,updated_at=? WHERE id=? AND status='PENDING'`,
    )
    .bind(now, acknowledgment.id)
    .run()
  const message: NormalizedEmailMessage = {
    id: String(source.gmail_message_id),
    threadId: String(source.gmail_thread_id),
    internalDate: String(source.internal_date),
    sender: String(source.sender),
    recipients: String(source.recipients),
    subject: String(source.subject),
    messageIdHeader: source.message_id_header ? String(source.message_id_header) : undefined,
    inReplyTo: source.in_reply_to ? String(source.in_reply_to) : undefined,
    references: source.references_header ? String(source.references_header) : undefined,
    textBody: '',
  }
  // Separate Graph send from D1 commit: if Graph accepts (202) but D1 commit
  // fails, the row must stay IN_FLIGHT (ambiguous), not FAILED — otherwise a
  // retry would send a second reply.
  let providerMessageId: string | undefined
  try {
    providerMessageId = await emailProvider.sendAcknowledgment(message, neutralAcknowledgment)
  } catch (error) {
    if (isAmbiguousSendError(error)) {
      // Ambiguous: Graph may have accepted. Keep IN_FLIGHT, do NOT auto-retry.
      // Owner must check Sent Items and reclaim manually.
      const code = error instanceof EmailProviderError ? error.code : 'AMBIGUOUS_SEND'
      const ambiguousAt = new Date().toISOString()
      await db
        .prepare(`UPDATE email_acknowledgments SET last_error_code=?,updated_at=? WHERE id=?`)
        .bind(`AMBIGUOUS:${code}`, ambiguousAt, acknowledgment.id)
        .run()
      await recordIntegrationEvent(db, 'ACKNOWLEDGMENT_AMBIGUOUS', complaintId, 'AMBIGUOUS', code)
      return 'AMBIGUOUS'
    }
    const code = error instanceof EmailProviderError ? error.code : 'MS_GRAPH_SEND_UNKNOWN'
    const failedAt = new Date().toISOString()
    await db.batch([
      db
        .prepare(
          `UPDATE email_acknowledgments SET status='FAILED',last_error_code=?,updated_at=? WHERE id=?`,
        )
        .bind(code, failedAt, acknowledgment.id),
      db
        .prepare(`UPDATE complaints SET acknowledgement_status='FAILED',updated_at=? WHERE id=?`)
        .bind(failedAt, complaintId),
    ])
    await recordIntegrationEvent(db, 'ACKNOWLEDGMENT_FAILED', complaintId, 'FAILED', code)
    return 'FAILED'
  }
  try {
    const sentAt = new Date().toISOString()
    await db.batch([
      db
        .prepare(
          `UPDATE email_acknowledgments SET status='SENT',provider_message_id=?,sent_at=?,updated_at=? WHERE id=? AND status='IN_FLIGHT'`,
        )
        .bind(providerMessageId ?? null, sentAt, sentAt, acknowledgment.id),
      db
        .prepare(
          `UPDATE complaints SET dunkin_acknowledged_at=?,acknowledgement_status='SENT',acknowledgment_body=?,updated_at=? WHERE id=?`,
        )
        .bind(sentAt, neutralAcknowledgment, sentAt, complaintId),
      db
        .prepare(
          `UPDATE gmail_messages SET acknowledgment_status='SENT',updated_at=? WHERE gmail_message_id=?`,
        )
        .bind(sentAt, message.id),
      db
        .prepare(
          `INSERT INTO complaint_events(id,complaint_id,event_type,actor,timestamp,metadata) VALUES(?,?,'DUNKIN_ACKNOWLEDGED','microsoft-graph',?,?)`,
        )
        .bind(
          crypto.randomUUID(),
          complaintId,
          sentAt,
          JSON.stringify({ provider: 'MICROSOFT_GRAPH', providerMessageId, manual: options.manual ?? false }),
        ),
    ])
    await recordIntegrationEvent(db, 'ACKNOWLEDGMENT_SENT', complaintId, 'SUCCESS')
    return 'SENT'
  } catch {
    // Graph accepted (202) but D1 commit failed: ambiguous, stay IN_FLIGHT.
    // Never mark FAILED here — a retry would send a duplicate reply.
    const failedAt = new Date().toISOString()
    await db
      .prepare(`UPDATE email_acknowledgments SET last_error_code=?,updated_at=? WHERE id=?`)
      .bind('AMBIGUOUS:COMMIT_FAILED', failedAt, acknowledgment.id)
      .run()
    await recordIntegrationEvent(db, 'ACKNOWLEDGMENT_AMBIGUOUS', complaintId, 'AMBIGUOUS', 'COMMIT_FAILED')
    return 'AMBIGUOUS'
  }
}

export type ReclaimAckResult =
  | 'RECLAIMED_SENT'
  | 'RECLAIMED_TO_PENDING'
  | 'NOT_IN_FLIGHT'
  | 'IN_FLIGHT_FRESH'
  | 'NEEDS_SENT_ITEMS_CHECK'
  | 'NEEDS_FURTHER_RECONCILIATION'

/**
 * Manual reclaim for stale IN_FLIGHT acks. Ambiguous sends stay in a
 * non-retryable pending-reconciliation state.
 * - sentConfirmed=true: mark SENT without sending again (no duplicate reply).
 * - sentConfirmed=false: DO NOT reset. Absence from Sent Items alone is NOT
 *   conclusive proof the message wasn't sent (Graph may have accepted but
 *   Sent Items hasn't updated, reply may be in another folder, check may
 *   be incomplete). Returns NEEDS_FURTHER_RECONCILIATION and keeps
 *   IN_FLIGHT. Only an explicit forceReset with a written reconciliation
 *   note (what was checked, how, and why it is conclusive) resets to
 *   PENDING for a manual retry — and that override is audited.
 * Never auto-reclaims, never sends.
 */
export async function reclaimStaleAck(
  db: D1Database,
  complaintId: string,
  options: {
    checkedSentItems: boolean
    sentConfirmed: boolean
    forceReset?: boolean
    reconciliationNote?: string
  },
): Promise<ReclaimAckResult> {
  if (!options.checkedSentItems) return 'NEEDS_SENT_ITEMS_CHECK'
  const row = await db
    .prepare('SELECT id,status,updated_at FROM email_acknowledgments WHERE complaint_id=?')
    .bind(complaintId)
    .first<{ id: string; status: string; updated_at: string }>()
  if (!row || row.status !== 'IN_FLIGHT') return 'NOT_IN_FLIGHT'
  const ageMs = Date.parse(new Date().toISOString()) - Date.parse(row.updated_at)
  if (ageMs < STALE_IN_FLIGHT_MS) return 'IN_FLIGHT_FRESH'
  const now = new Date().toISOString()
  if (options.sentConfirmed) {
    const source = await db
      .prepare(`SELECT gmail_message_id FROM gmail_messages WHERE complaint_id=? AND is_follow_up=0 ORDER BY first_seen_at LIMIT 1`)
      .bind(complaintId)
      .first<{ gmail_message_id: string }>()
    await db.batch([
      db
        .prepare(`UPDATE email_acknowledgments SET status='SENT',sent_at=?,updated_at=? WHERE id=? AND status='IN_FLIGHT'`)
        .bind(now, now, row.id),
      db
        .prepare(`UPDATE complaints SET dunkin_acknowledged_at=?,acknowledgement_status='SENT',acknowledgment_body=?,updated_at=? WHERE id=?`)
        .bind(now, neutralAcknowledgment, now, complaintId),
      ...(source
        ? [db.prepare(`UPDATE gmail_messages SET acknowledgment_status='SENT',updated_at=? WHERE gmail_message_id=?`).bind(now, source.gmail_message_id)]
        : []),
      db
        .prepare(`INSERT INTO complaint_events(id,complaint_id,event_type,actor,timestamp,metadata) VALUES(?,?,'DUNKIN_ACKNOWLEDGED','microsoft-graph',?,?)`)
        .bind(crypto.randomUUID(), complaintId, now, JSON.stringify({ reclaimed: true, via: 'SENT_ITEMS_CHECK' })),
    ])
    await recordIntegrationEvent(db, 'ACKNOWLEDGMENT_RECLAIMED_SENT', complaintId, 'SUCCESS')
    return 'RECLAIMED_SENT'
  }
  // sentConfirmed=false: stay IN_FLIGHT. Absence from Sent Items is not
  // conclusive. Only an explicit, audited forceReset with a note resets.
  if (!options.forceReset) {
    await db
      .prepare(`UPDATE email_acknowledgments SET last_error_code=?,updated_at=? WHERE id=?`)
      .bind('AMBIGUOUS:NEEDS_FURTHER_RECONCILIATION', now, row.id)
      .run()
    return 'NEEDS_FURTHER_RECONCILIATION'
  }
  if (!options.reconciliationNote?.trim()) return 'NEEDS_FURTHER_RECONCILIATION'
  await db.batch([
    db.prepare(`UPDATE email_acknowledgments SET status='PENDING',last_error_code=NULL,updated_at=? WHERE id=? AND status='IN_FLIGHT'`).bind(now, row.id),
    db.prepare(`UPDATE complaints SET acknowledgement_status='PENDING',updated_at=? WHERE id=?`).bind(now, complaintId),
  ])
  await recordIntegrationEvent(db, 'ACKNOWLEDGMENT_RECLAIMED_PENDING', complaintId, 'SUCCESS', `forceReset: ${options.reconciliationNote.trim().slice(0, 500)}`)
  return 'RECLAIMED_TO_PENDING'
}

export async function pollEmail(
  db: D1Database,
  emailProvider: EmailProvider,
  config: AppConfig,
): Promise<{ processed: number; failures: number }> {
  if (!config.emailIngestionEnabled || !emailProvider.ready) return { processed: 0, failures: 0 }
  await emailProvider.verifyConnection()
  const ids = await emailProvider.listMessageIds(config.emailLookbackDays ?? 30)
  let processed = 0
  let failures = 0
  for (const id of ids.reverse()) {
    try {
      const message = await emailProvider.getMessage(id)
      const result = await ingestEmailMessage(db, message)
      if (!['DUPLICATE', 'IN_PROGRESS'].includes(result.status)) processed += 1
      if (result.complaintId && config.emailAckEnabled)
        await acknowledgeComplaint(db, emailProvider, result.complaintId, config)
    } catch {
      failures += 1
    }
  }
  return { processed, failures }
}

// Dormant Gmail adapter compatibility; production uses Microsoft Graph through EmailProvider.
export const ingestGmailMessage = ingestEmailMessage
