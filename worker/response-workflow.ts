// Phase C — ack/response workflow scaffolding.
//
// Template storage (response_templates, new in migration 0007) plus an
// approval-gated lifecycle on the existing response_actions table:
//
//   DRAFT -> PENDING_APPROVAL -> APPROVED -> (send)
//                                        \-> REJECTED
//
// Sending is disabled in this phase, unconditionally:
//   - email_ack_enabled=false  -> SEND_DISABLED
//   - email_ack_enabled=true   -> SEND_NOT_WIRED (no transport exists yet)
// No external send can happen from this module. Ever, in Phase C.
//
// The identity rule is unaffected: drafts are human-authored from templates,
// never model-generated, and approving a response never creates or merges a
// complaint.

import type { D1Database } from './d1'

export const RESPONSE_POLICY_VERSION = 'response-policy.v1'

export type ResponseActionPolicyState =
  | 'DRAFT'
  | 'PENDING_APPROVAL'
  | 'APPROVED'
  | 'REJECTED'
  | 'SEND_BLOCKED'
  | 'SENT'
  | 'SENDING'
  | 'SEND_FAILED'
  | 'NEEDS_RECONCILIATION'

export class ResponseWorkflowError extends Error {
  constructor(
    readonly code: 'TEMPLATE_NOT_FOUND' | 'COMPLAINT_NOT_FOUND' | 'INVALID_TRANSITION' | 'SEND_DISABLED' | 'SEND_NOT_WIRED' | 'ACTION_NOT_FOUND' | 'NOT_APPROVED' | 'WRONG_RECIPIENT' | 'NO_SOURCE' | 'ALREADY_SENT' | 'MAILBOX_NOT_VERIFIED' | 'AMBIGUOUS_SEND',
    message: string,
  ) {
    super(message)
    this.name = 'ResponseWorkflowError'
  }
}

export interface ResponseTemplate {
  id: string
  name: string
  responseKind: string
  subjectTemplate?: string
  bodyTemplate: string
  version: number
  active: boolean
  createdAt: string
  updatedAt: string
}

export interface ResponseAction {
  id: string
  complaintId: string
  responseKind: string
  policyVersion: string
  policyState: ResponseActionPolicyState
  contentKind?: string
  contentRef?: string
  templateVersion?: string
  approvedBy?: string
  approvedAt?: string
  createdAt: string
  updatedAt: string
}

/**
 * Render a template with {{variable}} placeholders. Unknown variables are
 * left in place (visible in review) rather than silently dropped.
 */
export function renderTemplate(template: string, vars: Record<string, unknown>): string {
  return template.replace(/\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}/g, (match, key: string) =>
    Object.prototype.hasOwnProperty.call(vars, key) && vars[key] != null ? String(vars[key]) : match,
  )
}

const toTemplate = (row: Record<string, unknown>): ResponseTemplate => ({
  id: String(row.id),
  name: String(row.name),
  responseKind: String(row.response_kind),
  subjectTemplate: row.subject_template ? String(row.subject_template) : undefined,
  bodyTemplate: String(row.body_template),
  version: Number(row.version),
  active: Number(row.active) === 1,
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at),
})

const toAction = (row: Record<string, unknown>): ResponseAction => ({
  id: String(row.id),
  complaintId: String(row.complaint_id),
  responseKind: String(row.response_kind),
  policyVersion: String(row.policy_version),
  policyState: String(row.policy_state) as ResponseActionPolicyState,
  contentKind: row.content_kind ? String(row.content_kind) : undefined,
  contentRef: row.content_ref ? String(row.content_ref) : undefined,
  templateVersion: row.template_version ? String(row.template_version) : undefined,
  approvedBy: row.approved_by ? String(row.approved_by) : undefined,
  approvedAt: row.approved_at ? String(row.approved_at) : undefined,
  createdAt: String(row.created_at),
  updatedAt: String(row.updated_at),
})

export async function createResponseTemplate(
  db: D1Database,
  input: { name: string; responseKind: string; subjectTemplate?: string; bodyTemplate: string },
): Promise<ResponseTemplate> {
  const now = new Date().toISOString()
  const id = crypto.randomUUID()
  await db
    .prepare(
      `INSERT INTO response_templates(id,name,response_kind,subject_template,body_template,version,active,created_at,updated_at)
       VALUES(?,?,?,?,?,1,1,?,?)`,
    )
    .bind(id, input.name, input.responseKind, input.subjectTemplate ?? null, input.bodyTemplate, now, now)
    .run()
  const row = await db.prepare('SELECT * FROM response_templates WHERE id=?').bind(id).first<Record<string, unknown>>()
  if (!row) throw new ResponseWorkflowError('TEMPLATE_NOT_FOUND', 'template insert failed')
  return toTemplate(row)
}

export async function listResponseTemplates(db: D1Database): Promise<ResponseTemplate[]> {
  const rows = await db
    .prepare('SELECT * FROM response_templates WHERE active=1 ORDER BY name')
    .all<Record<string, unknown>>()
  return rows.results.map(toTemplate)
}

export async function getResponseTemplate(db: D1Database, id: string): Promise<ResponseTemplate | null> {
  const row = await db
    .prepare('SELECT * FROM response_templates WHERE id=?')
    .bind(id)
    .first<Record<string, unknown>>()
  return row ? toTemplate(row) : null
}

/**
 * Create a DRAFT response action for a complaint. Body comes from a template
 * (rendered with the given vars) or is supplied directly. Never sends.
 */
export async function createResponseDraft(
  db: D1Database,
  input: {
    complaintId: string
    responseKind?: string
    templateId?: string
    templateVars?: Record<string, unknown>
    body?: string
    createdBy?: string
  },
): Promise<ResponseAction> {
  const complaint = await db
    .prepare('SELECT id FROM complaints WHERE id=?')
    .bind(input.complaintId)
    .first<{ id: string }>()
  if (!complaint)
    throw new ResponseWorkflowError('COMPLAINT_NOT_FOUND', `complaint ${input.complaintId} not found`)

  let body = input.body ?? ''
  let templateVersion: string | undefined
  if (input.templateId) {
    const template = await getResponseTemplate(db, input.templateId)
    if (!template || !template.active)
      throw new ResponseWorkflowError('TEMPLATE_NOT_FOUND', `template ${input.templateId} not found`)
    body = renderTemplate(template.bodyTemplate, input.templateVars ?? {})
    templateVersion = `response-template.v${template.version}`
  }
  const now = new Date().toISOString()
  const id = crypto.randomUUID()
  await db
    .prepare(
      `INSERT INTO response_actions(
         id,complaint_id,response_kind,policy_version,policy_state,
         content_kind,content_ref,template_version,created_at,updated_at
       ) VALUES(?,?,?,?,?,?,?,?,?,?)`,
    )
    .bind(
      id,
      input.complaintId,
      input.responseKind ?? 'ACKNOWLEDGMENT',
      RESPONSE_POLICY_VERSION,
      'DRAFT',
      'draft-body',
      body,
      templateVersion ?? null,
      now,
      now,
    )
    .run()
  const row = await db.prepare('SELECT * FROM response_actions WHERE id=?').bind(id).first<Record<string, unknown>>()
  if (!row) throw new ResponseWorkflowError('ACTION_NOT_FOUND', 'draft insert failed')
  return toAction(row)
}

async function transitionState(
  db: D1Database,
  id: string,
  from: ResponseActionPolicyState[],
  to: ResponseActionPolicyState,
  extra?: { approvedBy?: string },
): Promise<ResponseAction> {
  const now = new Date().toISOString()
  const row = await db
    .prepare('SELECT * FROM response_actions WHERE id=?')
    .bind(id)
    .first<Record<string, unknown>>()
  if (!row) throw new ResponseWorkflowError('ACTION_NOT_FOUND', `response action ${id} not found`)
  const current = String(row.policy_state) as ResponseActionPolicyState
  if (!from.includes(current))
    throw new ResponseWorkflowError(
      'INVALID_TRANSITION',
      `cannot transition response action ${id} from ${current} to ${to}`,
    )
  await db
    .prepare(
      `UPDATE response_actions SET policy_state=?,approved_by=COALESCE(?,approved_by),
       approved_at=CASE WHEN ? IS NOT NULL THEN ? ELSE approved_at END, updated_at=?
       WHERE id=? AND policy_state=?`,
    )
    .bind(to, extra?.approvedBy ?? null, extra?.approvedBy ?? null, now, now, id, current)
    .run()
  const updated = await db
    .prepare('SELECT * FROM response_actions WHERE id=?')
    .bind(id)
    .first<Record<string, unknown>>()
  return toAction(updated!)
}

export const submitResponseForApproval = (db: D1Database, id: string): Promise<ResponseAction> =>
  transitionState(db, id, ['DRAFT'], 'PENDING_APPROVAL')

export const approveResponseAction = (
  db: D1Database,
  input: { id: string; approvedBy: string },
): Promise<ResponseAction> =>
  transitionState(db, input.id, ['PENDING_APPROVAL'], 'APPROVED', { approvedBy: input.approvedBy })

export const rejectResponseAction = (
  db: D1Database,
  input: { id: string; rejectedBy: string },
): Promise<ResponseAction> =>
  transitionState(db, input.id, ['PENDING_APPROVAL'], 'REJECTED', { approvedBy: input.rejectedBy })

/**
 * Dad-approved response send. Gated by explicit APPROVED state, NOT by the
 * global email_ack_enabled switch (which stays OFF — automatic acks are
 * separate from Dad-approved sends).
 *
 * Safety rules (mirroring acknowledgeComplaint):
 * - Action must be APPROVED (Dad explicitly approved this exact draft).
 * - Mailbox identity must be verified via Graph /me before sending.
 * - Recipient safety: only replies to verified Dunkin corporate source
 *   messages. Forwards and personal-domain senders are refused.
 * - Idempotency: SENT actions never re-send. IN_FLIGHT/ambiguous results
 *   become NEEDS_RECONCILIATION — never auto-retried, owner must check
 *   Sent Items and reclaim manually.
 * - The approved draft body (content_ref) is sent verbatim via Graph /reply
 *   into the original conversation thread.
 */
export async function sendResponseAction(
  db: D1Database,
  input: {
    id: string
    emailProvider: {
      sendAcknowledgment: (
        message: { id: string; threadId: string; internalDate: string; sender: string; recipients: string; subject: string; textBody: string },
        body: string,
      ) => Promise<unknown>
    }
    mailboxIdentityVerified: boolean
  },
): Promise<ResponseAction> {
  const row = await db
    .prepare('SELECT * FROM response_actions WHERE id=?')
    .bind(input.id)
    .first<Record<string, unknown>>()
  if (!row) throw new ResponseWorkflowError('ACTION_NOT_FOUND', `response action ${input.id} not found`)
  const action = toAction(row)
  if (action.policyState === 'SENT')
    throw new ResponseWorkflowError('ALREADY_SENT', `response action ${input.id} was already sent`)
  if (action.policyState === 'NEEDS_RECONCILIATION')
    throw new ResponseWorkflowError(
      'AMBIGUOUS_SEND',
      `response action ${input.id} has an uncertain send outcome — check Sent Items before retrying`,
    )
  if (action.policyState !== 'APPROVED')
    throw new ResponseWorkflowError(
      'NOT_APPROVED',
      `response action ${input.id} is ${action.policyState}, not APPROVED`,
    )
  if (!input.mailboxIdentityVerified)
    throw new ResponseWorkflowError('MAILBOX_NOT_VERIFIED', 'mailbox identity must be verified before sending')

  // Load the complaint's original source message for recipient safety + thread.
  // All validation happens BEFORE the atomic claim, so a refused send never
  // leaves the action stuck in SENDING.
  const source = await db
    .prepare(
      `SELECT gm.* FROM gmail_messages gm WHERE gm.complaint_id=? AND gm.is_follow_up=0 ORDER BY gm.first_seen_at LIMIT 1`,
    )
    .bind(action.complaintId)
    .first<Record<string, unknown>>()
  if (!source) throw new ResponseWorkflowError('NO_SOURCE', `complaint ${action.complaintId} has no source message`)

  // Recipient safety: same rules as acknowledgeComplaint.
  const senderRaw = String(source.sender ?? '')
  const senderEmail = (senderRaw.match(/<([^>]+)>/)?.[1] ?? senderRaw.trim()).toLowerCase()
  const senderDomain = senderEmail.split('@')[1] ?? ''
  const personalDomains = ['yahoo.com', 'gmail.com', 'aol.com', 'outlook.com', 'hotmail.com', 'icloud.com', 'me.com', 'mac.com']
  const isPersonal = personalDomains.includes(senderDomain)
  const isDunkinCorporate =
    senderEmail.endsWith('@dunkinbrands.com') ||
    (!isPersonal && senderRaw.toLowerCase().includes('dunkin'))
  if (!isDunkinCorporate)
    throw new ResponseWorkflowError('WRONG_RECIPIENT', 'source message is not from Dunkin corporate')
  const subjectRaw = String(source.subject ?? '').trim()
  if (/^(fwd|fw|re)\s*:/i.test(subjectRaw))
    throw new ResponseWorkflowError('WRONG_RECIPIENT', 'source message is a forward or reply, not an original Dunkin message')

  const body = action.contentRef ?? ''
  if (!body.trim())
    throw new ResponseWorkflowError('INVALID_TRANSITION', 'approved response has no body to send')

  // ATOMIC CLAIM: transition APPROVED -> SENDING in a single conditional
  // UPDATE. Only one caller can win the claim; repeated taps or concurrent
  // requests that lose the race see the action is no longer APPROVED and
  // refuse to send. This is the duplicate-send guard.
  const now = new Date().toISOString()
  const claim = await db
    .prepare(
      `UPDATE response_actions SET policy_state='SENDING',updated_at=? WHERE id=? AND policy_state='APPROVED'`,
    )
    .bind(now, input.id)
    .run()
  if ((claim.meta?.changes ?? 0) === 0) {
    // Lost the race (or state changed). Re-read to report accurately.
    const current = await db
      .prepare('SELECT policy_state FROM response_actions WHERE id=?')
      .bind(input.id)
      .first<{ policy_state: string }>()
    const state = current?.policy_state ?? 'UNKNOWN'
    if (state === 'SENT' || state === 'SENDING')
      throw new ResponseWorkflowError('ALREADY_SENT', `response action ${input.id} is already ${state}`)
    if (state === 'NEEDS_RECONCILIATION')
      throw new ResponseWorkflowError(
        'AMBIGUOUS_SEND',
        `response action ${input.id} has an uncertain send outcome — check Sent Items before retrying`,
      )
    throw new ResponseWorkflowError('NOT_APPROVED', `response action ${input.id} is ${state}, not APPROVED`)
  }

  const message = {
    id: String(source.gmail_message_id),
    threadId: String(source.gmail_thread_id),
    internalDate: String(source.internal_date),
    sender: String(source.sender),
    recipients: String(source.recipients),
    subject: String(source.subject),
    textBody: '',
  }

  try {
    await input.emailProvider.sendAcknowledgment(message, body)
  } catch (error) {
    // Ambiguous: Graph may have accepted. Mark NEEDS_RECONCILIATION, never auto-retry.
    const code = error instanceof Error ? error.message : 'AMBIGUOUS_SEND'
    await db
      .prepare(`UPDATE response_actions SET policy_state='NEEDS_RECONCILIATION',updated_at=? WHERE id=?`)
      .bind(now, input.id)
      .run()
    throw new ResponseWorkflowError('AMBIGUOUS_SEND', `send outcome uncertain: ${code}`)
  }

  // Success: mark SENT.
  await db
    .prepare(`UPDATE response_actions SET policy_state='SENT',updated_at=? WHERE id=?`)
    .bind(now, input.id)
    .run()
  const updated = await db
    .prepare('SELECT * FROM response_actions WHERE id=?')
    .bind(input.id)
    .first<Record<string, unknown>>()
  return toAction(updated!)
}

export async function getResponseAction(db: D1Database, id: string): Promise<ResponseAction | null> {
  const row = await db
    .prepare('SELECT * FROM response_actions WHERE id=?')
    .bind(id)
    .first<Record<string, unknown>>()
  return row ? toAction(row) : null
}

export async function listResponseActionsForComplaint(
  db: D1Database,
  complaintId: string,
): Promise<ResponseAction[]> {
  const rows = await db
    .prepare('SELECT * FROM response_actions WHERE complaint_id=? ORDER BY created_at')
    .bind(complaintId)
    .all<Record<string, unknown>>()
  return rows.results.map(toAction)
}

/**
 * Edit a DRAFT's body. Only DRAFT actions can be edited — once submitted
 * for approval, the text is frozen to preserve the review chain.
 */
export async function updateResponseDraft(
  db: D1Database,
  input: { id: string; body: string },
): Promise<ResponseAction> {
  const now = new Date().toISOString()
  const result = await db
    .prepare(
      `UPDATE response_actions SET content_ref=?,updated_at=? WHERE id=? AND policy_state='DRAFT'`,
    )
    .bind(input.body, now, input.id)
    .run()
  if ((result.meta?.changes ?? 0) === 0) {
    const row = await db
      .prepare('SELECT policy_state FROM response_actions WHERE id=?')
      .bind(input.id)
      .first<{ policy_state: string }>()
    if (!row) throw new ResponseWorkflowError('ACTION_NOT_FOUND', `response action ${input.id} not found`)
    throw new ResponseWorkflowError(
      'INVALID_TRANSITION',
      `cannot edit response action ${input.id} in state ${row.policy_state} (only DRAFT can be edited)`,
    )
  }
  const updated = await db
    .prepare('SELECT * FROM response_actions WHERE id=?')
    .bind(input.id)
    .first<Record<string, unknown>>()
  return toAction(updated!)
}
