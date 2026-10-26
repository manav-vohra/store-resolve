// Phase C: ack/response workflow — templates, approval-gated lifecycle,
// sending always blocked. No external sends anywhere in these tests.
import { describe, expect, it } from 'vitest'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import {
  ResponseWorkflowError,
  approveResponseAction,
  createResponseDraft,
  createResponseTemplate,
  getResponseTemplate,
  listResponseActionsForComplaint,
  listResponseTemplates,
  rejectResponseAction,
  renderTemplate,
  sendResponseAction,
  submitResponseForApproval,
  updateResponseDraft,
} from '../worker/response-workflow'

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
    return { results: this.db.all(this.sql, this.values) as T[], success: true }
  }
  async first<T>(): Promise<T | null> {
    return this.db.first(this.sql, this.values) as T | null
  }
  async run(): Promise<D1Result> {
    this.db.run(this.sql, this.values)
    const changes = (this.db as unknown as { lastChanges?: number }).lastChanges
    // Reset the marker so unrelated runs default to 1.
    if (changes !== undefined) (this.db as unknown as { lastChanges?: number }).lastChanges = undefined
    return { results: [], success: true, meta: { changes: changes ?? 1 } }
  }
}

class FakeDb implements D1Database {
  templates = new Map<string, Record<string, unknown>>()
  actions = new Map<string, Record<string, unknown>>()
  complaints = new Set<string>(['c1'])
  prepare(sql: string): D1PreparedStatement {
    return new FakeStatement(sql, this)
  }
  async batch<T>(): Promise<D1Result<T>[]> {
    return []
  }
  async exec(): Promise<unknown> {
    return null
  }
  first(sql: string, values: unknown[]): unknown {
    if (sql.includes('FROM complaints WHERE id=?'))
      return this.complaints.has(String(values[0])) ? { id: values[0] } : null
    if (sql.includes('FROM response_templates WHERE id=?'))
      return this.templates.get(String(values[0])) ?? null
    if (sql.includes('FROM response_actions WHERE id=?'))
      return this.actions.get(String(values[0])) ?? null
    if (sql.includes('FROM gmail_messages')) {
      // Fake Dunkin corporate source for complaint c1.
      if (String(values[0]) === 'c1')
        return {
          gmail_message_id: 'msg-1',
          gmail_thread_id: 'thread-1',
          internal_date: '2026-01-01T00:00:00Z',
          sender: 'Dunkin <cases@dunkinbrands.com>',
          recipients: 'bvohra@msn.com',
          subject: 'Complaint CCC123',
        }
      return null
    }
    if (sql.includes('SELECT policy_state FROM response_actions WHERE id=?')) {
      const action = this.actions.get(String(values[0]))
      return action ? { policy_state: action.policy_state } : null
    }
    return null
  }
  all(sql: string, values: unknown[]): unknown[] {
    if (sql.includes('FROM response_templates WHERE active=1'))
      return [...this.templates.values()].filter((t) => Number(t.active) === 1)
    if (sql.includes('FROM response_actions WHERE complaint_id=?'))
      return [...this.actions.values()].filter((a) => a.complaint_id === values[0])
    return []
  }
  run(sql: string, values: unknown[]): void {
    if (sql.includes('INSERT INTO response_templates')) {
      const [id, name, response_kind, subject_template, body_template, created_at, updated_at] = values
      this.templates.set(String(id), {
        id, name, response_kind, subject_template, body_template,
        version: 1, active: 1, created_at, updated_at,
      })
    }
    if (sql.includes('INSERT INTO response_actions')) {
      // (id,complaint_id,response_kind,policy_version,policy_state,content_kind,content_ref,template_version,created_at,updated_at)
      const [id, complaint_id, response_kind, policy_version, policy_state] = values
      this.actions.set(String(id), {
        id, complaint_id, response_kind, policy_version, policy_state,
        content_kind: values[5], content_ref: values[6], template_version: values[7],
        approved_by: null, approved_at: null, created_at: values[8], updated_at: values[9],
      })
    }
    if (sql.includes('UPDATE response_actions SET policy_state=?')) {
      // (to, approvedBy, approvedBy, now, now, id, current)
      const action = this.actions.get(String(values[5]))
      if (action && action.policy_state === values[6]) {
        action.policy_state = values[0]
        if (values[1]) {
          action.approved_by = values[1]
          action.approved_at = values[3]
        }
        action.updated_at = values[4]
      }
    }
    if (sql.includes(`UPDATE response_actions SET policy_state='SENDING'`)) {
      // Atomic claim: (now, id) — only if currently APPROVED.
      const action = this.actions.get(String(values[1]))
      if (action && action.policy_state === 'APPROVED') {
        action.policy_state = 'SENDING'
        action.updated_at = values[0]
        // Signal 1 change via a marker the FakeStatement can read.
        ;(this as unknown as { lastChanges: number }).lastChanges = 1
      } else {
        ;(this as unknown as { lastChanges: number }).lastChanges = 0
      }
    }
    if (sql.includes(`UPDATE response_actions SET policy_state='SENT'`)) {
      const action = this.actions.get(String(values[1]))
      if (action) {
        action.policy_state = 'SENT'
        action.updated_at = values[0]
      }
    }
    if (sql.includes(`UPDATE response_actions SET policy_state='NEEDS_RECONCILIATION'`)) {
      const action = this.actions.get(String(values[1]))
      if (action) {
        action.policy_state = 'NEEDS_RECONCILIATION'
        action.updated_at = values[0]
      }
    }
    if (sql.includes('UPDATE response_actions SET content_ref=?')) {
      // (body, now, id) — only if DRAFT.
      const action = this.actions.get(String(values[2]))
      if (action && action.policy_state === 'DRAFT') {
        action.content_ref = values[0]
        action.updated_at = values[1]
        ;(this as unknown as { lastChanges: number }).lastChanges = 1
      } else {
        ;(this as unknown as { lastChanges: number }).lastChanges = 0
      }
    }
  }
}

const db = () => new FakeDb() as unknown as D1Database

describe('renderTemplate', () => {
  it('substitutes variables and leaves unknown placeholders visible', () => {
    expect(renderTemplate('Hi {{name}}, case {{caseId}}.', { name: 'Jo' })).toBe(
      'Hi Jo, case {{caseId}}.',
    )
  })
})

describe('response templates', () => {
  it('creates and lists active templates', async () => {
    const database = db()
    const created = await createResponseTemplate(database, {
      name: 'Standard ack',
      responseKind: 'ACKNOWLEDGMENT',
      bodyTemplate: 'Thanks {{name}}, we received your note.',
    })
    expect(created.id).toBeTruthy()
    expect(created.version).toBe(1)
    const listed = await listResponseTemplates(database)
    expect(listed).toHaveLength(1)
    expect(listed[0]?.name).toBe('Standard ack')
  })

  it('returns null for an unknown template', async () => {
    expect(await getResponseTemplate(db(), 'nope')).toBeNull()
  })
})

describe('response action lifecycle', () => {
  const setup = async () => {
    const database = db()
    const template = await createResponseTemplate(database, {
      name: 'Ack',
      responseKind: 'ACKNOWLEDGMENT',
      bodyTemplate: 'Hello {{name}}, we got your message about {{topic}}.',
    })
    return { database, template }
  }

  it('creates a draft rendered from a template', async () => {
    const { database, template } = await setup()
    const draft = await createResponseDraft(database, {
      complaintId: 'c1',
      templateId: template.id,
      templateVars: { name: 'Jo' },
    })
    expect(draft.policyState).toBe('DRAFT')
    expect(draft.policyVersion).toBe('response-policy.v1')
    expect(draft.contentRef).toBe('Hello Jo, we got your message about {{topic}}.')
    expect(draft.templateVersion).toBe('response-template.v1')
  })

  it('rejects drafts for unknown complaints', async () => {
    const { database } = await setup()
    await expect(createResponseDraft(database, { complaintId: 'nope', body: 'hi' })).rejects.toMatchObject({
      code: 'COMPLAINT_NOT_FOUND',
    })
  })

  it('rejects drafts with unknown templates', async () => {
    const { database } = await setup()
    await expect(
      createResponseDraft(database, { complaintId: 'c1', templateId: 'nope' }),
    ).rejects.toMatchObject({ code: 'TEMPLATE_NOT_FOUND' })
  })

  it('walks DRAFT -> PENDING_APPROVAL -> APPROVED', async () => {
    const { database } = await setup()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'custom body' })
    const pending = await submitResponseForApproval(database, draft.id)
    expect(pending.policyState).toBe('PENDING_APPROVAL')
    const approved = await approveResponseAction(database, { id: draft.id, approvedBy: 'owner-1' })
    expect(approved.policyState).toBe('APPROVED')
    expect(approved.approvedBy).toBe('owner-1')
    expect(approved.approvedAt).toBeTruthy()
  })

  it('supports rejection from pending approval', async () => {
    const { database } = await setup()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'x' })
    await submitResponseForApproval(database, draft.id)
    const rejected = await rejectResponseAction(database, { id: draft.id, rejectedBy: 'owner-1' })
    expect(rejected.policyState).toBe('REJECTED')
  })

  it('rejects invalid transitions (approve straight from DRAFT)', async () => {
    const { database } = await setup()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'x' })
    await expect(
      approveResponseAction(database, { id: draft.id, approvedBy: 'owner-1' }),
    ).rejects.toMatchObject({ code: 'INVALID_TRANSITION' })
  })

  it('lists actions per complaint', async () => {
    const { database } = await setup()
    await createResponseDraft(database, { complaintId: 'c1', body: 'a' })
    await createResponseDraft(database, { complaintId: 'c1', body: 'b' })
    expect(await listResponseActionsForComplaint(database, 'c1')).toHaveLength(2)
  })
})

describe('sendResponseAction', () => {
  const fakeProvider = {
    sendAcknowledgment: async () => undefined,
  }

  it('throws NOT_APPROVED for drafts (approval gate, not the global switch)', async () => {
    const database = db()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'x' })
    await expect(
      sendResponseAction(database, {
        id: draft.id,
        emailProvider: fakeProvider,
        mailboxIdentityVerified: true,
      }),
    ).rejects.toMatchObject({ code: 'NOT_APPROVED' })
  })

  it('throws MAILBOX_NOT_VERIFIED when identity is not verified', async () => {
    const database = db()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'x' })
    await submitResponseForApproval(database, draft.id)
    await approveResponseAction(database, { id: draft.id, approvedBy: 'dad' })
    await expect(
      sendResponseAction(database, {
        id: draft.id,
        emailProvider: fakeProvider,
        mailboxIdentityVerified: false,
      }),
    ).rejects.toMatchObject({ code: 'MAILBOX_NOT_VERIFIED' })
  })

  it('throws ACTION_NOT_FOUND for unknown actions', async () => {
    await expect(
      sendResponseAction(db(), {
        id: 'nope',
        emailProvider: fakeProvider,
        mailboxIdentityVerified: true,
      }),
    ).rejects.toMatchObject({ code: 'ACTION_NOT_FOUND' })
  })

  it('is a ResponseWorkflowError (never a raw send attempt)', async () => {
    const database = db()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'x' })
    const error = await sendResponseAction(database, {
      id: draft.id,
      emailProvider: fakeProvider,
      mailboxIdentityVerified: true,
    }).catch((e) => e)
    expect(error).toBeInstanceOf(ResponseWorkflowError)
  })
})

describe('updateResponseDraft', () => {
  it('edits a DRAFT body', async () => {
    const database = db()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'original' })
    const updated = await updateResponseDraft(database, { id: draft.id, body: 'edited' })
    expect(updated.contentRef).toBe('edited')
    expect(updated.policyState).toBe('DRAFT')
  })

  it('refuses to edit a non-DRAFT action', async () => {
    const database = db()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'x' })
    await submitResponseForApproval(database, draft.id)
    await expect(
      updateResponseDraft(database, { id: draft.id, body: 'y' }),
    ).rejects.toMatchObject({ code: 'INVALID_TRANSITION' })
  })
})

describe('sendResponseAction atomic claim', () => {
  it('second concurrent send loses the race (ALREADY_SENT)', async () => {
    const database = db()
    const draft = await createResponseDraft(database, { complaintId: 'c1', body: 'resolution text' })
    await submitResponseForApproval(database, draft.id)
    await approveResponseAction(database, { id: draft.id, approvedBy: 'dad' })
    const provider = { sendAcknowledgment: async () => undefined }
    // First send wins.
    const sent = await sendResponseAction(database, {
      id: draft.id,
      emailProvider: provider,
      mailboxIdentityVerified: true,
    })
    expect(sent.policyState).toBe('SENT')
    // Second send is refused — never a duplicate.
    await expect(
      sendResponseAction(database, {
        id: draft.id,
        emailProvider: provider,
        mailboxIdentityVerified: true,
      }),
    ).rejects.toMatchObject({ code: 'ALREADY_SENT' })
  })
})
