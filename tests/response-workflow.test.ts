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
    return { results: [], success: true, meta: { changes: 1 } }
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
