import { beforeEach, describe, expect, it } from 'vitest'
import {
  createComplaint,
  eligibility,
  metrics,
  processDeadlines,
  updateComplaint,
} from '../src/lib/workflow'
import { initialState } from '../src/lib/seed'
import type { AppState, NewComplaint } from '../src/lib/types'
const input: NewComplaint = {
  externalCaseId: 'DKN-TEST-1',
  storeNumber: '41001',
  subject: 'Test complaint',
  complaintText: 'A fictional customer concern for automated testing.',
  category: 'Customer Service',
  severity: 'MEDIUM',
}
const received = '2026-08-11T12:00:00.000Z'
let state: AppState
beforeEach(() => {
  state = structuredClone(initialState)
})
const create = () => createComplaint(state, input, received)

describe('complaint creation and routing', () => {
  it('creates a complaint, event history, simulated acknowledgment and exact store assignment', () => {
    const r = create()
    expect(r.complaint.id).toBe('SR-2026-0001')
    expect(r.complaint.storeId).toBe('store-1')
    expect(r.complaint.routingConfidence).toBe('HIGH')
    expect(r.complaint.dunkinAcknowledgedAt).toBe(received)
    expect(r.complaint.events.map((e) => e.type)).toContain('COMPLAINT_RECEIVED')
  })
  it('leaves an unknown store reviewable and alerts ownership', () => {
    const r = createComplaint(
      state,
      { ...input, externalCaseId: 'X', storeNumber: '99999' },
      received,
    )
    expect(r.complaint.storeId).toBeUndefined()
    expect(r.complaint.routingConfidence).toBe('REVIEW')
    expect(r.complaint.notifications).toHaveLength(3)
  })
  it('creates complaint when store has NO assigned manager (prod: all stores manager_id NULL)', () => {
    // Oct 17: all 6 BACKFILL imports failed with FAILED_PERSISTENCE because
    // createComplaint unconditionally notified store.managerId, which was
    // null for every prod store. Must create the complaint + MANAGER_NOT_ASSIGNED
    // event instead of throwing.
    const noManagerState = structuredClone(state)
    for (const s of noManagerState.stores) s.managerId = undefined as unknown as string
    const r = createComplaint(noManagerState, input, received)
    expect(r.complaint.storeId).toBe('store-1')
    expect(r.complaint.status).toBe('MANAGER_NOTIFIED')
    expect(r.complaint.events.map((e) => e.type)).toContain('MANAGER_NOT_ASSIGNED')
    // Owner notifications still created; no manager notification.
    expect(r.complaint.notifications.length).toBeGreaterThan(0)
    expect(r.complaint.notifications.every((n) => n.eventType !== 'MANAGER_NOTIFIED')).toBe(true)
  })
  it('treats duplicate external case as follow-up without creating a second complaint', () => {
    const first = create()
    const second = createComplaint(
      first.state,
      { ...input, complaintText: 'Additional follow-up' },
      received,
    )
    expect(second.duplicate).toBe(true)
    expect(second.state.complaints).toHaveLength(1)
    expect(second.complaint.followUps).toHaveLength(1)
    expect(second.complaint.events.at(-1)?.type).toBe('FOLLOW_UP_RECEIVED')
  })
})
describe('independent notification safety', () => {
  it('creates separate ownership and manager records; delivery never acknowledges', () => {
    const c = create().complaint
    expect(
      c.notifications.filter((n) => ['father', 'uncle', 'grandfather'].includes(n.recipientUserId)),
    ).toHaveLength(3)
    expect(new Set(c.notifications.map((n) => n.id)).size).toBe(4)
    expect(c.managerAcknowledgedAt).toBeUndefined()
    expect(c.notifications.every((n) => n.status === 'SUPPRESSED')).toBe(true)
  })
  it('FAMILY_PILOT permits enabled owners but suppresses manager', () => {
    state.config = { mode: 'FAMILY_PILOT', externalNotificationsEnabled: true }
    state.users = state.users.map((u) => ({ ...u, smsEnabled: true }))
    const c = create().complaint
    expect(c.notifications.find((n) => n.recipientUserId === 'father')?.status).toBe('PENDING')
    expect(c.notifications.find((n) => n.recipientUserId === 'father')?.provider).toBe('SIGNALWIRE')
    expect(c.notifications.find((n) => n.recipientUserId === 'manager-1')?.status).toBe(
      'SUPPRESSED',
    )
  })
  it('global switch suppresses every external call regardless of mode', () => {
    state.config = { mode: 'FULL', externalNotificationsEnabled: false }
    state.users = state.users.map((u) => ({ ...u, smsEnabled: true }))
    const notifications = create().complaint.notifications
    expect(notifications.every((n) => n.status === 'SUPPRESSED')).toBe(true)
    expect(notifications.every((n) => n.failureReason === 'EXTERNAL_NOTIFICATIONS_DISABLED')).toBe(
      true,
    )
    expect(notifications.some((n) => n.status === 'PENDING' || n.status === 'SENT')).toBe(false)
    const events = create().complaint.events.map((event) => event.type)
    expect(events).toContain('OWNER_NOTIFICATION_SUPPRESSED')
    expect(events).not.toContain('OWNER_NOTIFIED')
    expect(events).toContain('MANAGER_NOTIFICATION_SUPPRESSED')
    expect(events).not.toContain('MANAGER_NOTIFIED')
  })
  it('never permits PILOT_ADMIN in complaint delivery, even in FULL mode', () => {
    const pilotAdmin = state.users.find((u) => u.recipientKind === 'PILOT_ADMIN')!
    expect(
      eligibility(
        { ...pilotAdmin, active: true, smsEnabled: true, phone: '+15551234567' },
        { mode: 'FULL', externalNotificationsEnabled: true },
      ),
    ).toMatchObject({ status: 'SUPPRESSED', reason: 'TEST_RECIPIENT_ONLY' })
    expect(create().complaint.notifications.some((n) => n.recipientUserId === pilotAdmin.id)).toBe(
      false,
    )
  })
})
describe('manager workflow and authorization', () => {
  it('records exact acknowledgment timestamp only on explicit action', () => {
    const r = create()
    const next = updateComplaint(
      r.state,
      r.complaint.id,
      'ACKNOWLEDGE',
      'manager-1',
      {},
      '2026-08-11T12:05:00.000Z',
    )
    expect(next.complaints[0].managerAcknowledgedAt).toBe('2026-08-11T12:05:00.000Z')
    expect(next.complaints[0].status).toBe('ACKNOWLEDGED')
  })
  it('supports investigation, contact and resolution without auto-closing', () => {
    const r = create()
    state = updateComplaint(r.state, r.complaint.id, 'ACKNOWLEDGE', 'manager-1', {}, received)
    state = updateComplaint(state, r.complaint.id, 'START_INVESTIGATION', 'manager-1', {}, received)
    state = updateComplaint(
      state,
      r.complaint.id,
      'CONTACT_CUSTOMER',
      'manager-1',
      { outcome: 'Reached customer' },
      received,
    )
    state = updateComplaint(
      state,
      r.complaint.id,
      'SUBMIT_RESOLUTION',
      'manager-1',
      { findings: 'Reviewed', correctiveAction: 'Coached', resolutionNotes: 'Complete' },
      received,
    )
    const c = state.complaints[0]
    expect(c.status).toBe('RESOLUTION_SUBMITTED')
    expect(c.customerContacted).toBe(true)
    expect(c.closedAt).toBeUndefined()
    expect(c.notifications.filter((n) => n.eventType === 'RESOLUTION_SUBMITTED')).toHaveLength(3)
  })
  it('prevents grandfather from owner mutation and other managers from actions', () => {
    const r = create()
    expect(() => updateComplaint(r.state, r.complaint.id, 'CLOSE', 'grandfather')).toThrow(
      'Owner access required',
    )
    expect(() => updateComplaint(r.state, r.complaint.id, 'ACKNOWLEDGE', 'manager-2')).toThrow(
      'Manager access required',
    )
  })
  it('allows an owner to close only a submitted resolution and records time', () => {
    const r = create()
    state = updateComplaint(r.state, r.complaint.id, 'ACKNOWLEDGE', 'manager-1', {}, received)
    state = updateComplaint(state, r.complaint.id, 'START_INVESTIGATION', 'manager-1', {}, received)
    state = updateComplaint(state, r.complaint.id, 'SUBMIT_RESOLUTION', 'manager-1', {}, received)
    state = updateComplaint(
      state,
      r.complaint.id,
      'CLOSE',
      'father',
      {},
      '2026-08-12T00:00:00.000Z',
    )
    expect(state.complaints[0].status).toBe('CLOSED')
    expect(state.complaints[0].closedAt).toBe('2026-08-12T00:00:00.000Z')
  })
})
describe('deadlines and reporting', () => {
  it('preserves historical timestamps without starting SLA work for backfill or test intake', () => {
    for (const ingestionMode of ['BACKFILL', 'TEST'] as const) {
      const result = createComplaint(state, input, received, { ingestionMode, acknowledged: false })
      const processed = processDeadlines(result.state, '2026-09-01T00:00:00.000Z')
      expect(processed.complaints[0].ingestionMode).toBe(ingestionMode)
      expect(processed.complaints[0].operationalStartedAt).toBeUndefined()
      expect(processed.complaints[0].isAckOverdue).toBe(false)
      expect(processed.complaints[0].isResolutionOverdue).toBe(false)
      expect(
        processed.complaints[0].notifications.every(
          (notification) => notification.status === 'SUPPRESSED',
        ),
      ).toBe(true)
    }
  })
  it('marks missed acknowledgment overdue, emits three alerts, and is idempotent', () => {
    const r = create()
    const once = processDeadlines(r.state, '2026-08-11T13:00:00.000Z')
    const twice = processDeadlines(once, '2026-08-11T14:00:00.000Z')
    const c = twice.complaints[0]
    expect(c.isAckOverdue).toBe(true)
    expect(c.notifications.filter((n) => n.eventType === 'MANAGER_ACK_OVERDUE')).toHaveLength(3)
    expect(c.events.filter((e) => e.type === 'MANAGER_ACK_OVERDUE')).toHaveLength(1)
  })
  it('marks missed resolution overdue and creates owner alerts', () => {
    const r = create()
    state = updateComplaint(r.state, r.complaint.id, 'ACKNOWLEDGE', 'manager-1', {}, received)
    state = processDeadlines(state, '2026-08-13T13:00:00.000Z')
    expect(state.complaints[0].isResolutionOverdue).toBe(true)
    expect(
      state.complaints[0].notifications.filter((n) => n.eventType === 'RESOLUTION_OVERDUE'),
    ).toHaveLength(3)
  })
  it('updates reporting metrics through close', () => {
    const r = create()
    expect(metrics(r.state).open).toBe(1)
    state = updateComplaint(
      r.state,
      r.complaint.id,
      'ACKNOWLEDGE',
      'manager-1',
      {},
      '2026-08-11T12:10:00Z',
    )
    state = updateComplaint(
      state,
      r.complaint.id,
      'START_INVESTIGATION',
      'manager-1',
      {},
      '2026-08-11T12:15:00Z',
    )
    state = updateComplaint(
      state,
      r.complaint.id,
      'SUBMIT_RESOLUTION',
      'manager-1',
      {},
      '2026-08-11T14:00:00Z',
    )
    state = updateComplaint(state, r.complaint.id, 'CLOSE', 'father', {}, '2026-08-11T15:00:00Z')
    expect(metrics(state)).toMatchObject({
      open: 0,
      closed: 1,
      avgAckMinutes: 10,
      avgResolutionHours: 2,
    })
  })
})
describe('fixture hygiene', () => {
  it('contains only reserved fictional contact domains and numbers', () => {
    expect(state.users.every((u) => u.email.endsWith('.invalid'))).toBe(true)
    expect(state.users.every((u) => !u.phone || u.phone.startsWith('+1555'))).toBe(true)
  })
})
