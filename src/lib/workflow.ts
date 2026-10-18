import type {
  AppConfig,
  AppState,
  Complaint,
  EventType,
  NewComplaint,
  Notification,
  Severity,
  User,
} from './types'

const ackMinutes: Record<Severity, number> = { LOW: 60, MEDIUM: 30, HIGH: 15, CRITICAL: 5 }
const resolutionHours: Partial<Record<Severity, number>> = { LOW: 48, MEDIUM: 24, HIGH: 12 }
const id = (prefix: string) => `${prefix}-${crypto.randomUUID()}`
const plus = (iso: string, amount: number, unit: 'minute' | 'hour') =>
  new Date(
    new Date(iso).getTime() + amount * (unit === 'minute' ? 60_000 : 3_600_000),
  ).toISOString()

export const acknowledgementTemplate = (caseId: string) =>
  `Thank you for bringing this matter to our attention.\n\nWe have received the complaint and are reviewing the matter with the appropriate store management team.\n\nWe will provide an update once the matter has been addressed.\n\nInternal Case ID: ${caseId}`

function event(
  complaintId: string,
  type: EventType,
  actor: string,
  timestamp: string,
  metadata?: Record<string, unknown>,
) {
  return { id: id('evt'), complaintId, type, actor, timestamp, metadata }
}
export function eligibility(
  recipient: User,
  config: AppConfig,
  storeId?: string,
): { status: Notification['status']; reason?: string; provider: Notification['provider'] } {
  if (!config.externalNotificationsEnabled)
    return {
      status: 'SUPPRESSED',
      reason: 'EXTERNAL_NOTIFICATIONS_DISABLED',
      provider: 'MOCK',
    }
  if (config.mode === 'MOCK')
    return { status: 'SUPPRESSED', reason: 'MOCK mode never sends externally', provider: 'MOCK' }
  if (!recipient.smsEnabled)
    return { status: 'SUPPRESSED', reason: 'SMS is disabled for this recipient', provider: 'MOCK' }
  if (recipient.recipientKind === 'PILOT_ADMIN')
    return {
      status: 'SUPPRESSED',
      reason: 'TEST_RECIPIENT_ONLY',
      provider: 'MOCK',
    }
  const owner = Boolean(recipient.complaintNotificationsEnabled)
  if (config.mode === 'FAMILY_PILOT' && !owner)
    return {
      status: 'SUPPRESSED',
      reason: 'FAMILY_PILOT',
      provider: 'MOCK',
    }
  if (config.mode === 'SINGLE_STORE_PILOT' && !owner && storeId !== config.pilotStoreId)
    return {
      status: 'SUPPRESSED',
      reason: 'Store is outside the single-store pilot',
      provider: 'MOCK',
    }
  return { status: 'PENDING', provider: 'SIGNALWIRE' }
}
function notification(
  state: AppState,
  complaint: Complaint,
  recipientId: string,
  type: EventType,
  message: string,
  now: string,
): Notification {
  const recipient = state.users.find((u) => u.id === recipientId)
  if (!recipient)
    throw new Error(`Notification recipient not found: ${String(recipientId)}`)
  const eligible = eligibility(recipient, state.config, complaint.storeId)
  return {
    id: id('ntf'),
    complaintId: complaint.id,
    eventType: type,
    recipientUserId: recipientId,
    channel: 'SMS',
    message,
    status: eligible.status,
    provider: eligible.provider,
    createdAt: now,
    sentAt: undefined,
    failureReason: eligible.reason,
  }
}
function notifyOwners(
  state: AppState,
  complaint: Complaint,
  type: EventType,
  message: string,
  now: string,
) {
  const notifications = state.users
    .filter((user) => user.complaintNotificationsEnabled && user.recipientKind !== 'PILOT_ADMIN')
    .map((owner) => notification(state, complaint, owner.id, type, message, now))
  complaint.notifications.push(...notifications)
  const planned = notifications.some((item) => item.status === 'PENDING')
  complaint.events.push(
    event(
      complaint.id,
      planned ? 'OWNER_NOTIFICATION_PLANNED' : 'OWNER_NOTIFICATION_SUPPRESSED',
      'system',
      now,
      { reason: type, notificationCount: notifications.length },
    ),
  )
}

export function createComplaint(
  state: AppState,
  input: NewComplaint,
  now = new Date().toISOString(),
  options: {
    source?: 'MANUAL' | 'GMAIL' | 'MICROSOFT_GRAPH'
    actor?: string
    acknowledged?: boolean
    ingestionMode?: 'LIVE' | 'BACKFILL' | 'TEST'
  } = {},
): { state: AppState; complaint: Complaint; duplicate: boolean } {
  const next = structuredClone(state)
  const existing = next.complaints.find(
    (c) => c.externalCaseId.toLowerCase() === input.externalCaseId.trim().toLowerCase(),
  )
  if (existing) {
    existing.followUps.push({ receivedAt: now, text: input.complaintText })
    existing.events.push(event(existing.id, 'FOLLOW_UP_RECEIVED', 'simulator', now))
    return { state: next, complaint: existing, duplicate: true }
  }
  const store = next.stores.find((s) => s.number === input.storeNumber.trim())
  const complaintId = id('cmp')
  const internalCase = `SR-${new Date(now).getUTCFullYear()}-${String(next.complaints.length + 1).padStart(4, '0')}`
  const acknowledged = options.acknowledged ?? true
  const actor = options.actor ?? 'simulator'
  const ingestionMode = options.ingestionMode ?? 'LIVE'
  const operationalStartedAt = ingestionMode === 'LIVE' ? now : undefined
  const complaint: Complaint = {
    id: complaintId,
    externalCaseId: input.externalCaseId.trim(),
    storeId: store?.id,
    assignedManagerId: store?.managerId,
    subject: input.subject,
    complaintText: input.complaintText,
    source: options.source ?? 'MANUAL',
    category: input.category,
    severity: input.severity,
    status: store ? 'MANAGER_NOTIFIED' : 'ROUTING_REVIEW',
    isAckOverdue: false,
    isResolutionOverdue: false,
    routingReason: store
      ? `Exact store number match: #${store.number}`
      : `No match for store number: ${input.storeNumber}`,
    routingConfidence: store ? 'HIGH' : 'REVIEW',
    receivedAt: now,
    ingestionMode,
    operationalStartedAt,
    dunkinAcknowledgedAt: acknowledged ? now : undefined,
    acknowledgementStatus: acknowledged ? 'SENT' : 'DISABLED',
    acknowledgementBody: acknowledgementTemplate(internalCase),
    ackDeadline: plus(
      operationalStartedAt ?? now,
      state.config.managerAckDeadlineMinutes ?? ackMinutes[input.severity],
      'minute',
    ),
    resolutionDeadline:
      (state.config.managerResolutionTargetHours ?? resolutionHours[input.severity])
        ? plus(
            operationalStartedAt ?? now,
            state.config.managerResolutionTargetHours ?? resolutionHours[input.severity]!,
            'hour',
          )
        : undefined,
    events: [],
    notifications: [],
    followUps: [],
  }
  complaint.id = internalCase
  complaint.events.push(
    event(complaint.id, 'COMPLAINT_RECEIVED', actor, now),
    event(complaint.id, store ? 'STORE_ASSIGNED' : 'ROUTING_REVIEW_REQUIRED', 'system', now, {
      reason: complaint.routingReason,
    }),
  )
  if (acknowledged)
    complaint.events.push(
      event(complaint.id, 'DUNKIN_ACKNOWLEDGED', 'system', now, {
        simulated: options.source === undefined || options.source === 'MANUAL',
        templateVersion: 'v1',
      }),
    )
  const storeLabel = store ? `#${store.number}` : 'Routing review required'
  notifyOwners(
    next,
    complaint,
    'COMPLAINT_RECEIVED',
    `NEW DUNKIN COMPLAINT\nStore: ${storeLabel}\nCategory: ${input.category}\nSeverity: ${input.severity}\nCase: ${complaint.id}`,
    now,
  )
  if (store) {
    // Stores may have NO assigned manager (prod: all 6 active stores have
    // manager_id NULL — Dad is OWNER, not a store manager). Previously this
    // unconditionally called notification() with a null managerId, which
    // threw (recipient not found) and failed the entire import with
    // FAILED_PERSISTENCE for every message. Only notify a manager when one
    // is actually assigned AND exists in state.users; otherwise record that
    // no manager was assigned and continue — the complaint is still created.
    const manager = store.managerId
      ? next.users.find((u) => u.id === store.managerId)
      : undefined
    if (manager) {
      const managerNotification = notification(
        next,
        complaint,
        manager.id,
        'MANAGER_NOTIFIED',
        `New complaint ${complaint.id} requires your acknowledgment.`,
        now,
      )
      complaint.notifications.push(managerNotification)
      complaint.events.push(
        event(
          complaint.id,
          managerNotification.status === 'PENDING'
            ? 'MANAGER_NOTIFICATION_PLANNED'
            : 'MANAGER_NOTIFICATION_SUPPRESSED',
          'system',
          now,
          { notificationId: managerNotification.id },
        ),
      )
    } else {
      complaint.events.push(
        event(complaint.id, 'MANAGER_NOT_ASSIGNED', 'system', now, {
          reason: 'STORE_HAS_NO_ASSIGNED_MANAGER',
          storeId: store.id,
        }),
      )
    }
  }
  next.complaints.unshift(complaint)
  return { state: next, complaint, duplicate: false }
}

export function updateComplaint(
  state: AppState,
  complaintId: string,
  action:
    | 'ACKNOWLEDGE'
    | 'START_INVESTIGATION'
    | 'CONTACT_CUSTOMER'
    | 'SUBMIT_RESOLUTION'
    | 'CLOSE'
    | 'REOPEN',
  actorId: string,
  data: Record<string, unknown> = {},
  now = new Date().toISOString(),
): AppState {
  const next = structuredClone(state)
  const c = next.complaints.find((item) => item.id === complaintId)
  if (!c) throw new Error('Complaint not found')
  const actor = next.users.find((u) => u.id === actorId)
  if (!actor) throw new Error('User not found')
  const isOwner = actor.role === 'OWNER' || actor.role === 'ADMIN'
  const isManager = actor.id === c.assignedManagerId
  if (
    ['ACKNOWLEDGE', 'START_INVESTIGATION', 'CONTACT_CUSTOMER', 'SUBMIT_RESOLUTION'].includes(
      action,
    ) &&
    !isManager
  )
    throw new Error('Manager access required')
  if (['CLOSE', 'REOPEN'].includes(action) && !isOwner) throw new Error('Owner access required')
  if (action === 'ACKNOWLEDGE') {
    if (c.status !== 'MANAGER_NOTIFIED') throw new Error('Complaint is not awaiting acknowledgment')
    c.status = 'ACKNOWLEDGED'
    c.managerAcknowledgedAt = now
    c.events.push(event(c.id, 'MANAGER_ACKNOWLEDGED', actor.name, now))
  }
  if (action === 'START_INVESTIGATION') {
    if (c.status !== 'ACKNOWLEDGED') throw new Error('Complaint must be acknowledged first')
    c.status = 'INVESTIGATING'
    c.investigationStartedAt = now
    c.events.push(event(c.id, 'INVESTIGATION_STARTED', actor.name, now))
  }
  if (action === 'CONTACT_CUSTOMER') {
    if (!['ACKNOWLEDGED', 'INVESTIGATING'].includes(c.status))
      throw new Error('Investigation must be active before customer contact')
    c.customerContacted = true
    c.customerContactedAt = now
    c.customerContactOutcome = String(data.outcome ?? 'Contact recorded')
    c.events.push(
      event(c.id, 'CUSTOMER_CONTACTED', actor.name, now, { outcome: c.customerContactOutcome }),
    )
  }
  if (action === 'SUBMIT_RESOLUTION') {
    if (c.status !== 'INVESTIGATING') throw new Error('Investigation must be started first')
    c.status = 'RESOLUTION_SUBMITTED'
    c.managerFindings = String(data.findings ?? '')
    c.correctiveAction = String(data.correctiveAction ?? '')
    c.resolutionNotes = String(data.resolutionNotes ?? '')
    c.resolutionSubmittedAt = now
    c.events.push(event(c.id, 'RESOLUTION_SUBMITTED', actor.name, now))
    notifyOwners(
      next,
      c,
      'RESOLUTION_SUBMITTED',
      `RESOLUTION SUBMITTED\nCase: ${c.id}\nThe store manager has submitted a resolution.`,
      now,
    )
  }
  if (action === 'CLOSE') {
    if (c.status !== 'RESOLUTION_SUBMITTED')
      throw new Error('Resolution must be submitted before closure')
    c.status = 'CLOSED'
    c.closedAt = now
    c.closedBy = actor.name
    c.ownerReviewedAt = now
    c.ownerNotes = String(data.ownerNotes ?? '')
    c.events.push(event(c.id, 'COMPLAINT_CLOSED', actor.name, now))
    notifyOwners(next, c, 'COMPLAINT_CLOSED', `COMPLAINT CLOSED\nCase: ${c.id}`, now)
  }
  if (action === 'REOPEN') {
    if (c.status !== 'CLOSED') throw new Error('Only a closed complaint can be reopened')
    const reason = String(data.reason ?? '').trim()
    if (!reason) throw new Error('A reopen reason is required')
    c.status = 'INVESTIGATING'
    c.closedAt = undefined
    c.closedBy = undefined
    c.reopenedAt = now
    c.reopenReason = reason
    c.events.push(event(c.id, 'COMPLAINT_REOPENED', actor.name, now, { reason }))
    notifyOwners(next, c, 'COMPLAINT_REOPENED', `COMPLAINT REOPENED\nCase: ${c.id}`, now)
  }
  return next
}

export function processDeadlines(state: AppState, now = new Date().toISOString()): AppState {
  const next = structuredClone(state)
  for (const c of next.complaints) {
    if (c.ingestionMode !== 'LIVE' || !c.operationalStartedAt) continue
    if (!c.managerAcknowledgedAt && new Date(now) > new Date(c.ackDeadline) && !c.isAckOverdue) {
      c.isAckOverdue = true
      c.events.push(event(c.id, 'MANAGER_ACK_OVERDUE', 'system', now))
      notifyOwners(
        next,
        c,
        'MANAGER_ACK_OVERDUE',
        `MANAGER ACKNOWLEDGMENT OVERDUE\nCase: ${c.id}`,
        now,
      )
    }
    if (
      c.resolutionDeadline &&
      !c.resolutionSubmittedAt &&
      c.status !== 'CLOSED' &&
      new Date(now) > new Date(c.resolutionDeadline) &&
      !c.isResolutionOverdue
    ) {
      c.isResolutionOverdue = true
      c.events.push(event(c.id, 'RESOLUTION_OVERDUE', 'system', now))
      notifyOwners(next, c, 'RESOLUTION_OVERDUE', `RESOLUTION OVERDUE\nCase: ${c.id}`, now)
    }
  }
  return next
}

export function metrics(state: AppState) {
  const all = state.complaints
  const closed = all.filter((c) => c.status === 'CLOSED')
  const avg = (xs: number[]) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0)
  return {
    total: all.length,
    open: all.length - closed.length,
    closed: closed.length,
    overdue: all.filter((c) => c.isAckOverdue || c.isResolutionOverdue).length,
    awaitingAck: all.filter((c) => !c.managerAcknowledgedAt && c.status !== 'CLOSED').length,
    awaitingResolution: all.filter((c) => c.managerAcknowledgedAt && !c.resolutionSubmittedAt)
      .length,
    avgAckMinutes: Math.round(
      avg(
        all
          .filter((c) => c.managerAcknowledgedAt)
          .map(
            (c) =>
              (new Date(c.managerAcknowledgedAt!).getTime() - new Date(c.receivedAt).getTime()) /
              60000,
          ),
      ),
    ),
    avgResolutionHours:
      Math.round(
        avg(
          all
            .filter((c) => c.resolutionSubmittedAt)
            .map(
              (c) =>
                (new Date(c.resolutionSubmittedAt!).getTime() - new Date(c.receivedAt).getTime()) /
                3600000,
            ),
        ) * 10,
      ) / 10,
  }
}
