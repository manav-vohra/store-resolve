import type { AppState, Complaint, Notification, Store, User } from '../src/lib/types'

export interface D1Result<T = Record<string, unknown>> {
  results: T[]
  success: boolean
  meta?: Record<string, unknown>
}
export interface D1PreparedStatement {
  bind(...values: unknown[]): D1PreparedStatement
  all<T = Record<string, unknown>>(): Promise<D1Result<T>>
  first<T = Record<string, unknown>>(): Promise<T | null>
  run(): Promise<D1Result>
}
export interface D1Database {
  prepare(query: string): D1PreparedStatement
  batch<T = unknown>(statements: D1PreparedStatement[]): Promise<D1Result<T>[]>
  exec(query: string): Promise<unknown>
}

const bool = (value: unknown) => Boolean(value)
export async function loadState(db: D1Database): Promise<AppState> {
  const [userRows, storeRows, aliasRows, complaintRows, eventRows, notificationRows, settingRows] =
    await Promise.all([
      db.prepare('SELECT * FROM users ORDER BY id').all<Record<string, unknown>>(),
      db.prepare('SELECT * FROM stores ORDER BY id').all<Record<string, unknown>>(),
      db
        .prepare('SELECT store_id,alias_normalized FROM store_aliases ORDER BY alias_normalized')
        .all<{
          store_id: string
          alias_normalized: string
        }>(),
      db
        .prepare('SELECT * FROM complaints ORDER BY received_at DESC')
        .all<Record<string, unknown>>(),
      db
        .prepare('SELECT * FROM complaint_events ORDER BY timestamp')
        .all<Record<string, unknown>>(),
      db.prepare('SELECT * FROM notifications ORDER BY created_at').all<Record<string, unknown>>(),
      db.prepare('SELECT * FROM settings').all<{ key: string; value: string }>(),
    ])
  const users: User[] = userRows.results.map((r) => ({
    id: String(r.id),
    name: String(r.name),
    email: String(r.email ?? ''),
    phone: String(r.phone ?? ''),
    role: r.role as User['role'],
    recipientKind: (r.recipient_kind ?? 'STANDARD') as User['recipientKind'],
    active: bool(r.active),
    smsEnabled: bool(r.sms_enabled),
    complaintNotificationsEnabled: bool(r.complaint_notifications_enabled),
    timezone: String(r.timezone),
  }))
  const aliasesByStore = new Map<string, string[]>()
  for (const row of aliasRows.results) {
    const aliases = aliasesByStore.get(row.store_id) ?? []
    aliases.push(row.alias_normalized)
    aliasesByStore.set(row.store_id, aliases)
  }
  const stores: Store[] = storeRows.results.map((r) => ({
    id: String(r.id),
    number: String(r.dunkin_store_number),
    name: String(r.name),
    address: String(r.address),
    city: String(r.city),
    state: String(r.state),
    postalCode: String(r.postal_code),
    phone: String(r.phone),
    active: bool(r.active),
    managerId: r.manager_id ? String(r.manager_id) : undefined,
    aliases: aliasesByStore.get(String(r.id)) ?? [],
  }))
  const eventMap = new Map<string, Complaint['events']>()
  for (const r of eventRows.results) {
    const list = eventMap.get(String(r.complaint_id)) ?? []
    list.push({
      id: String(r.id),
      complaintId: String(r.complaint_id),
      type: r.event_type as Complaint['events'][number]['type'],
      actor: String(r.actor),
      timestamp: String(r.timestamp),
      metadata: r.metadata ? JSON.parse(String(r.metadata)) : undefined,
    })
    eventMap.set(String(r.complaint_id), list)
  }
  const notificationMap = new Map<string, Notification[]>()
  for (const r of notificationRows.results) {
    const n: Notification = {
      id: String(r.id),
      complaintId: String(r.complaint_id ?? ''),
      eventType: r.event_type as Notification['eventType'],
      recipientUserId: String(r.recipient_user_id),
      channel: r.channel as Notification['channel'],
      message: String(r.message),
      status: r.status as Notification['status'],
      provider: r.provider as Notification['provider'],
      providerMessageId: r.provider_message_id ? String(r.provider_message_id) : undefined,
      createdAt: String(r.created_at),
      sentAt: r.sent_at ? String(r.sent_at) : undefined,
      deliveredAt: r.delivered_at ? String(r.delivered_at) : undefined,
      failedAt: r.failed_at ? String(r.failed_at) : undefined,
      failureReason: r.failure_reason ? String(r.failure_reason) : undefined,
    }
    const list = notificationMap.get(n.complaintId) ?? []
    list.push(n)
    notificationMap.set(n.complaintId, list)
  }
  const complaints: Complaint[] = complaintRows.results.map((r) => ({
    id: String(r.id),
    externalCaseId: String(r.external_case_id),
    storeId: r.store_id ? String(r.store_id) : undefined,
    assignedManagerId: r.assigned_manager_id ? String(r.assigned_manager_id) : undefined,
    subject: String(r.subject),
    complaintText: String(r.complaint_text),
    source: (r.source ?? 'MANUAL') as Complaint['source'],
    gmailMessageId: r.gmail_message_id ? String(r.gmail_message_id) : undefined,
    gmailThreadId: r.gmail_thread_id ? String(r.gmail_thread_id) : undefined,
    sourceSender: r.source_sender ? String(r.source_sender) : undefined,
    customerName: r.customer_name ? String(r.customer_name) : undefined,
    customerEmail: r.customer_email ? String(r.customer_email) : undefined,
    customerPhone: r.customer_phone ? String(r.customer_phone) : undefined,
    occurrenceAt: r.occurrence_at ? String(r.occurrence_at) : undefined,
    category: String(r.category),
    severity: r.severity as Complaint['severity'],
    status: r.status as Complaint['status'],
    isAckOverdue: bool(r.is_ack_overdue),
    isResolutionOverdue: bool(r.is_resolution_overdue),
    routingReason: String(r.routing_reason),
    routingConfidence: r.routing_confidence as Complaint['routingConfidence'],
    receivedAt: String(r.received_at),
    ingestionMode: (r.ingestion_mode ?? 'LIVE') as Complaint['ingestionMode'],
    operationalStartedAt: r.operational_started_at ? String(r.operational_started_at) : undefined,
    dunkinAcknowledgedAt: r.dunkin_acknowledged_at ? String(r.dunkin_acknowledged_at) : undefined,
    acknowledgementStatus: (r.acknowledgement_status ??
      'DISABLED') as Complaint['acknowledgementStatus'],
    acknowledgementBody: r.acknowledgment_body ? String(r.acknowledgment_body) : '',
    managerNotifiedAt: r.manager_notified_at ? String(r.manager_notified_at) : undefined,
    managerAcknowledgedAt: r.manager_acknowledged_at
      ? String(r.manager_acknowledged_at)
      : undefined,
    investigationStartedAt: r.investigation_started_at
      ? String(r.investigation_started_at)
      : undefined,
    resolutionSubmittedAt: r.resolution_submitted_at
      ? String(r.resolution_submitted_at)
      : undefined,
    closedAt: r.closed_at ? String(r.closed_at) : undefined,
    closedBy: r.closed_by ? String(r.closed_by) : undefined,
    ownerReviewedAt: r.owner_reviewed_at ? String(r.owner_reviewed_at) : undefined,
    reopenedAt: r.reopened_at ? String(r.reopened_at) : undefined,
    reopenReason: r.reopen_reason ? String(r.reopen_reason) : undefined,
    ownerNotes: r.owner_notes ? String(r.owner_notes) : undefined,
    ackDeadline: String(r.ack_deadline),
    resolutionDeadline: r.resolution_deadline ? String(r.resolution_deadline) : undefined,
    managerFindings: r.manager_findings ? String(r.manager_findings) : undefined,
    customerContacted: r.customer_contacted == null ? undefined : bool(r.customer_contacted),
    customerContactedAt: r.customer_contacted_at ? String(r.customer_contacted_at) : undefined,
    customerContactOutcome: r.customer_contact_outcome
      ? String(r.customer_contact_outcome)
      : undefined,
    correctiveAction: r.corrective_action ? String(r.corrective_action) : undefined,
    resolutionNotes: r.resolution_notes ? String(r.resolution_notes) : undefined,
    events: eventMap.get(String(r.id)) ?? [],
    notifications: notificationMap.get(String(r.id)) ?? [],
    followUps: r.follow_ups ? JSON.parse(String(r.follow_ups)) : [],
  }))
  const settings = Object.fromEntries(settingRows.results.map((r) => [r.key, r.value]))
  return {
    users,
    stores,
    complaints,
    activeUserId: '',
    testNotifications: notificationMap.get('') ?? [],
    config: {
      mode: (settings.notification_mode ?? 'FAMILY_PILOT') as AppState['config']['mode'],
      externalNotificationsEnabled: settings.external_notifications_enabled === 'true',
      pilotStoreId: settings.pilot_store_id || undefined,
      emailIngestionEnabled: settings.email_ingestion_enabled === 'true',
      emailAckEnabled: settings.email_ack_enabled === 'true',
      managerAckDeadlineMinutes: Number(settings.manager_ack_deadline_minutes ?? 30),
      managerResolutionTargetHours: Number(settings.manager_resolution_target_hours ?? 24),
      escalationIntervalMinutes: Number(settings.escalation_interval_minutes ?? 60),
      signalWireReconcileAfterMinutes: Number(settings.signalwire_reconcile_after_minutes ?? 10),
      emailLookbackDays: Number(settings.email_lookback_days ?? 30),
    },
  }
}

export async function persistState(db: D1Database, state: AppState): Promise<void> {
  const now = new Date().toISOString()
  const statements: D1PreparedStatement[] = []
  for (const u of state.users)
    statements.push(
      db
        .prepare(
          `UPDATE users SET name=?,email=?,phone=?,active=?,sms_enabled=?,complaint_notifications_enabled=?,updated_at=? WHERE id=?`,
        )
        .bind(
          u.name,
          u.email,
          u.phone,
          u.active ? 1 : 0,
          u.smsEnabled ? 1 : 0,
          u.complaintNotificationsEnabled ? 1 : 0,
          now,
          u.id,
        ),
    )
  for (const c of state.complaints) {
    statements.push(
      db
        .prepare(
          `INSERT INTO complaints (id,external_case_id,store_id,assigned_manager_id,subject,complaint_text,category,severity,status,is_ack_overdue,is_resolution_overdue,routing_reason,routing_confidence,received_at,ingestion_mode,operational_started_at,dunkin_acknowledged_at,acknowledgment_body,manager_notified_at,manager_acknowledged_at,investigation_started_at,resolution_submitted_at,closed_at,closed_by,ack_deadline,resolution_deadline,manager_findings,customer_contacted,customer_contacted_at,customer_contact_outcome,corrective_action,resolution_notes,follow_ups,created_at,updated_at,source,gmail_message_id,gmail_thread_id,source_sender,customer_name,customer_email,customer_phone,occurrence_at,acknowledgement_status,owner_reviewed_at,reopened_at,reopen_reason,owner_notes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET store_id=excluded.store_id,assigned_manager_id=excluded.assigned_manager_id,status=excluded.status,is_ack_overdue=excluded.is_ack_overdue,is_resolution_overdue=excluded.is_resolution_overdue,ingestion_mode=excluded.ingestion_mode,operational_started_at=excluded.operational_started_at,dunkin_acknowledged_at=excluded.dunkin_acknowledged_at,acknowledgment_body=excluded.acknowledgment_body,acknowledgement_status=excluded.acknowledgement_status,manager_notified_at=excluded.manager_notified_at,manager_acknowledged_at=excluded.manager_acknowledged_at,investigation_started_at=excluded.investigation_started_at,resolution_submitted_at=excluded.resolution_submitted_at,closed_at=excluded.closed_at,closed_by=excluded.closed_by,owner_reviewed_at=excluded.owner_reviewed_at,reopened_at=excluded.reopened_at,reopen_reason=excluded.reopen_reason,owner_notes=excluded.owner_notes,manager_findings=excluded.manager_findings,customer_contacted=excluded.customer_contacted,customer_contacted_at=excluded.customer_contacted_at,customer_contact_outcome=excluded.customer_contact_outcome,corrective_action=excluded.corrective_action,resolution_notes=excluded.resolution_notes,follow_ups=excluded.follow_ups,updated_at=excluded.updated_at`,
        )
        .bind(
          c.id,
          c.externalCaseId,
          c.storeId ?? null,
          c.assignedManagerId ?? null,
          c.subject,
          c.complaintText,
          c.category,
          c.severity,
          c.status,
          c.isAckOverdue ? 1 : 0,
          c.isResolutionOverdue ? 1 : 0,
          c.routingReason,
          c.routingConfidence,
          c.receivedAt,
          c.ingestionMode,
          c.operationalStartedAt ?? null,
          c.dunkinAcknowledgedAt ?? null,
          c.acknowledgementBody || null,
          c.managerNotifiedAt ?? null,
          c.managerAcknowledgedAt ?? null,
          c.investigationStartedAt ?? null,
          c.resolutionSubmittedAt ?? null,
          c.closedAt ?? null,
          c.closedBy ?? null,
          c.ackDeadline,
          c.resolutionDeadline ?? null,
          c.managerFindings ?? null,
          c.customerContacted == null ? null : c.customerContacted ? 1 : 0,
          c.customerContactedAt ?? null,
          c.customerContactOutcome ?? null,
          c.correctiveAction ?? null,
          c.resolutionNotes ?? null,
          JSON.stringify(c.followUps),
          c.receivedAt,
          now,
          c.source ?? 'MANUAL',
          c.gmailMessageId ?? null,
          c.gmailThreadId ?? null,
          c.sourceSender ?? null,
          c.customerName ?? null,
          c.customerEmail ?? null,
          c.customerPhone ?? null,
          c.occurrenceAt ?? null,
          c.acknowledgementStatus ?? 'DISABLED',
          c.ownerReviewedAt ?? null,
          c.reopenedAt ?? null,
          c.reopenReason ?? null,
          c.ownerNotes ?? null,
        ),
    )
    for (const e of c.events)
      statements.push(
        db
          .prepare(
            `INSERT OR IGNORE INTO complaint_events (id,complaint_id,event_type,actor,timestamp,metadata) VALUES (?,?,?,?,?,?)`,
          )
          .bind(
            e.id,
            c.id,
            e.type,
            e.actor,
            e.timestamp,
            e.metadata ? JSON.stringify(e.metadata) : null,
          ),
      )
    for (const n of c.notifications)
      statements.push(
        db
          .prepare(
            `INSERT INTO notifications (id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,delivered_at,failed_at,failure_reason) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET status=excluded.status,provider_message_id=COALESCE(excluded.provider_message_id,notifications.provider_message_id),sent_at=COALESCE(excluded.sent_at,notifications.sent_at),delivered_at=COALESCE(excluded.delivered_at,notifications.delivered_at),failed_at=COALESCE(excluded.failed_at,notifications.failed_at),failure_reason=excluded.failure_reason`,
          )
          .bind(
            n.id,
            n.complaintId,
            n.eventType,
            n.recipientUserId,
            n.channel,
            n.message,
            n.status,
            n.provider,
            (n as Notification & { providerMessageId?: string }).providerMessageId ?? null,
            n.createdAt,
            n.sentAt ?? null,
            n.deliveredAt ?? null,
            n.failedAt ?? null,
            n.failureReason ?? null,
          ),
      )
  }
  statements.push(
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('notification_mode',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(state.config.mode, now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('external_notifications_enabled',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(state.config.externalNotificationsEnabled), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('pilot_store_id',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(state.config.pilotStoreId ?? '', now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('email_ingestion_enabled',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(Boolean(state.config.emailIngestionEnabled)), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('email_ack_enabled',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(Boolean(state.config.emailAckEnabled)), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('manager_ack_deadline_minutes',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(state.config.managerAckDeadlineMinutes ?? 30), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('manager_resolution_target_hours',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(state.config.managerResolutionTargetHours ?? 24), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('escalation_interval_minutes',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(state.config.escalationIntervalMinutes ?? 60), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('signalwire_reconcile_after_minutes',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(state.config.signalWireReconcileAfterMinutes ?? 10), now),
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('email_lookback_days',?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(String(state.config.emailLookbackDays ?? 30), now),
  )
  // Bump the state version so lightweight clients can poll for changes
  // without reloading all seven tables every 10 seconds.
  statements.push(
    db
      .prepare(
        `INSERT INTO settings(key,value,updated_at) VALUES('state_version', COALESCE((SELECT CAST(value AS INTEGER) FROM settings WHERE key='state_version'), 0) + 1, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_at=excluded.updated_at`,
      )
      .bind(now),
  )
  for (let i = 0; i < statements.length; i += 50) await db.batch(statements.slice(i, i + 50))
}
