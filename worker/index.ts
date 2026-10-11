import { Hono } from 'hono'
import { zValidator } from '@hono/zod-validator'
import { z } from 'zod'
import {
  actionSchema,
  complaintInputSchema,
  configSchema,
  contactSchema,
  pushSubscribeSchema,
  pushUnsubscribeSchema,
  responseApprovalSchema,
  responseDraftSchema,
  responseTemplateSchema,
  reviewAssignSchema,
  reviewResolveSchema,
  signalWireCallbackSchema,
  signalWireReconcileSchema,
  storeAdminSchema,
  testNotificationSchema,
  userAdminSchema,
} from '../src/lib/api-schema'
import { createComplaint, processDeadlines, updateComplaint } from '../src/lib/workflow'
import type { AppState, User } from '../src/lib/types'
import { authenticate, canAdmin, canViewComplaint, maskEmail, maskPhone } from './auth'
import { REVIEW_QUEUE_DETAIL_QUERY, REVIEW_QUEUE_LIST_QUERY } from './review-queue'
import { loadState, persistState, type D1Database } from './d1'
import { SignalWireSmsProvider, WebPushProvider } from './providers'
import { applySignalWireCallback, reconcileSignalWireMessage } from './callbacks'
import {
  maskGraphIdentifier,
  MicrosoftGraphProvider,
  pilotBodyDiagnosticCandidates,
} from './microsoft-graph'
import { buildPilotUniquenessReport } from './pilot-uniqueness'
import { buildPilotBodyDiagnosticReport } from './pilot-body-diagnostic'
import { EmailProviderError } from './email-provider'
import { ingestApprovedPilotCasePair, ingestSinglePilotComplaint } from './ingestion'
import { GraphDeltaDiscoveryClient, runMailboxDiscovery } from './mail-discovery'
import {
  ResponseWorkflowError,
  approveResponseAction,
  createResponseDraft,
  createResponseTemplate,
  listResponseTemplates,
  rejectResponseAction,
  sendResponseAction,
  submitResponseForApproval,
} from './response-workflow'
import { runScheduledOperations } from './operations'
import { buildReport } from './reporting'

type Bindings = Env & {
  DEV_AUTH_USER_ID?: string
  SIGNALWIRE_PROJECT_ID?: string
  SIGNALWIRE_API_TOKEN?: string
  SIGNALWIRE_SPACE_URL?: string
  SIGNALWIRE_PHONE_NUMBER?: string
  PUBLIC_BASE_URL?: string
  VAPID_PUBLIC_KEY?: string
  VAPID_PRIVATE_KEY?: string
  MS_CLIENT_ID?: string
  MS_CLIENT_SECRET?: string
  MS_REFRESH_TOKEN?: string
  MS_MAILBOX_ADDRESS?: string
  MS_TENANT?: string
  PILOT_MAIL_TOOLING_ENABLED?: string
}
type Variables = { user: User }
const app = new Hono<{ Bindings: Bindings; Variables: Variables }>()
const ownerIds = ['father', 'uncle', 'grandfather']
const pilotAdminId = 'pilot-admin'
const testRecipientIds = [...ownerIds, pilotAdminId, 'dad']
const jsonError = (message: string) => ({ error: message })
const provider = (env: Bindings) =>
  new SignalWireSmsProvider({
    projectId: env.SIGNALWIRE_PROJECT_ID,
    apiToken: env.SIGNALWIRE_API_TOKEN,
    spaceUrl: env.SIGNALWIRE_SPACE_URL,
    from: env.SIGNALWIRE_PHONE_NUMBER,
    publicBaseUrl: env.PUBLIC_BASE_URL,
  })
const webPush = (env: Bindings) =>
  new WebPushProvider({
    publicKey: env.VAPID_PUBLIC_KEY,
    privateKey: env.VAPID_PRIVATE_KEY,
    subject: env.PUBLIC_BASE_URL ?? 'https://store-resolve.moogzie.workers.dev',
  })
const emailProvider = (env: Bindings) =>
  new MicrosoftGraphProvider({
    clientId: env.MS_CLIENT_ID,
    clientSecret: env.MS_CLIENT_SECRET,
    refreshToken: env.MS_REFRESH_TOKEN,
    mailboxAddress: env.MS_MAILBOX_ADDRESS,
    tenant: env.MS_TENANT,
  })
const auditAdminChange = async (
  db: D1Database,
  eventType: string,
  entityId: string,
  actorId: string,
) =>
  db
    .prepare(
      `INSERT INTO integration_events(id,integration,event_type,entity_id,outcome,detail_code,metadata,created_at) VALUES(?,'ADMIN',?,?,'SUCCESS',NULL,?,?)`,
    )
    .bind(
      crypto.randomUUID(),
      eventType,
      entityId,
      JSON.stringify({ actorId }),
      new Date().toISOString(),
    )
    .run()
const publicState = (state: AppState, user: User): AppState => ({
  ...state,
  activeUserId: user.id,
  users: state.users.map((u) => ({ ...u, email: maskEmail(u.email), phone: maskPhone(u.phone) })),
  complaints: state.complaints.filter((c) =>
    canViewComplaint(user, c.storeId, c.assignedManagerId),
  ),
})

app.get('/api/health', (c) => c.json({ ok: true, service: 'store-resolve' }))
app.post('/api/signalwire/status', async (c) => {
  let payload: unknown
  try {
    payload = c.req.header('content-type')?.includes('application/json')
      ? await c.req.json()
      : await c.req.parseBody()
  } catch {
    return c.json(jsonError('Invalid callback'), 400)
  }
  const parsed = signalWireCallbackSchema.safeParse(payload)
  if (!parsed.success) return c.json(jsonError('Invalid callback'), 400)
  if ('project_id' in parsed.data && parsed.data.project_id !== c.env.SIGNALWIRE_PROJECT_ID)
    return c.json(jsonError('SignalWire verification failed'), 403)
  const signalWire = provider(c.env)
  if (!signalWire.ready) return c.json(jsonError('Provider unavailable'), 503)
  let message
  try {
    message = await signalWire.retrieveMessage(
      'id' in parsed.data ? parsed.data.id : parsed.data.MessageSid,
    )
  } catch {
    return c.json(jsonError('SignalWire verification failed'), 403)
  }
  const result = await applySignalWireCallback(
    c.env.DB,
    message,
    c.env.SIGNALWIRE_PROJECT_ID!,
    c.env.SIGNALWIRE_PHONE_NUMBER!,
  )
  if (result === 'AUTHENTICITY_FAILED')
    return c.json(jsonError('SignalWire verification failed'), 403)
  if (result === 'NOT_FOUND') return c.json(jsonError('Notification not found'), 404)
  return c.json({ ok: true, result })
})

app.use('/api/*', async (c, next) => {
  const user = await authenticate(c.req.raw, c.env.DB, c.env)
  if (!user) return c.json(jsonError('Unauthorized'), 401)
  c.set('user', user)
  await next()
})
app.get('/api/bootstrap', async (c) => {
  const state = await loadState(c.env.DB)
  const user = c.get('user')
  const signalWire = provider(c.env)
  let providerReady = false
  if (signalWire.ready)
    try {
      await signalWire.verifyConnection()
      providerReady = true
    } catch {
      providerReady = false
    }
  const graph = emailProvider(c.env)
  let microsoftGraphReady = false
  if (graph.ready)
    try {
      await graph.verifyConnection()
      microsoftGraphReady = true
    } catch {
      microsoftGraphReady = false
    }
  const [lastEmailSync, lastBackgroundRun, failures] = await Promise.all([
    c.env.DB.prepare(
      `SELECT completed_at FROM background_job_runs WHERE job_name='OPERATIONS' AND outcome IN ('SUCCESS','PARTIAL') ORDER BY started_at DESC LIMIT 1`,
    ).first<{ completed_at: string }>(),
    c.env.DB.prepare(
      `SELECT completed_at FROM background_job_runs ORDER BY started_at DESC LIMIT 1`,
    ).first<{ completed_at: string }>(),
    c.env.DB.prepare(
      `SELECT integration,detail_code,created_at FROM integration_events WHERE outcome='FAILED' ORDER BY created_at DESC LIMIT 10`,
    ).all<{ integration: string; detail_code: string; created_at: string }>(),
  ])
  return c.json({
    state: publicState(state, user),
    currentUser: {
      ...user,
      phone: maskPhone(user.phone),
      email: maskEmail(user.email),
      maskedPhone: maskPhone(user.phone),
      maskedEmail: maskEmail(user.email),
    },
    providerReady,
    health: {
      microsoftGraphReady,
      emailIngestionEnabled: Boolean(state.config.emailIngestionEnabled),
      emailAckEnabled: Boolean(state.config.emailAckEnabled),
      lastEmailSyncAt: lastEmailSync?.completed_at,
      signalWireReady: providerReady,
      externalNotificationsEnabled: state.config.externalNotificationsEnabled,
      rolloutMode: state.config.mode,
      lastBackgroundRunAt: lastBackgroundRun?.completed_at,
      recentFailures: failures.results.map((row) => ({
        integration: row.integration,
        code: row.detail_code,
        occurredAt: row.created_at,
      })),
    },
  })
})
app.post(
  '/api/admin/signalwire/reconcile',
  zValidator('json', signalWireReconcileSchema),
  async (c) => {
    const user = c.get('user')
    if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
    const state = await loadState(c.env.DB)
    if (state.config.externalNotificationsEnabled)
      return c.json(jsonError('External notifications must be disabled'), 409)
    if (state.config.mode !== 'FAMILY_PILOT')
      return c.json(jsonError('FAMILY_PILOT mode required'), 409)
    const signalWire = provider(c.env)
    if (!signalWire.ready) return c.json(jsonError('Provider unavailable'), 503)
    let message
    try {
      message = await signalWire.retrieveMessage(c.req.valid('json').providerMessageId)
    } catch {
      return c.json(jsonError('SignalWire verification failed'), 403)
    }
    const result = await reconcileSignalWireMessage(
      c.env.DB,
      message,
      c.env.SIGNALWIRE_PROJECT_ID!,
      c.env.SIGNALWIRE_PHONE_NUMBER!,
      pilotAdminId,
    )
    if (result === 'AUTHENTICITY_FAILED')
      return c.json(jsonError('SignalWire verification failed'), 403)
    if (result === 'NOT_FOUND') return c.json(jsonError('Notification not found'), 404)
    if (result === 'NON_TERMINAL') return c.json(jsonError('Provider state is not terminal'), 409)
    return c.json({ ok: true, result, providerStatus: message.status.toUpperCase() })
  },
)
app.post('/api/complaints', zValidator('json', complaintInputSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  const result = createComplaint(state, c.req.valid('json'))
  await persistState(c.env.DB, result.state)
  await dispatchEligible(c.env, result.state)
  return c.json(publicState(await loadState(c.env.DB), user))
})
app.post('/api/complaints/:id/actions', zValidator('json', actionSchema), async (c) => {
  const user = c.get('user')
  const state = await loadState(c.env.DB)
  const complaint = state.complaints.find((x) => x.id === c.req.param('id'))
  if (!complaint) return c.json(jsonError('Complaint not found'), 404)
  if (!canViewComplaint(user, complaint.storeId, complaint.assignedManagerId))
    return c.json(jsonError('Forbidden'), 403)
  try {
    const body = c.req.valid('json')
    const next = updateComplaint(state, complaint.id, body.action, user.id, body.data)
    await persistState(c.env.DB, next)
    await dispatchEligible(c.env, next)
    return c.json(publicState(await loadState(c.env.DB), user))
  } catch (error) {
    return c.json(jsonError((error as Error).message), 403)
  }
})
app.post(
  '/api/admin/deadlines',
  zValidator('json', z.object({ advanceHours: z.number().min(0).max(720).default(0) })),
  async (c) => {
    const user = c.get('user')
    if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
    const state = await loadState(c.env.DB)
    const now = new Date(Date.now() + c.req.valid('json').advanceHours * 3600000).toISOString()
    const next = processDeadlines(state, now)
    await persistState(c.env.DB, next)
    await dispatchEligible(c.env, next)
    return c.json(publicState(await loadState(c.env.DB), user))
  },
)
app.put('/api/admin/config', zValidator('json', configSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  state.config = c.req.valid('json')
  await persistState(c.env.DB, state)
  await auditAdminChange(c.env.DB, 'SETTINGS_CHANGED', 'operational-settings', user.id)
  return c.json(publicState(await loadState(c.env.DB), user))
})

// Historical pilot tools remain available for audited recovery only. They are isolated from the
// production surface unless an operator deliberately enables the server-side binding.
app.use('/api/admin/email/pilot-*', async (c, next) => {
  if (c.env.PILOT_MAIL_TOOLING_ENABLED !== 'true') return c.json(jsonError('Not found'), 404)
  await next()
})

app.post('/api/admin/email/pilot-ingest', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  try {
    const result = await ingestSinglePilotComplaint(c.env.DB, emailProvider(c.env), state.config)
    return c.json({ ok: true, ...result })
  } catch (error) {
    const code = error instanceof EmailProviderError ? error.code : 'MS_GRAPH_PILOT_FAILED'
    return c.json(jsonError(code), 409)
  }
})

app.get('/api/admin/email/readiness', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const graph = emailProvider(c.env)
  if (!graph.ready)
    return c.json({ ok: false, ready: false, error: 'MS_GRAPH_NOT_CONFIGURED' }, 503)
  try {
    await graph.verifyConnection()
    return c.json({ ok: true, ready: true, check: 'INBOX_FOLDER_METADATA' })
  } catch (error) {
    const code = error instanceof EmailProviderError ? error.code : 'MS_GRAPH_READINESS_FAILED'
    return c.json({ ok: false, ready: false, error: code }, 503)
  }
})

app.get('/api/admin/email/pilot-diagnostic', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  if (
    state.config.mode !== 'FAMILY_PILOT' ||
    state.config.externalNotificationsEnabled ||
    state.config.emailIngestionEnabled ||
    state.config.emailAckEnabled
  )
    return c.json(jsonError('Production safety controls must remain locked down'), 409)
  const graph = emailProvider(c.env)
  if (!graph.ready)
    return c.json({ ok: false, ready: false, error: 'MS_GRAPH_NOT_CONFIGURED' }, 503)
  try {
    await graph.verifyConnection()
    const dateOnly = await graph.runDateOnlyIdentityDiagnostic()
    const senderOnly = await graph.runSenderOnlyIdentityDiagnostic()
    const serialize = (record: {
      id: string
      receivedDateTime: string
      senderAddress: string
    }) => ({
      messageId: maskGraphIdentifier(record.id),
      receivedDateTime: record.receivedDateTime,
      sender: maskEmail(record.senderAddress),
    })
    const result = {
      ok: true,
      scope: 'MAILBOX_IDENTITY_METADATA_ONLY',
      dateOnly: {
        count: dateOnly.records.length,
        hasMore: dateOnly.hasMore,
        records: dateOnly.records.map(serialize),
      },
      senderOnly: {
        count: senderOnly.records.length,
        hasMore: senderOnly.hasMore,
        records: senderOnly.records.map(serialize),
      },
    }
    if (c.req.header('accept')?.includes('text/html')) {
      const escaped = JSON.stringify(result, null, 2)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
      return c.html(`<pre>${escaped}</pre>`)
    }
    return c.json(result)
  } catch (error) {
    const code = error instanceof EmailProviderError ? error.code : 'MS_GRAPH_DIAGNOSTIC_FAILED'
    return c.json({ ok: false, error: code }, 503)
  }
})

app.get('/api/admin/email/pilot-uniqueness-diagnostic', (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  return c.html(`<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>StoreResolve metadata diagnostic</title></head>
<body><main><h1>Metadata-only uniqueness diagnostic</h1>
<p>This reads at most ten August 1 message metadata records. It does not read message bodies or run ingestion.</p>
<form method="post" action="/api/admin/email/pilot-uniqueness-diagnostic">
<button type="submit">Run metadata diagnostic once</button>
</form></main></body></html>`)
})

app.post('/api/admin/email/pilot-uniqueness-diagnostic', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  if (
    state.config.mode !== 'FAMILY_PILOT' ||
    state.config.externalNotificationsEnabled ||
    state.config.emailIngestionEnabled ||
    state.config.emailAckEnabled
  )
    return c.json(jsonError('Production safety controls must remain locked down'), 409)
  const graph = emailProvider(c.env)
  if (!graph.ready)
    return c.json({ ok: false, ready: false, error: 'MS_GRAPH_NOT_CONFIGURED' }, 503)
  try {
    await graph.verifyConnection()
    const metadata = await graph.listPilotUniquenessMetadata()
    const matches = metadata.records.filter(
      (record) =>
        record.senderMatched &&
        record.caseIdMatched &&
        record.subjectPhraseMatched &&
        record.storeNumberMatched,
    )
    const folderCache = new Map<string, string>()
    const matchingWithFolders = []
    for (const record of matches) {
      let folderName = folderCache.get(record.parentFolderId)
      if (!folderName) {
        folderName = (await graph.getMailFolderMetadata(record.parentFolderId)).displayName
        folderCache.set(record.parentFolderId, folderName)
      }
      matchingWithFolders.push({ ...record, folderName })
    }
    const report = await buildPilotUniquenessReport(metadata, matchingWithFolders)
    const runId = crypto.randomUUID()
    await c.env.DB.prepare(
      `INSERT INTO integration_events(id,integration,event_type,entity_id,outcome,detail_code,metadata,created_at) VALUES(?,'MICROSOFT_GRAPH','PILOT_UNIQUENESS_DIAGNOSTIC',?, 'SUCCESS',?,?,?)`,
    )
      .bind(
        crypto.randomUUID(),
        runId,
        report.classification,
        JSON.stringify(report),
        new Date().toISOString(),
      )
      .run()
    return c.json({ ok: true, diagnosticRunId: runId, ...report })
  } catch (error) {
    const code = error instanceof EmailProviderError ? error.code : 'MS_GRAPH_DIAGNOSTIC_FAILED'
    return c.json({ ok: false, error: code }, 503)
  }
})

app.get('/api/admin/email/pilot-uniqueness-diagnostic/latest', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const row = await c.env.DB.prepare(
    `SELECT entity_id,metadata,created_at FROM integration_events WHERE integration='MICROSOFT_GRAPH' AND event_type='PILOT_UNIQUENESS_DIAGNOSTIC' ORDER BY created_at DESC LIMIT 1`,
  ).first<{ entity_id: string; metadata: string; created_at: string }>()
  if (!row) return c.json(jsonError('No durable uniqueness diagnostic exists'), 404)
  return c.json({
    ok: true,
    diagnosticRunId: row.entity_id,
    createdAt: row.created_at,
    ...JSON.parse(row.metadata),
  })
})

app.get('/api/admin/email/pilot-body-diagnostic', (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  return c.html(`<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>StoreResolve body diagnostic</title></head>
<body><main><h1>Two-message body diagnostic</h1>
<p>This reads only the two approved case CCC11122413 messages. It does not access attachments or run ingestion.</p>
<form method="post" action="/api/admin/email/pilot-body-diagnostic">
<button type="submit">Run approved body diagnostic once</button>
</form></main></body></html>`)
})

app.post('/api/admin/email/pilot-body-diagnostic', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  if (
    state.config.mode !== 'FAMILY_PILOT' ||
    state.config.externalNotificationsEnabled ||
    state.config.emailIngestionEnabled ||
    state.config.emailAckEnabled
  )
    return c.json(jsonError('Production safety controls must remain locked down'), 409)
  const graph = emailProvider(c.env)
  if (!graph.ready)
    return c.json({ ok: false, ready: false, error: 'MS_GRAPH_NOT_CONFIGURED' }, 503)
  try {
    await graph.verifyConnection()
    const candidates = (await graph.findPilotBodyDiagnosticCandidates()).sort(
      (left, right) => Date.parse(left.receivedDateTime) - Date.parse(right.receivedDateTime),
    )
    const identitiesMatch =
      candidates.length === pilotBodyDiagnosticCandidates.length &&
      candidates.every(
        (candidate, index) =>
          maskGraphIdentifier(candidate.id) === pilotBodyDiagnosticCandidates[index].maskedId &&
          candidate.receivedDateTime === pilotBodyDiagnosticCandidates[index].receivedDateTime,
      )
    if (!identitiesMatch)
      return c.json({ ok: false, error: 'MS_GRAPH_BODY_DIAGNOSTIC_SELECTION_CHANGED' }, 409)
    const fetched = await Promise.all(
      candidates.map((candidate) => graph.getPilotBodyDiagnosticMessage(candidate.id)),
    )
    const bodyIdentitiesMatch = fetched.every(
      (message, index) =>
        message.id === candidates[index].id &&
        message.threadId === candidates[index].conversationId &&
        message.internalDate === new Date(candidates[index].receivedDateTime).toISOString() &&
        (message.sender.trim().toLowerCase() === 'customerservice@dunkinbrands.com' ||
          message.sender.toLowerCase().endsWith('<customerservice@dunkinbrands.com>')),
    )
    if (!bodyIdentitiesMatch)
      return c.json({ ok: false, error: 'MS_GRAPH_BODY_DIAGNOSTIC_IDENTITY_MISMATCH' }, 409)
    const report = await buildPilotBodyDiagnosticReport([fetched[0], fetched[1]])
    const runId = crypto.randomUUID()
    await c.env.DB.prepare(
      `INSERT INTO integration_events(id,integration,event_type,entity_id,outcome,detail_code,metadata,created_at) VALUES(?,'MICROSOFT_GRAPH','PILOT_BODY_DIAGNOSTIC',?,'SUCCESS',?,?,?)`,
    )
      .bind(
        crypto.randomUUID(),
        runId,
        report.classification,
        JSON.stringify(report),
        new Date().toISOString(),
      )
      .run()
    return c.json({ ok: true, diagnosticRunId: runId, ...report })
  } catch (error) {
    const code =
      error instanceof EmailProviderError
        ? error.code
        : error instanceof Error
          ? error.message
          : 'MS_GRAPH_BODY_DIAGNOSTIC_FAILED'
    return c.json({ ok: false, error: code }, 503)
  }
})

app.get('/api/admin/email/pilot-case-pair-ingest', (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  return c.html(`<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>StoreResolve controlled case ingestion</title></head>
<body><main><h1>Controlled case CCC11122413 ingestion</h1>
<p>This ingests only the approved initial source AAkA…DQAA and follow-up AAkA…EQAA. Acknowledgments and external SMS must remain off.</p>
<form method="post" action="/api/admin/email/pilot-case-pair-ingest">
<button type="submit">Ingest the approved source pair once</button>
</form></main></body></html>`)
})

app.post('/api/admin/email/pilot-case-pair-ingest', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  try {
    const result = await ingestApprovedPilotCasePair(c.env.DB, emailProvider(c.env), state.config)
    return c.json({ ok: true, ...result })
  } catch (error) {
    const code = error instanceof EmailProviderError ? error.code : 'MS_GRAPH_PILOT_PAIR_FAILED'
    return c.json(jsonError(code), 409)
  }
})

app.get('/diagnostics/pilot-mail', (c) => {
  if (c.env.PILOT_MAIL_TOOLING_ENABLED !== 'true') return c.text('Not found', 404)
  const url = new URL(c.req.url)
  url.pathname = '/api/admin/email/pilot-diagnostic'
  return app.fetch(new Request(url, { headers: c.req.raw.headers }), c.env)
})

app.get('/api/admin/reporting', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  return c.json(
    buildReport(state, {
      from: c.req.query('from'),
      to: c.req.query('to'),
      storeId: c.req.query('storeId'),
      status: c.req.query('status'),
      category: c.req.query('category'),
      severity: c.req.query('severity'),
    }),
  )
})

app.put('/api/admin/stores/:id', zValidator('json', storeAdminSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const id = c.req.param('id').trim()
  if (!id || id.length > 100) return c.json(jsonError('Invalid store identifier'), 400)
  const input = c.req.valid('json')
  const now = new Date().toISOString()
  const manager = input.managerId
    ? await c.env.DB.prepare(
        `SELECT id FROM users WHERE id=? AND role='STORE_MANAGER' AND active=1`,
      )
        .bind(input.managerId)
        .first<{ id: string }>()
    : null
  if (input.managerId && !manager) return c.json(jsonError('Active manager required'), 400)
  await c.env.DB.prepare(
    `INSERT INTO stores(id,organization_id,dunkin_store_number,name,address,city,state,postal_code,phone,active,manager_id,created_at,updated_at) VALUES(?,'org-1',?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET dunkin_store_number=excluded.dunkin_store_number,name=excluded.name,address=excluded.address,city=excluded.city,state=excluded.state,postal_code=excluded.postal_code,phone=excluded.phone,active=excluded.active,manager_id=excluded.manager_id,updated_at=excluded.updated_at`,
  )
    .bind(
      id,
      input.number,
      input.name,
      input.address,
      input.city,
      input.state,
      input.postalCode,
      input.phone,
      input.active ? 1 : 0,
      input.managerId ?? null,
      now,
      now,
    )
    .run()
  await c.env.DB.prepare('DELETE FROM store_aliases WHERE store_id=?').bind(id).run()
  if (input.aliases.length)
    await c.env.DB.batch(
      input.aliases.map((alias) =>
        c.env.DB.prepare(
          `INSERT INTO store_aliases(id,store_id,alias_normalized,created_at) VALUES(?,?,?,?)`,
        ).bind(crypto.randomUUID(), id, alias.toLowerCase().replace(/[^a-z0-9]/g, ''), now),
      ),
    )
  await auditAdminChange(c.env.DB, 'STORE_CHANGED', id, user.id)
  return c.json(publicState(await loadState(c.env.DB), user))
})

app.put('/api/admin/users/:id', zValidator('json', userAdminSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const id = c.req.param('id').trim()
  if (!id || id.length > 100 || id === pilotAdminId)
    return c.json(jsonError('Invalid managed user identifier'), 400)
  const input = c.req.valid('json')
  const now = new Date().toISOString()
  const existing = await c.env.DB.prepare('SELECT * FROM users WHERE id=?')
    .bind(id)
    .first<Record<string, unknown>>()
  if (!existing) return c.json(jsonError('User not found'), 404)
  await c.env.DB.prepare(
    `INSERT INTO users(id,organization_id,name,email,phone,role,active,sms_enabled,timezone,created_at,updated_at,recipient_kind,complaint_notifications_enabled) VALUES(?,'org-1',?,?,?,?,?,?,?, ?,?,'STANDARD',?) ON CONFLICT(id) DO UPDATE SET name=excluded.name,email=excluded.email,phone=excluded.phone,role=excluded.role,active=excluded.active,sms_enabled=excluded.sms_enabled,timezone=excluded.timezone,complaint_notifications_enabled=excluded.complaint_notifications_enabled,updated_at=excluded.updated_at`,
  )
    .bind(
      id,
      input.name ?? String(existing.name),
      input.email ?? String(existing.email),
      input.phone ?? String(existing.phone),
      input.role ?? String(existing.role),
      (input.active ?? Boolean(existing.active)) ? 1 : 0,
      (input.smsEnabled ?? Boolean(existing.sms_enabled)) ? 1 : 0,
      input.timezone ?? String(existing.timezone),
      now,
      now,
      (input.complaintNotificationsEnabled ?? Boolean(existing.complaint_notifications_enabled))
        ? 1
        : 0,
    )
    .run()
  if (input.storeIds) {
    await c.env.DB.prepare('DELETE FROM user_store_assignments WHERE user_id=?').bind(id).run()
  }
  if ((input.role ?? existing.role) === 'STORE_MANAGER' && input.storeIds?.length)
    await c.env.DB.batch(
      input.storeIds.map((storeId) =>
        c.env.DB.prepare(
          `INSERT INTO user_store_assignments(user_id,store_id,created_at) SELECT ?,id,? FROM stores WHERE id=?`,
        ).bind(id, now, storeId),
      ),
    )
  await auditAdminChange(c.env.DB, 'USER_CHANGED', id, user.id)
  return c.json(publicState(await loadState(c.env.DB), user))
})
app.put('/api/admin/contacts/:id', zValidator('json', contactSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const id = c.req.param('id')
  if (!testRecipientIds.includes(id))
    return c.json(jsonError('Only controlled recipients can be configured'), 400)
  const state = await loadState(c.env.DB)
  const target = state.users.find((u) => u.id === id)
  if (!target) return c.json(jsonError('Recipient not found'), 404)
  if (id === pilotAdminId) {
    if (target.recipientKind !== 'PILOT_ADMIN')
      return c.json(jsonError('Pilot recipient capability is unavailable'), 400)
    if (state.config.mode !== 'FAMILY_PILOT')
      return c.json(jsonError('Pilot recipient requires FAMILY_PILOT mode'), 403)
  }
  const contact = c.req.valid('json')
  target.name = contact.name
  target.smsEnabled = contact.smsEnabled
  target.active = contact.active
  if (contact.email) target.email = contact.email
  if (contact.phone) target.phone = contact.phone
  await persistState(c.env.DB, state)
  return c.json(publicState(await loadState(c.env.DB), user))
})
// ---------------------------------------------------------------------------
// Stage A: Web Push subscriptions and gated test delivery.
//
// All endpoints sit behind the existing Cloudflare Access authentication
// (app.use('/api/*')); subscriptions are always bound to the authenticated
// user. Real-recipient test sends are disabled by the `push_test_sends_enabled`
// settings kill switch until explicitly approved.
// ---------------------------------------------------------------------------

app.get('/api/push/public-key', (c) => {
  if (!c.env.VAPID_PUBLIC_KEY)
    return c.json(jsonError('Push notifications are not configured'), 503)
  return c.json({ publicKey: c.env.VAPID_PUBLIC_KEY })
})

app.post('/api/push/subscribe', zValidator('json', pushSubscribeSchema), async (c) => {
  const user = c.get('user')
  const { endpoint, keys, deviceLabel } = c.req.valid('json')
  const existing = await c.env.DB.prepare(
    'SELECT id, user_id, revoked_at FROM push_subscriptions WHERE endpoint=?',
  )
    .bind(endpoint)
    .first<{ id: string; user_id: string; revoked_at: string | null }>()
  if (existing && existing.user_id !== user.id)
    return c.json(jsonError('This device is already registered to another user'), 409)
  const now = new Date().toISOString()
  if (existing) {
    await c.env.DB.prepare(
      `UPDATE push_subscriptions SET p256dh_key=?, auth_key=?, device_label=COALESCE(?, device_label), revoked_at=NULL, last_used_at=? WHERE id=?`,
    )
      .bind(keys.p256dh, keys.auth, deviceLabel ?? null, now, existing.id)
      .run()
  } else {
    const org = await c.env.DB.prepare('SELECT organization_id FROM users WHERE id=?')
      .bind(user.id)
      .first<{ organization_id: string }>()
    await c.env.DB.prepare(
      `INSERT INTO push_subscriptions(id,organization_id,user_id,endpoint,p256dh_key,auth_key,device_label,created_at) VALUES(?,?,?,?,?,?,?,?)`,
    )
      .bind(
        crypto.randomUUID(),
        org?.organization_id ?? 'org-1',
        user.id,
        endpoint,
        keys.p256dh,
        keys.auth,
        deviceLabel ?? null,
        now,
      )
      .run()
  }
  return c.json({ ok: true })
})

app.post('/api/push/unsubscribe', zValidator('json', pushUnsubscribeSchema), async (c) => {
  const user = c.get('user')
  const { endpoint } = c.req.valid('json')
  await c.env.DB.prepare(
    `UPDATE push_subscriptions SET revoked_at=? WHERE endpoint=? AND user_id=? AND revoked_at IS NULL`,
  )
    .bind(new Date().toISOString(), endpoint, user.id)
    .run()
  return c.json({ ok: true })
})

app.post('/api/admin/push/test', zValidator('json', testNotificationSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const { recipientUserId } = c.req.valid('json')
  const state = await loadState(c.env.DB)
  const recipient = state.users.find((u) => u.id === recipientUserId)
  if (!recipient || !testRecipientIds.includes(recipient.id))
    return c.json(jsonError('Controlled test recipient required'), 400)
  // Kill switch: real-recipient push sends stay disabled until explicitly approved.
  const killSwitch = await c.env.DB.prepare(`SELECT value FROM settings WHERE key='push_test_sends_enabled'`)
    .first<{ value: string }>()
  if (killSwitch?.value !== 'true')
    return c.json(jsonError('Push test sends are disabled by the operator kill switch'), 409)
  const last = await c.env.DB.prepare(
    'SELECT last_sent_at FROM test_notification_rate_limits WHERE recipient_user_id=?',
  )
    .bind(recipient.id)
    .first<{ last_sent_at: string }>()
  if (last && Date.now() - new Date(last.last_sent_at).getTime() < 300000)
    return c.json(jsonError('Please wait five minutes before another test to this recipient'), 429)
  const subscription = await c.env.DB.prepare(
    `SELECT endpoint, p256dh_key, auth_key FROM push_subscriptions WHERE user_id=? AND revoked_at IS NULL ORDER BY last_used_at DESC LIMIT 1`,
  )
    .bind(recipient.id)
    .first<{ endpoint: string; p256dh_key: string; auth_key: string }>()
  const now = new Date().toISOString()
  const id = `pushtest-${crypto.randomUUID()}`
  let status = 'SUPPRESSED',
    reason: string | undefined
  if (!webPush(c.env).ready) reason = 'WEBPUSH_NOT_READY'
  else if (!subscription) reason = 'NO_PUSH_SUBSCRIPTION'
  else {
    const outstanding = state.complaints.filter(
      (complaint) =>
        complaint.assignedManagerId === recipient.id &&
        !['CLOSED'].includes(complaint.status),
    ).length
    try {
      const result = await webPush(c.env).sendToSubscription(
        {
          endpoint: subscription.endpoint,
          p256dhKey: subscription.p256dh_key,
          authKey: subscription.auth_key,
        },
        {
          title: 'StoreResolve test',
          body: 'TEST — push notifications are connected to this device. No action needed.',
          url: '/',
          tag: 'storeresolve-test-push',
          badgeCount: outstanding,
        },
      )
      if (result.expired) {
        status = 'FAILED'
        reason = 'SUBSCRIPTION_EXPIRED'
        await c.env.DB.prepare(
          `UPDATE push_subscriptions SET revoked_at=? WHERE endpoint=? AND revoked_at IS NULL`,
        )
          .bind(now, subscription.endpoint)
          .run()
      } else {
        status = 'SENT'
        await c.env.DB.prepare(`UPDATE push_subscriptions SET last_used_at=? WHERE endpoint=?`)
          .bind(now, subscription.endpoint)
          .run()
      }
    } catch (error) {
      status = 'FAILED'
      reason = (error as Error).message
    }
  }
  await c.env.DB.prepare(
    `INSERT INTO notifications(id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,failed_at,failure_reason) VALUES(?,NULL,'TEST',?,'PUSH',?,?,?,?,?,?, ?,?)`,
  )
    .bind(
      id,
      recipient.id,
      'StoreResolve test push notification',
      status,
      'WEBPUSH',
      null,
      now,
      status === 'SENT' ? now : null,
      status === 'FAILED' ? now : null,
      reason ?? null,
    )
    .run()
  await c.env.DB.prepare(
    'INSERT INTO test_notification_rate_limits(recipient_user_id,last_sent_at) VALUES(?,?) ON CONFLICT(recipient_user_id) DO UPDATE SET last_sent_at=excluded.last_sent_at',
  )
    .bind(recipient.id, now)
    .run()
  return c.json(publicState(await loadState(c.env.DB), user))
})

app.post('/api/admin/test-notifications', zValidator('json', testNotificationSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const { recipientUserId } = c.req.valid('json')
  const state = await loadState(c.env.DB)
  const recipient = state.users.find((u) => u.id === recipientUserId)
  if (!recipient || !testRecipientIds.includes(recipient.id))
    return c.json(jsonError('Controlled test recipient required'), 400)
  if (recipient.id === pilotAdminId && recipient.recipientKind !== 'PILOT_ADMIN')
    return c.json(jsonError('Pilot recipient capability is unavailable'), 400)
  const last = await c.env.DB.prepare(
    'SELECT last_sent_at FROM test_notification_rate_limits WHERE recipient_user_id=?',
  )
    .bind(recipient.id)
    .first<{ last_sent_at: string }>()
  if (last && Date.now() - new Date(last.last_sent_at).getTime() < 300000)
    return c.json(jsonError('Please wait five minutes before another test to this recipient'), 429)
  const now = new Date().toISOString()
  const id = `test-${crypto.randomUUID()}`
  const message =
    'STORERESOLVE TEST\nStoreResolve complaint alerts are connected to this phone.\nNo action is required.'
  let status = 'SUPPRESSED',
    reason: string | undefined,
    providerMessageId: string | undefined
  if (!state.config.externalNotificationsEnabled) reason = 'EXTERNAL_NOTIFICATIONS_DISABLED'
  else if (state.config.mode !== 'FAMILY_PILOT') reason = 'FAMILY_PILOT_REQUIRED'
  else if (!recipient.active || !recipient.smsEnabled || !recipient.phone)
    reason = 'RECIPIENT_NOT_READY'
  else if (!provider(c.env).ready) reason = 'SIGNALWIRE_NOT_READY'
  else {
    try {
      const result = await provider(c.env).send(
        {
          id,
          complaintId: '',
          eventType: 'TEST',
          recipientUserId: recipient.id,
          channel: 'SMS',
          message,
          status: 'PENDING',
          provider: 'SIGNALWIRE',
          createdAt: now,
        },
        recipient.phone,
      )
      status = result.status
      providerMessageId = result.providerMessageId
    } catch (error) {
      status = 'FAILED'
      reason = (error as Error).message
    }
  }
  await c.env.DB.prepare(
    `INSERT INTO notifications(id,complaint_id,event_type,recipient_user_id,channel,message,status,provider,provider_message_id,created_at,sent_at,failed_at,failure_reason) VALUES(?,NULL,'TEST',?,'SMS',?,?,'SIGNALWIRE',?,?,?, ?,?)`,
  )
    .bind(
      id,
      recipient.id,
      message,
      status,
      providerMessageId ?? null,
      now,
      status === 'SENT' ? now : null,
      status === 'FAILED' ? now : null,
      reason ?? null,
    )
    .run()
  await c.env.DB.prepare(
    'INSERT INTO test_notification_rate_limits(recipient_user_id,last_sent_at) VALUES(?,?) ON CONFLICT(recipient_user_id) DO UPDATE SET last_sent_at=excluded.last_sent_at',
  )
    .bind(recipient.id, now)
    .run()
  return c.json(publicState(await loadState(c.env.DB), user))
})

export async function dispatchEligible(env: Bindings, state: AppState) {
  for (const complaint of state.complaints)
    for (const n of complaint.notifications) {
      if (n.status !== 'PENDING' || n.providerMessageId) continue
      const recipient = state.users.find((u) => u.id === n.recipientUserId)
      if (!recipient) continue
      let reason: string | undefined
      if (!state.config.externalNotificationsEnabled) reason = 'EXTERNAL_NOTIFICATIONS_DISABLED'
      else if (recipient.recipientKind === 'PILOT_ADMIN') reason = 'TEST_RECIPIENT_ONLY'
      else if (state.config.mode === 'FAMILY_PILOT' && !recipient.complaintNotificationsEnabled)
        reason = 'FAMILY_PILOT'
      else if (
        state.config.mode === 'SINGLE_STORE_PILOT' &&
        recipient.role === 'STORE_MANAGER' &&
        complaint.storeId !== state.config.pilotStoreId
      )
        reason = 'SINGLE_STORE_PILOT'
      else if (!recipient.smsEnabled || !recipient.active || !recipient.phone)
        reason = 'RECIPIENT_NOT_READY'
      else if (!provider(env).ready) reason = 'SIGNALWIRE_NOT_READY'
      if (reason) {
        n.status = 'SUPPRESSED'
        n.failureReason = reason
        continue
      }
      try {
        const result = await provider(env).send(n, recipient.phone)
        n.providerMessageId = result.providerMessageId
        n.status = result.status
        n.sentAt = new Date().toISOString()
        if (result.status === 'SENT') {
          const sentAt = n.sentAt
          const isManager = recipient.id === complaint.assignedManagerId
          complaint.events.push({
            id: `evt-${crypto.randomUUID()}`,
            complaintId: complaint.id,
            type: isManager ? 'MANAGER_NOTIFIED' : 'OWNER_NOTIFIED',
            actor: 'system',
            timestamp: sentAt,
            metadata: {
              notificationId: n.id,
              provider: 'SIGNALWIRE',
              providerMessageId: result.providerMessageId,
            },
          })
          if (isManager) complaint.managerNotifiedAt = sentAt
        }
      } catch (error) {
        n.status = 'FAILED'
        n.failedAt = new Date().toISOString()
        n.failureReason = (error as Error).message
      }
    }
  await persistState(env.DB, state)
}

// ---------------------------------------------------------------------------
// Phase C: family review queue (mail_review_items, OPEN). Reads are available
// to any authenticated user; resolve/assign are owner/admin-only in this
// phase (whether store managers may resolve their own store's items is an
// open policy question for Manav). List/detail SQL lives in
// worker/review-queue.ts so a regression test runs it against the real schema.
// ---------------------------------------------------------------------------

app.get('/api/review-queue', async (c) => {
  const rows = await c.env.DB.prepare(REVIEW_QUEUE_LIST_QUERY).all<Record<string, unknown>>()
  return c.json({ items: rows.results })
})

app.get('/api/review-queue/:id', async (c) => {
  const row = await c.env.DB.prepare(REVIEW_QUEUE_DETAIL_QUERY)
    .bind(c.req.param('id'))
    .first<Record<string, unknown>>()
  if (!row) return c.json(jsonError('Review item not found'), 404)
  return c.json({ item: row })
})

app.post('/api/review-queue/:id/resolve', zValidator('json', reviewResolveSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const body = c.req.valid('json')
  // Identity rule: a review resolution may LINK to an existing complaint or
  // DISMISS the item. It can never create a complaint or an identity here.
  if (body.action === 'link') {
    if (!body.complaintId) return c.json(jsonError('complaintId is required to link'), 400)
    const target = await c.env.DB.prepare('SELECT id FROM complaints WHERE id=?')
      .bind(body.complaintId)
      .first<{ id: string }>()
    if (!target) return c.json(jsonError('Complaint not found'), 404)
  }
  const now = new Date().toISOString()
  const resolution = {
    action: body.action,
    complaintId: body.complaintId ?? null,
    note: body.note ?? null,
    resolvedBy: user.id,
    resolvedAt: now,
  }
  // When linking, also update the item's complaint_id column (not just the
  // resolution JSON) so downstream joins stay correct.
  const updated = await c.env.DB.prepare(
    `UPDATE mail_review_items SET status='RESOLVED',resolution_json=?,reviewed_by=?,resolved_at=?,
     complaint_id=CASE WHEN ? IS NOT NULL THEN ? ELSE complaint_id END,updated_at=?
     WHERE id=? AND status='OPEN'`,
  )
    .bind(
      JSON.stringify(resolution),
      user.id,
      now,
      body.complaintId ?? null,
      body.complaintId ?? null,
      now,
      c.req.param('id'),
    )
    .run()
  if (Number(updated.meta?.changes ?? 0) !== 1)
    return c.json(jsonError('Review item not found or already resolved'), 404)
  return c.json({ ok: true, resolution })
})

app.post('/api/review-queue/:id/assign', zValidator('json', reviewAssignSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const body = c.req.valid('json')
  const assignee = await c.env.DB.prepare('SELECT id FROM users WHERE id=? AND active=1')
    .bind(body.assigneeUserId)
    .first<{ id: string }>()
  if (!assignee) return c.json(jsonError('Assignee not found'), 404)
  const now = new Date().toISOString()
  const updated = await c.env.DB.prepare(
    `UPDATE mail_review_items SET reviewed_by=?,updated_at=? WHERE id=? AND status='OPEN'`,
  )
    .bind(body.assigneeUserId, now, c.req.param('id'))
    .run()
  if (Number(updated.meta?.changes ?? 0) !== 1)
    return c.json(jsonError('Review item not found or not open'), 404)
  return c.json({ ok: true, assigneeUserId: body.assigneeUserId })
})

// ---------------------------------------------------------------------------
// Phase C: response templates + approval-gated ack/response workflow.
// Draft -> human approve -> send. Sending is disabled in this phase.
// ---------------------------------------------------------------------------
app.get('/api/response-templates', async (c) => {
  return c.json({ templates: await listResponseTemplates(c.env.DB) })
})

app.post('/api/response-templates', zValidator('json', responseTemplateSchema), async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const body = c.req.valid('json')
  const template = await createResponseTemplate(c.env.DB, body)
  return c.json({ ok: true, template })
})

app.post(
  '/api/complaints/:id/response-drafts',
  zValidator('json', responseDraftSchema),
  async (c) => {
    const user = c.get('user')
    if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
    try {
      const draft = await createResponseDraft(c.env.DB, {
        complaintId: c.req.param('id'),
        ...c.req.valid('json'),
      })
      return c.json({ ok: true, draft })
    } catch (error) {
      if (error instanceof ResponseWorkflowError)
        return c.json(
          jsonError(error.code),
          error.code === 'COMPLAINT_NOT_FOUND' ? 404 : 400,
        )
      throw error
    }
  },
)

app.post(
  '/api/response-actions/:id/approve',
  zValidator('json', responseApprovalSchema),
  async (c) => {
    const user = c.get('user')
    if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
    try {
      // Approval is attributed to the authenticated reviewer, never a body field.
      const action = await approveResponseAction(c.env.DB, {
        id: c.req.param('id'),
        approvedBy: user.id,
      })
      return c.json({ ok: true, action })
    } catch (error) {
      if (error instanceof ResponseWorkflowError)
        return c.json(
          jsonError(error.code),
          error.code === 'ACTION_NOT_FOUND' ? 404 : 400,
        )
      throw error
    }
  },
)

app.post('/api/response-actions/:id/submit', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  try {
    const action = await submitResponseForApproval(c.env.DB, c.req.param('id'))
    return c.json({ ok: true, action })
  } catch (error) {
    if (error instanceof ResponseWorkflowError)
      return c.json(
        jsonError(error.code),
        error.code === 'ACTION_NOT_FOUND' ? 404 : 400,
      )
    throw error
  }
})

app.post('/api/response-actions/:id/reject', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  try {
    const action = await rejectResponseAction(c.env.DB, {
      id: c.req.param('id'),
      rejectedBy: user.id,
    })
    return c.json({ ok: true, action })
  } catch (error) {
    if (error instanceof ResponseWorkflowError)
      return c.json(
        jsonError(error.code),
        error.code === 'ACTION_NOT_FOUND' ? 404 : 400,
      )
    throw error
  }
})

app.post('/api/response-actions/:id/send', async (c) => {  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  try {
    await sendResponseAction(c.env.DB, {
      id: c.req.param('id'),
      emailAckEnabled: state.config.emailAckEnabled === true,
    })
    return c.json({ ok: true })
  } catch (error) {
    if (error instanceof ResponseWorkflowError) {
      const status = error.code === 'SEND_DISABLED' ? 403 : error.code === 'ACTION_NOT_FOUND' ? 404 : 409
      return c.json(jsonError(error.code), status as 403 | 404 | 409)
    }
    throw error
  }
})

// ---------------------------------------------------------------------------
// Phase C: manual delta-discovery trigger. Cron stays disabled; this is an
// explicit admin action, and discovery itself refuses to run unless
// email_ingestion_enabled=true.
// ---------------------------------------------------------------------------
app.post('/api/admin/email/discover', async (c) => {
  const user = c.get('user')
  if (!canAdmin(user)) return c.json(jsonError('Owner access required'), 403)
  const state = await loadState(c.env.DB)
  if (!state.config.emailIngestionEnabled)
    return c.json({ ok: false, ran: false, reason: 'INGESTION_DISABLED' })
  const graph = emailProvider(c.env)
  if (!graph.ready) return c.json(jsonError('MS_GRAPH_NOT_CONFIGURED'), 503)
  const result = await runMailboxDiscovery(c.env.DB, {
    provider: 'MICROSOFT_GRAPH',
    mailboxKey: 'production',
    ingestionEnabled: true,
    client: new GraphDeltaDiscoveryClient(graph),
  })
  return c.json({ ok: result.ran, ...result })
})

app.all('/api/*', (c) => c.json(jsonError('API route not found'), 404))
app.get('*', (c) =>
  c.env.ASSETS ? c.env.ASSETS.fetch(c.req.raw) : c.text('StoreResolve worker is healthy'),
)
export default {
  fetch: app.fetch,
  async scheduled(
    _controller: { cron: string; scheduledTime: number },
    env: Bindings,
  ): Promise<void> {
    await runScheduledOperations(env, emailProvider(env), provider(env), dispatchEligible)
  },
}
