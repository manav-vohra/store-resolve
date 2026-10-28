import { z } from 'zod'
import type { AppConfig, AppState, BootstrapResponse, NewComplaint, Store, User } from './types'
import {
  actionSchema,
  complaintInputSchema,
  configSchema,
  contactSchema,
  storeAdminSchema,
  userAdminSchema,
} from './api-schema'

const errorSchema = z.object({ error: z.string() })
async function request<T>(path: string, schema: z.ZodType<T>, init?: RequestInit): Promise<T> {
  const response = await fetch(`/api${path}`, {
    credentials: 'same-origin',
    headers: { 'content-type': 'application/json', ...init?.headers },
    ...init,
  })
  const json: unknown = await response.json()
  if (!response.ok) throw new Error(errorSchema.safeParse(json).data?.error ?? 'Request failed')
  return schema.parse(json)
}
const bootstrapSchema = z.custom<BootstrapResponse>()
const stateSchema = z.custom<AppState>()
const reconciliationSchema = z.object({
  ok: z.literal(true),
  result: z.enum(['UPDATED', 'UNCHANGED']),
  providerStatus: z.enum(['DELIVERED', 'FAILED', 'UNDELIVERED']),
})
const pilotEmailSchema = z.object({
  ok: z.literal(true),
  accessed: z.boolean(),
  status: z.enum(['PROCESSED', 'ROUTING_REVIEW', 'DUPLICATE', 'FOLLOW_UP']).optional(),
  complaintId: z.string().optional(),
  messageId: z.string().optional(),
  conversationId: z.string().optional(),
  matchCount: z.union([z.literal(0), z.literal(1), z.literal('2+')]).optional(),
  inspectedCount: z.number().int().min(0).max(3).optional(),
})
export interface ReviewQueueItem {
  id: string
  status: string
  reason_code: string
  disagreement_flags_json: string | null
  created_at: string
  subject: string
  sender_address: string
  received_at: string
  provider_message_id: string
  interpretation_confidence: number | null
  interpretation_disagreements_json: string | null
  interpretation_model: string | null
  interpretation_prompt_version: string | null
}
export interface ReviewQueueDetail extends Record<string, unknown> {
  id: string
  reason_code: string
  subject: string
  sender_address: string
}
const reviewQueueSchema = z.object({ items: z.array(z.custom<ReviewQueueItem>()) })
const reviewDetailSchema = z.object({ item: z.custom<ReviewQueueDetail>() })
const reviewResolveSchema = z.object({
  ok: z.literal(true),
  resolution: z.custom<Record<string, unknown>>(),
})
const reviewAssignSchema = z.object({ ok: z.literal(true), assigneeUserId: z.string() })
export const api = {
  reviewQueue: () => request('/review-queue', reviewQueueSchema),
  reviewItem: (id: string) => request(`/review-queue/${encodeURIComponent(id)}`, reviewDetailSchema),
  resolveReviewItem: (id: string, body: { action: 'dismiss' | 'link'; complaintId?: string; note?: string }) =>
    request(`/review-queue/${encodeURIComponent(id)}/resolve`, reviewResolveSchema, {
      method: 'POST',
      body: JSON.stringify(body),
    }),
  assignReviewItem: (id: string, assigneeUserId: string) =>
    request(`/review-queue/${encodeURIComponent(id)}/assign`, reviewAssignSchema, {
      method: 'POST',
      body: JSON.stringify({ assigneeUserId }),
    }),
  bootstrap: () => request('/bootstrap', bootstrapSchema),
  stateVersion: () => request('/state-version', z.object({ version: z.number() })),
  createComplaint: (input: NewComplaint) =>
    request('/complaints', stateSchema, {
      method: 'POST',
      body: JSON.stringify(complaintInputSchema.parse(input)),
    }),
  act: (
    id: string,
    action: z.infer<typeof actionSchema>['action'],
    data: Record<string, unknown> = {},
  ) =>
    request(`/complaints/${encodeURIComponent(id)}/actions`, stateSchema, {
      method: 'POST',
      body: JSON.stringify(actionSchema.parse({ action, data })),
    }),
  processDeadlines: (advanceHours = 0) =>
    request('/admin/deadlines', stateSchema, {
      method: 'POST',
      body: JSON.stringify({ advanceHours }),
    }),
  updateConfig: (config: AppConfig) =>
    request('/admin/config', stateSchema, {
      method: 'PUT',
      body: JSON.stringify(configSchema.parse(config)),
    }),
  updateContact: (
    id: string,
    contact: Pick<User, 'name' | 'email' | 'phone' | 'smsEnabled' | 'active'>,
  ) =>
    request('/admin/contacts/' + id, stateSchema, {
      method: 'PUT',
      body: JSON.stringify(contactSchema.parse(contact)),
    }),
  sendTest: (recipientUserId: 'father' | 'uncle' | 'grandfather' | 'pilot-admin' | 'dad') =>
    request('/admin/test-notifications', stateSchema, {
      method: 'POST',
      body: JSON.stringify({ recipientUserId, confirmed: true }),
    }),
  pushPublicKey: () =>
    request('/push/public-key', z.object({ publicKey: z.string().min(1) })),
  pushSubscribe: (input: { endpoint: string; keys: { p256dh: string; auth: string }; deviceLabel?: string }) =>
    request('/push/subscribe', z.object({ ok: z.literal(true) }), {
      method: 'POST',
      body: JSON.stringify(input),
    }),
  pushUnsubscribe: (endpoint: string) =>
    request('/push/unsubscribe', z.object({ ok: z.literal(true) }), {
      method: 'POST',
      body: JSON.stringify({ endpoint }),
    }),
  sendTestPush: (recipientUserId: 'father' | 'uncle' | 'grandfather' | 'pilot-admin' | 'dad') =>
    request('/admin/push/test', stateSchema, {
      method: 'POST',
      body: JSON.stringify({ recipientUserId, confirmed: true }),
    }),
  reconcileSignalWire: (providerMessageId: string) =>
    request('/admin/signalwire/reconcile', reconciliationSchema, {
      method: 'POST',
      body: JSON.stringify({ providerMessageId }),
    }),
  ingestPilotEmail: () =>
    request('/admin/email/pilot-ingest', pilotEmailSchema, {
      method: 'POST',
      body: JSON.stringify({}),
    }),
  previewMail: (params: { start: string; end: string; top?: number; pageToken?: string }) => {
    const query = new URLSearchParams({ start: params.start, end: params.end, top: String(params.top ?? 50) })
    if (params.pageToken) query.set('pageToken', params.pageToken)
    return request(
      `/admin/email/preview?${query.toString()}`,
      z.object({
        ok: z.literal(true),
        window: z.object({ receivedStart: z.string(), receivedEnd: z.string() }),
        count: z.number(),
        items: z.array(
          z.object({
            graphMessageId: z.string(),
            maskedId: z.string(),
            conversationId: z.string(),
            maskedConversationId: z.string(),
            receivedAt: z.string(),
            sender: z.string(),
            senderDisplay: z.string(),
            subject: z.string(),
            detectedCaseId: z.string().nullable(),
            detectedStoreNumber: z.string().nullable(),
            isForwardOrReply: z.boolean(),
            classification: z.enum([
              'COMPLAINT_CANDIDATE',
              'FOLLOW_UP_LIKELY',
              'NEEDS_REVIEW',
              'NOT_COMPLAINT',
              'ALREADY_INGESTED',
            ]),
            alreadyIngested: z.boolean(),
            classificationReason: z.string(),
            linkedComplaintId: z.string().nullable(),
          }),
        ),
        nextPageToken: z.string().nullable(),
      }),
    )
  },
  ackPreview: (complaintId: string) =>
    request(
      `/complaints/${encodeURIComponent(complaintId)}/acknowledge-preview`,
      z.object({
        ok: z.literal(true),
        complaintId: z.string(),
        externalCaseId: z.string().nullable(),
        complaintStatus: z.string(),
        acknowledgementStatus: z.string(),
        source: z.object({
          sender: z.string(),
          subject: z.string(),
          threadIdMasked: z.string(),
          internalDate: z.string(),
        }),
        template: z.string(),
        ack: z
          .object({
            status: z.string(),
            attempt_count: z.number(),
            last_error_code: z.string().nullable(),
            updated_at: z.string(),
          })
          .nullable(),
        mailbox: z.object({ matches: z.boolean(), reason: z.string().nullable() }).nullable(),
        canSend: z.boolean(),
      }),
    ),
  ackSend: (complaintId: string) =>
    request(
      `/complaints/${encodeURIComponent(complaintId)}/acknowledge`,
      z.object({ ok: z.boolean(), result: z.string() }),
      { method: 'POST', body: JSON.stringify({ confirmed: true }) },
    ),
  ackReclaim: (complaintId: string, body: { checkedSentItems: true; sentConfirmed: boolean; forceReset?: boolean; reconciliationNote?: string }) =>
    request(
      `/complaints/${encodeURIComponent(complaintId)}/acknowledge/reclaim`,
      z.object({ ok: z.boolean(), result: z.string() }),
      { method: 'POST', body: JSON.stringify(body) },
    ),
  listResponseActions: (complaintId: string) =>
    request(
      `/complaints/${encodeURIComponent(complaintId)}/response-actions`,
      z.object({
        ok: z.literal(true),
        actions: z.array(
          z.object({
            id: z.string(),
            complaintId: z.string(),
            responseKind: z.string(),
            policyState: z.string(),
            contentRef: z.string().optional(),
            approvedBy: z.string().optional(),
            approvedAt: z.string().optional(),
            createdAt: z.string(),
            updatedAt: z.string(),
          }),
        ),
      }),
    ),
  createResponseDraft: (complaintId: string, body: { body: string; responseKind?: string }) =>
    request(
      `/complaints/${encodeURIComponent(complaintId)}/response-drafts`,
      z.object({ ok: z.literal(true), draft: z.object({ id: z.string() }).passthrough() }),
      { method: 'POST', body: JSON.stringify(body) },
    ),
  updateResponseDraft: (id: string, body: string) =>
    request(
      `/response-actions/${encodeURIComponent(id)}`,
      z.object({ ok: z.literal(true), action: z.object({ id: z.string() }).passthrough() }),
      { method: 'PATCH', body: JSON.stringify({ body }) },
    ),
  submitResponseAction: (id: string) =>
    request(
      `/response-actions/${encodeURIComponent(id)}/submit`,
      z.object({ ok: z.literal(true) }),
      { method: 'POST' },
    ),
  approveResponseAction: (id: string) =>
    request(
      `/response-actions/${encodeURIComponent(id)}/approve`,
      z.object({ ok: z.literal(true) }),
      { method: 'POST', body: JSON.stringify({}) },
    ),
  rejectResponseAction: (id: string) =>
    request(
      `/response-actions/${encodeURIComponent(id)}/reject`,
      z.object({ ok: z.literal(true) }),
      { method: 'POST' },
    ),
  sendResponseAction: (id: string) =>
    request(
      `/response-actions/${encodeURIComponent(id)}/send`,
      z.object({ ok: z.literal(true) }),
      { method: 'POST' },
    ),
  notifyDraftReady: (id: string, force = false) =>
    request(
      `/admin/notify-draft-ready/${encodeURIComponent(id)}`,
      z.object({ ok: z.literal(true), sent: z.boolean(), reason: z.string().optional() }),
      { method: 'POST', body: JSON.stringify({ force }) },
    ),
  importPreviewMail: (body: {
    messageIds: string[]
    receivedStart: string
    receivedEnd: string
    ingestionMode?: 'LIVE' | 'BACKFILL'
  }) =>
    request(
      '/admin/email/preview-import',
      z.object({
        ok: z.literal(true),
        results: z.array(
          z.object({
            graphMessageId: z.string(),
            maskedId: z.string(),
            status: z.string(),
            complaintId: z.string().optional(),
            errorCode: z.string().optional(),
          }),
        ),
        imported: z.number(),
        failed: z.number(),
      }),
      { method: 'POST', body: JSON.stringify(body) },
    ),
  updateStore: (id: string, store: Store) =>
    request(`/admin/stores/${encodeURIComponent(id)}`, stateSchema, {
      method: 'PUT',
      body: JSON.stringify(
        storeAdminSchema.parse({
          number: store.number,
          name: store.name,
          address: store.address,
          city: store.city,
          state: store.state,
          postalCode: store.postalCode,
          phone: store.phone,
          active: store.active,
          managerId: store.managerId || undefined,
          aliases: store.aliases ?? [],
        }),
      ),
    }),
  updateUser: (id: string, changes: Partial<User> & { storeIds?: string[] }) =>
    request(`/admin/users/${encodeURIComponent(id)}`, stateSchema, {
      method: 'PUT',
      body: JSON.stringify(userAdminSchema.parse(changes)),
    }),
}
