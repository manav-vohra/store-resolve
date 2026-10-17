import {
  EmailProviderError,
  type EmailProvider,
  type NormalizedEmailMessage,
} from './email-provider'

export interface MicrosoftGraphConfig {
  clientId?: string
  clientSecret?: string
  refreshToken?: string
  mailboxAddress?: string
  tenant?: string
}

interface GraphAddress {
  emailAddress?: { name?: string; address?: string }
}

export interface GraphMessage {
  id?: string
  conversationId?: string
  parentFolderId?: string
  internetMessageId?: string
  receivedDateTime?: string
  subject?: string
  from?: GraphAddress
  sender?: GraphAddress
  toRecipients?: GraphAddress[]
  ccRecipients?: GraphAddress[]
  body?: { contentType?: string; content?: string }
  bodyPreview?: string
  internetMessageHeaders?: Array<{ name?: string; value?: string }>
}

export interface PilotMessageSelection {
  id: string
  conversationId: string
  parentFolderId: string
  receivedDateTime: string
  subject: string
  senderAddress: string
  senderMatched: boolean
  receivedWindowMatched: boolean
  caseIdMatched: boolean
  subjectPhraseMatched: boolean
  storeTokenMatched: boolean
  previousSubjectPrefixMatched: boolean
  previousWindowMatched: boolean
}

export interface MailboxIdentityDiagnosticMessage {
  id: string
  receivedDateTime: string
  senderAddress: string
}

export interface MailboxIdentityDiagnosticResult {
  records: MailboxIdentityDiagnosticMessage[]
  hasMore: boolean
}

export interface PilotUniquenessMetadata {
  id: string
  conversationId: string
  parentFolderId: string
  receivedDateTime: string
  subject: string
  senderAddress: string
  senderMatched: boolean
  caseIdMatched: boolean
  subjectPhraseMatched: boolean
  storeNumberMatched: boolean
}

export interface PilotUniquenessMetadataResult {
  records: PilotUniquenessMetadata[]
  hasMore: boolean
}

export const pilotBodyDiagnosticCandidates = [
  { maskedId: 'AAkA…DQAA', receivedDateTime: '2026-08-01T18:26:14Z' },
  { maskedId: 'AAkA…EQAA', receivedDateTime: '2026-08-01T18:27:50Z' },
] as const

export const pilotMessageSelector = {
  senderAddress: 'customerservice@dunkinbrands.com',
  caseId: 'CCC11122413',
  subjectPhrase: 'Slow Service',
  storeToken: '350909-DD',
  receivedStart: '2026-08-01T18:20:00.000Z',
  receivedEnd: '2026-08-01T18:35:00.000Z',
  previousSubjectPrefix: 'DBI Case # (CCC11122413) - Guest Contact: Slow Service',
  previousReceivedStart: '2026-08-01T18:25:00.000Z',
  previousReceivedEnd: '2026-08-01T18:30:00.000Z',
} as const

export const mailboxIdentityDiagnosticWindows = {
  dateOnly: {
    receivedStart: '2026-08-01T00:00:00.000Z',
    receivedEnd: '2026-08-02T00:00:00.000Z',
  },
  senderOnly: {
    senderAddress: 'customerservice@dunkinbrands.com',
    receivedStart: '2026-07-01T00:00:00.000Z',
    receivedEnd: '2026-08-25T00:00:00.000Z',
  },
} as const

export const pilotUniquenessDiagnostic = {
  receivedStart: '2026-08-01T00:00:00.000Z',
  receivedEnd: '2026-08-02T00:00:00.000Z',
  maxRecords: 10,
} as const

export const isApprovedPilotMessageMetadata = (
  message: Pick<
    PilotMessageSelection,
    | 'senderMatched'
    | 'receivedWindowMatched'
    | 'caseIdMatched'
    | 'subjectPhraseMatched'
    | 'storeTokenMatched'
  >,
) =>
  message.senderMatched &&
  message.receivedWindowMatched &&
  message.caseIdMatched &&
  message.subjectPhraseMatched &&
  message.storeTokenMatched

export const maskGraphIdentifier = (value: string) =>
  value.length <= 8 ? '••••' : `${value.slice(0, 4)}…${value.slice(-4)}`

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
const GRAPH_SCOPES = [
  'offline_access',
  'https://graph.microsoft.com/Mail.Read',
  'https://graph.microsoft.com/Mail.Send',
].join(' ')

const htmlToText = (value: string) =>
  value
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim()

const formatAddress = (value?: GraphAddress) => {
  const address = value?.emailAddress?.address?.replace(/[\r\n]+/g, ' ').trim() ?? ''
  const name = value?.emailAddress?.name?.replace(/[\r\n]+/g, ' ').trim() ?? ''
  return name && address ? `${name} <${address}>` : address
}

const header = (message: GraphMessage, name: string) =>
  message.internetMessageHeaders?.find((item) => item.name?.toLowerCase() === name.toLowerCase())
    ?.value

export function normalizeGraphMessage(message: GraphMessage): NormalizedEmailMessage {
  if (!message.id || !message.conversationId || !message.receivedDateTime)
    throw new EmailProviderError(
      'MS_GRAPH_INVALID_MESSAGE',
      'Microsoft Graph message is incomplete',
    )
  const content = message.body?.content ?? message.bodyPreview ?? ''
  const isHtml = message.body?.contentType?.toLowerCase() === 'html'
  return {
    id: message.id,
    threadId: message.conversationId,
    internalDate: new Date(message.receivedDateTime).toISOString(),
    sender: formatAddress(message.from ?? message.sender),
    recipients: [...(message.toRecipients ?? []), ...(message.ccRecipients ?? [])]
      .map(formatAddress)
      .filter(Boolean)
      .join(', '),
    subject: message.subject?.replace(/[\r\n]+/g, ' ').trim() || '(no subject)',
    messageIdHeader: message.internetMessageId || header(message, 'Message-ID'),
    inReplyTo: header(message, 'In-Reply-To'),
    references: header(message, 'References'),
    textBody: (isHtml ? htmlToText(content) : content).slice(0, 100_000),
    htmlBody: isHtml ? content.slice(0, 200_000) : undefined,
  }
}

export interface InboxDeltaPage {
  messages: GraphMessage[]
  /** Present while more pages remain; pass back to continue paging. */
  nextLink?: string
  /** Present on the final page; persist as the discovery cursor. */
  deltaLink?: string
}

export class MicrosoftGraphProvider implements EmailProvider {
  constructor(private readonly config: MicrosoftGraphConfig) {}

  get ready() {
    return Boolean(this.config.clientId && this.config.refreshToken && this.config.mailboxAddress)
  }

  private async accessToken(): Promise<string> {
    if (!this.ready)
      throw new EmailProviderError('MS_GRAPH_NOT_CONFIGURED', 'Microsoft Graph is not configured')
    const tenant = this.config.tenant?.trim() || 'consumers'
    const response = await fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: this.config.clientId!,
          refresh_token: this.config.refreshToken!,
          grant_type: 'refresh_token',
          scope: GRAPH_SCOPES,
          ...(this.config.clientSecret ? { client_secret: this.config.clientSecret } : {}),
        }),
      },
    )
    const payload: unknown = await response.json().catch(() => null)
    const accessToken =
      payload && typeof payload === 'object' && 'access_token' in payload
        ? (payload as { access_token?: unknown }).access_token
        : undefined
    if (!response.ok || typeof accessToken !== 'string')
      throw new EmailProviderError(
        response.status === 401 ? 'MS_GRAPH_OAUTH_401' : 'MS_GRAPH_OAUTH_FAILED',
        'Microsoft Graph OAuth refresh failed',
        response.status,
      )
    return accessToken
  }

  private async request<T>(path: string, init: RequestInit = {}, retryReads = true): Promise<T> {
    const method = init.method ?? 'GET'
    const attempts = retryReads && method === 'GET' ? 3 : 1
    for (let attempt = 1; attempt <= attempts; attempt += 1) {
      const token = await this.accessToken()
      // Delta links returned by Graph are absolute URLs; request paths are relative.
      const url = /^https?:\/\//i.test(path) ? path : `${GRAPH_ROOT}${path}`
      const response = await fetch(url, {
        ...init,
        headers: {
          authorization: `Bearer ${token}`,
          accept: 'application/json',
          Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"',
          ...init.headers,
        },
      })
      if (response.ok) {
        if (response.status === 202 || response.status === 204) return undefined as T
        return (await response.json()) as T
      }
      const retryable = response.status === 429 || response.status >= 500
      if (retryable && attempt < attempts) {
        await new Promise((resolve) => setTimeout(resolve, 250 * 2 ** (attempt - 1)))
        continue
      }
      const code =
        response.status === 401
          ? 'MS_GRAPH_API_401'
          : response.status === 403
            ? 'MS_GRAPH_API_403_SCOPE'
            : response.status === 429
              ? 'MS_GRAPH_API_RATE_LIMIT'
              : response.status >= 500
                ? 'MS_GRAPH_API_TEMPORARY'
                : 'MS_GRAPH_API_ERROR'
      throw new EmailProviderError(
        code,
        `Microsoft Graph request failed (${response.status})`,
        response.status,
      )
    }
    throw new EmailProviderError('MS_GRAPH_RETRY_EXHAUSTED', 'Microsoft Graph retry limit reached')
  }

  async verifyConnection(): Promise<void> {
    const folder = await this.request<{
      id?: string
      displayName?: string
      totalItemCount?: number
      unreadItemCount?: number
    }>('/me/mailFolders/inbox?$select=id,displayName,totalItemCount,unreadItemCount')
    if (!folder.id)
      throw new EmailProviderError(
        'MS_GRAPH_INVALID_MAILBOX_METADATA',
        'Microsoft Graph Inbox metadata is incomplete',
      )
  }

  async findPilotComplaintCandidates(): Promise<{
    candidates: PilotMessageSelection[]
    inspectedCandidates: PilotMessageSelection[]
    hasMore: boolean
  }> {
    const params = new URLSearchParams({
      $filter: [
        `receivedDateTime ge ${pilotMessageSelector.receivedStart}`,
        `receivedDateTime le ${pilotMessageSelector.receivedEnd}`,
        `from/emailAddress/address eq '${pilotMessageSelector.senderAddress}'`,
      ].join(' and '),
      $select: 'id,conversationId,parentFolderId,subject,receivedDateTime,from',
      $top: '3',
    })
    const page = await this.request<{
      value?: GraphMessage[]
      '@odata.nextLink'?: string
    }>(`/me/messages?${params.toString()}`)
    const inspectedCandidates = (page.value ?? []).map((candidate) => {
      const senderAddress = candidate.from?.emailAddress?.address?.trim() ?? ''
      const receivedAt = Date.parse(candidate.receivedDateTime ?? '')
      if (
        !candidate.id ||
        !candidate.conversationId ||
        !candidate.parentFolderId ||
        !candidate.receivedDateTime ||
        !candidate.subject ||
        !Number.isFinite(receivedAt)
      )
        throw new EmailProviderError(
          'MS_GRAPH_PILOT_INVALID_SELECTION',
          'Microsoft Graph pilot metadata was incomplete',
        )
      const normalizedSubject = candidate.subject.toLowerCase()
      return {
        id: candidate.id,
        conversationId: candidate.conversationId,
        parentFolderId: candidate.parentFolderId,
        receivedDateTime: candidate.receivedDateTime,
        subject: candidate.subject,
        senderAddress,
        senderMatched: senderAddress.toLowerCase() === pilotMessageSelector.senderAddress,
        receivedWindowMatched:
          receivedAt >= Date.parse(pilotMessageSelector.receivedStart) &&
          receivedAt <= Date.parse(pilotMessageSelector.receivedEnd),
        caseIdMatched: normalizedSubject.includes(pilotMessageSelector.caseId.toLowerCase()),
        subjectPhraseMatched: normalizedSubject.includes(
          pilotMessageSelector.subjectPhrase.toLowerCase(),
        ),
        storeTokenMatched: normalizedSubject.includes(
          pilotMessageSelector.storeToken.toLowerCase(),
        ),
        previousSubjectPrefixMatched: candidate.subject.startsWith(
          pilotMessageSelector.previousSubjectPrefix,
        ),
        previousWindowMatched:
          receivedAt >= Date.parse(pilotMessageSelector.previousReceivedStart) &&
          receivedAt <= Date.parse(pilotMessageSelector.previousReceivedEnd),
      }
    })
    const candidates = inspectedCandidates.filter(isApprovedPilotMessageMetadata)
    return { candidates, inspectedCandidates, hasMore: Boolean(page['@odata.nextLink']) }
  }

  async getMailFolderMetadata(id: string): Promise<{ id: string; displayName: string }> {
    const folder = await this.request<{ id?: string; displayName?: string }>(
      `/me/mailFolders/${encodeURIComponent(id)}?$select=id,displayName`,
    )
    if (!folder.id || !folder.displayName)
      throw new EmailProviderError(
        'MS_GRAPH_INVALID_FOLDER_METADATA',
        'Microsoft Graph folder metadata was incomplete',
      )
    return { id: folder.id, displayName: folder.displayName }
  }

  async listPilotUniquenessMetadata(): Promise<PilotUniquenessMetadataResult> {
    const params = new URLSearchParams({
      $filter: [
        `receivedDateTime ge ${pilotUniquenessDiagnostic.receivedStart}`,
        `receivedDateTime lt ${pilotUniquenessDiagnostic.receivedEnd}`,
        `from/emailAddress/address eq '${pilotMessageSelector.senderAddress}'`,
      ].join(' and '),
      $select: 'id,conversationId,parentFolderId,subject,receivedDateTime,from',
      $top: String(pilotUniquenessDiagnostic.maxRecords),
    })
    const page = await this.request<{
      value?: GraphMessage[]
      '@odata.nextLink'?: string
    }>(`/me/messages?${params.toString()}`)
    const records = (page.value ?? []).map((message) => {
      const senderAddress = message.from?.emailAddress?.address?.trim() ?? ''
      if (
        !message.id ||
        !message.conversationId ||
        !message.parentFolderId ||
        !message.subject ||
        !message.receivedDateTime ||
        !senderAddress
      )
        throw new EmailProviderError(
          'MS_GRAPH_INVALID_UNIQUENESS_METADATA',
          'Microsoft Graph uniqueness diagnostic metadata was incomplete',
        )
      const normalizedSubject = message.subject.toLowerCase()
      return {
        id: message.id,
        conversationId: message.conversationId,
        parentFolderId: message.parentFolderId,
        receivedDateTime: message.receivedDateTime,
        subject: message.subject,
        senderAddress,
        senderMatched:
          senderAddress.toLowerCase() === pilotMessageSelector.senderAddress.toLowerCase(),
        caseIdMatched: normalizedSubject.includes(pilotMessageSelector.caseId.toLowerCase()),
        subjectPhraseMatched: normalizedSubject.includes(
          pilotMessageSelector.subjectPhrase.toLowerCase(),
        ),
        storeNumberMatched: /350\D*909/i.test(message.subject),
      }
    })
    return { records, hasMore: Boolean(page['@odata.nextLink']) }
  }

  async findPilotBodyDiagnosticCandidates(): Promise<PilotUniquenessMetadata[]> {
    const params = new URLSearchParams({
      $filter: [
        `receivedDateTime ge ${pilotUniquenessDiagnostic.receivedStart}`,
        `receivedDateTime lt ${pilotUniquenessDiagnostic.receivedEnd}`,
        `from/emailAddress/address eq '${pilotMessageSelector.senderAddress}'`,
      ].join(' and '),
      $select: 'id,conversationId,parentFolderId,subject,receivedDateTime,from',
      $top: '3',
    })
    const page = await this.request<{ value?: GraphMessage[] }>(
      `/me/messages?${params.toString()}`,
      {},
      false,
    )
    return (page.value ?? []).flatMap((message) => {
      const senderAddress = message.from?.emailAddress?.address?.trim() ?? ''
      if (
        !message.id ||
        !message.conversationId ||
        !message.parentFolderId ||
        !message.subject ||
        !message.receivedDateTime ||
        !senderAddress
      )
        throw new EmailProviderError(
          'MS_GRAPH_INVALID_BODY_DIAGNOSTIC_METADATA',
          'Microsoft Graph body diagnostic metadata was incomplete',
        )
      const normalizedSubject = message.subject.toLowerCase()
      const record = {
        id: message.id,
        conversationId: message.conversationId,
        parentFolderId: message.parentFolderId,
        receivedDateTime: message.receivedDateTime,
        subject: message.subject,
        senderAddress,
        senderMatched:
          senderAddress.toLowerCase() === pilotMessageSelector.senderAddress.toLowerCase(),
        caseIdMatched: normalizedSubject.includes(pilotMessageSelector.caseId.toLowerCase()),
        subjectPhraseMatched: normalizedSubject.includes(
          pilotMessageSelector.subjectPhrase.toLowerCase(),
        ),
        storeNumberMatched: /350\D*909/i.test(message.subject),
      }
      return record.senderMatched &&
        record.caseIdMatched &&
        record.subjectPhraseMatched &&
        record.storeNumberMatched
        ? [record]
        : []
    })
  }

  async getPilotBodyDiagnosticMessage(id: string): Promise<NormalizedEmailMessage> {
    const select = [
      'id',
      'conversationId',
      'internetMessageId',
      'receivedDateTime',
      'subject',
      'from',
      'body',
    ].join(',')
    const message = await this.request<GraphMessage>(
      `/me/messages/${encodeURIComponent(id)}?$select=${encodeURIComponent(select)}`,
      {},
      false,
    )
    return normalizeGraphMessage(message)
  }

  private async listMailboxIdentityMetadata(
    receivedStart: string,
    receivedEnd: string,
    senderAddress?: string,
  ): Promise<MailboxIdentityDiagnosticResult> {
    const filter = [
      `receivedDateTime ge ${receivedStart}`,
      `receivedDateTime lt ${receivedEnd}`,
      ...(senderAddress
        ? [`from/emailAddress/address eq '${senderAddress.replace(/'/g, "''")}'`]
        : []),
    ].join(' and ')
    const params = new URLSearchParams({
      $filter: filter,
      $select: 'id,receivedDateTime,from',
      $top: '3',
    })
    const page = await this.request<{
      value?: GraphMessage[]
      '@odata.nextLink'?: string
    }>(`/me/messages?${params.toString()}`)
    const records = (page.value ?? []).map((message) => {
      const sender = message.from?.emailAddress?.address?.trim()
      if (!message.id || !message.receivedDateTime || !sender)
        throw new EmailProviderError(
          'MS_GRAPH_INVALID_IDENTITY_METADATA',
          'Microsoft Graph identity diagnostic metadata was incomplete',
        )
      return {
        id: message.id,
        receivedDateTime: message.receivedDateTime,
        senderAddress: sender,
      }
    })
    return { records, hasMore: Boolean(page['@odata.nextLink']) }
  }

  async runDateOnlyIdentityDiagnostic(): Promise<MailboxIdentityDiagnosticResult> {
    return this.listMailboxIdentityMetadata(
      mailboxIdentityDiagnosticWindows.dateOnly.receivedStart,
      mailboxIdentityDiagnosticWindows.dateOnly.receivedEnd,
    )
  }

  async runSenderOnlyIdentityDiagnostic(): Promise<MailboxIdentityDiagnosticResult> {
    return this.listMailboxIdentityMetadata(
      mailboxIdentityDiagnosticWindows.senderOnly.receivedStart,
      mailboxIdentityDiagnosticWindows.senderOnly.receivedEnd,
      mailboxIdentityDiagnosticWindows.senderOnly.senderAddress,
    )
  }

  /**
   * Incremental inbox discovery via Graph delta query. Pass the persisted
   * deltaLink (or nextLink) to continue; the final page returns a fresh
   * deltaLink to persist as the cursor. No live calls happen unless the
   * caller invokes this — tests use a fake DeltaDiscoveryClient instead.
   *
   * Phase 1 fix: initial delta requests are metadata-only (no body,
   * bodyPreview, or headers) and accept an optional receivedDateTime window.
   * The window is sent as Graph $filter AND enforced client-side by the
   * caller, because Graph may ignore $filter on delta. Full bodies are
   * fetched per-message only when that message needs processing, via
   * getMessage(). This prevents the initial delta from pulling bodies for
   * the whole historical inbox.
   */
  async listInboxDeltaPage(
    link?: string,
    options?: { receivedStart?: string; receivedEnd?: string },
  ): Promise<InboxDeltaPage> {
    const select = [
      'id',
      'conversationId',
      'internetMessageId',
      'receivedDateTime',
      'subject',
      'from',
      'sender',
      'toRecipients',
      'ccRecipients',
      'parentFolderId',
    ].join(',')
    let path: string
    if (link) {
      path = link
    } else {
      const params = new URLSearchParams({
        $select: select,
        $top: '25',
      })
      const filters: string[] = []
      if (options?.receivedStart) filters.push(`receivedDateTime ge ${options.receivedStart}`)
      if (options?.receivedEnd) filters.push(`receivedDateTime lt ${options.receivedEnd}`)
      if (filters.length > 0) params.set('$filter', filters.join(' and '))
      path = `/me/mailFolders/inbox/messages/delta?${params.toString()}`
    }
    const page = await this.request<{
      value?: GraphMessage[]
      '@odata.nextLink'?: string
      '@odata.deltaLink'?: string
    }>(path)
    return {
      messages: page.value ?? [],
      nextLink: page['@odata.nextLink'],
      deltaLink: page['@odata.deltaLink'],
    }
  }

  /**
   * Preview-only inbox listing for the Check Mail Now flow. Uses
   * /me/messages (NOT delta), so it never advances the ingestion cursor.
   * Returns metadata + bodyPreview (for classification) but never full
   * bodies or headers. Read-only: no D1 writes, no sends, no cursor change.
   */
  async listInboxPreviewPage(options: {
    receivedStart: string
    receivedEnd: string
    top?: number
    nextLink?: string
  }): Promise<{ messages: GraphMessage[]; nextLink?: string }> {
    if (options.nextLink) {
      const page = await this.request<{
        value?: GraphMessage[]
        '@odata.nextLink'?: string
      }>(options.nextLink)
      return { messages: page.value ?? [], nextLink: page['@odata.nextLink'] }
    }
    const select = [
      'id',
      'conversationId',
      'internetMessageId',
      'receivedDateTime',
      'subject',
      'from',
      'sender',
      'toRecipients',
      'ccRecipients',
      'parentFolderId',
      'bodyPreview',
    ].join(',')
    const params = new URLSearchParams({
      $select: select,
      $filter: `receivedDateTime ge ${options.receivedStart} and receivedDateTime lt ${options.receivedEnd}`,
      $orderby: 'receivedDateTime desc',
      $top: String(Math.max(1, Math.min(100, options.top ?? 50))),
    })
    const page = await this.request<{
      value?: GraphMessage[]
      '@odata.nextLink'?: string
    }>(`/me/messages?${params.toString()}`)
    return { messages: page.value ?? [], nextLink: page['@odata.nextLink'] }
  }

  async listMessageIds(lookbackDays: number, maxResults = 25): Promise<string[]> {    const since = new Date(
      Date.now() - Math.max(1, Math.min(365, lookbackDays)) * 86_400_000,
    ).toISOString()
    const params = new URLSearchParams({
      $select: 'id',
      $filter: `receivedDateTime ge ${since}`,
      $orderby: 'receivedDateTime asc',
      $top: String(Math.min(maxResults, 100)),
    })
    const page = await this.request<{ value?: Array<{ id?: string }> }>(
      `/me/mailFolders/inbox/messages?${params.toString()}`,
    )
    return (page.value ?? []).flatMap((item) => (item.id ? [item.id] : []))
  }

  async getMessage(id: string): Promise<NormalizedEmailMessage> {
    const select = [
      'id',
      'conversationId',
      'internetMessageId',
      'receivedDateTime',
      'subject',
      'from',
      'sender',
      'toRecipients',
      'ccRecipients',
      'body',
      'bodyPreview',
      'internetMessageHeaders',
    ].join(',')
    const message = await this.request<GraphMessage>(
      `/me/messages/${encodeURIComponent(id)}?$select=${encodeURIComponent(select)}`,
    )
    return normalizeGraphMessage(message)
  }

  async sendAcknowledgment(message: NormalizedEmailMessage, body: string): Promise<undefined> {
    await this.request<void>(
      `/me/messages/${encodeURIComponent(message.id)}/reply`,
      {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ comment: body }),
      },
      false,
    )
    // Graph's reply action returns 202 with no response body or sent-message identifier.
    return undefined
  }
}
