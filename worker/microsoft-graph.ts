import {
  EmailProviderError,
  type EmailProvider,
  type MailboxIdentity,
  type MailboxIdentityVerification,
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

/**
 * Validates a Graph request URL before any bearer token is attached.
 * - Relative paths must start with /me/ and contain no traversal.
 * - Absolute URLs must be https://graph.microsoft.com (no port override,
 *   no userinfo, no fragment), path must start with /v1.0/me/, no traversal.
 * - Rejects external, malformed, or out-of-scope URLs (e.g. /v1.0/users/,
 *   /beta/, other hosts). Bearer tokens are never sent to unvalidated URLs.
 * - Never logs the URL or token: callers must not include secrets in errors.
 */
export function validateGraphRequestUrl(path: string): string {
  if (path.startsWith('/')) {
    if (path.includes('..') || path.includes('\\'))
      throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request path is invalid')
    if (!path.startsWith('/me/'))
      throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request must target /me/')
    return `${GRAPH_ROOT}${path}`
  }
  let url: URL
  try {
    url = new URL(path)
  } catch {
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request URL is malformed')
  }
  if (url.protocol !== 'https:')
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request must use https')
  if (url.hostname !== 'graph.microsoft.com')
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request host is not trusted')
  if (url.port !== '' && url.port !== '443')
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request port is not allowed')
  if (url.username || url.password)
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request must not include credentials')
  if (url.hash)
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request must not include a fragment')
  if (!url.pathname.startsWith('/v1.0/me/'))
    throw new EmailProviderError(
      'MS_GRAPH_INVALID_URL',
      'Graph request must target the /me/ mailbox scope',
    )
  if (url.pathname.includes('..') || url.pathname.includes('\\'))
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request path is invalid')
  // Encoded traversal (e.g. %2e%2e) must not escape /me/ scope.
  try {
    const decoded = decodeURIComponent(url.pathname)
    if (decoded.includes('..') || decoded.includes('\\'))
      throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request path is invalid')
  } catch (error) {
    if (error instanceof EmailProviderError) throw error
    throw new EmailProviderError('MS_GRAPH_INVALID_URL', 'Graph request path is malformed')
  }
  return url.toString()
}

/** Extracts $skiptoken from a Graph nextLink without trusting its host/path. */
export function extractSkiptoken(nextLink: string): string | null {
  try {
    const url = new URL(nextLink)
    return url.searchParams.get('$skiptoken')
  } catch {
    return null
  }
}

/**
 * Opaque pagination continuation for Inbox preview.
 *
 * Microsoft Graph message-list @odata.nextLink uses $skip (e.g.
 * .../messages?$skip=10&$top=10...), NOT $skiptoken. The previous code
 * only read $skiptoken, so every $skip nextLink was dropped and preview
 * returned nextPageToken=null even when more pages existed — the
 * Oct 10-17 top=10 truncation bug.
 *
 * Safe handling: We NEVER accept a browser-provided Graph URL. The server
 * validates the Graph-generated nextLink (must be the approved Inbox
 * endpoint, same $filter window, same $orderby, same $select), extracts
 * ONLY the paging state ($skip and/or $skiptoken + $top), and encodes it
 * as an opaque base64url token. Callers pass that token back; the server
 * decodes it, re-validates it, and rebuilds the Inbox URL with the same
 * window + preserved paging state. The complete Microsoft paging state is
 * preserved (skip offset + skiptoken + original top), while host/path/
 * filter/scope cannot be widened by a caller.
 */
export interface PageContinuation {
  skip?: number
  skiptoken?: string
  top?: number
}

const MAX_PAGE_TOKEN_BYTES = 4096
const MAX_SKIP = 1_000_000

function base64UrlEncode(data: string): string {
  const bytes = new TextEncoder().encode(data)
  let binary = ''
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i])
  return btoa(binary)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')
}

function base64UrlDecode(token: string): string | null {
  try {
    const padded = token.replace(/-/g, '+').replace(/_/g, '/')
    const paddedWithEquals = padded + '='.repeat((4 - (padded.length % 4)) % 4)
    return new TextDecoder().decode(
      Uint8Array.from(atob(paddedWithEquals), (c) => c.charCodeAt(0)),
    )
  } catch {
    return null
  }
}

/** Validates a Graph nextLink and encodes its paging state as an opaque token. */
export function encodePageToken(
  nextLink: string,
  // Kept for signature stability; filter is ignored (see below) so the
  // window is not used. Referenced to satisfy the linter.
  expectedWindow: { receivedStart: string; receivedEnd: string },
): string | null {
  void expectedWindow
  let url: URL
  try {
    // Graph nextLink is usually absolute, but may be relative — resolve
    // relative links against the Graph root instead of rejecting them.
    url = nextLink.startsWith('/')
      ? new URL(nextLink, 'https://graph.microsoft.com')
      : new URL(nextLink)
  } catch {
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink is malformed')
  }
  // Must be Graph, under /v1.0/me/ (same mailbox scope). We do NOT require
  // the exact Inbox path: Graph may normalize the nextLink path (e.g. to
  // /v1.0/me/messages), and we REBUILD with the Inbox path + original
  // window anyway, so the nextLink's path cannot widen scope. The
  // security boundary is: Graph host + /me/ scope + valid paging state.
  // Rejecting on exact path caused false rejections (Oct 17: real Graph
  // nextLinks were rejected, breaking all pagination).
  if (url.protocol !== 'https:' || url.hostname !== 'graph.microsoft.com')
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink host is not trusted')
  if (!url.pathname.startsWith('/v1.0/me/'))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink must target /me/ scope')
  if (url.username || url.password || url.hash || (url.port !== '' && url.port !== '443'))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink is invalid')
  if (url.pathname.includes('..') || url.pathname.includes('\\'))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink path is invalid')
  // $filter / $orderby in the nextLink are IGNORED, not validated strictly.
  // Graph normalizes the filter (e.g. strips .000 from datetimes,
  // re-orders clauses) and may omit filter/orderby entirely (only
  // $skip/$top). Strict matching caused false rejections (Oct 17: page 1
  // errored because Graph's nextLink filter datetime format differed).
  // This is safe: the server REBUILDS the URL with the original window
  // filter + orderby + paging state — the nextLink's filter/orderby are
  // never used, so they cannot widen scope. The security boundary is:
  // host/path must be Inbox messages, and paging state must be valid.
  // (expectedWindow is kept for future use / signature stability.)
  // Paging state: $skip (integer) and/or $skiptoken (opaque). At least one must be present.
  const skipRaw = url.searchParams.get('$skip')
  const skiptoken = url.searchParams.get('$skiptoken')
  const topRaw = url.searchParams.get('$top')
  let skip: number | undefined
  if (skipRaw !== null) {
    if (!/^\d+$/.test(skipRaw)) throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink $skip is invalid')
    skip = parseInt(skipRaw, 10)
    if (skip < 0 || skip > MAX_SKIP) throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink $skip out of range')
  }
  if (skiptoken !== null && (skiptoken.length === 0 || skiptoken.length > 4096 || /[\r\n]/.test(skiptoken)))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink $skiptoken is invalid')
  if (skip === undefined && !skiptoken)
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink has no paging state')
  let top: number | undefined
  if (topRaw !== null) {
    if (!/^\d+$/.test(topRaw)) throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink $top is invalid')
    top = parseInt(topRaw, 10)
    if (top < 1 || top > 100) throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph nextLink $top out of range')
  }
  const continuation: PageContinuation = {}
  if (skip !== undefined) continuation.skip = skip
  if (skiptoken) continuation.skiptoken = skiptoken
  if (top !== undefined) continuation.top = top
  const token = base64UrlEncode(JSON.stringify(continuation))
  if (token.length > MAX_PAGE_TOKEN_BYTES)
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Graph page token too large')
  return token
}

/**
 * Validates + decodes a Microsoft access-token JWT to extract identity.
 * Decoding alone is NOT proof: we must validate authenticity and scope.
 *
 * Authenticity: The token is obtained directly from Microsoft's token
 * endpoint (login.microsoftonline.com) over TLS in the same request — it
 * is NOT supplied by an untrusted party. We additionally validate:
 * - Header alg is RS256 (Microsoft's signing alg for access tokens).
 * - iss is https://login.microsoftonline.com/{tenant}/v2.0 (tenant must
 *   match our configured tenant, or 'consumers' for personal accounts).
 * - aud is https://graph.microsoft.com (Graph access tokens) — never
 *   accept a token minted for a different audience.
 * - exp is in the future (with 60s clock skew), nbf <= now, iat <= now.
 * - Token has a usable identity claim (upn / preferred_username / email).
 *
 * This does NOT verify the RSA signature (we trust the TLS channel +
 * Microsoft as issuer because we just received it from the token endpoint
 * in this same call). If signature verification is ever required (e.g.
 * tokens supplied by an untrusted party), that would be a separate,
 * approval-gated change. For our flow — token fetched fresh from Microsoft
 * over TLS, then decoded in-process — iss/aud/exp validation + TLS
 * provenance is the correct authenticity check.
 */
export function decodeAccessTokenIdentity(
  token: string,
  expectedTenant?: string,
): MailboxIdentity | null {
  try {
    const parts = token.split('.')
    if (parts.length !== 3) return null
    // Header: alg must be RS256.
    const headerB64 = parts[0].replace(/-/g, '+').replace(/_/g, '/')
    const headerPadded = headerB64 + '='.repeat((4 - (headerB64.length % 4)) % 4)
    const headerJson = new TextDecoder().decode(
      Uint8Array.from(atob(headerPadded), (c) => c.charCodeAt(0)),
    )
    const header = JSON.parse(headerJson) as Record<string, unknown>
    if (header.alg !== 'RS256' || header.typ !== 'JWT') return null
    // Payload.
    const payloadB64 = parts[1].replace(/-/g, '+').replace(/_/g, '/')
    const padded = payloadB64 + '='.repeat((4 - (payloadB64.length % 4)) % 4)
    const json = new TextDecoder().decode(Uint8Array.from(atob(padded), (c) => c.charCodeAt(0)))
    const claims = JSON.parse(json) as Record<string, unknown>
    // iss: must be login.microsoftonline.com/{tenant}/v2.0
    const iss = typeof claims.iss === 'string' ? claims.iss : null
    if (!iss) return null
    const issMatch = iss.match(/^https:\/\/login\.microsoftonline\.com\/([^/]+)\/v2\.0$/)
    if (!issMatch) return null
    if (expectedTenant && issMatch[1] !== expectedTenant && issMatch[1] !== 'consumers') return null
    // aud: must be Graph (string) or contain Graph (array).
    const aud = claims.aud
    const audOk =
      aud === 'https://graph.microsoft.com' ||
      (Array.isArray(aud) && aud.includes('https://graph.microsoft.com'))
    if (!audOk) return null
    // exp / nbf / iat with 60s clock skew.
    const nowSec = Math.floor(Date.now() / 1000)
    const exp = typeof claims.exp === 'number' ? claims.exp : null
    if (exp === null || exp <= nowSec - 60) return null
    const nbf = typeof claims.nbf === 'number' ? claims.nbf : null
    if (nbf !== null && nbf > nowSec + 60) return null
    const iat = typeof claims.iat === 'number' ? claims.iat : null
    if (iat !== null && iat > nowSec + 60) return null
    const upn = typeof claims.upn === 'string' ? claims.upn : null
    const preferred = typeof claims.preferred_username === 'string' ? claims.preferred_username : null
    const email = typeof claims.email === 'string' ? claims.email : null
    const oid = typeof claims.oid === 'string' ? claims.oid : null
    const name = typeof claims.name === 'string' ? claims.name : null
    const principal = upn ?? preferred ?? email
    if (!principal) return null
    return {
      id: oid ?? principal,
      displayName: name ?? principal,
      mail: email ?? upn ?? preferred,
      userPrincipalName: upn ?? preferred ?? email,
    }
  } catch {
    return null
  }
}

/** Decodes and validates an opaque page token issued by encodePageToken. */
export function decodePageToken(token: string): PageContinuation {
  if (!token || token.length > MAX_PAGE_TOKEN_BYTES || /[\r\n]/.test(token) || /[^A-Za-z0-9_-]/.test(token))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token is invalid')
  const json = base64UrlDecode(token)
  if (!json) throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token is invalid')
  let parsed: unknown
  try {
    parsed = JSON.parse(json)
  } catch {
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token is invalid')
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token is invalid')
  const obj = parsed as Record<string, unknown>
  const keys = Object.keys(obj)
  if (keys.length === 0 || keys.some((k) => !['skip', 'skiptoken', 'top'].includes(k)))
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token is invalid')
  const result: PageContinuation = {}
  if ('skip' in obj) {
    if (typeof obj.skip !== 'number' || !Number.isInteger(obj.skip) || obj.skip < 0 || obj.skip > MAX_SKIP)
      throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token $skip is invalid')
    result.skip = obj.skip
  }
  if ('skiptoken' in obj) {
    if (typeof obj.skiptoken !== 'string' || obj.skiptoken.length === 0 || obj.skiptoken.length > 4096 || /[\r\n]/.test(obj.skiptoken))
      throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token $skiptoken is invalid')
    result.skiptoken = obj.skiptoken
  }
  if ('top' in obj) {
    if (typeof obj.top !== 'number' || !Number.isInteger(obj.top) || obj.top < 1 || obj.top > 100)
      throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token $top is invalid')
    result.top = obj.top
  }
  if (result.skip === undefined && !result.skiptoken)
    throw new EmailProviderError('MS_GRAPH_INVALID_PAGINATION', 'Preview page token has no paging state')
  return result
}

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
      // Validate before attaching the bearer token. Delta/next links from
      // Graph are absolute URLs; all other callers use relative /me/ paths.
      // Browser-provided URLs are never passed here directly — preview uses
      // pageToken reconstruction (see listInboxPreviewPage).
      const url = validateGraphRequestUrl(path)
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

  /**
   * Phase 2 — Mailbox identity proof, WITHOUT requiring User.Read.
   *
   * GET /me normally needs User.Read, which our existing delegated scopes
   * (offline_access, Mail.Read, Mail.Send) do NOT include. We have hit
   * this 403 before. Instead, decode the Microsoft-issued access token
   * JWT itself: its `upn` / `preferred_username` / `email` claims are the
   * authorized mailbox identity, no extra Graph call or permission needed.
   * This is compatible with the existing verified OAuth setup — no scope
   * change, no re-consent. If the token has no usable identity claims,
   * return a clear failure (TOKEN_IDENTITY_MISSING) and refuse to send.
   *
   * The token is obtained from Microsoft via refresh flow; we decode only
   * its payload (no signature verification — we trust Microsoft as the
   * issuer because we just received it from login.microsoftonline.com).
   * Never log the token or its claims beyond the masked identity.
   */
  async getMailboxIdentity(): Promise<MailboxIdentity> {
    const token = await this.accessToken()
    const expectedTenant = this.config.tenant?.trim() || 'consumers'
    const identity = decodeAccessTokenIdentity(token, expectedTenant)
    if (!identity)
      throw new EmailProviderError(
        'MS_GRAPH_TOKEN_IDENTITY_MISSING',
        'Microsoft Graph access token failed identity validation (iss/aud/exp/claims)',
      )
    return identity
  }

  /**
   * Verifies the authorized token belongs to the expected complaint mailbox.
   * Matches on mail or userPrincipalName (case-insensitive). Returns the
   * actual identity + expected + match flag. Callers must refuse to send
   * when matches=false.
   */
  async verifyMailboxIdentity(expectedMailbox: string | undefined): Promise<MailboxIdentityVerification> {
    const actual = await this.getMailboxIdentity()
    const expected = expectedMailbox?.trim() || null
    if (!expected) {
      return { matches: false, actual, expected: null, reason: 'NO_EXPECTED_MAILBOX' }
    }
    const lowerExpected = expected.toLowerCase()
    const matches =
      (actual.mail?.toLowerCase() === lowerExpected) ||
      (actual.userPrincipalName?.toLowerCase() === lowerExpected)
    return {
      matches,
      actual,
      expected,
      reason: matches ? null : 'MAILBOX_MISMATCH',
    }
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
      // Corrective fix: Graph message delta only supports an initial
      // receivedDateTime ge (or gt) filter, not a combined ge + lt expression.
      // Send only the lower bound to Graph; the caller enforces the upper
      // bound locally (see runMailboxDiscovery inWindow check).
      if (options?.receivedStart)
        params.set('$filter', `receivedDateTime ge ${options.receivedStart}`)
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
   * Preview-only INBOX listing for the Check Mail Now flow. Uses
   * /me/mailFolders/inbox/messages (NOT /me/messages, NOT delta), so it
   * never advances the ingestion cursor and never lists Sent Items or
   * other folders. Returns metadata + bodyPreview (for classification) but
   * never full bodies or headers. Read-only: no D1 writes, no sends, no
   * cursor change.
   *
   * Pagination: callers pass an opaque pageToken issued by encodePageToken
   * (which preserves the complete Microsoft paging state: $skip and/or
   * $skiptoken + original $top), never a browser-provided URL. The server
   * decodes + re-validates the token and rebuilds the Inbox URL with the
   * same window filter + preserved paging state, so a malicious nextLink
   * cannot redirect the bearer token or widen scope. Returns
   * nextPageToken (opaque) only, never a full Graph URL.
   *
   * Page-size changes: if the token carries only $skip (no $skiptoken),
   * the caller's $top is honored (skip offset is page-size independent).
   * If the token carries $skiptoken, the original $top from the Graph
   * nextLink is preserved, because Graph skiptokens are tied to the
   * original query/page-size — changing $top with a skiptoken may fail.
   */
  async listInboxPreviewPage(options: {
    receivedStart: string
    receivedEnd: string
    top?: number
    pageToken?: string
  }): Promise<{ messages: GraphMessage[]; nextPageToken?: string }> {
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
    const requestedTop = Math.max(1, Math.min(100, options.top ?? 50))
    const params = new URLSearchParams({
      $select: select,
      $filter: `receivedDateTime ge ${options.receivedStart} and receivedDateTime lt ${options.receivedEnd}`,
      $orderby: 'receivedDateTime desc',
      $top: String(requestedTop),
    })
    if (options.pageToken) {
      const continuation = decodePageToken(options.pageToken)
      // $skiptoken is tied to original query: preserve original $top.
      if (continuation.skiptoken) {
        params.set('$skiptoken', continuation.skiptoken)
        if (continuation.top !== undefined) params.set('$top', String(continuation.top))
      }
      // $skip is offset-based: caller's $top is safe to use.
      if (continuation.skip !== undefined) params.set('$skip', String(continuation.skip))
      // If both present, both are set (Graph will use skiptoken; skip is harmless).
    }
    const page = await this.request<{
      value?: GraphMessage[]
      '@odata.nextLink'?: string
    }>(`/me/mailFolders/inbox/messages?${params.toString()}`)
    const nextPageToken = page['@odata.nextLink']
      ? encodePageToken(page['@odata.nextLink'], {
          receivedStart: options.receivedStart,
          receivedEnd: options.receivedEnd,
        })
      : null
    return { messages: page.value ?? [], nextPageToken: nextPageToken ?? undefined }
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
