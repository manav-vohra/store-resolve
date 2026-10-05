import type { Notification } from '../src/lib/types'

export interface NotificationProvider {
  readonly ready: boolean
  send(
    notification: Notification,
    destination: string,
  ): Promise<{ providerMessageId: string; status: 'SENT' | 'FAILED' }>
}

export interface SignalWireMessage {
  sid: string
  accountSid: string
  status: string
  from: string
  to: string
  errorCode?: string
  errorMessage?: string
}

type SignalWireConfig = {
  projectId?: string
  apiToken?: string
  spaceUrl?: string
  from?: string
  publicBaseUrl?: string
}

const e164 = /^\+[1-9]\d{7,14}$/
const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i

function signalWireOrigin(value?: string): string | undefined {
  if (!value) return undefined
  try {
    const url = new URL(value.includes('://') ? value : `https://${value}`)
    if (
      url.protocol !== 'https:' ||
      !url.hostname.endsWith('.signalwire.com') ||
      url.username ||
      url.password ||
      url.port ||
      (url.pathname !== '/' && url.pathname !== '') ||
      url.search ||
      url.hash
    )
      return undefined
    return url.origin
  } catch {
    return undefined
  }
}

function publicOrigin(value?: string): string | undefined {
  if (!value) return undefined
  try {
    const url = new URL(value)
    if (
      url.protocol !== 'https:' ||
      url.username ||
      url.password ||
      url.port ||
      (url.pathname !== '/' && url.pathname !== '') ||
      url.search ||
      url.hash
    )
      return undefined
    return url.origin
  } catch {
    return undefined
  }
}

function safeDiagnostic(value: unknown, secrets: Array<string | undefined>): string {
  let message = typeof value === 'string' ? value : 'SignalWire request failed'
  for (const secret of secrets) if (secret) message = message.split(secret).join('[redacted]')
  return message.replace(/[\r\n\t]+/g, ' ').slice(0, 300)
}

function messageFrom(
  value: unknown,
  authenticatedProjectId?: string,
): SignalWireMessage | undefined {
  if (!value || typeof value !== 'object') return undefined
  const body = value as Record<string, unknown>
  const sid = typeof body.sid === 'string' ? body.sid : body.id
  const accountSid =
    typeof body.account_sid === 'string'
      ? body.account_sid
      : typeof body.project_id === 'string'
        ? body.project_id
        : authenticatedProjectId
  if (
    typeof sid !== 'string' ||
    typeof accountSid !== 'string' ||
    typeof body.status !== 'string' ||
    typeof body.from !== 'string' ||
    typeof body.to !== 'string'
  )
    return undefined
  return {
    sid,
    accountSid,
    status: body.status,
    from: body.from,
    to: body.to,
    errorCode: typeof body.error_code === 'string' ? body.error_code : undefined,
    errorMessage: typeof body.error_message === 'string' ? body.error_message : undefined,
  }
}

export class MockNotificationProvider implements NotificationProvider {
  readonly ready = true
  async send() {
    return { providerMessageId: `mock-${crypto.randomUUID()}`, status: 'SENT' as const }
  }
}

export class SignalWireSmsProvider implements NotificationProvider {
  constructor(private readonly config: SignalWireConfig) {}

  get ready() {
    return Boolean(
      this.config.projectId &&
      uuid.test(this.config.projectId) &&
      this.config.apiToken &&
      signalWireOrigin(this.config.spaceUrl) &&
      this.config.from &&
      e164.test(this.config.from) &&
      publicOrigin(this.config.publicBaseUrl),
    )
  }

  get statusCallbackUrl() {
    const origin = publicOrigin(this.config.publicBaseUrl)
    return origin ? `${origin}/api/signalwire/status` : undefined
  }

  private get compatibilityApiBase() {
    const origin = signalWireOrigin(this.config.spaceUrl)
    return origin && this.config.projectId
      ? `${origin}/api/laml/2010-04-01/Accounts/${encodeURIComponent(this.config.projectId)}`
      : undefined
  }

  private get authorization() {
    return this.config.projectId && this.config.apiToken
      ? `Basic ${btoa(`${this.config.projectId}:${this.config.apiToken}`)}`
      : undefined
  }

  private async responseBody(response: Response): Promise<Record<string, unknown>> {
    const contentType = response.headers.get('content-type')?.toLowerCase() ?? ''
    if (!contentType.includes('json'))
      throw new Error(
        `SignalWire returned an unexpected ${contentType.split(';')[0] || 'unknown'} response (${response.status})`,
      )
    try {
      return (await response.json()) as Record<string, unknown>
    } catch {
      throw new Error(`SignalWire returned a malformed response (${response.status})`)
    }
  }

  async send(
    notification: Notification,
    destination: string,
  ): Promise<{ providerMessageId: string; status: 'SENT' | 'FAILED' }> {
    const projectId = this.config.projectId
    if (
      !this.ready ||
      !this.compatibilityApiBase ||
      !this.authorization ||
      !this.config.from ||
      !projectId
    )
      throw new Error('SignalWire is not configured')
    if (!e164.test(destination)) throw new Error('SignalWire destination must use E.164 format')
    const form = new URLSearchParams({
      To: destination,
      From: this.config.from,
      Body: notification.message,
      StatusCallback: this.statusCallbackUrl!,
    })
    const response = await fetch(`${this.compatibilityApiBase}/Messages`, {
      method: 'POST',
      headers: {
        authorization: this.authorization,
        accept: 'application/json',
        'content-type': 'application/x-www-form-urlencoded;charset=UTF-8',
      },
      body: form.toString(),
    })
    const body = await this.responseBody(response)
    const message = messageFrom(body)
    if (response.status !== 201 || !message) {
      const code = body.code ?? body.error_code
      const detail = [code ? `[${code}]` : '', body.message ?? body.error_message ?? `SignalWire request failed (${response.status})`]
        .filter(Boolean)
        .join(' ')
      throw new Error(safeDiagnostic(detail, [this.config.projectId, this.config.apiToken]))
    }
    if (message.accountSid !== this.config.projectId || !uuid.test(message.sid))
      throw new Error('SignalWire returned an unexpected message identity')
    return { providerMessageId: message.sid, status: 'SENT' }
  }

  async retrieveMessage(messageSid: string): Promise<SignalWireMessage> {
    const projectId = this.config.projectId
    if (!this.ready || !this.compatibilityApiBase || !this.authorization || !projectId)
      throw new Error('SignalWire is not configured')
    if (!uuid.test(messageSid)) throw new Error('Invalid SignalWire message identifier')
    const compatibilityResponse = await fetch(
      `${this.compatibilityApiBase}/Messages/${encodeURIComponent(messageSid)}.json`,
      { headers: { accept: 'application/json', authorization: this.authorization } },
    )
    let compatibilityBody: Record<string, unknown> = {}
    try {
      compatibilityBody = await this.responseBody(compatibilityResponse)
    } catch {
      // A legacy Relay message can produce a non-JSON compatibility response.
      // The authenticated native log fallback below remains authoritative.
    }
    const compatibilityMessage = messageFrom(compatibilityBody)
    if (
      compatibilityResponse.ok &&
      compatibilityMessage?.sid === messageSid &&
      compatibilityMessage.accountSid === projectId
    )
      return compatibilityMessage

    // Native Relay messages are authoritative in SignalWire's native log but
    // are not always addressable through Compatibility Retrieve.
    const origin = signalWireOrigin(this.config.spaceUrl)
    const nativeResponse = await fetch(
      `${origin}/api/messaging/logs/${encodeURIComponent(messageSid)}`,
      { headers: { accept: 'application/json', authorization: this.authorization } },
    )
    const nativeBody = await this.responseBody(nativeResponse)
    const nativeMessage = messageFrom(nativeBody, projectId)
    if (!nativeResponse.ok || !nativeMessage)
      throw new Error(
        safeDiagnostic(
          nativeBody.message ??
            nativeBody.error_message ??
            compatibilityBody.message ??
            compatibilityBody.error_message ??
            `SignalWire verification failed (${nativeResponse.status})`,
          [this.config.projectId, this.config.apiToken],
        ),
      )
    if (nativeMessage.sid !== messageSid || nativeMessage.accountSid !== projectId)
      throw new Error('SignalWire returned an unexpected message identity')
    return nativeMessage
  }

  async verifyConnection(): Promise<void> {
    if (!this.ready || !this.compatibilityApiBase || !this.authorization)
      throw new Error('SignalWire is not configured')
    const response = await fetch(`${this.compatibilityApiBase}/Messages.json?PageSize=1`, {
      headers: { accept: 'application/json', authorization: this.authorization },
    })
    const body = await this.responseBody(response)
    if (response.status !== 200 || !Array.isArray(body.messages))
      throw new Error(
        safeDiagnostic(
          body.message ?? body.error_message ?? `SignalWire diagnostic failed (${response.status})`,
          [this.config.projectId, this.config.apiToken],
        ),
      )
  }
}
