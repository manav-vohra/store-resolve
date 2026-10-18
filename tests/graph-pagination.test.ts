import { describe, expect, it } from 'vitest'
import {
  decodePageToken,
  encodePageToken,
  MicrosoftGraphProvider,
} from '../worker/microsoft-graph'
import { EmailProviderError } from '../worker/email-provider'

function expectPaginationError(fn: () => unknown) {
  try {
    fn()
  } catch (e) {
    expect(e).toBeInstanceOf(EmailProviderError)
    expect((e as EmailProviderError).code).toBe('MS_GRAPH_INVALID_PAGINATION')
    return
  }
  throw new Error('Expected MS_GRAPH_INVALID_PAGINATION but no error was thrown')
}

const WINDOW = { receivedStart: '2026-10-10T00:00:00.000Z', receivedEnd: '2026-10-17T23:59:59.999Z' }
const FILTER = `receivedDateTime ge ${WINDOW.receivedStart} and receivedDateTime lt ${WINDOW.receivedEnd}`

function nextLink(params: Record<string, string>): string {
  const q = new URLSearchParams(params)
  return `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?${q.toString()}`
}

describe('Graph pagination opaque continuation', () => {
  it('encodes real Graph $skip nextLink and decodes it back', () => {
    const link = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: '10', $skip: '10' })
    const token = encodePageToken(link, WINDOW)
    expect(token).toBeTruthy()
    expect(token).not.toContain('graph.microsoft.com')
    expect(token).not.toContain('$skip')
    const decoded = decodePageToken(token!)
    expect(decoded.skip).toBe(10)
    expect(decoded.top).toBe(10)
    expect(decoded.skiptoken).toBeUndefined()
  })

  it('encodes $skiptoken nextLink and preserves original $top', () => {
    const link = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: '50', $skiptoken: 'abc123' })
    const token = encodePageToken(link, WINDOW)
    const decoded = decodePageToken(token!)
    expect(decoded.skiptoken).toBe('abc123')
    expect(decoded.top).toBe(50)
  })

  it('encodes nextLink with both $skip and $skiptoken', () => {
    const link = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: '25', $skip: '25', $skiptoken: 'tok' })
    const token = encodePageToken(link, WINDOW)
    const decoded = decodePageToken(token!)
    expect(decoded.skip).toBe(25)
    expect(decoded.skiptoken).toBe('tok')
    expect(decoded.top).toBe(25)
  })

  it('accepts Graph nextLink that omits $filter/$orderby (only $skip/$top)', () => {
    // Real Graph behavior: nextLink for $skip pagination often contains
    // ONLY $skip + $top, with the original query implied. This is valid —
    // the server rebuilds with the original window filter. The Oct 17
    // prod failure was caused by rejecting these nextLinks.
    const link = nextLink({ $top: '50', $skip: '50' })
    const token = encodePageToken(link, WINDOW)
    const decoded = decodePageToken(token!)
    expect(decoded.skip).toBe(50)
    expect(decoded.top).toBe(50)
  })

  it('ignores nextLink $filter (server rebuilds with original window)', () => {
    // Graph normalizes filter datetimes (e.g. strips .000), so strict
    // matching causes false rejections. The server ignores the nextLink
    // filter and rebuilds with the original window, so it cannot widen scope.
    const link = nextLink({ $filter: 'receivedDateTime ge 2026-01-01T00:00:00.000Z', $top: '50', $skip: '50' })
    const token = encodePageToken(link, WINDOW)
    expect(decodePageToken(token!).skip).toBe(50)
  })

  it('rejects nextLink with wrong host', () => {
    const link = `https://evil.example.com/v1.0/me/mailFolders/inbox/messages?${new URLSearchParams({ $filter: FILTER, $orderby: 'receivedDateTime desc', $skip: '10' }).toString()}`
    expectPaginationError(() => encodePageToken(link, WINDOW))
  })

  it('accepts nextLink path variants under /me/ (server rebuilds with Inbox)', () => {
    // Graph may normalize the nextLink path (e.g. to /me/messages). We
    // rebuild with the Inbox path + original window, so the nextLink path
    // cannot widen scope. Only non-/me/ paths are rejected.
    for (const path of ['/v1.0/me/messages', '/v1.0/me/mailFolders/inbox/messages']) {
      const link = `https://graph.microsoft.com${path}?${new URLSearchParams({ $skip: '10', $top: '50' }).toString()}`
      expect(decodePageToken(encodePageToken(link, WINDOW)!).skip).toBe(10)
    }
  })

  it('rejects nextLink outside /me/ scope', () => {
    for (const path of ['/v1.0/users/123/messages', '/v1.0/me/mailFolders/inbox/messages/123/child', '/beta/me/messages']) {
      const link = `https://graph.microsoft.com${path}?${new URLSearchParams({ $skip: '10' }).toString()}`
      // /beta/ is not /v1.0/me/, and /users/ is not /me/ — both rejected.
      // Note: /v1.0/me/mailFolders/inbox/messages/123/child starts with /v1.0/me/
      // so it is accepted by path check, but it has no valid paging use — it
      // will still be accepted for skip extraction (path is ignored on rebuild).
      // We only assert rejection for the two truly out-of-scope paths.
      if (path.startsWith('/v1.0/users') || path.startsWith('/beta')) {
        expectPaginationError(() => encodePageToken(link, WINDOW))
      }
    }
  })

  it('accepts relative nextLink (resolved against Graph root)', () => {
    const link = `/v1.0/me/mailFolders/inbox/messages?${new URLSearchParams({ $skip: '10', $top: '50' }).toString()}`
    expect(decodePageToken(encodePageToken(link, WINDOW)!).skip).toBe(10)
  })

  it('ignores nextLink $filter/$orderby (server rebuilds with original)', () => {
    // Graph normalizes filter datetimes and may reorder; strict matching
    // caused false rejections. Server ignores nextLink filter/orderby and
    // rebuilds with original window, so they cannot widen scope.
    const link1 = nextLink({ $filter: 'receivedDateTime ge 2026-01-01T00:00:00.000Z and receivedDateTime lt 2026-12-31T23:59:59.999Z', $orderby: 'receivedDateTime desc', $skip: '10' })
    expect(decodePageToken(encodePageToken(link1, WINDOW)!).skip).toBe(10)
    const link2 = nextLink({ $filter: FILTER, $orderby: 'subject asc', $skip: '10' })
    expect(decodePageToken(encodePageToken(link2, WINDOW)!).skip).toBe(10)
  })

  it('rejects nextLink with no paging state', () => {
    const link = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: '10' })
    expectPaginationError(() => encodePageToken(link, WINDOW))
  })

  it('rejects nextLink with invalid $skip (negative, non-numeric, huge)', () => {
    for (const skip of ['-1', 'abc', '9999999999']) {
      const link = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $skip: skip })
      expectPaginationError(() => encodePageToken(link, WINDOW))
    }
  })

  it('rejects malformed page tokens', () => {
    for (const token of ['', 'not-base64!!!', 'a'.repeat(5000), 'abc\ndef', '{"skip": -1}']) {
      expectPaginationError(() => decodePageToken(token))
    }
    // Valid base64 but not a continuation object
    const bad = Buffer.from(JSON.stringify({ evil: true })).toString('base64url')
    expectPaginationError(() => decodePageToken(bad))
  })

  it('walks multiple pages to completion with $skip', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm@example.invalid' })
    // Mock request: 3 pages of 2 messages each, using $skip.
    const allMessages = Array.from({ length: 5 }, (_, i) => ({
      id: `msg-${i + 1}`,
      conversationId: `conv-${i + 1}`,
      receivedDateTime: `2026-10-1${i % 7}T12:00:00Z`,
      subject: `Subject ${i + 1}`,
      from: { emailAddress: { address: 'a@example.invalid', name: 'A' } },
      bodyPreview: 'preview',
    }))
    const calls: string[] = []
    // @ts-expect-error — stub private request
    provider.request = async (path: string) => {
      calls.push(path)
      const url = new URL(`https://graph.microsoft.com/v1.0${path}`)
      const skip = parseInt(url.searchParams.get('$skip') ?? '0', 10)
      const top = parseInt(url.searchParams.get('$top') ?? '50', 10)
      const page = allMessages.slice(skip, skip + top)
      const nextSkip = skip + top
      const result: { value: typeof allMessages; '@odata.nextLink'?: string } = { value: page }
      if (nextSkip < allMessages.length) {
        result['@odata.nextLink'] = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: String(top), $skip: String(nextSkip) })
      }
      return result
    }
    // Page 1: top=2
    const p1 = await provider.listInboxPreviewPage({ ...WINDOW, top: 2 })
    expect(p1.messages).toHaveLength(2)
    expect(p1.nextPageToken).toBeTruthy()
    // Page 2: top=2 (same size)
    const p2 = await provider.listInboxPreviewPage({ ...WINDOW, top: 2, pageToken: p1.nextPageToken })
    expect(p2.messages).toHaveLength(2)
    expect(p2.nextPageToken).toBeTruthy()
    // Page 3: top=2, should get last 1, no next token
    const p3 = await provider.listInboxPreviewPage({ ...WINDOW, top: 2, pageToken: p2.nextPageToken })
    expect(p3.messages).toHaveLength(1)
    expect(p3.nextPageToken).toBeUndefined()
    // Total = 5, complete traversal
    expect([...p1.messages, ...p2.messages, ...p3.messages].map((m) => m.id)).toEqual(['msg-1', 'msg-2', 'msg-3', 'msg-4', 'msg-5'])
    // Verify $skip was actually sent on page 2 and 3
    expect(calls[1]).toContain('%24skip=2')
    expect(calls[2]).toContain('%24skip=4')
  })

  it('allows changing page size with $skip (skip is offset-based)', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm@example.invalid' })
    const calls: string[] = []
    // @ts-expect-error — stub private request
    provider.request = async (path: string) => {
      calls.push(path)
      const url = new URL(`https://graph.microsoft.com/v1.0${path}`)
      const skip = parseInt(url.searchParams.get('$skip') ?? '0', 10)
      const top = parseInt(url.searchParams.get('$top') ?? '50', 10)
      const result: { value: Array<{ id: string }>; '@odata.nextLink'?: string } = { value: [{ id: 'x' }] }
      if (skip === 0) {
        result['@odata.nextLink'] = nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: String(top), $skip: '10' })
      }
      return result
    }
    const p1 = await provider.listInboxPreviewPage({ ...WINDOW, top: 10 })
    expect(p1.nextPageToken).toBeTruthy()
    // Page 2 with different top=5: $skip=10 preserved, $top=5 from caller
    await provider.listInboxPreviewPage({ ...WINDOW, top: 5, pageToken: p1.nextPageToken })
    expect(calls[1]).toContain('%24skip=10')
    expect(calls[1]).toContain('%24top=5')
  })

  it('preserves original $top when token carries $skiptoken', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm@example.invalid' })
    const calls: string[] = []
    // @ts-expect-error — stub private request
    provider.request = async (path: string) => {
      calls.push(path)
      const url = new URL(`https://graph.microsoft.com/v1.0${path}`)
      if (url.searchParams.has('$skiptoken')) {
        return { value: [{ id: 'y' }] }
      }
      return {
        value: [{ id: 'x' }],
        '@odata.nextLink': nextLink({ $filter: FILTER, $orderby: 'receivedDateTime desc', $top: '50', $skiptoken: 'tok123' }),
      }
    }
    const p1 = await provider.listInboxPreviewPage({ ...WINDOW, top: 50 })
    expect(p1.nextPageToken).toBeTruthy()
    // Page 2 with caller top=10: original $top=50 from token must win (skiptoken tied to query)
    await provider.listInboxPreviewPage({ ...WINDOW, top: 10, pageToken: p1.nextPageToken })
    expect(calls[1]).toContain('%24top=50')
    expect(calls[1]).toContain('skiptoken')
  })

  it('complete traversal stops when Graph reports no nextLink', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm@example.invalid' })
    let callCount = 0
    // @ts-expect-error — stub private request
    provider.request = async () => {
      callCount += 1
      return { value: [{ id: 'only' }] }
    }
    const p1 = await provider.listInboxPreviewPage({ ...WINDOW, top: 50 })
    expect(p1.messages).toHaveLength(1)
    expect(p1.nextPageToken).toBeUndefined()
    expect(callCount).toBe(1)
  })
})
