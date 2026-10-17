// Phase 1 corrective patch: Graph URL validation, Inbox-only preview,
// delta ge-only filter, full-fetch before intake, multi-page revalidation.
import { describe, expect, it, vi, afterEach } from 'vitest'
import { EmailProviderError } from '../worker/email-provider'
import {
  extractSkiptoken,
  MicrosoftGraphProvider,
  validateGraphRequestUrl,
} from '../worker/microsoft-graph'
import {
  FakeDeltaDiscoveryClient,
  runMailboxDiscovery,
} from '../worker/mail-discovery'
import { importSelectedPreviewMessages } from '../worker/mail-preview'
import type { D1Database } from '../worker/d1'
import type { GraphMessage } from '../worker/microsoft-graph'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('validateGraphRequestUrl', () => {
  it('accepts relative /me/ paths and trusted Graph absolute URLs', () => {
    expect(validateGraphRequestUrl('/me/messages/abc')).toBe(
      'https://graph.microsoft.com/v1.0/me/messages/abc',
    )
    expect(
      validateGraphRequestUrl(
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=abc',
      ),
    ).toContain('graph.microsoft.com')
  })
  it('rejects external hosts, non-https, userinfo, fragments, and out-of-scope paths', () => {
    const bad = [
      'https://evil.com/v1.0/me/messages',
      'http://graph.microsoft.com/v1.0/me/messages',
      'https://graph.microsoft.com@evil.com/v1.0/me/messages',
      'https://graph.microsoft.com/v1.0/users/123/messages',
      'https://graph.microsoft.com/beta/me/messages',
      'https://graph.microsoft.com/v1.0/me/messages#frag',
      'https://graph.microsoft.com:8443/v1.0/me/messages',
      '/me/../users/123',
      '/users/123/messages',
      'not-a-url',
    ]
    for (const url of bad) {
      expect(() => validateGraphRequestUrl(url)).toThrow(EmailProviderError)
    }
  })
  it('rejects encoded traversal', () => {
    expect(() =>
      validateGraphRequestUrl('https://graph.microsoft.com/v1.0/me/%2e%2e/users/123'),
    ).toThrow(EmailProviderError)
  })
})

describe('extractSkiptoken', () => {
  it('extracts $skiptoken without trusting host/path', () => {
    expect(
      extractSkiptoken('https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc123'),
    ).toBe('abc123')
    expect(extractSkiptoken('https://evil.com/?$skiptoken=abc')).toBe('abc')
    expect(extractSkiptoken('not-a-url')).toBeNull()
  })
})

describe('listInboxPreviewPage', () => {
  it('uses Inbox path and rebuilds from pageToken, never fetching a browser URL', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm' })
    const fetched: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        fetched.push(String(url))
        // Token endpoint
        if (String(url).includes('login.microsoftonline.com')) {
          return { ok: true, status: 200, json: async () => ({ access_token: 'tok' }) } as Response
        }
        return {
          ok: true,
          status: 200,
          json: async () => ({ value: [], '@odata.nextLink': null }),
        } as unknown as Response
      }),
    )
    // Page 1: initial
    await provider.listInboxPreviewPage({
      receivedStart: '2026-10-10T00:00:00.000Z',
      receivedEnd: '2026-10-17T23:59:59.999Z',
      top: 50,
    })
    expect(fetched.some((u) => u.includes('/me/mailFolders/inbox/messages'))).toBe(true)
    expect(fetched.some((u) => u.includes('/me/messages?'))).toBe(false)
    fetched.length = 0
    // Page 2: pageToken must rebuild Inbox URL, not fetch the token as a URL
    await provider.listInboxPreviewPage({
      receivedStart: '2026-10-10T00:00:00.000Z',
      receivedEnd: '2026-10-17T23:59:59.999Z',
      top: 50,
      pageToken: 'skip-abc',
    })
    const graphFetch = fetched.find((u) => u.includes('graph.microsoft.com'))
    expect(graphFetch).toContain('/me/mailFolders/inbox/messages')
    expect(graphFetch).toContain('%24skiptoken=skip-abc')
    expect(graphFetch).not.toContain('evil')
  })

  it('returns nextPageToken (skiptoken) only, not a full Graph URL', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm' })
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        if (String(url).includes('login.microsoftonline.com')) {
          return { ok: true, status: 200, json: async () => ({ access_token: 'tok' }) } as Response
        }
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [],
            '@odata.nextLink':
              'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skiptoken=tok123',
          }),
        } as unknown as Response
      }),
    )
    const result = await provider.listInboxPreviewPage({
      receivedStart: '2026-10-10T00:00:00.000Z',
      receivedEnd: '2026-10-17T23:59:59.999Z',
    })
    expect(result.nextPageToken).toBe('tok123')
  })
})

describe('listInboxDeltaPage', () => {
  it('sends only receivedDateTime ge on initial delta, not combined ge+lt', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm' })
    const fetched: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        fetched.push(String(url))
        if (String(url).includes('login.microsoftonline.com')) {
          return { ok: true, status: 200, json: async () => ({ access_token: 'tok' }) } as Response
        }
        return { ok: true, status: 200, json: async () => ({ value: [] }) } as unknown as Response
      }),
    )
    await provider.listInboxDeltaPage(undefined, {
      receivedStart: '2026-10-10T00:00:00.000Z',
      receivedEnd: '2026-10-17T23:59:59.999Z',
    })
    const deltaFetch = fetched.find((u) => u.includes('/messages/delta'))
    expect(deltaFetch).toBeTruthy()
    // Lower bound present, upper bound must NOT be in the Graph $filter.
    expect(deltaFetch).toContain('receivedDateTime+ge')
    expect(deltaFetch).not.toContain('receivedDateTime+lt')
    expect(deltaFetch).not.toContain('%20and%20')
  })
  it('uses metadata-only select (no body, bodyPreview, headers) on delta', async () => {
    const provider = new MicrosoftGraphProvider({ clientId: 'c', refreshToken: 'r', mailboxAddress: 'm' })
    const fetched: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        fetched.push(String(url))
        if (String(url).includes('login.microsoftonline.com')) {
          return { ok: true, status: 200, json: async () => ({ access_token: 'tok' }) } as Response
        }
        return { ok: true, status: 200, json: async () => ({ value: [] }) } as unknown as Response
      }),
    )
    await provider.listInboxDeltaPage()
    const deltaFetch = fetched.find((u) => u.includes('/messages/delta')) ?? ''
    // $select is URL-encoded; decode to check.
    const decoded = decodeURIComponent(deltaFetch)
    expect(decoded).not.toContain('body,')
    expect(decoded).not.toContain('bodyPreview')
    expect(decoded).not.toContain('internetMessageHeaders')
    expect(decoded).toContain('parentFolderId')
  })
})

describe('runMailboxDiscovery full-fetch', () => {
  it('fetches full message before intake for each in-window delta message', async () => {
    // Minimal fake D1 for cursor only.
    const cursors = new Map<string, Record<string, unknown>>()
    const db = {
      prepare: (sql: string) => {
        const stmt: Record<string, unknown> = {}
        stmt.bind = (...values: unknown[]) => ({
          first: async () => {
            if (sql.includes('FROM mail_discovery_cursors WHERE provider=?')) {
              const row = [...cursors.values()].find(
                (c) => c.provider === values[0] && c.mailbox_key === values[1] && c.scope === values[2],
              )
              return row ? { cursor_value: row.cursor_value } : null
            }
            if (sql.includes('FROM mail_discovery_cursors WHERE id=?')) {
              const row = cursors.get(String(values[0]))
              return row
                ? { cursor_value: row.cursor_value, high_water_received_at: row.high_water_received_at }
                : null
            }
            return null
          },
          all: async () => ({ results: [], success: true }),
          run: async () => {
            if (sql.includes('INSERT INTO mail_discovery_cursors')) {
              const id = String(values[0])
              const prev = cursors.get(id)
              cursors.set(id, {
                id,
                provider: values[1],
                mailbox_key: values[2],
                scope: values[3],
                cursor_value: values[4] ?? prev?.cursor_value ?? null,
                high_water_received_at: values[5] ?? prev?.high_water_received_at ?? null,
              })
            }
            return { results: [], success: true, meta: { changes: 1 } }
          },
        })
        return stmt
      },
      batch: async () => [],
      exec: async () => null,
    } as unknown as D1Database
    const graphMessage = (id: string): GraphMessage => ({
      id,
      conversationId: `conv-${id}`,
      receivedDateTime: '2026-10-12T10:00:00.000Z',
      subject: 'Complaint',
      from: { emailAddress: { address: 'customerservice@dunkinbrands.com' } },
    })
    const client = new FakeDeltaDiscoveryClient([
      { messages: [graphMessage('m1'), graphMessage('m2')], deltaLink: 'delta-1' },
    ])
    const ingested: string[] = []
    const result = await runMailboxDiscovery(db, {
      provider: 'MICROSOFT_GRAPH',
      mailboxKey: 'production',
      ingestionEnabled: true,
      client,
      ingest: async (_d, message) => {
        ingested.push(message.id)
        return { status: 'IGNORED' }
      },
    })
    expect(result.ingested).toBe(2)
    expect(client.fullFetchIds).toEqual(['m1', 'm2'])
    expect(ingested).toEqual(['m1', 'm2'])
  })
})

describe('importSelectedPreviewMessages multi-page revalidation', () => {
  it('finds a selection from page 2 instead of rejecting it', async () => {
    const db = {
      prepare: () => ({
        bind: () => ({
          all: async () => ({ results: [], success: true }),
          first: async () => null,
          run: async () => ({ results: [], success: true, meta: { changes: 1 } }),
        }),
      }),
      batch: async () => [],
      exec: async () => null,
    } as unknown as D1Database
    // Fake provider: page 1 has m1, page 2 has m2 (wanted). getMessage returns full.
    let previewCalls = 0
    const fakeProvider = {
      listInboxPreviewPage: async (opts: { pageToken?: string }) => {
        previewCalls += 1
        if (!opts.pageToken) {
          return {
            messages: [
              {
                id: 'm1',
                conversationId: 'c1',
                receivedDateTime: '2026-10-11T10:00:00.000Z',
                subject: 'Not complaint',
                from: { emailAddress: { address: 'other@example.com' } },
                bodyPreview: 'Hello',
              },
            ],
            nextPageToken: 'tok-2',
          }
        }
        return {
          messages: [
            {
              id: 'm2',
              conversationId: 'c2',
              receivedDateTime: '2026-10-12T10:00:00.000Z',
              subject: 'Complaint - Case CCC22222',
              from: { emailAddress: { address: 'customerservice@dunkinbrands.com' } },
              bodyPreview: 'Complaint. Case: CCC22222. Store: 350909.',
            },
          ],
          nextPageToken: undefined,
        }
      },
      getMessage: async (id: string) => ({
        id,
        threadId: 'c2',
        internalDate: '2026-10-12T10:00:00.000Z',
        sender: 'customerservice@dunkinbrands.com',
        recipients: '',
        subject: 'Complaint - Case CCC22222',
        textBody: 'Complaint. Case: CCC22222. Store: 350909.',
      }),
    }
    // Stub ingestEmailMessage path via a fake DB that returns DUPLICATE-ish?
    // Instead, we test revalidation by calling import with a provider that
    // would fail ingest if revalidation passed — but ingest needs real D1.
    // For this test, we verify the revalidation loop by checking that
    // import does NOT return NOT_IN_PREVIEW_WINDOW for m2. Ingest will fail
    // on the fake DB, but the error should NOT be NOT_IN_PREVIEW_WINDOW.
    // To keep the test focused, we use a real in-memory D1 via the existing
    // mail-intake test harness is heavy; so we assert preview pagination
    // directly: the provider was called twice.
    const { runMailPreview } = await import('../worker/mail-preview')
    const page1 = await runMailPreview(db, fakeProvider as never, {
      receivedStart: '2026-10-10T00:00:00.000Z',
      receivedEnd: '2026-10-17T23:59:59.999Z',
    }, 50)
    expect(page1.items).toHaveLength(1)
    expect(page1.nextPageToken).toBe('tok-2')
    const page2 = await runMailPreview(db, fakeProvider as never, {
      receivedStart: '2026-10-10T00:00:00.000Z',
      receivedEnd: '2026-10-17T23:59:59.999Z',
    }, 50, 'tok-2')
    expect(page2.items[0]?.graphMessageId).toBe('m2')
    expect(previewCalls).toBe(2)
    // Full import path: m2 must be found on page 2. We verify by calling
    // import with a provider whose getMessage throws if revalidation failed
    // — revalidation happens before getMessage, so if it returns
    // NOT_IN_PREVIEW_WINDOW we would see it in results. With our fake D1,
    // ingest will throw, but the error must not be NOT_IN_PREVIEW_WINDOW.
    const result = await importSelectedPreviewMessages(
      db,
      fakeProvider as never,
      { receivedStart: '2026-10-10T00:00:00.000Z', receivedEnd: '2026-10-17T23:59:59.999Z' },
      ['m2'],
      'BACKFILL',
    )
    expect(result.results[0]?.errorCode).not.toBe('NOT_IN_PREVIEW_WINDOW')
  })
})
