import { generateKeyPairSync, sign } from 'node:crypto'
import { describe, expect, it, vi } from 'vitest'
import {
  authority,
  authorizeExpectedMailbox,
  confirmExpectedMailbox,
  graphReadinessUrl,
  graphSanityUrl,
  oidcConfigurationUrl,
  pollForTokenResponse,
  requestDeviceCode,
  scopes,
  validateIdToken,
} from '../scripts/microsoft-oauth.mjs'

const clientId = '00000000-0000-4000-8000-000000000001'
const expectedMailbox = 'complaints@example.test'
const consumerTenantId = '9188040d-6c67-4c5b-b112-36a304b66dad'
const issuer = `https://login.microsoftonline.com/${consumerTenantId}/v2.0`
const jwksUri = 'https://login.microsoftonline.com/consumers/discovery/v2.0/keys'
const keyId = 'test-signing-key'
const { privateKey, publicKey } = generateKeyPairSync('rsa', { modulusLength: 2048 })
const publicJwk = publicKey.export({ format: 'jwk' })

const encode = (value: unknown) => Buffer.from(JSON.stringify(value)).toString('base64url')

const createIdToken = (
  claims: Record<string, unknown> = {},
  signingKey = privateKey,
  nowSeconds = Math.floor(Date.now() / 1000),
) => {
  const header = encode({ alg: 'RS256', kid: keyId, typ: 'JWT' })
  const payload = encode({
    iss: issuer,
    aud: clientId,
    tid: consumerTenantId,
    iat: nowSeconds,
    nbf: nowSeconds - 10,
    exp: nowSeconds + 3600,
    email: expectedMailbox,
    preferred_username: expectedMailbox,
    ...claims,
  })
  const signature = sign('RSA-SHA256', Buffer.from(`${header}.${payload}`), signingKey)
  return `${header}.${payload}.${signature.toString('base64url')}`
}

const deviceResponse = () =>
  new Response(
    JSON.stringify({
      device_code: 'opaque-device-code',
      user_code: 'ABCD-EFGH',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 1,
    }),
    { status: 200 },
  )

const oidcFetch = (idToken = createIdToken()) =>
  vi.fn(async (input: string | URL | Request) => {
    const url = String(input)
    if (url === `${authority}/oauth2/v2.0/devicecode`) return deviceResponse()
    if (url === `${authority}/oauth2/v2.0/token`)
      return new Response(
        JSON.stringify({
          access_token: 'never-log-access-token',
          refresh_token: 'never-log-refresh-token',
          id_token: idToken,
        }),
        { status: 200 },
      )
    if (url === oidcConfigurationUrl)
      return new Response(JSON.stringify({ issuer, jwks_uri: jwksUri }), { status: 200 })
    if (url === jwksUri)
      return new Response(
        JSON.stringify({
          keys: [
            {
              ...publicJwk,
              kid: keyId,
              use: 'sig',
              alg: 'RS256',
              issuer,
            },
          ],
        }),
        { status: 200 },
      )
    if (url.startsWith('https://graph.microsoft.com/v1.0/me?'))
      return new Response(
        JSON.stringify({
          id: 'user-id-123',
          displayName: 'Complaints Mailbox',
          mail: expectedMailbox,
          userPrincipalName: expectedMailbox,
        }),
        { status: 200 },
      )
    if (url === graphReadinessUrl)
      return new Response(
        JSON.stringify({
          id: 'inbox',
          displayName: 'Inbox',
          totalItemCount: 12,
          unreadItemCount: 2,
        }),
        { status: 200 },
      )
    if (url === graphSanityUrl)
      return new Response(
        JSON.stringify({
          value: [
            {
              id: 'opaque-message-id',
              receivedDateTime: '2026-08-01T18:27:00Z',
              from: { emailAddress: { address: 'masked@example.test' } },
            },
          ],
        }),
        { status: 200 },
      )
    throw new Error(`Unexpected URL: ${url}`)
  })

describe('Microsoft device-code OAuth bootstrap', () => {
  it('requests the consumer device endpoint with the exact delegated scopes', async () => {
    const fetch = vi.fn().mockResolvedValue(deviceResponse())
    const result = await requestDeviceCode(clientId, fetch)
    expect(fetch).toHaveBeenCalledTimes(1)
    const [url, init] = fetch.mock.calls[0] as [string, RequestInit]
    expect(url).toBe(`${authority}/oauth2/v2.0/devicecode`)
    const body = new URLSearchParams(String(init.body))
    expect(body.get('client_id')).toBe(clientId)
    expect(body.get('scope')).toBe('openid profile email offline_access Mail.Read Mail.Send User.Read')
    expect(body.get('scope')).toBe(scopes)
    expect(body.has('client_secret')).toBe(false)
    expect(body.has('nonce')).toBe(false)
    expect(body.has('state')).toBe(false)
    expect(result).toMatchObject({
      userCode: 'ABCD-EFGH',
      verificationUri: 'https://microsoft.com/devicelogin',
    })
  })

  it('handles pending and slow-down responses without logging tokens', async () => {
    const fetch = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ error: 'authorization_pending' }), { status: 400 }),
      )
      .mockResolvedValueOnce(new Response(JSON.stringify({ error: 'slow_down' }), { status: 400 }))
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
            access_token: 'never-log-access',
            refresh_token: 'refresh-secret',
            id_token: 'identity-secret',
          }),
          { status: 200 },
        ),
      )
    const wait = vi.fn().mockResolvedValue(undefined)
    const tokens = await pollForTokenResponse(
      clientId,
      {
        deviceCode: 'opaque',
        userCode: 'ABCD',
        verificationUri: 'https://example.invalid',
        expiresIn: 900,
        interval: 1,
      },
      fetch,
      wait,
    )
    expect(tokens).toEqual({
      accessToken: 'never-log-access',
      refreshToken: 'refresh-secret',
      idToken: 'identity-secret',
    })
    expect(wait).toHaveBeenNthCalledWith(1, 1000)
    expect(wait).toHaveBeenNthCalledWith(2, 1000)
    expect(wait).toHaveBeenNthCalledWith(3, 6000)
  })

  it('validates the expected account before storing its refresh token', async () => {
    const storeToken = vi.fn().mockResolvedValue(undefined)
    const log = vi.fn()
    const idToken = createIdToken()
    const fetchImpl = oidcFetch(idToken)
    const confirmation = await authorizeExpectedMailbox({
      clientId,
      expectedMailbox,
      fetchImpl,
      wait: vi.fn().mockResolvedValue(undefined),
      storeToken,
      log,
    })
    expect(confirmation).toEqual({
      matches: true,
      maskedIdentity: 'c***s@example.test',
      reason: undefined,
      graphReady: true,
      sanityCount: 1,
    })
    expect(storeToken).toHaveBeenCalledExactlyOnceWith('never-log-refresh-token')
    const logged = log.mock.calls.flat().join('\n')
    expect(logged).toContain('AUTHORIZED_MICROSOFT_IDENTITY=c***s@example.test')
    expect(logged).toContain('EXPECTED_MAILBOX_MATCH=YES')
    expect(logged).toContain('MICROSOFT_GRAPH_READY=YES')
    expect(logged).toContain('AUGUST_1_METADATA_RECORD_COUNT=1')
    expect(logged).not.toContain('never-log-access-token')
    expect(logged).not.toContain('never-log-refresh-token')
    expect(logged).not.toContain(idToken)
    const graphCalls = (fetchImpl.mock.calls as unknown[][]).filter(([url]) =>
      String(url).startsWith('https://graph.microsoft.com/'),
    )
    expect(graphCalls.map(([url]) => url)).toEqual([
      expect.stringContaining('/v1.0/me?'),
      graphReadinessUrl,
      graphSanityUrl,
    ])
    expect(graphSanityUrl).toContain('$top=3')
    expect(graphSanityUrl).toContain('$select=id,receivedDateTime,from')
    expect(graphSanityUrl).not.toMatch(/subject|body|preview|recipient|attachment/i)
  })

  it('rejects a wrong Microsoft account before secret storage', async () => {
    const storeToken = vi.fn().mockResolvedValue(undefined)
    const log = vi.fn()
    const token = createIdToken({
      email: 'someone@example.com',
      preferred_username: 'someone@example.com',
    })
    await expect(
      authorizeExpectedMailbox({
        clientId,
        expectedMailbox,
        fetchImpl: oidcFetch(token),
        wait: vi.fn().mockResolvedValue(undefined),
        storeToken,
        log,
      }),
    ).rejects.toThrow('does not match MS_MAILBOX_ADDRESS')
    expect(storeToken).not.toHaveBeenCalled()
    expect(log).toHaveBeenCalledWith('AUTHORIZED_MICROSOFT_IDENTITY=s***e@example.com')
    expect(log).toHaveBeenCalledWith('EXPECTED_MAILBOX_MATCH=NO')
  })

  it('accepts an absent email claim when preferred_username matches', () => {
    expect(
      confirmExpectedMailbox({ preferred_username: 'COMPLAINTS@EXAMPLE.TEST' }, expectedMailbox),
    ).toEqual({ matches: true, maskedIdentity: 'C***S@EXAMPLE.TEST', reason: undefined })
  })

  it('rejects mismatching identity claims even when one claim matches', () => {
    expect(
      confirmExpectedMailbox(
        { email: expectedMailbox, preferred_username: 'different@example.test' },
        expectedMailbox,
      ),
    ).toMatchObject({ matches: false })
  })

  it.each([
    [
      'signature',
      createIdToken({}, generateKeyPairSync('rsa', { modulusLength: 2048 }).privateKey),
    ],
    ['audience', createIdToken({ aud: '00000000-0000-4000-8000-000000000099' })],
    ['issuer', createIdToken({ iss: 'https://login.microsoftonline.com/invalid/v2.0' })],
  ])('rejects an invalid ID-token %s', async (_case, token) => {
    await expect(validateIdToken(token, clientId, oidcFetch(token))).rejects.toThrow(/invalid/)
  })

  it('rejects an ID token outside its valid lifetime', async () => {
    const now = 2_000_000_000
    const expired = createIdToken({ exp: now - 120 }, privateKey, now)
    await expect(validateIdToken(expired, clientId, oidcFetch(expired), now)).rejects.toThrow(
      'expired',
    )
    const notActive = createIdToken({ nbf: now + 120 }, privateKey, now)
    await expect(validateIdToken(notActive, clientId, oidcFetch(notActive), now)).rejects.toThrow(
      'not active',
    )
  })

  it('reports the public-client Azure setting when device authorization is rejected', async () => {
    const fetch = vi
      .fn()
      .mockResolvedValue(
        new Response(JSON.stringify({ error: 'unauthorized_client' }), { status: 400 }),
      )
    await expect(requestDeviceCode(clientId, fetch)).rejects.toThrow('Allow public client flows')
  })
})
