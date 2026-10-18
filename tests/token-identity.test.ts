import { describe, expect, it } from 'vitest'
import { decodeAccessTokenIdentity } from '../worker/microsoft-graph'

function makeToken(claims: Record<string, unknown>, header: Record<string, unknown> = { alg: 'RS256', typ: 'JWT' }): string {
  const h = Buffer.from(JSON.stringify(header)).toString('base64url')
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url')
  return `${h}.${payload}.fakesig`
}

const VALID_CLAIMS = {
  iss: 'https://login.microsoftonline.com/consumers/v2.0',
  aud: 'https://graph.microsoft.com',
  exp: Math.floor(Date.now() / 1000) + 3600,
  nbf: Math.floor(Date.now() / 1000) - 10,
  iat: Math.floor(Date.now() / 1000) - 10,
}

describe('Access token identity decoding (no User.Read needed)', () => {
  it('decodes upn + preferred_username + email with valid iss/aud/exp', () => {
    const token = makeToken({
      ...VALID_CLAIMS,
      upn: 'ops@franchise.example',
      preferred_username: 'ops@franchise.example',
      email: 'ops@franchise.example',
      oid: 'oid-123',
      name: 'Ops Mailbox',
    })
    const identity = decodeAccessTokenIdentity(token, 'consumers')
    expect(identity).not.toBeNull()
    expect(identity?.userPrincipalName).toBe('ops@franchise.example')
    expect(identity?.mail).toBe('ops@franchise.example')
    expect(identity?.id).toBe('oid-123')
    expect(identity?.displayName).toBe('Ops Mailbox')
  })

  it('prefers upn when email differs', () => {
    const token = makeToken({ ...VALID_CLAIMS, upn: 'primary@example.invalid', email: 'alias@example.invalid' })
    const identity = decodeAccessTokenIdentity(token, 'consumers')
    expect(identity?.userPrincipalName).toBe('primary@example.invalid')
  })

  it('returns null for non-JWT, malformed, or no identity claims', () => {
    expect(decodeAccessTokenIdentity('not-a-jwt')).toBeNull()
    expect(decodeAccessTokenIdentity('a.b')).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ ...VALID_CLAIMS, oid: 'x' }))).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ ...VALID_CLAIMS }))).toBeNull()
  })

  it('rejects wrong iss, wrong aud, expired, wrong alg', () => {
    const base = { ...VALID_CLAIMS, upn: 'ops@example.invalid' }
    expect(decodeAccessTokenIdentity(makeToken({ ...base, iss: 'https://evil.example.com/tenant/v2.0' }), 'consumers')).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ ...base, aud: 'https://other.example.com' }), 'consumers')).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ ...base, exp: Math.floor(Date.now() / 1000) - 3600 }), 'consumers')).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ ...base }, { alg: 'none', typ: 'JWT' }), 'consumers')).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ ...base, iss: 'https://login.microsoftonline.com/other-tenant/v2.0' }), 'consumers')).toBeNull()
  })
})
