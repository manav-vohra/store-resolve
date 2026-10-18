import { describe, expect, it } from 'vitest'
import { decodeAccessTokenIdentity } from '../worker/microsoft-graph'

function makeToken(claims: Record<string, unknown>): string {
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url')
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url')
  return `${header}.${payload}.fakesig`
}

describe('Access token identity decoding (no User.Read needed)', () => {
  it('decodes upn + preferred_username + email', () => {
    const token = makeToken({
      upn: 'ops@franchise.example',
      preferred_username: 'ops@franchise.example',
      email: 'ops@franchise.example',
      oid: 'oid-123',
      name: 'Ops Mailbox',
    })
    const identity = decodeAccessTokenIdentity(token)
    expect(identity).not.toBeNull()
    expect(identity?.userPrincipalName).toBe('ops@franchise.example')
    expect(identity?.mail).toBe('ops@franchise.example')
    expect(identity?.id).toBe('oid-123')
    expect(identity?.displayName).toBe('Ops Mailbox')
  })

  it('prefers upn when email differs', () => {
    const token = makeToken({ upn: 'primary@example.invalid', email: 'alias@example.invalid' })
    const identity = decodeAccessTokenIdentity(token)
    expect(identity?.userPrincipalName).toBe('primary@example.invalid')
  })

  it('returns null for non-JWT, malformed, or no identity claims', () => {
    expect(decodeAccessTokenIdentity('not-a-jwt')).toBeNull()
    expect(decodeAccessTokenIdentity('a.b')).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({ oid: 'x' }))).toBeNull()
    expect(decodeAccessTokenIdentity(makeToken({}))).toBeNull()
  })
})
