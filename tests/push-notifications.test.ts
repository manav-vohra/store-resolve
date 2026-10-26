// Stage A: Web Push subscription endpoints, VAPID/crypto, gated test delivery,
// deep-link parsing, and payload privacy.
import { readFile } from 'node:fs/promises'
import { DatabaseSync, type SQLInputValue, type StatementSync } from 'node:sqlite'
import { fileURLToPath } from 'node:url'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import { parseComplaintHash } from '../src/lib/deep-links'
import {
  b64urlDecode,
  b64urlEncode,
  buildTestPushPayload,
  derivePushKeys,
  encryptPayload,
  sendWebPush,
} from '../worker/webpush'
import worker from '../worker/index'

class SqliteStatement implements D1PreparedStatement {
  private values: SQLInputValue[] = []
  constructor(private readonly statement: StatementSync) {}
  bind(...values: unknown[]) {
    const normalized: SQLInputValue[] = []
    for (const value of values) {
      if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'bigint') {
        normalized.push(value)
        continue
      }
      throw new TypeError('Unsupported SQLite test binding')
    }
    this.values = normalized
    return this
  }
  async all<T>(): Promise<D1Result<T>> {
    return { success: true, results: this.statement.all(...this.values) as T[] }
  }
  async first<T>(): Promise<T | null> {
    return (this.statement.get(...this.values) as T | undefined) ?? null
  }
  async run(): Promise<D1Result> {
    const result = this.statement.run(...this.values)
    return { success: true, results: [], meta: { changes: Number(result.changes) } }
  }
}

class SqliteD1 implements D1Database {
  constructor(private readonly database: DatabaseSync) {}
  prepare(sql: string) {
    return new SqliteStatement(this.database.prepare(sql))
  }
  async batch<T>(): Promise<D1Result<T>[]> {
    return []
  }
  async exec(sql: string): Promise<D1Result> {
    this.database.exec(sql)
    return { success: true, results: [] }
  }
}

const NOW = '2026-10-11T15:00:00.000Z'
const te = new TextEncoder()

async function hmac(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const k = await crypto.subtle.importKey('raw', key, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign'])
  return new Uint8Array(await crypto.subtle.sign('HMAC', k, data))
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let off = 0
  for (const p of parts) {
    out.set(p, off)
    off += p.length
  }
  return out
}

/**
 * Independent client-side decryption, written directly from the RFC 8291 §3.4
 * pseudocode (not by mirroring the sender's code path). The per-message salt
 * from the header is mixed into the PRK step.
 */
async function clientDecrypt(
  body: Uint8Array,
  clientPrivateKey: CryptoKey,
  clientPublicRaw: Uint8Array,
  authSecret: Uint8Array,
): Promise<string> {
  const salt = body.slice(0, 16)
  const serverPublicRaw = body.slice(21, 21 + 65)
  const serverKey = await crypto.subtle.importKey(
    'raw',
    serverPublicRaw,
    { name: 'ECDH', namedCurve: 'P-256' },
    true,
    [],
  )
  const ecdhSecret = new Uint8Array(
    await crypto.subtle.deriveBits({ name: 'ECDH', public: serverKey }, clientPrivateKey, 256),
  )
  const prkKey = await hmac(authSecret, ecdhSecret)
  const keyInfo = concat([
    te.encode('WebPush: info'),
    new Uint8Array([0x00]),
    clientPublicRaw,
    serverPublicRaw,
  ])
  const ikm = (await hmac(prkKey, concat([keyInfo, new Uint8Array([0x01])]))).slice(0, 32)
  const prk = await hmac(salt, ikm)
  const cek = (
    await hmac(prk, concat([te.encode('Content-Encoding: aes128gcm\0'), new Uint8Array([0x01])]))
  ).slice(0, 16)
  const nonce = (
    await hmac(prk, concat([te.encode('Content-Encoding: nonce\0'), new Uint8Array([0x01])]))
  ).slice(0, 12)
  const ciphertext = body.slice(21 + 65)
  const aesKey = await crypto.subtle.importKey('raw', cek, { name: 'AES-GCM' }, false, ['decrypt'])
  const record = new Uint8Array(
    await crypto.subtle.decrypt({ name: 'AES-GCM', iv: nonce }, aesKey, ciphertext),
  )
  expect(record[record.length - 1]).toBe(0x02)
  return new TextDecoder().decode(record.slice(0, -1))
}

describe('webpush crypto', () => {
  it('base64url round-trips binary data', () => {
    const bytes = crypto.getRandomValues(new Uint8Array(64))
    expect(b64urlDecode(b64urlEncode(bytes))).toEqual(bytes)
  })

  // Published test vector: RFC 8291, Appendix A ("Intermediate Values for
  // Encryption"). Only the base64url values are reproduced here, no prose.
  const RFC8291_A = {
    plaintext: 'V2hlbiBJIGdyb3cgdXAsIEkgd2FudCB0byBiZSBhIHdhdGVybWVsb24',
    as_public:
      'BP4z9KsN6nGRTbVYI_c7VJSPQTBtkgcy27mlmlMoZIIgDll6e3vCYLocInmYWAmS6TlzAC8wEqKK6PBru3jl7A8',
    as_private: 'yfWPiYE-n46HLnH0KqZOF1fJJU3MYrct3AELtAQ-oRw',
    ua_public:
      'BCVxsr7N_eNgVRqvHtD0zTZsEc6-VV-JvLexhqUzORcxaOzi6-AYWXvTBHm4bjyPjs7Vd8pZGH6SRpkNtoIAiw4',
    salt: 'DGv6ra1nlYgDCS1FRnbzlw',
    auth_secret: 'BTBZMqHH6r4Tts7J_aSIgg',
    ecdh_secret: 'kyrL1jIIOHEzg3sM2ZWRHDRB62YACZhhSlknJ672kSs',
    prk_key: 'Snr3JMxaHVDXHWJn5wdC52WjpCtd2EIEGBykDcZW32k',
    ikm: 'S4lYMb_L0FxCeq0WhDx813KgSYqU26kOyzWUdsXYyrg',
    prk: '09_eUZGrsvxChDCGRCdkLiDXrReGOEVeSCdCcPBSJSc',
    cek: 'oIhVW04MRdy2XN9CiKLxTg',
    nonce: '4h_95klXJ5E_qnoN',
    ciphertext:
      '8pfeW0KbunFT06SuDKoJH9Ql87S1QUrdirN6GcG7sFz1y1sqLgVi1VhjVkHsUoEsbI_0LpXMuGvnzQ',
  }

  async function rfcServerKeys(): Promise<CryptoKeyPair> {
    const asPubRaw = b64urlDecode(RFC8291_A.as_public)
    const publicKey = await crypto.subtle.importKey(
      'raw',
      asPubRaw,
      { name: 'ECDH', namedCurve: 'P-256' },
      true,
      [],
    )
    const privateKey = await crypto.subtle.importKey(
      'jwk',
      {
        kty: 'EC',
        crv: 'P-256',
        x: b64urlEncode(asPubRaw.slice(1, 33)),
        y: b64urlEncode(asPubRaw.slice(33, 65)),
        d: RFC8291_A.as_private,
      },
      { name: 'ECDH', namedCurve: 'P-256' },
      true,
      ['deriveBits'],
    )
    return { publicKey, privateKey }
  }

  it('matches the RFC 8291 Appendix A intermediate key values', async () => {
    const serverKeys = await rfcServerKeys()
    const keys = await derivePushKeys(
      b64urlDecode(RFC8291_A.ua_public),
      b64urlDecode(RFC8291_A.auth_secret),
      serverKeys.privateKey,
      b64urlDecode(RFC8291_A.as_public),
      b64urlDecode(RFC8291_A.salt),
    )
    expect(b64urlEncode(keys.ecdhSecret)).toBe(RFC8291_A.ecdh_secret)
    expect(b64urlEncode(keys.prkKey)).toBe(RFC8291_A.prk_key)
    expect(b64urlEncode(keys.ikm)).toBe(RFC8291_A.ikm)
    expect(b64urlEncode(keys.prk)).toBe(RFC8291_A.prk)
    expect(b64urlEncode(keys.cek)).toBe(RFC8291_A.cek)
    expect(b64urlEncode(keys.nonce)).toBe(RFC8291_A.nonce)
  })

  it('reproduces the RFC 8291 Appendix A encrypted message byte-for-byte', async () => {
    const serverKeys = await rfcServerKeys()
    const plaintext = new TextDecoder().decode(b64urlDecode(RFC8291_A.plaintext))
    expect(plaintext).toBe('When I grow up, I want to be a watermelon')
    const { bytes } = await encryptPayload(
      {
        endpoint: 'https://push.example.invalid/rfc',
        p256dh: RFC8291_A.ua_public,
        auth: RFC8291_A.auth_secret,
      },
      plaintext,
      { salt: b64urlDecode(RFC8291_A.salt), serverKeys },
    )
    // Expected: 86-octet header (salt || rs=4096 || idlen=65 || as_public)
    // followed by the RFC ciphertext.
    const asPubRaw = b64urlDecode(RFC8291_A.as_public)
    const expectedHeader = concat([
      b64urlDecode(RFC8291_A.salt),
      new Uint8Array([0x00, 0x00, 0x10, 0x00, 65]),
      asPubRaw,
    ])
    expect(expectedHeader.length).toBe(86)
    const expected = concat([expectedHeader, b64urlDecode(RFC8291_A.ciphertext)])
    expect(bytes).toEqual(expected)
  })

  it('verifies the VAPID JWT signature independently with the public key', async () => {
    const vapidKeys = await crypto.subtle.generateKey(
      { name: 'ECDSA', namedCurve: 'P-256' },
      true,
      ['sign', 'verify'],
    )
    const vapidJwk = await crypto.subtle.exportKey('jwk', vapidKeys.privateKey)
    const vapidPublicRaw = new Uint8Array([
      0x04,
      ...b64urlDecode(vapidJwk.x!),
      ...b64urlDecode(vapidJwk.y!),
    ])
    let captured: { headers: Record<string, string> } | undefined
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url: string, init: { headers: Record<string, string> }) => {
        captured = { headers: init.headers }
        return new Response(null, { status: 201 })
      }),
    )
    try {
      await sendWebPush(
        {
          publicKey: b64urlEncode(vapidPublicRaw),
          privateKey: vapidJwk.d!,
          subject: 'https://store-resolve.moogzie.workers.dev',
        },
        {
          endpoint: 'https://push.example.invalid/sub/vapid-1',
          p256dh: RFC8291_A.ua_public,
          auth: RFC8291_A.auth_secret,
        },
        { title: 't', body: 'b', url: '/' },
      )
      const authz = captured!.headers['Authorization']
      const match = /^vapid t=([^,]+), k=([A-Za-z0-9_-]+)$/.exec(authz)
      expect(match).toBeTruthy()
      const [, jwt, keyParam] = match!
      expect(keyParam).toBe(b64urlEncode(vapidPublicRaw))
      const [h, p, s] = jwt.split('.')
      // Claims.
      expect(JSON.parse(new TextDecoder().decode(b64urlDecode(h)))).toMatchObject({
        typ: 'JWT',
        alg: 'ES256',
      })
      const claims = JSON.parse(new TextDecoder().decode(b64urlDecode(p)))
      expect(claims.aud).toBe('https://push.example.invalid')
      expect(claims.sub).toBe('https://store-resolve.moogzie.workers.dev')
      expect(claims.exp).toBeGreaterThan(Date.now() / 1000)
      // Signature is raw P1363 r||s (64 bytes, not DER).
      const sig = b64urlDecode(s)
      expect(sig.length).toBe(64)
      // Independent verification with the public key only.
      const verifyKey = await crypto.subtle.importKey(
        'jwk',
        { kty: 'EC', crv: 'P-256', x: vapidJwk.x!, y: vapidJwk.y! },
        { name: 'ECDSA', namedCurve: 'P-256' },
        false,
        ['verify'],
      )
      const valid = await crypto.subtle.verify(
        { name: 'ECDSA', hash: 'SHA-256' },
        verifyKey,
        sig,
        te.encode(`${h}.${p}`),
      )
      expect(valid).toBe(true)
    } finally {
      vi.unstubAllGlobals()
    }
  })


  it('encrypts a payload the subscription owner can decrypt (round-trip)', async () => {
    const clientKeys = await crypto.subtle.generateKey(
      { name: 'ECDH', namedCurve: 'P-256' },
      true,
      ['deriveBits'],
    )
    const clientPublicRaw = new Uint8Array(await crypto.subtle.exportKey('raw', clientKeys.publicKey))
    const authSecret = crypto.getRandomValues(new Uint8Array(16))
    const vapidKeys = await crypto.subtle.generateKey(
      { name: 'ECDSA', namedCurve: 'P-256' },
      true,
      ['sign', 'verify'],
    )
    const vapidJwk = await crypto.subtle.exportKey('jwk', vapidKeys.privateKey)
    const vapidPublicRaw = new Uint8Array([
      0x04,
      ...b64urlDecode(vapidJwk.x!),
      ...b64urlDecode(vapidJwk.y!),
    ])
    let captured: { headers: Record<string, string>; body: Uint8Array } | undefined
    vi.stubGlobal(
      'fetch',
      vi.fn(async (_url: string, init: { headers: Record<string, string>; body: Uint8Array }) => {
        captured = { headers: init.headers, body: init.body }
        return new Response(null, { status: 201 })
      }),
    )
    try {
      const payload = { title: 'T', body: 'B', url: '/#/complaints/X', badgeCount: 2 }
      const result = await sendWebPush(
        {
          publicKey: b64urlEncode(vapidPublicRaw),
          privateKey: vapidJwk.d!,
          subject: 'https://example.invalid',
        },
        {
          endpoint: 'https://push.example.invalid/sub/1',
          p256dh: b64urlEncode(clientPublicRaw),
          auth: b64urlEncode(authSecret),
        },
        payload,
      )
      expect(result).toEqual({ accepted: true, expired: false })
      expect(captured).toBeDefined()
      expect(captured!.headers['Content-Encoding']).toBe('aes128gcm')
      expect(captured!.headers['Authorization']).toMatch(/^vapid t=[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+, k=[A-Za-z0-9_-]+$/)
      // VAPID JWT header decodes to ES256.
      const jwt = captured!.headers['Authorization'].match(/^vapid t=([^,]+),/)![1]
      const header = JSON.parse(new TextDecoder().decode(b64urlDecode(jwt.split('.')[0])))
      expect(header).toMatchObject({ typ: 'JWT', alg: 'ES256' })
      // Independent client-side decryption recovers the payload.
      const decrypted = await clientDecrypt(captured!.body, clientKeys.privateKey, clientPublicRaw, authSecret)
      expect(JSON.parse(decrypted)).toEqual(payload)
    } finally {
      vi.unstubAllGlobals()
    }
  })

  async function makeVapid() {
    const vapidKeys = await crypto.subtle.generateKey(
      { name: 'ECDSA', namedCurve: 'P-256' },
      true,
      ['sign', 'verify'],
    )
    const vapidJwk = await crypto.subtle.exportKey('jwk', vapidKeys.privateKey)
    const vapidPublicRaw = new Uint8Array([
      0x04,
      ...b64urlDecode(vapidJwk.x!),
      ...b64urlDecode(vapidJwk.y!),
    ])
    return {
      publicKey: b64urlEncode(vapidPublicRaw),
      privateKey: vapidJwk.d!,
      subject: 'https://example.invalid',
    }
  }

  async function makeSubscription() {
    const clientKeys = await crypto.subtle.generateKey(
      { name: 'ECDH', namedCurve: 'P-256' },
      true,
      ['deriveBits'],
    )
    const clientPublicRaw = new Uint8Array(await crypto.subtle.exportKey('raw', clientKeys.publicKey))
    return {
      endpoint: 'https://push.example.invalid/sub/1',
      p256dh: b64urlEncode(clientPublicRaw),
      auth: b64urlEncode(crypto.getRandomValues(new Uint8Array(16))),
    }
  }

  it('maps 404/410 to expired', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => new Response(null, { status: 410 })))
    try {
      await expect(
        sendWebPush(await makeVapid(), await makeSubscription(), { title: 't', body: 'b', url: '/' }),
      ).resolves.toEqual({ accepted: false, expired: true })
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('throws on other push-service failures', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => new Response('gone', { status: 500 })))
    try {
      await expect(
        sendWebPush(await makeVapid(), await makeSubscription(), { title: 't', body: 'b', url: '/' }),
      ).rejects.toThrow('500')
    } finally {
      vi.unstubAllGlobals()
    }
  })

  it('rejects non-https endpoints and missing VAPID config', async () => {
    await expect(
      sendWebPush(await makeVapid(), { ...(await makeSubscription()), endpoint: 'http://push.example.invalid/x' }, { title: 't', body: 'b', url: '/' }),
    ).rejects.toThrow('https')
    await expect(
      sendWebPush({ publicKey: '', privateKey: '', subject: '' }, await makeSubscription(), { title: 't', body: 'b', url: '/' }),
    ).rejects.toThrow('not configured')
  })
})

describe('deep links', () => {
  it('parses #/complaints/:id and rejects anything else', () => {
    expect(parseComplaintHash('#/complaints/SR-2026-0002')).toEqual({ id: 'SR-2026-0002', section: undefined })
    expect(parseComplaintHash('#/complaints/SR-2026-0002/')).toEqual({ id: 'SR-2026-0002', section: undefined })
    expect(parseComplaintHash('#/complaints/SR-2026-0002/reply')).toEqual({ id: 'SR-2026-0002', section: 'reply' })
    expect(parseComplaintHash('#/')).toBeUndefined()
    expect(parseComplaintHash('')).toBeUndefined()
    expect(parseComplaintHash('#/complaints/')).toBeUndefined()
    expect(parseComplaintHash('#/complaints/ABC/../../etc')).toBeUndefined()
    expect(parseComplaintHash('#/settings')).toBeUndefined()
  })
})

describe('test push payload privacy', () => {
  it('carries no customer PII', () => {
    const payload = buildTestPushPayload(3)
    const serialized = JSON.stringify(payload)
    expect(payload.url).toBe('/')
    expect(serialized).not.toMatch(/customer|guest|complaint_text/i)
    expect(payload.title).toMatch(/TEST/i)
  })
})

describe('push subscription endpoints', () => {
  let sqlite: DatabaseSync
  let db: D1Database

  const MIGRATIONS = [
    '0001_initial.sql',
    '0002_persistence_and_auth.sql',
    '0003_pilot_admin_recipient.sql',
    '0004_v1_operations.sql',
    '0005_microsoft_graph_provider.sql',
    '0006_intelligent_mail_foundation.sql',
    '0008_push_subscriptions.sql',
  ]

  async function seed() {
    await db
      .prepare(`INSERT OR IGNORE INTO organizations(id,name,created_at,updated_at) VALUES('org-1','Test Org',?,?)`)
      .bind(NOW, NOW)
      .run()
    for (const [id, role, kind] of [
      ['dad', 'OWNER', 'STANDARD'],
      ['mgr-1', 'STORE_MANAGER', 'STANDARD'],
    ] as const) {
      await db
        .prepare(
          `INSERT INTO users(id,organization_id,name,email,phone,role,active,sms_enabled,timezone,created_at,updated_at,recipient_kind,complaint_notifications_enabled)
           VALUES(?, 'org-1', ?, ?, '+10000000000', ?, 1, 1, 'America/New_York', ?, ?, ?, 1)`,
        )
        .bind(id, id === 'dad' ? 'Dad' : 'Manager', `${id}@example.invalid`, role, NOW, NOW, kind)
        .run()
    }
  }

  function envFor(userId: string, extra: Record<string, string> = {}) {
    return {
      DB: db,
      DEV_AUTH_USER_ID: userId,
      VAPID_PUBLIC_KEY: 'test-public-key',
      VAPID_PRIVATE_KEY: 'test-private-key',
      ...extra,
    } as unknown as Parameters<typeof worker.fetch>[1]
  }

  function authed(path: string, init?: RequestInit) {
    return new Request(`https://app.test${path}`, init)
  }

  beforeEach(async () => {
    sqlite = new DatabaseSync(':memory:')
    db = new SqliteD1(sqlite)
    for (const name of MIGRATIONS) {
      const path = fileURLToPath(new URL(`../drizzle/${name}`, import.meta.url))
      await db.exec(await readFile(path, 'utf8'))
    }
    await seed()
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('serves the VAPID public key, 503 when unconfigured', async () => {
    const ok = await worker.fetch(authed('/api/push/public-key'), envFor('dad'))
    expect(ok.status).toBe(200)
    expect(await ok.json()).toEqual({ publicKey: 'test-public-key' })

    const missing = await worker.fetch(authed('/api/push/public-key'), envFor('dad', { VAPID_PUBLIC_KEY: '' }))
    expect(missing.status).toBe(503)
  })

  it('subscribes, upserts, and scopes rows to the owning user', async () => {
    const body = {
      endpoint: 'https://push.example.invalid/sub/dad-1',
      keys: { p256dh: 'aGVsbG8td29ybGQ', auth: 'YXV0aHNlY3JldA' },
      deviceLabel: 'Test iPhone',
    }
    const first = await worker.fetch(
      authed('/api/push/subscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      }),
      envFor('dad'),
    )
    expect(first.status).toBe(200)

    const row = await db
      .prepare('SELECT user_id, device_label, revoked_at FROM push_subscriptions WHERE endpoint=?')
      .bind(body.endpoint)
      .first<{ user_id: string; device_label: string; revoked_at: string | null }>()
    expect(row?.user_id).toBe('dad')
    expect(row?.device_label).toBe('Test iPhone')
    expect(row?.revoked_at).toBeNull()

    // Same user re-subscribing the same endpoint upserts instead of duplicating.
    const second = await worker.fetch(
      authed('/api/push/subscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ ...body, deviceLabel: 'Renamed' }),
      }),
      envFor('dad'),
    )
    expect(second.status).toBe(200)
    const count = await db
      .prepare('SELECT COUNT(*) c FROM push_subscriptions WHERE endpoint=?')
      .bind(body.endpoint)
      .first<{ c: number }>()
    expect(count?.c).toBe(1)

    // Another user cannot claim the same endpoint.
    const hijack = await worker.fetch(
      authed('/api/push/subscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      }),
      envFor('mgr-1'),
    )
    expect(hijack.status).toBe(409)

    // Validation: non-https endpoints and malformed keys are rejected.
    const badUrl = await worker.fetch(
      authed('/api/push/subscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ ...body, endpoint: 'http://push.example.invalid/x' }),
      }),
      envFor('dad'),
    )
    expect(badUrl.status).toBe(400)
  })

  it('unsubscribes only the caller\u2019s own rows', async () => {
    const endpoint = 'https://push.example.invalid/sub/dad-2'
    await db
      .prepare(
        `INSERT INTO push_subscriptions(id,organization_id,user_id,endpoint,p256dh_key,auth_key,device_label,created_at)
         VALUES('sub-1','org-1','dad',?,'aGVsbG8td29ybGQ','YXV0aHNlY3JldA','Phone',?)`,
      )
      .bind(endpoint, NOW)
      .run()
    // Another user cannot revoke dad's subscription.
    await worker.fetch(
      authed('/api/push/unsubscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ endpoint }),
      }),
      envFor('mgr-1'),
    )
    let row = await db
      .prepare('SELECT revoked_at FROM push_subscriptions WHERE id=?')
      .bind('sub-1')
      .first<{ revoked_at: string | null }>()
    expect(row?.revoked_at).toBeNull()

    const res = await worker.fetch(
      authed('/api/push/unsubscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ endpoint }),
      }),
      envFor('dad'),
    )
    expect(res.status).toBe(200)
    row = await db
      .prepare('SELECT revoked_at FROM push_subscriptions WHERE id=?')
      .bind('sub-1')
      .first<{ revoked_at: string | null }>()
    expect(row?.revoked_at).not.toBeNull()
  })

  it('rejects unauthenticated push calls', async () => {
    const res = await worker.fetch(
      authed('/api/push/subscribe', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          endpoint: 'https://push.example.invalid/x',
          keys: { p256dh: 'aGVsbG8td29ybGQ', auth: 'YXV0aHNlY3JldA' },
        }),
      }),
      { DB: db } as unknown as Parameters<typeof worker.fetch>[1],
    )
    expect(res.status).toBe(401)
  })
})

describe('gated push test endpoint', () => {
  let sqlite: DatabaseSync
  let db: D1Database

  async function seed() {
    await db
      .prepare(`INSERT OR IGNORE INTO organizations(id,name,created_at,updated_at) VALUES('org-1','Test Org',?,?)`)
      .bind(NOW, NOW)
      .run()
    await db
      .prepare(
        `INSERT INTO users(id,organization_id,name,email,phone,role,active,sms_enabled,timezone,created_at,updated_at,recipient_kind,complaint_notifications_enabled)
         VALUES('dad','org-1','Dad','dad@example.invalid','+19086251648','OWNER',1,1,'America/New_York',?,?, 'STANDARD',1)`,
      )
      .bind(NOW, NOW)
      .run()
    await db
      .prepare(
        `INSERT INTO users(id,organization_id,name,email,phone,role,active,sms_enabled,timezone,created_at,updated_at,recipient_kind,complaint_notifications_enabled)
         VALUES('mgr-1','org-1','Manager','mgr@example.invalid','+10000000001','STORE_MANAGER',1,1,'America/New_York',?,?,'STANDARD',1)`,
      )
      .bind(NOW, NOW)
      .run()
  }

  function envFor(userId: string) {
    return {
      DB: db,
      DEV_AUTH_USER_ID: userId,
      VAPID_PUBLIC_KEY: 'test-public-key',
      VAPID_PRIVATE_KEY: 'test-private-key',
    } as unknown as Parameters<typeof worker.fetch>[1]
  }

  function testCall(userId: string) {
    return worker.fetch(
      new Request('https://app.test/api/admin/push/test', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ recipientUserId: 'dad', confirmed: true }),
      }),
      envFor(userId),
    )
  }

  beforeEach(async () => {
    sqlite = new DatabaseSync(':memory:')
    db = new SqliteD1(sqlite)
    for (const name of [
      '0001_initial.sql',
      '0002_persistence_and_auth.sql',
      '0003_pilot_admin_recipient.sql',
      '0004_v1_operations.sql',
      '0005_microsoft_graph_provider.sql',
      '0006_intelligent_mail_foundation.sql',
      '0008_push_subscriptions.sql',
    ]) {
      const path = fileURLToPath(new URL(`../drizzle/${name}`, import.meta.url))
      await db.exec(await readFile(path, 'utf8'))
    }
    await seed()
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('requires admin and honors the kill switch (default off)', async () => {
    const forbidden = await testCall('mgr-1')
    expect(forbidden.status).toBe(403)

    const killed = await testCall('dad')
    expect(killed.status).toBe(409)
    expect(await killed.json()).toMatchObject({ error: expect.stringMatching(/kill switch/i) })

    // Nothing was recorded and no send attempted.
    const rows = await db
      .prepare(`SELECT COUNT(*) c FROM notifications WHERE event_type='TEST' AND channel='PUSH'`)
      .bind()
      .first<{ c: number }>()
    expect(rows?.c).toBe(0)
  })

  it('records FAILED with NO_PUSH_SUBSCRIPTION when enabled but unsubscribed', async () => {
    await db
      .prepare(`INSERT INTO settings(key,value,updated_at) VALUES('push_test_sends_enabled','true',?)`)
      .bind(NOW)
      .run()
    const res = await testCall('dad')
    expect(res.status).toBe(200)
    const row = await db
      .prepare(
        `SELECT status, failure_reason, channel, provider FROM notifications WHERE event_type='TEST' AND channel='PUSH' ORDER BY created_at DESC LIMIT 1`,
      )
      .bind()
      .first<{ status: string; failure_reason: string; channel: string; provider: string }>()
    expect(row).toMatchObject({ status: 'FAILED', failure_reason: 'NO_PUSH_SUBSCRIPTION', channel: 'PUSH', provider: 'WEBPUSH' })
  })

  it('revokes the subscription on 410 and records the failure', async () => {
    await db
      .prepare(`INSERT INTO settings(key,value,updated_at) VALUES('push_test_sends_enabled','true',?)`)
      .bind(NOW)
      .run()
    const clientKeys = await crypto.subtle.generateKey(
      { name: 'ECDH', namedCurve: 'P-256' },
      true,
      ['deriveBits'],
    )
    const clientPublicRaw = new Uint8Array(await crypto.subtle.exportKey('raw', clientKeys.publicKey))
    const endpoint = 'https://push.example.invalid/sub/dad-9'
    await db
      .prepare(
        `INSERT INTO push_subscriptions(id,organization_id,user_id,endpoint,p256dh_key,auth_key,created_at)
         VALUES('sub-9','org-1','dad',?,?,?,?)`,
      )
      .bind(endpoint, b64urlEncode(clientPublicRaw), b64urlEncode(crypto.getRandomValues(new Uint8Array(16))), NOW)
      .run()

    // Real VAPID keys so encryption succeeds and the 410 path is exercised.
    const vapidKeys = await crypto.subtle.generateKey(
      { name: 'ECDSA', namedCurve: 'P-256' },
      true,
      ['sign', 'verify'],
    )
    const vapidJwk = await crypto.subtle.exportKey('jwk', vapidKeys.privateKey)
    const vapidPublicRaw = new Uint8Array([
      0x04,
      ...b64urlDecode(vapidJwk.x!),
      ...b64urlDecode(vapidJwk.y!),
    ])
    const withKeys = {
      DB: db,
      DEV_AUTH_USER_ID: 'dad',
      VAPID_PUBLIC_KEY: b64urlEncode(vapidPublicRaw),
      VAPID_PRIVATE_KEY: vapidJwk.d!,
    } as unknown as Parameters<typeof worker.fetch>[1]
    vi.stubGlobal('fetch', vi.fn(async () => new Response(null, { status: 410 })))
    const res = await worker.fetch(
      new Request('https://app.test/api/admin/push/test', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ recipientUserId: 'dad', confirmed: true }),
      }),
      withKeys,
    )
    expect(res.status).toBe(200)
    const sub = await db
      .prepare('SELECT revoked_at FROM push_subscriptions WHERE id=?')
      .bind('sub-9')
      .first<{ revoked_at: string | null }>()
    expect(sub?.revoked_at).not.toBeNull()
    const note = await db
      .prepare(
        `SELECT status, failure_reason FROM notifications WHERE event_type='TEST' AND channel='PUSH' ORDER BY created_at DESC LIMIT 1`,
      )
      .bind()
      .first<{ status: string; failure_reason: string }>()
    expect(note).toMatchObject({ status: 'FAILED', failure_reason: 'SUBSCRIPTION_EXPIRED' })
  })
})
