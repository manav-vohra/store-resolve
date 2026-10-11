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

async function hkdfExpand(prk: Uint8Array, info: Uint8Array, lengthBytes: number) {
  const key = await crypto.subtle.importKey('raw', prk, 'HKDF', false, ['deriveBits'])
  return new Uint8Array(
    await crypto.subtle.deriveBits(
      { name: 'HKDF', hash: 'SHA-256', salt: new Uint8Array(0), info },
      key,
      lengthBytes * 8,
    ),
  )
}

/** Independent client-side decryption of an RFC 8291 aes128gcm body. */
async function clientDecrypt(
  body: Uint8Array,
  clientPrivateKey: CryptoKey,
  clientPublicRaw: Uint8Array,
  authSecret: Uint8Array,
): Promise<string> {
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
  const hmacKey = await crypto.subtle.importKey(
    'raw',
    authSecret,
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign'],
  )
  const prk = new Uint8Array(await crypto.subtle.sign('HMAC', hmacKey, ecdhSecret))
  const keyInfo = new Uint8Array(
    te.encode('WebPush: info').length + 1 + clientPublicRaw.length + serverPublicRaw.length,
  )
  keyInfo.set(te.encode('WebPush: info'), 0)
  keyInfo[te.encode('WebPush: info').length] = 0x00
  keyInfo.set(clientPublicRaw, te.encode('WebPush: info').length + 1)
  keyInfo.set(serverPublicRaw, te.encode('WebPush: info').length + 1 + clientPublicRaw.length)
  const ikm = await hkdfExpand(prk, keyInfo, 32)
  const cek = await hkdfExpand(ikm, te.encode('Content-Encoding: aes128gcm\0'), 16)
  const nonce = await hkdfExpand(ikm, te.encode('Content-Encoding: nonce\0'), 12)
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
    expect(parseComplaintHash('#/complaints/SR-2026-0002')).toBe('SR-2026-0002')
    expect(parseComplaintHash('#/complaints/SR-2026-0002/')).toBe('SR-2026-0002')
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
