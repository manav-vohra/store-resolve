// Web Push sender for Cloudflare Workers (RFC 8291 content encryption,
// RFC 8292 VAPID authentication).
//
// Library evaluation (2026-10-11, Stage A correction #4): no maintained
// Workers-native Web Push library exists with meaningful adoption. The
// standard `web-push` npm package is Node-only (node:https / node:crypto and a
// Node HTTP stack — incompatible with the Workers runtime). This module
// therefore implements the RFC constructions directly on the Workers
// SubtleCrypto platform primitives (ECDSA P-256, ECDH, HKDF, AES-GCM). That is
// not custom cryptography: the primitives are the platform's audited WebCrypto
// implementation; only the RFC framing is hand-written. No new dependencies.

export interface WebPushSubscription {
  endpoint: string
  p256dh: string // base64url, client public key (uncompressed point)
  auth: string // base64url, client auth secret
}

export interface WebPushPayload {
  title: string
  body: string
  /** Same-origin path, e.g. "/#/complaints/SR-2026-0002". */
  url: string
  tag?: string
  badgeCount?: number
}

export interface VapidConfig {
  publicKey: string // base64url, uncompressed P-256 point
  privateKey: string // base64url, 32-byte scalar
  subject: string // mailto: or https: contact for the JWT `sub` claim
}

const te = new TextEncoder()

export function b64urlEncode(bytes: Uint8Array): string {
  let binary = ''
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i])
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

export function b64urlDecode(value: string): Uint8Array {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/')
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4)
  const binary = atob(padded)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

/** Split an uncompressed P-256 point into x/y base64url coordinates. */
function pointToJwkXY(publicKeyB64url: string): { x: string; y: string } {
  const raw = b64urlDecode(publicKeyB64url)
  if (raw.length !== 65 || raw[0] !== 0x04) throw new Error('Invalid VAPID public key')
  return { x: b64urlEncode(raw.slice(1, 33)), y: b64urlEncode(raw.slice(33, 65)) }
}

/** Convert a DER-encoded ECDSA signature to raw r||s (64 bytes). */
function derToRaw(der: Uint8Array): Uint8Array {
  let offset = 0
  const expect = (byte: number) => {
    if (der[offset++] !== byte) throw new Error('Invalid ECDSA signature encoding')
  }
  const readLength = () => {
    const first = der[offset++]
    if (first < 0x80) return first
    throw new Error('Unexpected long-form length in ECDSA signature')
  }
  const readInteger = () => {
    expect(0x02)
    const length = readLength()
    let bytes = der.slice(offset, offset + length)
    offset += length
    while (bytes.length > 32 && bytes[0] === 0x00) bytes = bytes.slice(1)
    const out = new Uint8Array(32)
    out.set(bytes, 32 - bytes.length)
    return out
  }
  expect(0x30)
  readLength()
  const r = readInteger()
  const s = readInteger()
  const raw = new Uint8Array(64)
  raw.set(r, 0)
  raw.set(s, 32)
  return raw
}

async function vapidJwt(vapid: VapidConfig, audience: string): Promise<string> {
  const { x, y } = pointToJwkXY(vapid.publicKey)
  const key = await crypto.subtle.importKey(
    'jwk',
    { kty: 'EC', crv: 'P-256', x, y, d: vapid.privateKey },
    { name: 'ECDSA', namedCurve: 'P-256' },
    false,
    ['sign'],
  )
  const header = b64urlEncode(te.encode(JSON.stringify({ typ: 'JWT', alg: 'ES256' })))
  const payload = b64urlEncode(
    te.encode(
      JSON.stringify({
        aud: audience,
        exp: Math.floor(Date.now() / 1000) + 12 * 60 * 60,
        sub: vapid.subject,
      }),
    ),
  )
  const signingInput = te.encode(`${header}.${payload}`)
  const sig = new Uint8Array(
    await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, signingInput),
  )
  // WebCrypto returns ECDSA signatures as raw r||s (64 bytes); accept a
  // DER encoding as a fallback for non-conforming runtimes.
  const raw = sig.length === 64 ? sig : derToRaw(sig)
  return `${header}.${payload}.${b64urlEncode(raw)}`
}

async function hmacSha256(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const cryptoKey = await crypto.subtle.importKey(
    'raw',
    key,
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign'],
  )
  return new Uint8Array(await crypto.subtle.sign('HMAC', cryptoKey, data))
}

function concatBytes(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

export interface DerivedPushKeys {
  ecdhSecret: Uint8Array
  prkKey: Uint8Array
  ikm: Uint8Array
  prk: Uint8Array
  cek: Uint8Array
  nonce: Uint8Array
  serverPublicRaw: Uint8Array
}

/**
 * RFC 8291 §3.4 key derivation, mirroring the RFC pseudocode step for step:
 *
 *   PRK_key = HMAC-SHA-256(auth_secret, ecdh_secret)
 *   IKM     = HMAC-SHA-256(PRK_key, key_info || 0x01)          [HKDF-Expand, 32]
 *   PRK     = HMAC-SHA-256(salt, IKM)                          [HKDF-Extract]
 *   CEK     = HMAC-SHA-256(PRK, cek_info || 0x01)[0..15]       [HKDF-Expand, 16]
 *   NONCE   = HMAC-SHA-256(PRK, nonce_info || 0x01)[0..11]     [HKDF-Expand, 12]
 *
 * The per-message `salt` is mixed in at the PRK step — it MUST be the same
 * salt written into the message header. Exported so tests can verify the
 * intermediate values against RFC 8291 Appendix A.
 */
export async function derivePushKeys(
  clientPublicRaw: Uint8Array,
  authSecret: Uint8Array,
  serverPrivateKey: CryptoKey,
  serverPublicRaw: Uint8Array,
  salt: Uint8Array,
): Promise<DerivedPushKeys> {
  if (salt.length !== 16) throw new Error('Push salt must be 16 bytes')
  const clientKey = await crypto.subtle.importKey(
    'raw',
    clientPublicRaw,
    { name: 'ECDH', namedCurve: 'P-256' },
    true,
    [],
  )
  const ecdhSecret = new Uint8Array(
    await crypto.subtle.deriveBits({ name: 'ECDH', public: clientKey }, serverPrivateKey, 256),
  )

  const prkKey = await hmacSha256(authSecret, ecdhSecret)

  const keyInfo = concatBytes([
    te.encode('WebPush: info'),
    new Uint8Array([0x00]),
    clientPublicRaw,
    serverPublicRaw,
  ])
  const ikm = (await hmacSha256(prkKey, concatBytes([keyInfo, new Uint8Array([0x01])]))).slice(0, 32)

  const prk = await hmacSha256(salt, ikm)

  const cekInfo = te.encode('Content-Encoding: aes128gcm\0')
  const cek = (await hmacSha256(prk, concatBytes([cekInfo, new Uint8Array([0x01])]))).slice(0, 16)

  const nonceInfo = te.encode('Content-Encoding: nonce\0')
  const nonce = (await hmacSha256(prk, concatBytes([nonceInfo, new Uint8Array([0x01])]))).slice(0, 12)

  return { ecdhSecret, prkKey, ikm, prk, cek, nonce, serverPublicRaw }
}

interface EncryptedBody {
  bytes: Uint8Array
}

export interface EncryptOverrides {
  /** Fixed 16-byte salt (tests only; random per message in production). */
  salt?: Uint8Array
  /** Fixed server ECDH keypair (tests only; ephemeral per message in production). */
  serverKeys?: CryptoKeyPair
}

/**
 * RFC 8291 §3 aes128gcm encryption of the payload for one subscription.
 * Exported for verification against the RFC 8291 Appendix A test vector.
 */
export async function encryptPayload(
  sub: WebPushSubscription,
  plaintext: string,
  overrides: EncryptOverrides = {},
): Promise<EncryptedBody> {
  const clientPublicRaw = b64urlDecode(sub.p256dh)
  if (clientPublicRaw.length !== 65 || clientPublicRaw[0] !== 0x04)
    throw new Error('Invalid subscription p256dh key')
  const authSecret = b64urlDecode(sub.auth)

  const serverKeys =
    overrides.serverKeys ??
    (await crypto.subtle.generateKey({ name: 'ECDH', namedCurve: 'P-256' }, true, ['deriveBits']))
  const serverPublicRaw = new Uint8Array(await crypto.subtle.exportKey('raw', serverKeys.publicKey))
  const salt = overrides.salt ?? crypto.getRandomValues(new Uint8Array(16))

  const { cek, nonce } = await derivePushKeys(
    clientPublicRaw,
    authSecret,
    serverKeys.privateKey,
    serverPublicRaw,
    salt,
  )

  // Plaintext record: message || 0x02 delimiter (no extra padding).
  const messageBytes = te.encode(plaintext)
  const record = new Uint8Array(messageBytes.length + 1)
  record.set(messageBytes, 0)
  record[messageBytes.length] = 0x02

  const aesKey = await crypto.subtle.importKey('raw', cek, { name: 'AES-GCM' }, false, ['encrypt'])
  const ciphertext = new Uint8Array(
    await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce }, aesKey, record),
  )

  const rs = 4096
  const out = new Uint8Array(16 + 4 + 1 + serverPublicRaw.length + ciphertext.length)
  out.set(salt, 0)
  out[16] = (rs >>> 24) & 0xff
  out[17] = (rs >>> 16) & 0xff
  out[18] = (rs >>> 8) & 0xff
  out[19] = rs & 0xff
  out[20] = serverPublicRaw.length
  out.set(serverPublicRaw, 21)
  out.set(ciphertext, 21 + serverPublicRaw.length)
  return { bytes: out }
}

export interface PushSendResult {
  /** The push service accepted the message (HTTP 200/201). Not device delivery. */
  accepted: boolean
  /** The subscription is gone (HTTP 404/410) — revoke it. */
  expired: boolean
}

/**
 * Privacy-safe test payload. Title/body carry no customer PII — only the app
 * name and a TEST marker. Deep-link URL is the app root.
 */
export function buildTestPushPayload(badgeCount: number): WebPushPayload {
  return {
    title: 'StoreResolve test',
    body: 'TEST — push notifications are connected to this device. No action needed.',
    url: '/',
    tag: 'storeresolve-test-push',
    badgeCount,
  }
}

/**
 * Send one encrypted push message. Throws on transport/provider errors.
 * A 404/410 resolves to `{expired: true}` instead of throwing so the caller
 * can revoke the subscription.
 */
export async function sendWebPush(
  vapid: VapidConfig,
  sub: WebPushSubscription,
  payload: WebPushPayload,
  ttlSeconds = 86400,
  overrides: EncryptOverrides = {},
): Promise<PushSendResult> {
  const endpoint = new URL(sub.endpoint)
  if (endpoint.protocol !== 'https:') throw new Error('Push endpoint must use https')
  if (!vapid.publicKey || !vapid.privateKey || !vapid.subject) throw new Error('Web Push is not configured')

  const { bytes } = await encryptPayload(sub, JSON.stringify(payload), overrides)

  const jwt = await vapidJwt(vapid, endpoint.origin)
  const response = await fetch(sub.endpoint, {
    method: 'POST',
    headers: {
      'Content-Encoding': 'aes128gcm',
      TTL: String(ttlSeconds),
      'Content-Type': 'application/octet-stream',
      Authorization: `vapid t=${jwt}, k=${vapid.publicKey}`,
    },
    body: bytes,
  })
  if (response.status === 404 || response.status === 410) return { accepted: false, expired: true }
  if (response.status === 200 || response.status === 201) return { accepted: true, expired: false }
  const detail = await response.text().catch(() => '')
  throw new Error(
    `Push service rejected the message (${response.status})${detail ? `: ${detail.slice(0, 160)}` : ''}`,
  )
}
