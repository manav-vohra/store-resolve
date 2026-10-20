import { createPublicKey, verify } from 'node:crypto'
import { spawn } from 'node:child_process'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { createInterface } from 'node:readline/promises'

export const authority = 'https://login.microsoftonline.com/consumers'
export const scopes = 'openid profile email offline_access Mail.Read Mail.Send User.Read'
export const oidcConfigurationUrl = `${authority}/v2.0/.well-known/openid-configuration`
export const graphReadinessUrl =
  'https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=id,displayName,totalItemCount,unreadItemCount'
export const graphSanityUrl =
  'https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime%20ge%202026-08-01T00%3A00%3A00Z%20and%20receivedDateTime%20lt%202026-08-02T00%3A00%3A00Z&$select=id,receivedDateTime,from&$top=3'

const formHeaders = { 'content-type': 'application/x-www-form-urlencoded' }
const sleep = (milliseconds) =>
  new Promise((resolveSleep) => setTimeout(resolveSleep, milliseconds))
const allowedClockSkewSeconds = 60

const safeJson = async (response) => {
  const value = await response.json().catch(() => null)
  return value && typeof value === 'object' && !Array.isArray(value) ? value : {}
}

const decodeJsonSegment = (segment, label) => {
  try {
    const value = JSON.parse(Buffer.from(segment, 'base64url').toString('utf8'))
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error()
    return value
  } catch {
    throw new Error(`Microsoft ID token has an invalid ${label}.`)
  }
}

const requireHttpsMicrosoftUrl = (value, label) => {
  if (typeof value !== 'string') throw new Error(`Microsoft OIDC ${label} is missing.`)
  const url = new URL(value)
  if (url.protocol !== 'https:' || url.hostname !== 'login.microsoftonline.com')
    throw new Error(`Microsoft OIDC ${label} is not trusted.`)
  return url.href
}

const signingKeyIssuerMatches = (keyIssuer, tokenIssuer, tenantId) => {
  if (typeof keyIssuer !== 'string') return true
  if (keyIssuer.includes('{tenantid}')) {
    if (typeof tenantId !== 'string' || !tenantId) return false
    return keyIssuer.replace('{tenantid}', tenantId) === tokenIssuer
  }
  return keyIssuer === tokenIssuer
}

export async function requestDeviceCode(clientId, fetchImpl = fetch) {
  const response = await fetchImpl(`${authority}/oauth2/v2.0/devicecode`, {
    method: 'POST',
    headers: formHeaders,
    body: new URLSearchParams({ client_id: clientId, scope: scopes }),
  })
  const body = await safeJson(response)
  if (
    !response.ok ||
    typeof body.device_code !== 'string' ||
    typeof body.user_code !== 'string' ||
    typeof body.verification_uri !== 'string'
  ) {
    throw new Error(
      'Microsoft device authorization failed. Confirm Authentication > Allow public client flows is Yes.',
    )
  }
  return {
    deviceCode: body.device_code,
    userCode: body.user_code,
    verificationUri: body.verification_uri,
    expiresIn: typeof body.expires_in === 'number' ? body.expires_in : 900,
    interval: typeof body.interval === 'number' ? Math.max(body.interval, 1) : 5,
  }
}

export async function pollForTokenResponse(clientId, device, fetchImpl = fetch, wait = sleep) {
  const deadline = Date.now() + device.expiresIn * 1000
  let intervalSeconds = device.interval
  while (Date.now() < deadline) {
    await wait(intervalSeconds * 1000)
    const response = await fetchImpl(`${authority}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: formHeaders,
      body: new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        client_id: clientId,
        device_code: device.deviceCode,
      }),
    })
    const body = await safeJson(response)
    if (response.ok) {
      if (
        typeof body.access_token !== 'string' ||
        typeof body.refresh_token !== 'string' ||
        typeof body.id_token !== 'string'
      )
        throw new Error('Microsoft authorization did not return the required identity credentials.')
      return {
        accessToken: body.access_token,
        refreshToken: body.refresh_token,
        idToken: body.id_token,
      }
    }
    if (body.error === 'authorization_pending') continue
    if (body.error === 'slow_down') {
      intervalSeconds += 5
      continue
    }
    if (body.error === 'authorization_declined')
      throw new Error('Microsoft authorization was declined.')
    if (body.error === 'expired_token') throw new Error('The Microsoft device code expired.')
    throw new Error('Microsoft device authorization failed during token polling.')
  }
  throw new Error('The Microsoft device code expired.')
}

export async function validateIdToken(
  idToken,
  clientId,
  fetchImpl = fetch,
  nowSeconds = Math.floor(Date.now() / 1000),
) {
  const segments = idToken.split('.')
  if (segments.length !== 3) throw new Error('Microsoft ID token is malformed.')
  const [encodedHeader, encodedPayload, encodedSignature] = segments
  const header = decodeJsonSegment(encodedHeader, 'header')
  const claims = decodeJsonSegment(encodedPayload, 'payload')
  if (header.alg !== 'RS256' || typeof header.kid !== 'string')
    throw new Error('Microsoft ID token uses an unsupported signing key.')

  const metadataResponse = await fetchImpl(oidcConfigurationUrl)
  const metadata = await safeJson(metadataResponse)
  if (!metadataResponse.ok) throw new Error('Microsoft OIDC discovery failed.')
  const issuer = requireHttpsMicrosoftUrl(metadata.issuer, 'issuer').replace(/\/$/, '')
  const jwksUri = requireHttpsMicrosoftUrl(metadata.jwks_uri, 'signing-key URL')

  const keysResponse = await fetchImpl(jwksUri)
  const keysBody = await safeJson(keysResponse)
  if (!keysResponse.ok || !Array.isArray(keysBody.keys))
    throw new Error('Microsoft signing-key retrieval failed.')
  const signingKey = keysBody.keys.find(
    (key) =>
      key &&
      typeof key === 'object' &&
      key.kid === header.kid &&
      key.kty === 'RSA' &&
      (key.use === undefined || key.use === 'sig') &&
      (key.alg === undefined || key.alg === 'RS256'),
  )
  if (!signingKey) throw new Error('Microsoft ID-token signing key was not found.')

  let signatureValid = false
  try {
    const publicKey = createPublicKey({ key: signingKey, format: 'jwk' })
    signatureValid = verify(
      'RSA-SHA256',
      Buffer.from(`${encodedHeader}.${encodedPayload}`),
      publicKey,
      Buffer.from(encodedSignature, 'base64url'),
    )
  } catch {
    signatureValid = false
  }
  if (!signatureValid) throw new Error('Microsoft ID-token signature is invalid.')

  if (claims.iss !== issuer) throw new Error('Microsoft ID-token issuer is invalid.')
  if (!signingKeyIssuerMatches(signingKey.issuer, claims.iss, claims.tid))
    throw new Error('Microsoft signing-key issuer is invalid.')
  const audiences = Array.isArray(claims.aud) ? claims.aud : [claims.aud]
  if (!audiences.includes(clientId)) throw new Error('Microsoft ID-token audience is invalid.')
  if (typeof claims.exp !== 'number' || claims.exp <= nowSeconds - allowedClockSkewSeconds)
    throw new Error('Microsoft ID token is expired.')
  if (typeof claims.nbf === 'number' && claims.nbf > nowSeconds + allowedClockSkewSeconds)
    throw new Error('Microsoft ID token is not active yet.')
  if (typeof claims.iat === 'number' && claims.iat > nowSeconds + allowedClockSkewSeconds)
    throw new Error('Microsoft ID token has an invalid issued-at time.')

  // Device Code Flow has no caller-supplied state or nonce parameter to correlate.
  return claims
}

export function maskMailbox(mailbox) {
  const separator = mailbox.lastIndexOf('@')
  if (separator <= 0) return '***'
  const local = mailbox.slice(0, separator)
  const domain = mailbox.slice(separator + 1)
  const visible = local.length > 1 ? `${local[0]}***${local.at(-1)}` : `${local[0]}***`
  return `${visible}@${domain}`
}

export function confirmExpectedMailbox(claims, expectedMailbox) {
  const expected = expectedMailbox.toLowerCase()
  const identifiers = [claims.email, claims.preferred_username].filter(
    (value) => typeof value === 'string' && value.length > 0,
  )
  if (identifiers.length === 0)
    return {
      matches: false,
      maskedIdentity: '***',
      reason: 'No mailbox identity claim was returned.',
    }
  const normalized = identifiers.map((value) => value.toLowerCase())
  const matchingIdentity = identifiers[normalized.indexOf(expected)]
  const matches = normalized.every((value) => value === expected)
  return {
    matches,
    maskedIdentity: maskMailbox(matchingIdentity ?? identifiers[0]),
    reason: matches
      ? undefined
      : 'The authorized Microsoft account does not match MS_MAILBOX_ADDRESS.',
  }
}

const graphHeaders = (accessToken) => ({ authorization: `Bearer ${accessToken}` })

export const graphIdentityUrl =
  'https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName'

export async function verifyGraphIdentity(accessToken, expectedMailbox, fetchImpl = fetch) {
  const response = await fetchImpl(graphIdentityUrl, { headers: graphHeaders(accessToken) })
  if (!response.ok)
    throw new Error('Microsoft Graph /me identity verification failed (User.Read consent may be missing).')
  const profile = await safeJson(response)
  const identifiers = [profile.mail, profile.userPrincipalName].filter(
    (value) => typeof value === 'string' && value.length > 0,
  )
  if (identifiers.length === 0)
    throw new Error('Microsoft Graph /me returned no usable mailbox identity.')
  const expected = expectedMailbox.toLowerCase()
  const normalized = identifiers.map((value) => value.toLowerCase())
  if (!normalized.includes(expected))
    throw new Error('Microsoft Graph /me identity does not match MS_MAILBOX_ADDRESS.')
  return { maskedIdentity: maskMailbox(identifiers[0]), matches: true }
}

export async function verifyGraphReadiness(accessToken, fetchImpl = fetch) {
  const response = await fetchImpl(graphReadinessUrl, { headers: graphHeaders(accessToken) })
  if (!response.ok) throw new Error('Microsoft Graph readiness verification failed.')
  const metadata = await safeJson(response)
  if (typeof metadata.id !== 'string')
    throw new Error('Microsoft Graph readiness metadata was incomplete.')
  return true
}

export async function runMailboxMetadataSanityCheck(accessToken, fetchImpl = fetch) {
  const response = await fetchImpl(graphSanityUrl, { headers: graphHeaders(accessToken) })
  if (!response.ok) throw new Error('Microsoft Graph mailbox metadata sanity check failed.')
  const body = await safeJson(response)
  if (!Array.isArray(body.value) || body.value.length > 3)
    throw new Error('Microsoft Graph mailbox metadata sanity response was invalid.')
  return body.value.length
}

export async function storeRefreshToken(refreshToken, cwd = process.cwd()) {
  const wranglerPath = resolve(cwd, 'node_modules', 'wrangler', 'bin', 'wrangler.js')
  await new Promise((resolveStore, reject) => {
    const child = spawn(process.execPath, [wranglerPath, 'secret', 'put', 'MS_REFRESH_TOKEN'], {
      cwd,
      stdio: ['pipe', 'inherit', 'inherit'],
      shell: false,
    })
    child.stdin.end(`${refreshToken}\n`)
    child.on('error', reject)
    child.on('exit', (code) =>
      code === 0 ? resolveStore() : reject(new Error('Wrangler could not store MS_REFRESH_TOKEN.')),
    )
  })
}

export async function authorizeExpectedMailbox({
  clientId,
  expectedMailbox,
  fetchImpl = fetch,
  wait = sleep,
  storeToken = storeRefreshToken,
  log = console.log,
}) {
  const device = await requestDeviceCode(clientId, fetchImpl)
  log(`MICROSOFT_DEVICE_LOGIN_URL=${device.verificationUri}`)
  log(`MICROSOFT_DEVICE_CODE=${device.userCode}`)
  const tokens = await pollForTokenResponse(clientId, device, fetchImpl, wait)
  const claims = await validateIdToken(tokens.idToken, clientId, fetchImpl)
  const confirmation = confirmExpectedMailbox(claims, expectedMailbox)
  log(`AUTHORIZED_MICROSOFT_IDENTITY=${confirmation.maskedIdentity}`)
  log(`EXPECTED_MAILBOX_MATCH=${confirmation.matches ? 'YES' : 'NO'}`)
  if (!confirmation.matches) throw new Error(confirmation.reason)
  // Verify /me identity (User.Read) AND inbox access BEFORE replacing the
  // production refresh token. Never risk interrupting working ingestion.
  const graphIdentity = await verifyGraphIdentity(tokens.accessToken, expectedMailbox, fetchImpl)
  log(`GRAPH_ME_IDENTITY=${graphIdentity.maskedIdentity}`)
  log(`GRAPH_ME_MATCH=YES`)
  await verifyGraphReadiness(tokens.accessToken, fetchImpl)
  log('MICROSOFT_GRAPH_READY=YES')
  const sanityCount = await runMailboxMetadataSanityCheck(tokens.accessToken, fetchImpl)
  log(`AUGUST_1_METADATA_RECORD_COUNT=${sanityCount}`)
  await storeToken(tokens.refreshToken)
  log('MS_REFRESH_TOKEN was written directly to Cloudflare and was not displayed.')
  return { ...confirmation, graphReady: true, sanityCount }
}

export async function main() {
  let clientId = process.env.MS_CLIENT_ID?.trim()
  let expectedMailbox = process.env.MS_MAILBOX_ADDRESS?.trim()
  const input = createInterface({ input: process.stdin, output: process.stdout })
  if (!clientId) clientId = (await input.question('Microsoft Application client ID: ')).trim()
  if (!expectedMailbox)
    expectedMailbox = (await input.question('Expected Microsoft mailbox address: ')).trim()
  input.close()
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId))
    throw new Error('MS_CLIENT_ID must be the Application (client) ID GUID.')
  if (!/^[^@\s]+@[^@\s]+$/.test(expectedMailbox))
    throw new Error('MS_MAILBOX_ADDRESS must be a valid mailbox address.')
  await authorizeExpectedMailbox({ clientId, expectedMailbox })
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : 'Microsoft device authorization failed.')
    process.exitCode = 1
  })
}
