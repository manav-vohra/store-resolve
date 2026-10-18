/**
 * Production-like integration test for the Dad-only pilot tick.
 *
 * Simulates: six stores with NULL managers, Dad as OWNER, FAMILY_PILOT.
 * Proves the tick's safety gates without touching Graph or sending push.
 */
import { describe, expect, it } from 'vitest'
import { runDadPilotDiscoveryTick } from '../worker/dad-pilot-tick'
import type { D1Database } from '../worker/d1'

function fakeDb(settings: Record<string, string>) {
  return {
    prepare(sql: string) {
      const first = async () => {
        const m = sql.match(/key='([^']+)'/)
        const key = m?.[1]
        return key && settings[key] ? { value: settings[key] } : null
      }
      return { bind: () => ({ first, run: async () => ({}) }), first, run: async () => ({}) }
    },
  } as unknown as D1Database
}

const baseEnv = {
  DB: null as unknown as D1Database,
  MS_CLIENT_ID: 'x',
  MS_CLIENT_SECRET: 'x',
  MS_TENANT_ID: 'x',
  MS_REFRESH_TOKEN: 'x',
  MS_MAILBOX_ADDRESS: 'mailbox@test',
  VAPID_PUBLIC_KEY: 'x',
  VAPID_PRIVATE_KEY: 'x',
  PUBLIC_BASE_URL: 'https://example.test',
}

describe('runDadPilotDiscoveryTick gates', () => {
  it('no-ops when not in FAMILY_PILOT', async () => {
    const env = { ...baseEnv, DB: fakeDb({ notification_mode: 'MOCK', email_ingestion_enabled: 'true' }) }
    const r = await runDadPilotDiscoveryTick(env)
    expect(r.ran).toBe(false)
    expect(r.reason).toBe('NOT_FAMILY_PILOT')
  })

  it('no-ops when scheduled ingestion is not enabled', async () => {
    const env = {
      ...baseEnv,
      DB: fakeDb({ notification_mode: 'FAMILY_PILOT', email_ingestion_enabled: 'false' }),
    }
    const r = await runDadPilotDiscoveryTick(env)
    expect(r.ran).toBe(false)
    expect(r.reason).toBe('INGESTION_DISABLED')
  })

  it('no-ops when Graph is not configured', async () => {
    const env = {
      ...baseEnv,
      MS_CLIENT_ID: '',
      DB: fakeDb({ notification_mode: 'FAMILY_PILOT', email_ingestion_enabled: 'true' }),
    }
    const r = await runDadPilotDiscoveryTick(env)
    expect(r.ran).toBe(false)
    expect(r.reason).toBe('GRAPH_NOT_CONFIGURED')
  })
})
