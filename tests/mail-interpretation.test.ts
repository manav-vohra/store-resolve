import { readFile } from 'node:fs/promises'
import { DatabaseSync, type SQLInputValue, type StatementSync } from 'node:sqlite'
import { fileURLToPath } from 'node:url'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { D1Database, D1PreparedStatement, D1Result } from '../worker/d1'
import type { NormalizedEmailMessage } from '../worker/email-provider'
import { ingestEmailMessage } from '../worker/ingestion'
import {
  DisabledModelTransport,
  FakeMailInterpreter,
  MailInterpretationError,
  OpenAiMailInterpreter,
  buildInterpretationInput,
  detectDisagreements,
  interpretWithConfig,
  interpretationPrompt,
  loadMailInterpretationConfig,
  validateInterpretationOutput,
  type MailInterpretationConfig,
  type MailInterpretationInput,
  type MailInterpretationOutput,
} from '../worker/mail-interpretation'

// --- Minimal fakes -----------------------------------------------------------

class FakeSettingsDb implements D1Database {
  constructor(private readonly settings: Record<string, string>) {}
  prepare() {
    const settings = this.settings
    const statement = {
      bind() {
        return {
          async all<T>(): Promise<D1Result<T>> {
            return {
              success: true,
              results: Object.entries(settings).map(([key, value]) => ({ key, value })) as T[],
            }
          },
          async first<T>(): Promise<T | null> {
            return null
          },
          async run(): Promise<D1Result> {
            return { success: true, results: [] }
          },
        }
      },
    }
    return statement as unknown as D1PreparedStatement
  }
  async batch<T>(): Promise<D1Result<T>[]> {
    throw new Error('not implemented')
  }
  async exec(): Promise<unknown> {
    throw new Error('not implemented')
  }
}

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
      if (value instanceof Uint8Array) {
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
  async batch<T>(statements: D1PreparedStatement[]): Promise<D1Result<T>[]> {
    this.database.exec('BEGIN')
    try {
      const results = []
      for (const statement of statements) results.push((await statement.run()) as D1Result<T>)
      this.database.exec('COMMIT')
      return results
    } catch (error) {
      this.database.exec('ROLLBACK')
      throw error
    }
  }
  async exec(sql: string): Promise<D1Result> {
    this.database.exec(sql)
    return { success: true, results: [] }
  }
}

const message = (overrides: Partial<NormalizedEmailMessage> = {}): NormalizedEmailMessage => ({
  id: 'graph-1',
  threadId: 'thread-1',
  internalDate: '2026-08-14T12:00:00.000Z',
  sender: 'Dunkin Guest Care <guestcare@example.invalid>',
  recipients: 'operations@example.invalid',
  subject: 'Complaint reference DD-1001',
  messageIdHeader: '<graph-1@example.invalid>',
  textBody: 'Complaint Reference ID: DD-1001\nStore Number: 41001\nComplaint: slow service',
  ...overrides,
})

const enabledConfig: MailInterpretationConfig = {
  enabled: true,
  model: 'gpt-4o-mini',
  promptVersion: 'mail-interpret.v1',
  shadowMode: true,
}

const validOutput = (overrides: Partial<MailInterpretationOutput> = {}): MailInterpretationOutput => ({
  schemaVersion: 'mail-interpret.output.v1',
  externalCaseId: 'DD-1001',
  storeNumber: '41001',
  storeConfidence: 0.9,
  issueCategory: 'Service',
  urgency: 'LOW',
  confidence: 0.9,
  summary: 'Guest reports slow service at store 41001, reference DD-1001.',
  evidenceQuotes: ['slow service'],
  ...overrides,
})

const testInput = (overrides: Partial<MailInterpretationInput> = {}): MailInterpretationInput => ({
  schemaVersion: 'mail-interpret.input.v1',
  promptVersion: 'mail-interpret.v1',
  normalized: {
    subject: 'Complaint reference DD-1001',
    senderAddress: 'Dunkin Guest Care <guestcare@example.invalid>',
    receivedAt: '2026-08-14T12:00:00.000Z',
    conversationId: 'thread-1',
    isComplaint: true,
    externalCaseId: 'DD-1001',
    storeNumber: '41001',
    category: 'Service',
    severity: 'LOW',
  },
  deterministicEvidence: {
    extractionRules: ['IS_COMPLAINT_KEYWORD', 'CASE_ID_LINE_ANCHORED', 'STORE_NUMBER_LABELED'],
    storeResolution: { storeNumber: '41001', reason: 'EXACT_STORE_NUMBER' },
  },
  sourceMetadata: {
    provider: 'MICROSOFT_GRAPH',
    mailboxKey: 'production',
    bodyExcerpt: 'Complaint Reference ID: DD-1001\nStore Number: 41001\nComplaint: slow service',
  },
  ...overrides,
})

// --- Config ------------------------------------------------------------------

describe('interpretation config', () => {
  it('defaults to disabled with safe model and prompt version', async () => {
    const config = await loadMailInterpretationConfig(new FakeSettingsDb({}))
    expect(config).toEqual({
      enabled: false,
      model: 'gpt-4o-mini',
      promptVersion: 'mail-interpret.v1',
      shadowMode: true,
    })
  })

  it('reads the enabled flag, model, and prompt version from settings', async () => {
    const config = await loadMailInterpretationConfig(
      new FakeSettingsDb({
        mail_interpretation_enabled: 'true',
        mail_interpretation_model: 'gpt-4o-mini',
        mail_interpretation_prompt_version: 'mail-interpret.v1',
      }),
    )
    expect(config.enabled).toBe(true)
    expect(config.model).toBe('gpt-4o-mini')
    expect(config.promptVersion).toBe('mail-interpret.v1')
  })

  it('treats any non-"true" enabled value as disabled', async () => {
    const config = await loadMailInterpretationConfig(new FakeSettingsDb({ mail_interpretation_enabled: 'yes' }))
    expect(config.enabled).toBe(false)
  })
})

// --- Prompts -----------------------------------------------------------------

describe('prompt registry', () => {
  it('serves the v1 prompt', () => {
    expect(interpretationPrompt('mail-interpret.v1')).toContain('never invent')
  })

  it('rejects unknown versions instead of silently falling back', () => {
    expect(() => interpretationPrompt('mail-interpret.v99')).toThrowError(MailInterpretationError)
  })
})

// --- Validation ---------------------------------------------------------------

describe('output validation', () => {
  it('accepts a well-formed output', () => {
    const result = validateInterpretationOutput(validOutput())
    expect(result.valid).toBe(true)
    expect(result.errors).toEqual([])
    expect(result.output?.externalCaseId).toBe('DD-1001')
  })

  it('rejects bad enums, out-of-range confidence, and storeConfidence mismatches', () => {
    for (const bad of [
      validOutput({ issueCategory: 'Billing' as never }),
      validOutput({ urgency: 'URGENT' as never }),
      validOutput({ confidence: 1.5 }),
      validOutput({ storeNumber: null, storeConfidence: 0.5 }),
      validOutput({ storeNumber: '41001', storeConfidence: 1.5 }),
      validOutput({ summary: 'x'.repeat(501) }),
      validOutput({ evidenceQuotes: ['a', 'b', 'c', 'd', 'e'] }),
      validOutput({ schemaVersion: 'wrong' as never }),
      'not an object',
    ]) {
      expect(validateInterpretationOutput(bad).valid).toBe(false)
    }
  })

  it('accepts null storeConfidence with a storeNumber (surfaced as LOW_CONFIDENCE)', () => {
    const result = validateInterpretationOutput(validOutput({ storeNumber: '41001', storeConfidence: null }))
    expect(result.valid).toBe(true)
    const disagreements = detectDisagreements(testInput(), result.output, [])
    expect(disagreements.map((d) => d.code)).toContain('LOW_CONFIDENCE')
  })

  it('rejects non-JSON-shaped input', () => {
    expect(validateInterpretationOutput(null).valid).toBe(false)
    expect(validateInterpretationOutput([1, 2]).valid).toBe(false)
  })
})

// --- Disagreements -------------------------------------------------------------

describe('disagreement detection', () => {
  it('reports no disagreements when model and deterministic evidence agree', () => {
    expect(detectDisagreements(testInput(), validOutput(), [])).toEqual([])
  })

  it('flags case ID mismatch, keeping the deterministic ID authoritative', () => {
    const disagreements = detectDisagreements(testInput(), validOutput({ externalCaseId: 'DD-9999' }), [])
    expect(disagreements.map((d) => d.code)).toContain('CASE_ID_MISMATCH')
    expect(disagreements.find((d) => d.code === 'CASE_ID_MISMATCH')?.detail).toContain('deterministic wins')
  })

  it('flags a model-only case ID as unverified, never as identity', () => {
    const input = testInput({
      normalized: { ...testInput().normalized, externalCaseId: undefined },
      deterministicEvidence: { extractionRules: [], storeResolution: { reason: 'NO_DETERMINISTIC_STORE_MATCH' } },
    })
    const disagreements = detectDisagreements(input, validOutput({ externalCaseId: 'DD-7777' }), [])
    expect(disagreements.map((d) => d.code)).toContain('CASE_ID_UNVERIFIED')
    expect(disagreements.map((d) => d.code)).not.toContain('CASE_ID_MISMATCH')
  })

  it('flags store mismatch and unverified store suggestions', () => {
    const mismatch = detectDisagreements(
      testInput({ deterministicEvidence: { extractionRules: [], storeResolution: { storeNumber: '350909', reason: 'EXACT_STORE_NUMBER' } } }),
      validOutput({ storeNumber: '41001' }),
      [],
    )
    expect(mismatch.map((d) => d.code)).toContain('STORE_MISMATCH')

    const unverified = detectDisagreements(
      testInput({
        deterministicEvidence: {
          extractionRules: [],
          storeResolution: { reason: 'NO_DETERMINISTIC_STORE_MATCH' },
        },
      }),
      validOutput({ storeNumber: '99999' }),
      [],
    )
    expect(unverified.map((d) => d.code)).toContain('STORE_UNVERIFIED')
    expect(unverified.find((d) => d.code === 'STORE_UNVERIFIED')?.detail).toContain('not invented')
  })

  it('flags severity mismatch and adds an escalation candidate when the model rates higher', () => {
    const disagreements = detectDisagreements(testInput(), validOutput({ urgency: 'CRITICAL' }), [])
    const codes = disagreements.map((d) => d.code)
    expect(codes).toContain('SEVERITY_MISMATCH')
    expect(codes).toContain('SEVERITY_ESCALATION_CANDIDATE')
  })

  it('does not flag an escalation candidate when the model rates lower', () => {
    const input = testInput({ normalized: { ...testInput().normalized, severity: 'HIGH' } })
    const codes = detectDisagreements(input, validOutput({ urgency: 'LOW' }), []).map((d) => d.code)
    expect(codes).toContain('SEVERITY_MISMATCH')
    expect(codes).not.toContain('SEVERITY_ESCALATION_CANDIDATE')
  })

  it('flags low confidence and hallucinated evidence quotes', () => {
    const disagreements = detectDisagreements(
      testInput(),
      validOutput({ confidence: 0.3, evidenceQuotes: ['the moon is made of cheese'] }),
      [],
    )
    const codes = disagreements.map((d) => d.code)
    expect(codes).toContain('LOW_CONFIDENCE')
    expect(codes).toContain('EVIDENCE_CONTRADICTION')
  })

  it('treats a null output as abstained and a validation failure as a schema violation', () => {
    expect(detectDisagreements(testInput(), null, []).map((d) => d.code)).toEqual(['ABSTAINED'])
    const codes = detectDisagreements(testInput(), null, ['confidence must be 0..1']).map((d) => d.code)
    expect(codes).toEqual(['SCHEMA_VIOLATION', 'ABSTAINED'])
  })
})

// --- Interpreter behavior -------------------------------------------------------

describe('interpretWithConfig', () => {
  it('skips entirely when disabled without touching the interpreter', async () => {
    let calls = 0
    const interpreter = new FakeMailInterpreter(() => {
      calls += 1
      return JSON.stringify(validOutput())
    })
    const result = await interpretWithConfig(interpreter, testInput(), {
      ...enabledConfig,
      enabled: false,
    })
    expect(result.skipped).toBe(true)
    expect(result.output).toBeNull()
    expect(calls).toBe(0)
  })

  it('converts live transport failures into abstained results instead of throwing', async () => {
    const failingTransport = {
      async complete() {
        throw new MailInterpretationError('MODEL_TRANSPORT_FAILED', 'live transport failed')
      },
    }
    const interpreter = new OpenAiMailInterpreter(enabledConfig, failingTransport)
    const result = await interpretWithConfig(interpreter, testInput(), enabledConfig)
    expect(result.skipped).toBe(false)
    expect(result.output).toBeNull()
    expect(result.disagreements.map((d) => d.code)).toContain('ABSTAINED')
  })

  it('treats a disabled transport as skipped, not ABSTAINED', async () => {
    const interpreter = new OpenAiMailInterpreter(enabledConfig, new DisabledModelTransport())
    const result = await interpretWithConfig(interpreter, testInput(), enabledConfig)
    expect(result.skipped).toBe(true)
    expect(result.output).toBeNull()
    expect(result.disagreements).toEqual([])
    expect(result.validation.errors).toContain('interpreter disabled - no live transport')
  })

  it('refuses non-complaint candidates without invoking the interpreter', async () => {
    let calls = 0
    const interpreter = new FakeMailInterpreter(() => {
      calls += 1
      return JSON.stringify(validOutput())
    })
    const nonCandidate = testInput({
      normalized: { ...testInput().normalized, isComplaint: false },
    })
    const result = await interpretWithConfig(interpreter, nonCandidate, enabledConfig)
    expect(result.skipped).toBe(true)
    expect(result.output).toBeNull()
    expect(result.validation.errors).toContain('not a deterministic complaint candidate')
    expect(calls).toBe(0)
  })

  it('throws only for unknown prompt versions', async () => {
    const interpreter = new FakeMailInterpreter(() => JSON.stringify(validOutput()))
    await expect(
      interpretWithConfig(interpreter, testInput(), { ...enabledConfig, promptVersion: 'nope' }),
    ).rejects.toThrowError(MailInterpretationError)
  })

  it('processes a valid fake response end to end', async () => {
    const interpreter = new FakeMailInterpreter(() => JSON.stringify(validOutput()))
    const result = await interpretWithConfig(interpreter, testInput(), enabledConfig)
    expect(result.output?.externalCaseId).toBe('DD-1001')
    expect(result.confidence).toBe(0.9)
    expect(result.disagreements).toEqual([])
    expect(result.inputSha256).toMatch(/^[0-9a-f]{64}$/)
    expect(result.outputSha256).toMatch(/^[0-9a-f]{64}$/)
  })

  it('marks non-JSON model output as a schema violation', async () => {
    const interpreter = new FakeMailInterpreter(() => 'definitely not json {{{')
    const result = await interpretWithConfig(interpreter, testInput(), enabledConfig)
    expect(result.output).toBeNull()
    expect(result.disagreements.map((d) => d.code)).toContain('SCHEMA_VIOLATION')
  })
})

describe('OpenAiMailInterpreter', () => {
  it('builds a deterministic JSON-mode request without touching the network', () => {
    const interpreter = new OpenAiMailInterpreter(enabledConfig, new DisabledModelTransport())
    const request = interpreter.buildRequest(testInput())
    expect(request.model).toBe('gpt-4o-mini')
    expect(request.temperature).toBe(0)
    expect(request.response_format).toEqual({ type: 'json_object' })
    expect(request.messages).toHaveLength(2)
    expect(request.messages[0]?.role).toBe('system')
    expect(request.messages[1]?.role).toBe('user')
  })

  it('the shipped disabled transport always throws', async () => {
    const transport = new DisabledModelTransport()
    await expect(transport.complete()).rejects.toThrowError(MailInterpretationError)
  })
})

// --- Integration: wiring through ingestEmailMessage ------------------------------

describe('Phase B intake wiring', () => {
  let sqlite: DatabaseSync
  let db: D1Database

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
    ]) {
      const path = fileURLToPath(new URL(`../drizzle/${name}`, import.meta.url))
      await db.exec(await readFile(path, 'utf8'))
    }
  })

  afterEach(() => sqlite.close())

  const runRow = (sourceId: string) =>
    db
      .prepare(
        `SELECT model_name,prompt_version,ai_structured_output_json,disagreement_flags_json,latency_ms,deterministic_evidence_json FROM mail_processing_runs WHERE source_message_id=?`,
      )
      .bind(sourceId)
      .first<{
        model_name: string | null
        prompt_version: string | null
        ai_structured_output_json: string | null
        disagreement_flags_json: string
        latency_ms: number | null
        deterministic_evidence_json: string | null
      }>()

  it('leaves run rows untouched when the interpreter is disabled (default)', async () => {
    const result = await ingestEmailMessage(db, message())
    expect(result.status).toBe('PROCESSED')
    const run = await runRow('graph-1')
    expect(run?.model_name).toBeNull()
    expect(run?.ai_structured_output_json).toBeNull()
    expect(run?.disagreement_flags_json).toBe('[]')
  })

  it('records the interpretation on the run and keeps the deterministic complaint outcome', async () => {
    const interpreter = new FakeMailInterpreter(() => JSON.stringify(validOutput()))
    const result = await ingestEmailMessage(db, message(), {
      interpreter,
      interpretationConfig: enabledConfig,
    })
    expect(result.status).toBe('PROCESSED')
    expect(result.complaintId).toBeDefined()

    const run = await runRow('graph-1')
    expect(run?.model_name).toBe('gpt-4o-mini')
    expect(run?.prompt_version).toBe('mail-interpret.v1')
    expect(JSON.parse(run?.ai_structured_output_json ?? '{}').externalCaseId).toBe('DD-1001')
    expect(JSON.parse(run?.deterministic_evidence_json ?? '{}').extractionRules).toContain('CASE_ID_LINE_ANCHORED')

    // Deterministic identity wins: the complaint carries the extracted case ID,
    // exactly one complaint exists, and nothing was sent anywhere.
    const complaints = await db.prepare('SELECT external_case_id FROM complaints').all<{ external_case_id: string }>()
    expect(complaints.results).toEqual([{ external_case_id: 'DD-1001' }])
    for (const table of ['outbound_deliveries', 'email_acknowledgments', 'response_actions']) {
      const count = await db.prepare(`SELECT COUNT(*) AS count FROM ${table}`).first<{ count: number }>()
      expect(count?.count).toBe(0)
    }
  })

  it('routes a model-only case ID to provisional review with the suggestion as an unverified hint', async () => {
    const interpreter = new FakeMailInterpreter(() =>
      JSON.stringify(
        validOutput({ externalCaseId: 'DD-7777', storeNumber: null, storeConfidence: null, confidence: 0.72 }),
      ),
    )
    const result = await ingestEmailMessage(
      db,
      message({
        id: 'noid-2',
        threadId: 'noid-thread-2',
        subject: 'Guest concern at an unknown store',
        textBody: 'Customer complaint with no location and no reference number.',
      }),
      { interpreter, interpretationConfig: enabledConfig },
    )
    expect(result.status).toBe('REVIEW_REQUIRED')
    expect(result.complaintId).toBeUndefined()

    // Identity rule: no complaint, no fabricated business ID.
    const complaints = await db.prepare('SELECT COUNT(*) AS count FROM complaints').first<{ count: number }>()
    expect(complaints?.count).toBe(0)

    const reviews = await db
      .prepare('SELECT reason_code,disagreement_flags_json FROM mail_review_items')
      .all<{ reason_code: string; disagreement_flags_json: string }>()
    expect(reviews.results).toHaveLength(1)
    expect(reviews.results[0]?.reason_code).toBe('IDENTITY_UNRESOLVED')
    expect(JSON.parse(reviews.results[0]?.disagreement_flags_json ?? '[]')).toContain('CASE_ID_UNVERIFIED')

    const events = await db
      .prepare('SELECT external_case_id,payload_json FROM canonical_mail_events')
      .all<{ external_case_id: string | null; payload_json: string }>()
    expect(events.results).toHaveLength(1)
    expect(events.results[0]?.external_case_id).toBeNull()
    const payload = JSON.parse(events.results[0]?.payload_json ?? '{}')
    expect(payload.modelSuggestedCaseId).toBe('DD-7777')
    expect(payload.modelSuggestedCaseIdNote).toBe('UNVERIFIED_MODEL_SUGGESTION_NOT_IDENTITY')
  })

  it('a model suggestion never overrides a deterministic new case ID', async () => {
    const interpreter = new FakeMailInterpreter(() =>
      JSON.stringify(validOutput({ externalCaseId: 'DD-9999', confidence: 0.95 })),
    )
    const result = await ingestEmailMessage(db, message(), {
      interpreter,
      interpretationConfig: enabledConfig,
    })
    expect(result.status).toBe('PROCESSED')
    const complaints = await db.prepare('SELECT external_case_id FROM complaints').all<{ external_case_id: string }>()
    expect(complaints.results).toEqual([{ external_case_id: 'DD-1001' }])
    const run = await runRow('graph-1')
    expect(JSON.parse(run?.disagreement_flags_json ?? '[]')).toContain('CASE_ID_MISMATCH')
  })

  it('skips interpretation (no record) and still completes intake when the transport is disabled but the flag is on', async () => {
    const result = await ingestEmailMessage(db, message(), {
      interpretationConfig: enabledConfig,
      // default interpreter -> OpenAiMailInterpreter + DisabledModelTransport: must never throw
    })
    expect(result.status).toBe('PROCESSED')
    const run = await runRow('graph-1')
    // Disabled transport = skipped: no interpretation record is written, so
    // model_name stays null and no ABSTAINED disagreement is recorded.
    expect(run?.model_name).toBeNull()
    expect(run?.ai_structured_output_json).toBeNull()
    expect(JSON.parse(run?.disagreement_flags_json ?? '[]')).not.toContain('ABSTAINED')
  })

  it('never invokes the interpreter for non-complaint mail, even when enabled', async () => {
    let calls = 0
    const interpreter = new FakeMailInterpreter(() => {
      calls += 1
      return JSON.stringify(validOutput())
    })
    const result = await ingestEmailMessage(
      db,
      message({
        id: 'noncomplaint-1',
        threadId: 'noncomplaint-thread',
        subject: 'Weekly newsletter',
        textBody: 'Here is our weekly newsletter with store updates and promotions.',
      }),
      { interpreter, interpretationConfig: enabledConfig },
    )
    expect(result.status).toBe('IGNORED')
    expect(calls).toBe(0)
    // No interpretation record: the run row is untouched by Stage 3b.
    const run = await runRow('noncomplaint-1')
    expect(run?.model_name).toBeNull()
    expect(run?.ai_structured_output_json).toBeNull()
    expect(run?.disagreement_flags_json).toBe('[]')
  })

  it('builds a bounded interpretation input from the deterministic extraction', async () => {
    const { extractComplaint } = await import('../worker/ingestion')
    const extraction = extractComplaint(message())
    expect(extraction.extractionRules).toContain('CASE_ID_LINE_ANCHORED')
    const input = buildInterpretationInput(message(), extraction, { reason: 'NO_DETERMINISTIC_STORE_MATCH' }, enabledConfig, 'production')
    expect(input.schemaVersion).toBe('mail-interpret.input.v1')
    expect(input.normalized.externalCaseId).toBe('DD-1001')
    expect(input.sourceMetadata.bodyExcerpt.length).toBeLessThanOrEqual(4000)
  })
})
