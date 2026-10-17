import type { Severity } from '../src/lib/types'
import type { D1Database } from './d1'
import type { NormalizedEmailMessage } from './email-provider'
import type { ComplaintExtraction } from './ingestion'

// --- Phase B: structured interpretation (shadow / review-only) ----------------
// This module defines the interpretation contract and a configurable interpreter
// interface. It NEVER creates/updates complaints, merges records, invents store
// mappings, sends messages, or touches rollout policy: it takes data in and
// returns a result object. All persistence happens in ingestion.ts via explicit
// statements. The only shipped model transport (DisabledModelTransport) throws —
// there is no fetch() call and no API key plumbing anywhere in this module.

export class MailInterpretationError extends Error {
  constructor(
    readonly code:
      | 'UNKNOWN_PROMPT_VERSION'
      | 'MODEL_TRANSPORT_DISABLED'
      | 'MODEL_TRANSPORT_FAILED'
      | 'INTERPRETATION_INVALID',
    message: string,
  ) {
    super(message)
    this.name = 'MailInterpretationError'
  }
}

// --- Configuration -----------------------------------------------------------

export interface MailInterpretationConfig {
  /** Master switch. Defaults OFF; Stage 3b is skipped entirely when false. */
  enabled: boolean
  /** Model id passed to the transport. Default 'gpt-4o-mini'. */
  model: string
  /** Immutable prompt version. Default 'mail-interpret.v1'. */
  promptVersion: string
  /** Phase B: always true. No code path lets the interpreter write or send. */
  shadowMode: boolean
}

export const DEFAULT_INTERPRETATION_MODEL = 'gpt-4o-mini'
export const DEFAULT_INTERPRETATION_PROMPT_VERSION = 'mail-interpret.v1'
export const INTERPRETATION_LOW_CONFIDENCE_THRESHOLD = 0.6
export const INTERPRETATION_INPUT_SCHEMA_VERSION = 'mail-interpret.input.v1'
export const INTERPRETATION_OUTPUT_SCHEMA_VERSION = 'mail-interpret.output.v1'
const INTERPRETATION_BODY_EXCERPT_CHARS = 4000

export const DISABLED_INTERPRETATION_CONFIG: MailInterpretationConfig = {
  enabled: false,
  model: DEFAULT_INTERPRETATION_MODEL,
  promptVersion: DEFAULT_INTERPRETATION_PROMPT_VERSION,
  shadowMode: true,
}

const SETTING_KEYS = [
  'mail_interpretation_enabled',
  'mail_interpretation_model',
  'mail_interpretation_prompt_version',
] as const

/** Reads the interpreter config from the settings table; every key defaults OFF/safe. */
export async function loadMailInterpretationConfig(
  db: D1Database,
): Promise<MailInterpretationConfig> {
  const rows = await db
    .prepare('SELECT key,value FROM settings WHERE key IN (?,?,?)')
    .bind(...SETTING_KEYS)
    .all<{ key: string; value: string }>()
  const values = Object.fromEntries(rows.results.map((row) => [row.key, row.value]))
  return {
    enabled: values.mail_interpretation_enabled === 'true',
    model: values.mail_interpretation_model || DEFAULT_INTERPRETATION_MODEL,
    promptVersion: values.mail_interpretation_prompt_version || DEFAULT_INTERPRETATION_PROMPT_VERSION,
    shadowMode: true,
  }
}

// --- Input schema ------------------------------------------------------------

export type InterpretationCategory = 'Cleanliness' | 'Service' | 'Product quality' | 'Other'

export interface InterpretationStoreResolution {
  storeNumber?: string
  reason: 'EXACT_STORE_NUMBER' | 'EXACT_ALIAS' | 'NO_DETERMINISTIC_STORE_MATCH'
}

export interface MailInterpretationInput {
  schemaVersion: typeof INTERPRETATION_INPUT_SCHEMA_VERSION
  promptVersion: string
  normalized: {
    subject: string
    senderAddress: string
    receivedAt: string
    conversationId?: string
    isComplaint: boolean
    externalCaseId?: string
    storeNumber?: string
    locationHint?: string
    customerName?: string
    customerEmail?: string
    customerPhone?: string
    occurrenceAt?: string
    category: InterpretationCategory
    severity: Severity
  }
  deterministicEvidence: {
    extractionRules: string[]
    storeResolution: InterpretationStoreResolution
  }
  sourceMetadata: {
    provider: 'MICROSOFT_GRAPH'
    mailboxKey: string
    bodyExcerpt: string
    internetMessageId?: string
  }
}

export function buildInterpretationInput(
  message: NormalizedEmailMessage,
  extraction: ComplaintExtraction,
  storeResolution: InterpretationStoreResolution,
  config: MailInterpretationConfig,
  mailboxKey: string,
): MailInterpretationInput {
  return {
    schemaVersion: INTERPRETATION_INPUT_SCHEMA_VERSION,
    promptVersion: config.promptVersion,
    normalized: {
      subject: message.subject,
      senderAddress: message.sender,
      receivedAt: message.internalDate,
      conversationId: message.threadId || undefined,
      isComplaint: extraction.isComplaint,
      externalCaseId: extraction.externalCaseId || undefined,
      storeNumber: extraction.storeNumber || undefined,
      locationHint: extraction.locationHint || undefined,
      customerName: extraction.customerName || undefined,
      customerEmail: extraction.customerEmail || undefined,
      customerPhone: extraction.customerPhone || undefined,
      occurrenceAt: extraction.occurrenceAt || undefined,
      category: extraction.category as InterpretationCategory,
      severity: extraction.severity,
    },
    deterministicEvidence: {
      extractionRules: extraction.extractionRules,
      storeResolution,
    },
    sourceMetadata: {
      provider: 'MICROSOFT_GRAPH',
      mailboxKey,
      bodyExcerpt: message.textBody.slice(0, INTERPRETATION_BODY_EXCERPT_CHARS),
      internetMessageId: message.messageIdHeader || undefined,
    },
  }
}

// --- Prompt registry (immutable, append-only) ---------------------------------

const MAIL_INTERPRET_V1 = `You are a precise information-extraction assistant for Dunkin' franchise guest-complaint email.
Extract ONLY what is explicitly stated in the message. Never infer, never invent.

RULES
- externalCaseId: copy a complaint/case/reference ID exactly as written, or null when none is stated. A guessed or plausible-looking ID is worse than null.
- storeNumber: copy the digits of a store/location number exactly as written, or null. Do NOT map names, addresses, or descriptions to numbers.
- issueCategory: exactly one of Cleanliness | Service | Product quality | Other — the single best fit.
- urgency: exactly one of LOW | MEDIUM | HIGH | CRITICAL, based on the text (safety, injury, contamination issues are HIGH or CRITICAL).
- confidence: 0 to 1, your overall confidence in this extraction. Below 0.6 means uncertain.
- summary: one or two sentences, at most 500 characters.
- evidenceQuotes: zero to four short quotes copied word-for-word from the message that support your extraction, each at most 280 characters. Every quote must appear verbatim in the message.

Return ONLY a JSON object with exactly these keys and no extra text:
{"schemaVersion":"mail-interpret.output.v1","externalCaseId":string|null,"storeNumber":string|null,"storeConfidence":number|null,"issueCategory":"Cleanliness"|"Service"|"Product quality"|"Other","urgency":"LOW"|"MEDIUM"|"HIGH"|"CRITICAL","confidence":number,"summary":string,"evidenceQuotes":string[]}`

export const INTERPRETATION_PROMPTS: Record<string, string> = {
  'mail-interpret.v1': MAIL_INTERPRET_V1,
}

export function interpretationPrompt(version: string): string {
  const prompt = INTERPRETATION_PROMPTS[version]
  if (!prompt)
    throw new MailInterpretationError(
      'UNKNOWN_PROMPT_VERSION',
      `Unknown interpretation prompt version: ${version}`,
    )
  return prompt
}

// --- Output schema + validation ----------------------------------------------

export interface MailInterpretationOutput {
  schemaVersion: typeof INTERPRETATION_OUTPUT_SCHEMA_VERSION
  externalCaseId: string | null
  storeNumber: string | null
  storeConfidence: number | null
  issueCategory: InterpretationCategory
  urgency: Severity
  confidence: number
  summary: string
  evidenceQuotes: string[]
}

const CATEGORIES: InterpretationCategory[] = ['Cleanliness', 'Service', 'Product quality', 'Other']
const SEVERITIES: Severity[] = ['LOW', 'MEDIUM', 'HIGH', 'CRITICAL']

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null

/** Strict validation: unknown shapes become SCHEMA_VIOLATION, never partial objects. */
export function validateInterpretationOutput(
  value: unknown,
): { valid: boolean; errors: string[]; output: MailInterpretationOutput | null } {
  const errors: string[] = []
  if (!isRecord(value)) return { valid: false, errors: ['output is not an object'], output: null }
  const get = (key: string) => value[key]
  const check = (condition: boolean, message: string) => {
    if (!condition) errors.push(message)
  }
  check(
    get('schemaVersion') === INTERPRETATION_OUTPUT_SCHEMA_VERSION,
    `schemaVersion must be ${INTERPRETATION_OUTPUT_SCHEMA_VERSION}`,
  )
  const externalCaseId = get('externalCaseId')
  check(
    externalCaseId === null || (typeof externalCaseId === 'string' && externalCaseId.length > 0 && externalCaseId.length <= 64),
    'externalCaseId must be a non-empty string or null',
  )
  const storeNumber = get('storeNumber')
  check(
    storeNumber === null || (typeof storeNumber === 'string' && /^\d{3,8}$/.test(storeNumber)),
    'storeNumber must be 3-8 digits or null',
  )
  const storeConfidence = get('storeConfidence')
  // A null storeConfidence alongside a storeNumber is a signal (the model is
  // unsure), not a schema error: disagreement detection turns it into
  // LOW_CONFIDENCE instead of SCHEMA_VIOLATION.
  check(
    (storeNumber === null && storeConfidence === null) ||
      (storeNumber !== null &&
        (storeConfidence === null ||
          (typeof storeConfidence === 'number' && storeConfidence >= 0 && storeConfidence <= 1))),
    'storeConfidence must be 0..1 or null when storeNumber is set, null otherwise',
  )
  check(
    typeof get('issueCategory') === 'string' && CATEGORIES.includes(get('issueCategory') as InterpretationCategory),
    `issueCategory must be one of ${CATEGORIES.join(', ')}`,
  )
  check(
    typeof get('urgency') === 'string' && SEVERITIES.includes(get('urgency') as Severity),
    `urgency must be one of ${SEVERITIES.join(', ')}`,
  )
  const confidence = get('confidence')
  check(typeof confidence === 'number' && confidence >= 0 && confidence <= 1, 'confidence must be 0..1')
  const summary = get('summary')
  check(typeof summary === 'string' && summary.length > 0 && summary.length <= 500, 'summary must be 1..500 chars')
  const quotes = get('evidenceQuotes')
  check(
    Array.isArray(quotes) &&
      quotes.length <= 4 &&
      quotes.every((quote) => typeof quote === 'string' && quote.length > 0 && quote.length <= 280),
    'evidenceQuotes must be an array of 0..4 strings of 1..280 chars',
  )
  if (errors.length > 0) return { valid: false, errors, output: null }
  return {
    valid: true,
    errors: [],
    output: {
      schemaVersion: INTERPRETATION_OUTPUT_SCHEMA_VERSION,
      externalCaseId: externalCaseId as string | null,
      storeNumber: storeNumber as string | null,
      storeConfidence: storeConfidence as number | null,
      issueCategory: get('issueCategory') as InterpretationCategory,
      urgency: get('urgency') as Severity,
      confidence: confidence as number,
      summary: summary as string,
      evidenceQuotes: quotes as string[],
    },
  }
}

// --- Disagreement taxonomy ----------------------------------------------------

export type DisagreementCode =
  | 'CASE_ID_MISMATCH'
  | 'CASE_ID_UNVERIFIED'
  | 'CASE_ID_MISSED'
  | 'STORE_MISMATCH'
  | 'STORE_UNVERIFIED'
  | 'STORE_MISSED'
  | 'CATEGORY_MISMATCH'
  | 'SEVERITY_MISMATCH'
  | 'SEVERITY_ESCALATION_CANDIDATE'
  | 'LOW_CONFIDENCE'
  | 'ABSTAINED'
  | 'EVIDENCE_CONTRADICTION'
  | 'SCHEMA_VIOLATION'

export interface InterpretationDisagreement {
  code: DisagreementCode
  detail: string
}

const SEVERITY_RANK: Record<Severity, number> = { LOW: 0, MEDIUM: 1, HIGH: 2, CRITICAL: 3 }
const normalizeWhitespace = (value: string) => value.replace(/\s+/g, ' ').trim()

/**
 * Pure function: compares model output against deterministic evidence.
 * The model is never trusted for identity — a model-only case ID yields
 * CASE_ID_UNVERIFIED (a review hint), never an identity.
 */
export function detectDisagreements(
  input: MailInterpretationInput,
  output: MailInterpretationOutput | null,
  validationErrors: string[],
): InterpretationDisagreement[] {
  const disagreements: InterpretationDisagreement[] = []
  if (validationErrors.length > 0)
    disagreements.push({
      code: 'SCHEMA_VIOLATION',
      detail: `output failed validation: ${validationErrors.join('; ')}`,
    })
  if (!output) {
    disagreements.push({ code: 'ABSTAINED', detail: 'model returned no usable output' })
    return disagreements
  }

  const deterministicCaseId = input.normalized.externalCaseId ?? null
  if (deterministicCaseId && output.externalCaseId && deterministicCaseId !== output.externalCaseId)
    disagreements.push({
      code: 'CASE_ID_MISMATCH',
      detail: `deterministic=${deterministicCaseId} model=${output.externalCaseId}; deterministic wins`,
    })
  if (!deterministicCaseId && output.externalCaseId)
    disagreements.push({
      code: 'CASE_ID_UNVERIFIED',
      detail: `model suggested ${output.externalCaseId} with no deterministic evidence; provisional review only`,
    })
  if (deterministicCaseId && !output.externalCaseId)
    disagreements.push({
      code: 'CASE_ID_MISSED',
      detail: `model missed deterministic case id ${deterministicCaseId}`,
    })

  const deterministicStore = input.deterministicEvidence.storeResolution.storeNumber ?? null
  if (deterministicStore && output.storeNumber && deterministicStore !== output.storeNumber)
    disagreements.push({
      code: 'STORE_MISMATCH',
      detail: `deterministic=${deterministicStore} model=${output.storeNumber}; deterministic wins`,
    })
  if (!deterministicStore && output.storeNumber)
    disagreements.push({
      code: 'STORE_UNVERIFIED',
      detail: `model suggested store ${output.storeNumber} with no DB match; mapping not invented`,
    })
  if (deterministicStore && !output.storeNumber)
    disagreements.push({
      code: 'STORE_MISSED',
      detail: `model missed deterministic store ${deterministicStore}`,
    })

  if (input.normalized.category !== output.issueCategory)
    disagreements.push({
      code: 'CATEGORY_MISMATCH',
      detail: `deterministic=${input.normalized.category} model=${output.issueCategory}`,
    })
  if (input.normalized.severity !== output.urgency) {
    disagreements.push({
      code: 'SEVERITY_MISMATCH',
      detail: `deterministic=${input.normalized.severity} model=${output.urgency}`,
    })
    if (SEVERITY_RANK[output.urgency] > SEVERITY_RANK[input.normalized.severity])
      disagreements.push({
        code: 'SEVERITY_ESCALATION_CANDIDATE',
        detail: `model rates urgency higher (${output.urgency}) than deterministic severity (${input.normalized.severity}); human confirms any escalation`,
      })
  }

  if (output.confidence < INTERPRETATION_LOW_CONFIDENCE_THRESHOLD)
    disagreements.push({
      code: 'LOW_CONFIDENCE',
      detail: `model confidence ${output.confidence} below ${INTERPRETATION_LOW_CONFIDENCE_THRESHOLD}`,
    })
  if (output.storeNumber !== null && output.storeConfidence === null)
    disagreements.push({
      code: 'LOW_CONFIDENCE',
      detail: `model suggested store ${output.storeNumber} with null storeConfidence; treat as uncertain`,
    })

  const excerpt = normalizeWhitespace(input.sourceMetadata.bodyExcerpt)
  for (const quote of output.evidenceQuotes)
    if (!excerpt.includes(normalizeWhitespace(quote)))
      disagreements.push({
        code: 'EVIDENCE_CONTRADICTION',
        detail: `evidence quote not found verbatim in source text: ${quote.slice(0, 80)}`,
      })

  return disagreements
}

// --- Interpreter interface ----------------------------------------------------

export interface RawModelResponse {
  rawText: string | null
  latencyMs: number
  errorCode?: string
}

export interface MailInterpreter {
  readonly name: string
  interpretRaw(input: MailInterpretationInput): Promise<RawModelResponse>
}

export interface MailInterpretationResult {
  model: string
  promptVersion: string
  output: MailInterpretationOutput | null
  rawText: string | null
  inputSha256: string
  outputSha256: string | null
  latencyMs: number
  confidence: number | null
  disagreements: InterpretationDisagreement[]
  validation: { valid: boolean; errors: string[] }
  skipped: boolean
}

const sha256Hex = async (text: string): Promise<string> => {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
  const bytes = new Uint8Array(digest)
  let hex = ''
  for (let i = 0; i < bytes.length; i++) hex += bytes[i]!.toString(16).padStart(2, '0')
  return hex
}

/** Shared post-processing: parse, validate, detect disagreements. Never throws for model issues. */
export async function processRawInterpretation(
  input: MailInterpretationInput,
  raw: RawModelResponse,
  config: MailInterpretationConfig,
): Promise<MailInterpretationResult> {
  const base = {
    model: config.model,
    promptVersion: config.promptVersion,
    inputSha256: await sha256Hex(JSON.stringify(input)),
    latencyMs: raw.latencyMs,
    skipped: false,
  }
  // Phase 1 fix: a disabled transport is NOT a model abstention. It means no
  // live model exists, so the pipeline must proceed deterministically and
  // the caller must treat this as skipped (no interpretation record), not as
  // a model signal. ABSTAINED is reserved for a live transport that ran but
  // returned no usable output.
  if (raw.errorCode === 'MODEL_TRANSPORT_DISABLED')
    return {
      ...base,
      output: null,
      rawText: null,
      outputSha256: null,
      confidence: null,
      disagreements: [],
      validation: { valid: false, errors: ['interpreter disabled - no live transport'] },
      skipped: true,
    }
  if (raw.errorCode || raw.rawText === null)
    return {
      ...base,
      output: null,
      rawText: raw.rawText,
      outputSha256: null,
      confidence: null,
      disagreements: [
        {
          code: 'ABSTAINED',
          detail: raw.errorCode ? `model transport failed: ${raw.errorCode}` : 'model returned no output',
        },
      ],
      validation: { valid: false, errors: [raw.errorCode ?? 'empty model output'] },
    }
  let parsed: unknown
  try {
    parsed = JSON.parse(raw.rawText)
  } catch {
    return {
      ...base,
      output: null,
      rawText: raw.rawText,
      outputSha256: await sha256Hex(raw.rawText),
      confidence: null,
      disagreements: [{ code: 'SCHEMA_VIOLATION', detail: 'model output was not valid JSON' }],
      validation: { valid: false, errors: ['model output was not valid JSON'] },
    }
  }
  const validation = validateInterpretationOutput(parsed)
  const disagreements = detectDisagreements(input, validation.output, validation.errors)
  return {
    ...base,
    output: validation.output,
    rawText: raw.rawText,
    outputSha256: await sha256Hex(raw.rawText),
    confidence: validation.output?.confidence ?? null,
    disagreements,
    validation: { valid: validation.valid, errors: validation.errors },
  }
}

/**
 * Runs the interpreter behind the config gate. Throws only for configuration
 * errors (unknown prompt version); model-level failures become ABSTAINED results.
 *
 * Architectural cost/privacy boundary: the interpreter is ONLY invoked for
 * messages the deterministic stage flagged as complaint candidates
 * (input.normalized.isComplaint). Non-candidates are refused here — no prompt is
 * built, no model is touched, no tokens are spent — regardless of caller.
 */
export async function interpretWithConfig(
  interpreter: MailInterpreter,
  input: MailInterpretationInput,
  config: MailInterpretationConfig,
): Promise<MailInterpretationResult> {
  const refused = (reason: string): MailInterpretationResult => ({
    model: config.model,
    promptVersion: config.promptVersion,
    output: null,
    rawText: null,
    inputSha256: '',
    outputSha256: null,
    latencyMs: 0,
    confidence: null,
    disagreements: [],
    validation: { valid: false, errors: [reason] },
    skipped: true,
  })
  if (!config.enabled) return refused('interpreter disabled')
  if (!input.normalized.isComplaint) return refused('not a deterministic complaint candidate')
  interpretationPrompt(config.promptVersion) // throws MailInterpretationError on unknown version
  const started = Date.now()
  let raw: RawModelResponse
  try {
    raw = await interpreter.interpretRaw(input)
  } catch (error) {
    raw = {
      rawText: null,
      latencyMs: Date.now() - started,
      errorCode: error instanceof MailInterpretationError ? error.code : 'INTERPRETER_FAILED',
    }
  }
  return processRawInterpretation(input, raw, config)
}

// --- Transports ----------------------------------------------------------------
// Only DisabledModelTransport ships. A live transport needs explicit approval and
// does not exist in this codebase yet.

export interface ModelCompletionRequest {
  model: string
  messages: { role: 'system' | 'user'; content: string }[]
  temperature: number
  max_tokens: number
  response_format: { type: 'json_object' }
}

export interface ModelCompletionResponse {
  contentText: string | null
  latencyMs: number
}

export interface ModelTransport {
  complete(request: ModelCompletionRequest): Promise<ModelCompletionResponse>
}

export class DisabledModelTransport implements ModelTransport {
  async complete(): Promise<ModelCompletionResponse> {
    throw new MailInterpretationError(
      'MODEL_TRANSPORT_DISABLED',
      'Live model calls are disabled: the Phase B transport is stubbed and paid model spend needs explicit approval.',
    )
  }
}

/**
 * OpenAI-compatible interpreter (chat-completions request shape). Constructed with
 * a transport; production wires DisabledModelTransport, so this can never reach
 * the network. The future live transport would POST the built request to the
 * chat-completions endpoint with a stored credential — not implemented here.
 */
export class OpenAiMailInterpreter implements MailInterpreter {
  readonly name = 'openai-mail-interpreter'
  constructor(
    private readonly config: MailInterpretationConfig,
    private readonly transport: ModelTransport,
  ) {}

  buildRequest(input: MailInterpretationInput): ModelCompletionRequest {
    return {
      model: this.config.model,
      messages: [
        { role: 'system', content: interpretationPrompt(this.config.promptVersion) },
        { role: 'user', content: JSON.stringify(input) },
      ],
      temperature: 0,
      max_tokens: 800,
      response_format: { type: 'json_object' },
    }
  }

  async interpretRaw(input: MailInterpretationInput): Promise<RawModelResponse> {
    try {
      const response = await this.transport.complete(this.buildRequest(input))
      return { rawText: response.contentText, latencyMs: response.latencyMs }
    } catch (error) {
      if (error instanceof MailInterpretationError) throw error
      throw new MailInterpretationError(
        'MODEL_TRANSPORT_FAILED',
        error instanceof Error ? error.message : 'model transport failed',
      )
    }
  }
}

/** Deterministic fake for tests: canned raw responses keyed by caller-supplied logic. */
export class FakeMailInterpreter implements MailInterpreter {
  readonly name = 'fake-mail-interpreter'
  constructor(
    private readonly respond: (input: MailInterpretationInput) => string | null,
    private readonly latencyMs = 5,
  ) {}
  async interpretRaw(input: MailInterpretationInput): Promise<RawModelResponse> {
    return { rawText: this.respond(input), latencyMs: this.latencyMs }
  }
}

// --- Persistence ---------------------------------------------------------------

/** Writes the full interpretation record onto the processing run. Existing columns only. */
export async function recordInterpretation(
  db: D1Database,
  runId: string,
  input: MailInterpretationInput,
  result: MailInterpretationResult,
): Promise<void> {
  const now = new Date().toISOString()
  await db
    .prepare(
      `UPDATE mail_processing_runs SET
         model_name=?, prompt_version=?, schema_version=?,
         normalized_output_json=?, deterministic_evidence_json=?,
         ai_structured_output_json=?, model_input_sha256=?, model_output_sha256=?,
         validation_result_json=?, disagreement_flags_json=?, latency_ms=?, updated_at=?
       WHERE id=?`,
    )
    .bind(
      result.model,
      result.promptVersion,
      result.output ? INTERPRETATION_OUTPUT_SCHEMA_VERSION : INTERPRETATION_INPUT_SCHEMA_VERSION,
      JSON.stringify(input.normalized),
      JSON.stringify(input.deterministicEvidence),
      result.output ? JSON.stringify(result.output) : null,
      result.inputSha256,
      result.outputSha256,
      JSON.stringify(result.validation),
      JSON.stringify(result.disagreements.map((disagreement) => disagreement.code)),
      result.latencyMs,
      now,
      runId,
    )
    .run()
}
