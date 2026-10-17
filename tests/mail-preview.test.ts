// Phase 1: preview classification, masking, window validation, forward/reply handling.
import { describe, expect, it } from 'vitest'
import {
  classifyPreviewMessage,
  maskGraphIdentifier,
  maskSenderAddress,
  validatePreviewWindow,
} from '../worker/mail-preview'
import { extractComplaint, isForwardOrReplyMessage } from '../worker/ingestion'
import { processRawInterpretation } from '../worker/mail-interpretation'
import type { GraphMessage } from '../worker/microsoft-graph'

const previewMessage = (overrides: Partial<GraphMessage> = {}): GraphMessage => ({
  id: 'AQMkAD1234567890abcdef',
  conversationId: 'conv-1',
  receivedDateTime: '2026-10-12T10:00:00.000Z',
  subject: 'Complaint - Case CCC12345 - Store 350909',
  from: { emailAddress: { address: 'customerservice@dunkinbrands.com' } },
  bodyPreview: 'Guest complaint about cold coffee. Case: CCC12345. Store: 350909.',
  ...overrides,
})

describe('mail preview', () => {
  it('masks Graph ids and personal senders, shows Dunkin sender fully', () => {
    expect(maskGraphIdentifier('AQMkAD1234567890abcdef')).toBe('AQMk…cdef')
    expect(maskSenderAddress('customerservice@dunkinbrands.com')).toBe(
      'customerservice@dunkinbrands.com',
    )
    expect(maskSenderAddress('john.smith@yahoo.com')).toBe('jo***@yahoo.com')
  })

  it('validates preview windows', () => {
    expect(
      validatePreviewWindow('2026-10-10T00:00:00.000Z', '2026-10-17T23:59:59.999Z'),
    ).toEqual({ ok: true })
    expect(validatePreviewWindow('2026-10-17T00:00:00.000Z', '2026-10-10T00:00:00.000Z')).toEqual({
      ok: false,
      error: 'INVALID_WINDOW',
    })
    expect(validatePreviewWindow('2026-01-01T00:00:00.000Z', '2026-03-01T00:00:00.000Z')).toEqual({
      ok: false,
      error: 'WINDOW_TOO_LARGE',
    })
  })

  it('classifies a complaint candidate with case id + store', () => {
    const result = classifyPreviewMessage(previewMessage(), new Set(), new Map(), new Map())
    expect(result.classification).toBe('COMPLAINT_CANDIDATE')
    expect(result.detectedCaseId).toBe('CCC12345')
    expect(result.detectedStoreNumber).toBe('350909')
  })

  it('flags Dunkin mail with no deterministic match as NEEDS_REVIEW, not NOT_COMPLAINT', () => {
    const result = classifyPreviewMessage(
      previewMessage({ subject: 'Your receipt', bodyPreview: 'Thanks for visiting.' }),
      new Set(),
      new Map(),
      new Map(),
    )
    expect(result.classification).toBe('NEEDS_REVIEW')
  })

  it('flags forwards/replies as NEEDS_REVIEW when preview has no complaint match', () => {
    const result = classifyPreviewMessage(
      previewMessage({
        subject: 'Fw: Something to check',
        from: { emailAddress: { address: 'manager@example.com' } },
        bodyPreview: 'FYI — see below.',
      }),
      new Set(),
      new Map(),
      new Map(),
    )
    expect(result.classification).toBe('NEEDS_REVIEW')
    expect(result.isForwardOrReply).toBe(true)
  })

  it('marks already-ingested ids as ALREADY_INGESTED', () => {
    const result = classifyPreviewMessage(
      previewMessage(),
      new Set(['AQMkAD1234567890abcdef']),
      new Map(),
      new Map(),
    )
    expect(result.classification).toBe('ALREADY_INGESTED')
    expect(result.alreadyIngested).toBe(true)
  })

  it('marks case-id matches to existing complaints as FOLLOW_UP_LIKELY', () => {
    const result = classifyPreviewMessage(
      previewMessage(),
      new Set(),
      new Map([['ccc12345', 'SR-2026-0001']]),
      new Map(),
    )
    expect(result.classification).toBe('FOLLOW_UP_LIKELY')
    expect(result.linkedComplaintId).toBe('SR-2026-0001')
  })

  it('detects forward/reply markers and case-id in forwarded complaint', () => {
    expect(isForwardOrReplyMessage('Fw: Complaint', 'FYI')).toBe(true)
    expect(isForwardOrReplyMessage('Re: Case CCC123', 'Thanks')).toBe(true)
    expect(isForwardOrReplyMessage('Complaint', 'Hello')).toBe(false)
    const extraction = extractComplaint({
      id: 'm1',
      threadId: 'conv-1',
      internalDate: '2026-10-12T10:00:00.000Z',
      subject: 'Fw: FYI',
      sender: 'manager@example.com',
      recipients: '',
      textBody: 'FYI\n---Original Message---\nComplaint about cold coffee. Case: CCC99999.',
      htmlBody: undefined,
      messageIdHeader: undefined,
      inReplyTo: undefined,
      references: undefined,
    })
    expect(extraction.isForwardOrReply).toBe(true)
    expect(extraction.hasCaseIdSignal).toBe(true)
    expect(extraction.isComplaint).toBe(true)
    expect(extraction.extractionRules).toContain('FORWARD_REPLY_WITH_CASE_ID')
  })

  it('treats a disabled model transport as skipped, not ABSTAINED', async () => {
    const input = {
      schemaVersion: 'mail-interpret.input.v1' as const,
      promptVersion: 'mail-interpret.v1',
      normalized: {
        subject: 'Complaint',
        senderAddress: 'customerservice@dunkinbrands.com',
        receivedAt: '2026-10-12T10:00:00.000Z',
        isComplaint: true,
        category: 'Service' as const,
        severity: 'LOW' as const,
      },
      deterministicEvidence: {
        extractionRules: [],
        storeResolution: { reason: 'NO_DETERMINISTIC_STORE_MATCH' as const },
      },
      sourceMetadata: {
        provider: 'MICROSOFT_GRAPH' as const,
        mailboxKey: 'production',
        bodyExcerpt: 'Complaint',
      },
    }
    const result = await processRawInterpretation(
      input,
      { rawText: null, latencyMs: 1, errorCode: 'MODEL_TRANSPORT_DISABLED' },
      { enabled: true, model: 'gpt-4o-mini', promptVersion: 'mail-interpret.v1', shadowMode: true },
    )
    expect(result.skipped).toBe(true)
    expect(result.disagreements).toEqual([])
  })
})
