import { useState } from 'react'
import { api } from '../lib/api'

type AckPreview = Awaited<ReturnType<typeof api.ackPreview>>

/**
 * Phase 2 — Manual Dunkin acknowledgment. Owner-only. Shows the exact
 * template, intended recipient/thread, mailbox identity, and ack status
 * before any send. Requires explicit confirm. Never auto-sends.
 * Global automatic ack switch stays OFF during manual pilot.
 */
export function ManualAck({
  complaintId,
  source,
  acknowledgementStatus,
  isOwner,
}: {
  complaintId: string
  source?: string
  acknowledgementStatus?: string
  isOwner: boolean
}) {
  const [preview, setPreview] = useState<AckPreview | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [sent, setSent] = useState(false)
  const [showReclaim, setShowReclaim] = useState(false)
  const [sentConfirmed, setSentConfirmed] = useState(false)

  if (!isOwner) return null
  if (source !== 'MICROSOFT_GRAPH') return null
  if (acknowledgementStatus === 'SENT') {
    return (
      <section className="panel">
        <div className="panel-head">
          <div>
            <h2>Dunkin acknowledgment</h2>
            <p>Acknowledgment already sent. No second reply will be sent.</p>
          </div>
        </div>
        <p style={{ color: 'var(--muted, #666)' }}>Status: SENT</p>
      </section>
    )
  }

  async function loadPreview() {
    setLoading(true)
    setError(null)
    try {
      const res = await api.ackPreview(complaintId)
      setPreview(res)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to load ack preview')
    } finally {
      setLoading(false)
    }
  }

  async function sendAck() {
    if (!window.confirm('Send this exact acknowledgment to Dunkin in this thread? This cannot be undone.')) return
    setLoading(true)
    setError(null)
    try {
      await api.ackSend(complaintId)
      setSent(true)
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Send failed'
      setError(msg)
      if (msg.includes('AMBIGUOUS') || msg.includes('IN_FLIGHT_STALE')) setShowReclaim(true)
    } finally {
      setLoading(false)
    }
  }

  async function reclaim() {
    if (!window.confirm('Confirm you checked Sent Items / the thread and are reporting what you found.')) return
    setLoading(true)
    setError(null)
    try {
      const res = await api.ackReclaim(complaintId, {
        checkedSentItems: true,
        sentConfirmed,
      })
      setError(`Reclaim result: ${res.result}. Refresh the page to see updated status.`)
      setShowReclaim(false)
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Reclaim failed'
      // Absence from Sent Items alone is NOT conclusive — the ack stays
      // IN_FLIGHT pending further reconciliation. Do not auto-retry.
      if (msg.includes('NEEDS_FURTHER_RECONCILIATION')) {
        setError('Not in Sent Items alone is not conclusive proof it was not sent. The ack stays in pending-reconciliation (IN_FLIGHT). Check the thread, wait, and re-check before any force-reset with a written note.')
      } else {
        setError(msg)
      }
    } finally {
      setLoading(false)
    }
  }

  if (sent) {
    return (
      <section className="panel">
        <div className="panel-head">
          <div>
            <h2>Dunkin acknowledgment</h2>
            <p>Acknowledgment sent. Refresh to see updated status.</p>
          </div>
        </div>
      </section>
    )
  }

  return (
    <section className="panel">
      <div className="panel-head">
        <div>
          <h2>Dunkin acknowledgment</h2>
          <p>
            Manual, one-complaint acknowledgment. Shows the exact message and recipient before
            sending. Global automatic acknowledgment stays OFF.
          </p>
        </div>
      </div>
      <p style={{ color: 'var(--muted, #666)' }}>
        Current ack status: {acknowledgementStatus ?? 'DISABLED'}
        {preview?.ack ? ` · Attempts: ${preview.ack.attempt_count} · Last error: ${preview.ack.last_error_code ?? 'none'}` : ''}
      </p>
      {!preview && (
        <button className="secondary" onClick={loadPreview} disabled={loading}>
          {loading ? 'Loading…' : 'Preview acknowledgment'}
        </button>
      )}
      {preview && (
        <div style={{ marginTop: 12 }}>
          <p>
            <strong>Recipient (source sender):</strong> {preview.source.sender}
          </p>
          <p>
            <strong>Thread:</strong> {preview.source.threadIdMasked}
          </p>
          <p>
            <strong>Subject:</strong> {preview.source.subject}
          </p>
          <p>
            <strong>Mailbox identity:</strong>{' '}
            {preview.mailbox?.matches ? 'Verified — matches expected complaint mailbox' : `NOT verified — ${preview.mailbox?.reason ?? 'unknown'}`}
          </p>
          <p>
            <strong>Exact message to send:</strong>
          </p>
          <blockquote style={{ borderLeft: '3px solid #ccc', paddingLeft: 12, margin: '8px 0' }}>
            {preview.template}
          </blockquote>
          {!preview.canSend && (
            <p style={{ color: '#b00' }}>
              Cannot send: mailbox identity is not verified. Check mailbox configuration before sending.
            </p>
          )}
          {preview.canSend && (
            <button className="primary" onClick={sendAck} disabled={loading}>
              {loading ? 'Sending…' : 'Send acknowledgment to Dunkin'}
            </button>
          )}
          {(preview.ack?.status === 'IN_FLIGHT' || showReclaim) && (
            <div style={{ marginTop: 16, borderTop: '1px solid #ddd', paddingTop: 12 }}>
              <p>
                <strong>Stale IN_FLIGHT / ambiguous send:</strong> Check Graph Sent Items / the
                thread for this reply. Report what you found — do not send again to &quot;make
                sure&quot;.
              </p>
              <label style={{ display: 'block', margin: '8px 0' }}>
                <input type="checkbox" checked={sentConfirmed} onChange={(e) => setSentConfirmed(e.target.checked)} />{' '}
                I checked Sent Items / the thread and the reply IS there (mark SENT without sending again)
              </label>
              <label style={{ display: 'block', margin: '8px 0' }}>
                <input type="checkbox" checked={!sentConfirmed} onChange={(e) => setSentConfirmed(!e.target.checked)} />{' '}
                I checked and the reply is NOT there (reset to PENDING for a manual retry)
              </label>
              <button className="secondary" onClick={reclaim} disabled={loading}>
                Reclaim
              </button>
            </div>
          )}
        </div>
      )}
      {error && <p style={{ color: '#b00', marginTop: 8 }}>{error}</p>}
    </section>
  )
}
