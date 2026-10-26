import { useEffect, useState } from 'react'
import { api } from '../lib/api'

type ResponseAction = {
  id: string
  complaintId: string
  responseKind: string
  policyState: string
  contentRef?: string
  approvedBy?: string
  approvedAt?: string
  createdAt: string
  updatedAt: string
}

/**
 * Dad's substantive reply workflow. Owner-only. Full chain:
 * draft -> edit -> submit for approval -> approve -> send -> delivery status.
 * Distinct from ManualAck (initial receipt acknowledgment). Sending is
 * gated by APPROVED state + mailbox identity, never by the global
 * email_ack_enabled switch (which stays OFF).
 */
export function ResponseComposer({
  complaintId,
  isOwner,
}: {
  complaintId: string
  isOwner: boolean
}) {
  const [actions, setActions] = useState<ResponseAction[]>([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [newBody, setNewBody] = useState('')
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editBody, setEditBody] = useState('')
  const [sendingId, setSendingId] = useState<string | null>(null)

  async function load() {
    setLoading(true)
    setError(null)
    try {
      const res = await api.listResponseActions(complaintId)
      setActions(res.actions)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to load responses')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    if (isOwner) load()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [complaintId, isOwner])

  if (!isOwner) return null

  async function createDraft() {
    if (!newBody.trim()) return
    setLoading(true)
    setError(null)
    try {
      await api.createResponseDraft(complaintId, { body: newBody.trim(), responseKind: 'RESOLUTION' })
      setNewBody('')
      await load()
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to create draft')
    } finally {
      setLoading(false)
    }
  }

  async function saveEdit(id: string) {
    if (!editBody.trim()) return
    setLoading(true)
    setError(null)
    try {
      await api.updateResponseDraft(id, editBody.trim())
      setEditingId(null)
      await load()
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to save edit')
    } finally {
      setLoading(false)
    }
  }

  async function transition(id: string, action: 'submit' | 'approve' | 'reject') {
    setLoading(true)
    setError(null)
    try {
      if (action === 'submit') await api.submitResponseAction(id)
      if (action === 'approve') await api.approveResponseAction(id)
      if (action === 'reject') {
        if (!window.confirm('Reject this response? It will not be sent.')) return
        await api.rejectResponseAction(id)
      }
      await load()
    } catch (e) {
      setError(e instanceof Error ? e.message : `${action} failed`)
    } finally {
      setLoading(false)
    }
  }

  async function send(id: string, body: string) {
    if (!window.confirm(`Send this exact response to Dunkin?\n\n${body}\n\nThis cannot be undone.`)) return
    setSendingId(id)
    setLoading(true)
    setError(null)
    try {
      await api.sendResponseAction(id)
      await load()
    } catch (e) {
      const msg = e instanceof Error ? e.message : 'Send failed'
      setError(msg)
      await load()
    } finally {
      setLoading(false)
      setSendingId(null)
    }
  }

  const stateLabel = (state: string) => {
    switch (state) {
      case 'DRAFT': return 'Draft — not yet submitted'
      case 'PENDING_APPROVAL': return 'Waiting for approval'
      case 'APPROVED': return 'Approved — ready to send'
      case 'SENT': return 'Accepted by Microsoft Graph'
      case 'SENDING': return 'Sending… (do not tap again)'
      case 'REJECTED': return 'Rejected — will not be sent'
      case 'NEEDS_RECONCILIATION': return 'Uncertain delivery — check Sent Items before retrying'
      default: return state
    }
  }

  return (
    <section className="panel" id="reply-to-dunkin">
      <div className="panel-head">
        <div>
          <h2>Reply to Dunkin</h2>
          <p>
            Substantive resolution response. Draft, edit, approve, then send into the original
            Dunkin thread. Separate from the initial receipt acknowledgment above.
          </p>
        </div>
      </div>

      <div style={{ marginBottom: 16 }}>
        <label style={{ display: 'block', marginBottom: 4, fontWeight: 600 }}>
          New response draft
        </label>
        <textarea
          value={newBody}
          onChange={(e) => setNewBody(e.target.value)}
          rows={4}
          placeholder="Write the response to Dunkin…"
          style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #ccc' }}
        />
        <button className="secondary" onClick={createDraft} disabled={loading || !newBody.trim()} style={{ marginTop: 8 }}>
          {loading ? 'Working…' : 'Create draft'}
        </button>
      </div>

      {actions.length === 0 && !loading && (
        <p style={{ color: 'var(--muted, #666)' }}>No response drafts yet for this complaint.</p>
      )}

      {actions.map((a) => (
        <div key={a.id} style={{ borderTop: '1px solid #ddd', paddingTop: 12, marginTop: 12 }}>
          <p>
            <strong>{a.responseKind}</strong> · {stateLabel(a.policyState)}
            {a.approvedBy && ` · Approved by ${a.approvedBy}`}
          </p>
          {editingId === a.id ? (
            <div>
              <textarea
                value={editBody}
                onChange={(e) => setEditBody(e.target.value)}
                rows={4}
                style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #ccc' }}
              />
              <div style={{ marginTop: 8, display: 'flex', gap: 8 }}>
                <button className="primary" onClick={() => saveEdit(a.id)} disabled={loading || !editBody.trim()}>
                  Save edit
                </button>
                <button className="secondary" onClick={() => setEditingId(null)} disabled={loading}>
                  Cancel
                </button>
              </div>
            </div>
          ) : (
            <blockquote style={{ borderLeft: '3px solid #ccc', paddingLeft: 12, margin: '8px 0', whiteSpace: 'pre-wrap' }}>
              {a.contentRef ?? '(empty)'}
            </blockquote>
          )}
          {editingId !== a.id && (
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 8 }}>
              {a.policyState === 'DRAFT' && (
                <>
                  <button className="secondary" onClick={() => { setEditingId(a.id); setEditBody(a.contentRef ?? '') }} disabled={loading}>
                    Edit
                  </button>
                  <button className="secondary" onClick={() => transition(a.id, 'submit')} disabled={loading}>
                    Submit for approval
                  </button>
                </>
              )}
              {a.policyState === 'PENDING_APPROVAL' && (
                <>
                  <button className="primary" onClick={() => transition(a.id, 'approve')} disabled={loading}>
                    Approve
                  </button>
                  <button className="secondary" onClick={() => transition(a.id, 'reject')} disabled={loading}>
                    Reject
                  </button>
                </>
              )}
              {a.policyState === 'APPROVED' && (
                <button
                  className="primary"
                  onClick={() => send(a.id, a.contentRef ?? '')}
                  disabled={loading || sendingId === a.id}
                >
                  {sendingId === a.id ? 'Sending…' : 'Send to Dunkin'}
                </button>
              )}
              {a.policyState === 'SENT' && (
                <p style={{ color: 'green' }}>✓ Accepted by Microsoft Graph — reply submitted to the original thread. Graph acceptance does not prove the recipient opened it.</p>
              )}
              {a.policyState === 'NEEDS_RECONCILIATION' && (
                <p style={{ color: '#b00' }}>
                  Delivery uncertain. Check Graph Sent Items / the thread before any retry. Do not
                  tap Send again to &quot;make sure&quot;.
                </p>
              )}
            </div>
          )}
        </div>
      ))}

      {error && <p style={{ color: '#b00', marginTop: 8 }}>{error}</p>}
    </section>
  )
}
