import { useState } from 'react'
import { api } from '../lib/api'

type PreviewItem = Awaited<ReturnType<typeof api.previewMail>>['items'][number]

const CLASS_LABEL: Record<PreviewItem['classification'], string> = {
  COMPLAINT_CANDIDATE: 'Complaint candidate',
  FOLLOW_UP_LIKELY: 'Follow-up likely',
  NEEDS_REVIEW: 'Needs review',
  NOT_COMPLAINT: 'Not complaint',
  ALREADY_INGESTED: 'Already ingested',
}

/** Owner-only Check Mail Now: preview first, import only explicitly selected messages. */
export function CheckMailNow() {
  const [startDate, setStartDate] = useState('2026-10-10')
  const [endDate, setEndDate] = useState('2026-10-17')
  const [items, setItems] = useState<PreviewItem[]>([])
  const [nextLink, setNextLink] = useState<string | null>(null)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [previewing, setPreviewing] = useState(false)
  const [importing, setImporting] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [importResults, setImportResults] = useState<
    Awaited<ReturnType<typeof api.importPreviewMail>> | null
  >(null)

  const windowISO = () => ({
    start: `${startDate}T00:00:00.000Z`,
    end: `${endDate}T23:59:59.999Z`,
  })

  const loadPreview = async (cursor?: string) => {
    setError(null)
    setImportResults(null)
    setPreviewing(true)
    try {
      const { start, end } = windowISO()
      const result = await api.previewMail({ start, end, top: 50, nextLink: cursor })
      if (cursor) {
        setItems((prev) => [...prev, ...result.items])
      } else {
        setItems(result.items)
        setSelected(new Set())
      }
      setNextLink(result.nextLink)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Preview failed')
    } finally {
      setPreviewing(false)
    }
  }

  const toggle = (id: string) => {
    setSelected((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }

  const importSelected = async () => {
    const ids = [...selected].filter((id) => {
      const item = items.find((i) => i.graphMessageId === id)
      return item && !item.alreadyIngested
    })
    if (ids.length === 0) return
    const confirmed = window.confirm(
      `Import ${ids.length} selected message(s) for ${startDate} to ${endDate}? ` +
        'This creates complaints / review items from the selected mail only. ' +
        'No notifications or acknowledgments will be sent.',
    )
    if (!confirmed) return
    setError(null)
    setImporting(true)
    try {
      const { start, end } = windowISO()
      const result = await api.importPreviewMail({
        messageIds: ids,
        receivedStart: start,
        receivedEnd: end,
        ingestionMode: 'BACKFILL',
      })
      setImportResults(result)
      setSelected(new Set())
      // Refresh preview so ALREADY_INGESTED states update.
      await loadPreview()
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Import failed')
    } finally {
      setImporting(false)
    }
  }

  return (
    <section className="panel settings-card">
      <h2>Check Mail Now</h2>
      <p>
        Preview recent Inbox mail first. Preview never creates complaints, never advances the
        ingestion cursor, and never sends notifications or acknowledgments. Import only the
        messages you explicitly select.
      </p>
      <div className="form-row">
        <label>
          From
          <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
        </label>
        <label>
          To
          <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
        </label>
        <button className="secondary" disabled={previewing} onClick={() => loadPreview()}>
          {previewing ? 'Previewing…' : 'Preview mail'}
        </button>
      </div>
      {error && <div className="danger-note">{error}</div>}
      {items.length > 0 && (
        <>
          <p>
            {items.length} message(s) in window. Select complaint candidates and needs-review
            items to import. Already-ingested messages cannot be re-imported.
          </p>
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Import</th>
                  <th>Received</th>
                  <th>Sender</th>
                  <th>Subject</th>
                  <th>Case ID</th>
                  <th>Store</th>
                  <th>Classification</th>
                </tr>
              </thead>
              <tbody>
                {items.map((item) => (
                  <tr key={item.graphMessageId}>
                    <td>
                      <input
                        type="checkbox"
                        disabled={item.alreadyIngested}
                        checked={selected.has(item.graphMessageId)}
                        onChange={() => toggle(item.graphMessageId)}
                        aria-label={`Select ${item.maskedId}`}
                      />
                    </td>
                    <td>{new Date(item.receivedAt).toLocaleString()}</td>
                    <td title={item.maskedId}>{item.senderDisplay}</td>
                    <td title={item.classificationReason}>{item.subject}</td>
                    <td>{item.detectedCaseId ?? '—'}</td>
                    <td>{item.detectedStoreNumber ?? '—'}</td>
                    <td>
                      {CLASS_LABEL[item.classification]}
                      {item.isForwardOrReply ? ' · Fwd/Re' : ''}
                      {item.linkedComplaintId ? ' · linked' : ''}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="form-row">
            {nextLink && (
              <button className="secondary" disabled={previewing} onClick={() => loadPreview(nextLink)}>
                Load more
              </button>
            )}
            <button
              className="primary"
              disabled={importing || selected.size === 0}
              onClick={importSelected}
            >
              {importing ? 'Importing…' : `Import selected (${selected.size})`}
            </button>
          </div>
          <small>
            IDs are masked in this view. Subjects may contain customer details — owner view only.
            Import uses BACKFILL mode for past windows so overdue deadlines are not applied to
            recovered mail.
          </small>
        </>
      )}
      {importResults && (
        <div className="readiness-list">
          <span>
            Imported <strong>{importResults.imported}</strong>, failed{' '}
            <strong>{importResults.failed}</strong>
          </span>
          {importResults.results.map((r) => (
            <span key={r.graphMessageId}>
              {r.maskedId}: {r.status}
              {r.complaintId ? ` → ${r.complaintId}` : ''}
              {r.errorCode ? ` (${r.errorCode})` : ''}
            </span>
          ))}
        </div>
      )}
    </section>
  )
}
