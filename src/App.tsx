import { useEffect, useMemo, useState } from 'react'
import {
  AlertTriangle,
  ArrowLeft,
  Bell,
  Check,
  ChevronRight,
  CircleGauge,
  Clock3,
  FileText,
  Inbox,
  LayoutDashboard,
  Menu,
  Plus,
  Search,
  Settings,
  ShieldCheck,
  Store as StoreIcon,
  Users,
} from 'lucide-react'
import { metrics } from './lib/workflow'
import { api } from './lib/api'
import { parseComplaintHash } from './lib/deep-links'
import {
  getPushStatus,
  isStandalone,
  subscribeForPush,
  unsubscribeLocalPush,
  type PushStatus,
} from './lib/push'
import ReviewQueue from './ReviewQueue'
import { CheckMailNow } from './components/CheckMailNow'
import { fixtures, initialState } from './lib/seed'
import type { AppState, Complaint, IntegrationHealth, NewComplaint, Severity } from './lib/types'

type View =
  | 'dashboard'
  | 'complaints'
  | 'simulator'
  | 'detail'
  | 'stores'
  | 'team'
  | 'reports'
  | 'settings'
  | 'review'
const fmt = (value?: string) =>
  value
    ? new Intl.DateTimeFormat('en-US', {
        month: 'short',
        day: 'numeric',
        hour: 'numeric',
        minute: '2-digit',
      }).format(new Date(value))
    : '—'
const cx = (...x: (string | false | undefined)[]) => x.filter(Boolean).join(' ')

// Deep links: `#/complaints/:id` opens the complaint detail view. Used by Web
// Push notification taps (`notificationclick` → `clients.openWindow(url)`) and
// shareable links. The hash keeps the URL stable behind Cloudflare Access and
// requires no router changes.
const syncHashForView = (next: string, id?: string) => {
  if (next === 'detail' && id) {
    const target = `#/complaints/${id}`
    if (window.location.hash !== target) window.location.hash = target
  } else if (window.location.hash.startsWith('#/complaints/')) {
    history.replaceState(null, '', window.location.pathname + window.location.search)
  }
}

function Badge({ children, tone = 'neutral' }: { children: React.ReactNode; tone?: string }) {
  return <span className={`badge ${tone}`}>{children}</span>
}
function Avatar({ name }: { name: string }) {
  return (
    <span className="avatar">
      {name
        .split(' ')
        .map((x) => x[0])
        .join('')
        .slice(0, 2)}
    </span>
  )
}

export default function App() {
  const [state, setState] = useState<AppState>(initialState)
  const [loading, setLoading] = useState(true)
  const [providerReady, setProviderReady] = useState(false)
  const [health, setHealth] = useState<IntegrationHealth>()
  const [view, setView] = useState<View>('dashboard')
  const [selectedId, setSelectedId] = useState<string>()
  const [toast, setToast] = useState<string>()
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const user = state.users.find((u) => u.id === state.activeUserId)!
  const selected = state.complaints.find((c) => c.id === selectedId)
  const show = (message: string) => {
    setToast(message)
    window.setTimeout(() => setToast(undefined), 2800)
  }
  const navigate = (next: View, id?: string) => {
    setView(next)
    setSelectedId(id)
    setSidebarOpen(false)
    window.scrollTo(0, 0)
    syncHashForView(next, id)
  }
  const visibleComplaints =
    user.role === 'STORE_MANAGER'
      ? state.complaints.filter((c) => c.assignedManagerId === user.id)
      : state.complaints
  useEffect(() => {
    api
      .bootstrap()
      .then((result) => {
        setState(result.state)
        setProviderReady(result.providerReady)
        setHealth(result.health)
      })
      .catch((error: Error) => show(error.message))
      .finally(() => setLoading(false))
  }, [])
  useEffect(() => {
    if (loading) return
    const timer = window.setInterval(
      () =>
        api
          .bootstrap()
          .then((result) => {
            setState(result.state)
            setProviderReady(result.providerReady)
            setHealth(result.health)
          })
          .catch(() => undefined),
      10_000,
    )
    return () => window.clearInterval(timer)
  }, [loading])
  // Deep-link handling: open `#/complaints/:id` on cold start (e.g. a push
  // notification tap) and on in-app hash changes. Idempotent — re-runs are
  // no-ops when the requested complaint is already open.
  useEffect(() => {
    if (loading) return
    const openFromHash = () => {
      const id = parseComplaintHash(window.location.hash)
      if (!id) return
      const allowed =
        user.role === 'STORE_MANAGER'
          ? state.complaints.some((c) => c.id === id && c.assignedManagerId === user.id)
          : state.complaints.some((c) => c.id === id)
      if (!allowed) {
        show('Complaint not found or not available to you')
        return
      }
      if (view === 'detail' && selectedId === id) return
      setView('detail')
      setSelectedId(id)
      setSidebarOpen(false)
      window.scrollTo(0, 0)
    }
    openFromHash()
    window.addEventListener('hashchange', openFromHash)
    return () => window.removeEventListener('hashchange', openFromHash)
  }, [loading, state.complaints, user, view, selectedId])
  if (loading)
    return (
      <div className="loading-screen">
        <ShieldCheck />
        <strong>Loading secure workspace…</strong>
      </div>
    )
  return (
    <div className="app-shell">
      {sidebarOpen && (
        <div className="sidebar-backdrop" onClick={() => setSidebarOpen(false)} />
      )}
      <aside className={cx('sidebar', sidebarOpen && 'open')}>
        <div className="brand">
          <span className="brand-mark">
            <Check size={18} />
          </span>
          <span>StoreResolve</span>
        </div>
        <nav>
          <button
            className={cx(view === 'dashboard' && 'active')}
            onClick={() => navigate('dashboard')}
          >
            <LayoutDashboard />
            Overview
          </button>
          <button
            className={cx((view === 'complaints' || view === 'detail') && 'active')}
            onClick={() => navigate('complaints')}
          >
            <FileText />
            Complaints <span className="nav-count">{visibleComplaints.length}</span>
          </button>
          {(user.role === 'OWNER' || user.role === 'ADMIN') && (
            <button
              className={cx(view === 'simulator' && 'active')}
              onClick={() => navigate('simulator')}
            >
              <Plus />
              Simulator
            </button>
          )}
          <div className="nav-label">Operations</div>
          <button className={cx(view === 'review' && 'active')} onClick={() => navigate('review')}>
            <Inbox />
            Review queue
          </button>
          <button className={cx(view === 'stores' && 'active')} onClick={() => navigate('stores')}>
            <StoreIcon />
            Stores
          </button>
          <button className={cx(view === 'team' && 'active')} onClick={() => navigate('team')}>
            <Users />
            Team
          </button>
          <button
            className={cx(view === 'reports' && 'active')}
            onClick={() => navigate('reports')}
          >
            <CircleGauge />
            Reports
          </button>
          {(user.role === 'OWNER' || user.role === 'ADMIN') && (
            <>
              <div className="nav-label">Administration</div>
              <button
                className={cx(view === 'settings' && 'active')}
                onClick={() => navigate('settings')}
              >
                <Settings />
                Pilot controls
              </button>
            </>
          )}
        </nav>
        <div className="safety-card">
          <ShieldCheck />
          <div>
            <strong>
              {state.config.externalNotificationsEnabled
                ? 'External messages eligible'
                : 'External messages off'}
            </strong>
            <span>
              Microsoft mail intake {state.config.emailIngestionEnabled ? 'on' : 'off'} · replies{' '}
              {state.config.emailAckEnabled ? 'on' : 'off'}
            </span>
          </div>
        </div>
        <div className="profile">
          <Avatar name={user.name} />
          <div>
            <strong>{user.name}</strong>
            <span>{user.role.replace('_', ' ')}</span>
          </div>
        </div>
      </aside>
      <main>
        <header>
          <button
            className="mobile-menu"
            onClick={() => setSidebarOpen(true)}
            aria-label="Open menu"
          >
            <Menu />
          </button>
          <div className="search">
            <Search />
            <input placeholder="Search cases, stores, managers…" />
          </div>
          <div className="header-actions">
            <Badge tone="safe">{state.config.mode.replaceAll('_', ' ')}</Badge>
            <button className="icon-button">
              <Bell />
              <i />
            </button>
            <span className="session-user">
              <ShieldCheck /> Signed in as {user.name}
            </span>
          </div>
        </header>
        <div className="page">
          {view === 'dashboard' && (
            <Dashboard
              state={{ ...state, complaints: visibleComplaints }}
              onOpen={(id) => navigate('detail', id)}
              onCreate={() => navigate('simulator')}
            />
          )}
          {view === 'review' && (
            <ReviewQueue
              users={state.users}
              isAdmin={user.role === 'OWNER' || user.role === 'ADMIN'}
              show={show}
            />
          )}
          {view === 'complaints' && (
            <ComplaintList
              state={{ ...state, complaints: visibleComplaints }}
              onOpen={(id) => navigate('detail', id)}
            />
          )}
          {view === 'stores' && (
            <StoresView
              state={state}
              canEdit={user.role === 'OWNER' || user.role === 'ADMIN'}
              onSave={(store) =>
                api
                  .updateStore(store.id, store)
                  .then((next) => {
                    setState(next)
                    show('Store configuration saved')
                  })
                  .catch((error: Error) => show(error.message))
              }
            />
          )}
          {view === 'team' && (
            <TeamView
              state={state}
              canEdit={user.role === 'OWNER' || user.role === 'ADMIN'}
              onSave={(id, changes) =>
                api
                  .updateUser(id, changes)
                  .then((next) => {
                    setState(next)
                    show('User configuration saved')
                  })
                  .catch((error: Error) => show(error.message))
              }
            />
          )}
          {view === 'reports' && (
            <ReportsView state={{ ...state, complaints: visibleComplaints }} />
          )}
          {view === 'simulator' && (
            <Simulator
              state={state}
              onCreate={async (input) => {
                try {
                  const next = await api.createComplaint(input)
                  setState(next)
                  const created = next.complaints.find(
                    (c) => c.externalCaseId === input.externalCaseId,
                  )
                  show('Complaint persisted and workflow started')
                  if (created) navigate('detail', created.id)
                } catch (error) {
                  show((error as Error).message)
                }
              }}
            />
          )}
          {view === 'detail' && selected && (
            <ComplaintDetail
              complaint={selected}
              state={state}
              onBack={() => navigate('complaints')}
              onAct={(action, data) =>
                api
                  .act(selected.id, action, data)
                  .then((next) => {
                    setState(next)
                    show('Case update persisted')
                  })
                  .catch((error: Error) => show(error.message))
              }
            />
          )}
          {view === 'settings' && (
            <>
              <PushNotificationSettings show={show} />
              <PilotControls
              state={state}
              setState={setState}
              providerReady={providerReady}
              health={health}
              onRun={() =>
                api
                  .processDeadlines(72)
                  .then((next) => {
                    setState(next)
                    show('Persistent deadlines processed')
                  })
                  .catch((error: Error) => show(error.message))
              }
              onSaveConfig={(config) =>
                api
                  .updateConfig(config)
                  .then((next) => {
                    setState(next)
                    show('Safety configuration saved')
                  })
                  .catch((error: Error) => show(error.message))
              }
              onSendTest={(id) =>
                api
                  .sendTest(id)
                  .then((next) => {
                    setState(next)
                    show('Test attempt recorded')
                  })
                  .catch((error: Error) => show(error.message))
              }
              onReconcile={async (providerMessageId) => {
                try {
                  const result = await api.reconcileSignalWire(providerMessageId)
                  const refreshed = await api.bootstrap()
                  setState(refreshed.state)
                  setProviderReady(refreshed.providerReady)
                  show(`Provider reconciliation ${result.result.toLowerCase()}`)
                } catch (error) {
                  show((error as Error).message)
                }
              }}
              onPilotIngest={async () => {
                try {
                  const result = await api.ingestPilotEmail()
                  const refreshed = await api.bootstrap()
                  setState(refreshed.state)
                  setProviderReady(refreshed.providerReady)
                  show(
                    result.accessed
                      ? `One-shot mail ingestion ${result.status?.toLowerCase() ?? 'completed'}`
                      : `Controlled selector matched ${result.matchCount ?? 0} messages; no content read`,
                  )
                } catch (error) {
                  show((error as Error).message)
                }
              }}
              onSendTestPush={(id) =>
                api
                  .sendTestPush(id)
                  .then((next) => {
                    setState(next)
                    show('Push test attempt recorded')
                  })
                  .catch((error: Error) => show(error.message))
              }
            />
            </>
          )}
        </div>
      </main>
      {toast && (
        <div className="toast">
          <Check />
          {toast}
        </div>
      )}
    </div>
  )
}

function PageTitle({
  eyebrow,
  title,
  body,
  action,
}: {
  eyebrow?: string
  title: string
  body: string
  action?: React.ReactNode
}) {
  return (
    <div className="page-title">
      <div>
        {eyebrow && <span className="eyebrow">{eyebrow}</span>}
        <h1>{title}</h1>
        <p>{body}</p>
      </div>
      {action}
    </div>
  )
}

function Dashboard({
  state,
  onOpen,
  onCreate,
}: {
  state: AppState
  onOpen: (id: string) => void
  onCreate: () => void
}) {
  const m = metrics(state)
  const today = state.complaints.filter(
    (c) => new Date(c.receivedAt).toDateString() === new Date().toDateString(),
  ).length
  const me = state.users.find((u) => u.id === state.activeUserId)
  const hour = new Date().getHours()
  const daypart = hour < 12 ? 'morning' : hour < 18 ? 'afternoon' : 'evening'
  const activeStores = state.stores.filter((s) => s.active).length
  const stats = [
    {
      label: 'Open complaints',
      value: m.open,
      note: `${m.awaitingAck} awaiting acknowledgment`,
      icon: <FileText />,
      tone: 'navy',
    },
    {
      label: 'Overdue',
      value: m.overdue,
      note: m.overdue ? 'Requires attention' : 'No missed deadlines',
      icon: <AlertTriangle />,
      tone: 'red',
    },
    {
      label: 'Complaints today',
      value: today,
      note: `${state.complaints.length} total recorded`,
      icon: <Clock3 />,
      tone: 'orange',
    },
    {
      label: 'Closed',
      value: m.closed,
      note: 'Ownership reviewed',
      icon: <Check />,
      tone: 'green',
    },
  ]
  return (
    <>
      <PageTitle
        eyebrow="Owner workspace"
        title={`Good ${daypart}, ${me?.name ?? 'there'}`}
        body={`Here’s what needs attention across your ${activeStores} stores.`}
        action={
          <button
            className="primary"
            onClick={onCreate}
            aria-label="New simulated complaint"
            title="New simulated complaint"
          >
            <Plus />
            New simulated complaint
          </button>
        }
      />
      <section className="stat-grid">
        {stats.map((s) => (
          <div className="stat-card" key={s.label}>
            <span className={`stat-icon ${s.tone}`}>{s.icon}</span>
            <div>
              <span>{s.label}</span>
              <strong>{s.value}</strong>
              <small>{s.note}</small>
            </div>
          </div>
        ))}
      </section>
      <section className="split">
        <div className="panel">
          <div className="panel-head">
            <div>
              <h2>Active complaints</h2>
              <p>Cases currently moving through the resolution workflow.</p>
            </div>
            <button className="text-button">
              View all <ChevronRight />
            </button>
          </div>
          <ComplaintTable state={state} onOpen={onOpen} />
        </div>
        <div className="panel attention">
          <div className="panel-head">
            <div>
              <h2>Needs attention</h2>
              <p>Deadlines and routing exceptions.</p>
            </div>
          </div>
          {m.overdue === 0 && !state.complaints.some((c) => !c.storeId) ? (
            <Empty
              icon={<ShieldCheck />}
              title="Everything is on track"
              body="No overdue or unassigned cases."
            />
          ) : (
            state.complaints
              .filter((c) => c.isAckOverdue || c.isResolutionOverdue || !c.storeId)
              .map((c) => (
                <button className="attention-row" onClick={() => onOpen(c.id)} key={c.id}>
                  <span className="alert-icon">
                    <AlertTriangle />
                  </span>
                  <div>
                    <strong>
                      {!c.storeId
                        ? 'Routing review'
                        : c.isAckOverdue
                          ? 'Manager acknowledgment overdue'
                          : 'Resolution overdue'}
                    </strong>
                    <span>
                      {c.id} · {fmt(c.receivedAt)}
                    </span>
                  </div>
                  <ChevronRight />
                </button>
              ))
          )}
        </div>
      </section>
      <section className="metrics-strip">
        <div>
          <span>Avg. acknowledgment</span>
          <strong>
            {m.avgAckMinutes || '—'}
            <small>{m.avgAckMinutes ? ' min' : ''}</small>
          </strong>
        </div>
        <div>
          <span>Avg. resolution</span>
          <strong>
            {m.avgResolutionHours || '—'}
            <small>{m.avgResolutionHours ? ' hr' : ''}</small>
          </strong>
        </div>
        <div>
          <span>Acknowledgment SLA</span>
          <strong>
            {state.complaints.length
              ? Math.round(
                  (state.complaints.filter((c) => !c.isAckOverdue).length /
                    state.complaints.length) *
                    100,
                )
              : 100}
            <small>%</small>
          </strong>
        </div>
        <div>
          <span>Resolution SLA</span>
          <strong>
            {state.complaints.length
              ? Math.round(
                  (state.complaints.filter((c) => !c.isResolutionOverdue).length /
                    state.complaints.length) *
                    100,
                )
              : 100}
            <small>%</small>
          </strong>
        </div>
      </section>
    </>
  )
}

function Empty({ icon, title, body }: { icon: React.ReactNode; title: string; body: string }) {
  return (
    <div className="empty">
      {icon}
      <strong>{title}</strong>
      <span>{body}</span>
    </div>
  )
}
function ComplaintTable({ state, onOpen }: { state: AppState; onOpen: (id: string) => void }) {
  if (!state.complaints.length)
    return (
      <Empty
        icon={<FileText />}
        title="No complaints yet"
        body="Use the simulator to create your first case."
      />
    )
  return (
    <div className="table-wrap">
      <table className="complaint-table">
        <thead>
          <tr>
            <th>Case</th>
            <th>Store</th>
            <th>Category</th>
            <th>Severity</th>
            <th className="hide-mobile">Manager</th>
            <th>Status</th>
            <th>Received</th>
            <th className="hide-mobile"></th>
          </tr>
        </thead>
        <tbody>
          {state.complaints.slice(0, 8).map((c) => {
            const store = state.stores.find((s) => s.id === c.storeId)
            const manager = state.users.find((u) => u.id === c.assignedManagerId)
            return (
              <tr key={c.id} onClick={() => onOpen(c.id)}>
                <td>
                  <strong>{c.id}</strong>
                  <span>{c.externalCaseId}</span>
                </td>
                <td>{store ? `#${store.number}` : <Badge tone="warning">Unassigned</Badge>}</td>
                <td>{c.category}</td>
                <td>
                  <Badge tone={c.severity.toLowerCase()}>{c.severity}</Badge>
                </td>
                <td className="hide-mobile">{manager?.name ?? '—'}</td>
                <td>
                  <Badge tone={c.isAckOverdue || c.isResolutionOverdue ? 'danger' : 'status'}>
                    {c.isAckOverdue || c.isResolutionOverdue
                      ? 'OVERDUE'
                      : c.status.replaceAll('_', ' ')}
                  </Badge>
                </td>
                <td>{fmt(c.receivedAt)}</td>
                <td className="hide-mobile">
                  <ChevronRight />
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

function ComplaintList({ state, onOpen }: { state: AppState; onOpen: (id: string) => void }) {
  const [query, setQuery] = useState('')
  const [severity, setSeverity] = useState('ALL')
  const list = useMemo(
    () =>
      state.complaints.filter(
        (c) =>
          (severity === 'ALL' || c.severity === severity) &&
          (c.id + c.subject + c.category).toLowerCase().includes(query.toLowerCase()),
      ),
    [state.complaints, query, severity],
  )
  return (
    <>
      <PageTitle
        eyebrow="Operations"
        title="Complaints"
        body="Track every case from receipt through ownership review."
      />
      <div className="panel">
        <div className="filters">
          <label>
            <Search />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search complaints"
            />
          </label>
          <select value={severity} onChange={(e) => setSeverity(e.target.value)}>
            <option value="ALL">All severities</option>
            {['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'].map((x) => (
              <option key={x}>{x}</option>
            ))}
          </select>
          <span>{list.length} cases</span>
        </div>
        <ComplaintTable state={{ ...state, complaints: list }} onOpen={onOpen} />
      </div>
    </>
  )
}

function Simulator({
  state,
  onCreate,
}: {
  state: AppState
  onCreate: (input: NewComplaint) => void
}) {
  const [form, setForm] = useState<NewComplaint>({
    externalCaseId: `DKN-${String(state.complaints.length + 1).padStart(3, '0')}`,
    storeNumber: '350909',
    subject: 'Service concern',
    complaintText: 'Customer reports an unhelpful interaction at the counter.',
    category: 'Customer Service',
    severity: 'MEDIUM',
  })
  const load = (f: (typeof fixtures)[number]) =>
    setForm({
      ...form,
      ...f,
      externalCaseId: `DKN-${String(Date.now()).slice(-6)}`,
    } as NewComplaint)
  return (
    <>
      <PageTitle
        eyebrow="Development tool"
        title="Complaint simulator"
        body="Run a realistic complaint through the same domain workflow used by Microsoft Graph ingestion."
      />
      <div className="sim-layout">
        <form
          className="panel form-card"
          onSubmit={(e) => {
            e.preventDefault()
            onCreate(form)
          }}
        >
          <div className="form-section">
            <h2>Incoming Dunkin complaint</h2>
            <p>All fields use fictional development data.</p>
          </div>
          <div className="form-grid">
            <label>
              External case ID
              <input
                required
                value={form.externalCaseId}
                onChange={(e) => setForm({ ...form, externalCaseId: e.target.value })}
              />
            </label>
            <label>
              Store number
              <input
                required
                value={form.storeNumber}
                onChange={(e) => setForm({ ...form, storeNumber: e.target.value })}
              />
            </label>
            <label className="full">
              Subject
              <input
                required
                value={form.subject}
                onChange={(e) => setForm({ ...form, subject: e.target.value })}
              />
            </label>
            <label>
              Category
              <select
                value={form.category}
                onChange={(e) => setForm({ ...form, category: e.target.value })}
              >
                {[
                  'Customer Service',
                  'Order Accuracy',
                  'Cleanliness',
                  'Employee Conduct',
                  'Safety',
                ].map((x) => (
                  <option key={x}>{x}</option>
                ))}
              </select>
            </label>
            <label>
              Severity
              <select
                value={form.severity}
                onChange={(e) => setForm({ ...form, severity: e.target.value as Severity })}
              >
                {['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'].map((x) => (
                  <option key={x}>{x}</option>
                ))}
              </select>
            </label>
            <label className="full">
              Complaint text
              <textarea
                required
                rows={7}
                value={form.complaintText}
                onChange={(e) => setForm({ ...form, complaintText: e.target.value })}
              />
            </label>
          </div>
          <div className="form-actions">
            <span>
              <ShieldCheck /> No external messages will be sent
            </span>
            <button className="primary" type="submit">
              Process complaint <ChevronRight />
            </button>
          </div>
        </form>
        <aside className="fixture-panel">
          <h3>Quick fixtures</h3>
          <p>Load a prepared scenario.</p>
          {fixtures.map((f) => (
            <button key={f.label} onClick={() => load(f)}>
              <span>{f.label}</span>
              <small>
                {f.category} · {f.severity}
              </small>
              <ChevronRight />
            </button>
          ))}
        </aside>
      </div>
    </>
  )
}

function ComplaintDetail({
  complaint: c,
  state,
  onBack,
  onAct,
}: {
  complaint: Complaint
  state: AppState
  onBack: () => void
  onAct: (
    action:
      | 'ACKNOWLEDGE'
      | 'START_INVESTIGATION'
      | 'CONTACT_CUSTOMER'
      | 'SUBMIT_RESOLUTION'
      | 'CLOSE'
      | 'REOPEN',
    data?: Record<string, unknown>,
  ) => void
}) {
  const store = state.stores.find((s) => s.id === c.storeId)
  const manager = state.users.find((u) => u.id === c.assignedManagerId)
  const current = state.users.find((u) => u.id === state.activeUserId)!
  const isManager = current.id === c.assignedManagerId
  const canClose =
    (current.role === 'OWNER' || current.role === 'ADMIN') && c.status === 'RESOLUTION_SUBMITTED'
  const [resolution, setResolution] = useState({
    findings: 'Reviewed the complaint details and spoke with the shift team.',
    correctiveAction: 'Coached the team and reviewed the applicable service standard.',
    resolutionNotes: 'Customer concern addressed and store follow-up completed.',
  })
  return (
    <>
      <button className="back" onClick={onBack}>
        <ArrowLeft />
        All complaints
      </button>
      <div className="detail-title">
        <div>
          <div className="title-badges">
            <Badge tone={c.severity.toLowerCase()}>{c.severity}</Badge>
            <Badge tone={c.isAckOverdue || c.isResolutionOverdue ? 'danger' : 'status'}>
              {c.isAckOverdue || c.isResolutionOverdue ? 'OVERDUE' : c.status.replaceAll('_', ' ')}
            </Badge>
          </div>
          <h1>{c.subject}</h1>
          <p>
            {c.id} · External case {c.externalCaseId}
          </p>
        </div>
        {canClose && (
          <button className="primary" onClick={() => onAct('CLOSE')}>
            <Check />
            Close complaint
          </button>
        )}
      </div>
      <div className="detail-grid">
        <div className="detail-main">
          <section className="panel complaint-copy">
            <span>Customer complaint</span>
            <p>{c.complaintText}</p>
          </section>
          {!c.storeId && (
            <section className="routing-alert">
              <AlertTriangle />
              <div>
                <strong>Manual routing required</strong>
                <span>
                  {c.routingReason}. Ownership should assign this case before manager workflow
                  begins.
                </span>
              </div>
            </section>
          )}
          <section className="panel">
            <div className="panel-head">
              <div>
                <h2>Case activity</h2>
                <p>A complete, immutable history of this complaint.</p>
              </div>
            </div>
            <div className="timeline">
              {[...c.events].reverse().map((e, i) => (
                <div className="timeline-row" key={e.id}>
                  <span className={i === 0 ? 'current' : ''}>
                    {e.type === 'COMPLAINT_CLOSED' ? <Check /> : <Clock3 />}
                  </span>
                  <div>
                    <strong>
                      {e.type
                        .replaceAll('_', ' ')
                        .toLowerCase()
                        .replace(/^./, (x) => x.toUpperCase())}
                    </strong>
                    <p>
                      {e.actor} · {fmt(e.timestamp)}
                    </p>
                  </div>
                </div>
              ))}
            </div>
          </section>
          {isManager && c.status !== 'CLOSED' && (
            <section className="panel action-card">
              <div className="panel-head">
                <div>
                  <h2>Manager actions</h2>
                  <p>Each step requires an explicit action.</p>
                </div>
              </div>
              {!c.managerAcknowledgedAt ? (
                <button className="primary wide" onClick={() => onAct('ACKNOWLEDGE')}>
                  Acknowledge complaint
                </button>
              ) : !c.investigationStartedAt ? (
                <button className="primary wide" onClick={() => onAct('START_INVESTIGATION')}>
                  Start investigation
                </button>
              ) : !c.customerContacted ? (
                <button
                  className="secondary wide"
                  onClick={() =>
                    onAct('CONTACT_CUSTOMER', {
                      outcome: 'Customer contacted; follow-up discussed.',
                    })
                  }
                >
                  Record customer contact
                </button>
              ) : !c.resolutionSubmittedAt ? (
                <div className="resolution-form">
                  <label>
                    Manager findings
                    <textarea
                      value={resolution.findings}
                      onChange={(e) => setResolution({ ...resolution, findings: e.target.value })}
                    />
                  </label>
                  <label>
                    Corrective action
                    <textarea
                      value={resolution.correctiveAction}
                      onChange={(e) =>
                        setResolution({ ...resolution, correctiveAction: e.target.value })
                      }
                    />
                  </label>
                  <label>
                    Resolution notes
                    <textarea
                      value={resolution.resolutionNotes}
                      onChange={(e) =>
                        setResolution({ ...resolution, resolutionNotes: e.target.value })
                      }
                    />
                  </label>
                  <button
                    className="primary"
                    onClick={() => onAct('SUBMIT_RESOLUTION', resolution)}
                  >
                    Submit resolution
                  </button>
                </div>
              ) : (
                <div className="success-box">
                  <Check />
                  Resolution submitted for ownership review.
                </div>
              )}
            </section>
          )}
        </div>
        <aside className="detail-side">
          <section className="panel summary">
            <h3>Case summary</h3>
            <dl>
              <div>
                <dt>Store</dt>
                <dd>{store ? `#${store.number} · ${store.name}` : 'Unassigned'}</dd>
              </div>
              <div>
                <dt>Manager</dt>
                <dd>{manager?.name ?? '—'}</dd>
              </div>
              <div>
                <dt>Category</dt>
                <dd>{c.category}</dd>
              </div>
              <div>
                <dt>Received</dt>
                <dd>{fmt(c.receivedAt)}</dd>
              </div>
              <div>
                <dt>Acknowledgment due</dt>
                <dd className={c.isAckOverdue ? 'red-text' : ''}>{fmt(c.ackDeadline)}</dd>
              </div>
              <div>
                <dt>Resolution due</dt>
                <dd>{fmt(c.resolutionDeadline)}</dd>
              </div>
              <div>
                <dt>Routing</dt>
                <dd>{c.routingReason}</dd>
              </div>
            </dl>
          </section>
          <section className="panel summary">
            <h3>Workflow timestamps</h3>
            <dl>
              {[
                ['Dunkin acknowledged', c.dunkinAcknowledgedAt],
                ['Manager notified', c.managerNotifiedAt],
                ['Manager acknowledged', c.managerAcknowledgedAt],
                ['Investigation started', c.investigationStartedAt],
                ['Resolution submitted', c.resolutionSubmittedAt],
                ['Closed', c.closedAt],
              ].map(([k, v]) => (
                <div key={k}>
                  <dt>{k}</dt>
                  <dd>{fmt(v)}</dd>
                </div>
              ))}
            </dl>
          </section>
          {state.users.filter(
            (u) => u.id === c.assignedManagerId || c.notifications.some((n) => n.recipientUserId === u.id),
          ).length > 0 && (
            <section className="panel notifications">
              <h3>Notification delivery</h3>
              {state.users
                .filter(
                  (u) =>
                    u.id === c.assignedManagerId ||
                    c.notifications.some((n) => n.recipientUserId === u.id),
                )
                .map((u) => {
                const n = [...c.notifications].reverse().find((x) => x.recipientUserId === u.id)
                return (
                  <div key={u.id}>
                    <Avatar name={u.name} />
                    <span>
                      <strong>{u.name}</strong>
                      <small>{n?.failureReason ?? n?.provider ?? 'Not created'}</small>
                    </span>
                    <Badge
                      tone={
                        n?.status === 'FAILED'
                          ? 'danger'
                          : n?.status === 'SUPPRESSED'
                            ? 'warning'
                            : 'safe'
                      }
                    >
                      {n?.status ?? 'NONE'}
                    </Badge>
                  </div>
                )
              })}
            </section>
          )}
        </aside>
      </div>
    </>
  )
}

function StoresView({
  state,
  canEdit,
  onSave,
}: {
  state: AppState
  canEdit: boolean
  onSave: (store: AppState['stores'][number]) => void
}) {
  return (
    <>
      <PageTitle
        eyebrow="Operations"
        title="Stores"
        body="Authoritative routing locations, aliases, and manager assignments."
      />
      <div className="settings-grid">
        {state.stores.map((store) => (
          <StoreEditor
            key={store.id}
            store={store}
            managers={state.users.filter((candidate) => candidate.role === 'STORE_MANAGER')}
            canEdit={canEdit}
            onSave={onSave}
          />
        ))}
      </div>
    </>
  )
}

function StoreEditor({
  store,
  managers,
  canEdit,
  onSave,
}: {
  store: AppState['stores'][number]
  managers: AppState['users']
  canEdit: boolean
  onSave: (store: AppState['stores'][number]) => void
}) {
  const [draft, setDraft] = useState(store)
  useEffect(() => setDraft(store), [store])
  return (
    <section className="panel settings-card">
      <h2>
        #{draft.number} Â· {draft.name}
      </h2>
      <p>
        {draft.address}, {draft.city}, {draft.state} {draft.postalCode}
      </p>
      <label>
        Assigned manager
        <select
          value={draft.managerId}
          disabled={!canEdit}
          onChange={(event) => setDraft({ ...draft, managerId: event.target.value })}
        >
          <option value="">Unassigned</option>
          {managers.map((manager) => (
            <option value={manager.id} key={manager.id}>
              {manager.name}
            </option>
          ))}
        </select>
      </label>
      <label>
        Deterministic routing aliases
        <input
          disabled={!canEdit}
          value={(draft.aliases ?? []).join(', ')}
          onChange={(event) =>
            setDraft({
              ...draft,
              aliases: event.target.value
                .split(',')
                .map((alias) => alias.trim())
                .filter(Boolean),
            })
          }
        />
      </label>
      <label className="confirm-row">
        <input
          type="checkbox"
          disabled={!canEdit}
          checked={draft.active}
          onChange={(event) => setDraft({ ...draft, active: event.target.checked })}
        />{' '}
        Active for routing
      </label>
      {canEdit && (
        <button className="secondary" onClick={() => onSave(draft)}>
          Save store
        </button>
      )}
    </section>
  )
}

function TeamView({
  state,
  canEdit,
  onSave,
}: {
  state: AppState
  canEdit: boolean
  onSave: (
    id: string,
    changes: Partial<AppState['users'][number]> & { storeIds?: string[] },
  ) => void
}) {
  return (
    <>
      <PageTitle
        eyebrow="Administration"
        title="Users and authority"
        body="Access identities, roles, store assignments, and notification eligibility."
      />
      <div className="settings-grid">
        {state.users.map((member) => (
          <TeamEditor
            key={member.id}
            member={member}
            stores={state.stores}
            canEdit={canEdit && member.recipientKind !== 'PILOT_ADMIN'}
            onSave={onSave}
          />
        ))}
      </div>
    </>
  )
}

function TeamEditor({
  member,
  stores,
  canEdit,
  onSave,
}: {
  member: AppState['users'][number]
  stores: AppState['stores']
  canEdit: boolean
  onSave: (
    id: string,
    changes: Partial<AppState['users'][number]> & { storeIds?: string[] },
  ) => void
}) {
  const assigned = stores.filter((store) => store.managerId === member.id).map((store) => store.id)
  const [role, setRole] = useState(member.role)
  const [active, setActive] = useState(member.active)
  const [smsEnabled, setSmsEnabled] = useState(member.smsEnabled)
  const [complaintEligible, setComplaintEligible] = useState(
    Boolean(member.complaintNotificationsEnabled),
  )
  const [storeIds, setStoreIds] = useState(assigned)
  const [accessEmail, setAccessEmail] = useState('')
  return (
    <section className="panel settings-card">
      <h2>{member.name}</h2>
      <p>
        {member.email} Â· {member.recipientKind}
      </p>
      <label>
        Role
        <select
          value={role}
          disabled={!canEdit}
          onChange={(event) => setRole(event.target.value as typeof role)}
        >
          {['OWNER', 'ADMIN', 'VIEW_ONLY', 'STORE_MANAGER'].map((candidate) => (
            <option key={candidate}>{candidate}</option>
          ))}
        </select>
      </label>
      {role === 'STORE_MANAGER' && (
        <label>
          Assigned stores
          <select
            multiple
            value={storeIds}
            disabled={!canEdit}
            onChange={(event) =>
              setStoreIds(Array.from(event.target.selectedOptions).map((option) => option.value))
            }
          >
            {stores.map((store) => (
              <option value={store.id} key={store.id}>
                #{store.number} {store.name}
              </option>
            ))}
          </select>
        </label>
      )}
      {canEdit && (
        <label>
          Replace Access email (leave blank to preserve)
          <input
            type="email"
            value={accessEmail}
            onChange={(event) => setAccessEmail(event.target.value)}
          />
        </label>
      )}
      <label className="confirm-row">
        <input
          type="checkbox"
          checked={active}
          disabled={!canEdit}
          onChange={(event) => setActive(event.target.checked)}
        />{' '}
        Active account
      </label>
      <label className="confirm-row">
        <input
          type="checkbox"
          checked={smsEnabled}
          disabled={!canEdit}
          onChange={(event) => setSmsEnabled(event.target.checked)}
        />{' '}
        SMS enabled
      </label>
      <label className="confirm-row">
        <input
          type="checkbox"
          checked={complaintEligible}
          disabled={!canEdit}
          onChange={(event) => setComplaintEligible(event.target.checked)}
        />{' '}
        Ownership complaint notification eligibility
      </label>
      {canEdit && (
        <button
          className="secondary"
          onClick={() =>
            onSave(member.id, {
              role,
              active,
              smsEnabled,
              complaintNotificationsEnabled: complaintEligible,
              storeIds: role === 'STORE_MANAGER' ? storeIds : [],
              ...(accessEmail ? { email: accessEmail } : {}),
            })
          }
        >
          Save user
        </button>
      )}
    </section>
  )
}

function ReportsView({ state }: { state: AppState }) {
  const [storeId, setStoreId] = useState('')
  const [status, setStatus] = useState('')
  const [category, setCategory] = useState('')
  const [severity, setSeverity] = useState('')
  const filtered = state.complaints.filter(
    (complaint) =>
      (!storeId || complaint.storeId === storeId) &&
      (!status || complaint.status === status) &&
      (!category || complaint.category === category) &&
      (!severity || complaint.severity === severity),
  )
  const report = metrics({ ...state, complaints: filtered })
  const categories = Array.from(new Set(state.complaints.map((complaint) => complaint.category)))
  return (
    <>
      <PageTitle
        eyebrow="Reporting"
        title="Complaint performance"
        body="Metrics calculated from authoritative D1 complaint records."
      />
      <section className="panel settings-card">
        <div className="filters">
          <select value={storeId} onChange={(event) => setStoreId(event.target.value)}>
            <option value="">All stores</option>
            {state.stores.map((store) => (
              <option value={store.id} key={store.id}>
                #{store.number} {store.name}
              </option>
            ))}
          </select>
          <select value={status} onChange={(event) => setStatus(event.target.value)}>
            <option value="">All statuses</option>
            {Array.from(new Set(state.complaints.map((complaint) => complaint.status))).map(
              (candidate) => (
                <option key={candidate}>{candidate}</option>
              ),
            )}
          </select>
          <select value={category} onChange={(event) => setCategory(event.target.value)}>
            <option value="">All categories</option>
            {categories.map((candidate) => (
              <option key={candidate}>{candidate}</option>
            ))}
          </select>
          <select value={severity} onChange={(event) => setSeverity(event.target.value)}>
            <option value="">All severities</option>
            {['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'].map((candidate) => (
              <option key={candidate}>{candidate}</option>
            ))}
          </select>
        </div>
      </section>
      <div className="metrics-grid">
        {[
          ['Complaints', report.total],
          ['Open', report.open],
          ['Closed', report.closed],
          ['Overdue', report.overdue],
          ['Average acknowledgment', `${report.avgAckMinutes} min`],
          ['Average resolution', `${report.avgResolutionHours} hr`],
        ].map(([label, value]) => (
          <section className="metric-card" key={label}>
            <span>{label}</span>
            <strong>{value}</strong>
          </section>
        ))}
      </div>
      <section className="panel">
        <ComplaintTable state={{ ...state, complaints: filtered }} onOpen={() => undefined} />
      </section>
    </>
  )
}

function PushNotificationSettings({ show }: { show: (message: string) => void }) {
  const [status, setStatus] = useState<PushStatus>('prompt')
  const [busy, setBusy] = useState(false)
  const [problem, setProblem] = useState<string>()
  const standalone = isStandalone()

  useEffect(() => {
    getPushStatus()
      .then(setStatus)
      .catch(() => setStatus('unsupported'))
  }, [])

  const enable = async () => {
    setBusy(true)
    setProblem(undefined)
    try {
      const { publicKey } = await api.pushPublicKey()
      const keys = await subscribeForPush(publicKey)
      await api.pushSubscribe({
        endpoint: keys.endpoint,
        keys: { p256dh: keys.p256dh, auth: keys.auth },
        deviceLabel: standalone ? 'Home Screen app' : 'Browser',
      })
      setStatus('subscribed')
      show('Push notifications enabled on this device')
    } catch (error) {
      setProblem((error as Error).message)
      setStatus(await getPushStatus().catch(() => 'unsupported' as PushStatus))
    } finally {
      setBusy(false)
    }
  }

  const disable = async () => {
    setBusy(true)
    setProblem(undefined)
    try {
      const endpoint = await unsubscribeLocalPush()
      if (endpoint) await api.pushUnsubscribe(endpoint)
      setStatus('unsubscribed')
      show('Push notifications disabled on this device')
    } catch (error) {
      setProblem((error as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <section className="panel settings-card">
      <h2>Push notifications</h2>
      <p>
        Get StoreResolve alerts on this device, even when the app is closed. Tapping a
        notification opens the complaint directly.
      </p>
      {status === 'unsupported' && (
        <p>Push notifications aren&apos;t supported on this device or browser.</p>
      )}
      {status === 'denied' && (
        <p>
          Notifications are blocked for StoreResolve. Enable them in your device settings, then
          return here.
        </p>
      )}
      {!standalone && (status === 'prompt' || status === 'unsubscribed') && (
        <p>
          <strong>iPhone tip:</strong> for notifications to work, install StoreResolve first —
          tap Share <strong>→ Add to Home Screen</strong>, then open it from your Home Screen
          and enable notifications here.
        </p>
      )}
      {(status === 'prompt' || status === 'unsubscribed') && (
        <button className="secondary" disabled={busy} onClick={enable}>
          {busy ? 'Enabling…' : 'Enable push notifications'}
        </button>
      )}
      {status === 'subscribed' && (
        <>
          <p>
            <strong>Enabled on this device.</strong> You&apos;ll receive complaint alerts here.
          </p>
          <button className="secondary" disabled={busy} onClick={disable}>
            {busy ? 'Disabling…' : 'Disable on this device'}
          </button>
        </>
      )}
      {problem && <p className="form-error">{problem}</p>}
      <small>
        Notifications never include customer details — only the store and case reference.
      </small>
    </section>
  )
}

function PilotControls({
  state,
  setState,
  onRun,
  providerReady,
  health,
  onSaveConfig,
  onSendTest,
  onSendTestPush,
  onReconcile,
  onPilotIngest,
}: {
  state: AppState
  setState: (s: AppState) => void
  onRun: () => void
  providerReady: boolean
  health?: IntegrationHealth
  onSaveConfig: (config: AppState['config']) => void
  onSendTest: (id: 'father' | 'uncle' | 'grandfather' | 'pilot-admin' | 'dad') => void
  onSendTestPush: (id: 'father' | 'uncle' | 'grandfather' | 'pilot-admin' | 'dad') => void
  onReconcile: (providerMessageId: string) => void
  onPilotIngest: () => void
}) {
  const owners = state.users.filter((u) =>
    ['father', 'uncle', 'grandfather', 'dad'].includes(u.id),
  )
  const pilotAdmin = state.users.find((u) => u.recipientKind === 'PILOT_ADMIN')
  const testRecipients = pilotAdmin ? [...owners, pilotAdmin] : owners
  const [recipient, setRecipient] = useState<
    'father' | 'uncle' | 'grandfather' | 'pilot-admin' | 'dad'
  >('pilot-admin')
  const [confirmed, setConfirmed] = useState(false)
  const [pushConfirmed, setPushConfirmed] = useState(false)
  const chosen = testRecipients.find((candidate) => candidate.id === recipient)
  const recentTests = state.testNotifications
    .filter((notification) => notification.recipientUserId === recipient)
    .slice(-3)
    .reverse()
  return (
    <>
      <PageTitle
        eyebrow="Administration"
        title="Pilot controls"
        body="Guardrails for the staged external notification rollout."
      />
      <div className="settings-grid">
        <section className="panel settings-card">
          <h2>Notification safety</h2>
          <label>
            Rollout mode
            <select
              value={state.config.mode}
              onChange={(e) =>
                onSaveConfig({
                  ...state.config,
                  mode: e.target.value as AppState['config']['mode'],
                })
              }
            >
              {['MOCK', 'FAMILY_PILOT', 'SINGLE_STORE_PILOT', 'FULL'].map((x) => (
                <option key={x}>{x}</option>
              ))}
            </select>
            <small>
              MOCK is the safe default. Family pilot allows only configured ownership recipients.
            </small>
          </label>
          {state.config.mode === 'SINGLE_STORE_PILOT' && (
            <label>
              Single-store pilot location
              <select
                value={state.config.pilotStoreId ?? ''}
                onChange={(event) =>
                  onSaveConfig({
                    ...state.config,
                    pilotStoreId: event.target.value || undefined,
                  })
                }
              >
                <option value="">Select an active store</option>
                {state.stores
                  .filter((store) => store.active)
                  .map((store) => (
                    <option key={store.id} value={store.id}>
                      {store.number} · {store.name}
                    </option>
                  ))}
              </select>
              <small>Only this store's configured manager is eligible in this rollout mode.</small>
            </label>
          )}
          <label className="toggle-row">
            <span>
              <strong>External notifications</strong>
              <small>Master backend kill switch</small>
            </span>
            <button
              className={cx('toggle', state.config.externalNotificationsEnabled && 'on')}
              onClick={() =>
                onSaveConfig({
                  ...state.config,
                  externalNotificationsEnabled: !state.config.externalNotificationsEnabled,
                })
              }
            >
              <i />
            </button>
          </label>
          <div className="danger-note">
            <ShieldCheck />
            <span>
              <strong>
                {state.config.externalNotificationsEnabled
                  ? 'External sending eligibility enabled'
                  : 'External communication is blocked'}
              </strong>
              Every attempted message is still independently recorded.
            </span>
          </div>
        </section>
        <section className="panel settings-card">
          <h2>Microsoft Graph mail</h2>
          <p>
            Intake and outbound acknowledgment are independent server-enforced controls. Both remain
            off until the mailbox OAuth secrets are configured and a pilot is approved.
          </p>
          <div className="readiness-list">
            <span>
              Microsoft Graph provider{' '}
              <strong>{health?.microsoftGraphReady ? 'READY' : 'NOT READY'}</strong>
            </span>
            <span>
              Last background run <strong>{fmt(health?.lastBackgroundRunAt)}</strong>
            </span>
          </div>
          <label className="toggle-row">
            <span>
              <strong>Email ingestion</strong>
              <small>Reads and persists matching messages; does not send a reply.</small>
            </span>
            <button
              className={cx('toggle', state.config.emailIngestionEnabled && 'on')}
              onClick={() =>
                onSaveConfig({
                  ...state.config,
                  emailIngestionEnabled: !state.config.emailIngestionEnabled,
                })
              }
            >
              <i />
            </button>
          </label>
          <label className="toggle-row">
            <span>
              <strong>Automatic acknowledgment</strong>
              <small>Replies once in the source Outlook conversation when enabled.</small>
            </span>
            <button
              className={cx('toggle', state.config.emailAckEnabled && 'on')}
              onClick={() =>
                onSaveConfig({ ...state.config, emailAckEnabled: !state.config.emailAckEnabled })
              }
            >
              <i />
            </button>
          </label>
          <label>
            Inbox lookback (days)
            <input
              type="number"
              min="1"
              max="365"
              value={state.config.emailLookbackDays ?? 30}
              onChange={(event) =>
                onSaveConfig({ ...state.config, emailLookbackDays: Number(event.target.value) })
              }
            />
          </label>
          <button
            className="secondary"
            disabled={
              !health?.microsoftGraphReady ||
              state.config.mode !== 'FAMILY_PILOT' ||
              state.config.externalNotificationsEnabled ||
              Boolean(state.config.emailIngestionEnabled) ||
              Boolean(state.config.emailAckEnabled)
            }
            onClick={onPilotIngest}
          >
            Ingest one recent Dunkin complaint
          </button>
          <small>
            One-shot pilot: Inbox metadata must match the approved sender, subject prefix, and
            five-minute window. At most two metadata rows prove uniqueness before one body read.
          </small>
        </section>
        <CheckMailNow />
        <section className="panel settings-card">
          <h2>SLA and background processing</h2>
          <label>
            Manager acknowledgment deadline (minutes)
            <input
              type="number"
              min="5"
              value={state.config.managerAckDeadlineMinutes ?? 30}
              onChange={(event) =>
                onSaveConfig({
                  ...state.config,
                  managerAckDeadlineMinutes: Number(event.target.value),
                })
              }
            />
          </label>
          <label>
            Resolution target (hours)
            <input
              type="number"
              min="1"
              value={state.config.managerResolutionTargetHours ?? 24}
              onChange={(event) =>
                onSaveConfig({
                  ...state.config,
                  managerResolutionTargetHours: Number(event.target.value),
                })
              }
            />
          </label>
          <label>
            Escalation interval (minutes)
            <input
              type="number"
              min="15"
              value={state.config.escalationIntervalMinutes ?? 60}
              onChange={(event) =>
                onSaveConfig({
                  ...state.config,
                  escalationIntervalMinutes: Number(event.target.value),
                })
              }
            />
          </label>
          <label>
            SignalWire reconciliation threshold (minutes)
            <input
              type="number"
              min="5"
              value={state.config.signalWireReconcileAfterMinutes ?? 10}
              onChange={(event) =>
                onSaveConfig({
                  ...state.config,
                  signalWireReconcileAfterMinutes: Number(event.target.value),
                })
              }
            />
          </label>
        </section>
        <section className="panel settings-card">
          <h2>Ownership recipients</h2>
          <p>Contact values live only in D1 and are returned masked.</p>
          {owners.map((owner) => (
            <OwnerContact key={owner.id} owner={owner} onSaved={setState} />
          ))}
        </section>
        <section className="panel settings-card">
          <h2>Controlled test recipient</h2>
          <p>
            This PILOT_ADMIN destination is only for a manual FAMILY_PILOT test. It is never an
            ownership, complaint-fanout, or escalation recipient.
          </p>
          {state.config.mode !== 'FAMILY_PILOT' ? (
            <div className="danger-note">
              <ShieldCheck />
              <span>
                <strong>Unavailable outside FAMILY_PILOT</strong>
                Switch modes before securely configuring this destination.
              </span>
            </div>
          ) : pilotAdmin ? (
            <OwnerContact owner={pilotAdmin} onSaved={setState} />
          ) : (
            <p>PILOT_ADMIN migration is not available.</p>
          )}
        </section>
        <section className="panel settings-card">
          <h2>Deadline test clock</h2>
          <p>
            Advance the effective clock by three days to exercise acknowledgment and resolution
            escalation idempotently.
          </p>
          <button className="secondary" onClick={onRun}>
            <Clock3 />
            Advance 3 days & process
          </button>
          <button
            className="secondary"
            onClick={() =>
              api
                .processDeadlines(0)
                .then((next) => {
                  setState(next)
                })
                .catch(() => {})
            }
          >
            Send pending notifications now
          </button>
        </section>
        <section className="panel settings-card">
          <h2>Test notification</h2>
          <p>
            Select exactly one controlled recipient. PILOT_ADMIN is manual-test-only and can never
            receive complaint or escalation messages.
          </p>
          <select
            value={recipient}
            onChange={(e) => {
              setRecipient(e.target.value as typeof recipient)
              setConfirmed(false)
            }}
          >
            {testRecipients.map((candidate) => (
              <option value={candidate.id} key={candidate.id}>
                {candidate.name} · {candidate.phone}
              </option>
            ))}
          </select>
          <div className="readiness-list">
            <span>
              Rollout mode <strong>{state.config.mode}</strong>
            </span>
            <span>
              External notifications{' '}
              <strong>{state.config.externalNotificationsEnabled ? 'ENABLED' : 'DISABLED'}</strong>
            </span>
            <span>
              SignalWire provider <strong>{providerReady ? 'READY' : 'NOT READY'}</strong>
            </span>
            <span>
              Destination <strong>{chosen?.phone ?? 'Not configured'}</strong>
            </span>
          </div>
          <label className="confirm-row">
            <input
              type="checkbox"
              checked={confirmed}
              onChange={(e) => setConfirmed(e.target.checked)}
            />{' '}
            I confirm one harmless test to {chosen?.name ?? 'the selected recipient'}
          </label>
          <button
            className="secondary"
            disabled={
              !confirmed ||
              !chosen ||
              state.config.mode !== 'FAMILY_PILOT' ||
              !state.config.externalNotificationsEnabled ||
              !providerReady ||
              !chosen.active ||
              !chosen.smsEnabled ||
              !chosen.phone
            }
            onClick={() => {
              onSendTest(recipient)
              setConfirmed(false)
            }}
          >
            Send one confirmed test
          </button>
          <small>
            No manager or send-to-all target exists. The server rechecks every safety gate and
            permanently excludes PILOT_ADMIN from complaint delivery.
          </small>
          {recentTests.length > 0 && (
            <div className="test-results">
              <strong>Recent delivery state</strong>
              {recentTests.map((notification) => (
                <div key={notification.id}>
                  <span>{fmt(notification.createdAt)}</span>
                  <Badge
                    tone={
                      notification.status === 'DELIVERED'
                        ? 'safe'
                        : notification.status === 'FAILED' || notification.status === 'UNDELIVERED'
                          ? 'danger'
                          : 'warning'
                    }
                  >
                    {notification.status}
                  </Badge>
                  {recipient === 'pilot-admin' &&
                    notification.provider === 'SIGNALWIRE' &&
                    notification.providerMessageId &&
                    state.config.mode === 'FAMILY_PILOT' &&
                    !state.config.externalNotificationsEnabled && (
                      <button
                        className="secondary"
                        onClick={() => onReconcile(notification.providerMessageId!)}
                      >
                        Reconcile provider state
                      </button>
                    )}
                </div>
              ))}
            </div>
          )}
        </section>
        <section className="panel settings-card">
          <h2>Test push notification</h2>
          <p>
            Sends a harmless test push to the selected recipient&apos;s subscribed device.
            Requires the recipient to have enabled push notifications in the section above.
          </p>
          <div className="readiness-list">
            <span>
              Recipient <strong>{chosen?.name ?? 'Not selected'}</strong>
            </span>
            <span>
              Needs <strong>an active push subscription</strong> for the recipient
            </span>
          </div>
          <label className="confirm-row">
            <input
              type="checkbox"
              checked={pushConfirmed}
              onChange={(e) => setPushConfirmed(e.target.checked)}
            />{' '}
            I confirm one harmless test push to {chosen?.name ?? 'the selected recipient'}
          </label>
          <button
            className="secondary"
            disabled={!pushConfirmed || !chosen}
            onClick={() => {
              onSendTestPush(recipient)
              setPushConfirmed(false)
            }}
          >
            Send one confirmed test push
          </button>
          <small>
            Disabled until the operator enables push test sends. The server rechecks every
            safety gate; the kill switch stays off by default.
          </small>
        </section>
      </div>
    </>
  )
}

function OwnerContact({
  owner,
  onSaved,
}: {
  owner: AppState['users'][number]
  onSaved: (state: AppState) => void
}) {
  const [editing, setEditing] = useState(false)
  const [draft, setDraft] = useState({
    name: owner.name,
    email: '',
    phone: '',
    smsEnabled: owner.smsEnabled,
    active: owner.active,
  })
  if (!editing)
    return (
      <div className="recipient">
        <Avatar name={owner.name} />
        <div>
          <strong>{owner.name}</strong>
          <span>
            {owner.phone} · {owner.email}
          </span>
        </div>
        <Badge tone={owner.smsEnabled ? 'safe' : 'warning'}>
          {owner.smsEnabled ? 'SMS ON' : 'SMS OFF'}
        </Badge>
        <button className="text-button" onClick={() => setEditing(true)}>
          Configure
        </button>
      </div>
    )
  return (
    <div className="contact-editor">
      <label>
        Name
        <input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} />
      </label>
      <label>
        New email
        <input
          type="email"
          placeholder="name@example.com"
          value={draft.email}
          onChange={(e) => setDraft({ ...draft, email: e.target.value })}
        />
      </label>
      <label>
        New phone
        <input
          placeholder="+15551234567"
          value={draft.phone}
          onChange={(e) => setDraft({ ...draft, phone: e.target.value })}
        />
      </label>
      <label>
        <input
          type="checkbox"
          checked={draft.smsEnabled}
          onChange={(e) => setDraft({ ...draft, smsEnabled: e.target.checked })}
        />{' '}
        SMS enabled
      </label>
      <label>
        <input
          type="checkbox"
          checked={draft.active}
          onChange={(e) => setDraft({ ...draft, active: e.target.checked })}
        />{' '}
        Active
      </label>
      <div>
        <button className="secondary" onClick={() => setEditing(false)}>
          Cancel
        </button>
        <button
          className="primary"
          onClick={() =>
            api.updateContact(owner.id, draft).then((state) => {
              onSaved(state)
              setEditing(false)
            })
          }
        >
          Save securely
        </button>
      </div>
    </div>
  )
}
