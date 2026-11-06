import { useCallback, useEffect, useState } from 'react';
import { api, ApiError, type AuditResponse, type ListMessage, type MessageDetail, type Summary } from './api';
import { Card, ClassificationBanner, dueText, Empty, fmtDate, LifecyclePill, PrecedencePill, Stat } from './ui';

type Toast = (text: string, tone?: 'ok' | 'error') => void;

function errText(e: unknown) {
  return e instanceof ApiError || e instanceof Error ? e.message : String(e);
}

// ---------------------------------------------------------------- Home

export function HomePage({ toast, go }: { toast: Toast; go: (h: string) => void }) {
  const [s, setS] = useState<Summary | null>(null);
  const [err, setErr] = useState('');
  const load = useCallback(() => {
    api.summary().then(setS, (e) => setErr(errText(e)));
  }, []);
  useEffect(load, [load]);

  async function ack(m: ListMessage) {
    try {
      await api.acknowledge(m.id);
      toast(`Acknowledged ${m.reference}.`);
      load();
    } catch (e) {
      toast(errText(e), 'error');
    }
  }

  if (err) return <Empty>{err}</Empty>;
  if (!s) return <Empty>Loading...</Empty>;
  return (
    <div className="page">
      <div className="page-head">
        <h1>Home</h1>
        <p className="muted">{s.commandName}</p>
      </div>
      <div className="stats">
        <Stat label="Acknowledgements due" value={s.ackDueCount} tone={s.ackDueCount ? 'warn' : 'ok'} />
        <Stat label="Unread" value={s.unread} tone="info" />
        <Stat label="Messages on file" value={s.total} />
      </div>
      <Card title="Acknowledgement required">
        {s.ackDue.length === 0 ? (
          <Empty>Nothing waits on your acknowledgement.</Empty>
        ) : (
          <ul className="ack-list">
            {s.ackDue.map((m) => {
              const due = dueText(m.ackBy);
              return (
                <li key={m.id}>
                  <PrecedencePill value={m.precedence} />
                  <button className="link ack-subject" onClick={() => go(`#/messages/${m.id}`)}>
                    <span className="ref">{m.reference}</span> {m.subject}
                  </button>
                  <span className={due.overdue ? 'due overdue' : 'due'}>{due.text}</span>
                  <button className="btn" onClick={() => ack(m)}>
                    Acknowledge
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </Card>
      <Card title="Latest traffic" action={<button className="link" onClick={() => go('#/messages')}>All messages</button>}>
        <MessageTable rows={s.latest} onOpen={(id) => go(`#/messages/${id}`)} />
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------- Messages

function MessageTable({ rows, selected, onOpen, compact }: { rows: ListMessage[]; selected?: number; onOpen: (id: number) => void; compact?: boolean }) {
  if (rows.length === 0) return <Empty>No messages match.</Empty>;
  if (compact) {
    // Beside the open message: two lines a row, so the subject has room.
    return (
      <ul className="compact-list">
        {rows.map((m) => (
          <li key={m.id}>
            <button className={`compact-row ${m.readState === 'unread' ? 'unread' : ''} ${selected === m.id ? 'selected' : ''}`} onClick={() => onOpen(m.id)}>
              <span className="compact-top">
                <span className="dot">{m.readState === 'unread' ? '\u25CF' : ''}</span>
                <PrecedencePill value={m.precedence} />
                <span className="ref">{m.reference}</span>
                <span className="muted compact-dtg">{m.dtg}</span>
              </span>
              <span className="compact-subject">
                {m.subject} <LifecyclePill value={m.lifecycle} />
                {m.ackDue ? <span className="pill ack-due">Ack due</span> : null}
              </span>
            </button>
          </li>
        ))}
      </ul>
    );
  }
  return (
    <div className="table-wrap">
      <table className="grid">
        <thead>
          <tr>
            <th aria-label="Unread" />
            <th>Precedence</th>
            <th>DTG</th>
            <th>Reference</th>
            <th>Subject</th>
            <th>Ack</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((m) => (
            <tr
              key={m.id}
              className={`${m.readState === 'unread' ? 'unread' : ''} ${selected === m.id ? 'selected' : ''}`}
              onClick={() => onOpen(m.id)}
              tabIndex={0}
              onKeyDown={(e) => e.key === 'Enter' && onOpen(m.id)}
            >
              <td className="dot">{m.readState === 'unread' ? '●' : ''}</td>
              <td>
                <PrecedencePill value={m.precedence} />
              </td>
              <td className="nowrap">{m.dtg}</td>
              <td className="nowrap ref">{m.reference}</td>
              <td className="subject">
                {m.subject} <LifecyclePill value={m.lifecycle} />
                {m.classification !== 'UNCLASSIFIED' && <span className="pill classif-pill">{m.classification} (DEMO)</span>}
              </td>
              <td className="nowrap">
                {m.ackDue ? <span className="pill ack-due">Ack due</span> : m.ackOn ? <span className="muted">Acknowledged</span> : ''}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// True on a phone-width window, where the list always uses two-line rows.
function useNarrow() {
  const q = '(max-width: 760px)';
  const [narrow, setNarrow] = useState(() => matchMedia(q).matches);
  useEffect(() => {
    const mq = matchMedia(q);
    const on = () => setNarrow(mq.matches);
    mq.addEventListener('change', on);
    return () => mq.removeEventListener('change', on);
  }, []);
  return narrow;
}

export function MessagesPage({ id, toast, go }: { id?: number; toast: Toast; go: (h: string) => void }) {
  const [q, setQ] = useState('');
  const [precedence, setPrecedence] = useState('');
  const [unread, setUnread] = useState(false);
  const [ackDue, setAckDue] = useState(false);
  const [rows, setRows] = useState<ListMessage[] | null>(null);
  const [err, setErr] = useState('');
  const narrow = useNarrow();

  const load = useCallback(() => {
    api.messages({ q, precedence, unread, ackDue }).then(
      (r) => {
        setRows(r.messages);
        setErr('');
      },
      (e) => setErr(errText(e)),
    );
  }, [q, precedence, unread, ackDue]);

  useEffect(() => {
    const t = setTimeout(load, 200);
    return () => clearTimeout(t);
  }, [load]);

  return (
    <div className="page">
      <div className="page-head">
        <h1>Messages</h1>
      </div>
      <div className="filters">
        <input type="search" placeholder="Search subject, reference, originator, type" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search" />
        <select value={precedence} onChange={(e) => setPrecedence(e.target.value)} aria-label="Precedence">
          <option value="">Any precedence</option>
          <option>FLASH</option>
          <option>IMMEDIATE</option>
          <option>PRIORITY</option>
          <option>ROUTINE</option>
        </select>
        <label className="check">
          <input type="checkbox" checked={unread} onChange={(e) => setUnread(e.target.checked)} /> Unread only
        </label>
        <label className="check">
          <input type="checkbox" checked={ackDue} onChange={(e) => setAckDue(e.target.checked)} /> Ack due only
        </label>
      </div>
      <div className={`split ${id ? 'has-preview' : ''}`}>
        <Card>{err ? <Empty>{err}</Empty> : rows ? <MessageTable rows={rows} selected={id} compact={!!id || narrow} onOpen={(m) => go(`#/messages/${m}`)} /> : <Empty>Loading...</Empty>}</Card>
        {id ? <MessageView key={id} id={id} toast={toast} onChanged={load} onClose={() => go('#/messages')} /> : null}
      </div>
    </div>
  );
}

function MessageView({ id, toast, onChanged, onClose }: { id: number; toast: Toast; onChanged: () => void; onClose: () => void }) {
  const [d, setD] = useState<MessageDetail | null>(null);
  const [err, setErr] = useState('');
  const [busy, setBusy] = useState(false);
  const load = useCallback(() => {
    api.message(id).then(
      (r) => {
        setD(r);
        onChanged(); // opening it marked it read
      },
      (e) => setErr(errText(e)),
    );
  }, [id]);
  useEffect(load, [load]);

  async function ack() {
    setBusy(true);
    try {
      await api.acknowledge(id);
      toast('Acknowledged. Your receipt and its audit entry are saved.');
      load();
    } catch (e) {
      toast(errText(e), 'error');
    } finally {
      setBusy(false);
    }
  }

  if (err) return <Card><Empty>{err}</Empty></Card>;
  if (!d) return <Card><Empty>Loading...</Empty></Card>;
  const m = d.message;
  const s = (k: string) => (m[k] == null ? '' : String(m[k]));
  const due = dueText(s('AckDue'));
  return (
    <section className="card preview" aria-label="Message">
      <ClassificationBanner value={s('Classification')} />
      <header className="preview-head">
        <div className="preview-pills">
          <PrecedencePill value={s('Precedence')} />
          <LifecyclePill value={s('Lifecycle')} />
          <span className="muted">{s('MsgType')}</span>
        </div>
        <button className="icon-btn" onClick={onClose} aria-label="Close message">
          {'✕'}
        </button>
      </header>
      <h2 className="preview-subject">{s('Subject')}</h2>
      <dl className="fields">
        <dt>Reference</dt>
        <dd>{s('Reference')}</dd>
        <dt>DTG</dt>
        <dd>{s('Dtg')}</dd>
        <dt>From</dt>
        <dd>{s('Originator')}</dd>
        <dt>To</dt>
        <dd>{s('ToLine')}</dd>
        {s('InfoLine') && (
          <>
            <dt>Info</dt>
            <dd>{s('InfoLine')}</dd>
          </>
        )}
        {s('ReplacedBy') && (
          <>
            <dt>Replaced by</dt>
            <dd>{s('ReplacedBy')}</dd>
          </>
        )}
      </dl>

      {d.ackDue ? (
        <div className="ack-panel">
          <div>
            <strong>Acknowledgement required.</strong> <span className={due.overdue ? 'due overdue' : 'due'}>{due.text}</span>
          </div>
          <button className="btn" onClick={ack} disabled={busy}>
            {busy ? 'Saving...' : 'Acknowledge'}
          </button>
        </div>
      ) : d.ackOn ? (
        <div className="ack-panel ack-done">You acknowledged this {fmtDate(d.ackOn)}.</div>
      ) : null}

      {s('ActionText') && (
        <div className="action">
          <strong>Action:</strong> {s('ActionText')} {s('ActionSuspense') && <span className="muted">(suspense {s('ActionSuspense')}, {s('ActionAssignedTo')})</span>}
        </div>
      )}

      {s('AiSummary') && (
        <div className="ai-summary">
          <div className="ai-label">Summary -- Not authoritative</div>
          {s('AiSummary')}
        </div>
      )}

      <h3>Message</h3>
      {d.sections.length > 0 ? (
        d.sections.map((x) => (
          <div key={x.SortOrder} className="section">
            <h4>
              {x.SortOrder}. {x.Heading}
            </h4>
            <p>{x.Body}</p>
          </div>
        ))
      ) : (
        <p className="body-text">{s('BodyText')}</p>
      )}

      {d.attachments.length > 0 && (
        <>
          <h3>Attachments</h3>
          <ul className="attachments">
            {d.attachments.map((a) => (
              <li key={a.FileName}>
                <span className="kind">{a.Kind.toUpperCase()}</span> {a.FileName} <span className="muted">{a.SizeText}</span>
              </li>
            ))}
          </ul>
        </>
      )}

      {d.receipts.length > 0 && (
        <>
          <h3>Receipts from addressees</h3>
          <ul className="receipts">
            {d.receipts.map((r) => (
              <li key={r.CommandName + r.ReceivedOn}>
                {r.CommandName} <span className="muted">{fmtDate(r.ReceivedOn)}</span>
              </li>
            ))}
          </ul>
        </>
      )}

      <details className="official">
        <summary>Official text</summary>
        <pre>{s('RawText').replace(/\\n/g, '\n')}</pre>
      </details>

      <div className="integrity">
        <span>Integrity</span> SHA-256 {s('Sha256') || 'not recorded'} {s('C2oixId') && <> &middot; {s('C2oixId')}</>}
      </div>
    </section>
  );
}

// ---------------------------------------------------------------- Audit

export function AuditPage() {
  const [a, setA] = useState<AuditResponse | null>(null);
  const [err, setErr] = useState('');
  const load = useCallback(() => {
    setA(null);
    api.audit().then(setA, (e) => setErr(errText(e)));
  }, []);
  useEffect(load, [load]);
  if (err) return <Empty>{err}</Empty>;
  return (
    <div className="page">
      <div className="page-head">
        <h1>Audit</h1>
        <p className="muted">Append-only and hash chained. Web and Access entries share one chain.</p>
      </div>
      {!a ? (
        <Empty>Checking the chain...</Empty>
      ) : (
        <>
          <div className={`chain ${a.chain.ok ? 'chain-ok' : 'chain-bad'}`}>
            {a.chain.ok ? `Chain verified: all ${a.chain.checked} entries match.` : a.chain.problem}
            <button className="btn btn-outline" onClick={load}>
              Recompute
            </button>
          </div>
          <Card>
            <div className="table-wrap">
              <table className="grid">
                <thead>
                  <tr>
                    <th>When</th>
                    <th>Who</th>
                    <th>Action</th>
                    <th>Resource</th>
                    <th>Device</th>
                    <th>Result</th>
                  </tr>
                </thead>
                <tbody>
                  {a.rows.map((r) => (
                    <tr key={r.AuditID} title={r.Details ?? ''}>
                      <td className="nowrap">{fmtDate(r.LoggedAt)}</td>
                      <td>{r.UserName}</td>
                      <td>{r.Action}</td>
                      <td>{r.Resource}</td>
                      <td className="nowrap">{r.Device}</td>
                      <td>{r.Result}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </>
      )}
    </div>
  );
}
