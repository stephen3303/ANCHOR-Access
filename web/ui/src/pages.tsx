import { useCallback, useEffect, useState } from 'react';
import { Card, ClassificationBanner, dueText, Empty, errText, fmtDate, go, PageHead, Pill, Stat, useNarrow, useToast, type Tone } from '@kit';
import { api, type AuditResponse, type ListMessage, type MessageDetail, type Summary } from './api';

// ---------------------------------------------------------------- ANCHOR pills

const PRECEDENCE_TONE: Record<string, Tone> = { FLASH: 'solid-danger', IMMEDIATE: 'solid-warning', PRIORITY: 'info', ROUTINE: 'neutral' };
const LIFECYCLE_TONE: Record<string, Tone> = { CANCELLED: 'neutral', EXPIRED: 'neutral', SUPERSEDED: 'warning', AMENDED: 'info', CORRECTED: 'info' };

function Precedence({ value }: { value: string }) {
  return <Pill tone={PRECEDENCE_TONE[value] ?? 'neutral'}>{value}</Pill>;
}
function Lifecycle({ value }: { value: string }) {
  return value === 'ACTIVE' ? null : <Pill tone={LIFECYCLE_TONE[value] ?? 'neutral'}>{value}</Pill>;
}
function DemoMark({ value }: { value: string }) {
  return value === 'UNCLASSIFIED' ? null : <Pill tone="danger">{value} (DEMO)</Pill>;
}

// ---------------------------------------------------------------- Home

export function HomePage() {
  const toast = useToast();
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
    <div className="kit-page">
      <PageHead title="Home">{s.commandName}</PageHead>
      <div className="kit-stats">
        <Stat label="Acknowledgements due" value={s.ackDueCount} tone={s.ackDueCount ? 'warn' : 'ok'} />
        <Stat label="Unread" value={s.unread} tone="info" />
        <Stat label="Messages on file" value={s.total} />
      </div>
      <Card title="Acknowledgement required">
        {s.ackDue.length === 0 ? (
          <Empty>Nothing waits on your acknowledgement.</Empty>
        ) : (
          <ul className="an-ack-list">
            {s.ackDue.map((m) => {
              const due = dueText(m.ackBy);
              return (
                <li key={m.id}>
                  <Precedence value={m.precedence} />
                  <button className="kit-link an-ack-subject" onClick={() => go(`#/messages/${m.id}`)}>
                    <span className="an-ref">{m.reference}</span> {m.subject}
                  </button>
                  <span className={due.overdue ? 'an-due an-overdue' : 'an-due'}>{due.text}</span>
                  <button className="kit-btn" onClick={() => ack(m)}>
                    Acknowledge
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </Card>
      <Card title="Latest traffic" action={<button className="kit-link" onClick={() => go('#/messages')}>All messages</button>}>
        <MessageTable rows={s.latest} onOpen={(id) => go(`#/messages/${id}`)} />
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------- Messages

function MessageTable({ rows, selected, onOpen, compact }: { rows: ListMessage[]; selected?: number; onOpen: (id: number) => void; compact?: boolean }) {
  if (rows.length === 0) return <Empty>No messages match.</Empty>;
  if (compact) {
    // Beside the open message, or on a phone: two lines a row.
    return (
      <ul className="an-compact-list">
        {rows.map((m) => (
          <li key={m.id}>
            <button className={`an-compact-row ${m.readState === 'unread' ? 'unread' : ''} ${selected === m.id ? 'selected' : ''}`} onClick={() => onOpen(m.id)}>
              <span className="an-compact-top">
                <span className="an-dot">{m.readState === 'unread' ? '●' : ''}</span>
                <Precedence value={m.precedence} />
                <span className="an-ref">{m.reference}</span>
                <span className="kit-muted an-compact-dtg">{m.dtg}</span>
              </span>
              <span className="an-compact-subject">
                {m.subject} <Lifecycle value={m.lifecycle} /> <DemoMark value={m.classification} />
                {m.ackDue ? <Pill tone="warning">Ack due</Pill> : null}
              </span>
            </button>
          </li>
        ))}
      </ul>
    );
  }
  return (
    <div className="kit-table-wrap">
      <table className="kit-grid an-messages">
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
            <tr key={m.id} className={`kit-clickable ${m.readState === 'unread' ? 'unread' : ''} ${selected === m.id ? 'selected' : ''}`}
              onClick={() => onOpen(m.id)} tabIndex={0} onKeyDown={(e) => e.key === 'Enter' && onOpen(m.id)}>
              <td className="an-dot">{m.readState === 'unread' ? '●' : ''}</td>
              <td><Precedence value={m.precedence} /></td>
              <td className="kit-nowrap">{m.dtg}</td>
              <td className="kit-nowrap an-ref">{m.reference}</td>
              <td className="an-subject">
                {m.subject} <Lifecycle value={m.lifecycle} /> <DemoMark value={m.classification} />
              </td>
              <td className="kit-nowrap">{m.ackDue ? <Pill tone="warning">Ack due</Pill> : m.ackOn ? <span className="kit-muted">Acknowledged</span> : ''}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function MessagesPage({ id }: { id?: number }) {
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
    <div className="kit-page">
      <PageHead title="Messages" />
      <div className="kit-filters">
        <input type="search" placeholder="Search subject, reference, originator, type" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search" />
        <select value={precedence} onChange={(e) => setPrecedence(e.target.value)} aria-label="Precedence">
          <option value="">Any precedence</option>
          <option>FLASH</option>
          <option>IMMEDIATE</option>
          <option>PRIORITY</option>
          <option>ROUTINE</option>
        </select>
        <label className="kit-check">
          <input type="checkbox" checked={unread} onChange={(e) => setUnread(e.target.checked)} /> Unread only
        </label>
        <label className="kit-check">
          <input type="checkbox" checked={ackDue} onChange={(e) => setAckDue(e.target.checked)} /> Ack due only
        </label>
      </div>
      <div className={`an-split ${id ? 'has-preview' : ''}`}>
        <Card>
          {err ? <Empty>{err}</Empty> : rows ? <MessageTable rows={rows} selected={id} compact={!!id || narrow} onOpen={(m) => go(`#/messages/${m}`)} /> : <Empty>Loading...</Empty>}
        </Card>
        {id ? <MessageView key={id} id={id} onChanged={load} onClose={() => go('#/messages')} /> : null}
      </div>
    </div>
  );
}

function MessageView({ id, onChanged, onClose }: { id: number; onChanged: () => void; onClose: () => void }) {
  const toast = useToast();
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
    <section className="kit-card an-preview" aria-label="Message">
      <div className="an-preview-classif"><ClassificationBanner value={s('Classification')} /></div>
      <header className="an-preview-head">
        <div className="an-preview-pills">
          <Precedence value={s('Precedence')} />
          <Lifecycle value={s('Lifecycle')} />
          <span className="kit-muted">{s('MsgType')}</span>
        </div>
        <button className="kit-icon-btn" onClick={onClose} aria-label="Close message">{'✕'}</button>
      </header>
      <h2 className="an-preview-subject">{s('Subject')}</h2>
      <dl className="an-fields">
        <dt>Reference</dt><dd>{s('Reference')}</dd>
        <dt>DTG</dt><dd>{s('Dtg')}</dd>
        <dt>From</dt><dd>{s('Originator')}</dd>
        <dt>To</dt><dd>{s('ToLine')}</dd>
        {s('InfoLine') && (<><dt>Info</dt><dd>{s('InfoLine')}</dd></>)}
        {s('ReplacedBy') && (<><dt>Replaced by</dt><dd>{s('ReplacedBy')}</dd></>)}
      </dl>

      {d.ackDue ? (
        <div className="kit-banner kit-banner-warn an-block">
          <div>
            <strong>Acknowledgement required.</strong> <span className={due.overdue ? 'an-due an-overdue' : 'an-due'}>{due.text}</span>
          </div>
          <button className="kit-btn" onClick={ack} disabled={busy}>{busy ? 'Saving...' : 'Acknowledge'}</button>
        </div>
      ) : d.ackOn ? (
        <div className="kit-banner kit-banner-ok an-block">You acknowledged this {fmtDate(d.ackOn)}.</div>
      ) : null}

      {s('ActionText') && (
        <div className="an-action an-block">
          <strong>Action:</strong> {s('ActionText')} {s('ActionSuspense') && <span className="kit-muted">(suspense {s('ActionSuspense')}, {s('ActionAssignedTo')})</span>}
        </div>
      )}

      {s('AiSummary') && (
        <div className="an-ai-summary">
          <div className="an-ai-label">Summary -- Not authoritative</div>
          {s('AiSummary')}
        </div>
      )}

      <h3>Message</h3>
      {d.sections.length > 0 ? (
        d.sections.map((x) => (
          <div key={x.SortOrder} className="an-section">
            <h4>{x.SortOrder}. {x.Heading}</h4>
            <p>{x.Body}</p>
          </div>
        ))
      ) : (
        <p className="an-body-text">{s('BodyText')}</p>
      )}

      {d.attachments.length > 0 && (
        <>
          <h3>Attachments</h3>
          <ul className="an-plain-list">
            {d.attachments.map((a) => (
              <li key={a.FileName}><span className="an-kind">{a.Kind.toUpperCase()}</span> {a.FileName} <span className="kit-muted">{a.SizeText}</span></li>
            ))}
          </ul>
        </>
      )}

      {d.receipts.length > 0 && (
        <>
          <h3>Receipts from addressees</h3>
          <ul className="an-plain-list">
            {d.receipts.map((r) => (
              <li key={r.CommandName + r.ReceivedOn}>{r.CommandName} <span className="kit-muted">{fmtDate(r.ReceivedOn)}</span></li>
            ))}
          </ul>
        </>
      )}

      <details className="an-official">
        <summary>Official text</summary>
        <pre>{s('RawText').replace(/\\n/g, '\n')}</pre>
      </details>

      <div className="an-integrity">
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
    <div className="kit-page">
      <PageHead title="Audit">Append-only and hash chained. Web and Access entries share one chain.</PageHead>
      {!a ? (
        <Empty>Checking the chain...</Empty>
      ) : (
        <>
          <div className={`kit-banner ${a.chain.ok ? 'kit-banner-ok' : 'kit-banner-bad'}`}>
            <strong>{a.chain.ok ? `Chain verified: all ${a.chain.checked} entries match.` : a.chain.problem}</strong>
            <button className="kit-btn kit-btn-outline" onClick={load}>Recompute</button>
          </div>
          <Card>
            <div className="kit-table-wrap">
              <table className="kit-grid">
                <thead>
                  <tr><th>When</th><th>Who</th><th>Action</th><th>Resource</th><th>Device</th><th>Result</th></tr>
                </thead>
                <tbody>
                  {a.rows.map((r) => (
                    <tr key={r.AuditID} title={r.Details ?? ''}>
                      <td className="kit-nowrap">{fmtDate(r.LoggedAt)}</td>
                      <td>{r.UserName}</td>
                      <td>{r.Action}</td>
                      <td>{r.Resource}</td>
                      <td className="kit-nowrap">{r.Device}</td>
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
