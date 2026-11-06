import { useCallback, useEffect, useMemo, useState, type KeyboardEvent } from 'react';
import { Card, ClassificationBanner, Empty, errText, fmtDate, go as goHash, PageHead, Pill, useToast, type Tone } from '@kit';
import { api, NE_LABEL, STATUSES, type Attribute, type Band, type HistoryRow, type InspectionDetail, type InspectionList, type Node, type Status } from './api';

// Paths inside SABER: go('/inspections/12') opens #/saber/inspections/12.
const go = (path: string) => goHash(`#/saber${path}`);

// ---------------------------------------------------------------- pills and dates

const STATUS_TONE: Record<Status, Tone> = { active: 'success', configuring: 'info', closed: 'neutral', archived: 'neutral' };
const STATUS_LABEL: Record<'all' | Status, string> = { all: 'All', active: 'Active', configuring: 'Configuring', closed: 'Closed', archived: 'Archived' };

function StatusPill({ status }: { status: Status }) {
  return <Pill tone={STATUS_TONE[status] ?? 'neutral'}>{STATUS_LABEL[status] ?? status}</Pill>;
}

// Bands are listed lowest first; the top two read as good, the bottom two as bad.
function bandTone(bands: Band[], label: string | null): Tone {
  if (!label) return 'neutral';
  if (label === NE_LABEL) return 'neutral';
  const i = bands.findIndex((b) => b.label === label);
  if (i < 0) return 'neutral';
  const fromTop = bands.length - 1 - i;
  if (fromTop <= 1) return 'success';
  if (i <= 1) return 'danger';
  return fromTop === 2 ? 'info' : 'warning';
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
// "05 Nov 2026", as the Access edition's lists show dates.
function day(iso: string | null | undefined) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return `${String(d.getDate()).padStart(2, '0')} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}
function dateRange(a: string | null, b: string | null) {
  return [day(a), day(b)].filter(Boolean).join(' - ');
}

function Progress({ pct }: { pct: number | null }) {
  if (pct === null) return <span className="kit-muted">No schema</span>;
  return (
    <span className="sb-progress" title={`${pct}% of attributes graded`}>
      <span className="sb-bar"><span style={{ width: `${pct}%` }} /></span>
      <span className="sb-pct">{pct}%</span>
    </span>
  );
}

// The installation's marking (system_setting classification_level). The top
// bar already says UNCLASSIFIED, so only another marking adds a banner (with
// the kit's demonstration-only wording).
function Marking({ value }: { value: string }) {
  return (value || 'UNCLASSIFIED').toUpperCase() === 'UNCLASSIFIED' ? null : <ClassificationBanner value={value} />;
}

// ---------------------------------------------------------------- Inspections (pgDashboard)

export function InspectionsPage() {
  const [q, setQ] = useState('');
  const [year, setYear] = useState<number | null>(new Date().getFullYear());   // the Access edition opens on this year
  const [status, setStatus] = useState<'all' | Status>('all');
  const [data, setData] = useState<InspectionList | null>(null);
  const [err, setErr] = useState('');

  useEffect(() => {
    const t = setTimeout(() => {
      api.inspections({ q, year, status }).then((d) => { setData(d); setErr(''); }, (e) => setErr(errText(e)));
    }, q ? 250 : 0);
    return () => clearTimeout(t);
  }, [q, year, status]);

  if (err && !data) return <Empty>{err}</Empty>;
  if (!data) return <Empty>Loading...</Empty>;
  return (
    <div className="kit-page">
      <Marking value={data.classification} />
      <PageHead title="Inspections">Every inspection in this SABER data file. Open one to grade it.</PageHead>
      <div className="sb-cards" role="group" aria-label="Filter by status">
        {(['all', ...STATUSES] as const).map((k) => (
          <button key={k} className={`sb-card ${status === k ? 'active' : ''}`} aria-pressed={status === k} onClick={() => setStatus(k)}>
            <span className="sb-card-value">{data.counts[k] ?? 0}</span>
            <span className="sb-card-label">{STATUS_LABEL[k]}</span>
          </button>
        ))}
      </div>
      <div className="kit-filters">
        <input type="search" placeholder="Search name, inspected or inspecting command" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search inspections" />
        <select value={year ?? ''} onChange={(e) => setYear(e.target.value ? Number(e.target.value) : null)} aria-label="Year">
          <option value="">All years</option>
          {data.years.map((y) => <option key={y} value={y}>{y}</option>)}
        </select>
        {err && <span className="kit-error">{err}</span>}
      </div>
      <Card>
        {data.inspections.length === 0 ? (
          <Empty>
            No inspections match.{' '}
            {year && <button className="kit-link" onClick={() => setYear(null)}>Show all years</button>}
          </Empty>
        ) : (
          <div className="kit-table-wrap">
            <table className="kit-grid sb-list">
              <thead>
                <tr><th>Inspection</th><th>Inspected command</th><th>Inspecting command</th><th>Dates</th><th>Status</th><th>Graded</th></tr>
              </thead>
              <tbody>
                {data.inspections.map((e) => (
                  <tr key={e.id} className="kit-clickable" onClick={() => go(`/inspections/${e.id}`)}>
                    <td>
                      <a className="sb-name" href={`#/saber/inspections/${e.id}`} onClick={(ev) => ev.stopPropagation()}>{e.name}</a>
                      {e.boatType && <span className="kit-muted sb-boat">{e.boatType}</span>}
                    </td>
                    <td>{e.inspectedCommand}</td>
                    <td>{e.inspectingCommand}</td>
                    <td className="kit-nowrap">{dateRange(e.startDate, e.endDate)}</td>
                    <td><StatusPill status={e.status} /></td>
                    <td><Progress pct={e.pctDone} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------- Grading (pgGrading, subTreeGrading)

type Tree = { ma: Node; groups: { g: Node; caps: { c: Node; attrs: Attribute[] }[] }[] }[];

function buildTree(d: InspectionDetail): Tree {
  const by = <T extends { parentId?: number }>(xs: T[]) => {
    const m = new Map<number, T[]>();
    for (const x of xs) m.set(x.parentId!, [...(m.get(x.parentId!) ?? []), x]);
    return m;
  };
  const groups = by(d.groups), caps = by(d.capabilities), attrs = by(d.attributes);
  return d.missionAreas.map((ma) => ({
    ma,
    groups: (groups.get(ma.id) ?? []).map((g) => ({ g, caps: (caps.get(g.id) ?? []).map((c) => ({ c, attrs: attrs.get(c.id) ?? [] })) })),
  }));
}

const count = (xs: Attribute[]) => ({ graded: xs.filter((a) => a.categorical).length, total: xs.length });

function Count({ xs }: { xs: Attribute[] }) {
  const n = count(xs);
  return <span className={`sb-count ${n.total && n.graded === n.total ? 'done' : ''}`}>{n.graded}/{n.total}</span>;
}

export function InspectionPage({ id }: { id: number }) {
  const toast = useToast();
  const [d, setD] = useState<InspectionDetail | null>(null);
  const [err, setErr] = useState('');
  const [open, setOpen] = useState<Set<string>>(new Set());
  const [onlyUngraded, setOnlyUngraded] = useState(false);
  const [busy, setBusy] = useState<number | null>(null);
  const [historyFor, setHistoryFor] = useState<Attribute | null>(null);
  const [historyKey, setHistoryKey] = useState(0);

  const load = useCallback(() => api.inspection(id).then((x) => { setD(x); setErr(''); }, (e) => setErr(errText(e))), [id]);
  useEffect(() => {
    setD(null);
    load();
  }, [load]);
  // Mission areas start open, groups and capabilities closed (modRollup.BuildTree).
  useEffect(() => {
    if (d) setOpen((o) => (o.size ? o : new Set(d.missionAreas.map((m) => `ma${m.id}`))));
  }, [d]);

  const tree = useMemo(() => (d ? buildTree(d) : []), [d]);
  if (err && !d) return <Empty>{err}</Empty>;
  if (!d) return <Empty>Loading...</Empty>;
  const ev = d.inspection;
  const toggle = (k: string) => setOpen((o) => { const n = new Set(o); n.has(k) ? n.delete(k) : n.add(k); return n; });
  const expandAll = (on: boolean) => {
    if (!on) return setOpen(new Set());
    const n = new Set<string>();
    tree.forEach((m) => { n.add(`ma${m.ma.id}`); m.groups.forEach((g) => { n.add(`g${g.g.id}`); g.caps.forEach((c) => n.add(`c${c.c.id}`)); }); });
    setOpen(n);
  };

  async function save(a: Attribute, body: { score: number | null } | { band: string } | 'clear') {
    setBusy(a.id);
    try {
      const r = body === 'clear' ? await api.clearGrade(id, a.id) : await api.grade(id, a.id, body);
      if (r.action !== 'unchanged') toast(r.action === 'cleared' ? 'Grade cleared.' : `Graded ${r.categorical}${r.score != null ? ` (${r.score})` : ''}.`);
      await load();
      setHistoryKey((k) => k + 1);
    } catch (e) {
      toast(errText(e), 'error');
      await load();
    } finally {
      setBusy(null);
    }
  }

  const pct = d.progress.total ? Math.round((d.progress.graded * 100) / d.progress.total) : null;
  const show = (xs: Attribute[]) => (onlyUngraded ? xs.filter((a) => !a.categorical) : xs);
  const all = (m: Tree[number]) => m.groups.flatMap((g) => g.caps.flatMap((c) => c.attrs));

  return (
    <div className="kit-page">
      <Marking value={d.classification} />
      <div className="sb-head">
        <button className="kit-link" onClick={() => go('/')}>‹ Inspections</button>
        <div className="sb-title">
          <h1>{ev.name}</h1>
          <StatusPill status={ev.status} />
          {ev.boatType && <Pill>{ev.boatType}</Pill>}
        </div>
        <div className="sb-meta kit-muted">
          {[ev.inspectedCommand && `Inspected: ${ev.inspectedCommand}`, ev.inspectingCommand && `Inspecting: ${ev.inspectingCommand}`,
            ev.conductingCommand && `Conducting: ${ev.conductingCommand}`, dateRange(ev.startDate, ev.endDate)].filter(Boolean).join(' · ')}
        </div>
      </div>
      <div className="kit-stats">
        <div className="kit-stat kit-stat-info"><div className="kit-stat-value">{d.progress.graded}/{d.progress.total}</div><div className="kit-stat-label">Attributes graded</div></div>
        <div className="kit-stat"><div className="kit-stat-value"><Progress pct={pct} /></div><div className="kit-stat-label">Complete</div></div>
        <div className="kit-stat"><div className="kit-stat-value sb-role">{d.you.stage ? d.you.stageLabel : 'No role'}</div><div className="kit-stat-label">Your role on this inspection</div></div>
      </div>
      {!d.you.canGrade && <div className="kit-banner kit-banner-warn" role="note">{d.you.whyNot} Grades are shown read-only.</div>}

      <div className={`sb-split ${historyFor ? 'has-panel' : ''}`}>
        <Card title="Grading" action={
          <div className="sb-tools">
            <label className="kit-check"><input type="checkbox" checked={onlyUngraded} onChange={(e) => setOnlyUngraded(e.target.checked)} /> Ungraded only</label>
            <button className="kit-btn kit-btn-outline kit-btn-sm" onClick={() => expandAll(true)}>Expand all</button>
            <button className="kit-btn kit-btn-outline kit-btn-sm" onClick={() => expandAll(false)}>Collapse all</button>
          </div>
        }>
          {tree.length === 0 ? (
            <Empty>This inspection has no schema yet. Import one on the inspection's Schema tab (Access edition until that page is ported).</Empty>
          ) : (
            <div className="sb-tree" role="tree" aria-label="Schema">
              {tree.map((m) => (
                <div key={m.ma.id} role="treeitem" aria-expanded={open.has(`ma${m.ma.id}`)}>
                  <button className="sb-row sb-l1" onClick={() => toggle(`ma${m.ma.id}`)}>
                    <span className="sb-caret">{open.has(`ma${m.ma.id}`) ? '▾' : '▸'}</span>
                    <span className="sb-row-title">{m.ma.title}</span>
                    <Count xs={all(m)} />
                  </button>
                  {open.has(`ma${m.ma.id}`) && m.groups.map((g) => (
                    <div key={g.g.id} role="group">
                      <button className="sb-row sb-l2" onClick={() => toggle(`g${g.g.id}`)} aria-expanded={open.has(`g${g.g.id}`)}>
                        <span className="sb-caret">{open.has(`g${g.g.id}`) ? '▾' : '▸'}</span>
                        <span className="sb-row-title">{g.g.title}</span>
                        <Count xs={g.caps.flatMap((c) => c.attrs)} />
                      </button>
                      {open.has(`g${g.g.id}`) && g.caps.map((c) => {
                        const rows = show(c.attrs);
                        if (onlyUngraded && rows.length === 0) return null;
                        return (
                          <div key={c.c.id} role="group">
                            <button className="sb-row sb-l3" onClick={() => toggle(`c${c.c.id}`)} aria-expanded={open.has(`c${c.c.id}`)}>
                              <span className="sb-caret">{open.has(`c${c.c.id}`) ? '▾' : '▸'}</span>
                              <span className="sb-row-title">{c.c.code && <span className="sb-code">{c.c.code}</span>}{c.c.title}</span>
                              <Count xs={c.attrs} />
                            </button>
                            {open.has(`c${c.c.id}`) && rows.map((a) => (
                              <AttributeRow key={a.id} a={a} n={c.attrs.indexOf(a) + 1} bands={d.bands} canGrade={d.you.canGrade} busy={busy === a.id}
                                selected={historyFor?.id === a.id} onSave={(b) => save(a, b)} onHistory={() => setHistoryFor(a)} />
                            ))}
                          </div>
                        );
                      })}
                    </div>
                  ))}
                </div>
              ))}
            </div>
          )}
        </Card>
        {historyFor && <HistoryPanel key={`${historyFor.id}-${historyKey}`} inspectionId={id} a={historyFor} onClose={() => setHistoryFor(null)} />}
      </div>
    </div>
  );
}

function AttributeRow({ a, n, bands, canGrade, busy, selected, onSave, onHistory }: {
  a: Attribute; n: number; bands: Band[]; canGrade: boolean; busy: boolean; selected: boolean;
  onSave: (b: { score: number | null } | { band: string } | 'clear') => void; onHistory: () => void;
}) {
  const [text, setText] = useState(a.score == null ? '' : String(a.score));
  const [bad, setBad] = useState('');
  useEffect(() => { setText(a.score == null ? '' : String(a.score)); setBad(''); }, [a.score, a.categorical]);

  // Blank clears the grade; otherwise a whole number 0-100 (subTreeGrading.txtScore_AfterUpdate).
  function commit() {
    const v = text.trim();
    const was = a.score == null ? '' : String(a.score);
    if (v === was) return;                      // includes a Not Evaluated row left blank
    if (v === '') { if (a.categorical) onSave('clear'); return; }
    if (!/^\d{1,3}$/.test(v) || Number(v) > 100) { setBad('Enter a whole number from 0 to 100.'); return; }
    setBad('');
    onSave({ score: Number(v) });
  }
  const key = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
    if (e.key === 'Escape') { setText(a.score == null ? '' : String(a.score)); setBad(''); }
  };
  const label = a.title || a.criterion;
  return (
    <div className={`sb-attr ${selected ? 'selected' : ''}`} role="treeitem">
      <span className="sb-num">{n}.</span>
      <div className="sb-crit">
        <div title={a.criterion}>{label}</div>
        {a.title && a.criterion && a.title !== a.criterion && <div className="kit-muted sb-crit-sub">{a.criterion}</div>}
        {bad && <div className="kit-error">{bad}</div>}
      </div>
      <span className="sb-grade">
        {a.categorical ? <Pill tone={bandTone(bands, a.categorical)}>{a.categorical === NE_LABEL ? 'N/E' : a.categorical}</Pill> : <span className="kit-muted">Ungraded</span>}
        {a.isPreliminary && a.categorical && <span className="sb-prelim" title="Preliminary: a comment on this capability is still in review.">Prelim</span>}
      </span>
      {canGrade ? (
        <span className="sb-entry">
          <input className="sb-score" inputMode="numeric" aria-label={`Score for attribute ${n}`} placeholder="0-100" value={text} disabled={busy}
            onChange={(e) => setText(e.target.value)} onBlur={commit} onKeyDown={key} aria-invalid={!!bad} />
          <select className="sb-band" aria-label={`Band for attribute ${n}`} value="" disabled={busy}
            onChange={(e) => { const v = e.target.value; if (v === '(clear)') onSave('clear'); else if (v === NE_LABEL) onSave({ score: null }); else if (v) onSave({ band: v }); }}>
            <option value="">Band…</option>
            {[...bands].reverse().map((b) => <option key={b.id} value={b.label}>{b.label} ({b.minScore}-{b.maxScore})</option>)}
            <option value={NE_LABEL}>Not Evaluated</option>
            {a.categorical && <option value="(clear)">Clear grade</option>}
          </select>
        </span>
      ) : (
        <span className="sb-entry kit-muted">{a.score ?? ''}</span>
      )}
      <button className="kit-icon-btn" title="Grade history" aria-label={`Grade history for attribute ${n}`} onClick={onHistory}>⏱</button>
    </div>
  );
}

const ACTION: Record<string, string> = { graded: 'Graded', changed: 'Changed', cleared: 'Cleared' };
const gradeText = (score: number | null, cat: string | null) => (cat ? `${cat === NE_LABEL ? 'N/E' : cat}${score != null ? ` (${score})` : ''}` : '-');

// dlgGradeHistory for one attribute: who changed its grade, when, from what to what.
function HistoryPanel({ inspectionId, a, onClose }: { inspectionId: number; a: Attribute; onClose: () => void }) {
  const [rows, setRows] = useState<HistoryRow[] | null>(null);
  const [err, setErr] = useState('');
  useEffect(() => {
    api.history(inspectionId, a.id).then((r) => setRows(r.history), (e) => setErr(errText(e)));
  }, [inspectionId, a.id]);
  return (
    <Card className="sb-panel" title="Grade history" action={<button className="kit-icon-btn" aria-label="Close history" onClick={onClose}>✕</button>}>
      <p className="sb-panel-attr">{a.title || a.criterion}</p>
      {err ? <p className="kit-error">{err}</p> : !rows ? <Empty>Loading...</Empty> : rows.length === 0 ? <Empty>No changes recorded yet.</Empty> : (
        <ol className="sb-history">
          {rows.map((h) => (
            <li key={h.id}>
              <div><strong>{ACTION[h.action] ?? h.action}</strong> {gradeText(h.oldScore, h.oldCategorical)} → {gradeText(h.newScore, h.newCategorical)}</div>
              <div className="kit-muted">{[h.actorRank, h.actor].filter(Boolean).join(' ') || 'Unknown'} · {fmtDate(h.occurredAt)}</div>
            </li>
          ))}
        </ol>
      )}
    </Card>
  );
}
