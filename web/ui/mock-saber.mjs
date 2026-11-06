// SABER stand-in for UI work (npm run mock): the /api/saber/ routes of
// web/apps/saber/api.ps1 over a small invented inspection, in memory. All
// names are placeholders. Not used by the real server.
const NE = 'Not Evaluated';
const now = () => new Date().toISOString().slice(0, 19);
const bands = [
  ['Unsatisfactory', 0, 24], ['Significantly Below Standards', 25, 59], ['Below Standards', 60, 69],
  ['At Standards', 70, 79], ['Above Standards', 80, 89], ['Outstanding', 90, 100],
].map(([label, minScore, maxScore], i) => ({ id: i + 1, label, minScore, maxScore }));
const inspections = [
  { id: 1, name: 'FY27 Readiness Evaluation', inspectedCommand: 'USS Example (SSN-800)', inspectingCommand: 'TYCOM Assessment Team', status: 'active', boatType: 'SSN', startDate: `${new Date().getFullYear()}-11-02T00:00:00`, endDate: `${new Date().getFullYear()}-11-13T00:00:00`, createdAt: '2026-10-01T09:00:00' },
  { id: 2, name: 'Pre-deployment Assessment', inspectedCommand: 'USS Placeholder (SSBN-750)', inspectingCommand: 'TYCOM Assessment Team', status: 'configuring', boatType: 'SSBN', startDate: `${new Date().getFullYear()}-12-01T00:00:00`, endDate: null, createdAt: '2026-10-20T09:00:00' },
  { id: 3, name: 'Spring Readiness Evaluation', inspectedCommand: 'USS Sample (SSGN-730)', inspectingCommand: 'TYCOM Assessment Team', status: 'closed', boatType: 'SSGN', startDate: `${new Date().getFullYear()}-04-06T00:00:00`, endDate: `${new Date().getFullYear()}-04-17T00:00:00`, createdAt: '2026-03-01T09:00:00' },
];
const missionAreas = [{ id: 1, title: 'Mobility', description: '' }, { id: 2, title: 'Strike Warfare', description: '' }];
const groups = [{ id: 1, parentId: 1, title: 'Ship Control' }, { id: 2, parentId: 1, title: 'Navigation' }, { id: 3, parentId: 2, title: 'Weapons Handling' }];
const capabilities = [
  { id: 1, parentId: 1, code: 'MOB 1.1', title: 'Maintain depth and course control' }, { id: 2, parentId: 2, code: 'MOB 2.1', title: 'Plan and execute safe navigation' },
  { id: 3, parentId: 3, code: 'STW 1.1', title: 'Load and stow weapons safely' },
];
const crit = ['Watchstanders demonstrate correct procedures.', 'Supervisors verify readings and take timely action.', 'Records are complete and reviewed.'];
const attributes = capabilities.flatMap((c) => crit.map((t, i) => ({ id: c.id * 10 + i, parentId: c.id, title: null, criterion: t, score: null, categorical: null, isPreliminary: null, gradedAt: null, gradedBy: null })));
const history = [];
const me = { name: 'Senior Member', rank: 'CAPT' };
const grade = (a, score) => {
  const before = { score: a.score, cat: a.categorical };
  a.score = score;
  a.categorical = score === null ? NE : bands.find((b) => b.minScore <= score && score <= b.maxScore).label;
  a.isPreliminary = a.parentId === 1; a.gradedAt = now(); a.gradedBy = me.name;
  const action = before.cat ? 'changed' : 'graded';
  if (before.cat !== a.categorical || before.score !== a.score) history.push({ id: history.length + 1, attr: a.id, action, oldScore: before.score, oldCategorical: before.cat, newScore: a.score, newCategorical: a.categorical, occurredAt: now(), actor: me.name, actorRank: me.rank });
  return { status: 200, ok: true, action, score: a.score, categorical: a.categorical };
};
grade(attributes[0], 85); grade(attributes[1], 72); grade(attributes[0], 91); grade(attributes[3], null);

const row = (e) => {
  const attrs = e.id === 1 ? attributes : e.id === 3 ? attributes : [];
  const n = attrs.length, g = e.id === 1 ? attributes.filter((a) => a.categorical).length : n;
  return { ...e, statusLabel: e.status[0].toUpperCase() + e.status.slice(1), attributes: n, graded: g, pctDone: n ? Math.round((g * 100) / n) : null };
};

export function saberMock(req, res, u, send, body) {
  let m;
  const p = u.pathname.replace('/api/saber', '');
  if (p === '/inspections') {
    const q = (u.searchParams.get('q') ?? '').toLowerCase(), y = Number(u.searchParams.get('year') || 0), st = u.searchParams.get('status') || 'all';
    const rows = inspections.map(row).filter((e) => (!q || [e.name, e.inspectedCommand, e.inspectingCommand].some((v) => (v ?? '').toLowerCase().includes(q))) &&
      (!y || [e.startDate, e.endDate].some((d) => d && Number(d.slice(0, 4)) === y)));
    const counts = { all: rows.length, active: 0, configuring: 0, closed: 0, archived: 0 };
    rows.forEach((e) => counts[e.status]++);
    return send(res, 200, { inspections: rows.filter((e) => st === 'all' || e.status === st), counts, years: [new Date().getFullYear(), 2025], status: st, classification: 'UNCLASSIFIED' });
  }
  if ((m = p.match(/^\/inspections\/(\d+)$/))) {
    const e = inspections.find((x) => x.id === Number(m[1]));
    if (!e) return send(res, 404, { error: 'No such inspection.' });
    const has = e.id !== 2;
    const editable = e.status === 'active';
    return send(res, 200, {
      inspection: { ...row(e), conductingCommand: 'TYCOM', closedAt: null }, classification: 'UNCLASSIFIED', bands,
      missionAreas: has ? missionAreas : [], groups: has ? groups : [], capabilities: has ? capabilities : [], attributes: has ? attributes : [],
      progress: { graded: has ? attributes.filter((a) => a.categorical).length : 0, total: has ? attributes.length : 0 },
      you: { stage: 4, stageLabel: 'Senior Member', canGrade: editable, whyNot: editable ? '' : 'Grading is only open while the inspection is Active.' },
    });
  }
  if ((m = p.match(/^\/inspections\/(\d+)\/grades\/(\d+)$/))) {
    const a = attributes.find((x) => x.id === Number(m[2]));
    if (!a) return send(res, 404, { error: 'No such attribute on this inspection.' });
    if (req.method === 'DELETE') {
      if (!a.categorical) return send(res, 200, { status: 200, ok: true, action: 'unchanged' });
      history.push({ id: history.length + 1, attr: a.id, action: 'cleared', oldScore: a.score, oldCategorical: a.categorical, newScore: null, newCategorical: null, occurredAt: now(), actor: me.name, actorRank: me.rank });
      Object.assign(a, { score: null, categorical: null, isPreliminary: null, gradedAt: null, gradedBy: null });
      return send(res, 200, { status: 200, ok: true, action: 'cleared' });
    }
    if ('band' in body) {
      if (body.band === NE) return send(res, 200, grade(a, null));
      const b = bands.find((x) => x.label === body.band);
      return b ? send(res, 200, grade(a, Math.floor((b.minScore + b.maxScore) / 2 + 0.5))) : send(res, 400, { error: 'No such band.' });
    }
    if (body.score !== null && (!Number.isInteger(body.score) || body.score < 0 || body.score > 100)) return send(res, 400, { error: 'Score must be a whole number from 0 to 100.' });
    return send(res, 200, grade(a, body.score));
  }
  if ((m = p.match(/^\/inspections\/(\d+)\/attributes\/(\d+)\/history$/))) {
    return send(res, 200, { history: history.filter((h) => h.attr === Number(m[2])).reverse() });
  }
  send(res, 404, { error: 'Not found.' });
}
