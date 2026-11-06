// A stand-in for web/server/AnchorWeb.ps1 for UI work on a machine without
// Access: the same /api/ routes over the seed rows in access/backend/seed,
// in memory. npm run mock, then npm run dev. Not used by the real server.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const seed = path.join(here, '../../access/backend/seed');
const tsv = (name) => {
  const [head, ...lines] = fs.readFileSync(path.join(seed, name), 'utf8').split(/\r?\n/).filter(Boolean);
  const cols = head.split('\t');
  return lines.map((l) => Object.fromEntries(l.split('\t').map((v, i) => [cols[i], v.replace(/\\n/g, '\n')])));
};
const iso = (s) => (s ? s.replace(' ', 'T') : null);
const messages = tsv('tblMessages.tsv').map((m, i) => ({ ...m, MessageID: i + 1, AckRequired: m.AckRequired === '1' }));
const byKey = Object.fromEntries(messages.map((m) => [m.MsgKey, m]));
const children = (file) => tsv(file).map((r) => ({ ...r, MessageID: byKey[r.MsgKey]?.MessageID }));
const sections = children('tblMessageSections.tsv');
const attachments = children('tblAttachments.tsv');
const addressees = children('tblAddressees.tsv');
const receipts = children('tblAckReceipts.tsv');
const audit = tsv('tblAudit.tsv').map((a, i) => ({ AuditID: i + 1, ...a, LoggedAt: iso(a.LoggedAt), RowHash: 'mock' }));
const state = new Map(); // MessageID -> { readState, ackOn }
const me = { personId: 1, account: 'MOCK\\demo01', username: 'LT Jordan Taylor', billet: 'Communications Officer', unit: 'COMSUBLANT N6', isAdmin: true };
const SRC = 'COMSUBLANT';

const list = (m) => {
  const s = state.get(m.MessageID);
  return {
    id: m.MessageID, readState: s?.readState ?? 'unread', precedence: m.Precedence, dtg: m.Dtg, reference: m.Reference,
    subject: m.Subject, originator: m.Originator, msgType: m.MsgType, classification: m.Classification, lifecycle: m.Lifecycle,
    ackDue: m.AckRequired && m.Source !== SRC && !s?.ackOn ? 1 : 0, ackOn: s?.ackOn ?? null, ackBy: iso(m.AckDue), receivedOn: iso(m.ReceivedOn),
  };
};
const log = (action, resource) =>
  audit.push({ AuditID: audit.length + 1, LoggedAt: new Date().toISOString().slice(0, 19), UserName: me.username, Action: action, Resource: resource, Classification: 'UNCLASSIFIED', Result: 'Success', Device: 'MOCK (web)', Location: me.unit, Details: '', RowHash: 'mock' });

const send = (res, code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
http.createServer((req, res) => {
  const u = new URL(req.url, 'http://x');
  const all = [...messages].sort((a, b) => b.ReceivedOn.localeCompare(a.ReceivedOn)).map(list);
  let m;
  if (u.pathname === '/') { res.writeHead(200, { 'Content-Type': 'text/html' }); return res.end(fs.readFileSync(path.join(here, '../dist/index.html'))); }
  if (u.pathname === '/api/me') return send(res, 200, { me: { personId: 1, account: me.account, username: me.username }, csrf: 'mock', demoSignIn: false, title: 'Navy Apps', server: 'MOCK',
    apps: [{ key: 'anchor', name: 'ANCHOR', available: true, error: '', profile: { billet: me.billet, unit: me.unit, isAdmin: true } },
           { key: 'examgen', name: 'Exam Generator', available: true, error: '', profile: { billet: 'SME, Approver', roles: ['SME', 'Approver'], isAdmin: false } }] });
  if (u.pathname === '/api/examgen/overview') return send(res, 200, {
    questionsByStatus: [{ status: 'Draft', n: 4 }, { status: 'In review', n: 3 }, { status: 'Approved', n: 41 }, { status: 'Retired', n: 2 }],
    examsByStatus: [{ status: 'Draft', n: 1 }, { status: 'Topic review', n: 2 }, { status: 'Approved', n: 6 }],
    topics: [{ id: 1, name: 'Reactor Plant Fundamentals', defaultCount: 10, questions: 18 }, { id: 2, name: 'Electrical Systems', defaultCount: 8, questions: 14 }, { id: 3, name: 'Damage Control', defaultCount: 6, questions: 16 }],
    templates: [{ id: 1, name: 'EOOW Qualification' }, { id: 2, name: 'Basic Engineering' }],
    difficulty: [{ name: 'Basic' }, { name: 'Intermediate' }, { name: 'Advanced' }],
    classLevels: [{ marking: 'UNCLASSIFIED', portion: 'U' }, { marking: 'CUI', portion: 'CUI' }] });
  if (u.pathname.startsWith('/api/anchor/')) u.pathname = u.pathname.replace('/api/anchor', '/api');
  if (u.pathname === '/api/summary') {
    const ackDue = all.filter((x) => x.ackDue);
    return send(res, 200, { total: all.length, unread: all.filter((x) => x.readState === 'unread').length, ackDueCount: ackDue.length, ackDue, latest: all.slice(0, 8), commandName: 'Commander, Submarine Force Atlantic' });
  }
  if (u.pathname === '/api/messages') {
    const q = (u.searchParams.get('q') ?? '').toLowerCase();
    const p = u.searchParams.get('precedence');
    return send(res, 200, { messages: all.filter((x) =>
      (!q || [x.subject, x.reference, x.originator, x.msgType].some((v) => (v ?? '').toLowerCase().includes(q))) &&
      (!p || x.precedence === p) && (u.searchParams.get('unread') !== '1' || x.readState === 'unread') &&
      (u.searchParams.get('ackDue') !== '1' || x.ackDue)) });
  }
  if ((m = u.pathname.match(/^\/api\/messages\/(\d+)$/))) {
    const msg = messages.find((x) => x.MessageID === Number(m[1]));
    if (!msg) return send(res, 404, { error: 'No such message.' });
    const s = state.get(msg.MessageID);
    if (!s) { state.set(msg.MessageID, { readState: 'read', ackOn: null }); log('View', `Message (${msg.Reference})`); } else s.readState = 'read';
    const st = state.get(msg.MessageID);
    const own = (arr) => arr.filter((r) => r.MessageID === msg.MessageID);
    return send(res, 200, { message: { ...msg, ReceivedOn: iso(msg.ReceivedOn), AckDue: iso(msg.AckDue) }, sections: own(sections).map((x) => ({ ...x, SortOrder: Number(x.SortOrder) })),
      attachments: own(attachments), addressees: own(addressees), receipts: own(receipts).map((r) => ({ ...r, ReceivedOn: iso(r.ReceivedOn) })),
      ackOn: st.ackOn, ackDue: msg.AckRequired && msg.Source !== SRC && !st.ackOn });
  }
  if ((m = u.pathname.match(/^\/api\/messages\/(\d+)\/ack$/)) && req.method === 'POST') {
    const msg = messages.find((x) => x.MessageID === Number(m[1]));
    const s = state.get(msg.MessageID) ?? { readState: 'read', ackOn: null };
    if (!msg.AckRequired || msg.Source === SRC || s.ackOn) return send(res, 409, { error: "This message doesn't need your acknowledgement." });
    s.ackOn = new Date().toISOString().slice(0, 19); s.readState = 'read'; state.set(msg.MessageID, s);
    log('Acknowledge', `Message (${msg.Reference})`);
    return send(res, 200, { status: 200, ok: true });
  }
  if (u.pathname === '/api/audit') return send(res, 200, { rows: [...audit].reverse(), chain: { ok: true, checked: audit.length, problem: '' } });
  send(res, 404, { error: 'Not found.' });
}).listen(8788, () => console.log('mock API on http://localhost:8788'));
