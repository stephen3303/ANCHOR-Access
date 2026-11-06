import { useEffect, useState } from 'react';
import { api, setDemoPerson, type MeResponse, type Person } from './api';
import { AuditPage, HomePage, MessagesPage } from './pages';
import { Empty } from './ui';

// Hash routes, so the server only ever serves one page: #/, #/messages,
// #/messages/12, #/audit.
function useRoute() {
  const [hash, setHash] = useState(location.hash || '#/');
  useEffect(() => {
    const on = () => setHash(location.hash || '#/');
    addEventListener('hashchange', on);
    return () => removeEventListener('hashchange', on);
  }, []);
  return hash;
}

const NAV = [
  { href: '#/', label: 'Home' },
  { href: '#/messages', label: 'Messages' },
  { href: '#/audit', label: 'Audit' },
];

export default function App() {
  const hash = useRoute();
  const [me, setMe] = useState<MeResponse | null>(null);
  const [err, setErr] = useState('');
  const [toast, setToast] = useState<{ text: string; tone: 'ok' | 'error' } | null>(null);

  const loadMe = () => api.me().then(setMe, (e) => setErr(String(e.message ?? e)));
  useEffect(() => {
    loadMe();
  }, []);
  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(t);
  }, [toast]);

  const go = (h: string) => {
    location.hash = h;
  };
  const say = (text: string, tone: 'ok' | 'error' = 'ok') => setToast({ text, tone });

  if (err) return <Empty>{err}</Empty>;
  if (!me) return <Empty>Signing in...</Empty>;
  if (!me.me.personId) {
    return me.demoSignIn ? (
      <DemoSignIn onPick={(id) => { setDemoPerson(id); loadMe(); }} />
    ) : (
      <div className="signin">
        <h1>ANCHOR</h1>
        <p>Your Windows account ({me.me.account}) has no active ANCHOR profile. Ask an ANCHOR admin to add you.</p>
      </div>
    );
  }

  const msgMatch = hash.match(/^#\/messages(?:\/(\d+))?$/);
  let page;
  if (msgMatch) page = <MessagesPage id={msgMatch[1] ? Number(msgMatch[1]) : undefined} toast={say} go={go} />;
  else if (hash === '#/audit') page = <AuditPage />;
  else page = <HomePage toast={say} go={go} />;
  const active = msgMatch ? '#/messages' : hash === '#/audit' ? '#/audit' : '#/';

  return (
    <div className="shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark" aria-hidden>{'⚓'}</span> ANCHOR
        </div>
        <nav>
          {NAV.map((n) => (
            <a key={n.href} href={n.href} className={active === n.href ? 'active' : ''}>
              {n.label}
            </a>
          ))}
        </nav>
        <div className="sidebar-foot">
          Web front end (prototype)
          <br />
          Data: {me.dataFile} on {me.server}
        </div>
      </aside>
      <div className="main">
        <header className="topbar">
          <div className="topbar-classif">UNCLASSIFIED</div>
          <div className="who">
            <span className="avatar" aria-hidden>{(me.me.username ?? '?').replace(/^\W+/, '').charAt(0)}</span>
            <div>
              <div className="who-name">{me.me.username}</div>
              <div className="who-billet">{[me.me.billet, me.me.unit].filter(Boolean).join(', ')}</div>
            </div>
            {me.demoSignIn && (
              <button className="btn btn-outline btn-sm" onClick={() => { setDemoPerson(null); loadMe(); }}>
                Switch person
              </button>
            )}
          </div>
        </header>
        {me.demoSignIn && <div className="demo-strip">Demo sign-in: for review only. Real use signs in with your Windows account.</div>}
        <main>{page}</main>
      </div>
      {toast && (
        <div className={`toast toast-${toast.tone}`} role="status">
          {toast.text}
        </div>
      )}
    </div>
  );
}

function DemoSignIn({ onPick }: { onPick: (id: number) => void }) {
  const [people, setPeople] = useState<Person[] | null>(null);
  const [err, setErr] = useState('');
  useEffect(() => {
    api.people().then((r) => setPeople(r.people), (e) => setErr(String(e.message ?? e)));
  }, []);
  return (
    <div className="signin">
      <h1>ANCHOR</h1>
      <p>Demo sign-in. Pick the person to act as. Real use signs in with your Windows account.</p>
      {err && <p className="error">{err}</p>}
      {!people ? (
        <p>Loading...</p>
      ) : people.length === 0 ? (
        <p>tblPeople is empty. Seed the data file with -Personnel first (backend\seed-backend.ps1).</p>
      ) : (
        <ul className="people">
          {people.map((p) => (
            <li key={p.PersonID}>
              <button className="person" onClick={() => onPick(p.PersonID)}>
                <strong>{p.Name}</strong>
                <span className="muted">{[p.Billet, p.Unit].filter(Boolean).join(', ')}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
