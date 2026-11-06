import { createContext, useContext, useEffect, useState, type ReactNode } from 'react';
import { kitApi, setAppKey, store, stored, errText, type MeResponse, type Person } from './api';
import { Empty } from './components';
import { PALETTES } from './palettes';

// The app shell every kit web app uses: sign-in (Windows, or the demo
// picker), sidebar with the app's pages, top bar, palette picker and toasts.
// An app gives its name, its nav and a function from the route to a page.

export type NavItem = { href: string; label: string };
export type KitAppConfig = {
  name: string;               // "ANCHOR"
  mark?: string;              // a character shown before the name
  nav: NavItem[];             // hash routes: "#/", "#/messages"
  page: (hash: string) => { node: ReactNode; active: string };
  defaultPalette?: string;    // a key from palettes.ts
};

type Toast = (text: string, tone?: 'ok' | 'error') => void;
const ToastContext = createContext<Toast>(() => {});
const MeContext = createContext<MeResponse | null>(null);
export const useToast = () => useContext(ToastContext);
export const useMe = () => useContext(MeContext)!;
export const go = (hash: string) => {
  location.hash = hash;
};

function useHash() {
  const [hash, setHash] = useState(location.hash || '#/');
  useEffect(() => {
    const on = () => setHash(location.hash || '#/');
    addEventListener('hashchange', on);
    return () => removeEventListener('hashchange', on);
  }, []);
  return hash;
}

function usePalette(fallback: string) {
  const [key, setKey] = useState(() => stored('palette') || fallback);
  useEffect(() => {
    document.documentElement.dataset.palette = key;
  }, [key]);
  return [key, (k: string) => { store('palette', k); setKey(k); }] as const;
}

export function KitApp({ config }: { config: KitAppConfig }) {
  setAppKey(config.name);
  const hash = useHash();
  const [palette, setPalette] = usePalette(config.defaultPalette ?? PALETTES[0].key);
  const [me, setMe] = useState<MeResponse | null>(null);
  const [err, setErr] = useState('');
  const [toast, setToast] = useState<{ text: string; tone: 'ok' | 'error' } | null>(null);

  const loadMe = () => kitApi.me().then(setMe, (e) => setErr(errText(e)));
  useEffect(() => {
    document.title = config.name;
    loadMe();
  }, []);
  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(t);
  }, [toast]);
  const say: Toast = (text, tone = 'ok') => setToast({ text, tone });

  if (err) return <Empty>{err}</Empty>;
  if (!me) return <Empty>Signing in...</Empty>;
  if (!me.me.personId) {
    return me.demoSignIn ? (
      <DemoSignIn app={config.name} onPick={(id) => { store('demoPerson', String(id)); loadMe(); }} />
    ) : (
      <div className="kit-signin">
        <h1>{config.name}</h1>
        <p>Your Windows account ({me.me.account}) has no active {config.name} profile. Ask a {config.name} admin to add you.</p>
      </div>
    );
  }

  const { node, active } = config.page(hash);
  const initial = (me.me.username ?? '?').replace(/^\W+/, '').charAt(0);
  return (
    <MeContext.Provider value={me}>
      <ToastContext.Provider value={say}>
        <div className="kit-shell">
          <aside className="kit-sidebar">
            <div className="kit-brand">
              {config.mark && <span className="kit-brand-mark" aria-hidden>{config.mark}</span>} {config.name}
            </div>
            <nav>
              {config.nav.map((n) => (
                <a key={n.href} href={n.href} className={active === n.href ? 'active' : ''} aria-current={active === n.href ? 'page' : undefined}>
                  {n.label}
                </a>
              ))}
            </nav>
            <div className="kit-sidebar-foot">
              <label>
                Palette
                <select value={palette} onChange={(e) => setPalette(e.target.value)}>
                  {PALETTES.map((p) => (
                    <option key={p.key} value={p.key}>{p.name}</option>
                  ))}
                </select>
              </label>
              <div>Data: {me.dataFile} on {me.server}</div>
            </div>
          </aside>
          <div className="kit-main">
            <header className="kit-topbar">
              <div className="kit-topbar-classif">UNCLASSIFIED</div>
              <div className="kit-who">
                <span className="kit-avatar" aria-hidden>{initial}</span>
                <div>
                  <div className="kit-who-name">{me.me.username}</div>
                  <div className="kit-who-billet">{[me.me.billet, me.me.unit].filter(Boolean).join(', ')}</div>
                </div>
                {me.demoSignIn && (
                  <button className="kit-btn kit-btn-outline kit-btn-sm" onClick={() => { store('demoPerson', null); loadMe(); }}>
                    Switch person
                  </button>
                )}
              </div>
            </header>
            {me.demoSignIn && <div className="kit-demo-strip">Demo sign-in: for review only. Real use signs in with your Windows account.</div>}
            <main>{node}</main>
          </div>
          {toast && (
            <div className={`kit-toast kit-toast-${toast.tone}`} role="status">
              {toast.text}
            </div>
          )}
        </div>
      </ToastContext.Provider>
    </MeContext.Provider>
  );
}

function DemoSignIn({ app, onPick }: { app: string; onPick: (id: number) => void }) {
  const [people, setPeople] = useState<Person[] | null>(null);
  const [err, setErr] = useState('');
  useEffect(() => {
    kitApi.people().then((r) => setPeople(r.people), (e) => setErr(errText(e)));
  }, []);
  return (
    <div className="kit-signin">
      <h1>{app}</h1>
      <p>Demo sign-in. Pick the person to act as. Real use signs in with your Windows account.</p>
      {err && <p className="kit-error">{err}</p>}
      {!people ? (
        <p>Loading...</p>
      ) : people.length === 0 ? (
        <p>Nobody has a {app} profile in this data file yet.</p>
      ) : (
        <ul className="kit-people">
          {people.map((p) => (
            <li key={p.PersonID}>
              <button className="kit-person" onClick={() => onPick(p.PersonID)}>
                <strong>{p.Name}</strong>
                <span className="kit-muted">{[p.Billet, p.Unit].filter(Boolean).join(', ')}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
