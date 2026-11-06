import { createContext, useContext, useEffect, useState, type ReactNode } from 'react';
import { kitApi, setStoreKey, store, stored, errText, type AppEntry, type MeResponse, type Person } from './api';
import { Card, Empty, PageHead } from './components';
import { PALETTES } from './palettes';

// The one shell every kit app runs in: sign-in (Windows, or the demo picker),
// an app switcher listing the apps the person has a profile in, the current
// app's pages in the sidebar, the top bar, the palette picker and toasts.
// Each app brings a KitAppModule; the host lists them.

export type NavItem = { path: string; label: string };   // path inside the app: "/", "/messages"
export type KitAppModule = {
  key: string;                // matches the server's app key: routes are #/<key>/<path>
  name: string;
  mark?: string;              // a character shown before the name
  nav: NavItem[];
  page: (path: string) => { node: ReactNode; active: string };
};
export type KitHostConfig = { title: string; apps: KitAppModule[]; defaultPalette?: string };

type Toast = (text: string, tone?: 'ok' | 'error') => void;
type AppContextValue = { module: KitAppModule; entry: AppEntry; me: MeResponse; go: (path: string) => void };
const ToastContext = createContext<Toast>(() => {});
const AppContext = createContext<AppContextValue | null>(null);
export const useToast = () => useContext(ToastContext);
export const useApp = () => useContext(AppContext)!;
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

export function KitHost({ config }: { config: KitHostConfig }) {
  setStoreKey(config.title);
  const hash = useHash();
  const [palette, setPalette] = usePalette(config.defaultPalette ?? PALETTES[0].key);
  const [me, setMe] = useState<MeResponse | null>(null);
  const [err, setErr] = useState('');
  const [toast, setToast] = useState<{ text: string; tone: 'ok' | 'error' } | null>(null);

  const loadMe = () => kitApi.me().then(setMe, (e) => setErr(errText(e)));
  useEffect(() => {
    document.title = config.title;
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
      <DemoSignIn title={config.title} onPick={(id) => { store('demoPerson', String(id)); loadMe(); }} />
    ) : (
      <div className="kit-signin">
        <h1>{config.title}</h1>
        <p>Your Windows account ({me.me.account}) isn't in the Personnel directory. Ask an admin to add you.</p>
      </div>
    );
  }

  // The apps this person may open: a profile in that app's own data file.
  const mine = config.apps
    .map((module) => ({ module, entry: me.apps.find((a) => a.key === module.key) }))
    .filter((x): x is { module: KitAppModule; entry: AppEntry } => !!x.entry && !!x.entry.profile && x.entry.available);
  const unavailable = me.apps.filter((a) => !a.available);

  const m = hash.match(/^#\/([a-z][a-z0-9-]*)(\/.*)?$/);
  // With one app, "#/" opens it; with several, "#/" lists them.
  const current = m ? mine.find((x) => x.module.key === m[1]) : mine.length === 1 ? mine[0] : undefined;
  const path = m?.[2] || '/';
  const page = current ? current.module.page(path) : null;
  const profile = (current?.entry.profile ?? {}) as Record<string, unknown>;
  const subtitle = [profile.billet, profile.unit].filter(Boolean).join(', ');
  const initial = me.me.username.replace(/^\W+/, '').charAt(0) || '?';

  return (
    <ToastContext.Provider value={say}>
      <div className="kit-shell">
        <aside className="kit-sidebar">
          <a className="kit-brand" href="#/">{config.title}</a>
          {mine.length > 1 && (
            <div className="kit-apps" role="navigation" aria-label="Apps">
              {mine.map(({ module }) => (
                <a key={module.key} href={`#/${module.key}/`} className={current?.module.key === module.key ? 'active' : ''}>
                  {module.mark && <span aria-hidden>{module.mark}</span>} {module.name}
                </a>
              ))}
            </div>
          )}
          {current && (
            <nav aria-label={current.module.name}>
              <div className="kit-nav-title">{current.module.mark} {current.module.name}</div>
              {current.module.nav.map((n) => (
                <a key={n.path} href={`#/${current.module.key}${n.path}`} className={page?.active === n.path ? 'active' : ''}
                  aria-current={page?.active === n.path ? 'page' : undefined}>
                  {n.label}
                </a>
              ))}
            </nav>
          )}
          <div className="kit-sidebar-foot">
            <label>
              Palette
              <select value={palette} onChange={(e) => setPalette(e.target.value)}>
                {PALETTES.map((p) => (
                  <option key={p.key} value={p.key}>{p.name}</option>
                ))}
              </select>
            </label>
            <div>Server: {me.server}</div>
          </div>
        </aside>
        <div className="kit-main">
          <header className="kit-topbar">
            <div className="kit-topbar-classif">UNCLASSIFIED</div>
            <div className="kit-who">
              <span className="kit-avatar" aria-hidden>{initial}</span>
              <div>
                <div className="kit-who-name">{me.me.username}</div>
                <div className="kit-who-billet">{subtitle}</div>
              </div>
              {me.demoSignIn && (
                <button className="kit-btn kit-btn-outline kit-btn-sm" onClick={() => { store('demoPerson', null); go('#/'); loadMe(); }}>
                  Switch person
                </button>
              )}
            </div>
          </header>
          {me.demoSignIn && <div className="kit-demo-strip">Demo sign-in: for review only. Real use signs in with your Windows account.</div>}
          <main>
            {current ? (
              <AppContext.Provider value={{ module: current.module, entry: current.entry, me, go: (p) => go(`#/${current.module.key}${p}`) }}>
                {page!.node}
              </AppContext.Provider>
            ) : (
              <AppsPage mine={mine} unavailable={unavailable} />
            )}
          </main>
        </div>
        {toast && (
          <div className={`kit-toast kit-toast-${toast.tone}`} role="status">
            {toast.text}
          </div>
        )}
      </div>
    </ToastContext.Provider>
  );
}

// "#/": the apps the person can open, and any that can't be reached.
function AppsPage({ mine, unavailable }: { mine: { module: KitAppModule; entry: AppEntry }[]; unavailable: AppEntry[] }) {
  return (
    <div className="kit-page">
      <PageHead title="Apps" />
      {mine.length === 0 ? (
        <Empty>You don't have a profile in any app yet. Ask an app's admin to add you.</Empty>
      ) : (
        <div className="kit-app-cards">
          {mine.map(({ module, entry }) => (
            <a key={module.key} className="kit-app-card" href={`#/${module.key}/`}>
              <span className="kit-app-card-mark" aria-hidden>{module.mark}</span>
              <strong>{module.name}</strong>
              <span className="kit-muted">{String((entry.profile as Record<string, unknown>).billet ?? '')}</span>
            </a>
          ))}
        </div>
      )}
      {unavailable.length > 0 && (
        <Card title="Unavailable right now">
          <ul className="kit-plain">
            {unavailable.map((a) => <li key={a.key}><strong>{a.name}</strong>: {a.error}</li>)}
          </ul>
        </Card>
      )}
    </div>
  );
}

function DemoSignIn({ title, onPick }: { title: string; onPick: (id: number) => void }) {
  const [people, setPeople] = useState<Person[] | null>(null);
  const [err, setErr] = useState('');
  useEffect(() => {
    kitApi.people().then((r) => setPeople(r.people), (e) => setErr(errText(e)));
  }, []);
  return (
    <div className="kit-signin">
      <h1>{title}</h1>
      <p>Demo sign-in. Pick the person to act as. Real use signs in with your Windows account.</p>
      {err && <p className="kit-error">{err}</p>}
      {!people ? (
        <p>Loading...</p>
      ) : people.length === 0 ? (
        <p>Nobody has a profile in any app's data file yet.</p>
      ) : (
        <ul className="kit-people">
          {people.map((p) => (
            <li key={p.PersonID}>
              <button className="kit-person" onClick={() => onPick(p.PersonID)}>
                <strong>{p.Name}</strong>
                <span className="kit-muted">{[p.Billet, p.Unit].filter(Boolean).join(', ')}</span>
                <span className="kit-muted">Apps: {p.Apps}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
