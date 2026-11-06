import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { KitApp, type KitAppConfig } from '@kit';
import { AuditPage, HomePage, MessagesPage } from './pages';
import './anchor.css';

// ANCHOR's pages on the kit shell. Hash routes: #/, #/messages, #/messages/12, #/audit.
const config: KitAppConfig = {
  name: 'ANCHOR',
  mark: '⚓',
  nav: [
    { href: '#/', label: 'Home' },
    { href: '#/messages', label: 'Messages' },
    { href: '#/audit', label: 'Audit' },
  ],
  page: (hash) => {
    const m = hash.match(/^#\/messages(?:\/(\d+))?$/);
    if (m) return { node: <MessagesPage id={m[1] ? Number(m[1]) : undefined} />, active: '#/messages' };
    if (hash === '#/audit') return { node: <AuditPage />, active: '#/audit' };
    return { node: <HomePage />, active: '#/' };
  },
};

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <KitApp config={config} />
  </StrictMode>,
);
