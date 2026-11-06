import type { KitAppModule } from '@kit';
import { AuditPage, HomePage, MessagesPage } from './pages';
import './anchor.css';

// ANCHOR in the kit shell: #/anchor/, #/anchor/messages, #/anchor/messages/12, #/anchor/audit.
export const anchorApp: KitAppModule = {
  key: 'anchor',
  name: 'ANCHOR',
  mark: '⚓',
  nav: [
    { path: '/', label: 'Home' },
    { path: '/messages', label: 'Messages' },
    { path: '/audit', label: 'Audit' },
  ],
  page: (path) => {
    const m = path.match(/^\/messages(?:\/(\d+))?$/);
    if (m) return { node: <MessagesPage id={m[1] ? Number(m[1]) : undefined} />, active: '/messages' };
    if (path === '/audit') return { node: <AuditPage />, active: '/audit' };
    return { node: <HomePage />, active: '/' };
  },
};
