import type { ReactNode } from 'react';

// Small shared pieces: pills, the classification banner, cards, dates.

export function PrecedencePill({ value }: { value: string }) {
  return <span className={`pill prec-${value.toLowerCase()}`}>{value}</span>;
}

export function LifecyclePill({ value }: { value: string }) {
  if (value === 'ACTIVE') return null;
  return <span className={`pill life-${value.toLowerCase()}`}>{value}</span>;
}

// Stephen 2026-11-05: ANCHOR allows markings up to SECRET for demonstration
// only; anything other than UNCLASSIFIED shows this banner with it.
export function ClassificationBanner({ value }: { value: string }) {
  const v = (value || 'UNCLASSIFIED').toUpperCase();
  if (v === 'UNCLASSIFIED') return <div className="classif classif-u">UNCLASSIFIED</div>;
  return (
    <div className="classif classif-demo" role="note">
      {v} -- DEMONSTRATION MARKING ONLY. CONTENT IS UNCLASSIFIED.
    </div>
  );
}

export function Card({ title, action, children }: { title?: string; action?: ReactNode; children: ReactNode }) {
  return (
    <section className="card">
      {(title || action) && (
        <header className="card-head">
          {title && <h2>{title}</h2>}
          {action}
        </header>
      )}
      {children}
    </section>
  );
}

export function Stat({ label, value, tone }: { label: string; value: number | string; tone?: 'warn' | 'info' | 'ok' }) {
  return (
    <div className={`stat ${tone ? `stat-${tone}` : ''}`}>
      <div className="stat-value">{value}</div>
      <div className="stat-label">{label}</div>
    </div>
  );
}

const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

// "24 OCT 26 1830" from the server's ISO text (local time of the data).
export function fmtDate(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getDate())} ${MONTHS[d.getMonth()]} ${String(d.getFullYear()).slice(2)} ${p(d.getHours())}${p(d.getMinutes())}`;
}

// "Due in 5 h", "Overdue by 2 d" (modMessages.AckDueText).
export function dueText(iso: string | null | undefined): { text: string; overdue: boolean } {
  if (!iso) return { text: '', overdue: false };
  const ms = new Date(iso).getTime() - Date.now();
  const abs = Math.abs(ms);
  const span = abs >= 86400000 ? `${Math.round(abs / 86400000)} d` : `${Math.max(1, Math.round(abs / 3600000))} h`;
  return ms < 0 ? { text: `Overdue by ${span}`, overdue: true } : { text: `Due in ${span}`, overdue: false };
}

export function Empty({ children }: { children: ReactNode }) {
  return <div className="empty">{children}</div>;
}
