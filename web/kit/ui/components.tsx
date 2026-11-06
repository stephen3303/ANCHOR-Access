import { useEffect, useState, type ReactNode } from 'react';

// Shared pieces every app's pages use.

export type Tone = 'neutral' | 'info' | 'success' | 'warning' | 'danger' | 'solid-danger' | 'solid-warning';

export function Pill({ tone = 'neutral', children }: { tone?: Tone; children: ReactNode }) {
  return <span className={`kit-pill kit-pill-${tone}`}>{children}</span>;
}

// Markings above UNCLASSIFIED are for demonstration only in these apps
// (Stephen, 2026-11-05): anything else shows this banner with it.
export function ClassificationBanner({ value }: { value: string | null | undefined }) {
  const v = (value || 'UNCLASSIFIED').toUpperCase();
  if (v === 'UNCLASSIFIED') return <div className="kit-classif kit-classif-u">UNCLASSIFIED</div>;
  return (
    <div className="kit-classif kit-classif-demo" role="note">
      {v} -- DEMONSTRATION MARKING ONLY. CONTENT IS UNCLASSIFIED.
    </div>
  );
}

export function Card({ title, action, className = '', children }: { title?: string; action?: ReactNode; className?: string; children: ReactNode }) {
  return (
    <section className={`kit-card ${className}`}>
      {(title || action) && (
        <header className="kit-card-head">
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
    <div className={`kit-stat ${tone ? `kit-stat-${tone}` : ''}`}>
      <div className="kit-stat-value">{value}</div>
      <div className="kit-stat-label">{label}</div>
    </div>
  );
}

export function Empty({ children }: { children: ReactNode }) {
  return <div className="kit-empty">{children}</div>;
}

export function PageHead({ title, children }: { title: string; children?: ReactNode }) {
  return (
    <div className="kit-page-head">
      <h1>{title}</h1>
      {children && <p className="kit-muted">{children}</p>}
    </div>
  );
}

const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

// "24 OCT 26 1830" from the server's ISO text.
export function fmtDate(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getDate())} ${MONTHS[d.getMonth()]} ${String(d.getFullYear()).slice(2)} ${p(d.getHours())}${p(d.getMinutes())}`;
}

// "Due in 5 h", "Overdue by 2 d".
export function dueText(iso: string | null | undefined): { text: string; overdue: boolean } {
  if (!iso) return { text: '', overdue: false };
  const ms = new Date(iso).getTime() - Date.now();
  const abs = Math.abs(ms);
  const span = abs >= 86400000 ? `${Math.round(abs / 86400000)} d` : `${Math.max(1, Math.round(abs / 3600000))} h`;
  return ms < 0 ? { text: `Overdue by ${span}`, overdue: true } : { text: `Due in ${span}`, overdue: false };
}

// True on a phone-width window.
export function useNarrow(maxWidth = 760) {
  const q = `(max-width: ${maxWidth}px)`;
  const [narrow, setNarrow] = useState(() => matchMedia(q).matches);
  useEffect(() => {
    const mq = matchMedia(q);
    const on = () => setNarrow(mq.matches);
    mq.addEventListener('change', on);
    return () => mq.removeEventListener('change', on);
  }, [q]);
  return narrow;
}
