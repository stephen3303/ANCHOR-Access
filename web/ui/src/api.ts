// Calls to web/server/AnchorWeb.ps1. The browser signs in with Windows on its
// own (the server asks for Negotiate/NTLM); writes also carry the token that
// /api/me hands out.

export type ListMessage = {
  id: number;
  readState: 'read' | 'unread';
  precedence: string;
  dtg: string | null;
  reference: string;
  subject: string;
  originator: string | null;
  msgType: string | null;
  classification: string;
  lifecycle: string;
  ackDue: number;
  ackOn: string | null;
  ackBy: string | null;
  receivedOn: string;
};

export type Me = {
  personId: number;
  account: string;
  username?: string;
  billet?: string;
  unit?: string;
  isAdmin?: boolean;
};

export type MeResponse = { me: Me; csrf: string; demoSignIn: boolean; server: string; dataFile: string };

export type Summary = {
  total: number;
  unread: number;
  ackDueCount: number;
  ackDue: ListMessage[];
  latest: ListMessage[];
  commandName: string;
};

export type MessageDetail = {
  message: Record<string, string | number | boolean | null>;
  sections: { SortOrder: number; Heading: string; Body: string }[];
  attachments: { FileName: string; SizeText: string; Kind: string }[];
  addressees: { AddrName: string; AddrRole: string; AddrKind: string }[];
  receipts: { CommandName: string; ReceivedOn: string }[];
  ackOn: string | null;
  ackDue: boolean;
};

export type AuditRow = {
  AuditID: number;
  LoggedAt: string;
  UserName: string;
  Action: string;
  Resource: string | null;
  Classification: string | null;
  Result: string | null;
  Device: string | null;
  Location: string | null;
  Details: string | null;
  RowHash: string | null;
};

export type AuditResponse = { rows: AuditRow[]; chain: { ok: boolean; checked: number; problem: string } };

export type Person = { PersonID: number; Billet: string | null; Unit: string | null; Name: string };

let csrf = '';
const DEMO_KEY = 'anchor.demoPerson';

export function demoPerson(): string {
  try {
    return localStorage.getItem(DEMO_KEY) ?? '';
  } catch {
    return '';
  }
}

export function setDemoPerson(id: number | null) {
  try {
    if (id === null) localStorage.removeItem(DEMO_KEY);
    else localStorage.setItem(DEMO_KEY, String(id));
  } catch {
    /* private window: the choice lasts until reload */
  }
}

export class ApiError extends Error {
  constructor(public status: number, message: string) {
    super(message);
  }
}

async function call<T>(path: string, init: RequestInit = {}): Promise<T> {
  const headers: Record<string, string> = { Accept: 'application/json' };
  const as = demoPerson();
  if (as) headers['X-Anchor-As'] = as;
  if (init.method && init.method !== 'GET') headers['X-Anchor-Csrf'] = csrf;
  const res = await fetch(path, { ...init, headers, credentials: 'same-origin' });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new ApiError(res.status, body.error ?? `The server answered ${res.status}.`);
  return body as T;
}

export const api = {
  async me() {
    const r = await call<MeResponse>('/api/me');
    csrf = r.csrf;
    return r;
  },
  people: () => call<{ people: Person[] }>('/api/people'),
  summary: () => call<Summary>('/api/summary'),
  messages: (f: { q?: string; precedence?: string; unread?: boolean; ackDue?: boolean }) => {
    const p = new URLSearchParams();
    if (f.q) p.set('q', f.q);
    if (f.precedence) p.set('precedence', f.precedence);
    if (f.unread) p.set('unread', '1');
    if (f.ackDue) p.set('ackDue', '1');
    return call<{ messages: ListMessage[] }>(`/api/messages?${p}`);
  },
  message: (id: number) => call<MessageDetail>(`/api/messages/${id}`),
  acknowledge: (id: number) => call<{ ok: boolean }>(`/api/messages/${id}/ack`, { method: 'POST' }),
  audit: () => call<AuditResponse>('/api/audit'),
};
