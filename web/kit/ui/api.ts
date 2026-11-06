// Calls to the kit web server (kit/server/KitWeb.ps1). The browser signs in
// with Windows on its own (the server asks for Negotiate/NTLM); writes also
// carry the token /api/me hands out.

export type Me = {
  personId: number;
  account: string;
  username?: string;
  billet?: string | null;
  unit?: string | null;
  isAdmin?: boolean;
  [key: string]: unknown;
};

export type MeResponse = { me: Me; csrf: string; demoSignIn: boolean; app: string; server: string; dataFile: string };
export type Person = { PersonID: number; Billet: string | null; Unit: string | null; Name: string };

let csrf = '';
let appKey = 'kit';

export class ApiError extends Error {
  constructor(public status: number, message: string) {
    super(message);
  }
}

// Per-viewer conveniences (demo person, palette). Storage can be blocked.
export function stored(name: string): string {
  try {
    return localStorage.getItem(`${appKey}.${name}`) ?? '';
  } catch {
    return '';
  }
}
export function store(name: string, value: string | null) {
  try {
    if (value === null) localStorage.removeItem(`${appKey}.${name}`);
    else localStorage.setItem(`${appKey}.${name}`, value);
  } catch {
    /* lasts until reload */
  }
}

export function setAppKey(key: string) {
  appKey = key.toLowerCase();
}

async function call<T>(method: string, path: string, body?: unknown): Promise<T> {
  const headers: Record<string, string> = { Accept: 'application/json' };
  const as = stored('demoPerson');
  if (as) headers['X-Kit-As'] = as;
  if (method !== 'GET') headers['X-Kit-Csrf'] = csrf;
  if (body !== undefined) headers['Content-Type'] = 'application/json';
  const res = await fetch(path, { method, headers, credentials: 'same-origin', body: body === undefined ? undefined : JSON.stringify(body) });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new ApiError(res.status, data.error ?? `The server answered ${res.status}.`);
  return data as T;
}

export const kitApi = {
  get: <T>(path: string) => call<T>('GET', path),
  post: <T>(path: string, body?: unknown) => call<T>('POST', path, body),
  put: <T>(path: string, body?: unknown) => call<T>('PUT', path, body),
  del: <T>(path: string) => call<T>('DELETE', path),
  async me() {
    const r = await call<MeResponse>('GET', '/api/me');
    csrf = r.csrf;
    return r;
  },
  people: () => call<{ people: Person[] }>('GET', '/api/people'),
};

export function errText(e: unknown) {
  return e instanceof Error ? e.message : String(e);
}
