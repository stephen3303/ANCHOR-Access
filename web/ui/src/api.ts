// ANCHOR's endpoints (web/server/AnchorApi.ps1), over the kit client.
import { kitApi } from '@kit';

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

export type Summary = { total: number; unread: number; ackDueCount: number; ackDue: ListMessage[]; latest: ListMessage[]; commandName: string };

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
  AuditID: number; LoggedAt: string; UserName: string; Action: string; Resource: string | null; Classification: string | null;
  Result: string | null; Device: string | null; Location: string | null; Details: string | null; RowHash: string | null;
};
export type AuditResponse = { rows: AuditRow[]; chain: { ok: boolean; checked: number; problem: string } };

export const api = {
  summary: () => kitApi.get<Summary>('/api/summary'),
  messages: (f: { q?: string; precedence?: string; unread?: boolean; ackDue?: boolean }) => {
    const p = new URLSearchParams();
    if (f.q) p.set('q', f.q);
    if (f.precedence) p.set('precedence', f.precedence);
    if (f.unread) p.set('unread', '1');
    if (f.ackDue) p.set('ackDue', '1');
    return kitApi.get<{ messages: ListMessage[] }>(`/api/messages?${p}`);
  },
  message: (id: number) => kitApi.get<MessageDetail>(`/api/messages/${id}`),
  acknowledge: (id: number) => kitApi.post<{ ok: boolean }>(`/api/messages/${id}/ack`),
  audit: () => kitApi.get<AuditResponse>('/api/audit'),
};
