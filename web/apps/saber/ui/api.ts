// SABER's endpoints (web/apps/saber/api.ps1, under /api/saber/), over the kit client.
import { appApi } from '@kit';

const saber = appApi('saber');

export type Status = 'configuring' | 'active' | 'closed' | 'archived';
export const STATUSES: Status[] = ['active', 'configuring', 'closed', 'archived'];

export type InspectionRow = {
  id: number;
  name: string;
  inspectedCommand: string | null;
  inspectingCommand: string | null;
  status: Status;
  statusLabel: string;
  boatType: string | null;
  startDate: string | null;
  endDate: string | null;
  createdAt: string;
  attributes: number;
  graded: number;
  pctDone: number | null;
};

export type InspectionList = {
  inspections: InspectionRow[];
  counts: Record<'all' | Status, number>;
  years: number[];
  status: 'all' | Status;
  classification: string;
};

export type Band = { id: number; label: string; minScore: number; maxScore: number };
export type Node = { id: number; parentId?: number; title: string; code?: string | null; description?: string | null };
export type Attribute = {
  id: number;
  parentId: number;
  title: string | null;
  criterion: string;
  score: number | null;
  categorical: string | null;
  isPreliminary: boolean | null;
  gradedAt: string | null;
  gradedBy: string | null;
};

export type InspectionDetail = {
  inspection: Omit<InspectionRow, 'attributes' | 'graded' | 'pctDone' | 'createdAt'> & { conductingCommand: string | null; closedAt: string | null };
  classification: string;
  bands: Band[];
  missionAreas: Node[];
  groups: Node[];
  capabilities: Node[];
  attributes: Attribute[];
  progress: { graded: number; total: number };
  you: { stage: number; stageLabel: string; canGrade: boolean; whyNot: string };
};

export type GradeResult = { ok: boolean; action: 'graded' | 'changed' | 'cleared' | 'unchanged'; score?: number | null; categorical?: string };

export type HistoryRow = {
  id: number;
  action: string;
  oldScore: number | null;
  oldCategorical: string | null;
  newScore: number | null;
  newCategorical: string | null;
  occurredAt: string;
  actor: string | null;
  actorRank: string | null;
};

export const NE_LABEL = 'Not Evaluated';

export const api = {
  inspections: (f: { q?: string; year?: number | null; status?: string }) => {
    const p = new URLSearchParams();
    if (f.q) p.set('q', f.q);
    if (f.year) p.set('year', String(f.year));
    if (f.status && f.status !== 'all') p.set('status', f.status);
    return saber.get<InspectionList>(`/inspections?${p}`);
  },
  inspection: (id: number) => saber.get<InspectionDetail>(`/inspections/${id}`),
  // score: a whole number 0-100, or null for Not Evaluated. band: a band label (scored at its midpoint).
  grade: (id: number, attributeId: number, body: { score: number | null } | { band: string }) =>
    saber.put<GradeResult>(`/inspections/${id}/grades/${attributeId}`, body),
  clearGrade: (id: number, attributeId: number) => saber.del<GradeResult>(`/inspections/${id}/grades/${attributeId}`),
  history: (id: number, attributeId: number) => saber.get<{ history: HistoryRow[] }>(`/inspections/${id}/attributes/${attributeId}/history`),
};
