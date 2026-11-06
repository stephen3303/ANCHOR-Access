import type { KitAppModule } from '@kit';
import { InspectionPage, InspectionsPage } from './pages';
import './saber.css';

// SABER in the kit shell: #/saber/ (Inspections), #/saber/inspections/12 (grading).
// Later phases add Comments, Review, Progress, Caucus, Reports and Admin (PORT-PLAN.md).
export const saberApp: KitAppModule = {
  key: 'saber',
  name: 'SABER',
  mark: '◈',
  nav: [{ path: '/', label: 'Inspections' }],
  page: (path) => {
    const m = path.match(/^\/inspections\/(\d+)$/);
    if (m) return { node: <InspectionPage key={m[1]} id={Number(m[1])} />, active: '/' };
    return { node: <InspectionsPage />, active: '/' };
  },
};
