import { useEffect, useState } from 'react';
import './examgen.css';
import { appApi, Card, Empty, errText, PageHead, Stat, type KitAppModule } from '@kit';

// Exam Generator in the kit shell (read-only first cut): #/examgen/.
type Count = { status: string; n: number };
type Overview = {
  questionsByStatus: Count[];
  examsByStatus: Count[];
  topics: { id: number; name: string; defaultCount: number | null; questions: number }[];
  templates: { id: number; name: string }[];
  difficulty: { name: string }[];
  classLevels: { marking: string; portion: string }[];
};

const api = appApi('examgen');

function OverviewPage() {
  const [o, setO] = useState<Overview | null>(null);
  const [err, setErr] = useState('');
  useEffect(() => {
    api.get<Overview>('/overview').then(setO, (e) => setErr(errText(e)));
  }, []);
  if (err) return <Empty>{err}</Empty>;
  if (!o) return <Empty>Loading...</Empty>;
  const total = (xs: Count[]) => xs.reduce((a, x) => a + Number(x.n), 0);
  return (
    <div className="kit-page">
      <PageHead title="Overview">Question bank and exams in this Exam Generator data file.</PageHead>
      <div className="kit-stats">
        <Stat label="Current questions" value={total(o.questionsByStatus)} tone="info" />
        <Stat label="Exams" value={total(o.examsByStatus)} />
        <Stat label="Active topics" value={o.topics.length} />
        <Stat label="Templates" value={o.templates.length} />
      </div>
      <div className="eg-two">
        <Card title="Questions by status">
          <table className="kit-grid">
            <tbody>
              {o.questionsByStatus.map((x) => <tr key={x.status}><td>{x.status}</td><td className="eg-num">{x.n}</td></tr>)}
            </tbody>
          </table>
        </Card>
        <Card title="Exams by status">
          <table className="kit-grid">
            <tbody>
              {o.examsByStatus.map((x) => <tr key={x.status}><td>{x.status}</td><td className="eg-num">{x.n}</td></tr>)}
            </tbody>
          </table>
        </Card>
      </div>
      <Card title="Topics">
        {o.topics.length === 0 ? (
          <Empty>No topics yet.</Empty>
        ) : (
          <table className="kit-grid">
            <thead><tr><th>Topic</th><th className="eg-num">Questions</th><th className="eg-num">Default per exam</th></tr></thead>
            <tbody>
              {o.topics.map((t) => <tr key={t.id}><td>{t.name}</td><td className="eg-num">{t.questions}</td><td className="eg-num">{t.defaultCount ?? ''}</td></tr>)}
            </tbody>
          </table>
        )}
      </Card>
      <Card title="Lookups">
        <p className="kit-muted">Difficulty: {o.difficulty.map((d) => d.name).join(', ') || 'none'}</p>
        <p className="kit-muted">Markings: {o.classLevels.map((c) => `${c.marking} (${c.portion})`).join(', ') || 'none'}</p>
      </Card>
    </div>
  );
}

export const examGenApp: KitAppModule = {
  key: 'examgen',
  name: 'Exam Generator',
  mark: '✎',
  nav: [{ path: '/', label: 'Overview' }],
  page: () => ({ node: <OverviewPage />, active: '/' }),
};
