import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { KitHost, type KitHostConfig } from '@kit';
import { anchorApp } from '../../apps/anchor/ui';
import { examGenApp } from '../../apps/examgen/ui';
import { saberApp } from '../../apps/saber/ui';

// Every app's pages in one UI. The server's settings file decides which apps
// (and which data files) a copy serves; the shell shows each person only the
// apps they have a profile in. Add an app: its module here and in web.config.json.
const config: KitHostConfig = {
  title: 'Navy Apps',
  apps: [anchorApp, examGenApp, saberApp],
};

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <KitHost config={config} />
  </StrictMode>,
);
