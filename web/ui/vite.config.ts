import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { fileURLToPath } from 'node:url';

// One self-contained HTML file (script and styles inline), written to
// web/dist/index.html, which web/server/AnchorWeb.ps1 serves. The built file
// is committed, so the Windows machine that runs the server needs no Node.
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  // The kit web UI (web/kit/ui): the shell, components and palettes every app shares.
  resolve: { dedupe: ['react', 'react-dom'], alias: { '@kit': fileURLToPath(new URL('../kit/ui/index.ts', import.meta.url)) } },
  build: { outDir: '../dist', emptyOutDir: true },
  // npm run dev talks to the mock server (npm run mock), not to Access.
  server: { proxy: { '/api': 'http://localhost:8788' } },
});
