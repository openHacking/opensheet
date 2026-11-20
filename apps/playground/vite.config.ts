import { defineConfig } from 'vite';
import { aliases } from '../../scripts/aliases';
export default defineConfig({
  root: 'apps/playground',
  base: process.env.OPENSHEET_BASE_PATH ?? '/',
  resolve: { alias: aliases },
  // The file Workers load SheetJS on first use, outside the page dependency scan.
  optimizeDeps: { include: ['xlsx'] },
  server: { port: 5173, strictPort: true },
  build: { outDir: 'dist', emptyOutDir: true },
});
