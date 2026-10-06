import { defineConfig } from 'vite';
import { aliases } from '../../scripts/aliases';
export default defineConfig({
  root: 'apps/playground',
  base: process.env.OPENSHEET_BASE_PATH ?? '/',
  resolve: { alias: aliases },
  server: { port: 5173, strictPort: true },
  build: { outDir: 'dist', emptyOutDir: true },
});
