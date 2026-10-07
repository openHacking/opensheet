import { defineConfig } from 'vitest/config';
import { aliases } from './scripts/aliases';
export default defineConfig({
  resolve: { alias: aliases },
  test: { setupFiles: ['tests/storage-setup.ts'], include: ['tests/**/*.test.ts'] },
});
