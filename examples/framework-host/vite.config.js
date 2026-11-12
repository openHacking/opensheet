import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

// Keep the module-relative storage Worker URL attached to the installed core package.
export default defineConfig({
  base: '/embedded/sheets/',
  build: { rollupOptions: { input: ['index.html', 'react.html', 'vue.html'] } },
  resolve: {
    alias: [
      { find: /^zod$/, replacement: fileURLToPath(new URL('./csp-zod.js', import.meta.url)) },
    ],
  },
  preview: {
    headers: {
      'Content-Security-Policy':
        "default-src 'none'; script-src 'self'; worker-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; font-src 'self'; base-uri 'none'; object-src 'none'; frame-ancestors 'none'",
    },
  },
  optimizeDeps: {
    exclude: [
      'opensheet',
      '@opensheetjs/core',
      '@opensheetjs/renderer',
      '@opensheetjs/formula',
      '@opensheetjs/plugin-sdk',
      '@opensheetjs/react',
      '@opensheetjs/vue',
    ],
  },
});
