import { createRequire } from 'node:module';
import { writeFileSync, mkdirSync } from 'node:fs';
import { snapshotSchema, commandSchemas } from '../packages/core/dist/index.js';
const { z } = createRequire(new URL('../packages/core/package.json', import.meta.url))('zod');
mkdirSync('schemas', { recursive: true });
writeFileSync(
  'schemas/workbook.schema.json',
  JSON.stringify(z.toJSONSchema(snapshotSchema), null, 2) + '\n',
);
writeFileSync(
  'schemas/commands.schema.json',
  JSON.stringify(
    {
      $schema: 'https://json-schema.org/draft/2020-12/schema',
      oneOf: Object.entries(commandSchemas).map(([type, payload]) => ({
        type: 'object',
        required: ['type', 'payload'],
        properties: { type: { const: type }, payload: z.toJSONSchema(payload) },
        additionalProperties: false,
      })),
    },
    null,
    2,
  ) + '\n',
);

writeFileSync(
  'api-manifest.json',
  JSON.stringify(
    {
      version: '0.1.0',
      schemaVersion: 1,
      pluginApiVersion: '0.1.0',
      entrypoints: {
        opensheet: ['createOpenSheet', 'OpenSheet'],
        '@opensheetjs/core': [
          'createWorkbook',
          'openWorkbook',
          'Workbook',
          'snapshotSchema',
          'commandSchemas',
        ],
        '@opensheetjs/adapter-sheetjs': ['fromSheetJS', 'toSheetJS'],
        '@opensheetjs/formats': [
          'exportRange',
          'streamExport',
          'parseDelimited',
          'stringifyDelimited',
        ],
        '@opensheetjs/plugin-sdk': ['definePlugin', 'PluginRegistry'],
      },
      commands: Object.keys(commandSchemas),
      errors: [
        'INVALID_ARGUMENT',
        'INVALID_RANGE',
        'READ_ONLY',
        'LIMIT_EXCEEDED',
        'UNSUPPORTED_FEATURE',
        'REVISION_CONFLICT',
        'PLUGIN_CONFLICT',
        'DISPOSED',
        'ABORTED',
        'BUSY',
        'STORAGE_BUDGET',
        'STORAGE_UNAVAILABLE',
        'CORRUPT_STORAGE',
        'WORKER_UNAVAILABLE',
        'WORKER_FAILED',
        'WORKER_TIMEOUT',
        'QuotaExceededError',
      ],
      docs: ['docs/api.md', 'docs/plugins.md', 'docs/implementation-status.md'],
      schemas: ['schemas/workbook.schema.json', 'schemas/commands.schema.json'],
    },
    null,
    2,
  ) + '\n',
);
