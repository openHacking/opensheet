import { defineConfig } from '@playwright/test';
import base from './playwright.config';

// Diagnostics may intentionally demonstrate browser-dependent failures; they are not a CI pass gate.
export default defineConfig({ ...base, testDir: 'tests/diagnostics' });
