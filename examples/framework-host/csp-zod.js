// Execute before dependent schemas are constructed, in both page and Worker bundles.
import { z } from 'zod/v4';
z.config({ jitless: true });
export * from 'zod/v4';
