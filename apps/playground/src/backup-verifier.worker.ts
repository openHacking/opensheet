import { verifyBinary, verifyJSON } from '@opensheetjs/core';

self.onmessage = async (event: MessageEvent<{ file: File; json: boolean }>) => {
  try {
    const summary = event.data.json
      ? await verifyJSON(event.data.file)
      : await verifyBinary(event.data.file);
    self.postMessage({ summary });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : String(error) });
  }
};
