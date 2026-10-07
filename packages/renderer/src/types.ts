import type { Selection } from '@opensheetjs/core';
export type GridOptions = {
  readOnly?: boolean;
  onSelection?: (selection: Selection) => void;
  onError?: (error: unknown) => void;
  onStatus?: (message: string) => void;
};
