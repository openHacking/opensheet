import type { Selection } from '@opensheetjs/core';
export type GridOptions = {
  readOnly?: boolean;
  onSelection?: (selection: Selection) => void;
  onError?: (error: unknown) => void;
  onDraft?: (dirty: boolean) => void;
  onStatus?: (message: string) => void;
  onKeyboardReference?: () => boolean;
  referencePicker?: {
    begin(): boolean;
    update(selection: Selection): void;
    end(): void;
  };
};
