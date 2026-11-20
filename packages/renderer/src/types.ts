import type { Selection } from '@opensheetjs/core';
export type GridTarget =
  | { kind: 'cell'; row: number; column: number }
  | { kind: 'row' | 'column'; index: number };

export type GridOptions = {
  readOnly?: boolean;
  onContextMenu?: (target: GridTarget, point: { x: number; y: number }) => void;
  preserveContextFocus?: () => boolean;
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
