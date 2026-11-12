import type { OpenSheet, PersistenceState } from 'opensheet';
import { $ } from './dom.js';
import type { PlaygroundState } from './state.js';

export function showPersistence(
  state: PersistenceState,
  playground: PlaygroundState,
  hasDraft = false,
  canRequestWriteLock = false,
) {
  const label = $('save-state');
  const reason =
    state.error?.code === 'STORAGE_BUDGET' || state.error?.code === 'QuotaExceededError'
      ? ' · Storage full'
      : state.error?.code === 'WORKER_FAILED' || state.error?.code === 'WORKER_TIMEOUT'
        ? ' · Reopen workbook'
        : '';
  label.textContent =
    state.error?.code === 'SYNC_UNAVAILABLE'
      ? 'Sync unavailable · Reopen workbook'
      : state.state === 'saving'
        ? 'Saving…'
        : state.state === 'failed'
          ? `Save unconfirmed${reason}`
          : hasDraft
            ? 'Editing…'
            : state.state === 'read-only'
              ? 'Read-only'
              : playground.importedWorkbook
                ? 'Imported locally · Saved'
                : 'Saved locally';
  label.title = state.error?.message ?? '';
  $<HTMLButtonElement>('save-dismiss').hidden = !state.acknowledgeable;
  $<HTMLButtonElement>('take-write-lock').hidden = !canRequestWriteLock;
}

export function refreshPersistence(app: OpenSheet, playground: PlaygroundState) {
  showPersistence(
    app.getWorkbook().persistenceState,
    playground,
    app.hasDraft,
    app.getWorkbook().canRequestWriteLock,
  );
}
