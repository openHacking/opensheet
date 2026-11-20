import { createOpenSheet } from 'opensheet';
import { verifyBackupInWorker } from './backup-verifier.js';
import 'opensheet/style.css';
import { createBindings } from './bindings.js';
import { createPerformance } from './performance.js';
import { createCodePreview } from './code-preview.js';
import { $ } from './dom.js';
import { bindFeatureGuide } from './feature-guide.js';
import { createFeedback } from './feedback.js';
import { createFileActions } from './files.js';
import { renderInsights } from './insights.js';
import { createInteractions } from './interactions.js';
import { createWorkbench } from './layout.js';
import { mountPage } from './page.js';
import { showPersistence } from './persistence.js';
import { createRouting, RetainedBackupError } from './routing.js';
import { createState } from './state.js';
import './style.css';
import { sumPlugin } from './sum-plugin.js';

mountPage();
const state = createState();
const feedback = createFeedback();
const app = createOpenSheet({
  container: '#spreadsheet',
  onError: (e) => feedback.notice(e instanceof Error ? e.message : String(e)),
});
app.use(sumPlugin);
const workbench = createWorkbench();
const stopFeatureGuide = bindFeatureGuide(app, feedback.notice);
const refresh = () => {
  const showGuide = state.currentScene === 'features' && !state.importedWorkbook;
  $('feature-guide').hidden = !showGuide;
  $('scene-insights').hidden = showGuide;
  $('scene-hint').hidden = showGuide;
  $('tab-insights').textContent = showGuide ? 'Try features' : 'Data view';
  if (showGuide) {
    preview.refresh();
    return;
  }
  if (state.currentScene === 'performance') {
    performanceLab.update();
    return;
  }
  void renderInsights(app, state).catch((e) => feedback.notice(String(e)));
  preview.refresh();
};
const files = createFileActions(app, state, feedback, refresh);
const performanceLab = createPerformance(app, state, files.download);
const preview = createCodePreview(app, feedback, files.download);
const interactions = createInteractions(app, state, feedback);
const routing = createRouting(
  app,
  state,
  files.cancelImport,
  refresh,
  workbench.showPanel,
  performanceLab.start,
  performanceLab.leave,
  feedback.notice,
);
const stopCommit = app.on('workbook:committed', () => {
  state.dirty = state.currentScene !== 'performance';
  refresh();
});
let pendingWrites = 0;
const stopDraft = app.on('editor:draft', () => {
  showPersistence(
    app.getWorkbook().persistenceState,
    state,
    app.hasDraft,
    app.getWorkbook().canRequestWriteLock,
  );
});
const stopPersistence = app.on('workbook:persistence', (value) => {
  pendingWrites = value.pending;
  showPersistence(value, state, app.hasDraft, app.getWorkbook().canRequestWriteLock);
});
const stopSelection = app.on('selection:changed', refresh);
const bindings = createBindings();
bindings.on($('save-dismiss'), 'click', () => app.getWorkbook().acknowledgePersistenceFailure());
bindings.on($('take-write-lock'), 'click', async () => {
  const button = $<HTMLButtonElement>('take-write-lock');
  button.disabled = true;
  try {
    const acquired = await app.getWorkbook().tryAcquireWriteLock();
    feedback.notice(acquired ? 'Editing enabled' : 'Another tab is still editing');
  } catch (error) {
    feedback.notice(error instanceof Error ? error.message : String(error));
  } finally {
    button.disabled = false;
  }
});
// Expose the public API for learning and browser contract tests.
Object.assign(window, { opensheet: app });
window.addEventListener(
  'beforeunload',
  (event) => {
    if (state.dirty || pendingWrites > 0 || app.hasDraft) {
      event.preventDefault();
      event.returnValue = '';
    }
  },
  { signal: bindings.signal },
);
window.addEventListener(
  'pagehide',
  (event) => {
    // A cached page can become active again without running this module anew.
    if (event.persisted) return;
    files.dispose();
    preview.dispose();
    performanceLab.dispose();
    interactions.dispose();
    routing.dispose();
    workbench.dispose();
    stopFeatureGuide();
    feedback.dispose();
    stopCommit();
    stopPersistence();
    stopDraft();
    stopSelection();
    bindings.dispose();
    app.dispose();
  },
  { signal: bindings.signal },
);
const recovery = document.createElement('section');
recovery.className = 'startup-recovery';
recovery.hidden = true;
recovery.setAttribute('role', 'alert');
recovery.innerHTML = `<div><h1>Unable to open workbook</h1><p id="startup-recovery-message"></p><p>Check browser storage access or restore a backup if storage is damaged.</p><button class="button primary" type="button">Try again</button> <button class="button secondary" id="open-retained-backup" type="button" hidden>Open retained backup</button> <button class="button secondary" id="restore-backup" type="button">Restore backup</button> <button class="button secondary" id="verify-backup" type="button">Verify backup</button> <button class="button secondary" id="cancel-verification" type="button" hidden>Cancel verification</button> <button class="button secondary" id="recovery-budget" type="button" hidden>Open Budget Simulator</button><input id="restore-backup-file" type="file" accept=".opensheet,.json" hidden><input id="verify-backup-file" type="file" accept=".opensheet,.json" hidden><p id="verify-backup-status" role="status" aria-live="polite"></p></div>`;
document.body.append(recovery);
const retry = recovery.querySelector<HTMLButtonElement>('button')!;
const root = $<HTMLDivElement>('app');
const restore = $<HTMLButtonElement>('restore-backup');
const openRetained = $<HTMLButtonElement>('open-retained-backup');
let retainedBackupId: string | undefined;
const backupFile = $<HTMLInputElement>('restore-backup-file');
const verify = $<HTMLButtonElement>('verify-backup');
const cancelVerification = $<HTMLButtonElement>('cancel-verification');
const verifyFile = $<HTMLInputElement>('verify-backup-file');
const verifyStatus = $('verify-backup-status');
const recoveryBudget = $<HTMLButtonElement>('recovery-budget');
recoveryBudget.hidden = location.hash !== '#performance';
if (!recoveryBudget.hidden) restore.hidden = true;
recoveryBudget.addEventListener(
  'click',
  () => {
    location.hash = 'budget';
    location.reload();
  },
  { signal: bindings.signal },
);
const restoreController = new AbortController();
bindings.signal.addEventListener('abort', () => restoreController.abort(), { once: true });
async function start() {
  retry.disabled = true;
  restore.disabled = true;
  verify.disabled = true;
  openRetained.disabled = true;
  try {
    await routing.start();
    recovery.hidden = true;
    root.inert = false;
    retainedBackupId = undefined;
    openRetained.hidden = true;
  } catch (error) {
    const code = (error as { code?: string })?.code;
    const message = error instanceof Error ? error.message : String(error);
    $('startup-recovery-message').textContent = code ? `${code}: ${message}` : message;
    root.inert = true;
    recovery.hidden = false;
    retry.disabled = false;
    retry.focus();
  } finally {
    retry.disabled = false;
    restore.disabled = false;
    verify.disabled = false;
    openRetained.disabled = false;
  }
}
openRetained.addEventListener(
  'click',
  async () => {
    if (!retainedBackupId) return;
    retry.disabled = restore.disabled = verify.disabled = openRetained.disabled = true;
    try {
      await routing.openRetainedBackup(retainedBackupId);
      recovery.hidden = true;
      root.inert = false;
      retainedBackupId = undefined;
      openRetained.hidden = true;
      app.getGrid()?.element.focus();
      noticeRestoreSuccess();
    } catch (error) {
      $('startup-recovery-message').textContent =
        error instanceof Error ? error.message : String(error);
    } finally {
      retry.disabled = restore.disabled = verify.disabled = openRetained.disabled = false;
    }
  },
  { signal: bindings.signal },
);
restore.addEventListener('click', () => backupFile.click(), { signal: bindings.signal });
backupFile.addEventListener(
  'change',
  async () => {
    const file = backupFile.files?.[0];
    if (!file) return;
    retry.disabled = true;
    restore.disabled = true;
    verify.disabled = true;
    openRetained.disabled = true;
    try {
      await routing.restoreBackup(file, restoreController.signal);
      recovery.hidden = true;
      root.inert = false;
      retainedBackupId = undefined;
      openRetained.hidden = true;
      noticeRestoreSuccess();
    } catch (error) {
      if (error instanceof RetainedBackupError) {
        retainedBackupId = error.workbookId;
        openRetained.hidden = false;
      }
      $('startup-recovery-message').textContent =
        error instanceof Error ? error.message : String(error);
    } finally {
      backupFile.value = '';
      retry.disabled = false;
      restore.disabled = false;
      verify.disabled = false;
      openRetained.disabled = false;
    }
  },
  { signal: bindings.signal },
);
let verificationController: AbortController | undefined;
let cancelledFocus = false;
bindings.signal.addEventListener('abort', () => verificationController?.abort(), { once: true });
cancelVerification.addEventListener(
  'click',
  () => {
    cancelledFocus = true;
    verificationController?.abort();
    cancelVerification.disabled = true;
    verifyStatus.textContent = 'Cancelling verification…';
  },
  { signal: bindings.signal },
);
verify.addEventListener('click', () => verifyFile.click(), { signal: bindings.signal });
verifyFile.addEventListener(
  'change',
  async () => {
    const file = verifyFile.files?.[0];
    if (!file) return;
    retry.disabled = true;
    restore.disabled = true;
    verify.disabled = true;
    recoveryBudget.disabled = true;
    openRetained.disabled = true;
    verifyStatus.textContent = 'Checking backup…';
    const controller = new AbortController();
    verificationController = controller;
    cancelledFocus = false;
    cancelVerification.hidden = false;
    cancelVerification.disabled = false;
    cancelVerification.focus();
    try {
      if (!/\.(opensheet|json)$/i.test(file.name))
        throw new Error('Choose a native or JSON backup');
      const summary = await verifyBackupInWorker(file, controller.signal);
      verifyStatus.textContent = `Structure verified: ${summary.sheets} sheets, ${summary.cells} cells. Completeness is not confirmed.`;
    } catch (error) {
      verifyStatus.textContent = controller.signal.aborted
        ? 'Verification cancelled.'
        : error instanceof Error
          ? error.message
          : String(error);
    } finally {
      const returnFocus =
        document.activeElement === cancelVerification ||
        (cancelledFocus &&
          (document.activeElement === document.body ||
            document.activeElement === document.documentElement));
      cancelledFocus = false;
      verificationController = undefined;
      cancelVerification.hidden = true;
      verifyFile.value = '';
      retry.disabled = false;
      restore.disabled = false;
      verify.disabled = false;
      recoveryBudget.disabled = false;
      openRetained.disabled = false;
      if (returnFocus && !bindings.signal.aborted) verify.focus();
    }
  },
  { signal: bindings.signal },
);
function noticeRestoreSuccess() {
  feedback.notice('Backup restored locally');
}
retry.addEventListener('click', () => void start(), { signal: bindings.signal });
await start();
