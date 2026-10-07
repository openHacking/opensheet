import { createOpenSheet } from 'opensheet';
import 'opensheet/style.css';
import { createBindings } from './bindings.js';
import { createPerformance } from './performance.js';
import { createCodePreview } from './code-preview.js';
import { $ } from './dom.js';
import { createFeedback } from './feedback.js';
import { createFileActions } from './files.js';
import { renderInsights } from './insights.js';
import { createInteractions } from './interactions.js';
import { createWorkbench } from './layout.js';
import { mountPage } from './page.js';
import { createRouting } from './routing.js';
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
const refresh = () => {
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
);
const stopCommit = app.on('workbook:committed', () => {
  state.dirty = state.currentScene !== 'performance';
  $('save-state').textContent = 'Saved in IndexedDB · Available on this device';
  refresh();
});
const stopSelection = app.on('selection:changed', refresh);
const bindings = createBindings();
// Expose the public API for learning and browser contract tests.
Object.assign(window, { opensheet: app });
window.addEventListener(
  'beforeunload',
  (event) => {
    if (state.dirty) {
      event.preventDefault();
      event.returnValue = '';
    }
  },
  { signal: bindings.signal },
);
window.addEventListener(
  'pagehide',
  () => {
    files.dispose();
    preview.dispose();
    performanceLab.dispose();
    interactions.dispose();
    routing.dispose();
    workbench.dispose();
    feedback.dispose();
    stopCommit();
    stopSelection();
    bindings.dispose();
    app.dispose();
  },
  { signal: bindings.signal },
);
await routing.start();
