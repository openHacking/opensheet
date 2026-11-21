import { createWorkbook, Workbook, type OpenSheet } from 'opensheet';
import { createBindings } from './bindings.js';
import { sample } from './budget.js';
import { $ } from './dom.js';
import { createScene, scenes, type SceneId } from './scenes.js';
import type { PlaygroundState } from './state.js';
import { refreshPersistence } from './persistence.js';
export class RetainedBackupError extends Error {
  constructor(
    readonly workbookId: string,
    message: string,
    cause: unknown,
  ) {
    super(message, { cause });
    this.name = 'RetainedBackupError';
  }
}
export function createRouting(
  app: OpenSheet,
  state: PlaygroundState,
  cancelImport: () => void,
  refresh: () => void,
  showPanel: (panel: 'code' | 'insights') => void,
  startPerformance: () => Promise<void>,
  leavePerformance: () => Promise<void>,
  notice: (message: string) => void,
) {
  const bindings = createBindings();
  function showPage(page: 'spreadsheet' | 'api' | 'plugins', focus = false) {
    for (const id of ['spreadsheet', 'api', 'plugins'] as const) {
      $(id === 'spreadsheet' ? 'spreadsheet-page' : `${id}-page`).hidden = id !== page;
      const button = $(id === 'spreadsheet' ? 'nav-sheet' : `nav-${id}`);
      button.classList.toggle('active', id === page);
      if (id === page) button.setAttribute('aria-current', 'page');
      else button.removeAttribute('aria-current');
    }
    const title =
      page === 'api' ? 'API quick start' : page === 'plugins' ? 'Plugin example' : 'Playground';
    $('page-name').textContent = title;
    document.title = `${page === 'spreadsheet' ? scenes[state.currentScene].title : title} · OpenSheet`;
    if (focus) {
      if (page === 'spreadsheet')
        document.querySelector<HTMLElement>('#spreadsheet .os-grid')?.focus();
      else $(`${page}-title`).focus();
    }
  }
  bindings.on($('nav-api'), 'click', () => {
    location.hash = 'api';
  });
  bindings.on($('nav-plugins'), 'click', () => {
    location.hash = 'plugins';
  });
  bindings.on($('nav-sheet'), 'click', () => {
    location.hash = state.currentScene;
  });
  for (const id of ['api-back', 'plugins-back'])
    bindings.on($(id), 'click', () => {
      location.hash = state.currentScene;
    });
  async function loadScene(id: SceneId) {
    $<HTMLSelectElement>('scene-select').disabled = true;
    cancelImport();
    if (state.currentScene === 'performance' && id !== 'performance') await leavePerformance();
    state.currentScene = id;
    if (id === 'performance') await startPerformance();
    else {
      const saved = localStorage.getItem(`opensheet.scene.${id}`);
      if (saved) await app.open(saved);
      else {
        const current = () => {
          try {
            return app.getWorkbook();
          } catch (error) {
            if ((error as { code?: string }).code === 'INVALID_ARGUMENT') return undefined;
            throw error;
          }
        };
        if (app.hasDraft) throw new Error('Finish or cancel editing before changing scenes');
        const previous = current();
        const revision = previous?.revision;
        let candidate: Workbook | undefined;
        let published = false;
        try {
          const snapshot = id === 'budget' ? await sample() : await createScene(id);
          candidate = new Workbook({ ...snapshot, workbookId: `wb-${crypto.randomUUID()}` });
          await candidate.ready();
          if (
            current() !== previous ||
            previous?.revision !== revision ||
            previous?.persistenceState.pending ||
            app.hasDraft
          )
            throw new Error('Workbook changed during scene creation; try again');
          localStorage.setItem(`opensheet.scene.${id}`, candidate.id);
          published = true;
          await app.attachWorkbook(candidate);
        } catch (error) {
          if (candidate && !published) {
            try {
              await candidate.deleteStorage();
            } catch {
              candidate.dispose();
            }
          } else if (candidate && published) {
            try {
              if (current() !== candidate) await candidate.close();
            } catch {
              candidate.dispose();
            }
            const message = error instanceof Error ? error.message : String(error);
            throw new Error(
              `${message}. Scene workbook ${candidate.id} was retained; reload to open the saved scene.`,
              { cause: error },
            );
          }
          throw error;
        }
      }
    }
    completeScene(id);
  }
  function completeScene(id: SceneId) {
    state.dirty = false;
    state.importedWorkbook = false;
    state.foundIndex = -1;
    $<HTMLInputElement>('read-only').checked = false;
    app.setMode('edit');
    $<HTMLInputElement>('document-name').value =
      id === 'budget' ? 'Launch budget' : scenes[id].title;
    refreshPersistence(app, state);
    $('scene-title').textContent = scenes[id].title;
    $('scene-label').textContent = scenes[id].label;
    $('scene-description').textContent = scenes[id].description;
    const page = location.hash.slice(1);
    document.title = `${page === 'api' ? 'API quick start' : page === 'plugins' ? 'Plugin example' : scenes[id].title} · OpenSheet`;
    $<HTMLSelectElement>('scene-select').value = id;
    showPanel(id === 'code' ? 'code' : 'insights');
    refresh();
    $<HTMLSelectElement>('scene-select').disabled = false;
  }
  bindings.on($<HTMLSelectElement>('scene-select'), 'change', (event) => {
    const next = (event.target as HTMLSelectElement).value as SceneId;
    if (next === state.currentScene) return;
    location.hash = next;
  });
  bindings.on($('reset-demo'), 'click', async () => {
    if (routeBusy) return;
    if (app.getWorkbook().persistenceState.pending) {
      notice('Wait for saving to finish');
      return;
    }
    if (state.dirty && !confirm('Reset this demo? Download JSON first to keep your changes.'))
      return;
    routeBusy = true;
    const scene = state.currentScene;
    const button = $<HTMLButtonElement>('reset-demo');
    button.disabled = true;
    let candidate: Workbook | undefined;
    let published = false;
    try {
      if (scene === 'performance') {
        localStorage.removeItem(`opensheet.scene.${scene}`);
        await loadScene(scene);
      } else {
        const previous = app.getWorkbook();
        const revision = previous.revision;
        const snapshot = scene === 'budget' ? await sample() : await createScene(scene);
        candidate = new Workbook({ ...snapshot, workbookId: `wb-${crypto.randomUUID()}` });
        await candidate.ready();
        if (
          app.getWorkbook() !== previous ||
          previous.revision !== revision ||
          previous.persistenceState.pending
        )
          throw new Error('Workbook changed during reset; try again');
        localStorage.setItem(`opensheet.scene.${scene}`, candidate.id);
        published = true;
        await app.attachWorkbook(candidate);
        completeScene(scene);
      }
    } catch (error) {
      if (candidate && !published) {
        try {
          await candidate.deleteStorage();
        } catch {
          candidate.dispose();
        }
      } else if (candidate && published) {
        try {
          if (app.getWorkbook() !== candidate) await candidate.close();
        } catch {
          candidate.dispose();
        }
      }
      $<HTMLSelectElement>('scene-select').disabled = false;
      const message = error instanceof Error ? error.message : String(error);
      notice(
        published
          ? `${message}. Reset workbook ${candidate?.id} was retained; reload to open the saved scene.`
          : message,
      );
    } finally {
      routeBusy = false;
      button.disabled = false;
    }
  });
  function requestedScene(): SceneId {
    const id = location.hash.slice(1);
    return Object.hasOwn(scenes, id) ? (id as SceneId) : 'budget';
  }
  let routeBusy = false;
  window.addEventListener(
    'hashchange',
    async () => {
      const page = location.hash.slice(1);
      if (page === 'api' || page === 'plugins') {
        showPage(page, true);
        return;
      }
      const next = requestedScene();
      if (routeBusy) {
        if (next !== state.currentScene) {
          history.replaceState(null, '', `#${state.currentScene}`);
          notice('Wait for the current scene operation to finish');
        }
        $<HTMLSelectElement>('scene-select').value = state.currentScene;
        showPage('spreadsheet', true);
        return;
      }
      if (next === state.currentScene) {
        $<HTMLSelectElement>('scene-select').value = state.currentScene;
        showPage('spreadsheet', true);
        return;
      }
      if (app.getWorkbook().persistenceState.pending) {
        history.replaceState(null, '', `#${state.currentScene}`);
        $<HTMLSelectElement>('scene-select').value = state.currentScene;
        notice('Wait for saving to finish');
        return;
      }
      if (state.dirty && !confirm('Switch demos? Download JSON first to keep your changes.')) {
        history.replaceState(null, '', `#${state.currentScene}`);
        $<HTMLSelectElement>('scene-select').value = state.currentScene;
        return;
      }
      routeBusy = true;
      const previous = state.currentScene;
      try {
        await loadScene(next);
        const page = location.hash.slice(1);
        if (page === 'api' || page === 'plugins') showPage(page);
        else showPage('spreadsheet', true);
      } catch (error) {
        state.currentScene = previous;
        const page = location.hash.slice(1);
        if (page !== 'api' && page !== 'plugins') history.replaceState(null, '', `#${previous}`);
        $<HTMLSelectElement>('scene-select').value = previous;
        $<HTMLSelectElement>('scene-select').disabled = false;
        notice(error instanceof Error ? error.message : String(error));
      } finally {
        routeBusy = false;
      }
    },
    { signal: bindings.signal },
  );
  return {
    restoreBackup: async (file: File, signal: AbortSignal) => {
      if (!/\.(opensheet|json)$/i.test(file.name))
        throw new Error('Choose a native or JSON backup');
      const scene = requestedScene();
      if (scene === 'performance') throw new Error('Open Budget Simulator to restore a backup');
      signal.throwIfAborted();
      const candidate = await createWorkbook();
      let attached = false;
      const key = `opensheet.scene.${scene}`;
      let referenceWritten = false;
      try {
        if (/\.json$/i.test(file.name)) await candidate.importJSON(file, { signal });
        else await candidate.importBinary(file, { signal });
        signal.throwIfAborted();
        localStorage.setItem(key, candidate.id);
        referenceWritten = true;
        signal.throwIfAborted();
        await app.attachWorkbook(candidate);
        attached = true;
        signal.throwIfAborted();
        state.currentScene = scene;
        completeScene(scene);
        showPage('spreadsheet');
        signal.throwIfAborted();
      } catch (error) {
        try {
          if (!attached && app.getWorkbook() === candidate) attached = true;
        } catch {
          /* No workbook is attached. */
        }
        if (attached || referenceWritten) {
          if (!attached) {
            try {
              await candidate.close();
            } catch {
              candidate.dispose();
            }
          }
          const message = error instanceof Error ? error.message : String(error);
          throw new RetainedBackupError(
            candidate.id,
            `Backup restored as workbook ${candidate.id}, but recovery could not finish: ${message}. The restored workbook was retained. Try again opens the currently saved scene, which may have changed in another tab.`,
            error,
          );
        }
        try {
          await candidate.deleteStorage();
        } catch {
          /* Preserve the original error; storage may be unavailable. */
        }
        candidate.dispose();
        throw error;
      }
    },
    openRetainedBackup: async (id: string) => {
      const scene = requestedScene();
      if (scene === 'performance') throw new Error('Open Budget Simulator to restore a backup');
      await app.open(id);
      localStorage.setItem(`opensheet.scene.${scene}`, id);
      state.currentScene = scene;
      completeScene(scene);
      showPage('spreadsheet');
    },
    start: async () => {
      routeBusy = true;
      try {
        await loadScene(requestedScene());
        const page = location.hash.slice(1);
        showPage(page === 'api' || page === 'plugins' ? page : 'spreadsheet');
      } finally {
        routeBusy = false;
      }
    },
    dispose: bindings.dispose,
  };
}
