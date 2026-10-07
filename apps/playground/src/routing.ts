import type { OpenSheet } from 'opensheet';
import { createBindings } from './bindings.js';
import { sample } from './budget.js';
import { $ } from './dom.js';
import { createScene, scenes, type SceneId } from './scenes.js';
import type { PlaygroundState } from './state.js';
export function createRouting(
  app: OpenSheet,
  state: PlaygroundState,
  cancelImport: () => void,
  refresh: () => void,
  showPanel: (panel: 'code' | 'insights') => void,
  startPerformance: () => Promise<void>,
  leavePerformance: () => Promise<void>,
) {
  const bindings = createBindings();
  async function loadScene(id: SceneId) {
    $<HTMLSelectElement>('scene-select').disabled = true;
    cancelImport();
    if (state.currentScene === 'performance' && id !== 'performance') await leavePerformance();
    state.currentScene = id;
    if (id === 'performance') await startPerformance();
    else {
      const saved = localStorage.getItem(`opensheet.scene.${id}`);
      if (saved) {
        try {
          await app.open(saved);
        } catch {
          await app.load(id === 'budget' ? await sample() : await createScene(id));
        }
      } else await app.load(id === 'budget' ? await sample() : await createScene(id));
      localStorage.setItem(`opensheet.scene.${id}`, app.getWorkbook().id);
    }
    state.dirty = false;
    state.importedWorkbook = false;
    state.foundIndex = -1;
    $<HTMLInputElement>('read-only').checked = false;
    app.setMode('edit');
    $<HTMLInputElement>('document-name').value =
      id === 'budget' ? 'Launch budget' : scenes[id].title;
    $('save-state').textContent = 'Example workbook · All changes stay local';
    $('scene-title').textContent = scenes[id].title;
    $('scene-label').textContent = scenes[id].label;
    $('scene-description').textContent = scenes[id].description;
    document.title = `${scenes[id].title} · OpenSheet`;
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
    if (state.dirty && !confirm('Reset this demo? Download JSON first to keep your changes.'))
      return;
    localStorage.removeItem(`opensheet.scene.${state.currentScene}`);
    if (state.currentScene !== 'performance') await app.getWorkbook().deleteStorage();
    await loadScene(state.currentScene);
  });
  function requestedScene(): SceneId {
    const id = location.hash.slice(1);
    return Object.hasOwn(scenes, id) ? (id as SceneId) : 'budget';
  }
  let routeBusy = false;
  window.addEventListener(
    'hashchange',
    async () => {
      const next = requestedScene();
      if (routeBusy || next === state.currentScene) {
        $<HTMLSelectElement>('scene-select').value = state.currentScene;
        return;
      }
      if (state.dirty && !confirm('Switch demos? Download JSON first to keep your changes.')) {
        history.replaceState(null, '', `#${state.currentScene}`);
        $<HTMLSelectElement>('scene-select').value = state.currentScene;
        return;
      }
      routeBusy = true;
      try {
        await loadScene(next);
      } finally {
        routeBusy = false;
      }
    },
    { signal: bindings.signal },
  );
  return {
    start: async () => {
      routeBusy = true;
      try {
        await loadScene(requestedScene());
      } finally {
        routeBusy = false;
      }
    },
    dispose: bindings.dispose,
  };
}
