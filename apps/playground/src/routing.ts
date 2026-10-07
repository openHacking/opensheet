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
) {
  const bindings = createBindings();
  async function loadScene(id: SceneId) {
    cancelImport();
    await app.load(id === 'budget' ? sample() : createScene(id));
    state.currentScene = id;
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
  }
  bindings.on($<HTMLSelectElement>('scene-select'), 'change', (event) => {
    const next = (event.target as HTMLSelectElement).value as SceneId;
    if (next === state.currentScene) return;
    location.hash = next;
  });
  bindings.on($('reset-demo'), 'click', async () => {
    if (state.dirty && !confirm('Reset this demo? Download JSON first to keep your changes.'))
      return;
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
  return { start: () => loadScene(requestedScene()), dispose: bindings.dispose };
}
