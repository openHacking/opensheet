import { createBindings } from './bindings.js';
import { $ } from './dom.js';
export function createWorkbench() {
  const bindings = createBindings();
  type Panel = 'insights' | 'code';
  function showPanel(panel: Panel) {
    for (const name of ['insights', 'code'] as const) {
      const active = name === panel;
      $(`panel-${name}`).hidden = !active;
      $(`tab-${name}`).setAttribute('aria-selected', String(active));
      $<HTMLButtonElement>(`tab-${name}`).tabIndex = active ? 0 : -1;
    }
  }
  for (const name of ['insights', 'code'] as const) {
    bindings.on($(`tab-${name}`), 'click', () => showPanel(name));
    bindings.on($(`tab-${name}`), 'keydown', (event) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      const next = name === 'insights' ? 'code' : 'insights';
      showPanel(next);
      $(`tab-${next}`).focus();
    });
  }
  const workbench = $('workbench');
  const splitter = $('splitter');
  const defaultSplitFraction = 0.67;
  let splitFraction = defaultSplitFraction;
  function resizeWorkbench(fraction = splitFraction) {
    if (getComputedStyle(workbench).display !== 'grid') return;
    const available = workbench.clientWidth - splitter.offsetWidth;
    if (available <= 0) return;
    const min = Math.min(520, Math.max(0, available - 320));
    const max = Math.max(min, available - 320);
    const left = Math.max(min, Math.min(max, available * fraction));
    splitFraction = left / available;
    workbench.style.setProperty('--sheet-width', `${left}px`);
    splitter.setAttribute('aria-valuemin', String(Math.round((min / available) * 100)));
    splitter.setAttribute('aria-valuemax', String(Math.round((max / available) * 100)));
    splitter.setAttribute('aria-valuenow', String(Math.round(splitFraction * 100)));
  }
  const resize = new ResizeObserver(() => resizeWorkbench());
  resize.observe(workbench);
  bindings.on(splitter, 'pointerdown', (event) => {
    if (event.button !== 0) return;
    splitter.setPointerCapture(event.pointerId);
    splitter.classList.add('dragging');
    event.preventDefault();
  });
  bindings.on(splitter, 'pointermove', (event) => {
    if (!splitter.hasPointerCapture(event.pointerId)) return;
    const available = workbench.clientWidth - splitter.offsetWidth;
    resizeWorkbench((event.clientX - workbench.getBoundingClientRect().left) / available);
  });
  bindings.on(splitter, 'pointerup', (event) => {
    if (splitter.hasPointerCapture(event.pointerId))
      splitter.releasePointerCapture(event.pointerId);
    splitter.classList.remove('dragging');
  });
  bindings.on(splitter, 'pointercancel', () => splitter.classList.remove('dragging'));
  bindings.on(splitter, 'dblclick', () => resizeWorkbench(defaultSplitFraction));
  bindings.on(splitter, 'keydown', (event) => {
    const available = workbench.clientWidth - splitter.offsetWidth;
    if (event.key === 'ArrowLeft') resizeWorkbench(splitFraction - 24 / available);
    else if (event.key === 'ArrowRight') resizeWorkbench(splitFraction + 24 / available);
    else if (event.key === 'Home') resizeWorkbench(0);
    else if (event.key === 'End') resizeWorkbench(1);
    else return;
    event.preventDefault();
  });
  return {
    showPanel,
    dispose() {
      resize.disconnect();
      bindings.dispose();
    },
  };
}
