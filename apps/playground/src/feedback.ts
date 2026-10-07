import { createBindings } from './bindings.js';
import { $ } from './dom.js';
export function createFeedback() {
  const bindings = createBindings();
  let timer: ReturnType<typeof setTimeout> | undefined;
  function notice(message: string) {
    $('notice').textContent = message;
    $('notice').hidden = false;
    clearTimeout(timer);
    timer = setTimeout(() => ($('notice').hidden = true), 5000);
  }
  function modal(title: string, content: HTMLElement) {
    $('dialog-title').textContent = title;
    $('dialog-content').replaceChildren(content);
    $('dialog').scrollTop = 0;
    ($('dialog') as HTMLDialogElement).showModal();
  }
  bindings.on($('dialog-close'), 'click', () => ($('dialog') as HTMLDialogElement).close());
  function textModal(title: string, text: string) {
    const pre = document.createElement('pre');
    pre.textContent = text;
    modal(title, pre);
  }
  return {
    notice,
    modal,
    textModal,
    dispose() {
      clearTimeout(timer);
      bindings.dispose();
    },
  };
}
export type Feedback = ReturnType<typeof createFeedback>;
