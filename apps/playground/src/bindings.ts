export function createBindings() {
  const abort = new AbortController();
  return {
    signal: abort.signal,
    on<K extends keyof HTMLElementEventMap>(
      element: HTMLElement,
      event: K,
      handler: (event: HTMLElementEventMap[K]) => void,
    ) {
      element.addEventListener(event, handler, { signal: abort.signal });
    },
    dispose() {
      abort.abort();
    },
  };
}
