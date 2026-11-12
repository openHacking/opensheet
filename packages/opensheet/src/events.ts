import type { EventMap } from './types.js';
export class EditorEvents {
  private listeners: { [K in keyof EventMap]?: Set<(value: EventMap[K]) => void> } = {};
  constructor(private onError: (error: unknown) => void) {}
  on<K extends keyof EventMap>(event: K, listener: (value: EventMap[K]) => void) {
    const set = (this.listeners[event] ??= new Set<
      (value: EventMap[K]) => void
    >() as (typeof this.listeners)[K] & Set<(value: EventMap[K]) => void>);
    set.add(listener);
    return () => {
      set.delete(listener);
    };
  }
  emit<K extends keyof EventMap>(event: K, value: EventMap[K]) {
    const report = (error: unknown) => {
      try {
        const result: unknown = this.onError(error);
        if (result && typeof (result as PromiseLike<unknown>).then === 'function')
          void Promise.resolve(result).catch((observerError) => console.error(observerError));
      } catch (observerError) {
        console.error(observerError);
      }
    };
    for (const fn of this.listeners[event] ?? []) {
      try {
        const result: unknown = fn(value);
        if (result && typeof (result as PromiseLike<unknown>).then === 'function')
          void Promise.resolve(result).catch(report);
      } catch (error) {
        report(error);
      }
    }
  }
  clear() {
    this.listeners = {};
  }
}
