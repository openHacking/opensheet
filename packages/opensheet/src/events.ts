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
    for (const fn of this.listeners[event] ?? []) {
      try {
        fn(value);
      } catch (error) {
        this.onError(error);
      }
    }
  }
  clear() {
    this.listeners = {};
  }
}
