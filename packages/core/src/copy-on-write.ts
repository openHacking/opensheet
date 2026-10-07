/** A bounded transaction overlay. Iteration is reserved for explicit whole-book operations. */
export function overlay<T>(base: Record<string, T>) {
  const changed = new Map<string, T | undefined>();
  const record: Record<string, T> = new Proxy(Object.create(null) as Record<string, T>, {
    get: (_target, key) =>
      typeof key === 'string' ? (changed.has(key) ? changed.get(key) : base[key]) : undefined,
    set: (_target, key, value) => {
      changed.set(String(key), value);
      return true;
    },
    deleteProperty: (_target, key) => {
      changed.set(String(key), undefined);
      return true;
    },
    ownKeys: (): string[] =>
      [...new Set([...Object.keys(base), ...changed.keys()])].filter(
        (key) => record[key] !== undefined,
      ),
    getOwnPropertyDescriptor: (_target, key) =>
      record[String(key)] === undefined
        ? undefined
        : { configurable: true, enumerable: true, writable: true, value: record[String(key)] },
    has: (_target, key) => record[String(key)] !== undefined,
  });
  return { record, changed };
}
