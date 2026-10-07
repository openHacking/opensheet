import type { Rect } from './types.js';
type Node = {
  key: string;
  range: Rect;
  value: string;
  priority: number;
  max: number;
  left?: Node;
  right?: Node;
};
const update = (node: Node) => {
  node.max = Math.max(node.range.endRow, node.left?.max ?? 0, node.right?.max ?? 0);
  return node;
};
const merge = (a: Node | undefined, b: Node | undefined): Node | undefined => {
  if (!a) return b;
  if (!b) return a;
  if (a.priority < b.priority) {
    a.right = merge(a.right, b);
    return update(a);
  }
  b.left = merge(a, b.left);
  return update(b);
};
const less = (range: Rect, key: string, node: Node) =>
  range.startRow < node.range.startRow ||
  (range.startRow === node.range.startRow && key < node.key);
const insert = (root: Node | undefined, node: Node): Node => {
  if (!root) return node;
  if (less(node.range, node.key, root)) {
    root.left = insert(root.left, node);
    if (root.left.priority < root.priority) {
      const next = root.left;
      root.left = next.right;
      next.right = update(root);
      return update(next);
    }
  } else {
    root.right = insert(root.right, node);
    if (root.right.priority < root.priority) {
      const next = root.right;
      root.right = next.left;
      next.left = update(root);
      return update(next);
    }
  }
  return update(root);
};
const remove = (root: Node | undefined, range: Rect, key: string): Node | undefined => {
  if (!root) return;
  if (root.key === key) return merge(root.left, root.right);
  if (less(range, key, root)) root.left = remove(root.left, range, key);
  else root.right = remove(root.right, range, key);
  return update(root);
};
/** Augmented interval treap: range dependencies do not expand to individual cells. */
export class ReferenceIndex {
  private roots = new Map<string, Node | undefined>();
  private entries = new Map<string, Array<{ sheet: string; range: Rect; key: string }>>();
  private sequence = 0;
  add(value: string, references: Array<{ sheet: string; range: Rect }>) {
    this.delete(value);
    const entries = references.map((ref) => {
      const key = `${value}/${this.sequence++}`;
      let hash = 2166136261;
      for (const c of key) hash = Math.imul(hash ^ c.charCodeAt(0), 16777619);
      this.roots.set(
        ref.sheet,
        insert(this.roots.get(ref.sheet), {
          ...ref,
          key,
          value,
          priority: hash >>> 0,
          max: ref.range.endRow,
        }),
      );
      return { ...ref, key };
    });
    this.entries.set(value, entries);
  }
  delete(value: string) {
    for (const entry of this.entries.get(value) ?? [])
      this.roots.set(entry.sheet, remove(this.roots.get(entry.sheet), entry.range, entry.key));
    this.entries.delete(value);
  }
  query(sheet: string, range: Rect) {
    const values = new Set<string>();
    const visit = (node: Node | undefined) => {
      if (!node || node.max <= range.startRow) return;
      visit(node.left);
      if (node.range.startRow >= range.endRow) return;
      if (
        node.range.endRow > range.startRow &&
        node.range.startColumn < range.endColumn &&
        node.range.endColumn > range.startColumn
      )
        values.add(node.value);
      visit(node.right);
    };
    visit(this.roots.get(sheet));
    return values;
  }
  clear() {
    this.roots.clear();
    this.entries.clear();
  }
}
