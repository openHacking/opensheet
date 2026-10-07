import { assert } from './types.js';
type Node = {
  id: number;
  values: Uint32Array;
  left?: Node;
  right?: Node;
  parent?: Node;
  count: number;
  priority: number;
};
const size = (node: Node | undefined) => node?.count ?? 0;
const update = (node: Node) => {
  node.count = size(node.left) + node.values.length + size(node.right);
  if (node.left) node.left.parent = node;
  if (node.right) node.right.parent = node;
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
/** Chunked implicit treap with sparse typed reverse directories; no string identity maps. */
export class AxisIndex {
  private root?: Node;
  private nodes = new Map<number, Node>();
  private directory = new Map<number, Uint32Array>();
  private nextLeaf = 1;
  private owned = new Set<number>();
  constructor(ids: readonly number[] = []) {
    for (let i = 0; i < ids.length; i += 128)
      this.root = merge(this.root, this.leaf(ids.slice(i, i + 128)));
    if (this.root) this.root.parent = undefined;
  }
  private location(id: number, value: number) {
    const page = Math.floor(id / 64);
    let entries = this.directory.get(page);
    if (!entries || !this.owned.has(page)) {
      entries = entries ? entries.slice() : new Uint32Array(64);
      this.directory.set(page, entries);
      this.owned.add(page);
    }
    entries[id % 64] = value;
  }
  private leaf(values: readonly number[]): Node {
    const id = this.nextLeaf++;
    assert(id < 0x1ffffff, 'LIMIT_EXCEEDED', 'Axis index exhausted');
    let priority = Math.imul(id, 2654435761) >>> 0;
    priority ^= priority >>> 16;
    const node: Node = {
      id,
      values: new Uint32Array(values),
      count: values.length,
      priority: priority >>> 0,
    };
    this.nodes.set(id, node);
    values.forEach((value, i) => this.location(value, id * 128 + i + 1));
    return node;
  }
  private split(node: Node | undefined, at: number): [Node | undefined, Node | undefined] {
    if (!node) return [undefined, undefined];
    const before = size(node.left),
      end = before + node.values.length;
    if (at < before) {
      const [a, b] = this.split(node.left, at);
      node.left = b;
      return [a, update(node)];
    }
    if (at > end) {
      const [a, b] = this.split(node.right, at - end);
      node.right = a;
      return [update(node), b];
    }
    const left =
      at === before
        ? node.left
        : merge(node.left, this.leaf([...node.values.subarray(0, at - before)]));
    const right =
      at === end
        ? node.right
        : merge(this.leaf([...node.values.subarray(at - before)]), node.right);
    this.nodes.delete(node.id);
    return [left, right];
  }
  get length() {
    return size(this.root);
  }
  get(id: number) {
    const entry = this.directory.get(Math.floor(id / 64))?.[id % 64];
    if (!entry) return;
    const encoded = entry - 1,
      node = this.nodes.get(Math.floor(encoded / 128));
    if (!node) return;
    let rank = size(node.left) + (encoded % 128),
      current = node;
    while (current.parent) {
      if (current.parent.right === current)
        rank += size(current.parent.left) + current.parent.values.length;
      current = current.parent;
    }
    return rank;
  }
  at(index: number) {
    assert(index >= 0 && index < this.length, 'INVALID_RANGE', 'Unknown axis position');
    let node = this.root!,
      at = index;
    while (true) {
      const before = size(node.left);
      if (at < before) node = node.left!;
      else if (at >= before + node.values.length) {
        at -= before + node.values.length;
        node = node.right!;
      } else return node.values[at - before];
    }
  }
  toArray() {
    const ids: number[] = [];
    const walk = (node: Node | undefined) => {
      if (!node) return;
      walk(node.left);
      ids.push(...node.values);
      walk(node.right);
    };
    walk(this.root);
    return ids;
  }
  clone() {
    const index = new AxisIndex();
    index.directory = new Map(this.directory);
    this.owned.clear();
    index.owned.clear();
    index.nextLeaf = this.nextLeaf;
    const copy = (node: Node | undefined): Node | undefined => {
      if (!node) return;
      const cloned = { ...node, left: copy(node.left), right: copy(node.right), parent: undefined };
      index.nodes.set(cloned.id, cloned);
      return update(cloned);
    };
    index.root = copy(this.root);
    return index;
  }
  splice(start: number, remove: number, ids: number[] = []) {
    assert(
      start >= 0 && remove >= 0 && start + remove <= this.length,
      'INVALID_RANGE',
      'Invalid axis operation',
    );
    const [left, rest] = this.split(this.root, start),
      [deleted, right] = this.split(rest, remove);
    const clear = (node: Node | undefined) => {
      if (!node) return;
      clear(node.left);
      clear(node.right);
      for (const id of node.values) this.location(id, 0);
      this.nodes.delete(node.id);
    };
    clear(deleted);
    let inserted: Node | undefined;
    for (let i = 0; i < ids.length; i += 128)
      inserted = merge(inserted, this.leaf(ids.slice(i, i + 128)));
    this.root = merge(merge(left, inserted), right);
    if (this.root) this.root.parent = undefined;
  }
  get bytes() {
    return this.directory.size * 256 + this.nodes.size * 128 + this.length * 4;
  }
}
