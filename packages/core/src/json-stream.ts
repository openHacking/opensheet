import { assert, LIMITS } from './types.js';
import { yieldToTask } from './task-yield.js';
import type { FileBlock, WorkbookFile } from './file-codec.js';
export type JSONSource = string | Blob | ReadableStream<Uint8Array>;
type Token = { text: string; punctuation: boolean };
async function* tokens(
  source: JSONSource,
  checkpoint?: () => Promise<void>,
  signal?: AbortSignal,
): AsyncGenerator<Token[]> {
  signal?.throwIfAborted();
  const stream =
    typeof source === 'string'
      ? new Blob([source]).stream()
      : source instanceof Blob
        ? source.stream()
        : source;
  const reader = stream.getReader(),
    decoder = new TextDecoder('utf-8', { fatal: true });
  let text = '',
    quoted = false,
    escaped = false;
  let input: Uint8Array | undefined,
    inputOffset = 0;
  let pending: Token[] = [],
    pendingChars = 0,
    chunksSinceCheck = 0,
    charsSinceCheck = 0,
    lastYield = performance.now();
  const yieldIfNeeded = async () => {
    if (performance.now() - lastYield < 8) return;
    await yieldToTask();
    lastYield = performance.now();
    signal?.throwIfAborted();
  };
  const emit = (token: Token) => {
    pending.push(token);
    pendingChars += token.text.length;
  };
  const abort = () => {
    void reader.cancel().catch(() => {});
  };
  signal?.addEventListener('abort', abort, { once: true });
  try {
    while (true) {
      signal?.throwIfAborted();
      let done = false;
      if (!input || inputOffset === input.length) {
        const item = await reader.read();
        signal?.throwIfAborted();
        done = item.done;
        assert(
          done ||
            (ArrayBuffer.isView(item.value) &&
              Object.prototype.toString.call(item.value) === '[object Uint8Array]'),
          'INVALID_ARGUMENT',
          'JSON stream must contain Uint8Array chunks',
        );
        input = item.value;
        inputOffset = 0;
      }
      let chunk: string;
      try {
        if (done) chunk = decoder.decode();
        else {
          const end = Math.min(inputOffset + 65536, input!.length);
          chunk = decoder.decode(input!.subarray(inputOffset, end), { stream: true });
          inputOffset = end;
        }
      } catch {
        assert(false, 'INVALID_ARGUMENT', 'Invalid UTF-8 in JSON stream');
      }
      for (const char of chunk) {
        if (quoted) {
          text += char;
          if (escaped) escaped = false;
          else if (char === '\\') escaped = true;
          else if (char === '"') {
            emit({ text, punctuation: false });
            text = '';
            quoted = false;
          }
        } else if (char === '"') {
          assert(text === '', 'INVALID_ARGUMENT', 'Invalid JSON token');
          text = char;
          quoted = true;
        } else if ('{}[]:,'.includes(char) || /[ \t\r\n]/.test(char)) {
          if (text) {
            emit({ text, punctuation: false });
            text = '';
          }
          if (!/[ \t\r\n]/.test(char)) emit({ text: char, punctuation: true });
        } else text += char;
        assert(
          text.length <= LIMITS.text * 6 + 2,
          'LIMIT_EXCEEDED',
          'JSON token exceeds text limit',
        );
        if (pending.length >= 512 || pendingChars >= 65536) {
          yield pending;
          pending = [];
          pendingChars = 0;
        }
        if (++charsSinceCheck >= 8192) {
          charsSinceCheck = 0;
          await yieldIfNeeded();
        }
      }
      if (!done && pending.length) {
        yield pending;
        pending = [];
        pendingChars = 0;
      }
      if (done) break;
      if (checkpoint) await checkpoint();
      // Even streams that continuously produce tokens must allow cancellation tasks to run.
      if (++chunksSinceCheck >= 64) {
        chunksSinceCheck = 0;
        await yieldIfNeeded();
      }
    }
    assert(!quoted, 'INVALID_ARGUMENT', 'Unterminated JSON string');
    if (text) emit({ text, punctuation: false });
    if (pending.length) yield pending;
  } finally {
    signal?.removeEventListener('abort', abort);
    // The source can leave cancellation pending indefinitely. Release our lock
    // after requesting cancellation without delaying the parser result.
    void reader.cancel().catch(() => {});
    reader.releaseLock();
  }
}
type Frame = {
  value: any;
  path: Array<string | number>;
  array: boolean;
  state: 'key' | 'colon' | 'value' | 'comma';
  key?: string;
  count: number;
  keys: Set<string>;
  trailing: boolean;
};
/** Emits complete blocks without retaining them in the root JSON object. */
export async function* parseWorkbookJSON(
  source: JSONSource,
  options: { checkpoint?: () => Promise<void>; signal?: AbortSignal } = {},
): AsyncGenerator<{ sheet: number; block: FileBlock } | { metadata: WorkbookFile }> {
  const stack: Frame[] = [];
  let root: unknown,
    finished = false,
    metadataBytes = 0,
    blockBytes = 0;
  const path = () => {
    const f = stack.at(-1);
    return f ? [...f.path, f.array ? f.count : f.key!] : [];
  };
  const attach = (value: unknown, discard = false) => {
    const f = stack.at(-1);
    if (!f) {
      assert(root === undefined && !finished, 'INVALID_ARGUMENT', 'Multiple JSON roots');
      root = value;
      finished = true;
      return;
    }
    assert(f.state === 'value', 'INVALID_ARGUMENT', 'Unexpected JSON value');
    if (f.array) {
      if (!discard) f.value.push(value);
      f.count++;
    } else f.value[f.key!] = value;
    f.state = 'comma';
    f.trailing = false;
  };
  for await (const batch of tokens(source, options.checkpoint, options.signal)) {
    for (const token of batch) {
      const f = stack.at(-1),
        t = token.text;
      if (f?.path[0] === 'sheets' && f.path[2] === 'blocks' && f.path.length >= 4) {
        blockBytes += t.length * 2;
        assert(
          blockBytes <= 32 * 1024 * 1024,
          'LIMIT_EXCEEDED',
          'JSON block exceeds 32 MiB working budget',
        );
      }
      if (!f || !(f.path[0] === 'sheets' && f.path[2] === 'blocks' && f.path.length >= 4)) {
        metadataBytes += t.length * 2;
        assert(metadataBytes <= 32 * 1024 * 1024, 'LIMIT_EXCEEDED', 'JSON metadata exceeds 32 MiB');
      }
      if (token.punctuation) {
        if (t === '{' || t === '[') {
          assert(
            !finished && (!f || f.state === 'value'),
            'INVALID_ARGUMENT',
            'Unexpected JSON container',
          );
          assert(stack.length < 32, 'LIMIT_EXCEEDED', 'JSON nesting limit');
          stack.push({
            value: t === '[' ? [] : Object.create(null),
            path: path(),
            array: t === '[',
            state: t === '[' ? 'value' : 'key',
            count: 0,
            keys: new Set(),
            trailing: false,
          });
        } else if (t === '}' || t === ']') {
          assert(
            f &&
              f.array === (t === ']') &&
              !f.trailing &&
              (f.state === 'comma' ||
                (f.array
                  ? f.state === 'value' && f.count === 0
                  : f.state === 'key' && f.keys.size === 0)),
            'INVALID_ARGUMENT',
            'Invalid JSON container ending',
          );
          stack.pop();
          const isBlock =
            f.path.length === 4 &&
            f.path[0] === 'sheets' &&
            typeof f.path[1] === 'number' &&
            f.path[2] === 'blocks';
          attach(f.value, isBlock);
          if (isBlock) {
            blockBytes = 0;
            yield { sheet: f.path[1] as number, block: f.value };
          }
        } else if (t === ',') {
          assert(f?.state === 'comma', 'INVALID_ARGUMENT', 'Unexpected JSON comma');
          f.state = f.array ? 'value' : 'key';
          f.trailing = true;
        } else {
          assert(
            t === ':' && f && !f.array && f.state === 'colon',
            'INVALID_ARGUMENT',
            'Unexpected JSON colon',
          );
          f.state = 'value';
        }
      } else {
        let value: unknown;
        try {
          value = JSON.parse(t);
        } catch {
          assert(false, 'INVALID_ARGUMENT', 'Invalid JSON token');
        }
        if (f && !f.array && f.state === 'key') {
          assert(
            typeof value === 'string' &&
              !f.keys.has(value) &&
              !['__proto__', 'prototype', 'constructor'].includes(value),
            'INVALID_ARGUMENT',
            'Duplicate or unsafe JSON key',
          );
          f.keys.add(value);
          f.key = value;
          f.state = 'colon';
        } else attach(value);
      }
    }
  }
  assert(
    finished && stack.length === 0 && root !== null && typeof root === 'object',
    'INVALID_ARGUMENT',
    'Incomplete workbook JSON',
  );
  yield { metadata: root as WorkbookFile };
}
