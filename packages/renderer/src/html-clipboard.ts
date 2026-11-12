import {
  LIMITS,
  OpenSheetError,
  parseInput,
  inputSchema,
  scalarSchema,
  styleSchema,
  cellSchema,
  type CellRecord,
  type Scalar,
  type CellInput,
  type CellStyle,
} from '@opensheetjs/core';

export type HtmlClipboardCell = {
  input: CellInput;
  style: CellStyle;
  externalText?: string;
  value?: Scalar;
  source?: { row: number; column: number };
  note?: string | null;
  link?: CellRecord['link'] | null;
};

function hasOnlyKeys(value: unknown, keys: readonly string[]): boolean {
  return (
    !!value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.keys(value).every((key) => keys.includes(key))
  );
}

function hasOnlyScalarFields(value: unknown, formula: boolean): boolean {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const type = (value as { type?: unknown }).type;
  const field =
    type === 'blank'
      ? undefined
      : type === 'formula' && formula
        ? 'expression'
        : type === 'error'
          ? 'code'
          : ['number', 'string', 'boolean'].includes(String(type))
            ? 'value'
            : undefined;
  return hasOnlyKeys(value, field ? ['type', field] : ['type']);
}

function color(value: string): string | undefined {
  if (/^#[0-9a-f]{6}$/i.test(value)) return value.toLowerCase();
  if (/^#[0-9a-f]{3}$/i.test(value))
    return (
      '#' +
      [...value.slice(1)]
        .map((c) => c + c)
        .join('')
        .toLowerCase()
    );
  const rgb = /^rgba?\(\s*(\d+),\s*(\d+),\s*(\d+)(?:,\s*(1|1\.0+))?\s*\)$/.exec(value);
  if (rgb && rgb.slice(1, 4).every((v) => Number(v) <= 255))
    return (
      '#' +
      rgb
        .slice(1, 4)
        .map((v) => Number(v).toString(16).padStart(2, '0'))
        .join('')
    );
  return (
    {
      black: '#000000',
      white: '#ffffff',
      red: '#ff0000',
      blue: '#0000ff',
      green: '#008000',
      yellow: '#ffff00',
      gray: '#808080',
    } as Record<string, string>
  )[value.toLowerCase()];
}

function textOf(cell: HTMLElement): string {
  let text = '',
    generatedBreak = false;
  const tag = (node: Node) => (node instanceof Element ? node.tagName.toUpperCase() : '');
  const block = (node: Node) => /^(P|DIV|LI|PRE|BLOCKQUOTE|H[1-6])$/.test(tag(node));
  function append(value: string, generated = false) {
    if (!value) return;
    text += value;
    generatedBreak = generated;
    if (text.length - (generatedBreak ? 1 : 0) > LIMITS.text)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Pasted cell text is too long');
  }
  function boundary() {
    if (text && !text.endsWith('\n')) append('\n', true);
  }
  // Iterative depth-first traversal avoids recursion and a sibling-sized work queue.
  let node: Node | null = cell.firstChild;
  while (node) {
    const excluded =
      /^(SCRIPT|STYLE|IFRAME|OBJECT|EMBED|TEMPLATE|NOSCRIPT|SVG|MATH|CANVAS|AUDIO|VIDEO)$/.test(
        tag(node),
      );
    if (!excluded) {
      if (node.nodeType === Node.TEXT_NODE) append(node.textContent ?? '');
      else if (tag(node) === 'BR') append('\n');
      else if (block(node)) boundary();
      if (node.firstChild) {
        node = node.firstChild;
        continue;
      }
    }
    while (node && node !== cell) {
      if (block(node)) boundary();
      if (node.nextSibling) {
        node = node.nextSibling;
        break;
      }
      node = node.parentNode;
    }
    if (node === cell) break;
  }
  return generatedBreak ? text.slice(0, -1) : text;
}

function styleOf(cell: HTMLElement): CellStyle {
  const css = cell.style,
    style: CellStyle = {};
  if (css.fontWeight === 'bold' || Number(css.fontWeight) >= 600) style.bold = true;
  else if (css.fontWeight === 'normal' || (css.fontWeight && Number(css.fontWeight) < 600))
    style.bold = false;
  if (css.fontStyle === 'italic' || css.fontStyle === 'oblique') style.italic = true;
  else if (css.fontStyle === 'normal') style.italic = false;
  if (css.textDecoration.includes('underline')) style.underline = true;
  if (['left', 'center', 'right'].includes(css.textAlign))
    style.align = css.textAlign as CellStyle['align'];
  const foreground = color(css.color),
    background = color(css.backgroundColor);
  if (foreground) style.color = foreground;
  if (background) style.background = background;
  const fontSize = /^(\d+(?:\.\d+)?)(px|pt)$/.exec(css.fontSize);
  if (fontSize) {
    const size = Number(fontSize[1]) * (fontSize[2] === 'pt' ? 4 / 3 : 1);
    if (size >= 8 && size <= 72) style.fontSize = size;
  }
  if (['normal', 'pre-wrap', 'pre-line'].includes(css.whiteSpace)) style.wrap = true;
  else if (css.whiteSpace === 'nowrap') style.wrap = false;
  return style;
}

/** Parse in an inert template; never attach clipboard DOM to the live document. */
export function parseClipboardHTML(html: string): HtmlClipboardCell[][] | null {
  if (html.length > 2 * 1024 * 1024)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste is too large');
  const template = document.createElement('template');
  template.innerHTML = html;
  const table = template.content.querySelector('table');
  if (!table) return null;
  const version = table.getAttribute('data-opensheet-clipboard');
  let source: { row: number; column: number } | undefined;
  if (version !== null) {
    if (version !== '1')
      throw new OpenSheetError('UNSUPPORTED_FEATURE', 'Unknown structured clipboard version');
    const row = table.getAttribute('data-row') ?? '',
      column = table.getAttribute('data-column') ?? '';
    if (
      !/^\d+$/.test(row) ||
      !/^\d+$/.test(column) ||
      Number(row) >= LIMITS.rows ||
      Number(column) >= LIMITS.columns
    )
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid structured clipboard origin');
    source = { row: Number(row), column: Number(column) };
  }
  if (template.content.querySelectorAll('table').length !== 1)
    throw new OpenSheetError(
      'UNSUPPORTED_FEATURE',
      'Paste one table at a time; nested tables are not supported',
    );
  if (table.querySelector('table'))
    throw new OpenSheetError('UNSUPPORTED_FEATURE', 'Nested clipboard tables are not supported');
  const rows: HtmlClipboardCell[][] = [];
  let width = 0;
  for (const tr of table.rows) {
    const row: HtmlClipboardCell[] = [];
    for (const cell of tr.cells) {
      if (cell.rowSpan !== 1 || cell.colSpan !== 1)
        throw new OpenSheetError('UNSUPPORTED_FEATURE', 'Unmerge the source table before pasting');
      if (source) {
        let payload: any;
        try {
          payload = JSON.parse(cell.getAttribute('data-opensheet-cell') ?? '');
        } catch {
          throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid structured clipboard cell');
        }
        const input = inputSchema.safeParse(payload?.input),
          value = scalarSchema.safeParse(payload?.value),
          style = styleSchema.strict().safeParse(payload?.style),
          note = cellSchema.shape.note.nullable().safeParse(payload?.note),
          link = cellSchema.shape.link.nullable().safeParse(payload?.link);
        if (
          !hasOnlyKeys(payload, ['input', 'value', 'style', 'note', 'link']) ||
          !hasOnlyScalarFields(payload.input, true) ||
          !hasOnlyScalarFields(payload.value, false) ||
          !input.success ||
          !value.success ||
          !style.success ||
          !note.success ||
          !link.success ||
          source.row + rows.length >= LIMITS.rows ||
          source.column + row.length >= LIMITS.columns
        )
          throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid structured clipboard cell');
        row.push({
          input: input.data,
          value: value.data,
          style: style.data,
          note: note.data,
          link: link.data,
          source: { row: source.row + rows.length, column: source.column + row.length },
        });
      } else {
        const text = textOf(cell);
        row.push({ input: parseInput(text), style: styleOf(cell), externalText: text });
      }
      width = Math.max(width, row.length);
      if (width > LIMITS.columns || (rows.length + 1) * width > LIMITS.rangeCells)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste dimensions are too large');
    }
    if (rows.length + 1 > LIMITS.rows)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste has too many rows');
    // Ragged rows are padded below, including entirely empty rows. Bound the
    // eventual rectangular allocation even when this row contains no cells.
    if ((rows.length + 1) * width > LIMITS.rangeCells)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste dimensions are too large');
    rows.push(row);
  }
  if (!rows.length || !width)
    throw new OpenSheetError('INVALID_ARGUMENT', 'Clipboard table has no cells');
  if (source && rows.some((row) => row.length !== width))
    throw new OpenSheetError('INVALID_ARGUMENT', 'Structured clipboard must be rectangular');
  return rows.map((row) =>
    Array.from({ length: width }, (_, i) => row[i] ?? { input: { type: 'blank' }, style: {} }),
  );
}
