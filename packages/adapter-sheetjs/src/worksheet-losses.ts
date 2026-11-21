import { LIMITS, OpenSheetError, address, parseAddress, parseRange } from '@opensheetjs/core';
import { add } from './report.js';
import type { CompatibilityReport, SheetJSWorkbook } from './types.js';

const worksheetContentType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
// SheetJS resolves ZIP paths case-insensitively and accepts a leading slash.
const partName = (name: string) => name.replace(/^\//, '').toLowerCase();

function xmlEncoding(content: Uint8Array): 'utf-8' | 'utf-16le' | 'utf-16be' {
  if (content[0] === 0xff && content[1] === 0xfe) return 'utf-16le';
  if (content[0] === 0xfe && content[1] === 0xff) return 'utf-16be';
  if (content[1] === 0 && content[3] === 0 && content[0] !== 0 && content[2] !== 0)
    return 'utf-16le';
  if (content[0] === 0 && content[2] === 0 && content[1] !== 0 && content[3] !== 0)
    return 'utf-16be';
  return 'utf-8';
}

// SheetJS resolves OOXML escape-shaped strings while parsing XML. The raw text is
// ambiguous: an escape may be intentional or a literal that the parser changes.
function containsEscapeSequence(
  content: Uint8Array,
  budget: { remaining: number },
): boolean | undefined {
  const bomLittle = content[0] === 0xff && content[1] === 0xfe;
  const bomBig = content[0] === 0xfe && content[1] === 0xff;
  const little = xmlEncoding(content) === 'utf-16le';
  const big = xmlEncoding(content) === 'utf-16be';
  const width = little || big ? 2 : 1;
  const unit = (offset: number) =>
    width === 1
      ? content[offset]
      : little
        ? content[offset] | (content[offset + 1] << 8)
        : (content[offset] << 8) | content[offset + 1];
  const hex = (n: number) => (n >= 48 && n <= 57) || (n >= 65 && n <= 70) || (n >= 97 && n <= 102);
  // Numeric XML character references are decoded before SheetJS recognizes
  // _xHHHH_. Limit each reference as well as the total number of candidates.
  let tooLong = false;
  const read = (start: number): { value: number; next: number } | undefined => {
    if (start < 0 || start + width > content.length) return;
    const first = unit(start);
    if (first !== 38 || start + 2 * width > content.length || unit(start + width) !== 35)
      return { value: first, next: start + width };
    let position = start + 2 * width;
    const isHex = unit(position) === 120 || unit(position) === 88;
    if (isHex) position += width;
    let text = '';
    while (position + width <= content.length && unit(position) !== 59) {
      const char = unit(position);
      if (!(isHex ? hex(char) : char >= 48 && char <= 57))
        return { value: first, next: start + width };
      if (text.length >= 20) {
        tooLong = true;
        return;
      }
      text += String.fromCharCode(char);
      position += width;
    }
    if (!text || position + width > content.length) return { value: first, next: start + width };
    return { value: Number.parseInt(text, isHex ? 16 : 10) & 0xffff, next: position + width };
  };
  const startOffset = bomLittle || bomBig ? 2 : 0;
  let underscore = content.indexOf(95, startOffset);
  let ampersand = content.indexOf(38, startOffset);
  while (underscore >= 0 || ampersand >= 0) {
    const offset =
      underscore < 0 ? ampersand : ampersand < 0 ? underscore : Math.min(underscore, ampersand);
    if (offset === underscore) underscore = content.indexOf(95, offset + 1);
    if (offset === ampersand) ampersand = content.indexOf(38, offset + 1);
    if (budget.remaining-- <= 0) return undefined;
    const start = big ? offset - 1 : offset;
    if (start < 0 || (width !== 1 && start % 2 !== 0)) continue;
    let position = start;
    let matches = true;
    for (let i = 0; i < 7; i++) {
      const next = read(position);
      if (tooLong) return undefined;
      if (
        !next ||
        (i === 0 || i === 6
          ? next.value !== 95
          : i === 1
            ? next.value !== 120 && next.value !== 88
            : !hex(next.value))
      ) {
        matches = false;
        break;
      }
      position = next.next;
    }
    if (matches) return true;
  }
  return false;
}

function decodeXml(content: Uint8Array) {
  return new TextDecoder(xmlEncoding(content)).decode(content);
}

function xmlAttribute(value: string) {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|quot|apos|lt|gt);/gi, (entity, code: string) => {
    if (code[0] === '#') {
      const point = /^#x/i.test(code) ? Number.parseInt(code.slice(2), 16) : Number(code.slice(1));
      return point > 0 && point <= 0x10ffff ? String.fromCodePoint(point) : entity;
    }
    return (
      ({ amp: '&', quot: '"', apos: "'", lt: '<', gt: '>' } as Record<string, string>)[code] ??
      entity
    );
  });
}

/** Scan markup once, skipping non-element constructs and respecting quoted attributes. */
function* elementTags(xml: string): Generator<string> {
  let position = 0;
  while (position < xml.length) {
    const start = xml.indexOf('<', position);
    if (start < 0) return;
    let terminator: string | undefined;
    let offset = 0;
    if (xml.startsWith('<!--', start)) {
      terminator = '-->';
      offset = 4;
    } else if (xml.startsWith('<![CDATA[', start)) {
      terminator = ']]>';
      offset = 9;
    } else if (xml.startsWith('<?', start)) {
      terminator = '?>';
      offset = 2;
    }
    if (terminator) {
      const end = xml.indexOf(terminator, start + offset);
      if (end < 0) throw new OpenSheetError('INVALID_ARGUMENT', 'Unterminated XLSX XML markup');
      position = end + terminator.length;
      continue;
    }
    let quote = '';
    let end = start + 1;
    for (; end < xml.length; end++) {
      const char = xml[end];
      if (quote) {
        if (char === quote) quote = '';
      } else if (char === '"' || char === "'") quote = char;
      else if (char === '>') break;
    }
    if (end === xml.length)
      throw new OpenSheetError('INVALID_ARGUMENT', 'Unterminated XLSX XML markup');
    yield xml.slice(start, end + 1);
    position = end + 1;
  }
}

function declaredWorksheets(xml: string) {
  const paths = new Set<string>();
  for (const tag of elementTags(xml)) {
    const attributes = /^<(?:[\w.-]+:)?Override(?=[\s/>])([^]*)>$/.exec(tag)?.[1];
    if (!attributes) continue;
    const fields = new Map<string, string>();
    const name = /[\w.-]+/y;
    for (let i = 0; i < attributes.length; ) {
      name.lastIndex = i;
      const match = name.exec(attributes);
      if (!match) {
        i++;
        continue;
      }
      i = name.lastIndex;
      while (i < attributes.length && /\s/.test(attributes[i])) i++;
      if (attributes[i] !== '=') continue;
      i++;
      while (i < attributes.length && /\s/.test(attributes[i])) i++;
      const quote = attributes[i];
      if (quote !== '"' && quote !== "'") continue;
      const end = attributes.indexOf(quote, i + 1);
      if (end < 0) throw new OpenSheetError('INVALID_ARGUMENT', 'Unterminated XLSX XML markup');
      fields.set(match[0], xmlAttribute(attributes.slice(i + 1, end)));
      i = end + 1;
    }
    const path = fields.get('PartName');
    if (fields.get('ContentType') === worksheetContentType && path?.startsWith('/')) {
      paths.add(partName(path));
      try {
        paths.add(partName(decodeURI(path)));
      } catch {
        /* Keep the literal path. */
      }
    }
  }
  return paths;
}

type Freeze = { rows: number; columns: number };

// Return no mapping on ambiguity: importing a pane onto the wrong sheet is worse than a warning.
function attributes(tag: string): Map<string, string> | undefined {
  if (tag.length > 16 * 1024) return;
  const head = /^<[\w.:-]+(?=[\s/>])/.exec(tag);
  if (!head) return;
  const fields = new Map<string, string>();
  let position = head[0].length;
  const name = /[\w.:-]+/y;
  while (position < tag.length) {
    while (/\s/.test(tag[position] ?? '')) position++;
    if (tag[position] === '>' || (tag[position] === '/' && tag[position + 1] === '>'))
      return fields;
    name.lastIndex = position;
    const match = name.exec(tag);
    if (!match || fields.has(match[0])) return;
    position += match[0].length;
    while (/\s/.test(tag[position] ?? '')) position++;
    if (tag[position++] !== '=') return;
    while (/\s/.test(tag[position] ?? '')) position++;
    const quote = tag[position++];
    if (quote !== '"' && quote !== "'") return;
    const end = tag.indexOf(quote, position);
    if (end < 0) return;
    fields.set(match[0], xmlAttribute(tag.slice(position, end)));
    position = end + 1;
  }
}

function relationshipTarget(target: string): string | undefined {
  // Match only canonical local paths; encoded or dot-segment paths are left as reported loss.
  if (/[\\:?#%\u0000-\u001f\u007f]/.test(target)) return;
  const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  const parts = path.split('/');
  if (parts.some((component) => !component || component === '.' || component === '..')) return;
  return partName(path);
}

const mainNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships';

function rawWorkbookSheets(content: Uint8Array): Array<{ name: string; id: string }> | undefined {
  const stack: string[] = [];
  const sheets: Array<{ name: string; id: string }> = [];
  let rootSeen = false;
  let relNamespace: string | undefined;
  for (const tag of elementTags(decodeXml(content))) {
    const node = /^<(\/?)([\w.:-]+)(?=[\s/>])/.exec(tag);
    if (!node) {
      if (tag.startsWith('<!')) return;
      continue;
    }
    const [closing, name] = [node[1] === '/', node[2]];
    if (closing) {
      if (stack.pop() !== name) return;
      continue;
    }
    if (!rootSeen) {
      rootSeen = true;
      const root = attributes(tag);
      if (name !== 'workbook' || root?.get('xmlns') !== mainNamespace) return;
      relNamespace = root.get('xmlns:r');
    } else if (!stack.length) return;
    if (name === 'sheets' && stack.length === 1 && stack[0] === 'workbook') {
      const fields = attributes(tag);
      if (!fields || (fields.has('xmlns') && fields.get('xmlns') !== mainNamespace)) return;
    }
    if (
      name === 'sheet' &&
      stack.length === 2 &&
      stack[0] === 'workbook' &&
      stack[1] === 'sheets'
    ) {
      const fields = attributes(tag);
      if (
        !fields ||
        (fields.has('xmlns') && fields.get('xmlns') !== mainNamespace) ||
        (fields.get('xmlns:r') ?? relNamespace) !== relationshipNamespace
      )
        return;
      const sheetName = fields.get('name'),
        id = fields.get('r:id');
      if (!sheetName || !id) return;
      sheets.push({ name: sheetName, id });
      if (sheets.length > LIMITS.sheets) return;
    }
    if (!/\/\s*>$/.test(tag)) {
      if (stack.length >= 128) return;
      stack.push(name);
    }
  }
  if (!rootSeen || stack.length) return;
  return sheets;
}

function worksheetMappings(
  parts: Map<string, { content?: Uint8Array } | undefined>,
  names: string[],
  sheets: any,
): Map<string, string> {
  const output = new Map<string, string>();
  if (!Array.isArray(sheets) || sheets.length !== names.length) return output;
  const workbook = parts.get('xl/workbook.xml')?.content;
  if (!(workbook instanceof Uint8Array) || workbook.byteLength > 4 * 1024 * 1024) return output;
  const rawSheets = rawWorkbookSheets(workbook);
  if (
    !rawSheets ||
    rawSheets.length !== names.length ||
    rawSheets.some(
      (sheet, i) =>
        sheet.name !== names[i] || sheets[i]?.name !== sheet.name || sheets[i]?.id !== sheet.id,
    )
  )
    return output;
  const rels = parts.get('xl/_rels/workbook.xml.rels')?.content;
  if (!(rels instanceof Uint8Array) || rels.byteLength > 4 * 1024 * 1024) return output;
  const links = new Map<string, string>();
  const duplicates = new Set<string>();
  const seenIds = new Set<string>();
  let rootSeen = false;
  const stack: string[] = [];
  for (const tag of elementTags(decodeXml(rels))) {
    const node = /^<(\/?)([\w.:-]+)(?=[\s/>])/.exec(tag);
    if (!node) {
      if (tag.startsWith('<!')) return output;
      continue;
    }
    if (node[1]) {
      if (stack.pop() !== node[2]) return output;
      continue;
    }
    if (!rootSeen) {
      rootSeen = true;
      if (
        node[2] !== 'Relationships' ||
        attributes(tag)?.get('xmlns') !== packageRelationshipNamespace
      )
        return output;
    } else if (!stack.length) return output;
    const isRelationship =
      node[2] === 'Relationship' && stack.length === 1 && stack[0] === 'Relationships';
    if (!/\/\s*>$/.test(tag)) {
      if (stack.length >= 128) return output;
      stack.push(node[2]);
    }
    if (!isRelationship) continue;
    const fields = attributes(tag);
    if (!fields) return output;
    if (fields.has('xmlns') && fields.get('xmlns') !== packageRelationshipNamespace) return output;
    const id = fields?.get('Id'),
      target = fields?.get('Target');
    if (id && seenIds.has(id)) duplicates.add(id);
    else if (id) seenIds.add(id);
    if (
      !id ||
      !target ||
      (fields?.has('TargetMode') && fields.get('TargetMode') !== 'Internal') ||
      fields?.get('Type') !==
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
    )
      continue;
    const path = relationshipTarget(target);
    if (!path) continue;
    if (links.has(id)) duplicates.add(id);
    else links.set(id, path);
  }
  if (!rootSeen || stack.length) return output;
  const pathOwners = new Map<string, string[]>();
  for (let i = 0; i < names.length; i++) {
    const sheet = sheets[i];
    if (
      !sheet ||
      sheet.name !== names[i] ||
      typeof sheet.id !== 'string' ||
      duplicates.has(sheet.id)
    )
      continue;
    const path = links.get(sheet.id);
    if (!path) continue;
    const owners = pathOwners.get(path) ?? [];
    owners.push(names[i]);
    pathOwners.set(path, owners);
  }
  for (const [path, owners] of pathOwners) if (owners.length === 1) output.set(path, owners[0]);
  return output;
}

/** Return only unambiguous raw external hyperlink relationships within a bounded scan. */
export function rawHyperlinkTargets(
  files: SheetJSWorkbook['files'],
  names: string[],
  sheets: unknown,
): Map<
  string,
  {
    links: Map<string, { target: string; byteDecoded: string }>;
    cells: Map<string, { id: string; ref: string; location?: string }>;
  }
> {
  const output = new Map<
    string,
    {
      links: Map<string, { target: string; byteDecoded: string }>;
      cells: Map<string, { id: string; ref: string; location?: string }>;
    }
  >();
  if (!files) return output;
  const parts = new Map<string, { content?: Uint8Array }>();
  for (const [name, part] of Object.entries(files)) {
    const path = partName(name);
    if (parts.has(path)) return new Map();
    parts.set(path, part);
  }
  const mappings = worksheetMappings(parts, names, sheets);
  let totalBytes = 0;
  let totalRelations = 0;
  let totalCells = 0;
  for (const [path, name] of mappings) {
    const slash = path.lastIndexOf('/');
    const relationshipPath = `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`;
    const content = parts.get(relationshipPath)?.content;
    if (!(content instanceof Uint8Array)) continue;
    totalBytes += content.byteLength;
    if (content.byteLength > 4 * 1024 * 1024 || totalBytes > 16 * 1024 * 1024) return new Map();
    const links = new Map<string, { target: string; byteDecoded: string }>();
    const seen = new Set<string>();
    const duplicate = new Set<string>();
    const stack: string[] = [];
    let rootSeen = false;
    try {
      const xml = new TextDecoder(xmlEncoding(content), { fatal: true }).decode(content);
      for (const tag of elementTags(xml)) {
        const node = /^<(\/?)([\w.:-]+)(?=[\s/>])/.exec(tag);
        if (!node) {
          if (tag.startsWith('<!')) return new Map();
          continue;
        }
        if (node[1]) {
          if (stack.pop() !== node[2]) return new Map();
          continue;
        }
        if (!rootSeen) {
          rootSeen = true;
          if (
            node[2] !== 'Relationships' ||
            attributes(tag)?.get('xmlns') !== packageRelationshipNamespace
          )
            return new Map();
        } else if (!stack.length) return new Map();
        const relationship =
          node[2] === 'Relationship' && stack.length === 1 && stack[0] === 'Relationships';
        if (rootSeen && stack.length && !relationship) return new Map();
        if (!/\/\s*>$/.test(tag)) {
          if (stack.length >= 128) return new Map();
          stack.push(node[2]);
        }
        if (!relationship) continue;
        if (++totalRelations > 20000) return new Map();
        const fields = attributes(tag);
        if (
          !fields ||
          (fields.has('xmlns') && fields.get('xmlns') !== packageRelationshipNamespace)
        )
          return new Map();
        const id = fields.get('Id');
        if (!id) return new Map();
        if (seen.has(id)) duplicate.add(id);
        else seen.add(id);
        const target = fields.get('Target');
        if (
          target &&
          target.length <= 4096 &&
          /[^\x00-\x7f]/.test(target) &&
          !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(target) &&
          fields.get('Type') === `${relationshipNamespace}/hyperlink` &&
          fields.get('TargetMode') === 'External'
        ) {
          const bytes = new TextEncoder().encode(target);
          let byteDecoded = '';
          for (let offset = 0; offset < bytes.length; offset += 8192)
            byteDecoded += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
          links.set(id, { target, byteDecoded });
        }
      }
      if (!rootSeen || stack.length) return new Map();
    } catch {
      return new Map();
    }
    for (const id of duplicate) links.delete(id);
    if (!links.size) continue;
    const worksheet = parts.get(path)?.content;
    if (!(worksheet instanceof Uint8Array)) continue;
    totalBytes += worksheet.byteLength;
    if (worksheet.byteLength > 4 * 1024 * 1024 || totalBytes > 16 * 1024 * 1024) return new Map();
    const cells = new Map<string, { id: string; ref: string; location?: string }>();
    const duplicateCells = new Set<string>();
    const tree: string[] = [];
    let worksheetRoot = false;
    let hyperlinkContainer = false;
    let rootRelationshipNamespace: string | undefined;
    let containerRelationshipNamespace: string | undefined;
    try {
      const xml = new TextDecoder(xmlEncoding(worksheet), { fatal: true }).decode(worksheet);
      for (const tag of elementTags(xml)) {
        const node = /^<(\/?)([\w.:-]+)(?=[\s/>])/.exec(tag);
        if (!node) {
          if (tag.startsWith('<!')) return new Map();
          continue;
        }
        if (node[1]) {
          if (tree.pop() !== node[2]) return new Map();
          continue;
        }
        if (!worksheetRoot) {
          worksheetRoot = true;
          const fields = attributes(tag);
          if (
            node[2] !== 'worksheet' ||
            fields?.get('xmlns') !== mainNamespace ||
            (fields.has('xmlns:r') && fields.get('xmlns:r') !== relationshipNamespace)
          )
            return new Map();
          rootRelationshipNamespace = fields.get('xmlns:r');
        } else if (!tree.length) return new Map();
        if (node[2] === 'hyperlinks') {
          const fields = attributes(tag);
          if (
            tree.length !== 1 ||
            tree[0] !== 'worksheet' ||
            hyperlinkContainer ||
            !fields ||
            (fields.has('xmlns') && fields.get('xmlns') !== mainNamespace) ||
            (fields.has('xmlns:r') && fields.get('xmlns:r') !== relationshipNamespace)
          )
            return new Map();
          hyperlinkContainer = true;
          containerRelationshipNamespace = fields.get('xmlns:r') ?? rootRelationshipNamespace;
        }
        if (node[2] === 'hyperlink') {
          const fields = attributes(tag);
          if (
            tree.length !== 2 ||
            tree[0] !== 'worksheet' ||
            tree[1] !== 'hyperlinks' ||
            !fields ||
            (fields.has('xmlns') && fields.get('xmlns') !== mainNamespace) ||
            (fields.has('xmlns:r') && fields.get('xmlns:r') !== relationshipNamespace) ||
            (fields.has('r:id') &&
              (fields.get('xmlns:r') ?? containerRelationshipNamespace) !== relationshipNamespace)
          )
            return new Map();
          if (++totalRelations > 20000) return new Map();
          const ref = fields.get('ref');
          if (!ref || !/^[A-Z]+[1-9]\d*(?::[A-Z]+[1-9]\d*)?$/.test(ref)) return new Map();
          const range = parseRange(ref);
          const endpoints = ref.split(':');
          if (
            endpoints[0] !== address(range.startRow, range.startColumn) ||
            (endpoints[1] !== undefined &&
              endpoints[1] !== address(range.endRow - 1, range.endColumn - 1))
          )
            return new Map();
          const size = (range.endRow - range.startRow) * (range.endColumn - range.startColumn);
          // Bound total expansion, including overlapping and non-external hyperlinks.
          totalCells += size;
          if (totalCells > 16384) return new Map();
          for (let row = range.startRow; row < range.endRow; row++) {
            for (let column = range.startColumn; column < range.endColumn; column++) {
              const cell = address(row, column);
              if (cells.has(cell)) duplicateCells.add(cell);
              else
                cells.set(cell, {
                  id: fields.get('r:id') ?? '',
                  ref,
                  ...(fields.has('location') ? { location: fields.get('location')! } : {}),
                });
            }
          }
        }
        if (!/\/\s*>$/.test(tag)) {
          if (tree.length >= 128) return new Map();
          tree.push(node[2]);
        }
      }
      if (!worksheetRoot || tree.length) return new Map();
    } catch {
      return new Map();
    }
    for (const ref of duplicateCells) cells.delete(ref);
    output.set(name, { links, cells });
  }
  return output;
}

/** Verify generated workbook relationships before patching a worksheet part. */
export function generatedWorksheetPaths(
  files: SheetJSWorkbook['files'],
  names: string[],
): Map<string, string> {
  const parts = new Map<string, { content?: Uint8Array }>();
  for (const [name, part] of Object.entries(files ?? {})) {
    const path = partName(name);
    if (parts.has(path)) throw new OpenSheetError('INVALID_ARGUMENT', 'Ambiguous XLSX part names');
    parts.set(path, part);
  }
  const workbook = parts.get('xl/workbook.xml')?.content;
  if (!(workbook instanceof Uint8Array) || workbook.byteLength > 4 * 1024 * 1024)
    throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid generated XLSX workbook');
  const sheets = rawWorkbookSheets(workbook);
  if (!sheets) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid generated XLSX worksheets');
  const mapping = worksheetMappings(parts, names, sheets);
  if (mapping.size !== names.length)
    throw new OpenSheetError('INVALID_ARGUMENT', 'Ambiguous generated XLSX worksheet mapping');
  return mapping;
}

function importableFreeze(tag: string): Freeze | undefined {
  const fields = attributes(tag);
  if (!fields || fields.get('state') !== 'frozen') return;
  if (
    [...fields.keys()].some(
      (key) =>
        !['state', 'xSplit', 'ySplit', 'topLeftCell', 'activePane', 'xmlns'].includes(key) &&
        !key.startsWith('xmlns:'),
    )
  )
    return;
  const number = (key: string) => {
    const value = fields.get(key);
    if (value === undefined) return 0;
    if (!/^\d+(?:\.0+)?$/.test(value)) return NaN;
    return Number(value);
  };
  const columns = number('xSplit'),
    rows = number('ySplit');
  if (
    !Number.isSafeInteger(rows) ||
    !Number.isSafeInteger(columns) ||
    (!rows && !columns) ||
    rows >= LIMITS.rows ||
    columns >= LIMITS.columns
  )
    return;
  const pane = fields.get('activePane');
  const expectedPane = rows && columns ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  if (pane !== undefined && pane !== expectedPane) return;
  const topLeft = fields.get('topLeftCell');
  if (!topLeft) return;
  try {
    const position = parseAddress(topLeft);
    if (position.row !== rows || position.column !== columns) return;
  } catch {
    return;
  }
  return { rows, columns };
}

/** Neutral view settings seen in independently saved files do not change pane geometry. */
function neutralSheetView(view: Map<string, string>): boolean {
  const neutral: Record<string, readonly string[]> = {
    workbookViewId: ['0'],
    showFormulas: ['false', '0'],
    showGridLines: ['true', '1'],
    showRowColHeaders: ['true', '1'],
    showZeros: ['true', '1'],
    rightToLeft: ['false', '0'],
    showOutlineSymbols: ['true', '1'],
    defaultGridColor: ['true', '1'],
    view: ['normal'],
    topLeftCell: ['A1'],
    colorId: ['64'],
    zoomScale: ['100'],
    zoomScaleNormal: ['100'],
    zoomScalePageLayoutView: ['100'],
    tabSelected: ['true', 'false', '0', '1'],
  };
  return (
    view.get('workbookViewId') === '0' &&
    [...view].every(([key, value]) => Object.hasOwn(neutral, key) && neutral[key].includes(value))
  );
}

/** Inspect raw worksheet parts because the parsed SheetJS cell model omits these rules. */
export function reportWorksheetLosses(
  files: SheetJSWorkbook['files'],
  report: CompatibilityReport,
  sheetNames: string[] = [],
  workbookSheets?: unknown,
): Map<string, Freeze> {
  const freezes = new Map<string, Freeze>();
  if (!files) return freezes;
  let bytes = 0;
  let validation = 0;
  let formatting = 0;
  let panes = 0;
  let selections = 0;
  let selectedTabs = 0;
  const parts = new Map<string, (typeof files)[string]>();
  for (const [name, part] of Object.entries(files)) {
    const path = partName(name);
    if (parts.has(path)) throw new OpenSheetError('INVALID_ARGUMENT', 'Ambiguous XLSX part names');
    parts.set(path, part);
  }
  const manifest = parts.get('[content_types].xml')?.content;
  if (manifest instanceof Uint8Array && manifest.byteLength > 4 * 1024 * 1024)
    throw new OpenSheetError(
      'LIMIT_EXCEEDED',
      'XLSX content type manifest exceeds inspection budget',
    );
  const declared =
    manifest instanceof Uint8Array ? declaredWorksheets(decodeXml(manifest)) : new Set<string>();
  let inspectedBytes = 0;
  let ambiguousParts = 0;
  let incompleteParts = 0;
  const escapeBudget = { remaining: 1_000_000 };
  for (const [name, part] of parts) {
    if (
      !(
        declared.has(name) ||
        /^xl\/(?:workbook|sharedstrings|styles|comments\d*)\.xml$/.test(name) ||
        /^xl\/worksheets\/[^/]+\.xml$/.test(name) ||
        /^xl\/(?:_rels|worksheets\/_rels)\/[^/]+\.rels$/.test(name)
      )
    )
      continue;
    const content = part?.content;
    if (!(content instanceof Uint8Array)) continue;
    inspectedBytes += content.byteLength;
    if (inspectedBytes > 200 * 1024 * 1024)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'XLSX text inspection exceeds budget');
    const found = containsEscapeSequence(content, escapeBudget);
    if (found === undefined) incompleteParts++;
    else if (found) ambiguousParts++;
  }
  if (ambiguousParts)
    add(report, {
      code: 'XLSX_ESCAPE_IMPORT_AMBIGUOUS',
      feature: 'text',
      action: 'approximated',
      severity: 'warning',
      message:
        'XLSX escape-shaped text may have changed during parsing; verify text, formulas, sheet names, notes and links',
      count: ambiguousParts,
    });
  if (incompleteParts)
    add(report, {
      code: 'XLSX_ESCAPE_INSPECTION_INCOMPLETE',
      feature: 'text',
      action: 'approximated',
      severity: 'warning',
      message: 'XLSX escape inspection reached its work limit; verify imported text and links',
      count: incompleteParts,
    });
  const mappings = worksheetMappings(parts, sheetNames, workbookSheets);
  for (const [name, part] of parts) {
    if (declared.size ? !declared.has(name) : !/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
    const content = part?.content;
    if (!(content instanceof Uint8Array)) continue;
    bytes += content.byteLength;
    if (bytes > 200 * 1024 * 1024)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Worksheet XML exceeds inspection budget');
    const xml = decodeXml(content);
    // Unfrozen sheets can still carry an active tab or selection that the
    // native model does not preserve. Do not limit inspection to pane tags.
    const inspectTree =
      xml.includes('pane') || xml.includes('tabSelected') || xml.includes('<selection');
    let viewCount = 0;
    let safeTree = inspectTree;
    let candidateInView = false;
    let rootSeen = false;
    const stack: string[] = [];
    let paneCount = 0;
    let firstPane: string | undefined;
    for (const tag of elementTags(xml)) {
      // Avoid a regular-expression match allocation for every cell element in
      // large worksheets while retaining the same accepted ASCII name grammar.
      const closing = inspectTree && tag[1] === '/';
      const nameStart = closing ? 2 : 1;
      let nameEnd = nameStart;
      if (inspectTree)
        while (nameEnd < tag.length) {
          const code = tag.charCodeAt(nameEnd);
          if (
            (code >= 65 && code <= 90) ||
            (code >= 97 && code <= 122) ||
            (code >= 48 && code <= 57) ||
            code === 95 ||
            code === 46 ||
            code === 58 ||
            code === 45
          )
            nameEnd++;
          else break;
        }
      const delimiter = tag[nameEnd];
      const validName =
        inspectTree &&
        nameEnd > nameStart &&
        (delimiter === '/' || delimiter === '>' || /\s/.test(delimiter));
      if (validName) {
        const nodeName = tag.slice(nameStart, nameEnd);
        if (closing) {
          if (stack.pop() !== nodeName) safeTree = false;
        } else {
          if (!rootSeen) {
            rootSeen = true;
            const root = attributes(tag);
            if (
              nodeName !== 'worksheet' ||
              root?.get('xmlns') !== 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            )
              safeTree = false;
          } else if (!stack.length) safeTree = false;
          else if (
            /\sxmlns(?::[\w.-]+)?\s*=/.test(tag) &&
            (nodeName === 'sheetViews' || nodeName === 'sheetView' || nodeName === 'pane')
          )
            safeTree = false;
          if (
            nodeName === 'selection' &&
            stack.length === 3 &&
            stack[0] === 'worksheet' &&
            stack[1] === 'sheetViews' &&
            stack[2] === 'sheetView'
          ) {
            const selection = attributes(tag);
            if (
              !selection ||
              (selection.get('activeCell') !== undefined && selection.get('activeCell') !== 'A1') ||
              (selection.get('sqref') !== undefined && selection.get('sqref') !== 'A1') ||
              [...selection.keys()].some((key) => !['pane', 'activeCell', 'sqref'].includes(key))
            )
              selections++;
          }
          if (nodeName === 'sheetView') {
            viewCount++;
            const view = attributes(tag);
            if (!view || !neutralSheetView(view)) safeTree = false;
            if (
              view &&
              stack.length === 2 &&
              stack[0] === 'worksheet' &&
              stack[1] === 'sheetViews' &&
              ['true', '1'].includes(view.get('tabSelected') ?? '')
            )
              selectedTabs++;
          }
          if (
            nodeName === 'pane' &&
            stack.length === 3 &&
            stack[0] === 'worksheet' &&
            stack[1] === 'sheetViews' &&
            stack[2] === 'sheetView'
          )
            candidateInView = true;
          if (!/\/\s*>$/.test(tag)) {
            if (stack.length >= 128) safeTree = false;
            else stack.push(nodeName);
          }
        }
      } else if (tag.startsWith('<!')) safeTree = false;
      const match = /^<(?:[\w.-]+:)?(dataValidation|conditionalFormatting|pane)(?=[\s/>])/.exec(
        tag,
      );
      if (!match) continue;
      if (match[1] === 'dataValidation') validation++;
      else if (match[1] === 'conditionalFormatting') formatting++;
      else {
        paneCount++;
        if (paneCount === 1) firstPane = tag;
      }
    }
    if (stack.length) safeTree = false;
    const sheetName = mappings.get(name);
    const freeze =
      safeTree && candidateInView && paneCount === 1 && viewCount === 1 && sheetName && firstPane
        ? importableFreeze(firstPane)
        : undefined;
    if (freeze && sheetName) freezes.set(sheetName, freeze);
    else panes += paneCount;
  }
  if (validation)
    add(report, {
      code: 'DATA_VALIDATION_UNSUPPORTED',
      feature: 'data-validation',
      action: 'dropped',
      severity: 'warning',
      message: 'Data validation rules are not imported',
      count: validation,
    });
  if (formatting)
    add(report, {
      code: 'CONDITIONAL_FORMATTING_UNSUPPORTED',
      feature: 'conditional-formatting',
      action: 'dropped',
      severity: 'warning',
      message: 'Conditional formatting is not imported',
      count: formatting,
    });
  if (panes)
    add(report, {
      code: 'WORKSHEET_PANE_UNSUPPORTED',
      feature: 'worksheet-pane',
      action: 'dropped',
      severity: 'warning',
      message: 'Some frozen or split worksheet panes are not imported',
      count: panes,
    });
  if (selections)
    add(report, {
      code: 'WORKSHEET_SELECTION_DROPPED',
      feature: 'worksheet-selection',
      action: 'dropped',
      severity: 'warning',
      message: 'Worksheet selection and active cell are not imported',
      count: selections,
    });
  if (selectedTabs)
    add(report, {
      code: 'WORKSHEET_ACTIVE_TAB_DROPPED',
      feature: 'worksheet-view',
      action: 'dropped',
      severity: 'warning',
      message: 'Worksheet tab selection is not imported',
      count: selectedTabs,
    });
  return freezes;
}
