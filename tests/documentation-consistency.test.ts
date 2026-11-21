import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { supportedFunctions } from '@opensheetjs/formula';

const read = (path: string) => readFileSync(path, 'utf8');

// Replace inline code with spaces so examples are not mistaken for live links.
// An unmatched delimiter remains prose, as in Markdown.
const withoutInlineCode = (text: string): string => {
  let result = '';
  let position = 0;
  const delimiters = /`+/g;
  for (let opening = delimiters.exec(text); opening; opening = delimiters.exec(text)) {
    let backslashes = 0;
    for (let i = opening.index - 1; i >= 0 && text[i] === '\\'; i--) backslashes++;
    if (backslashes % 2) continue;
    let closing = delimiters.exec(text);
    while (closing && closing[0].length !== opening[0].length) closing = delimiters.exec(text);
    if (!closing) {
      delimiters.lastIndex = opening.index + opening[0].length;
      continue;
    }
    result +=
      text.slice(position, opening.index) + ' '.repeat(delimiters.lastIndex - opening.index);
    position = delimiters.lastIndex;
  }
  return result + text.slice(position);
};

// This deliberately does not attempt to resolve reference usages or anchors as a full
// Markdown parser would. Both link forms must use the same fence rules.
const proseLines = (source: string): { line: number; text: string }[] => {
  const lines: { line: number; text: string }[] = [];
  let fence: { marker: string; length: number } | undefined;
  for (const [index, line] of source.split('\n').entries()) {
    const delimiter = line.match(/^ {0,3}(`{3,}|~{3,})/);
    if (delimiter) {
      const marker = delimiter[1][0];
      if (!fence) fence = { marker, length: delimiter[1].length };
      else if (
        marker === fence.marker &&
        delimiter[1].length >= fence.length &&
        line.slice(delimiter[0].length).trim() === ''
      )
        fence = undefined;
      continue;
    }
    if (fence) continue;
    lines.push({ line: index + 1, text: line });
  }
  return lines;
};

const referenceDestinations = (source: string): { line: number; href: string }[] => {
  return proseLines(source).flatMap(({ line, text }) => {
    const definition = text.match(/^ {0,3}\[[^\]]+\]:\s*(<[^>]+>|\S+)/);
    return definition ? [{ line, href: definition[1].replace(/^<|>$/g, '') }] : [];
  });
};

describe('selected capability documentation', () => {
  it('keeps relative inline documentation links pointed at existing files', () => {
    const markdownFiles = (directory: string): string[] =>
      readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
        const path = join(directory, entry.name);
        return entry.isDirectory() ? markdownFiles(path) : path.endsWith('.md') ? [path] : [];
      });
    const files = [
      'README.md',
      'ROADMAP.md',
      'TODO.md',
      ...markdownFiles('docs'),
      ...readdirSync('packages', { withFileTypes: true })
        .filter((entry) => entry.isDirectory())
        .map((entry) => join('packages', entry.name, 'README.md'))
        .filter(existsSync),
    ];
    const missing: string[] = [];
    for (const file of files) {
      for (const { line, text } of proseLines(read(file))) {
        for (const match of withoutInlineCode(text).matchAll(
          /!?\[[^\]]+\]\((<[^>]+>|[^\s)]+)(?:\s+"[^"]*")?\)/g,
        )) {
          const href = match[1].replace(/^<|>$/g, '');
          if (/^(?:[a-z][a-z\d+.-]*:|\/|#)/i.test(href)) continue;
          const path = href.split(/[?#]/, 1)[0];
          if (!path) continue;
          let target: string;
          try {
            target = resolve(dirname(file), decodeURIComponent(path));
          } catch {
            missing.push(`${file}:${line}: invalid URL ${href}`);
            continue;
          }
          if (!existsSync(target)) missing.push(`${file}:${line}: ${href}`);
        }
      }
      for (const { line, href } of referenceDestinations(read(file))) {
        if (/^(?:[a-z][a-z\d+.-]*:|\/|#)/i.test(href)) continue;
        const path = href.split(/[?#]/, 1)[0];
        try {
          if (path && !existsSync(resolve(dirname(file), decodeURIComponent(path)))) {
            missing.push(`${file}:${line}: ${href}`);
          }
        } catch {
          missing.push(`${file}:${line}: invalid URL ${href}`);
        }
      }
    }
    expect(missing).toEqual([]);
  });

  it('ignores inline code examples while retaining real links and unmatched delimiters', () => {
    expect(withoutInlineCode('`[example](missing.md)` [real](ok.md)')).toBe(
      '                        [real](ok.md)',
    );
    expect(withoutInlineCode('`` ` [example](missing.md) ` `` [real](ok.md)')).toMatch(
      /^\s+\[real\]\(ok\.md\)$/,
    );
    expect(withoutInlineCode('`unclosed [real](ok.md)')).toBe('`unclosed [real](ok.md)');
    expect(withoutInlineCode('`unclosed ``[example](missing.md)`` [real](ok.md)')).toMatch(
      /^`unclosed\s+\[real\]\(ok\.md\)$/,
    );
    expect(withoutInlineCode('\\` [real](ok.md)')).toBe('\\` [real](ok.md)');
  });

  it('extracts reference destinations outside matching fences', () => {
    expect(
      referenceDestinations(
        '[one]: <../a%20b.md#section> \"title\"\n' +
          '````md\n[ignored]: missing.md\n```\n[still-ignored]: missing.md\n````\n' +
          '~~~\n[also-ignored]: missing.md\n~~~\n[two]: other.md',
      ),
    ).toEqual([
      { line: 1, href: '../a%20b.md#section' },
      { line: 10, href: 'other.md' },
    ]);
  });

  it('ignores short and mismatched fences until the opening fence closes', () => {
    expect(
      proseLines(
        'before\n````md\n[ignored](missing.md)\n```\n~~~\n[ignored too](missing.md)\n````\n[kept](real.md)',
      ),
    ).toEqual([
      { line: 1, text: 'before' },
      { line: 8, text: '[kept](real.md)' },
    ]);
  });

  it('reports the actual runtime formula registry size', () => {
    const count = supportedFunctions.length;
    const formulas = read('docs/formulas.md').match(/registry now contains (\d+) functions/);
    const inventory = read('docs/capabilities.md').match(
      /\| (\d+) formula functions and reference transforms/,
    );
    expect(formulas).not.toBeNull();
    expect(inventory).not.toBeNull();
    expect(Number(formulas![1])).toBe(count);
    expect(Number(inventory![1])).toBe(count);
  });

  it('identifies the supported unversioned directory-head exception on current storage pages', () => {
    for (const path of [
      'docs/storage-format.md',
      'docs/02-data-and-commands.md',
      'docs/implementation-status.md',
    ]) {
      const text = read(path);
      expect(text, path).toMatch(/unversioned directory head/);
      expect(text, path).toMatch(/pre-0\.1\.1/i);
      expect(text, path).not.toMatch(
        /older files and persisted engine formats are rejected|old files and databases are not migrated/i,
      );
    }
  });
});
