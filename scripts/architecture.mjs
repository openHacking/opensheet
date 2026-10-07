import { readdirSync, readFileSync } from 'node:fs';
import { dirname, posix, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const layers = {
  formula: [],
  core: ['formula'],
  renderer: ['core'],
  'plugin-sdk': ['core'],
  'adapter-sheetjs': ['core'],
  formats: ['core'],
  opensheet: ['core', 'renderer', 'plugin-sdk'],
  react: ['opensheet'],
  vue: ['opensheet'],
  testing: ['core'],
};
const owner = (file) => /^packages\/([^/]+)\/src\//.exec(file)?.[1];
const entry = (name) => `packages/${name}/src/index.ts`;
const packageName = (specifier) =>
  specifier === 'opensheet' ? 'opensheet' : /^@opensheetjs\/([^/]+)$/.exec(specifier)?.[1];

function imports(source) {
  const result = [];
  const add = (specifier, runtime) => {
    if (ts.isStringLiteralLike(specifier)) result.push({ specifier: specifier.text, runtime });
  };
  function visit(node) {
    if (ts.isImportDeclaration(node)) {
      const clause = node.importClause;
      const named = clause?.namedBindings;
      const runtime =
        !clause ||
        (!clause.isTypeOnly &&
          (!!clause.name ||
            !named ||
            ts.isNamespaceImport(named) ||
            named.elements.length === 0 ||
            named.elements.some((item) => !item.isTypeOnly)));
      add(node.moduleSpecifier, runtime);
    } else if (ts.isExportDeclaration(node) && node.moduleSpecifier) {
      const named = node.exportClause;
      add(
        node.moduleSpecifier,
        !node.isTypeOnly &&
          (!named ||
            !ts.isNamedExports(named) ||
            named.elements.length === 0 ||
            named.elements.some((item) => !item.isTypeOnly)),
      );
    } else if (
      ts.isCallExpression(node) &&
      (node.expression.kind === ts.SyntaxKind.ImportKeyword ||
        (ts.isIdentifier(node.expression) && node.expression.text === 'require')) &&
      node.arguments[0]
    ) {
      add(node.arguments[0], true);
    } else if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument)) {
      add(node.argument.literal, false);
    }
    ts.forEachChild(node, visit);
  }
  visit(source);
  return result;
}

/** Check an in-memory source graph so boundary rules can be tested without editing files. */
export function analyzeArchitecture(files) {
  const graph = new Map();
  const errors = [];
  for (const [file, text] of files) {
    const current = owner(file);
    const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
    const edges = [];
    for (const { specifier, runtime } of imports(source)) {
      const dependency = packageName(specifier);
      let target;
      if (specifier.startsWith('.')) {
        const base = posix.normalize(posix.join(posix.dirname(file), specifier));
        target = [base.replace(/\.js$/, '.ts'), `${base}.ts`, `${base}/index.ts`, base].find(
          (candidate) => files.has(candidate),
        );
        if (!target && !specifier.endsWith('.css'))
          errors.push(`${file}: unresolved local import ${specifier}`);
        if (current && target && owner(target) !== current)
          errors.push(`${file}: cross-package source import ${specifier}`);
      } else if (dependency && files.has(entry(dependency))) {
        target = entry(dependency);
        if (current && dependency !== current && !layers[current]?.includes(dependency))
          errors.push(`${file}: forbidden dependency on ${dependency}`);
      } else if (current && /^(?:opensheet|@opensheetjs\/[^/]+)\//.test(specifier)) {
        errors.push(`${file}: package subpath import ${specifier}`);
      }
      if (current && file !== entry(current) && target === entry(current))
        errors.push(`${file}: internal module imports its public entry`);
      if (
        current &&
        ['core', 'formula'].includes(current) &&
        (/^(?:react|vue|lucide|xlsx)(?:\/|$)/.test(specifier) || target?.startsWith('apps/'))
      )
        errors.push(`${file}: browser/framework dependency ${specifier}`);
      if (target && runtime) edges.push(target);
    }
    graph.set(file, edges);
  }
  const visiting = new Set(),
    visited = new Set(),
    path = [];
  function visit(file) {
    if (visiting.has(file)) {
      errors.push(`Runtime cycle: ${[...path.slice(path.indexOf(file)), file].join(' -> ')}`);
      return;
    }
    if (visited.has(file)) return;
    visiting.add(file);
    path.push(file);
    for (const target of graph.get(file) ?? []) visit(target);
    path.pop();
    visiting.delete(file);
    visited.add(file);
  }
  for (const file of graph.keys()) visit(file);
  return errors;
}

function sources(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((item) => {
    const file = `${directory}/${item.name}`;
    if (item.isDirectory()) return sources(file);
    return /\.tsx?$/.test(file) ? [[file, readFileSync(file, 'utf8')]] : [];
  });
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  process.chdir(resolve(dirname(fileURLToPath(import.meta.url)), '..'));
  const files = new Map([
    ...Object.keys(layers).flatMap((name) => sources(`packages/${name}/src`)),
    ...sources('apps/playground/src'),
  ]);
  const errors = analyzeArchitecture(files);
  if (errors.length) {
    console.error(errors.join('\n'));
    process.exitCode = 1;
  } else
    console.log(
      `PASS: ${files.size} source modules obey package boundaries and have no runtime cycles.`,
    );
}
