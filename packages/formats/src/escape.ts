export const escapeHTML = (text: string) =>
  text.replace(
    /[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!,
  );
export const escapeLatex = (text: string) =>
  text
    .replace(
      /[&%$#_{}~^\\]/g,
      (c) =>
        ({
          '&': '\\&',
          '%': '\\%',
          $: '\\$',
          '#': '\\#',
          _: '\\_',
          '{': '\\{',
          '}': '\\}',
          '~': '\\textasciitilde{}',
          '^': '\\textasciicircum{}',
          '\\': '\\textbackslash{}',
        })[c]!,
    )
    .replace(/\r?\n/g, ' ');
