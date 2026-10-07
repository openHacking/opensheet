import { createElement, type IconNode } from 'lucide';
export function button(
  label: string,
  run: () => void,
  title = label,
  icon?: IconNode,
  signal?: AbortSignal,
) {
  const b = document.createElement('button');
  b.type = 'button';
  if (icon) {
    const svg = createElement(icon, {
      width: 16,
      height: 16,
      'aria-hidden': 'true',
      focusable: 'false',
    });
    b.append(svg);
    if (label) b.append(document.createTextNode(label));
  } else b.textContent = label;
  b.title = title;
  b.setAttribute('aria-label', title);
  b.addEventListener('click', run, { signal });
  return b;
}
