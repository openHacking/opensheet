/** Explicit external-link profile used for persisted metadata and user navigation. */
export function isSafeLinkTarget(target: string): boolean {
  if (!target || target.length > 4096 || /[\u0000-\u0020\u007f\\]/.test(target)) return false;
  try {
    const url = new URL(target);
    if (url.protocol === 'mailto:') return /^mailto:/i.test(target) && url.pathname.length > 0;
    return (
      /^https?:\/\//i.test(target) &&
      (url.protocol === 'http:' || url.protocol === 'https:') &&
      !!url.hostname
    );
  } catch {
    return false;
  }
}
