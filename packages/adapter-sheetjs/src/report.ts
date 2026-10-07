import { OpenSheetError } from '@opensheetjs/core';
import type { AdapterOptions, CompatibilityReport, Issue } from './types.js';
export function report(): CompatibilityReport {
  return {
    adapterVersion: '0.1.1',
    sourceFormat: 'SheetJS',
    summary: { exact: 0, approximated: 0, dropped: 0, blocked: 0 },
    issues: [],
  };
}
export function add(r: CompatibilityReport, issue: Issue) {
  const field = issue.action === 'preserved' ? 'exact' : issue.action;
  r.summary[field]++;
  const prior = r.issues.find((x) => x.code === issue.code && x.sheetId === issue.sheetId);
  if (prior) prior.count = (prior.count ?? 1) + 1;
  else if (r.issues.length < 200) r.issues.push(issue);
}
export function finish(r: CompatibilityReport, o: AdapterOptions) {
  if (o.unsupported === 'strict' && r.issues.some((i) => i.action !== 'preserved'))
    throw new OpenSheetError('UNSUPPORTED_FEATURE', 'Strict compatibility check failed', r);
}
