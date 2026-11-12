import { OpenSheetError } from '@opensheetjs/core';
import type { AdapterOptions, CompatibilityReport, Issue } from './types.js';
export function report(): CompatibilityReport {
  return {
    adapterVersion: '0.2.0-dev.0',
    sourceFormat: 'SheetJS',
    summary: { exact: 0, approximated: 0, dropped: 0, blocked: 0 },
    issues: [],
    omittedIssueOccurrences: 0,
  };
}
export function add(r: CompatibilityReport, issue: Issue) {
  if (
    !issue ||
    !['preserved', 'approximated', 'dropped', 'blocked'].includes(issue.action) ||
    (issue.count !== undefined && (!Number.isSafeInteger(issue.count) || issue.count < 1))
  )
    throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid compatibility issue action or count');
  const field = issue.action === 'preserved' ? 'exact' : issue.action;
  const count = issue.count ?? 1;
  if (!Number.isSafeInteger(r.summary[field] + count))
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Compatibility occurrence count exceeds safe range');
  r.summary[field] += count;
  const prior = r.issues.find(
    (x) =>
      x.code === issue.code &&
      x.sheetId === issue.sheetId &&
      x.action === issue.action &&
      x.severity === issue.severity &&
      x.feature === issue.feature,
  );
  if (prior) prior.count = (prior.count ?? 1) + count;
  else if (r.issues.length < 200) r.issues.push({ ...issue, count });
  else {
    // Keep actionable blockers visible when informational or warning groups
    // filled the bounded detail list first.
    const replace =
      issue.action === 'blocked' ? r.issues.findIndex((entry) => entry.action !== 'blocked') : -1;
    const omitted = replace < 0 ? count : (r.issues[replace].count ?? 1);
    if (!Number.isSafeInteger((r.omittedIssueOccurrences ?? 0) + omitted))
      throw new OpenSheetError(
        'LIMIT_EXCEEDED',
        'Compatibility occurrence count exceeds safe range',
      );
    r.omittedIssueOccurrences = (r.omittedIssueOccurrences ?? 0) + omitted;
    if (replace >= 0) r.issues[replace] = { ...issue, count };
  }
}
export function finish(r: CompatibilityReport, o: AdapterOptions) {
  if (
    o.unsupported === 'strict' &&
    (r.summary.approximated > 0 || r.summary.dropped > 0 || r.summary.blocked > 0)
  )
    throw new OpenSheetError('UNSUPPORTED_FEATURE', 'Strict compatibility check failed', r);
}
