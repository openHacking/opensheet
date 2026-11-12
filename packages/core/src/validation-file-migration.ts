import {
  validateWorkbookFile,
  validateWorkbookFilePages,
  type FileBlock,
  type WorkbookFile,
} from './file-codec.js';
import { validateJSONValue } from './model.js';
import { assert } from './types.js';
import { validateValidationEntries, type ValidationEntry } from './validation-model.js';
import { verificationYield } from './verification-yield.js';
import { validateConditionalRules, type ConditionalRule } from './conditional-format.js';

/** Provisional in-memory migration building block; not a supported public file format. */
type ValidationSheet = WorkbookFile['sheets'][number] & { validation: ValidationEntry[] };
export type ValidationWorkbookFile = Omit<WorkbookFile, 'schemaVersion' | 'sheets'> & {
  schemaVersion: 4;
  sheets: ValidationSheet[];
};
type FileMetadata<T extends WorkbookFile | ValidationWorkbookFile | ConditionalWorkbookFile> = Omit<
  T,
  'sheets'
> & {
  sheets: Omit<T['sheets'][number], 'blocks'>[];
};
type StagePage = (sheetIndex: number, block: FileBlock, signal?: AbortSignal) => Promise<void>;

/** Stage validated pages in order, returning publication metadata only on success.
 * The caller owns the input for the duration and must discard staging on failure.
 * The sink must not mutate pages or publish anything itself. It must honor the signal
 * and settle only after its page operation has stopped, so caller cleanup cannot race it.
 */
export async function stageValidationUpgrade(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
): Promise<FileMetadata<ValidationWorkbookFile>> {
  const file = await stagePages(input, stage, signal);
  return {
    ...file,
    schemaVersion: 4,
    sheets: file.sheets.map((sheet) => ({ ...sheet, validation: [] })),
  };
}

async function stagePages(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
  onMetadata?: (file: WorkbookFile) => void,
): Promise<FileMetadata<WorkbookFile>> {
  const checkpoint = verificationYield(signal);
  signal?.throwIfAborted();
  const pages = validateWorkbookFilePages(input, onMetadata);
  try {
    for (let step = pages.next(); ; step = pages.next()) {
      if (step.done) {
        return structuredClone({
          ...step.value,
          sheets: step.value.sheets.map(({ blocks: _blocks, ...sheet }) => sheet),
        });
      }
      await checkpoint();
      await stage(step.value.sheetIndex, step.value.block, signal);
      signal?.throwIfAborted();
    }
  } finally {
    pages.return(undefined as never);
  }
}

function splitV4(input: unknown) {
  assert(
    input && typeof input === 'object' && !Array.isArray(input),
    'INVALID_ARGUMENT',
    'Invalid validation workbook',
  );
  const candidate = input as Record<string, any>;
  assert(
    candidate.schemaVersion === 4 && Array.isArray(candidate.sheets),
    'INVALID_ARGUMENT',
    'Validation workbook schema 4 required',
  );
  const validation: unknown[] = [];
  const sheets = candidate.sheets.map((sheet: unknown) => {
    assert(
      sheet &&
        typeof sheet === 'object' &&
        !Array.isArray(sheet) &&
        Object.hasOwn(sheet, 'validation'),
      'INVALID_ARGUMENT',
      'Missing validation rules',
    );
    const { validation: rules, ...rest } = sheet as Record<string, unknown>;
    validateJSONValue(rules);
    validation.push(rules);
    return rest;
  });
  return { legacy: { ...candidate, schemaVersion: 3, sheets }, validation };
}

function validatedRules(legacy: WorkbookFile, validation: unknown[]) {
  return legacy.sheets.map((sheet, index) => {
    // The legacy validator already checked every axis run and its identities.
    return validateValidationEntries(
      validation[index],
      sheet.rows.order.reduce((length, [, count]) => length + count, 0),
      sheet.columns.order.reduce((length, [, count]) => length + count, 0),
    );
  });
}

function rejectRules(legacy: WorkbookFile, validation: unknown[]) {
  const rules = validatedRules(legacy, validation);
  assert(
    rules.every((entries) => entries.length === 0),
    'UNSUPPORTED_FEATURE',
    'Validation rules cannot be represented in schema 3',
  );
}

/** Stage a provisional v4 backup without discarding its rules. No head is published here. */
export async function stageValidationFile(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
): Promise<FileMetadata<ValidationWorkbookFile>> {
  signal?.throwIfAborted();
  const { legacy, validation } = splitV4(input);
  let rules: ValidationEntry[][] | undefined;
  const metadata = await stagePages(legacy, stage, signal, (file) => {
    rules = validatedRules(file, validation);
  });
  assert(rules, 'INVALID_ARGUMENT', 'Missing validation metadata');
  return {
    ...metadata,
    schemaVersion: 4,
    sheets: metadata.sheets.map((sheet, index) => ({ ...sheet, validation: rules![index] })),
  };
}

/** The caller must separately back up and verify durable storage before any upgrade. */
export function upgradeValidationFile(input: unknown): ValidationWorkbookFile {
  const file = structuredClone(validateWorkbookFile(input));
  return {
    ...file,
    schemaVersion: 4,
    sheets: file.sheets.map((sheet) => ({ ...sheet, validation: [] })),
  };
}

/** Downgrading must never silently remove active validation semantics. */
export function downgradeValidationFile(input: unknown): WorkbookFile {
  const { legacy, validation } = splitV4(input);
  const file = validateWorkbookFile(legacy);
  rejectRules(file, validation);
  return structuredClone(file);
}

/** Streaming parser metadata contains empty page arrays; the importer checks staged pages. */
export function downgradeValidationMetadata(input: unknown): WorkbookFile {
  const { legacy, validation } = splitV4(input);
  assert(
    legacy.sheets.every((sheet) => Array.isArray(sheet.blocks) && sheet.blocks.length === 0),
    'INVALID_ARGUMENT',
    'Migration metadata must not contain cell pages',
  );
  const file = validateWorkbookFile(legacy);
  rejectRules(file, validation);
  return file;
}

/** Refuse lossy rollback before calling the page sink. Caller owns staging cleanup. */
export async function stageValidationDowngrade(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
) {
  signal?.throwIfAborted();
  const { legacy, validation } = splitV4(input);
  return stagePages(legacy, stage, signal, (file) => rejectRules(file, validation));
}

// Experimental schema 5 keeps the provisional schema 4 validation rules and
// adds conditional formatting. Neither version is a public persistence format.
type ConditionalSheet = ValidationSheet & { conditionalFormats: ConditionalRule[] };
export type ConditionalWorkbookFile = Omit<WorkbookFile, 'schemaVersion' | 'sheets'> & {
  schemaVersion: 5;
  sheets: ConditionalSheet[];
};

function splitV5(input: unknown) {
  assert(
    input && typeof input === 'object' && !Array.isArray(input),
    'INVALID_ARGUMENT',
    'Invalid conditional workbook',
  );
  const candidate = input as Record<string, any>;
  assert(
    candidate.schemaVersion === 5 && Array.isArray(candidate.sheets),
    'INVALID_ARGUMENT',
    'Conditional workbook schema 5 required',
  );
  const conditional: unknown[] = [];
  const sheets = candidate.sheets.map((sheet: unknown) => {
    assert(
      sheet &&
        typeof sheet === 'object' &&
        !Array.isArray(sheet) &&
        Object.hasOwn(sheet, 'conditionalFormats'),
      'INVALID_ARGUMENT',
      'Missing conditional formatting rules',
    );
    const { conditionalFormats, ...rest } = sheet as Record<string, unknown>;
    validateJSONValue(conditionalFormats);
    conditional.push(conditionalFormats);
    return rest;
  });
  return { ...splitV4({ ...candidate, schemaVersion: 4, sheets }), conditional };
}

function validatedConditionalRules(file: WorkbookFile, conditional: unknown[]) {
  return file.sheets.map((sheet, index) =>
    validateConditionalRules(
      conditional[index],
      sheet.rows.order.reduce((length, [, count]) => length + count, 0),
      sheet.columns.order.reduce((length, [, count]) => length + count, 0),
    ),
  );
}

function rejectConditionalRules(rules: ConditionalRule[][]) {
  assert(
    rules.every((entries) => entries.length === 0),
    'UNSUPPORTED_FEATURE',
    'Conditional formatting cannot be represented in schema 4',
  );
}

/** Upgrade an already supported provisional schema 4 object, preserving validation rules. */
export function upgradeConditionalFile(input: unknown): ConditionalWorkbookFile {
  const { legacy, validation } = splitV4(input);
  const file = validateWorkbookFile(legacy);
  const rules = validatedRules(file, validation);
  return structuredClone({
    ...file,
    schemaVersion: 5 as const,
    sheets: file.sheets.map((sheet, index) => ({
      ...sheet,
      validation: rules[index],
      conditionalFormats: [],
    })),
  });
}

/** Stage a rule-free conditional upgrade without materializing another page table. */
export async function stageConditionalUpgrade(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
): Promise<FileMetadata<ConditionalWorkbookFile>> {
  signal?.throwIfAborted();
  const { legacy, validation } = splitV4(input);
  let validationRules: ValidationEntry[][] | undefined;
  const metadata = await stagePages(legacy, stage, signal, (file) => {
    validationRules = validatedRules(file, validation);
  });
  assert(validationRules, 'INVALID_ARGUMENT', 'Missing validation metadata');
  return {
    ...metadata,
    schemaVersion: 5,
    sheets: metadata.sheets.map((sheet, index) => ({
      ...sheet,
      validation: validationRules![index],
      conditionalFormats: [],
    })),
  };
}

/** Reject loss of formatting semantics and preserve schema 4 validation rules. */
export function downgradeConditionalFile(input: unknown): ValidationWorkbookFile {
  const { legacy, validation, conditional } = splitV5(input);
  const file = validateWorkbookFile(legacy);
  const rules = validatedRules(file, validation);
  rejectConditionalRules(validatedConditionalRules(file, conditional));
  return structuredClone({
    ...file,
    schemaVersion: 4 as const,
    sheets: file.sheets.map((sheet, index) => ({ ...sheet, validation: rules[index] })),
  });
}

/** Validate both rule sets before staging the first page; never publishes a head. */
export async function stageConditionalFile(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
): Promise<FileMetadata<ConditionalWorkbookFile>> {
  signal?.throwIfAborted();
  const { legacy, validation, conditional } = splitV5(input);
  let validationRules: ValidationEntry[][] | undefined;
  let conditionalRules: ConditionalRule[][] | undefined;
  const metadata = await stagePages(legacy, stage, signal, (file) => {
    validationRules = validatedRules(file, validation);
    conditionalRules = validatedConditionalRules(file, conditional);
  });
  assert(validationRules && conditionalRules, 'INVALID_ARGUMENT', 'Missing conditional metadata');
  return {
    ...metadata,
    schemaVersion: 5,
    sheets: metadata.sheets.map((sheet, index) => ({
      ...sheet,
      validation: validationRules![index],
      conditionalFormats: conditionalRules![index],
    })),
  };
}

/** A downgrade must reject active formatting rules before staging any pages. */
export async function stageConditionalDowngrade(
  input: unknown,
  stage: StagePage,
  signal?: AbortSignal,
): Promise<FileMetadata<ValidationWorkbookFile>> {
  signal?.throwIfAborted();
  const { legacy, validation, conditional } = splitV5(input);
  let validationRules: ValidationEntry[][] | undefined;
  const metadata = await stagePages(legacy, stage, signal, (file) => {
    validationRules = validatedRules(file, validation);
    rejectConditionalRules(validatedConditionalRules(file, conditional));
  });
  assert(validationRules, 'INVALID_ARGUMENT', 'Missing validation metadata');
  return {
    ...metadata,
    schemaVersion: 4,
    sheets: metadata.sheets.map((sheet, index) => ({
      ...sheet,
      validation: validationRules![index],
    })),
  };
}
