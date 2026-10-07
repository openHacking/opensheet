import { axisCommand } from './commands/axis.js';
import { cellsCommand } from './commands/cells.js';
import { sheetsCommand } from './commands/sheets.js';
import { assert, type Command, type WorkbookSnapshot } from './types.js';
export { commandSchemas, validateCommand } from './commands/schemas.js';
export function reduceCommand(state: WorkbookSnapshot, command: Command): void {
  if (command.type.startsWith('core.cells.')) cellsCommand(state, command);
  else if (command.type.startsWith('core.axis.')) axisCommand(state, command);
  else if (command.type.startsWith('core.sheet.')) sheetsCommand(state, command);
  else assert(false, 'UNSUPPORTED_FEATURE', 'Unknown command');
}
