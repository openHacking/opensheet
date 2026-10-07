import type { Workbook } from '@opensheetjs/core';
import { Plus, type IconNode } from 'lucide';
import { button } from './button.js';
export interface TabActions {
  book(): Workbook;
  activeSheet(): string;
  setSheet(id: string): void;
  run(action: () => unknown): void;
}
export class SheetTabs {
  readonly element = document.createElement('div');
  private abort = new AbortController();
  constructor(private actions: TabActions) {
    this.element.className = 'os-tabs';
    this.element.setAttribute('role', 'tablist');
    this.element.setAttribute('aria-label', 'Sheets');
  }
  private button(label: string, action: () => void, title = label, icon?: IconNode) {
    return button(label, () => this.actions.run(action), title, icon, this.abort.signal);
  }
  render() {
    this.abort.abort();
    this.abort = new AbortController();
    this.element.replaceChildren();
    for (const sheet of this.actions.book().getSheets()) {
      if (this.actions.book().sheetData(sheet.id).hidden) continue;
      const b = this.button(sheet.name, () => {
        this.actions.setSheet(sheet.id);
        this.render();
      });
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(sheet.id === this.actions.activeSheet()));
      this.element.append(b);
    }
    this.element.append(
      this.button(
        '',
        async () => {
          const names = new Set(
            this.actions
              .book()
              .getSheets()
              .map((s) => s.name),
          );
          let i = 1;
          while (names.has(`Sheet${i}`)) i++;
          const sh = await this.actions.book().addSheet(`Sheet${i}`);
          this.actions.setSheet(sh.id);
          this.render();
        },
        'Add sheet',
        Plus,
      ),
    );
  }
  dispose() {
    this.abort.abort();
  }
}
