export type WorkbookDialogFeedback = {
  modal(title: string, content: HTMLElement): void;
  notice(message: string): void;
  isCurrent(): boolean;
  isTargetCurrent?(): boolean;
};

export function currentWorkbookDialog(
  form: HTMLElement,
  status: HTMLElement,
  isCurrent: () => boolean,
) {
  if (!form.closest('dialog')?.open) return false;
  if (isCurrent()) return true;
  status.textContent = 'Workbook changed. Check the previous operation before trying again.';
  return false;
}

export class WorkbookDialog {
  readonly element = document.createElement('dialog');
  private title = document.createElement('h2');
  private content = document.createElement('div');
  private close = document.createElement('button');

  constructor(container: HTMLElement) {
    this.element.className = 'os-workbook-dialog';
    this.title.id = `os-dialog-${crypto.randomUUID()}`;
    this.element.setAttribute('aria-labelledby', this.title.id);
    this.close.type = 'button';
    this.close.setAttribute('aria-label', 'Close dialog');
    this.close.textContent = '×';
    this.close.onclick = () => this.element.close();
    const header = document.createElement('header');
    header.append(this.title, this.close);
    this.element.append(header, this.content);
    container.append(this.element);
  }

  show(title: string, content: HTMLElement) {
    this.title.textContent = title;
    this.content.replaceChildren(content);
    this.element.scrollTop = 0;
    if (!this.element.open) this.element.showModal();
  }
}
