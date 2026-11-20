export type MenuItem = {
  label: string;
  disabled?: boolean;
  action?: () => unknown;
  items?: MenuItem[];
};

export class ActionMenu {
  private element = document.createElement('div');
  private abort = new AbortController();
  private trigger?: HTMLElement;
  private fallback?: () => void;
  private stack: MenuItem[][] = [];
  private point = { x: 0, y: 0 };

  constructor(private run: (action: () => unknown) => void) {
    const signal = this.abort.signal;
    this.element.className = 'os-action-menu';
    this.element.setAttribute('role', 'menu');
    this.element.hidden = true;
    this.element.tabIndex = -1;
    this.element.addEventListener('keydown', (event) => this.key(event), { signal });
  }
  private outside?: AbortController;

  private watchOutside() {
    this.outside = new AbortController();
    const signal = this.outside.signal;
    const dismiss = (event: Event) => {
      if (!this.element.contains(event.target as Node)) this.close(false);
    };
    document.addEventListener('pointerdown', dismiss, { signal });
    document.addEventListener('wheel', dismiss, { signal, capture: true });
    window.addEventListener('resize', () => this.close(false), { signal });
  }

  open(
    label: string,
    items: MenuItem[],
    point: { x: number; y: number },
    fallback: () => void,
    trigger?: HTMLElement,
  ) {
    this.close(false);
    if (!items.length) return;
    if (!this.element.isConnected) document.body.append(this.element);
    this.element.setAttribute('aria-label', label);
    this.trigger = trigger;
    this.trigger?.setAttribute('aria-expanded', 'true');
    this.fallback = fallback;
    this.point = point;
    this.stack = [items];
    this.watchOutside();
    this.render();
  }

  close(restore = true) {
    if (this.element.hidden) return;
    this.element.hidden = true;
    this.outside?.abort();
    this.outside = undefined;
    this.trigger?.setAttribute('aria-expanded', 'false');
    if (restore) {
      if (this.trigger?.isConnected && this.trigger.getClientRects().length)
        this.trigger.focus({ preventScroll: true });
      else this.fallback?.();
    }
    this.trigger = undefined;
    this.fallback = undefined;
    this.stack = [];
  }

  private render() {
    const items = this.stack.at(-1)!;
    this.element.replaceChildren();
    this.element.hidden = false;
    const add = (item: MenuItem) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.tabIndex = -1;
      button.setAttribute('role', 'menuitem');
      button.disabled = !!item.disabled;
      button.textContent = item.label;
      if (item.items) {
        button.setAttribute('aria-haspopup', 'menu');
        const indicator = document.createElement('span');
        indicator.setAttribute('aria-hidden', 'true');
        indicator.textContent = '›';
        button.append(indicator);
      }
      button.onclick = () => {
        if (item.items) {
          this.element.dataset.parentIndex = String(this.buttons().indexOf(button));
          this.stack.push(item.items);
          this.render();
        } else if (item.action) {
          this.close();
          this.run(item.action);
        }
      };
      this.element.append(button);
    };
    if (this.stack.length > 1) {
      add({ label: 'Back' });
      this.element.firstElementChild!.addEventListener('click', () => this.back());
    }
    items.forEach(add);
    const margin = 6;
    const clamp = (coordinate: number, size: number, limit: number) =>
      Math.max(margin, Math.min(coordinate, limit - size - margin));
    this.element.style.left = `${clamp(this.point.x, this.element.offsetWidth, window.innerWidth)}px`;
    this.element.style.top = `${clamp(this.point.y, this.element.offsetHeight, window.innerHeight)}px`;
    const buttons = this.buttons();
    (buttons[this.stack.length > 1 && buttons.length > 1 ? 1 : 0] ?? this.element).focus({
      preventScroll: true,
    });
  }

  private buttons() {
    return [...this.element.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
  }
  private back() {
    if (this.stack.length === 1) return;
    this.stack.pop();
    const index = this.element.dataset.parentIndex;
    this.render();
    if (index !== undefined) this.buttons()[Number(index)]?.focus({ preventScroll: true });
  }
  private key(event: KeyboardEvent) {
    const buttons = this.buttons();
    const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
    let next: number;
    switch (event.key) {
      case 'ArrowDown':
        next = (index + 1) % buttons.length;
        break;
      case 'ArrowUp':
        next = (index <= 0 ? buttons.length : index) - 1;
        break;
      case 'Home':
        next = 0;
        break;
      case 'End':
        next = buttons.length - 1;
        break;
      case 'ArrowRight':
        if (buttons[index]?.hasAttribute('aria-haspopup')) {
          buttons[index].click();
        }
        next = -1;
        break;
      case 'ArrowLeft':
        this.back();
        next = -1;
        break;
      case 'Escape':
        if (this.stack.length > 1) this.back();
        else this.close();
        next = -1;
        break;
      case 'Tab':
        this.close();
        return;
      default:
        return;
    }
    event.preventDefault();
    if (next >= 0) buttons[next]?.focus();
  }
  dispose() {
    this.close(false);
    this.abort.abort();
    this.element.remove();
  }
}
