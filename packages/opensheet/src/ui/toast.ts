export class Toast {
  readonly element = document.createElement('div');
  private timer?: ReturnType<typeof setTimeout>;
  constructor() {
    this.element.className = 'os-toast';
    this.element.hidden = true;
    this.element.setAttribute('role', 'status');
  }
  show(message: string) {
    this.element.textContent = message;
    this.element.hidden = false;
    clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      this.element.hidden = true;
    }, 5000);
  }
  dispose() {
    clearTimeout(this.timer);
  }
}
