export class GridScrollbars {
  private tracks: { element: HTMLDivElement; thumb: HTMLDivElement; horizontal: boolean }[];
  constructor(
    container: HTMLElement,
    private scroller: HTMLElement,
    signal: AbortSignal,
  ) {
    this.tracks = [true, false].map((horizontal) => {
      const element = document.createElement('div');
      const thumb = document.createElement('div');
      element.className = `os-scrollbar os-scrollbar-${horizontal ? 'horizontal' : 'vertical'}`;
      thumb.className = 'os-scrollbar-thumb';
      element.tabIndex = 0;
      element.setAttribute('role', 'scrollbar');
      element.setAttribute('aria-label', horizontal ? 'Horizontal scroll' : 'Vertical scroll');
      element.setAttribute('aria-orientation', horizontal ? 'horizontal' : 'vertical');
      element.setAttribute('aria-controls', scroller.id);
      element.append(thumb);
      container.append(element);
      const metrics = () => {
        const extent = horizontal ? scroller.clientWidth : scroller.clientHeight;
        const total = horizontal ? scroller.scrollWidth : scroller.scrollHeight;
        const length = horizontal ? element.clientWidth : element.clientHeight;
        const size = Math.min(length, Math.max(24, (length * extent) / total));
        return { extent, max: Math.max(0, total - extent), travel: length - size };
      };
      const get = () => (horizontal ? scroller.scrollLeft : scroller.scrollTop);
      const set = (value: number) => {
        if (horizontal) scroller.scrollLeft = value;
        else scroller.scrollTop = value;
        this.update();
      };
      let drag: { pointer: number; coordinate: number; offset: number } | undefined;
      element.addEventListener(
        'pointerdown',
        (event) => {
          event.stopPropagation();
          event.preventDefault();
          if (event.button !== 0) return;
          element.focus({ preventScroll: true });
          const coordinate = horizontal ? event.clientX : event.clientY;
          const rect = thumb.getBoundingClientRect();
          const start = horizontal ? rect.left : rect.top;
          const end = horizontal ? rect.right : rect.bottom;
          if (coordinate < start || coordinate > end) {
            set(get() + (coordinate < start ? -1 : 1) * metrics().extent);
          } else {
            drag = { pointer: event.pointerId, coordinate, offset: get() };
            element.setPointerCapture(event.pointerId);
          }
        },
        { signal },
      );
      element.addEventListener(
        'pointermove',
        (event) => {
          if (!drag || event.pointerId !== drag.pointer) return;
          const { max, travel } = metrics();
          if (travel > 0)
            set(
              drag.offset +
                (((horizontal ? event.clientX : event.clientY) - drag.coordinate) * max) / travel,
            );
        },
        { signal },
      );
      element.addEventListener(
        'lostpointercapture',
        () => {
          drag = undefined;
        },
        { signal },
      );
      element.addEventListener('dblclick', (event) => event.stopPropagation(), { signal });
      element.addEventListener(
        'keydown',
        (event) => {
          event.stopPropagation();
          const { extent, max } = metrics();
          const steps: Record<string, number> = {
            ArrowLeft: -40,
            ArrowRight: 40,
            ArrowUp: -40,
            ArrowDown: 40,
            PageUp: -extent,
            PageDown: extent,
          };
          if (event.key === 'Home') set(0);
          else if (event.key === 'End') set(max);
          else if (event.key in steps) set(get() + steps[event.key]);
          else return;
          event.preventDefault();
        },
        { signal },
      );
      return { element, thumb, horizontal };
    });
  }
  update() {
    for (const { element, thumb, horizontal } of this.tracks) {
      const extent = horizontal ? this.scroller.clientWidth : this.scroller.clientHeight;
      const total = horizontal ? this.scroller.scrollWidth : this.scroller.scrollHeight;
      const offset = horizontal ? this.scroller.scrollLeft : this.scroller.scrollTop;
      const length = horizontal ? element.clientWidth : element.clientHeight;
      const size = Math.min(length, Math.max(24, (length * extent) / total));
      const max = Math.max(0, total - extent);
      const position = max ? (offset / max) * (length - size) : 0;
      thumb.style[horizontal ? 'width' : 'height'] = `${Math.max(0, size - 4)}px`;
      thumb.style[horizontal ? 'left' : 'top'] = `${position + 2}px`;
      element.setAttribute('aria-valuemin', '0');
      element.setAttribute('aria-valuemax', String(max));
      element.setAttribute('aria-valuenow', String(offset));
      element.setAttribute('aria-disabled', String(max === 0));
    }
  }
}
