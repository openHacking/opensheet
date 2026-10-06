import { createElement, useEffect, useRef, type ReactElement, type CSSProperties } from 'react';
import { createOpenSheet, type OpenSheet, type WorkbookSnapshot } from 'opensheet';
export interface OpenSheetProps {
  initialSnapshot?: WorkbookSnapshot;
  readOnly?: boolean;
  className?: string;
  style?: CSSProperties;
  onReady?: (app: OpenSheet) => void;
  onChange?: (snapshot: WorkbookSnapshot) => void;
}
/** initialSnapshot is read once. Use the instance API to load another workbook. */
export function OpenSheetView(props: OpenSheetProps): ReactElement {
  const host = useRef<HTMLDivElement>(null),
    app = useRef<OpenSheet | null>(null),
    callbacks = useRef(props);
  callbacks.current = props;
  useEffect(() => {
    const instance = createOpenSheet({
      container: host.current!,
      mode: callbacks.current.readOnly ? 'read' : 'edit',
    });
    app.current = instance;
    let active = true;
    const start = async () => {
      if (callbacks.current.initialSnapshot) await instance.load(callbacks.current.initialSnapshot);
      else instance.createWorkbook();
      if (!active) return;
      instance.on('workbook:committed', () =>
        callbacks.current.onChange?.(instance.getWorkbook().toJSON()),
      );
      callbacks.current.onReady?.(instance);
    };
    void start();
    return () => {
      active = false;
      instance.dispose();
      app.current = null;
    };
  }, []);
  useEffect(() => {
    app.current?.setMode(props.readOnly ? 'read' : 'edit');
  }, [props.readOnly]);
  return createElement('div', {
    ref: host,
    className: props.className,
    style: { height: 500, ...props.style },
  });
}
