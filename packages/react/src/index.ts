import { createElement, useEffect, useRef, type ReactElement, type CSSProperties } from 'react';
import { createOpenSheet, type OpenSheet, type Commit, type WorkbookSnapshot } from 'opensheet';
export interface OpenSheetProps {
  initialSnapshot?: WorkbookSnapshot;
  readOnly?: boolean;
  className?: string;
  style?: CSSProperties;
  onError?: (error: unknown) => void;
  onReady?: (app: OpenSheet) => void;
  onChange?: (commit: Commit) => void;
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
      onError: (error) => callbacks.current.onError?.(error),
    });
    app.current = instance;
    let active = true;
    const start = async () => {
      if (callbacks.current.initialSnapshot) await instance.load(callbacks.current.initialSnapshot);
      else await instance.createWorkbook();
      if (!active) return;
      instance.on('workbook:committed', (commit) => callbacks.current.onChange?.(commit));
      callbacks.current.onReady?.(instance);
    };
    void start().catch((error) => {
      if (active) {
        if (callbacks.current.onError) callbacks.current.onError(error);
        else console.error(error);
      }
    });
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
