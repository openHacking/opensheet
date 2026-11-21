import { createElement, useEffect, useRef, type ReactElement, type CSSProperties } from 'react';
import { createOpenSheet, type OpenSheet, type Commit, type WorkbookFile } from 'opensheet';
export interface OpenSheetProps {
  initialFile?: WorkbookFile;
  readOnly?: boolean;
  className?: string;
  style?: CSSProperties;
  onError?: (error: unknown) => void;
  onReady?: (app: OpenSheet) => void;
  onChange?: (commit: Commit) => void;
}
/** initialFile is read once. Use the instance API to load another workbook. */
export function OpenSheetView(props: OpenSheetProps): ReactElement {
  const host = useRef<HTMLDivElement>(null),
    app = useRef<OpenSheet | null>(null),
    callbacks = useRef(props);
  callbacks.current = props;
  useEffect(() => {
    const reportInitializationError = (error: unknown) => {
      const current = callbacks.current;
      const handler = current.onError;
      if (!handler) {
        console.error(error);
        return;
      }
      try {
        const result: unknown = handler.call(current, error);
        if (result && typeof (result as PromiseLike<unknown>).then === 'function')
          void Promise.resolve(result).catch((observerError) => console.error(observerError));
      } catch (observerError) {
        console.error(observerError);
      }
    };
    let instance: OpenSheet;
    try {
      instance = createOpenSheet({
        container: host.current!,
        mode: callbacks.current.readOnly ? 'read' : 'edit',
        onError: (error) => callbacks.current.onError?.(error),
      });
    } catch (error) {
      reportInitializationError(error);
      return;
    }
    app.current = instance;
    let active = true;
    const start = async () => {
      if (callbacks.current.initialFile) await instance.load(callbacks.current.initialFile);
      else await instance.createWorkbook();
      if (!active) return;
      instance.on('workbook:committed', (commit) => callbacks.current.onChange?.(commit));
      callbacks.current.onReady?.(instance);
    };
    void start().catch((error) => {
      if (active) reportInitializationError(error);
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
    className: ['os-framework-host', props.className].filter(Boolean).join(' '),
    style: props.style,
  });
}
