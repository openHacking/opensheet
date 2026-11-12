import { expect, it, vi } from 'vitest';
import { EditorEvents } from '../packages/opensheet/src/events.js';

it('isolates a throwing error observer and continues event delivery', () => {
  const log = vi.spyOn(console, 'error').mockImplementation(() => {});
  try {
    const onError = vi.fn(() => {
      throw new Error('error observer failed');
    });
    const events = new EditorEvents(onError);
    const later = vi.fn();
    events.on('editor:draft', () => {
      throw new Error('event observer failed');
    });
    events.on('editor:draft', later);
    events.emit('editor:draft', true);
    expect(onError).toHaveBeenCalledOnce();
    expect(later).toHaveBeenCalledWith(true);
    expect(log).toHaveBeenCalledOnce();
  } finally {
    log.mockRestore();
  }
});

it('observes asynchronous event and error-handler rejections', async () => {
  const log = vi.spyOn(console, 'error').mockImplementation(() => {});
  try {
    const onError = vi.fn(async () => {
      throw new Error('async error observer');
    });
    const events = new EditorEvents(onError);
    const later = vi.fn();
    events.on('editor:draft', async () => {
      throw new Error('async event observer');
    });
    events.on('editor:draft', later);
    events.emit('editor:draft', true);
    expect(later).toHaveBeenCalledWith(true);
    await vi.waitFor(() => expect(log).toHaveBeenCalledOnce());
    expect(onError).toHaveBeenCalledOnce();
    expect((onError.mock.calls[0] as unknown as [Error])[0].message).toBe('async event observer');
  } finally {
    log.mockRestore();
  }
});
