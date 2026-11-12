/** Yield to another task without retaining message ports after the callback. */
export function yieldToTask(): Promise<void> {
  if (typeof MessageChannel === 'undefined')
    return new Promise((resolve) => setTimeout(resolve, 0));
  return new Promise((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      channel.port2.close();
      resolve();
    };
    channel.port2.postMessage(0);
  });
}
