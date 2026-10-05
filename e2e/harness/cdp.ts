import { waitUntil, type WaitOptions } from './wait';

export type CdpTarget = {
  type: string;
  title: string;
  url: string;
  webSocketDebuggerUrl: string;
};

export type CdpPage = {
  /** Sends a raw DevTools protocol command, for anything the helpers below do not cover. */
  send: (method: string, params?: Record<string, unknown>) => Promise<Record<string, unknown>>;
  /** Evaluates an expression in the page, awaiting promises, and returns its JSON value. */
  evaluate: <T>(expression: string) => Promise<T>;
  /** Polls an expression until it is truthy and returns that value. */
  waitFor: <T>(expression: string, options: WaitOptions) => Promise<NonNullable<T>>;
  /** PNG of the page as rendered by Chromium; needs no OS screen-capture permission. */
  screenshot: () => Promise<Buffer>;
  close: () => void;
};

type CdpReply = {
  id?: number;
  error?: { message: string };
  result?: Record<string, unknown>;
};

type EvaluateResult = {
  result: { value?: unknown };
  exceptionDetails?: { text: string; exception?: { description?: string } };
};

export async function listCdpTargets(port: number): Promise<CdpTarget[]> {
  const response = await fetch(`http://127.0.0.1:${port}/json/list`);
  return (await response.json()) as CdpTarget[];
}

export async function connectCdpPage(target: CdpTarget): Promise<CdpPage> {
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise<void>((resolve, reject) => {
    socket.addEventListener('open', () => resolve(), { once: true });
    socket.addEventListener(
      'error',
      () => reject(new Error(`CDP connection to ${target.url} failed`)),
      { once: true },
    );
  });

  const pending = new Map<number, (reply: CdpReply) => void>();
  let nextId = 1;
  socket.addEventListener('message', (event) => {
    const reply: CdpReply = JSON.parse(String(event.data));
    if (reply.id === undefined) return;
    pending.get(reply.id)?.(reply);
    pending.delete(reply.id);
  });

  const send = (method: string, params: Record<string, unknown> = {}) =>
    new Promise<Record<string, unknown>>((resolve, reject) => {
      const id = nextId++;
      pending.set(id, (reply) => {
        if (reply.error) reject(new Error(`CDP ${method} failed: ${reply.error.message}`));
        else resolve(reply.result ?? {});
      });
      socket.send(JSON.stringify({ id, method, params }));
    });

  const evaluate = async <T>(expression: string): Promise<T> => {
    const reply = (await send('Runtime.evaluate', {
      expression,
      returnByValue: true,
      awaitPromise: true,
    })) as EvaluateResult;
    if (reply.exceptionDetails) {
      const { text, exception } = reply.exceptionDetails;
      throw new Error(`Page evaluation threw: ${exception?.description ?? text}`);
    }
    return reply.result.value as T;
  };

  return {
    send,
    evaluate,
    waitFor: <T>(expression: string, options: WaitOptions) =>
      waitUntil(() => evaluate<T>(expression), options),
    screenshot: async () => {
      const reply = await send('Page.captureScreenshot', { format: 'png' });
      return Buffer.from(String(reply.data), 'base64');
    },
    close: () => socket.close(),
  };
}
