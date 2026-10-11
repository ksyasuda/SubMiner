import { waitUntil, type Truthy, type WaitOptions } from './wait';

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
  waitFor: <T>(expression: string, options: WaitOptions) => Promise<Truthy<T>>;
  /** Console output and uncaught exceptions seen since connecting, oldest first. */
  console: string[];
  /** PNG of the page as rendered by Chromium; needs no OS screen-capture permission. */
  screenshot: () => Promise<Buffer>;
  close: () => void;
};

type CdpReply = {
  id?: number;
  method?: string;
  params?: Record<string, unknown>;
  error?: { message: string };
  result?: Record<string, unknown>;
};

type RemoteObject = { value?: unknown; description?: string };

// Renders a console/exception event as one log line.
function formatConsoleEvent(method: string, params: Record<string, unknown>): string | null {
  if (method === 'Runtime.consoleAPICalled') {
    const args = Array.isArray(params.args) ? (params.args as RemoteObject[]) : [];
    const text = args.map((arg) => arg.description ?? JSON.stringify(arg.value)).join(' ');
    return `console.${String(params.type)}: ${text}`;
  }
  if (method === 'Runtime.exceptionThrown') {
    const details = params.exceptionDetails as
      | { text?: string; exception?: RemoteObject }
      | undefined;
    return `exception: ${details?.exception?.description ?? details?.text ?? 'unknown'}`;
  }
  if (method === 'Log.entryAdded') {
    const entry = params.entry as { level?: string; source?: string; text?: string } | undefined;
    return `${String(entry?.source)}.${String(entry?.level)}: ${String(entry?.text)}`;
  }
  return null;
}

type EvaluateResult = {
  result: { value?: unknown };
  exceptionDetails?: { text: string; exception?: { description?: string } };
};

export async function listCdpTargets(port: number): Promise<CdpTarget[]> {
  const response = await fetch(`http://127.0.0.1:${port}/json/list`);
  return (await response.json()) as CdpTarget[];
}

export function connectCdpPage(target: CdpTarget): Promise<CdpPage> {
  return connectCdp(target.webSocketDebuggerUrl, target.url);
}

/** Connects to any DevTools endpoint: a renderer page or Electron's main-process inspector. */
export async function connectCdp(webSocketDebuggerUrl: string, label: string): Promise<CdpPage> {
  const socket = new WebSocket(webSocketDebuggerUrl);
  await new Promise<void>((resolve, reject) => {
    socket.addEventListener('open', () => resolve(), { once: true });
    socket.addEventListener('error', () => reject(new Error(`CDP connection to ${label} failed`)), {
      once: true,
    });
  });

  const pending = new Map<number, (reply: CdpReply) => void>();
  const consoleLines: string[] = [];
  let nextId = 1;
  socket.addEventListener('message', (event) => {
    const reply: CdpReply = JSON.parse(String(event.data));
    if (reply.id === undefined) {
      if (reply.method) {
        const line = formatConsoleEvent(reply.method, reply.params ?? {});
        if (line) consoleLines.push(line);
      }
      return;
    }
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

  await send('Runtime.enable');
  // The Node inspector behind Electron's main process has no Log domain.
  await send('Log.enable').catch(() => undefined);

  return {
    send,
    evaluate,
    console: consoleLines,
    waitFor: <T>(expression: string, options: WaitOptions) =>
      waitUntil(() => evaluate<T>(expression), options),
    screenshot: async () => {
      const reply = await send('Page.captureScreenshot', { format: 'png' });
      return Buffer.from(String(reply.data), 'base64');
    },
    close: () => socket.close(),
  };
}
