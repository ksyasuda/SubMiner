import net from 'node:net';
import {
  encodeAppControlRequest,
  getAppControlSocketPath,
  parseAppControlResponseLine,
  type AppControlSocketPathOptions,
} from './app-control';

export interface AppControlClientOptions extends AppControlSocketPathOptions {
  socketPath?: string;
  timeoutMs?: number;
}

export interface AppControlCommandResult {
  ok: boolean;
  unavailable?: boolean;
  error?: string;
}

function resolveSocketPath(options: AppControlClientOptions): string {
  return options.socketPath ?? getAppControlSocketPath(options);
}

export function isAppControlServerAvailable(
  options: AppControlClientOptions = {},
): Promise<boolean> {
  const socketPath = resolveSocketPath(options);
  const timeoutMs = options.timeoutMs ?? 350;

  return new Promise<boolean>((resolve) => {
    const socket = net.createConnection(socketPath);
    let settled = false;

    const finish = (available: boolean): void => {
      if (settled) return;
      settled = true;
      try {
        socket.destroy();
      } catch {
        // ignore
      }
      resolve(available);
    };

    socket.once('connect', () => finish(typeof socket.write === 'function'));
    socket.once('error', () => finish(false));
    socket.setTimeout(timeoutMs, () => finish(false));
  });
}

export function sendAppControlCommand(
  argv: string[],
  options: AppControlClientOptions = {},
): Promise<AppControlCommandResult> {
  const socketPath = resolveSocketPath(options);
  const timeoutMs = options.timeoutMs ?? 1000;

  return new Promise<AppControlCommandResult>((resolve) => {
    const socket = net.createConnection(socketPath);
    let settled = false;
    let connected = false;
    let responseBuffer = '';

    const finish = (result: AppControlCommandResult): void => {
      if (settled) return;
      settled = true;
      try {
        socket.destroy();
      } catch {
        // ignore
      }
      resolve(result);
    };

    socket.once('connect', () => {
      connected = true;
      if (typeof socket.write !== 'function') {
        finish({ ok: false, unavailable: true, error: 'App control socket is not writable' });
        return;
      }
      socket.write(encodeAppControlRequest(argv));
    });
    socket.on('data', (chunk) => {
      responseBuffer += chunk.toString('utf8');
      const newlineIndex = responseBuffer.indexOf('\n');
      if (newlineIndex < 0) return;
      try {
        finish(parseAppControlResponseLine(responseBuffer.slice(0, newlineIndex)));
      } catch (error) {
        finish({ ok: false, error: error instanceof Error ? error.message : String(error) });
      }
    });
    socket.once('error', (error) => {
      finish({ ok: false, unavailable: !connected, error: error.message });
    });
    socket.once('close', () => {
      finish({ ok: false, unavailable: !connected, error: 'App control socket closed' });
    });
    socket.setTimeout(timeoutMs, () => {
      finish({ ok: false, unavailable: !connected, error: 'App control socket timed out' });
    });
  });
}
