import fs from 'node:fs';
import net from 'node:net';
import path from 'node:path';
import {
  encodeAppControlResponse,
  parseAppControlRequestLine,
  type AppControlResponse,
} from '../../shared/app-control';

export interface AppControlServerOptions {
  socketPath: string;
  platform?: NodeJS.Platform;
  handleArgv: (argv: string[]) => void;
  logDebug?: (message: string) => void;
  logWarn?: (message: string, error?: unknown) => void;
}

export interface AppControlServerHandle {
  close: () => void;
}

function prepareSocketPath(socketPath: string, platform: NodeJS.Platform): void {
  if (platform === 'win32') return;
  fs.mkdirSync(path.dirname(socketPath), { recursive: true });
  fs.rmSync(socketPath, { force: true });
}

function cleanupSocketPath(socketPath: string, platform: NodeJS.Platform): void {
  if (platform === 'win32') return;
  try {
    fs.rmSync(socketPath, { force: true });
  } catch {
    // ignore
  }
}

function writeResponse(socket: net.Socket, response: AppControlResponse): void {
  socket.end(encodeAppControlResponse(response));
}

export function startAppControlServer(options: AppControlServerOptions): AppControlServerHandle {
  const platform = options.platform ?? process.platform;
  prepareSocketPath(options.socketPath, platform);

  const server = net.createServer((socket) => {
    let buffer = '';
    let handled = false;

    socket.on('data', (chunk) => {
      if (handled) return;
      buffer += chunk.toString('utf8');
      if (buffer.length > 65536) {
        handled = true;
        writeResponse(socket, { ok: false, error: 'App control request too large' });
        return;
      }

      const newlineIndex = buffer.indexOf('\n');
      if (newlineIndex < 0) return;
      handled = true;

      try {
        const request = parseAppControlRequestLine(buffer.slice(0, newlineIndex));
        options.handleArgv(request.argv);
        writeResponse(socket, { ok: true });
      } catch (error) {
        options.logWarn?.('Failed to handle app control command.', error);
        writeResponse(socket, {
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    });
  });

  server.on('error', (error) => {
    options.logWarn?.(`App control socket failed: ${options.socketPath}`, error);
  });
  server.listen(options.socketPath, () => {
    options.logDebug?.(`App control socket listening: ${options.socketPath}`);
  });

  let closed = false;
  return {
    close: () => {
      if (closed) return;
      closed = true;
      try {
        server.close();
      } catch {
        // ignore
      }
      cleanupSocketPath(options.socketPath, platform);
    },
  };
}
