import { spawn, type ChildProcess } from 'node:child_process';
import path from 'node:path';
import { createServer } from 'node:net';
import { AnimeBridgeClient } from './bridge-client';
import type { BundleBinaries } from './sidecar-bundle';

/** Cold start includes JVM boot plus AndroidCompat init; be generous. */
export const DEFAULT_READY_TIMEOUT_MS = 30_000;
const READY_POLL_INTERVAL_MS = 500;

/** Ask the OS for a free loopback port, then hand it to the JVM. */
export async function allocatePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      if (address === null || typeof address === 'string') {
        server.close();
        reject(new Error('Could not allocate a loopback port for the anime bridge.'));
        return;
      }
      const { port } = address;
      server.close(() => resolve(port));
    });
  });
}

export interface SidecarHandle {
  baseUrl: string;
  port: number;
  client: AnimeBridgeClient;
  stop: () => Promise<void>;
}

export interface StartSidecarOptions {
  binaries: BundleBinaries;
  port?: number;
  readyTimeoutMs?: number;
  spawnImpl?: typeof spawn;
  onLog?: (line: string) => void;
}

/**
 * Launch the bridge and wait until it reports the capabilities this client
 * needs. The desktop launch contract is `java -jar MExtensionServer.jar <port>`,
 * run from the JAR's own directory.
 */
export async function startSidecar(options: StartSidecarOptions): Promise<SidecarHandle> {
  const { binaries } = options;
  const port = options.port ?? (await allocatePort());
  const baseUrl = `http://127.0.0.1:${port}`;
  const spawnProcess = options.spawnImpl ?? spawn;

  const child: ChildProcess = spawnProcess(
    binaries.javaPath,
    ['-jar', binaries.jarPath, String(port)],
    {
      cwd: path.dirname(binaries.jarPath),
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  );

  const log = options.onLog;
  if (log) {
    child.stdout?.on('data', (chunk: Buffer) => log(chunk.toString().trimEnd()));
    child.stderr?.on('data', (chunk: Buffer) => log(chunk.toString().trimEnd()));
  }

  let exited: { code: number | null; signal: NodeJS.Signals | null } | null = null;
  child.once('exit', (code, signal) => {
    exited = { code, signal };
  });

  const stop = async (): Promise<void> => {
    if (exited !== null) return;
    // Graceful first: the server exposes a shutdown endpoint.
    try {
      await fetch(`${baseUrl}/stop`, { signal: AbortSignal.timeout(2000) });
    } catch {
      // Falling through to a signal is fine; the endpoint may already be gone.
    }
    if (exited === null) child.kill();
  };

  const client = new AnimeBridgeClient({ baseUrl });
  const deadline = Date.now() + (options.readyTimeoutMs ?? DEFAULT_READY_TIMEOUT_MS);

  while (Date.now() < deadline) {
    if (exited !== null) {
      const { code, signal } = exited as { code: number | null; signal: NodeJS.Signals | null };
      throw new Error(
        `Anime bridge exited before becoming ready (code ${code}, signal ${signal}).`,
      );
    }
    if (await client.isReady()) return { baseUrl, port, client, stop };
    await new Promise((resolve) => setTimeout(resolve, READY_POLL_INTERVAL_MS));
  }

  await stop();
  throw new Error(
    `Anime bridge did not become ready within ${options.readyTimeoutMs ?? DEFAULT_READY_TIMEOUT_MS}ms.`,
  );
}
