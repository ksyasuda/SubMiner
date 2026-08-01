import { spawn, type ChildProcess } from 'node:child_process';
import path from 'node:path';
import { createServer } from 'node:net';
import { AnimeBridgeClient } from './bridge-client';
import type { BundleBinaries } from './sidecar-bundle';

/** Cold start includes JVM boot plus AndroidCompat init; be generous. */
export const DEFAULT_READY_TIMEOUT_MS = 30_000;
const READY_POLL_INTERVAL_MS = 500;
/** How long to wait for the child to go after each signal before escalating. */
const STOP_TIMEOUT_MS = 5_000;

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

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
  /**
   * Fires once when the child process goes away, however it goes away —
   * including deliberate stops. Fires immediately for a subscriber that
   * attaches after the death, so a caller cannot miss it.
   */
  onExit: (
    listener: (info: { code: number | null; signal: NodeJS.Signals | null }) => void,
  ) => void;
  client: AnimeBridgeClient;
  stop: () => Promise<void>;
}

export interface StartSidecarOptions {
  binaries: BundleBinaries;
  port?: number;
  readyTimeoutMs?: number;
  /** How long to wait for the child to go after each stop signal. Tests only. */
  stopTimeoutMs?: number;
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
  let spawnError: Error | null = null;
  let killError: Error | null = null;
  const stopTimeoutMs = options.stopTimeoutMs ?? STOP_TIMEOUT_MS;
  const hasExited = new Promise<void>((resolve) => {
    child.once('exit', (code, signal) => {
      exited = { code, signal };
      resolve();
    });
    // A ChildProcess is an EventEmitter: without this listener a failed spawn
    // (a missing or non-executable java) throws in the main process instead of
    // failing the readiness loop below. `error` never proves the child is gone
    // -- only `exit` does -- so it must never set `exited`, or a kill that
    // failed would look like a clean shutdown. A spawn failure is told apart
    // from a later kill failure by the missing pid.
    child.on('error', (error: Error) => {
      if (child.pid === undefined) {
        spawnError = error;
        resolve();
        return;
      }
      killError = error;
    });
  });

  const stop = async (): Promise<void> => {
    if (exited !== null) return;
    // Nothing was ever spawned, so there is nothing to signal or wait for.
    if (spawnError !== null) return;
    // Graceful first: the server exposes a shutdown endpoint.
    try {
      await fetch(`${baseUrl}/stop`, { signal: AbortSignal.timeout(2000) });
    } catch {
      // Falling through to a signal is fine; the endpoint may already be gone.
    }
    if (exited !== null) return;
    child.kill();
    // kill() only sends the signal. Wait for the process to actually go, so a
    // restart cannot race the old one still holding the port.
    await Promise.race([hasExited, delay(stopTimeoutMs)]);
    if (exited === null) {
      child.kill('SIGKILL');
      await Promise.race([hasExited, delay(stopTimeoutMs)]);
    }
    // Never report success while the child may still hold the port: a caller
    // that restarts on the same port would race the survivor.
    if (exited === null) {
      const failedKill = killError as Error | null;
      if (failedKill !== null) {
        throw new Error(`Anime bridge could not be killed: ${failedKill.message}`, {
          cause: failedKill,
        });
      }
      throw new Error('Anime bridge did not exit after SIGKILL.');
    }
  };

  const client = new AnimeBridgeClient({ baseUrl });
  const deadline = Date.now() + (options.readyTimeoutMs ?? DEFAULT_READY_TIMEOUT_MS);

  while (Date.now() < deadline) {
    if (spawnError !== null) {
      throw new Error(`Anime bridge could not start: ${(spawnError as Error).message}`);
    }
    if (exited !== null) {
      const { code, signal } = exited as { code: number | null; signal: NodeJS.Signals | null };
      throw new Error(
        `Anime bridge exited before becoming ready (code ${code}, signal ${signal}).`,
      );
    }
    // Cap the probe at the time actually left, so a short readiness budget is
    // not overrun by a single stalled capabilities request.
    if (await client.isReady(deadline - Date.now())) {
      const onExit: SidecarHandle['onExit'] = (listener) => {
        void hasExited.then(() => listener(exited ?? { code: null, signal: null }));
      };
      return { baseUrl, port, client, stop, onExit };
    }
    // Sleeping past the deadline would only delay the timeout report.
    const remaining = deadline - Date.now();
    if (remaining <= 0) break;
    await delay(Math.min(READY_POLL_INTERVAL_MS, remaining));
  }

  const timeout = new Error(
    `Anime bridge did not become ready within ${options.readyTimeoutMs ?? DEFAULT_READY_TIMEOUT_MS}ms.`,
  );
  // A failed shutdown must not mask why startup failed, but it must not be
  // dropped either: a surviving child still holds the port, which decides
  // whether a caller may retry on it.
  await stop().catch((error: unknown) => {
    timeout.cause = error;
  });
  throw timeout;
}
