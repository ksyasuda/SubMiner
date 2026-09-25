import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:net';
import { request as httpRequest } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const START_TIMEOUT_MS = 15_000;
const STOP_TIMEOUT_MS = 5_000;
const POLL_INTERVAL_MS = 25;
const MAX_CHILD_OUTPUT_BYTES = 64 * 1024;

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRootArgIndex = process.argv.indexOf('--repo-root');
const repoRoot = resolve(
  repoRootArgIndex === -1 ? join(scriptDir, '..') : (process.argv[repoRootArgIndex + 1] ?? ''),
);
const paths = {
  mainEntry: join(repoRoot, 'dist', 'main-entry.js'),
  daemonEntry: join(repoRoot, 'dist', 'stats-daemon-runner.js'),
  statsServer: join(repoRoot, 'dist', 'core', 'services', 'stats-server.js'),
  tracker: join(repoRoot, 'dist', 'core', 'services', 'immersion-tracker-service.js'),
  statsIndex: join(repoRoot, 'stats', 'dist', 'index.html'),
};

function requireCompiledArtifacts() {
  const missing = Object.values(paths).filter((artifactPath) => !existsSync(artifactPath));
  if (missing.length > 0) {
    throw new Error(
      `Compiled runtime artifacts are missing. Run \`bun run build\` before this check:\n${missing
        .map((artifactPath) => `  - ${artifactPath}`)
        .join('\n')}`,
    );
  }
}

function requireElectronNodeRuntime() {
  if (!process.versions.electron || process.env.ELECTRON_RUN_AS_NODE !== '1') {
    throw new Error(
      'This check must run with Electron in Node mode. Use `bun run test:smoke:dist`.',
    );
  }
}

function delay(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

function listen(server, port = 0) {
  return new Promise((resolvePromise, reject) => {
    const onError = (error) => {
      server.off('listening', onListening);
      reject(error);
    };
    const onListening = () => {
      server.off('error', onError);
      const address = server.address();
      if (!address || typeof address === 'string') {
        reject(new Error('Could not resolve the stats smoke port.'));
        return;
      }
      resolvePromise(address.port);
    };
    server.once('error', onError);
    server.once('listening', onListening);
    server.listen(port, '127.0.0.1');
  });
}

function closeServer(server) {
  return new Promise((resolvePromise, reject) => {
    server.close((error) => {
      if (error) {
        reject(error);
        return;
      }
      resolvePromise();
    });
  });
}

async function findAvailablePort() {
  const reservation = createServer();
  const port = await listen(reservation);
  await closeServer(reservation);
  return port;
}

function writeSmokeConfig(userDataPath, port) {
  writeFileSync(
    join(userDataPath, 'config.json'),
    `${JSON.stringify({ stats: { serverPort: port } })}\n`,
  );
}

function spawnDaemon(userDataPath, responsePath) {
  const child = spawn(
    process.execPath,
    [
      paths.daemonEntry,
      '--stats-user-data-path',
      userDataPath,
      '--stats-response-path',
      responsePath,
    ],
    {
      cwd: repoRoot,
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  );
  let output = '';
  const appendOutput = (chunk) => {
    if (output.length >= MAX_CHILD_OUTPUT_BYTES) return;
    output += String(chunk);
    if (output.length > MAX_CHILD_OUTPUT_BYTES) {
      output = `${output.slice(0, MAX_CHILD_OUTPUT_BYTES)}\n[child output truncated]\n`;
    }
  };
  child.stdout.on('data', appendOutput);
  child.stderr.on('data', appendOutput);
  let exited = false;
  const exit = new Promise((resolvePromise) => {
    child.once('error', (error) => {
      exited = true;
      resolvePromise({ code: null, signal: null, error });
    });
    child.once('exit', (code, signal) => {
      exited = true;
      resolvePromise({ code, signal, error: null });
    });
  });
  return { child, exit, hasExited: () => exited, getOutput: () => output };
}

async function waitForResponse(responsePath, daemon) {
  const deadline = Date.now() + START_TIMEOUT_MS;
  while (Date.now() < deadline) {
    if (existsSync(responsePath)) {
      try {
        return JSON.parse(readFileSync(responsePath, 'utf8'));
      } catch {
        // The daemon may still be finishing the response file write.
      }
    }
    if (daemon.hasExited()) {
      const result = await daemon.exit;
      throw new Error(
        `Stats daemon exited before writing a startup response (${formatExit(result)}).\n${daemon.getOutput()}`,
      );
    }
    await delay(POLL_INTERVAL_MS);
  }
  throw new Error(`Timed out waiting for stats daemon startup.\n${daemon.getOutput()}`);
}

function formatExit(result) {
  if (result.error) return result.error.message;
  if (result.signal) return `signal ${result.signal}`;
  return `exit code ${result.code}`;
}

async function waitForExit(daemon, timeoutMs) {
  return await Promise.race([daemon.exit, delay(timeoutMs).then(() => null)]);
}

async function stopDaemon(daemon) {
  if (daemon.hasExited()) {
    return await daemon.exit;
  }
  daemon.child.kill('SIGTERM');
  const result = await waitForExit(daemon, STOP_TIMEOUT_MS);
  if (result) return result;
  daemon.child.kill('SIGKILL');
  await daemon.exit;
  throw new Error(`Stats daemon did not stop after SIGTERM.\n${daemon.getOutput()}`);
}

async function assertPortCanBind(port) {
  const server = createServer();
  try {
    await listen(server, port);
  } finally {
    if (server.listening) await closeServer(server);
  }
}

async function fetchOverview(url, daemon) {
  const deadline = Date.now() + START_TIMEOUT_MS;
  let lastError = null;
  while (Date.now() < deadline) {
    try {
      return await fetch(`${url}/api/stats/overview`, {
        signal: AbortSignal.timeout(START_TIMEOUT_MS),
      });
    } catch (error) {
      lastError = error;
    }
    if (daemon.hasExited()) {
      const result = await daemon.exit;
      throw new Error(
        `Stats daemon exited before accepting HTTP requests (${formatExit(result)}).\n${daemon.getOutput()}`,
      );
    }
    await delay(POLL_INTERVAL_MS);
  }
  throw new Error(
    `Timed out waiting for the stats HTTP server: ${lastError instanceof Error ? lastError.message : String(lastError)}\n${daemon.getOutput()}`,
  );
}

async function runHealthyStartup(userDataPath, port, responseName) {
  writeSmokeConfig(userDataPath, port);
  const responsePath = join(userDataPath, responseName);
  const statePath = join(userDataPath, 'stats-daemon.json');
  const databasePath = join(userDataPath, 'immersion.sqlite');
  const daemon = spawnDaemon(userDataPath, responsePath);
  let shutdownResult;
  let unfinishedRequest;

  try {
    const startup = await waitForResponse(responsePath, daemon);
    assert.deepEqual(startup, { ok: true, url: `http://127.0.0.1:${port}` });

    const response = await fetchOverview(startup.url, daemon);
    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-type') ?? '', /^application\/json\b/);
    const overview = await response.json();
    assert.equal(typeof overview, 'object');
    assert.ok(overview !== null);
    assert.ok(Array.isArray(overview.sessions));
    assert.ok(Array.isArray(overview.rollups));
    assert.equal(typeof overview.hints, 'object');

    const dashboard = await fetch(`${startup.url}/?overlay=1`, {
      signal: AbortSignal.timeout(START_TIMEOUT_MS),
    });
    assert.equal(dashboard.status, 200);
    assert.match(dashboard.headers.get('content-type') ?? '', /^text\/html\b/);
    assert.match(await dashboard.text(), /id="root"/);

    assert.ok(existsSync(databasePath), 'The compiled tracker did not create its SQLite database.');
    assert.ok(
      statSync(databasePath).size > 0,
      'The compiled tracker created an empty SQLite file.',
    );

    // Leave a real request body unfinished so shutdown must bound its drain wait.
    unfinishedRequest = httpRequest(new URL('/api/stats/anki/notesInfo', startup.url), {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': '1000',
        Expect: '100-continue',
      },
      signal: AbortSignal.timeout(START_TIMEOUT_MS),
    });
    await new Promise((resolvePromise, reject) => {
      unfinishedRequest.on('error', reject);
      unfinishedRequest.once('continue', () => {
        unfinishedRequest.write('{');
        resolvePromise();
      });
      unfinishedRequest.flushHeaders();
    });
  } finally {
    try {
      shutdownResult = await stopDaemon(daemon);
    } finally {
      unfinishedRequest?.destroy();
    }
  }

  assert.equal(shutdownResult.code, 0, `Stats daemon shutdown failed.\n${daemon.getOutput()}`);
  assert.equal(existsSync(statePath), false, 'Stats daemon state remained after shutdown.');
  await assertPortCanBind(port);
}

async function runConflictRecovery() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'subminer-compiled-conflict-'));
  const reservation = createServer();
  const port = await listen(reservation);
  const responsePath = join(userDataPath, 'conflict-response.json');
  const statePath = join(userDataPath, 'stats-daemon.json');
  writeSmokeConfig(userDataPath, port);
  const daemon = spawnDaemon(userDataPath, responsePath);
  const failures = [];

  try {
    const response = await waitForResponse(responsePath, daemon);
    if (response.ok !== false) {
      failures.push('The stats daemon reported success while its configured port was occupied.');
    }
    const result = await waitForExit(daemon, STOP_TIMEOUT_MS);
    if (!result) {
      failures.push('The stats daemon did not exit after its configured port failed to bind.');
    } else if (result.code === 0) {
      failures.push(
        'The stats daemon exited successfully after its configured port failed to bind.',
      );
    }
    if (existsSync(statePath)) {
      failures.push(
        'The stats daemon left ownership state behind after its configured port failed to bind.',
      );
    }
  } finally {
    await closeServer(reservation);
    if (!daemon.hasExited()) {
      await stopDaemon(daemon);
    }
  }

  try {
    await runHealthyStartup(userDataPath, port, 'recovery-response.json');
  } finally {
    rmSync(userDataPath, { recursive: true, force: true });
  }

  assert.deepEqual(failures, []);
}

async function main() {
  requireCompiledArtifacts();
  requireElectronNodeRuntime();

  const userDataPath = mkdtempSync(join(tmpdir(), 'subminer-compiled-runtime-'));
  try {
    await runHealthyStartup(userDataPath, await findAvailablePort(), 'startup-response.json');
  } finally {
    rmSync(userDataPath, { recursive: true, force: true });
  }
  await runConflictRecovery();

  process.stdout.write(
    `Compiled runtime smoke passed with Electron ${process.versions.electron}, Node ${process.versions.node}, HTTP, and native SQLite.\n`,
  );
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack : String(error)}\n`);
  process.exitCode = 1;
});
