import assert from 'node:assert/strict';
import test from 'node:test';

import {
  CHANGELOG_REQUEST_TIMEOUT_MS,
  createChangelogRuntime,
  withRequestTimeout,
} from './changelog-runtime';
import type { FetchLike, FetchResponseLike } from '../update/release-assets';
import { createCurlFetch } from '../update/fetch-adapter';

function okResponse(body: string): FetchResponseLike {
  return {
    ok: true,
    status: 200,
    json: async () => JSON.parse(body),
    text: async () => body,
    arrayBuffer: async () => new ArrayBuffer(0),
  };
}

test('request timeout wrapper attaches an abort signal to every request', async () => {
  const seen: Array<{ url: string; init?: Record<string, unknown> }> = [];
  const wrapped = withRequestTimeout(async (url, init) => {
    seen.push({ url, init });
    return okResponse('body');
  }, 1_234);

  await wrapped('https://example.test/a');
  await wrapped('https://example.test/b', { headers: { 'User-Agent': 'SubMiner' } });

  assert.equal(seen.length, 2);
  for (const request of seen) {
    assert.ok(request.init?.signal instanceof AbortSignal, 'each request carries a signal');
    assert.equal((request.init?.signal as AbortSignal).aborted, false);
  }
  // Existing init is preserved rather than replaced.
  assert.deepEqual(seen[1]?.init?.headers, { 'User-Agent': 'SubMiner' });
});

test('request timeout wrapper aborts a request that never settles', async () => {
  let observed: AbortSignal | undefined;
  const wrapped = withRequestTimeout((_url, init) => {
    observed = init?.signal as AbortSignal;
    return new Promise<FetchResponseLike>(() => {
      // Never resolves, standing in for a stalled connection.
    });
  }, 10);

  void wrapped('https://example.test/stalled');
  await new Promise((resolve) => setTimeout(resolve, 40));

  assert.equal(observed?.aborted, true);
});

test('changelog runtime times out both requests instead of hanging the modal', async () => {
  const signals: AbortSignal[] = [];
  // Stays pending until the deadline fires, standing in for a stalled server.
  const stalling: FetchLike = (_url, init) => {
    const signal = init?.signal as AbortSignal;
    signals.push(signal);
    return new Promise((_resolve, reject) => {
      signal.addEventListener('abort', () => reject(signal.reason), { once: true });
    });
  };

  const runtime = createChangelogRuntime({
    getInstalledVersion: () => '0.19.2',
    getUpdateChannel: () => 'stable',
    resourcesPath: '/res',
    appPath: '/app',
    dirname: '/app/dist/main',
    joinPath: (...parts) => parts.join('/'),
    fileExists: () => false,
    readFile: () => '',
    logWarn: () => {},
    createFetch: () => withRequestTimeout(stalling, 10),
  });

  const snapshot = await runtime.getChangelogSnapshot();

  // Release lookup and changelog download both hit the deadline rather than hang.
  assert.equal(signals.length, 2);
  assert.ok(
    signals.every((signal) => signal.aborted),
    'every stalled request was aborted',
  );
  // No bundled copy is readable here, so the timeout surfaces as an error state
  // that does not claim a bundled changelog is on screen.
  assert.match(snapshot.error ?? '', /^Changelog unavailable: /);
  assert.match(snapshot.error ?? '', /timed out|timeout|abort/i);
  assert.deepEqual(snapshot.entries, []);
});

test('changelog timeout reaches the curl transport, not just global fetch', async () => {
  // The POSIX transport is curl, whose own --max-time is 60s; the changelog
  // deadline is shorter, so the signal has to actually terminate the process.
  let killed: string | undefined;
  const curlFetch = createCurlFetch({
    execFile: ((
      _file: string,
      _args: readonly string[],
      _options: unknown,
      _callback: unknown,
    ) => ({
      kill: (signal?: string) => {
        killed = signal;
        return true;
      },
    })) as never,
  });

  const wrapped = withRequestTimeout(curlFetch, 10);
  await assert.rejects(wrapped('https://example.test/stalled'));

  assert.equal(killed, 'SIGKILL', 'the stalled curl process is killed at the changelog deadline');
});

test('changelog request timeout is finite', () => {
  assert.ok(Number.isFinite(CHANGELOG_REQUEST_TIMEOUT_MS));
  assert.ok(CHANGELOG_REQUEST_TIMEOUT_MS > 0);
});
