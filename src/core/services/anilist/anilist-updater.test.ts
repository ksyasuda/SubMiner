import test from 'node:test';
import assert from 'node:assert/strict';
import * as childProcess from 'child_process';

import { guessAnilistMediaInfo, updateAnilistPostWatchProgress } from './anilist-updater';

function createJsonResponse(payload: unknown): Response {
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

test('guessAnilistMediaInfo uses guessit output when available', async () => {
  const originalExecFile = childProcess.execFile;
  (
    childProcess as unknown as {
      execFile: typeof childProcess.execFile;
    }
  ).execFile = ((...args: unknown[]) => {
    const callback = args[args.length - 1];
    const cb =
      typeof callback === 'function'
        ? (callback as (error: Error | null, stdout: string, stderr: string) => void)
        : null;
    cb?.(null, JSON.stringify({ title: 'Guessit Title', episode: 7 }), '');
    return {} as childProcess.ChildProcess;
  }) as typeof childProcess.execFile;

  try {
    const result = await guessAnilistMediaInfo('/tmp/demo.mkv', null);
    assert.deepEqual(result, {
      title: 'Guessit Title',
      episode: 7,
      source: 'guessit',
    });
  } finally {
    (
      childProcess as unknown as {
        execFile: typeof childProcess.execFile;
      }
    ).execFile = originalExecFile;
  }
});

test('guessAnilistMediaInfo falls back to parser when guessit fails', async () => {
  const originalExecFile = childProcess.execFile;
  (
    childProcess as unknown as {
      execFile: typeof childProcess.execFile;
    }
  ).execFile = ((...args: unknown[]) => {
    const callback = args[args.length - 1];
    const cb =
      typeof callback === 'function'
        ? (callback as (error: Error | null, stdout: string, stderr: string) => void)
        : null;
    cb?.(new Error('guessit not found'), '', '');
    return {} as childProcess.ChildProcess;
  }) as typeof childProcess.execFile;

  try {
    const result = await guessAnilistMediaInfo('/tmp/My Anime S01E03.mkv', null);
    assert.deepEqual(result, {
      title: 'My Anime',
      episode: 3,
      source: 'fallback',
    });
  } finally {
    (
      childProcess as unknown as {
        execFile: typeof childProcess.execFile;
      }
    ).execFile = originalExecFile;
  }
});

test('updateAnilistPostWatchProgress updates progress when behind', async () => {
  const originalFetch = globalThis.fetch;
  let call = 0;
  globalThis.fetch = (async () => {
    call += 1;
    if (call === 1) {
      return createJsonResponse({
        data: {
          Page: {
            media: [
              {
                id: 11,
                episodes: 24,
                title: { english: 'Demo Show', romaji: 'Demo Show' },
              },
            ],
          },
        },
      });
    }
    if (call === 2) {
      return createJsonResponse({
        data: {
          Media: {
            id: 11,
            mediaListEntry: { progress: 2, status: 'CURRENT' },
          },
        },
      });
    }
    return createJsonResponse({
      data: { SaveMediaListEntry: { progress: 3, status: 'CURRENT' } },
    });
  }) as typeof fetch;

  try {
    const result = await updateAnilistPostWatchProgress('token', 'Demo Show', 3);
    assert.equal(result.status, 'updated');
    assert.match(result.message, /episode 3/i);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('updateAnilistPostWatchProgress skips when progress already reached', async () => {
  const originalFetch = globalThis.fetch;
  let call = 0;
  globalThis.fetch = (async () => {
    call += 1;
    if (call === 1) {
      return createJsonResponse({
        data: {
          Page: {
            media: [{ id: 22, episodes: 12, title: { english: 'Skip Show' } }],
          },
        },
      });
    }
    return createJsonResponse({
      data: {
        Media: { id: 22, mediaListEntry: { progress: 12, status: 'CURRENT' } },
      },
    });
  }) as typeof fetch;

  try {
    const result = await updateAnilistPostWatchProgress('token', 'Skip Show', 10);
    assert.equal(result.status, 'skipped');
    assert.match(result.message, /already at episode/i);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('updateAnilistPostWatchProgress returns error when search fails', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    createJsonResponse({
      errors: [{ message: 'bad request' }],
    })) as typeof fetch;

  try {
    const result = await updateAnilistPostWatchProgress('token', 'Bad', 1);
    assert.equal(result.status, 'error');
    assert.match(result.message, /search failed/i);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
