import test from 'node:test';
import assert from 'node:assert/strict';
import {
  filterMpvPollResultBySocketPath,
  matchesMpvSocketPathInCommandLine,
} from './mpv-socket-match';
import type { MpvPollResult } from './win32';

function createPollResult(commandLines: Array<string | null>): MpvPollResult {
  return {
    matches: commandLines.map((commandLine, index) => ({
      hwnd: index + 1,
      bounds: { x: index * 10, y: 0, width: 1280, height: 720 },
      area: 1280 * 720,
      isForeground: index === 0,
      commandLine,
    })),
    focusState: true,
    windowState: 'visible',
  };
}

test('matchesMpvSocketPathInCommandLine accepts equals and space-delimited socket flags', () => {
  assert.equal(
    matchesMpvSocketPathInCommandLine(
      'mpv.exe --input-ipc-server=\\\\.\\pipe\\subminer-a video.mkv',
      '\\\\.\\pipe\\subminer-a',
    ),
    true,
  );
  assert.equal(
    matchesMpvSocketPathInCommandLine(
      'mpv.exe --input-ipc-server "\\\\.\\pipe\\subminer-b" video.mkv',
      '\\\\.\\pipe\\subminer-b',
    ),
    true,
  );
  assert.equal(
    matchesMpvSocketPathInCommandLine(
      'mpv.exe --input-ipc-server=\\\\.\\pipe\\subminer-a video.mkv',
      '\\\\.\\pipe\\subminer-b',
    ),
    false,
  );
});

test('filterMpvPollResultBySocketPath keeps only matches for the requested socket path', () => {
  const result = filterMpvPollResultBySocketPath(
    createPollResult([
      'mpv.exe --input-ipc-server=\\\\.\\pipe\\subminer-a video-a.mkv',
      'mpv.exe --input-ipc-server=\\\\.\\pipe\\subminer-b video-b.mkv',
      null,
    ]),
    '\\\\.\\pipe\\subminer-b',
  );

  assert.deepEqual(
    result.matches.map((match) => match.hwnd),
    [2],
  );
  assert.equal(result.windowState, 'visible');
});

test('matchesMpvSocketPathInCommandLine rejects socket-path prefix matches', () => {
  assert.equal(
    matchesMpvSocketPathInCommandLine(
      'mpv.exe --input-ipc-server=\\\\.\\pipe\\subminer-10 video.mkv',
      '\\\\.\\pipe\\subminer-1',
    ),
    false,
  );
});
