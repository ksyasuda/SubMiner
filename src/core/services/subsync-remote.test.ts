import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as http from 'http';
import * as os from 'os';
import * as path from 'path';
import { runSubsyncManual } from './subsync';
import type { TriggerSubsyncFromConfigDeps } from './subsync';

const SUBTITLE_BODY = '1\n00:00:01,000 --> 00:00:02,000\nhello\n';

/**
 * Extension-backed and Jellyfin playback add subtitles by URL, which subsync
 * used to reject outright with "Subtitle file not found: https://…".
 */
async function startStubHost(): Promise<{
  url: (pathname: string) => string;
  close: () => Promise<void>;
}> {
  const server = http.createServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end(SUBTITLE_BODY);
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  const port = typeof address === 'object' && address ? address.port : 0;

  return {
    url: (pathname) => `http://127.0.0.1:${port}${pathname}`,
    close: () =>
      new Promise<void>((resolve) => {
        server.close(() => resolve());
      }),
  };
}

test('runSubsyncManual syncs stream subtitle tracks served over http', async (t) => {
  if (process.platform === 'win32') {
    t.skip('stub shell scripts are not executable on Windows');
    return;
  }

  const host = await startStubHost();
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-remote-'));
  const alassLogPath = path.join(tmpDir, 'alass-args.log');
  const alassPath = path.join(tmpDir, 'alass.sh');

  // Record argv and copy the inputs aside — subsync deletes its temp files as
  // soon as the run finishes — then write the third argument (alass's output).
  fs.writeFileSync(
    alassPath,
    `#!/bin/sh\n: > "${alassLogPath}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${alassLogPath}"; done\ncp "$1" "${tmpDir}/reference.copy"\ncp "$2" "${tmpDir}/target.copy"\nprintf '%s' "retimed" > "$3"\nexit 0\n`,
    { mode: 0o755 },
  );

  const primaryUrl = host.url('/subs/ja.srt');
  const sourceUrl = host.url('/subs/en.srt');
  const sentCommands: Array<Array<string | number>> = [];

  const deps: Pick<TriggerSubsyncFromConfigDeps, 'getMpvClient' | 'getResolvedConfig'> = {
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: (payload) => {
        sentCommands.push(payload.command);
      },
      requestProperty: async (name: string) => {
        if (name === 'path') return host.url('/stream/video.mp4');
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              lang: 'ja',
              'external-filename': primaryUrl,
            },
            {
              id: 2,
              type: 'sub',
              selected: false,
              external: true,
              lang: 'en',
              'external-filename': sourceUrl,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath: '',
      // Points nowhere on purpose: both tracks are external, so resolving ffmpeg
      // at all would throw. An empty path would just auto-discover the real
      // ffmpeg on a dev machine and prove nothing.
      ffmpegPath: path.join(tmpDir, 'no-such-ffmpeg'),
      replace: true,
    }),
  };

  try {
    const result = await runSubsyncManual({ engine: 'alass', referenceTrackId: 2 }, deps);

    assert.equal(result.ok, true, result.message);
    assert.equal(result.message, 'Subtitle synchronized with alass');

    const alassArgs = fs.readFileSync(alassLogPath, 'utf8').trim().split('\n');
    // reference, target, output — both inputs were pulled down from the host.
    assert.equal(alassArgs.length, 3);
    assert.equal(fs.readFileSync(path.join(tmpDir, 'reference.copy'), 'utf8'), SUBTITLE_BODY);
    assert.equal(fs.readFileSync(path.join(tmpDir, 'target.copy'), 'utf8'), SUBTITLE_BODY);

    const loadCommand = sentCommands.find((command) => command[0] === 'sub-add');
    assert.equal(loadCommand?.[1], alassArgs[2]);
    assert.equal(fs.readFileSync(String(loadCommand?.[1]), 'utf8'), 'retimed');
  } finally {
    await host.close();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});
