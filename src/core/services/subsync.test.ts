import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { pathToFileURL } from 'url';
import {
  TriggerSubsyncFromConfigDeps,
  runSubsyncManual,
  triggerSubsyncFromConfig,
} from './subsync';
import type { SubsyncManualPayload } from '../../types';

function makeDeps(
  overrides: Partial<TriggerSubsyncFromConfigDeps> = {},
): TriggerSubsyncFromConfigDeps {
  const mpvClient = {
    connected: true,
    currentAudioStreamIndex: null,
    send: () => {},
    requestProperty: async (name: string) => {
      if (name === 'path') return '/tmp/video.mkv';
      if (name === 'sid') return 1;
      if (name === 'secondary-sid') return null;
      if (name === 'track-list') {
        return [
          { id: 1, type: 'sub', selected: true, lang: 'jpn' },
          {
            id: 2,
            type: 'sub',
            selected: false,
            external: true,
            lang: 'eng',
            'external-filename': '/tmp/ref.srt',
          },
          { id: 3, type: 'audio', selected: true, 'ff-index': 1 },
        ];
      }
      return null;
    },
  };

  return {
    getMpvClient: () => mpvClient,
    getResolvedConfig: () => ({
      alassPath: '/usr/bin/alass',
      ffsubsyncPath: '/usr/bin/ffsubsync',
      ffmpegPath: '/usr/bin/ffmpeg',
    }),
    isSubsyncInProgress: () => false,
    setSubsyncInProgress: () => {},
    showMpvOsd: () => {},
    runWithSubsyncSpinner: async <T>(task: () => Promise<T>) => task(),
    openManualPicker: () => {},
    ...overrides,
  };
}

test('triggerSubsyncFromConfig returns early when already in progress', async () => {
  const osd: string[] = [];
  await triggerSubsyncFromConfig(
    makeDeps({
      isSubsyncInProgress: () => true,
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );
  assert.deepEqual(osd, ['Subsync already running']);
});

test('triggerSubsyncFromConfig opens manual picker', async () => {
  const osd: string[] = [];
  let payloadTrackCount = 0;
  let ffsubsyncAvailable: boolean | null = null;
  let payloadMediaPath: string | null | undefined = undefined;
  let inProgressState: boolean | null = null;

  await triggerSubsyncFromConfig(
    makeDeps({
      getResolvedConfig: () => ({
        mode: 'manual',
        alassPath: process.execPath,
        ffsubsyncPath: process.execPath,
        ffmpegPath: process.execPath,
      }),
      openManualPicker: (payload) => {
        payloadTrackCount = payload.subtitleTracks.length;
        ffsubsyncAvailable = payload.ffsubsyncAvailable;
        payloadMediaPath = payload.mediaPath;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
      setSubsyncInProgress: (value) => {
        inProgressState = value;
      },
    }),
  );

  assert.equal(payloadTrackCount, 2);
  assert.equal(ffsubsyncAvailable, true);
  assert.equal(payloadMediaPath, '/tmp/video.mkv');
  assert.ok(osd.includes('Subsync: choose engine and subtitles'));
  assert.equal(inProgressState, false);
});

test('triggerSubsyncFromConfig marks ffsubsync unavailable for remote media paths', async () => {
  let ffsubsyncAvailable: boolean | null = null;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return 'https://jellyfin.example/Videos/movie/stream.mkv';
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              { id: 1, type: 'sub', selected: true, lang: 'jpn' },
              {
                id: 2,
                type: 'sub',
                selected: false,
                external: true,
                lang: 'eng',
                'external-filename': 'https://jellyfin.example/subs/eng.srt',
              },
            ];
          }
          return null;
        },
      }),
      openManualPicker: (payload) => {
        ffsubsyncAvailable = payload.ffsubsyncAvailable;
      },
    }),
  );

  assert.equal(ffsubsyncAvailable, false);
});

test('triggerSubsyncFromConfig dedupes repeated subtitle source tracks', async () => {
  let payloadTrackCount = 0;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return '/tmp/video.mkv';
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'track-list') {
            return [
              { id: 1, type: 'sub', selected: true, lang: 'jpn' },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': '/tmp/ref.srt',
              },
              {
                id: 3,
                type: 'sub',
                selected: false,
                external: true,
                lang: 'eng',
                'external-filename': '/tmp/ref.srt',
              },
            ];
          }
          return null;
        },
      }),
      openManualPicker: (payload) => {
        payloadTrackCount = payload.subtitleTracks.length;
      },
    }),
  );

  assert.equal(payloadTrackCount, 2);
});

test('triggerSubsyncFromConfig keeps both active tracks when they share a file', async () => {
  let payload: SubsyncManualPayload | null = null;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return '/tmp/video.mkv';
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'track-list') {
            // mpv appends a duplicate entry when the same file is re-added, so
            // the primary and secondary slots can point at one path.
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                'external-filename': '/tmp/ref.srt',
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                'external-filename': '/tmp/ref.srt',
              },
              {
                id: 3,
                type: 'sub',
                selected: false,
                external: true,
                'external-filename': '/tmp/ref.srt',
              },
            ];
          }
          return null;
        },
      }),
      openManualPicker: (nextPayload) => {
        payload = nextPayload;
      },
    }),
  );

  assert.ok(payload);
  const resolved = payload as SubsyncManualPayload;
  assert.deepEqual(
    resolved.subtitleTracks.map((track) => track.id),
    [1, 2],
  );
  assert.equal(resolved.defaultReferenceTrackId, 2);
  assert.equal(resolved.defaultTargetTrackId, 1);
});

test('triggerSubsyncFromConfig reports failures to OSD', async () => {
  const osd: string[] = [];
  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => null,
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.ok(osd.some((line) => line.startsWith('Subsync failed: MPV not connected')));
});

test('runSubsyncManual requires a reference track for alass', async () => {
  const result = await runSubsyncManual({ engine: 'alass', referenceTrackId: null }, makeDeps());

  assert.deepEqual(result, {
    ok: false,
    message: 'Select a reference subtitle track for alass',
  });
});

test('runSubsyncManual rejects alass when reference and target are the same track', async () => {
  const result = await runSubsyncManual(
    { engine: 'alass', referenceTrackId: 2, targetTrackId: 2 },
    makeDeps(),
  );

  assert.deepEqual(result, {
    ok: false,
    message: 'Reference and out-of-sync subtitles must be different tracks',
  });
});

test('runSubsyncManual rejects an unknown target track', async () => {
  const result = await runSubsyncManual(
    { engine: 'alass', referenceTrackId: 2, targetTrackId: 99 },
    makeDeps(),
  );

  assert.deepEqual(result, {
    ok: false,
    message: 'Select the out-of-sync subtitle track to retime',
  });
});

test('runSubsyncManual rejects the video reference for remote media', async () => {
  const result = await runSubsyncManual(
    { engine: 'alass', referenceMode: 'video' },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return 'https://jellyfin.example/Videos/movie/stream.mkv';
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [{ id: 1, type: 'sub', selected: true, lang: 'jpn' }];
          }
          return null;
        },
      }),
    }),
  );

  assert.equal(result.ok, false);
  assert.match(result.message, /cannot use a stream URL as reference/);
});

test('openSubsyncManualPicker defaults the reference to the secondary subtitle track', async () => {
  let payload: SubsyncManualPayload | null = null;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return '/tmp/video.mkv';
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 3;
          if (name === 'track-list') {
            return [
              { id: 1, type: 'sub', selected: true, lang: 'jpn' },
              {
                id: 2,
                type: 'sub',
                selected: false,
                external: true,
                lang: 'eng',
                'external-filename': '/tmp/other.srt',
              },
              {
                id: 3,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': '/tmp/secondary.srt',
              },
            ];
          }
          return null;
        },
      }),
      openManualPicker: (nextPayload) => {
        payload = nextPayload;
      },
    }),
  );

  assert.ok(payload);
  const resolved = payload as SubsyncManualPayload;
  assert.deepEqual(
    resolved.subtitleTracks.map((track) => track.id),
    [1, 2, 3],
  );
  assert.equal(resolved.defaultReferenceTrackId, 3);
  assert.equal(resolved.defaultTargetTrackId, 1);
  assert.equal(resolved.videoReferenceAvailable, true);
});

test('openSubsyncManualPicker never defaults to a reference missing from the track list', async () => {
  let payload: SubsyncManualPayload | null = null;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return '/tmp/video.mkv';
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'track-list') {
            return [
              { id: 1, type: 'sub', selected: true, lang: 'jpn' },
              // Secondary track with no usable file path: filtered out of the picker.
              { id: 2, type: 'sub', selected: true, external: true, 'external-filename': '' },
              { id: 3, type: 'sub', selected: false, lang: 'eng' },
            ];
          }
          return null;
        },
      }),
      openManualPicker: (nextPayload) => {
        payload = nextPayload;
      },
    }),
  );

  assert.ok(payload);
  const resolved = payload as SubsyncManualPayload;
  assert.deepEqual(
    resolved.subtitleTracks.map((track) => track.id),
    [1, 3],
  );
  assert.equal(resolved.defaultReferenceTrackId, 3);
});

test('triggerSubsyncFromConfig does not validate sync tool paths before manual selection', async () => {
  const osd: string[] = [];
  const inProgress: boolean[] = [];
  let payloadTrackCount = 0;

  await triggerSubsyncFromConfig(
    makeDeps({
      getResolvedConfig: () => ({
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
      }),
      setSubsyncInProgress: (value) => {
        inProgress.push(value);
      },
      openManualPicker: (payload) => {
        payloadTrackCount = payload.subtitleTracks.length;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.deepEqual(inProgress, [true, false]);
  assert.equal(payloadTrackCount, 2);
  assert.deepEqual(osd, ['Subsync: choose engine and subtitles']);
});

function writeExecutableScript(filePath: string, content: string): void {
  fs.writeFileSync(filePath, content, { encoding: 'utf8', mode: 0o755 });
  fs.chmodSync(filePath, 0o755);
}

function toShellPath(filePath: string): string {
  if (process.platform !== 'win32') {
    return filePath;
  }

  return filePath.replace(/\\/g, '/').replace(/^([A-Za-z]):\//, (_, driveLetter: string) => {
    return `/mnt/${driveLetter.toLowerCase()}/`;
  });
}

function fromShellPath(filePath: string): string {
  if (process.platform !== 'win32') {
    return filePath;
  }

  return filePath
    .replace(/^\/mnt\/([a-z])\//, (_, driveLetter: string) => {
      return `${driveLetter.toUpperCase()}:/`;
    })
    .replace(/\//g, '\\');
}

test('runSubsyncManual constructs ffsubsync command and returns success', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-ffsubsync-'));
  const ffsubsyncLogPath = path.join(tmpDir, 'ffsubsync-args.log');
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'sub');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(alassPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    ffsubsyncPath,
    `#!/bin/sh\n: > "${toShellPath(ffsubsyncLogPath)}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${toShellPath(ffsubsyncLogPath)}"; done\nout=\"\"\nprev=\"\"\nfor arg in \"$@\"; do\n  if [ \"$prev\" = \"-o\" ]; then out=\"$arg\"; fi\n  prev=\"$arg\"\ndone\nif [ -n \"$out\" ]; then : > \"$out\"; fi\nexit 0\n`,
  );

  const sentCommands: Array<Array<string | number>> = [];
  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: 2,
      send: (payload) => {
        sentCommands.push(payload.command);
      },
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              'external-filename': primaryPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManual({ engine: 'ffsubsync' }, deps);

  assert.equal(result.ok, true);
  assert.equal(result.message, 'Subtitle synchronized with ffsubsync');
  const ffArgs = fs.readFileSync(ffsubsyncLogPath, 'utf8').trim().split('\n');
  assert.equal(ffArgs[0], toShellPath(videoPath));
  assert.ok(ffArgs.includes('-i'));
  assert.ok(ffArgs.includes(toShellPath(primaryPath)));
  assert.ok(ffArgs.includes('--reference-stream'));
  assert.ok(ffArgs.includes('0:2'));
  const ffOutputFlagIndex = ffArgs.indexOf('-o');
  assert.equal(ffOutputFlagIndex >= 0, true);
  const stagedArg = ffArgs[ffOutputFlagIndex + 1]!;
  assert.equal(path.dirname(fromShellPath(stagedArg)), tmpDir);
  assert.equal(path.extname(stagedArg), '.srt');
  assert.notEqual(stagedArg, toShellPath(primaryPath));
  assert.equal(fs.existsSync(fromShellPath(stagedArg)), false);
  assert.equal(fs.existsSync(primaryPath), true);
  assert.deepEqual(sentCommands[0], ['sub-add', primaryPath]);
  assert.deepEqual(sentCommands[1], ['set_property', 'sub-delay', 0]);
});

test('runSubsyncManual writes deterministic _retimed filename when replace is false', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-ffsubsync-no-replace-'));
  const ffsubsyncLogPath = path.join(tmpDir, 'ffsubsync-args.log');
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'episode.ja.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'sub');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(alassPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    ffsubsyncPath,
    `#!/bin/sh\n: > "${toShellPath(ffsubsyncLogPath)}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${toShellPath(ffsubsyncLogPath)}"; done\nout=\"\"\nprev=\"\"\nfor arg in \"$@\"; do\n  if [ \"$prev\" = \"-o\" ]; then out=\"$arg\"; fi\n  prev=\"$arg\"\ndone\nif [ -n \"$out\" ]; then : > \"$out\"; fi\nexit 0\n`,
  );

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              'external-filename': primaryPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
      replace: false,
    }),
  });

  const result = await runSubsyncManual({ engine: 'ffsubsync' }, deps);

  assert.equal(result.ok, true);
  const ffArgs = fs.readFileSync(ffsubsyncLogPath, 'utf8').trim().split('\n');
  const ffOutputFlagIndex = ffArgs.indexOf('-o');
  assert.equal(ffOutputFlagIndex >= 0, true);
  const stagedOutputArg = ffArgs[ffOutputFlagIndex + 1]!;
  const expectedRetimedPath = path.join(tmpDir, 'episode.ja_retimed.srt');
  assert.equal(path.dirname(fromShellPath(stagedOutputArg)), tmpDir);
  assert.equal(path.extname(stagedOutputArg), '.srt');
  assert.equal(fs.existsSync(fromShellPath(stagedOutputArg)), false);
  assert.equal(fs.existsSync(expectedRetimedPath), true);
});

test('runSubsyncManual reports ffsubsync command failures with details', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-ffsubsync-failure-'));
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'sub');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(alassPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(ffsubsyncPath, '#!/bin/sh\necho "reference audio missing" >&2\nexit 1\n');

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              'external-filename': primaryPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManual({ engine: 'ffsubsync' }, deps);

  assert.equal(result.ok, false);
  assert.equal(result.message.startsWith('ffsubsync synchronization failed'), true);
  assert.match(result.message, /code=1/);
  assert.match(result.message, /reference audio missing/);
});

test('runSubsyncManual constructs alass command and returns failure on non-zero exit', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-alass-'));
  const alassLogPath = path.join(tmpDir, 'alass-args.log');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');
  const sourcePath = path.join(tmpDir, 'source.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'sub');
  fs.writeFileSync(sourcePath, 'sub2');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(ffsubsyncPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    alassPath,
    `#!/bin/sh\n: > "${toShellPath(alassLogPath)}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${toShellPath(alassLogPath)}"; done\nexit 1\n`,
  );

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              'external-filename': primaryPath,
            },
            {
              id: 2,
              type: 'sub',
              selected: false,
              external: true,
              'external-filename': sourcePath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManual({ engine: 'alass', referenceTrackId: 2 }, deps);

  assert.equal(result.ok, false);
  assert.equal(typeof result.message, 'string');
  assert.equal(result.message.startsWith('alass synchronization failed'), true);
  const alassArgs = fs.readFileSync(alassLogPath, 'utf8').trim().split('\n');
  assert.equal(alassArgs[0], toShellPath(sourcePath));
  assert.equal(alassArgs[1], toShellPath(primaryPath));
});

function makeAlassSelectionDeps(tmpDir: string): {
  deps: TriggerSubsyncFromConfigDeps;
  alassLogPath: string;
  videoPath: string;
  primaryPath: string;
  sourcePath: string;
  sentCommands: Array<Array<string | number>>;
} {
  const alassLogPath = path.join(tmpDir, 'alass-args.log');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');
  const sourcePath = path.join(tmpDir, 'source.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'sub');
  fs.writeFileSync(sourcePath, 'sub2');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(ffsubsyncPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    alassPath,
    `#!/bin/sh\n: > "${toShellPath(alassLogPath)}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${toShellPath(alassLogPath)}"; done\n: > "$3"\nexit 0\n`,
  );

  const trackList: Array<Record<string, unknown>> = [
    { id: 1, type: 'sub', selected: true, external: true, 'external-filename': primaryPath },
    { id: 2, type: 'sub', selected: true, external: true, 'external-filename': sourcePath },
  ];
  const sentCommands: Array<Array<string | number>> = [];
  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: (payload) => {
        sentCommands.push(payload.command);
        if (payload.command[0] === 'sub-add' || payload.command[0] === 'sub_add') {
          trackList.push({
            id: trackList.length + 1,
            type: 'sub',
            selected: false,
            external: true,
            'external-filename': payload.command[1],
          });
        }
      },
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return trackList;
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
      replace: false,
    }),
  });

  return { deps, alassLogPath, videoPath, primaryPath, sourcePath, sentCommands };
}

test('runSubsyncManual uses the video file as alass reference when requested', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-alass-video-ref-'));
  const { deps, alassLogPath, videoPath, primaryPath } = makeAlassSelectionDeps(tmpDir);

  const result = await runSubsyncManual({ engine: 'alass', referenceMode: 'video' }, deps);

  assert.equal(result.ok, true);
  const alassArgs = fs.readFileSync(alassLogPath, 'utf8').trim().split('\n');
  assert.equal(alassArgs[0], toShellPath(videoPath));
  assert.equal(alassArgs[1], toShellPath(primaryPath));
});

test('runSubsyncManual retimes the selected target track instead of the primary', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-alass-target-'));
  const { deps, alassLogPath, primaryPath, sourcePath, sentCommands } =
    makeAlassSelectionDeps(tmpDir);

  const result = await runSubsyncManual(
    { engine: 'alass', referenceTrackId: 1, targetTrackId: 2 },
    deps,
  );

  assert.equal(result.ok, true);
  const alassArgs = fs.readFileSync(alassLogPath, 'utf8').trim().split('\n');
  assert.equal(alassArgs[0], toShellPath(primaryPath));
  assert.equal(alassArgs[1], toShellPath(sourcePath));
  assert.equal(sentCommands[0]?.[0], 'sub-add');
  assert.equal(sentCommands[0]?.[2], 'auto');
});

test('runSubsyncManual keeps a retimed secondary track in the secondary slot', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-alass-secondary-slot-'));
  const alassPath = path.join(tmpDir, 'alass.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'ja.srt');
  const secondaryPath = path.join(tmpDir, 'en.srt');
  const retimedPath = path.join(tmpDir, 'en_retimed.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'ja');
  fs.writeFileSync(secondaryPath, 'en');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(ffsubsyncPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(alassPath, '#!/bin/sh\n: > "$3"\nexit 0\n');

  const trackList: Array<Record<string, unknown>> = [
    { id: 1, type: 'sub', selected: true, external: true, 'external-filename': primaryPath },
    { id: 2, type: 'sub', selected: true, external: true, 'external-filename': secondaryPath },
  ];
  const sentCommands: Array<Array<string | number>> = [];
  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: (payload) => {
        sentCommands.push(payload.command);
        if (payload.command[0] === 'sub-add') {
          trackList.push({
            id: 3,
            type: 'sub',
            selected: false,
            external: true,
            'external-filename': payload.command[1],
          });
        }
      },
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return 2;
        if (name === 'track-list') return trackList;
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
      replace: false,
    }),
  });

  const result = await runSubsyncManual(
    { engine: 'alass', referenceTrackId: 1, targetTrackId: 2 },
    deps,
  );

  assert.equal(result.ok, true);
  assert.deepEqual(sentCommands[0], ['sub-add', retimedPath, 'auto']);
  assert.deepEqual(sentCommands[1], ['set_property', 'secondary-sub-delay', 0]);
  assert.deepEqual(sentCommands[2], ['set_property', 'secondary-sid', 3]);
  assert.equal(
    sentCommands.some((command) => command[1] === 'sub-delay'),
    false,
  );
  assert.equal(
    sentCommands.some((command) => command[1] === 'sid'),
    false,
  );
  assert.equal(
    sentCommands.some((command) => command[1] === 'sid'),
    false,
  );
});

test('runSubsyncManual keeps internal alass source file alive until sync finishes', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-alass-internal-source-'));
  const alassPath = path.join(tmpDir, 'alass.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'sub');
  writeExecutableScript(ffsubsyncPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    ffmpegPath,
    '#!/bin/sh\nout=""\nfor arg in "$@"; do out="$arg"; done\nif [ -n "$out" ]; then : > "$out"; fi\nexit 0\n',
  );
  writeExecutableScript(
    alassPath,
    '#!/bin/sh\nsleep 0.2\nif [ ! -f "$1" ]; then echo "missing reference subtitle" >&2; exit 1; fi\nif [ ! -f "$2" ]; then echo "missing input subtitle" >&2; exit 1; fi\n: > "$3"\nexit 0\n',
  );

  const sentCommands: Array<Array<string | number>> = [];
  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: (payload) => {
        sentCommands.push(payload.command);
      },
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              'external-filename': primaryPath,
            },
            {
              id: 2,
              type: 'sub',
              selected: false,
              external: false,
              'ff-index': 2,
              codec: 'ass',
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManual({ engine: 'alass', referenceTrackId: 2 }, deps);

  assert.equal(result.ok, true);
  assert.equal(result.message, 'Subtitle synchronized with alass');
  assert.equal(sentCommands[0]?.[0], 'sub-add');
  assert.deepEqual(sentCommands[1], ['set_property', 'sub-delay', 0]);
});

test('runSubsyncManual resolves string sid values from mpv stream properties', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-stream-sid-'));
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const ffsubsyncLogPath = path.join(tmpDir, 'ffsubsync-args.log');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const videoPath = path.join(tmpDir, 'video.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');

  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(primaryPath, 'subtitle');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(alassPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    ffsubsyncPath,
    `#!/bin/sh\nmkdir -p "${toShellPath(tmpDir)}"\n: > "${toShellPath(ffsubsyncLogPath)}"\nfor arg in "$@"; do printf '%s\\n' "$arg" >> "${toShellPath(ffsubsyncLogPath)}"; done\nprev=""\nout=""\nfor arg in "$@"; do\n  if [ "$prev" = "--reference-stream" ]; then :; fi\n  if [ "$prev" = "-o" ]; then out="$arg"; fi\n  prev="$arg"\ndone\nif [ -n "$out" ]; then : > "$out"; fi`,
  );

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return '1';
        if (name === 'secondary-sid') return '2';
        if (name === 'track-list') {
          return [
            {
              id: '1',
              type: 'sub',
              selected: true,
              external: true,
              'external-filename': primaryPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath,
      ffsubsyncPath,
      ffmpegPath,
    }),
  });

  const result = await runSubsyncManual({ engine: 'ffsubsync' }, deps);

  assert.equal(result.ok, true);
  assert.equal(result.message, 'Subtitle synchronized with ffsubsync');
  const ffArgs = fs.readFileSync(ffsubsyncLogPath, 'utf8').trim().split('\n');
  const syncOutputIndex = ffArgs.indexOf('-o');
  assert.equal(syncOutputIndex >= 0, true);
  const outputPath = ffArgs[syncOutputIndex + 1];
  assert.equal(typeof outputPath, 'string');
  assert.ok(outputPath!.length > 0);
  assert.equal(fs.existsSync(fromShellPath(outputPath!)), false);
  assert.equal(fs.readFileSync(primaryPath, 'utf8'), '');
});

function generateSyntheticSpans(
  seedInit: bigint,
  step: bigint,
  limitSec = 900,
): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number }> = [];
  let seed = seedInit;
  const mask = (1n << 64n) - 1n;
  let t = 5.0;
  while (t < limitSec) {
    seed = (seed * 6364136223846793005n + step) & mask;
    const len = 0.8 + (Number(seed >> 33n) / Number(1n << 31n)) * 2.5;
    const gap = 0.5 + (Number((seed >> 13n) & 0xffffn) / 65536.0) * 3.5;
    out.push({ start: t, end: t + len });
    t += len + gap;
  }
  return out;
}

function fmtSrtTime(sec: number): string {
  const totalMs = Math.max(0, Math.round(sec * 1000));
  const ms = totalMs % 1000;
  const totalSec = Math.floor(totalMs / 1000);
  const s = totalSec % 60;
  const totalMin = Math.floor(totalSec / 60);
  const m = totalMin % 60;
  const h = Math.floor(totalMin / 60);
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')},${String(ms).padStart(3, '0')}`;
}

function fmtAssTime(sec: number): string {
  const totalCs = Math.max(0, Math.round(sec * 100));
  const cs = totalCs % 100;
  const totalSec = Math.floor(totalCs / 100);
  const s = totalSec % 60;
  const totalMin = Math.floor(totalSec / 60);
  const m = totalMin % 60;
  const h = Math.floor(totalMin / 60);
  return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}.${String(cs).padStart(2, '0')}`;
}

function writeSrtFromSpans(
  filePath: string,
  spans: ReadonlyArray<{ start: number; end: number }>,
  shiftSec = 0,
): void {
  const content = spans
    .map(
      (span, idx) =>
        `${idx + 1}\n${fmtSrtTime(span.start + shiftSec)} --> ${fmtSrtTime(span.end + shiftSec)}\nセリフ ${idx + 1}\n`,
    )
    .join('\n');
  fs.writeFileSync(filePath, content, 'utf8');
}

function writeAssWithSignsFromSpans(
  filePath: string,
  spans: ReadonlyArray<{ start: number; end: number }>,
  shiftSec = 0,
): void {
  const lines = [
    '[Script Info]',
    'Title: Test',
    '',
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
  ];
  for (let i = 0; i < spans.length; i += 1) {
    const s = spans[i]!;
    lines.push(
      `Dialogue: 0,${fmtAssTime(s.start + shiftSec)},${fmtAssTime(s.end + shiftSec)},Default,,0,0,0,,Line ${i + 1}`,
    );
  }
  // Add heavy stacked karaoke/song & sign noise that would confuse naive syncers.
  for (let k = 0; k < 10; k += 1) {
    lines.push(
      `Dialogue: 0,${fmtAssTime(910)},${fmtAssTime(990)},OPRomaji,,0,0,0,,{\\k20}song ${k}`,
    );
    lines.push(
      `Dialogue: 0,${fmtAssTime(20 + k * 5)},${fmtAssTime(24 + k * 5)},Sign,,0,0,0,,{\\pos(100,200)}Sign ${k}`,
    );
  }
  fs.writeFileSync(filePath, lines.join('\n'), 'utf8');
}

function writeSpeechPcmFile(
  pcmPath: string,
  spans: ReadonlyArray<{ start: number; end: number }>,
  durationSec = 920,
): void {
  const sampleRate = 8000;
  const samples = new Int16Array(sampleRate * durationSec);
  for (let i = 0; i < samples.length; i += 1) {
    const t = i / sampleRate;
    samples[i] = Math.round(Math.sin(2 * Math.PI * 500 * t) * 120);
  }
  for (const span of spans) {
    const sIdx = Math.max(0, Math.floor(span.start * sampleRate));
    const eIdx = Math.min(samples.length, Math.floor(span.end * sampleRate));
    for (let i = sIdx; i < eIdx; i += 1) {
      const t = i / sampleRate;
      const sig =
        0.5 * Math.sin(2 * Math.PI * 420 * t) +
        0.3 * Math.sin(2 * Math.PI * 1150 * t) +
        0.2 * Math.sin(2 * Math.PI * 2100 * t);
      samples[i] = Math.round(sig * 12000);
    }
  }
  fs.writeFileSync(pcmPath, Buffer.from(samples.buffer));
}

test('triggerSubsyncFromConfig automatically syncs with built-in alass and strips reference signs/songs', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-auto-builtin-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const refPath = path.join(tmpDir, 'ep01.en.ass');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 900);

  fs.writeFileSync(videoPath, 'video');
  writeSrtFromSpans(targetPath, spans, 3.2);
  writeAssWithSignsFromSpans(refPath, spans, 0);

  let modalOpened = false;
  const osd: string[] = [];
  const sentCommands: Array<Array<string | number>> = [];

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: (payload) => {
          sentCommands.push(payload.command);
        },
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 1000;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': refPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: false,
      }),
      openManualPicker: () => {
        modalOpened = true;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.equal(modalOpened, false);
  const retimedPath = path.join(tmpDir, 'ep01.ja_retimed.srt');
  assert.equal(fs.existsSync(retimedPath), true);
  assert.equal(sentCommands[0]?.[0], 'sub-add');
  assert.equal(sentCommands[0]?.[1], retimedPath);
  assert.deepEqual(sentCommands[1], ['set_property', 'sub-delay', 0]);
  assert.ok(osd.some((line) => line.startsWith('Synced via') && line.includes('-3.20s')));
});

test('triggerSubsyncFromConfig prefers an audio-aligned unloaded sidecar over a misaligned secondary track', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-auto-audio-verify-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const badSecondaryPath = path.join(tmpDir, 'ep01.bad.en.srt');
  const goodSidecarPath = path.join(tmpDir, 'ep01.good.en.srt');
  const pcmPath = path.join(tmpDir, 'speech.pcm');
  const fakeFfmpeg = path.join(tmpDir, 'ffmpeg.sh');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 400);

  fs.writeFileSync(videoPath, 'video');
  // Target is +2.5s off audio; loaded secondary is +5.0s off audio (misaligned!); unloaded sidecar matches audio!
  writeSrtFromSpans(targetPath, spans, 2.5);
  writeSrtFromSpans(badSecondaryPath, spans, 5.0);
  writeSrtFromSpans(goodSidecarPath, spans, 0);
  writeSpeechPcmFile(pcmPath, spans, 420);
  writeExecutableScript(fakeFfmpeg, `#!/bin/sh\ncat "${toShellPath(pcmPath)}"\nexit 0\n`);

  let modalOpened = false;
  const osd: string[] = [];

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: 0,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 420;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': badSecondaryPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: fakeFfmpeg,
        replace: false,
      }),
      openManualPicker: () => {
        modalOpened = true;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.equal(modalOpened, false);
  assert.ok(
    osd.some((line) => line.includes('ep01.good.en.srt') && line.includes('-2.5')),
    `expected sync via aligned sidecar around -2.5s, got: ${osd.join(' | ')}`,
  );
});

test('triggerSubsyncFromConfig falls back to built-in audio speech sync when no reference subtitle exists', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-auto-audio-only-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const pcmPath = path.join(tmpDir, 'speech.pcm');
  const fakeFfmpeg = path.join(tmpDir, 'ffmpeg.sh');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 400);

  fs.writeFileSync(videoPath, 'video');
  writeSrtFromSpans(targetPath, spans, 3.0);
  writeSpeechPcmFile(pcmPath, spans, 420);
  writeExecutableScript(fakeFfmpeg, `#!/bin/sh\ncat "${toShellPath(pcmPath)}"\nexit 0\n`);

  let modalOpened = false;
  const osd: string[] = [];

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: 0,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'duration') return 420;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: fakeFfmpeg,
        replace: false,
      }),
      openManualPicker: () => {
        modalOpened = true;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.equal(modalOpened, false);
  assert.ok(
    osd.some((line) => line.includes('video audio') && line.includes('-3.')),
    `expected audio-only sync around -3s, got: ${osd.join(' | ')}`,
  );
});

test('triggerSubsyncFromConfig opens guided modal with explanation when auto-sync confidence is low', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-auto-low-conf-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'wrong-ep.ja.srt');
  const refPath = path.join(tmpDir, 'ep01.en.srt');
  const spansA = generateSyntheticSpans(7n, 1442695040888963407n, 600);
  const spansB = generateSyntheticSpans(99999n, 0x9e3779b97f4a7c15n, 600);

  fs.writeFileSync(videoPath, 'video');
  writeSrtFromSpans(targetPath, spansB, 0);
  writeSrtFromSpans(refPath, spansA, 0);

  let modalPayload: SubsyncManualPayload | null = null;
  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 620;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': refPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: false,
      }),
      openManualPicker: (payload) => {
        modalPayload = payload;
      },
    }),
  );

  assert.ok(modalPayload);
  const resolved = modalPayload as SubsyncManualPayload;
  assert.match(resolved.guidanceMessage ?? '', /low confidence/i);
  assert.equal(resolved.defaultEngine, 'builtin');
});

test('runSubsyncManual supports built-in file reference and manual offset engine', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-manual-builtin-offset-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const externalRefPath = path.join(tmpDir, 'custom-ref.srt');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 500);

  fs.writeFileSync(videoPath, 'video');
  writeSrtFromSpans(targetPath, spans, 2.0);
  writeSrtFromSpans(externalRefPath, spans, 0);

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              lang: 'jpn',
              'external-filename': targetPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath: '/missing/alass',
      ffsubsyncPath: '/missing/ffsubsync',
      ffmpegPath: '/missing/ffmpeg',
      replace: false,
    }),
  });

  const builtinResult = await runSubsyncManual(
    {
      engine: 'builtin',
      referenceMode: 'file',
      referenceFilePath: externalRefPath,
      stripSigns: true,
      allowSplits: true,
    },
    deps,
  );
  assert.equal(builtinResult.ok, true);
  assert.match(builtinResult.message, /Synced via custom-ref\.srt/);

  const offsetResult = await runSubsyncManual({ engine: 'offset', offsetMs: -1250 }, deps);
  assert.equal(offsetResult.ok, true);
  assert.equal(offsetResult.message, 'Shifted subtitle by -1.25s');
});

test('triggerSubsyncFromConfig extracts multiple embedded tracks and audio in a single ffmpeg pass and skips signs tracks', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-single-pass-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const badEmbeddedPath = path.join(tmpDir, 'bad-embedded.srt');
  const goodEmbeddedPath = path.join(tmpDir, 'good-embedded.srt');
  const pcmPath = path.join(tmpDir, 'speech.pcm');
  const ffmpegCallsLog = path.join(tmpDir, 'ffmpeg-calls.log');
  const fakeFfmpeg = path.join(tmpDir, 'ffmpeg.sh');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 400);

  fs.writeFileSync(videoPath, 'video');
  writeSrtFromSpans(targetPath, spans, 2.2);
  writeSrtFromSpans(badEmbeddedPath, spans, 4.5);
  writeSrtFromSpans(goodEmbeddedPath, spans, 0);
  writeSpeechPcmFile(pcmPath, spans, 420);

  // Mock ffmpeg that records each invocation, writes mapped embedded tracks 0:2 and 0:4, and streams audio PCM to stdout.
  writeExecutableScript(
    fakeFfmpeg,
    `#!/bin/sh
echo "CALL:$*" >> "${toShellPath(ffmpegCallsLog)}"
prev=""
cur_map=""
for arg in "$@"; do
  if [ "$prev" = "-map" ]; then cur_map="$arg"; fi
  case "$arg" in
    *.srt)
      if [ "$cur_map" = "0:2" ]; then cp "${toShellPath(badEmbeddedPath)}" "$arg"; fi
      if [ "$cur_map" = "0:4" ]; then cp "${toShellPath(goodEmbeddedPath)}" "$arg"; fi
      ;;
    pipe:1)
      cat "${toShellPath(pcmPath)}"
      ;;
  esac
  prev="$arg"
done
exit 0
`,
  );

  let modalOpened = false;
  const osd: string[] = [];

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: 1,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 420;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: false,
                lang: 'eng',
                title: 'English (Bad)',
                codec: 'subrip',
                'ff-index': 2,
              },
              {
                id: 3,
                type: 'sub',
                selected: false,
                external: false,
                lang: 'eng',
                title: 'Signs & Songs',
                codec: 'ass',
                'ff-index': 3,
              },
              {
                id: 4,
                type: 'sub',
                selected: false,
                external: false,
                lang: 'eng',
                title: 'English (Full)',
                codec: 'subrip',
                'ff-index': 4,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: fakeFfmpeg,
        replace: false,
      }),
      openManualPicker: () => {
        modalOpened = true;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.equal(modalOpened, false);
  const calls = fs.readFileSync(ffmpegCallsLog, 'utf8').trim().split('\n');
  assert.equal(
    calls.length,
    1,
    `expected single ffmpeg pass, got ${calls.length}: ${calls.join(' | ')}`,
  );
  assert.equal(
    calls[0]!.includes('0:3'),
    false,
    'signs/songs embedded track 0:3 should not be demuxed',
  );
  assert.ok(calls[0]!.includes('0:2') && calls[0]!.includes('0:4') && calls[0]!.includes('0:1'));
  assert.ok(
    osd.some((line) => line.includes('English (Full)') && line.includes('-2.2')),
    `expected sync via aligned embedded track #4 around -2.2s, got: ${osd.join(' | ')}`,
  );
});

test('triggerSubsyncFromConfig verifies external references via fast sampled audio without reading full video', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-sampled-fastpath-'));
  const videoPath = path.join(tmpDir, 'movie.mkv');
  const targetPath = path.join(tmpDir, 'movie.ja.srt');
  const refPath = path.join(tmpDir, 'movie.en.srt');
  const pcmPath = path.join(tmpDir, 'speech.pcm');
  const ffmpegCallsLog = path.join(tmpDir, 'ffmpeg-calls.log');
  const fakeFfmpeg = path.join(tmpDir, 'ffmpeg.sh');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 900);

  fs.writeFileSync(videoPath, 'video');
  writeSrtFromSpans(targetPath, spans, 2.6);
  writeSrtFromSpans(refPath, spans, 0);
  writeSpeechPcmFile(pcmPath, spans, 920);

  // Seek-aware mock ffmpeg: honors -ss and -t by slicing the 8kHz s16le PCM file (16000 bytes/sec).
  writeExecutableScript(
    fakeFfmpeg,
    `#!/bin/sh
echo "CALL:$*" >> "${toShellPath(ffmpegCallsLog)}"
ss=0
dur=0
prev=""
for arg in "$@"; do
  if [ "$prev" = "-ss" ]; then ss="$arg"; fi
  if [ "$prev" = "-t" ]; then dur="$arg"; fi
  prev="$arg"
done
if [ "$dur" -gt 0 ] 2>/dev/null; then
  dd if="${toShellPath(pcmPath)}" bs=16000 skip="$ss" count="$dur" 2>/dev/null
else
  cat "${toShellPath(pcmPath)}"
fi
exit 0
`,
  );

  let modalOpened = false;
  const osd: string[] = [];

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: 0,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 920;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': refPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: fakeFfmpeg,
        replace: false,
      }),
      openManualPicker: () => {
        modalOpened = true;
      },
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.equal(modalOpened, false);
  const calls = fs.readFileSync(ffmpegCallsLog, 'utf8').trim().split('\n');
  assert.ok(
    calls.length >= 2 && calls.every((line) => line.includes('-ss ') && line.includes('-t 60')),
    `expected only sampled -ss/-t 60 calls without full decode, got: ${calls.join(' | ')}`,
  );
  assert.ok(
    osd.some((line) => line.startsWith('Synced via') && line.includes('-2.6')),
    `expected fast-path sync around -2.6s, got: ${osd.join(' | ')}`,
  );
});

test('triggerSubsyncFromConfig discards stale result when playback switches to another video mid-sync', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-stale-session-'));
  const videoA = path.join(tmpDir, 'ep01.mkv');
  const videoB = path.join(tmpDir, 'ep02.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const refPath = path.join(tmpDir, 'ep01.en.srt');
  const spans = generateSyntheticSpans(7n, 1442695040888963407n, 500);

  fs.writeFileSync(videoA, 'videoA');
  fs.writeFileSync(videoB, 'videoB');
  writeSrtFromSpans(targetPath, spans, 2.5);
  writeSrtFromSpans(refPath, spans, 0);
  const originalContent = fs.readFileSync(targetPath, 'utf8');

  let pathCalls = 0;
  const sentCommands: Array<Array<string | number>> = [];
  let modalOpened = false;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: (payload) => {
          sentCommands.push(payload.command);
        },
        requestProperty: async (name: string) => {
          if (name === 'path') {
            pathCalls += 1;
            // First call sees ep01, subsequent staleness check sees ep02.
            return pathCalls === 1 ? videoA : videoB;
          }
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 520;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': refPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: true,
      }),
      openManualPicker: () => {
        modalOpened = true;
      },
    }),
  );

  assert.equal(modalOpened, false);
  assert.deepEqual(sentCommands, []);
  assert.equal(fs.readFileSync(targetPath, 'utf8'), originalContent);
});

test('runSubsyncManual refuses destructive replace when target subtitle encoding is lossy', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-lossy-replace-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'lossy.srt');

  fs.writeFileSync(videoPath, 'video');
  // Valid SRT header + corrupt non-UTF8/non-SJIS/non-EUCJP bytes in dialogue text.
  const corruptBytes = Buffer.concat([
    Buffer.from('1\n00:00:01,000 --> 00:00:02,500\n', 'utf8'),
    Buffer.from([0x80, 0xff, 0xfe, 0x80]),
    Buffer.from('\n', 'utf8'),
  ]);
  fs.writeFileSync(targetPath, corruptBytes);

  const deps = makeDeps({
    getMpvClient: () => ({
      connected: true,
      currentAudioStreamIndex: null,
      send: () => {},
      requestProperty: async (name: string) => {
        if (name === 'path') return videoPath;
        if (name === 'sid') return 1;
        if (name === 'secondary-sid') return null;
        if (name === 'track-list') {
          return [
            {
              id: 1,
              type: 'sub',
              selected: true,
              external: true,
              lang: 'jpn',
              'external-filename': targetPath,
            },
          ];
        }
        return null;
      },
    }),
    getResolvedConfig: () => ({
      alassPath: '/missing/alass',
      ffsubsyncPath: '/missing/ffsubsync',
      ffmpegPath: '/missing/ffmpeg',
      replace: true,
    }),
  });

  const result = await runSubsyncManual({ engine: 'offset', offsetMs: 500 }, deps);
  assert.equal(result.ok, false);
  assert.match(result.message, /losslessly/i);
  assert.deepEqual(fs.readFileSync(targetPath), corruptBytes);
});

test('runSubsyncManual rejects stale submissions when active video changed since picker opened', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-stale-media-'));
  const videoA = path.join(tmpDir, 'ep01.mkv');
  const videoB = path.join(tmpDir, 'ep02.mkv');
  const targetB = path.join(tmpDir, 'ep02.ja.srt');
  const original = '1\n00:00:01,000 --> 00:00:03,000\nこんにちは\n';
  fs.writeFileSync(videoA, 'a');
  fs.writeFileSync(videoB, 'b');
  fs.writeFileSync(targetB, original, 'utf8');

  const result = await runSubsyncManual(
    {
      engine: 'offset',
      mediaPath: videoA,
      targetTrackId: 1,
      offsetMs: 1000,
    },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoB;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetB,
              },
            ];
          }
          return null;
        },
      }),
    }),
  );

  assert.equal(result.ok, false);
  assert.match(result.message, /Active video changed/i);
  assert.equal(fs.readFileSync(targetB, 'utf8'), original);
});

test('triggerSubsyncFromConfig normalizes file:// video URLs as local media', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-file-url-'));
  const videoPath = path.join(tmpDir, 'ep#01.mkv');
  fs.writeFileSync(videoPath, 'video');
  const fileUrl = pathToFileURL(videoPath).href;
  let capturedPayload: SubsyncManualPayload | null = null;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return fileUrl;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [{ id: 1, type: 'sub', selected: true, lang: 'jpn' }];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'manual',
        alassPath: process.execPath,
        ffsubsyncPath: process.execPath,
        ffmpegPath: process.execPath,
      }),
      openManualPicker: (payload) => {
        capturedPayload = payload;
      },
    }),
  );

  assert.ok(capturedPayload);
  const resolved = capturedPayload as SubsyncManualPayload;
  assert.equal(resolved.videoReferenceAvailable, true);
  assert.equal(resolved.ffsubsyncAvailable, true);
  assert.equal(resolved.mediaPath, videoPath);
});

test('runSubsyncManual allows retiming an unchanged inactive target track', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-inactive-target-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const activeSub = path.join(tmpDir, 'active.srt');
  const inactiveSub = path.join(tmpDir, 'inactive.srt');
  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(activeSub, '1\n00:00:01,000 --> 00:00:02,000\nactive\n', 'utf8');
  fs.writeFileSync(inactiveSub, '1\n00:00:05,000 --> 00:00:07,000\ninactive\n', 'utf8');

  const commands: Array<(string | number)[]> = [];
  const result = await runSubsyncManual(
    {
      engine: 'offset',
      mediaPath: videoPath,
      targetTrackId: 3,
      offsetMs: 1000,
    },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: (payload) => {
          commands.push(payload.command);
        },
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': activeSub,
              },
              {
                id: 3,
                type: 'sub',
                selected: false,
                external: true,
                lang: 'jpn',
                'external-filename': inactiveSub,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: true,
      }),
    }),
  );

  assert.equal(result.ok, true);
  assert.ok(fs.readFileSync(inactiveSub, 'utf8').includes('00:00:06,000 --> 00:00:08,000'));
  assert.ok(commands.some((cmd) => cmd[0] === 'sub-add' && cmd[1] === inactiveSub));
});

test('triggerSubsyncFromConfig discards auto-sync when user disables target slot mid-sync', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-disabled-slot-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const refPath = path.join(tmpDir, 'ep01.en.srt');
  fs.writeFileSync(videoPath, 'video');

  const lines: string[] = [];
  const targetLines: string[] = [];
  for (let i = 0; i < 40; i += 1) {
    const start = 10 + i * 5;
    const end = start + 2;
    const fmt = (sec: number) => {
      const m = Math.floor(sec / 60);
      const s = sec % 60;
      return `00:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')},000`;
    };
    lines.push(`${i + 1}\n${fmt(start)} --> ${fmt(end)}\nLine ${i + 1}\n`);
    targetLines.push(`${i + 1}\n${fmt(start + 3)} --> ${fmt(end + 3)}\nセリフ ${i + 1}\n`);
  }
  const originalTarget = targetLines.join('\n');
  fs.writeFileSync(refPath, lines.join('\n'), 'utf8');
  fs.writeFileSync(targetPath, originalTarget, 'utf8');

  let sidCalls = 0;
  const sentCommands: Array<(string | number)[]> = [];
  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: (payload) => {
          sentCommands.push(payload.command);
        },
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') {
            sidCalls += 1;
            return sidCalls === 1 ? 1 : 'no';
          }
          if (name === 'secondary-sid') return 2;
          if (name === 'duration') return 400;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': refPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: true,
      }),
    }),
  );

  assert.deepEqual(sentCommands, []);
  assert.equal(fs.readFileSync(targetPath, 'utf8'), originalTarget);
});

test('triggerSubsyncFromConfig skips early extraction fallback picker when playback switches', async () => {
  let pathCalls = 0;
  let pickerOpened = false;

  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') {
            pathCalls += 1;
            return pathCalls === 1 ? '/tmp/videoA.mkv' : '/tmp/videoB.mkv';
          }
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [{ id: 1, type: 'sub', selected: true, lang: 'jpn' }];
          }
          return null;
        },
      }),
      openManualPicker: () => {
        pickerOpened = true;
      },
    }),
  );

  assert.equal(pickerOpened, false);
});

test('triggerSubsyncFromConfig resets non-zero sub-delay when already in sync without rewriting', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-reset-delay-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ep01.ja.srt');
  const refPath = path.join(tmpDir, 'ep01.en.srt');
  fs.writeFileSync(videoPath, 'video');

  let t = 5;
  const spans: Array<[number, number]> = [];
  for (let i = 0; i < 60; i += 1) {
    const dur = 1.2 + ((i * 7) % 5) * 0.4;
    const gap = 0.6 + ((i * 13) % 7) * 0.7;
    spans.push([t, t + dur]);
    t += dur + gap;
  }
  const fmt = (sec: number) => {
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = Math.floor(sec % 60);
    const ms = Math.round((sec - Math.floor(sec)) * 1000);
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')},${String(ms).padStart(3, '0')}`;
  };
  const srt = spans
    .map(([s, e], i) => `${i + 1}\n${fmt(s)} --> ${fmt(e)}\nLine ${i + 1}\n`)
    .join('\n');
  fs.writeFileSync(refPath, srt, 'utf8');
  fs.writeFileSync(targetPath, srt, 'utf8');
  const mtimeBefore = fs.statSync(targetPath).mtimeMs;

  const sentCommands: Array<(string | number)[]> = [];
  const osd: string[] = [];
  await triggerSubsyncFromConfig(
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: (payload) => {
          sentCommands.push(payload.command);
        },
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return 2;
          if (name === 'sub-delay') return 2.5;
          if (name === 'duration') return 600;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
              {
                id: 2,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'eng',
                'external-filename': refPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        mode: 'auto',
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: true,
      }),
      showMpvOsd: (text) => {
        osd.push(text);
      },
    }),
  );

  assert.deepEqual(sentCommands, [['set_property', 'sub-delay', 0]]);
  assert.equal(fs.statSync(targetPath).mtimeMs, mtimeBefore);
  assert.ok(osd.some((line) => line.includes('Reset subtitle delay to 0s')));
});

test('runSubsyncManual stages external engine output and preserves original when playback switches mid-run', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-external-stage-'));
  const ffsubsyncPath = path.join(tmpDir, 'ffsubsync.sh');
  const ffmpegPath = path.join(tmpDir, 'ffmpeg.sh');
  const alassPath = path.join(tmpDir, 'alass.sh');
  const videoA = path.join(tmpDir, 'ep01.mkv');
  const videoB = path.join(tmpDir, 'ep02.mkv');
  const primaryPath = path.join(tmpDir, 'primary.srt');
  const originalContent = '1\n00:00:01,000 --> 00:00:02,000\noriginal\n';

  fs.writeFileSync(videoA, 'a');
  fs.writeFileSync(videoB, 'b');
  fs.writeFileSync(primaryPath, originalContent, 'utf8');
  writeExecutableScript(ffmpegPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(alassPath, '#!/bin/sh\nexit 0\n');
  writeExecutableScript(
    ffsubsyncPath,
    '#!/bin/sh\nout=""\nprev=""\nfor arg in "$@"; do\n  if [ "$prev" = "-o" ]; then out="$arg"; fi\n  prev="$arg"\ndone\nif [ -n "$out" ]; then printf "overwritten" > "$out"; fi\nexit 0\n',
  );

  let pathCalls = 0;
  const sentCommands: Array<Array<string | number>> = [];
  const result = await runSubsyncManual(
    { engine: 'ffsubsync', mediaPath: videoA },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: (payload) => {
          sentCommands.push(payload.command);
        },
        requestProperty: async (name: string) => {
          if (name === 'path') {
            pathCalls += 1;
            return pathCalls === 1 ? videoA : videoB;
          }
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                'external-filename': primaryPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        alassPath,
        ffsubsyncPath,
        ffmpegPath,
        replace: true,
      }),
    }),
  );

  assert.equal(result.ok, false);
  assert.match(result.message, /Playback changed/i);
  assert.equal(fs.readFileSync(primaryPath, 'utf8'), originalContent);
  assert.deepEqual(sentCommands, []);
  const leftoverStaged = fs
    .readdirSync(tmpDir)
    .filter((name) => name.startsWith('.subminer-subsync-'));
  assert.deepEqual(leftoverStaged, []);
});

test('runSubsyncManual preserves POSIX trailing whitespace and backslash distinctions in mediaPath', async () => {
  if (process.platform === 'win32') return;
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-posix-paths-'));
  const trailingSpaceVideo = path.join(tmpDir, 'episode.mkv ');
  const subPath = path.join(tmpDir, 'episode.srt');
  fs.writeFileSync(trailingSpaceVideo, 'video');
  fs.writeFileSync(subPath, '1\n00:00:01,000 --> 00:00:02,000\nhello\n', 'utf8');

  const okResult = await runSubsyncManual(
    { engine: 'offset', mediaPath: trailingSpaceVideo, targetTrackId: 1, offsetMs: 500 },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return trailingSpaceVideo;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                'external-filename': subPath,
              },
            ];
          }
          return null;
        },
      }),
    }),
  );
  assert.equal(okResult.ok, true);

  const slashVideo = path.join(tmpDir, 'show', 'episode.mkv');
  const backslashVideo = path.join(tmpDir, 'show\\episode.mkv');
  const mismatchResult = await runSubsyncManual(
    { engine: 'offset', mediaPath: backslashVideo, targetTrackId: 1, offsetMs: 500 },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return slashVideo;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                'external-filename': subPath,
              },
            ];
          }
          return null;
        },
      }),
    }),
  );
  assert.equal(mismatchResult.ok, false);
  assert.match(mismatchResult.message, /Active video changed/i);
});

test('runSubsyncManual refuses to write ambiguous Shift_JIS/EUC-JP output even when replace is false', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-ambiguous-noreplace-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const targetPath = path.join(tmpDir, 'ambiguous.srt');
  const retimedPath = path.join(tmpDir, 'ambiguous_retimed.srt');

  fs.writeFileSync(videoPath, 'video');
  const ambiguousBytes = Buffer.concat([
    Buffer.from('1\n00:00:01,000 --> 00:00:02,500\n', 'utf8'),
    Buffer.from([0xa4, 0xa2]),
    Buffer.from('\n', 'utf8'),
  ]);
  fs.writeFileSync(targetPath, ambiguousBytes);

  const result = await runSubsyncManual(
    { engine: 'offset', offsetMs: 500 },
    makeDeps({
      getMpvClient: () => ({
        connected: true,
        currentAudioStreamIndex: null,
        send: () => {},
        requestProperty: async (name: string) => {
          if (name === 'path') return videoPath;
          if (name === 'sid') return 1;
          if (name === 'secondary-sid') return null;
          if (name === 'track-list') {
            return [
              {
                id: 1,
                type: 'sub',
                selected: true,
                external: true,
                lang: 'jpn',
                'external-filename': targetPath,
              },
            ];
          }
          return null;
        },
      }),
      getResolvedConfig: () => ({
        alassPath: '/missing/alass',
        ffsubsyncPath: '/missing/ffsubsync',
        ffmpegPath: '/missing/ffmpeg',
        replace: false,
      }),
    }),
  );

  assert.equal(result.ok, false);
  assert.match(result.message, /losslessly/i);
  assert.deepEqual(fs.readFileSync(targetPath), ambiguousBytes);
  assert.equal(fs.existsSync(retimedPath), false);
});
