import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { EventEmitter } from 'node:events';
import test from 'node:test';
import type { spawn as spawnFn } from 'node:child_process';
import { SOURCE_TYPE_LOCAL } from './types';
import { getLocalVideoMetadata, guessAnimeVideoMetadata, runFfprobe } from './metadata';

type Spawn = typeof spawnFn;

function createSpawnStub(options: {
  stdout?: string;
  stderr?: string;
  emitError?: boolean;
}): Spawn {
  return (() => {
    const child = new EventEmitter() as EventEmitter & {
      stdout: EventEmitter;
      stderr: EventEmitter;
    };
    child.stdout = new EventEmitter();
    child.stderr = new EventEmitter();

    queueMicrotask(() => {
      if (options.emitError) {
        child.emit('error', new Error('ffprobe failed'));
        return;
      }
      if (options.stderr) {
        child.stderr.emit('data', Buffer.from(options.stderr));
      }
      if (options.stdout !== undefined) {
        child.stdout.emit('data', Buffer.from(options.stdout));
      }
      child.emit('close', 0);
    });

    return child as unknown as ReturnType<Spawn>;
  }) as Spawn;
}

test('runFfprobe parses valid JSON from stream and format sections', async () => {
  const metadata = await runFfprobe('/tmp/video.mp4', {
    spawn: createSpawnStub({
      stdout: JSON.stringify({
        format: { duration: '12.34', bit_rate: '3456000' },
        streams: [
          {
            codec_type: 'video',
            codec_tag_string: 'avc1',
            width: 1920,
            height: 1080,
            avg_frame_rate: '24000/1001',
          },
          {
            codec_type: 'audio',
            codec_tag_string: 'mp4a',
          },
        ],
      }),
    }),
  });

  assert.equal(metadata.durationMs, 12340);
  assert.equal(metadata.bitrateKbps, 3456);
  assert.equal(metadata.widthPx, 1920);
  assert.equal(metadata.heightPx, 1080);
  assert.equal(metadata.fpsX100, 2398);
  assert.equal(metadata.containerId, 0);
  assert.ok(Number(metadata.codecId) > 0);
  assert.ok(Number(metadata.audioCodecId) > 0);
});

test('runFfprobe returns empty metadata for invalid JSON and process errors', async () => {
  const invalidJsonMetadata = await runFfprobe('/tmp/broken.mp4', {
    spawn: createSpawnStub({ stdout: '{invalid' }),
  });
  assert.deepEqual(invalidJsonMetadata, {
    durationMs: null,
    codecId: null,
    containerId: null,
    widthPx: null,
    heightPx: null,
    fpsX100: null,
    bitrateKbps: null,
    audioCodecId: null,
  });

  const errorMetadata = await runFfprobe('/tmp/error.mp4', {
    spawn: createSpawnStub({ emitError: true }),
  });
  assert.deepEqual(errorMetadata, {
    durationMs: null,
    codecId: null,
    containerId: null,
    widthPx: null,
    heightPx: null,
    fpsX100: null,
    bitrateKbps: null,
    audioCodecId: null,
  });
});

test('getLocalVideoMetadata derives title and falls back to null hash on read errors', async () => {
  const successMetadata = await getLocalVideoMetadata('/tmp/Episode 01.mkv', {
    spawn: createSpawnStub({ stdout: JSON.stringify({ format: { duration: '0' }, streams: [] }) }),
    fs: {
      createReadStream: () => {
        const stream = new EventEmitter();
        queueMicrotask(() => {
          stream.emit('data', Buffer.from('hello world'));
          stream.emit('end');
        });
        return stream as unknown as ReturnType<typeof import('node:fs').createReadStream>;
      },
      promises: {
        stat: (async () => ({ size: 1234 }) as unknown) as typeof import('node:fs').promises.stat,
      },
    } as never,
  });

  assert.equal(successMetadata.sourceType, SOURCE_TYPE_LOCAL);
  assert.equal(successMetadata.canonicalTitle, 'Episode 01');
  assert.equal(successMetadata.fileSizeBytes, 1234);
  assert.equal(
    successMetadata.hashSha256,
    createHash('sha256').update('hello world').digest('hex'),
  );

  const hashFallbackMetadata = await getLocalVideoMetadata('/tmp/Episode 02.mkv', {
    spawn: createSpawnStub({ stdout: JSON.stringify({ format: {}, streams: [] }) }),
    fs: {
      createReadStream: () => {
        const stream = new EventEmitter();
        queueMicrotask(() => {
          stream.emit('error', new Error('read failed'));
        });
        return stream as unknown as ReturnType<typeof import('node:fs').createReadStream>;
      },
      promises: {
        stat: (async () => ({ size: 5678 }) as unknown) as typeof import('node:fs').promises.stat,
      },
    } as never,
  });

  assert.equal(hashFallbackMetadata.canonicalTitle, 'Episode 02');
  assert.equal(hashFallbackMetadata.hashSha256, null);
});

test('guessAnimeVideoMetadata uses guessit basename output first when available', async () => {
  const seenTargets: string[] = [];
  const parsed = await guessAnimeVideoMetadata(
    '/tmp/Little Witch Academia S02E05.mkv',
    'Episode 5',
    {
      runGuessit: async (target) => {
        seenTargets.push(target);
        return JSON.stringify({
          title: 'Little Witch Academia',
          season: 2,
          episode: 5,
        });
      },
    },
  );

  assert.deepEqual(seenTargets, ['Little Witch Academia S02E05.mkv']);
  assert.deepEqual(parsed, {
    parsedBasename: 'Little Witch Academia S02E05.mkv',
    parsedTitle: 'Little Witch Academia',
    parsedSeason: 2,
    parsedEpisode: 5,
    parserSource: 'guessit',
    parserConfidence: 1,
    parseMetadataJson: JSON.stringify({
      filename: 'Little Witch Academia S02E05.mkv',
      source: 'guessit',
    }),
  });
});

test('guessAnimeVideoMetadata falls back to parser when guessit throws', async () => {
  const parsed = await guessAnimeVideoMetadata(
    '/tmp/Little Witch Academia S02E05.mkv',
    'Episode 5',
    {
      runGuessit: async () => {
        throw new Error('guessit unavailable');
      },
    },
  );

  assert.deepEqual(parsed, {
    parsedBasename: 'Little Witch Academia S02E05.mkv',
    parsedTitle: 'Little Witch Academia',
    parsedSeason: 2,
    parsedEpisode: 5,
    parserSource: 'fallback',
    parserConfidence: 1,
    parseMetadataJson: JSON.stringify({
      confidence: 'high',
      filename: 'Little Witch Academia S02E05.mkv',
      rawTitle: 'Little Witch Academia S02E05',
      source: 'fallback',
    }),
  });
});

test('guessAnimeVideoMetadata falls back when guessit output is incomplete', async () => {
  const parsed = await guessAnimeVideoMetadata('/tmp/[SubsPlease] Frieren - 03 (1080p).mkv', null, {
    runGuessit: async () => JSON.stringify({ episode: 3 }),
  });

  assert.deepEqual(parsed, {
    parsedBasename: '[SubsPlease] Frieren - 03 (1080p).mkv',
    parsedTitle: 'Frieren - 03 (1080p)',
    parsedSeason: null,
    parsedEpisode: null,
    parserSource: 'fallback',
    parserConfidence: 0.2,
    parseMetadataJson: JSON.stringify({
      confidence: 'low',
      filename: '[SubsPlease] Frieren - 03 (1080p).mkv',
      rawTitle: 'Frieren - 03 (1080p)',
      source: 'fallback',
    }),
  });
});
