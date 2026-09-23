import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { DEFAULT_SUBTITLE_GENERATION_CONFIG } from '../../shared/subtitle-generation';
import { generateJapaneseSubtitles } from './subtitle-generation';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { resolveSubtitleGenerationTools } from './subtitle-generation-tools';

test(
  'real FFmpeg extracts concurrent authenticated HLS downloads in audio order and cleans up',
  {
    skip: process.platform === 'win32' ? 'Requires a POSIX executable fixture.' : false,
  },
  async (t) => {
    const tools = await resolveSubtitleGenerationTools(DEFAULT_SUBTITLE_GENERATION_CONFIG);
    if (tools.ffmpeg.kind !== 'found' || tools.ffprobe.kind !== 'found') {
      t.skip('FFmpeg and ffprobe are required for the network extraction test.');
      return;
    }
    const directory = await mkdtemp(path.join(tmpdir(), 'subminer-network-generation-'));
    const requests: string[] = [];
    let activeSegments = 0;
    let peakSegments = 0;
    const server = createServer(async (req, res) => {
      if (
        req.headers.referer !== 'https://anime.example/' ||
        req.headers['user-agent'] !== 'SubMiner test'
      ) {
        res.writeHead(403).end();
        return;
      }
      const name = req.url?.slice(1) ?? '';
      if (!/^(episode\.m3u8|segment\d+\.ts)$/.test(name)) {
        res.writeHead(404).end();
        return;
      }
      requests.push(name);
      const segment = name.endsWith('.ts');
      if (segment) {
        activeSegments += 1;
        peakSegments = Math.max(peakSegments, activeSegments);
      }
      try {
        if (segment)
          await new Promise((resolve) => setTimeout(resolve, name === 'segment0.ts' ? 80 : 20));
        res.end(await readFile(path.join(directory, name)));
      } catch {
        res.writeHead(404).end();
      } finally {
        if (segment) activeSegments -= 1;
      }
    });
    try {
      await runSubtitleGenerationProcess({
        command: tools.ffmpeg.path,
        args: [
          '-v',
          'error',
          '-f',
          'lavfi',
          '-i',
          'aevalsrc=sin(2*PI*(220+220*floor(t))*t):d=6',
          '-c:a',
          'aac',
          '-f',
          'hls',
          '-hls_time',
          '1',
          '-hls_playlist_type',
          'vod',
          '-hls_segment_filename',
          path.join(directory, 'segment%d.ts'),
          path.join(directory, 'episode.m3u8'),
        ],
      });
      const expectedWav = path.join(directory, 'expected.wav');
      await runSubtitleGenerationProcess({
        command: tools.ffmpeg.path,
        args: [
          '-v',
          'error',
          '-i',
          path.join(directory, 'episode.m3u8'),
          '-map',
          '0:0',
          '-vn',
          '-af',
          'asetpts=PTS-STARTPTS',
          '-ac',
          '1',
          '-ar',
          '16000',
          '-c:a',
          'pcm_s16le',
          expectedWav,
        ],
      });
      await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
      const address = server.address();
      assert.ok(address && typeof address === 'object');
      const modelPath = path.join(directory, 'model.bin');
      const model = Buffer.alloc(8);
      model.writeUInt32LE(0x67676d6c, 0);
      model.writeInt32LE(51865, 4);
      await writeFile(modelPath, model);
      const whisperPath = path.join(directory, 'whisper-test');
      // Recognition is deterministic here; probing, HTTP requests and decoding use real FFmpeg.
      await writeFile(
        whisperPath,
        `#!${process.execPath}
const fs = require('node:fs');
const assert = require('node:assert/strict');
const args = process.argv.slice(2);
const wav = args[args.indexOf('-f') + 1];
const bytes = fs.readFileSync(wav);
assert.equal(bytes.toString('ascii', 0, 4), 'RIFF');
assert.ok(bytes.length > 90000);
assert.deepEqual(bytes, fs.readFileSync(${JSON.stringify(expectedWav)}));
fs.writeFileSync(${JSON.stringify(path.join(directory, 'audio-path'))}, wav);
fs.writeFileSync(args[args.indexOf('-of') + 1] + '.srt', '1\\n00:00:00,500 --> 00:00:01,500\\nこんにちは\\n');
`,
        { mode: 0o755 },
      );
      const cacheDirectory = path.join(directory, 'cache');
      const input = {
        config: {
          ...DEFAULT_SUBTITLE_GENERATION_CONFIG,
          modelPath,
          whisperPath,
          ffmpegPath: tools.ffmpeg.path,
          ffprobePath: tools.ffprobe.path,
        },
        modelDirectory: directory,
        mediaPath: `http://127.0.0.1:${address.port}/episode.m3u8`,
        audioStreamIndex: 0,
        remote: {
          cacheDirectory,
          sessionDirectory: cacheDirectory,
          httpHeaders: {
            headers: { Referer: 'https://anime.example/' },
            userAgent: 'SubMiner test',
          },
        },
      };
      const failedWhisper = path.join(directory, 'whisper-failure');
      await writeFile(failedWhisper, `#!${process.execPath}\nprocess.exit(1);\n`, { mode: 0o755 });
      await assert.rejects(
        generateJapaneseSubtitles({
          ...input,
          config: { ...input.config, whisperPath: failedWhisper },
        }),
        /status 1/,
      );
      assert.ok(requests.includes('episode.m3u8'));
      assert.ok(requests.includes('segment2.ts'));
      assert.ok(peakSegments >= 3, `Expected concurrent downloads, saw ${peakSegments}`);
      assert.ok(peakSegments <= 4, `Download limit exceeded: ${peakSegments}`);
      requests.length = 0;
      const output = await generateJapaneseSubtitles(input);
      assert.deepEqual(requests, [], 'A Whisper failure must not discard completed audio');
      assert.match(await readFile(output, 'utf8'), /00:00:00,500 --> 00:00:01,500\nこんにちは/);
      const wav = await readFile(path.join(directory, 'audio-path'), 'utf8');
      await assert.rejects(readFile(wav), /ENOENT/);
      requests.length = 0;
      const progress: string[] = [];
      const secondOutput = await generateJapaneseSubtitles({
        ...input,
        config: { ...input.config, managedModel: 'base' },
        onProgress: (update) => progress.push(update.message ?? ''),
      });
      assert.notEqual(output, secondOutput);
      assert.deepEqual(requests, []);
      assert.ok(progress.includes('Reusing downloaded audio...'));
      await assert.rejects(
        generateJapaneseSubtitles({
          ...input,
          remote: { ...input.remote, httpHeaders: { headers: {}, userAgent: null } },
        }),
        /403/,
      );
      assert.deepEqual(
        (await readdir(cacheDirectory)).sort(),
        ['audio', path.basename(output), path.basename(secondOutput)].sort(),
      );
    } finally {
      server.closeAllConnections();
      await new Promise<void>((resolve, reject) =>
        server.close((error) =>
          error && (!('code' in error) || error.code !== 'ERR_SERVER_NOT_RUNNING')
            ? reject(error)
            : resolve(),
        ),
      );
      await rm(directory, { recursive: true, force: true });
    }
  },
);
