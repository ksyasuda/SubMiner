import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { probeSubtitleGenerationAudio } from './subtitle-generation-probe';
import { selectSubtitleGenerationAlternative } from './subtitle-generation-alternative';
import { prepareSubtitleGenerationAudio } from './subtitle-generation-audio';
import { runSubtitleGenerationProcess } from './subtitle-generation-process';
import { resolveSubtitleGenerationTools } from './subtitle-generation-tools';
import { DEFAULT_SUBTITLE_GENERATION_CONFIG } from '../../shared/subtitle-generation';

test(
  'real FFmpeg verifies matching alternatives, rejects different dialogue and reuses the extract',
  { timeout: 30_000 },
  async (t) => {
    const tools = await resolveSubtitleGenerationTools(DEFAULT_SUBTITLE_GENERATION_CONFIG);
    if (tools.ffmpeg.kind !== 'found' || tools.ffprobe.kind !== 'found') {
      t.skip('Requires FFmpeg');
      return;
    }
    const directory = await mkdtemp(path.join(tmpdir(), 'subtitle-alternative-test-'));
    let requests = 0;
    const requestedFiles: string[] = [];
    const server = createServer(async (req, res) => {
      requests += 1;
      if (req.headers.referer !== 'https://anime.test/') {
        res.writeHead(403).end();
        return;
      }
      const name = req.url?.slice(1) ?? '';
      requestedFiles.push(name);
      if (
        !['original.mkv', 'small.mkv', 'audio.mka', 'dub.mka', 'short.mka', 'shifted.mka'].includes(
          name,
        ) &&
        !/^(?:high|low)(?:\d+\.ts|\.m3u8)$/.test(name)
      ) {
        res.writeHead(404).end();
        return;
      }
      const bytes = await readFile(path.join(directory, name));
      const range = /^bytes=(\d+)-(\d*)$/.exec(req.headers.range ?? '');
      if (range) {
        const start = Number(range[1]);
        const end = range[2] ? Math.min(bytes.length - 1, Number(range[2])) : bytes.length - 1;
        if (start >= bytes.length) {
          res.writeHead(416).end();
          return;
        }
        res.writeHead(206, {
          'Content-Range': `bytes ${start}-${end}/${bytes.length}`,
          'Content-Length': end - start + 1,
          'Accept-Ranges': 'bytes',
        });
        res.end(bytes.subarray(start, end + 1));
      } else
        res.writeHead(200, { 'Content-Length': bytes.length, 'Accept-Ranges': 'bytes' }).end(bytes);
    });
    const ffmpeg = tools.ffmpeg.path;
    try {
      await runSubtitleGenerationProcess({
        command: ffmpeg,
        args: [
          '-v',
          'error',
          '-f',
          'lavfi',
          '-i',
          'testsrc2=size=128x72:rate=10:duration=20',
          '-f',
          'lavfi',
          '-i',
          'aevalsrc=0.2*sin(2*PI*(220+20*t)*t):d=20:s=48000',
          '-c:v',
          'mpeg4',
          '-c:a',
          'pcm_s16le',
          '-metadata:s:a:0',
          'language=jpn',
          path.join(directory, 'original.mkv'),
        ],
      });
      for (const [name, args] of [
        ['audio.mka', ['-vn', '-c:a', 'copy']],
        ['short.mka', ['-vn', '-c:a', 'copy', '-t', '15']],
        ['shifted.mka', ['-vn', '-af', 'adelay=250', '-c:a', 'pcm_s16le', '-t', '20']],
        ['small.mkv', ['-vf', 'scale=64:36', '-c:v', 'mpeg4', '-c:a', 'copy']],
      ] as const) {
        await runSubtitleGenerationProcess({
          command: ffmpeg,
          args: [
            '-v',
            'error',
            '-i',
            path.join(directory, 'original.mkv'),
            ...args,
            path.join(directory, name),
          ],
        });
      }
      await runSubtitleGenerationProcess({
        command: ffmpeg,
        args: [
          '-v',
          'error',
          '-f',
          'lavfi',
          '-i',
          'sine=frequency=750:duration=20',
          '-c:a',
          'pcm_s16le',
          path.join(directory, 'dub.mka'),
        ],
      });
      await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
      const address = server.address();
      assert.ok(address && typeof address === 'object');
      const origin = `http://127.0.0.1:${address.port}`;
      const httpHeaders = { headers: { Referer: 'https://anime.test/' }, userAgent: null };
      const config = {
        ffmpeg,
        ffprobe: tools.ffprobe.path,
        mediaPath: `${origin}/original.mkv`,
        audioStreamIndex: 1,
      };
      const referenceDirectory = path.join(directory, 'reference');
      await mkdir(referenceDirectory);
      const reference = await prepareSubtitleGenerationAudio({
        ...config,
        directory: referenceDirectory,
        mediaPath: path.join(directory, 'original.mkv'),
      });
      // Reproduce an HLS seek that exits successfully without producing audio.
      const emptyOnceFfmpeg = path.join(directory, 'ffmpeg-empty-once');
      const marker = path.join(directory, 'empty-sample-produced');
      await writeFile(
        emptyOnceFfmpeg,
        `#!${process.execPath}
const fs = require('node:fs');
const { spawnSync } = require('node:child_process');
const args = process.argv.slice(2);
const output = args.at(-1);
if (output.endsWith('.pcm') && !fs.existsSync(${JSON.stringify(marker)})) {
  fs.writeFileSync(${JSON.stringify(marker)}, '');
  fs.writeFileSync(output, '');
  process.exit(0);
}
const result = spawnSync(${JSON.stringify(ffmpeg)}, args, {stdio: 'inherit'});
if (result.status === 0 && output.endsWith('.pcm') && fs.statSync(output).size % 4 === 0) fs.appendFileSync(output, Buffer.alloc(2));
process.exit(result.status ?? 1);
`,
        { mode: 0o755 },
      );
      for (const [name, size, timestamp] of [
        ['high', '128:72', '0'],
        ['low', '64:36', '10'],
      ] as const) {
        await runSubtitleGenerationProcess({
          command: ffmpeg,
          args: [
            '-v',
            'error',
            '-i',
            path.join(directory, 'original.mkv'),
            '-vf',
            `scale=${size}`,
            '-c:v',
            'libx264',
            '-g',
            '20',
            '-sc_threshold',
            '0',
            '-c:a',
            'aac',
            '-output_ts_offset',
            timestamp,
            '-f',
            'hls',
            '-hls_time',
            '2',
            '-hls_list_size',
            '0',
            path.join(directory, `${name}.m3u8`),
          ],
        });
      }
      const highProbe = await probeSubtitleGenerationAudio({
        command: tools.ffprobe.path,
        mediaPath: `${origin}/high.m3u8`,
        httpHeaders,
      });
      const lowProbe = await probeSubtitleGenerationAudio({
        command: tools.ffprobe.path,
        mediaPath: `${origin}/low.m3u8`,
        httpHeaders,
      });
      assert.notEqual(
        highProbe.startTime,
        lowProbe.startTime,
        'Fixture must exercise different timestamp origins',
      );
      const hlsDirectory = path.join(directory, 'hls-samples');
      await mkdir(hlsDirectory);
      const hlsResult = await selectSubtitleGenerationAlternative({
        original: { url: `${origin}/high.m3u8`, probe: highProbe, httpHeaders },
        alternatives: [{ kind: 'video', url: `${origin}/low.m3u8`, label: '480p', httpHeaders }],
        ffmpeg,
        ffprobe: tools.ffprobe.path,
        directory: hlsDirectory,
      });
      assert.equal(hlsResult.label, '480p');
      const emptyDirectory = path.join(directory, 'empty-first-sample');
      await mkdir(emptyDirectory);
      const emptyMessages: string[] = [];
      await prepareSubtitleGenerationAudio({
        ...config,
        ffmpeg: emptyOnceFfmpeg,
        directory: emptyDirectory,
        remote: {
          cacheDirectory: path.join(directory, 'empty-cache'),
          httpHeaders,
          alternatives: [
            { kind: 'video', url: `${origin}/small.mkv`, label: 'small.mkv', httpHeaders },
          ],
        },
        onProgress: (progress) => emptyMessages.push(progress.message ?? ''),
      });
      assert.ok(
        emptyMessages.includes('Extracting audio from small.mkv...'),
        emptyMessages.join('\n'),
      );
      const mismatchReasons: string[] = [];
      const originalProbe = await probeSubtitleGenerationAudio({
        command: tools.ffprobe.path,
        mediaPath: config.mediaPath,
        httpHeaders,
      });
      const mismatched = await selectSubtitleGenerationAlternative({
        original: { url: config.mediaPath, probe: originalProbe, httpHeaders },
        alternatives: [{ kind: 'audio', url: `${origin}/dub.mka`, label: 'dub', httpHeaders }],
        ffmpeg: emptyOnceFfmpeg,
        ffprobe: tools.ffprobe.path,
        directory: emptyDirectory,
        onDiagnostic: (event) => mismatchReasons.push(event.reason),
      });
      assert.equal(mismatched.label, 'current audio track');
      assert.deepEqual(
        mismatchReasons,
        ['audio-mismatch'],
        'Odd sample counts must not turn a mismatch into a process error',
      );
      for (const variant of ['audio.mka', 'small.mkv'] as const) {
        const jobDirectory = path.join(directory, variant + '-job');
        const cacheDirectory = path.join(directory, variant + '-cache');
        await mkdir(jobDirectory);
        const messages: string[] = [];
        const result = await prepareSubtitleGenerationAudio({
          ...config,
          directory: jobDirectory,
          remote: {
            cacheDirectory,
            sessionDirectory: cacheDirectory,
            httpHeaders,
            alternatives: [
              { kind: 'audio', url: `${origin}/short.mka`, label: 'wrong duration', httpHeaders },
              {
                kind: 'audio',
                url: `${origin}/shifted.mka`,
                label: 'shifted dialogue',
                httpHeaders,
              },
              { kind: 'audio', url: `${origin}/dub.mka`, label: 'different dialogue', httpHeaders },
              {
                kind: variant === 'audio.mka' ? 'audio' : 'video',
                url: `${origin}/${variant}`,
                label: variant,
                httpHeaders,
              },
            ],
          },
          onProgress: (progress) => messages.push(progress.message ?? ''),
        });
        assert.ok(messages.includes(`Extracting audio from ${variant}...`), messages.join('\n'));
        assert.deepEqual(await readFile(result.wavPath), await readFile(reference.wavPath));
        assert.equal(result.offset, reference.offset);
        const retryDirectory = path.join(directory, variant + '-retry');
        await mkdir(retryDirectory);
        const before = requests;
        const reused = await prepareSubtitleGenerationAudio({
          ...config,
          directory: retryDirectory,
          remote: { cacheDirectory, sessionDirectory: cacheDirectory, httpHeaders },
        });
        assert.equal(requests, before, 'Reuse must not re-probe or download the stream');
        assert.deepEqual(await readFile(reused.wavPath), await readFile(reference.wavPath));
      }
      const fallbackDirectory = path.join(directory, 'fallback');
      await mkdir(fallbackDirectory);
      const messages: string[] = [];
      await prepareSubtitleGenerationAudio({
        ...config,
        directory: fallbackDirectory,
        remote: {
          cacheDirectory: path.join(directory, 'fallback-cache'),
          httpHeaders,
          alternatives: [{ kind: 'audio', url: `${origin}/dub.mka`, label: 'dub', httpHeaders }],
        },
        onProgress: (progress) => messages.push(progress.message ?? ''),
      });
      assert.ok(messages.includes('Extracting audio...'));
      assert.ok(!messages.includes('Extracting audio from dub...'));
      const externalDirectory = path.join(directory, 'external');
      await mkdir(externalDirectory);
      const before = requestedFiles.length;
      const external = await prepareSubtitleGenerationAudio({
        ...config,
        audioStreamIndex: undefined,
        directory: externalDirectory,
        remote: {
          cacheDirectory: path.join(directory, 'external-cache'),
          httpHeaders,
          selectedAudio: {
            url: `${origin}/audio.mka`,
            audioStreamIndex: 0,
            delaySeconds: 0.25,
            httpHeaders,
          },
        },
      });
      assert.equal(external.offset, reference.offset + 0.25);
      assert.ok(requestedFiles.length > before);
      assert.ok(
        requestedFiles.slice(before).every((name) => name === 'audio.mka'),
        'Selected external audio must not download video',
      );
    } finally {
      server.closeAllConnections();
      await new Promise<void>((resolve) => server.close(() => resolve()));
      await rm(directory, { recursive: true, force: true });
    }
  },
);
