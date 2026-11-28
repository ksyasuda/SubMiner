import test from 'node:test';
import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import {
  cacheEmbeddedSubtitleTrack,
  clearAudioSpeechCache,
  detectSpeechSpansFromPcm,
  extractAudioSpeechSpans,
  extractSinglePassRipMedia,
  getCachedAudioSpeechSpans,
  getCachedEmbeddedSubtitles,
  getMediaFileSignature,
  parseRawSileroVadSegments,
} from './audio-speech';
import { parseSubtitleFileContent } from './subtitle-lines';

test('detectSpeechSpansFromPcm detects voice-band bursts separated by quiet background', () => {
  const sampleRate = 8000;
  const totalSec = 20;
  const samples = new Int16Array(sampleRate * totalSec);

  const expectedSpans: Array<[number, number]> = [
    [2.0, 4.0],
    [7.5, 9.5],
    [13.0, 15.5],
  ];

  // Fill quiet ambient floor across whole clip.
  for (let i = 0; i < samples.length; i += 1) {
    const t = i / sampleRate;
    samples[i] = Math.round(Math.sin(2 * Math.PI * 500 * t) * 180);
  }

  // Overlay speech-like multi-frequency formant signal inside expected spans.
  for (const [start, end] of expectedSpans) {
    const sIdx = Math.floor(start * sampleRate);
    const eIdx = Math.floor(end * sampleRate);
    for (let i = sIdx; i < eIdx; i += 1) {
      const t = i / sampleRate;
      const sig =
        0.5 * Math.sin(2 * Math.PI * 420 * t) +
        0.3 * Math.sin(2 * Math.PI * 1150 * t) +
        0.2 * Math.sin(2 * Math.PI * 2100 * t);
      samples[i] = Math.round(sig * 12000);
    }
  }

  const { spans, durationSec } = detectSpeechSpansFromPcm(samples, sampleRate);
  assert.equal(durationSec, totalSec);
  assert.equal(spans.length, expectedSpans.length);
  for (let i = 0; i < expectedSpans.length; i += 1) {
    assert.ok(Math.abs(spans[i]!.start - expectedSpans[i]![0]) <= 0.15);
    assert.ok(Math.abs(spans[i]!.end - expectedSpans[i]![1]) <= 0.25);
  }
});

test('parseRawSileroVadSegments parses fine-grained speech segments from whisper-vad output', () => {
  const stdout = [
    'Detected 2 speech segments:',
    'Speech segment 0: start = 125.00, end = 340.00',
    'Speech segment 1: start = 510.50, end = 690.00',
  ].join('\n');

  assert.deepEqual(parseRawSileroVadSegments(stdout), [
    { start: 1.25, end: 3.4 },
    { start: 5.105, end: 6.9 },
  ]);
});

test('audio speech cache keys distinguish built-in VAD from Silero VAD config', async () => {
  if (process.platform === 'win32') return;
  clearAudioSpeechCache();
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subsync-vad-cache-'));
  const videoPath = path.join(tmpDir, 'ep01.mkv');
  const modelPath = path.join(tmpDir, 'silero.bin');
  const ffmpegScript = path.join(tmpDir, 'ffmpeg.sh');
  const vadScript = path.join(tmpDir, 'whisper-vad.sh');
  fs.writeFileSync(videoPath, 'video');
  fs.writeFileSync(modelPath, 'model');

  // Fake ffmpeg outputs silence for pipe:1 (built-in VAD) and a tiny WAV when given an output file (Silero).
  fs.writeFileSync(
    ffmpegScript,
    `#!/bin/sh
last=""
for arg in "$@"; do last="$arg"; done
if [ "$last" = "pipe:1" ]; then
  dd if=/dev/zero bs=1600 count=1 2>/dev/null
else
  printf 'RIFF....WAVEdata................................................' > "$last"
fi
exit 0
`,
    { mode: 0o755 },
  );
  fs.chmodSync(ffmpegScript, 0o755);

  fs.writeFileSync(
    vadScript,
    `#!/bin/sh
echo "Speech segment 0: start = 100.00, end = 300.00"
exit 0
`,
    { mode: 0o755 },
  );
  fs.chmodSync(vadScript, 0o755);

  try {
    const builtin = await extractAudioSpeechSpans({
      ffmpegPath: ffmpegScript,
      videoPath,
      audioStreamIndex: 1,
    });
    assert.equal(builtin.source, 'builtin-vad');
    assert.equal(builtin.spans.length, 0);
    assert.ok(getCachedAudioSpeechSpans(videoPath, 1) !== null);
    assert.equal(
      getCachedAudioSpeechSpans(videoPath, 1, {
        vadExecutablePath: vadScript,
        vadModelPath: modelPath,
      }),
      null,
    );

    const silero = await extractAudioSpeechSpans({
      ffmpegPath: ffmpegScript,
      videoPath,
      audioStreamIndex: 1,
      vadExecutablePath: vadScript,
      vadModelPath: modelPath,
    });
    assert.equal(silero.source, 'silero-vad');
    assert.deepEqual(silero.spans, [{ start: 1.0, end: 3.0 }]);
    assert.equal(
      getCachedAudioSpeechSpans(videoPath, 1, {
        vadExecutablePath: vadScript,
        vadModelPath: modelPath,
      })?.source,
      'silero-vad',
    );
  } finally {
    clearAudioSpeechCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('cacheEmbeddedSubtitleTrack populates single-pass cache, ignores empty parses, and guards file signature', async () => {
  clearAudioSpeechCache();
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-embedded-cache-'));
  try {
    const videoPath = path.join(tmpDir, 'episode.mkv');
    fs.writeFileSync(videoPath, 'video-v1');
    const sigV1 = getMediaFileSignature(videoPath);
    assert.ok(sigV1 !== null);

    const validParsed = parseSubtitleFileContent(
      '1\n00:00:01,000 --> 00:00:02,500\nHello world\n',
      'track_3.srt',
    );
    const emptyParsed = parseSubtitleFileContent('', 'track_3.srt');

    cacheEmbeddedSubtitleTrack(videoPath, 3, validParsed, sigV1);
    cacheEmbeddedSubtitleTrack(videoPath, 3, emptyParsed, sigV1);

    const cachedV1 = getCachedEmbeddedSubtitles(videoPath);
    assert.equal(cachedV1?.get(3)?.events.length, 1);

    const rip = await extractSinglePassRipMedia({
      ffmpegPath: '/nonexistent/ffmpeg',
      videoPath,
      audioStreamIndex: 1,
      embeddedTracks: [{ ffIndex: 3, extension: 'srt' }],
      includeAudio: false,
    });

    assert.ok(rip !== null);
    assert.equal(rip.embeddedSubtitles.get(3)?.events.length, 1);

    // Replacing the file invalidates sigV1; publishing under sigV1 must not poison v2.
    fs.writeFileSync(videoPath, 'video-v2-remuxed');
    const sigV2 = getMediaFileSignature(videoPath);
    assert.notEqual(sigV2, sigV1);

    cacheEmbeddedSubtitleTrack(videoPath, 4, validParsed, sigV1);
    assert.equal(getCachedEmbeddedSubtitles(videoPath), null);
  } finally {
    clearAudioSpeechCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('extractSinglePassRipMedia skips TTML tracks and does not cache empty tracks', async () => {
  if (process.platform === 'win32') return;
  clearAudioSpeechCache();
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-single-pass-'));
  try {
    const videoPath = path.join(tmpDir, 'episode.mkv');
    const argsLog = path.join(tmpDir, 'args.log');
    const ffmpegScript = path.join(tmpDir, 'ffmpeg.sh');
    fs.writeFileSync(videoPath, 'video');
    // Fake ffmpeg: track_2 gets one cue, track_5 parses to zero cues.
    fs.writeFileSync(
      ffmpegScript,
      `#!/bin/sh
echo "$@" > "${argsLog}"
for arg in "$@"; do
  case "$arg" in
    *track_2.srt) printf '1\\n00:00:01,000 --> 00:00:02,000\\nHello\\n' > "$arg" ;;
    *track_5.srt) : > "$arg" ;;
  esac
done
exit 0
`,
      { mode: 0o755 },
    );
    fs.chmodSync(ffmpegScript, 0o755);

    const rip = await extractSinglePassRipMedia({
      ffmpegPath: ffmpegScript,
      videoPath,
      audioStreamIndex: 1,
      embeddedTracks: [
        { ffIndex: 2, extension: 'srt' },
        { ffIndex: 4, extension: 'ttml' },
        { ffIndex: 5, extension: 'srt' },
      ],
      includeAudio: false,
    });

    assert.ok(rip !== null);
    assert.ok(!fs.readFileSync(argsLog, 'utf8').includes('0:4'));
    assert.equal(rip.embeddedSubtitles.get(2)?.events.length, 1);
    assert.equal(rip.embeddedSubtitles.get(5)?.events.length, 0);
    const cached = getCachedEmbeddedSubtitles(videoPath);
    assert.equal(cached?.get(2)?.events.length, 1);
    assert.equal(cached?.has(5), false);
  } finally {
    clearAudioSpeechCache();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});
