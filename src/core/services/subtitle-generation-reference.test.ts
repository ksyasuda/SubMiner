import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  loadSubtitleGenerationReference,
  readSubtitleGenerationReferences,
  subtitleGenerationReferences,
} from './subtitle-generation-reference';

const embedded = { type: 'sub', codec: 'ass', 'ff-index': 2, lang: 'eng' };

test('references exclude signs, songs, forced, bitmap and generated tracks even when selected', () => {
  const excluded = [
    'Signs & Songs',
    'SignsSongs',
    'Signs/Songs',
    'S&S',
    'S+S',
    'Forced',
    'Karaoke',
    'OP',
    'ED',
    'English lyrics',
    'Generated Japanese',
  ];
  assert.deepEqual(
    subtitleGenerationReferences([
      ...excluded.map((title) => ({ ...embedded, title, selected: true })),
      { ...embedded, forced: true },
      { ...embedded, codec: 'hdmv_pgs_subtitle' },
      {
        ...embedded,
        title: 'English',
        external: true,
        'external-filename': '/subs/show.en.signs.ass',
      },
      { ...embedded, title: 'English Full' },
    ]).map((reference) => reference.label),
    ['English Full'],
  );
});

test('references rank English dialogue first and resolve loaded external files against mpv cwd', () => {
  const refs = subtitleGenerationReferences(
    [
      { ...embedded, title: 'French Full', lang: 'fra', selected: true },
      { ...embedded, title: 'English' },
      { type: 'sub', external: true, 'external-filename': 'subs/show.en.full.srt' },
      { type: 'sub', external: true, 'external-filename': 'https://example.com/en.srt' },
      { type: 'sub', 'ff-index': -1 },
      null,
    ],
    '/mpv',
  );
  assert.deepEqual(
    refs.map((ref) => ref.label),
    ['show.en.full.srt', 'English', 'French Full'],
  );
  assert.deepEqual(refs[0]?.source, { kind: 'external', path: '/mpv/subs/show.en.full.srt' });
  assert.deepEqual(
    subtitleGenerationReferences([
      { type: 'sub', external: true, 'external-filename': 'relative.en.srt' },
    ]),
    [],
  );
});

test('reference discovery captures primary and secondary subtitle delays', async () => {
  const properties: Record<string, unknown> = {
    'working-directory': '/mpv',
    sid: 1,
    'secondary-sid': 2,
    'sub-delay': 1.5,
    'secondary-sub-delay': -2,
  };
  const refs = await readSubtitleGenerationReferences(
    [
      { ...embedded, id: 1 },
      { ...embedded, id: 2 },
      { ...embedded, id: 3 },
    ],
    async (name) => properties[name],
  );
  assert.deepEqual(
    refs.map((ref) => ref.delaySeconds),
    [1.5, -2, 0],
  );
});

test('reference extraction retries unreadable tracks, restores audio offset and ignores marked lyrics', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'generation-reference-'));
  try {
    const ffmpegPath = path.join(directory, 'ffmpeg');
    await writeFile(
      ffmpegPath,
      `#!${process.execPath}
const args = process.argv.slice(2);
if (args[args.indexOf('-map') + 1] === '0:2') process.exit(1);
require('node:fs').writeFileSync(args.at(-1), '1\\n00:00:10,000 --> 00:00:12,000\\nHello\\n\\n2\\n00:00:20,000 --> 00:00:22,000\\n♪ Song ♪\\n\\n3\\n00:00:30,000 --> 00:00:32,000\\nWorld\\n');
`,
      { mode: 0o755 },
    );
    const refs = subtitleGenerationReferences([{ ...embedded }, { ...embedded, 'ff-index': 3 }]);
    const hints = await loadSubtitleGenerationReference({
      references: refs,
      mediaPath: '/video.mkv',
      ffmpegPath,
      directory,
      audioOffset: 2.5,
    });
    assert.deepEqual(hints, [7.5, 27.5]);
    assert.deepEqual(
      await loadSubtitleGenerationReference({
        references: refs.slice(0, 1),
        mediaPath: '/video.mkv',
        ffmpegPath,
        directory,
        audioOffset: 0,
      }),
      [],
    );
    await assert.rejects(
      loadSubtitleGenerationReference({
        references: refs,
        mediaPath: '/video.mkv',
        ffmpegPath,
        directory,
        audioOffset: 0,
        signal: AbortSignal.abort(),
      }),
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
