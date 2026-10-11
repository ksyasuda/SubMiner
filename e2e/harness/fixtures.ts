import { execFile } from 'node:child_process';
import fs from 'node:fs';
import { promisify } from 'node:util';
import { writeStoredZipAsync } from '../../src/shared/stored-zip';

const execFileAsync = promisify(execFile);

export const FIXTURE_DICTIONARY_TITLE = 'SubMiner E2E Fixture';
export const FIXTURE_CLIP_SECONDS = 15;

/** Cues of the synthetic clip. Seek to `at` to land inside a cue. */
export const FIXTURE_CUES = [
  { start: 1, end: 4, at: 2, text: '今日はいい天気ですね。' },
  { start: 5, end: 8, at: 6, text: '猫が好きです。' },
  { start: 9, end: 12, at: 10, text: '明日は学校に行きます。' },
] as const;

// [term, reading, part of speech, gloss] for every content word in FIXTURE_CUES.
const FIXTURE_TERMS = [
  ['今日', 'きょう', 'n', 'today'],
  ['いい', 'いい', 'adj-i', 'good'],
  ['天気', 'てんき', 'n', 'weather'],
  ['猫', 'ねこ', 'n', 'cat'],
  ['好き', 'すき', 'adj-na', 'liked'],
  ['明日', 'あした', 'n', 'tomorrow'],
  ['学校', 'がっこう', 'n', 'school'],
  ['行く', 'いく', 'v5', 'to go'],
] as const;

function srtTimestamp(seconds: number): string {
  return `00:00:${String(seconds).padStart(2, '0')},000`;
}

export function writeFixtureSubtitles(outputPath: string): void {
  const blocks = FIXTURE_CUES.map(
    (cue, index) =>
      `${index + 1}\n${srtTimestamp(cue.start)} --> ${srtTimestamp(cue.end)}\n${cue.text}\n`,
  );
  fs.writeFileSync(outputPath, blocks.join('\n'), 'utf8');
}

/** Renders a deterministic test-pattern clip with a sine tone, so no real media is needed. */
export async function generateFixtureClip(outputPath: string): Promise<void> {
  await execFileAsync('ffmpeg', [
    '-nostdin',
    '-v',
    'error',
    '-y',
    '-f',
    'lavfi',
    '-i',
    `testsrc2=size=640x360:rate=24:duration=${FIXTURE_CLIP_SECONDS}`,
    '-f',
    'lavfi',
    '-i',
    `sine=frequency=440:duration=${FIXTURE_CLIP_SECONDS}`,
    '-c:v',
    'libx264',
    '-preset',
    'ultrafast',
    '-pix_fmt',
    'yuv420p',
    '-c:a',
    'aac',
    outputPath,
  ]);
}

/** Writes a minimal Yomitan term dictionary so the tokenizer has real entries to match. */
export async function buildFixtureDictionary(outputPath: string): Promise<void> {
  const index = {
    title: FIXTURE_DICTIONARY_TITLE,
    revision: '1',
    format: 3,
    author: 'SubMiner',
    description: 'Synthetic dictionary for the e2e harness.',
  };
  const termBank = FIXTURE_TERMS.map(([term, reading, partOfSpeech, gloss], sequence) => [
    term,
    reading,
    '',
    partOfSpeech,
    0,
    [gloss],
    sequence + 1,
    '',
  ]);
  await writeStoredZipAsync(outputPath, [
    { name: 'index.json', data: Buffer.from(JSON.stringify(index), 'utf8') },
    { name: 'term_bank_1.json', data: Buffer.from(JSON.stringify(termBank), 'utf8') },
  ]);
}
