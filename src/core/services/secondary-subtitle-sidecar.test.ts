import { afterEach, it } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  clearRetimedSecondarySubtitleCache,
  resolveRetimedSecondarySubtitleTextFromSidecar,
  type RetimedSubtitleCommandRunner,
} from './secondary-subtitle-sidecar.js';

const ORIGINAL_ENGLISH = `1
00:00:09,000 --> 00:00:10,000
Stale English subtitle
`;

const ALIGNED_ENGLISH = `1
00:00:01,000 --> 00:00:02,000
Aligned English subtitle
`;

type SidecarFixture = {
  sourcePath: string;
  japanesePath: string;
  englishPath: string;
  alassPath: string;
};

/** Writes an episode with a correctly timed Japanese sidecar and a mistimed English one. */
async function withSidecarFixture(fn: (fixture: SidecarFixture) => Promise<void>): Promise<void> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-secondary-sidecar-test-'));
  try {
    const fixture = {
      sourcePath: path.join(dir, 'episode.mkv'),
      japanesePath: path.join(dir, 'episode.ja.srt'),
      englishPath: path.join(dir, 'episode.en.srt'),
      alassPath: path.join(dir, 'alass-cli'),
    };
    fs.writeFileSync(fixture.sourcePath, 'fake media');
    fs.writeFileSync(fixture.alassPath, 'fake alass');
    fs.writeFileSync(fixture.japanesePath, '1\n00:00:01,000 --> 00:00:02,000\n猫を見た\n');
    fs.writeFileSync(fixture.englishPath, ORIGINAL_ENGLISH);
    await fn(fixture);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

afterEach(() => {
  clearRetimedSecondarySubtitleCache();
});

it('retimes secondary sidecar subtitles against the Japanese sidecar and caches the output', async () => {
  await withSidecarFixture(async ({ sourcePath, japanesePath, englishPath, alassPath }) => {
    let alassRuns = 0;
    const input = { sourcePath, startMs: 1_000, endMs: 2_000, alassPath };

    const first = await resolveRetimedSecondarySubtitleTextFromSidecar({
      ...input,
      runAlass: async (_alassPath, referencePath, inputPath, outputPath) => {
        alassRuns += 1;
        assert.equal(referencePath, japanesePath);
        assert.equal(inputPath, englishPath);
        fs.writeFileSync(outputPath, ALIGNED_ENGLISH);
        return { ok: true, code: 0, stdout: '', stderr: '' };
      },
    });
    const second = await resolveRetimedSecondarySubtitleTextFromSidecar({
      ...input,
      runAlass: async () => {
        alassRuns += 1;
        return { ok: false, code: 1, stdout: '', stderr: 'should use cache' };
      },
    });

    assert.equal(first, 'Aligned English subtitle');
    assert.equal(second, 'Aligned English subtitle');
    assert.equal(alassRuns, 1);
    assert.equal(fs.readFileSync(englishPath, 'utf8'), ORIGINAL_ENGLISH);
  });
});

it('shares in-flight retimed secondary subtitle work for concurrent requests', async () => {
  await withSidecarFixture(async ({ sourcePath, alassPath }) => {
    let alassRuns = 0;
    let releaseAlass!: () => void;
    const alassGate = new Promise<void>((resolve) => {
      releaseAlass = resolve;
    });
    const runAlass: RetimedSubtitleCommandRunner = async (_alass, _reference, _input, output) => {
      alassRuns += 1;
      await alassGate;
      fs.writeFileSync(output, ALIGNED_ENGLISH);
      return { ok: true, code: 0, stdout: '', stderr: '' };
    };
    const input = { sourcePath, startMs: 1_000, endMs: 2_000, alassPath, runAlass };

    const first = resolveRetimedSecondarySubtitleTextFromSidecar(input);
    const second = resolveRetimedSecondarySubtitleTextFromSidecar(input);
    releaseAlass();

    assert.deepEqual(await Promise.all([first, second]), [
      'Aligned English subtitle',
      'Aligned English subtitle',
    ]);
    assert.equal(alassRuns, 1);
  });
});
