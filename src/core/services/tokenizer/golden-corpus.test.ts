import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { tokenizeSubtitle } from '../tokenizer';
import {
  createReplayTokenizerDeps,
  projectGoldenTokens,
  type GoldenFixture,
} from './golden-corpus-harness';

// End-to-end regression corpus: each fixture replays recorded Yomitan backend
// responses and raw MeCab tokens through the real tokenizeSubtitle pipeline
// (scan-token merge, MeCab enrichment, frequency ranks, annotation stage,
// noise suppression) and asserts the final annotated tokens.
//
// Record new fixtures with:
//   bun run record-tokenizer-fixture:electron -- --name <slug> <text>
// See __fixtures__/golden/README.md for the format and options.

const FIXTURE_DIR = path.join(__dirname, '__fixtures__', 'golden');

function loadFixtures(): GoldenFixture[] {
  if (!fs.existsSync(FIXTURE_DIR)) {
    return [];
  }
  return fs
    .readdirSync(FIXTURE_DIR)
    .filter((entry) => entry.endsWith('.json'))
    .sort()
    .map((entry) => {
      const filePath = path.join(FIXTURE_DIR, entry);
      const fixture = JSON.parse(fs.readFileSync(filePath, 'utf8')) as GoldenFixture;
      assert.ok(fixture.name, `${entry}: fixture is missing "name"`);
      assert.equal(
        `${fixture.name}.json`,
        entry,
        `${entry}: fixture "name" must match its file name`,
      );
      return fixture;
    });
}

const fixtures = loadFixtures();

// Fixtures are not copied into dist by tsc; the corpus only runs from src.
test('golden corpus has fixtures', { skip: !fs.existsSync(FIXTURE_DIR) }, () => {
  assert.ok(
    fixtures.length > 0,
    `no golden fixtures found in ${FIXTURE_DIR}; record one with record-tokenizer-fixture:electron`,
  );
});

for (const fixture of fixtures) {
  const label = fixture.issueRefs?.length
    ? `${fixture.name} (${fixture.issueRefs.join(', ')})`
    : fixture.name;
  test(`golden: ${label}`, async () => {
    const deps = createReplayTokenizerDeps(fixture);
    const subtitleData = await tokenizeSubtitle(fixture.input.text, deps);
    assert.deepEqual(
      projectGoldenTokens(subtitleData.tokens),
      fixture.expected.tokens,
      fixture.description
        ? `${fixture.name}: ${fixture.description}`
        : `${fixture.name}: annotated tokens diverged from recorded expectation`,
    );
  });
}
