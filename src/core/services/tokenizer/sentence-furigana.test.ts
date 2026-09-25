import assert from 'node:assert/strict';
import test from 'node:test';
import { generateSentenceFurigana } from './sentence-furigana';
import { createDeps, runInjectedYomitanScript } from './yomitan-scan-test-harness';

test('generates sentence readings through the parser runtime bridge', async () => {
  const requests: string[] = [];
  const deps = createDeps((script) =>
    runInjectedYomitanScript(script, (action, params) => {
      requests.push(action);
      if (action === 'optionsGetFull')
        return { profileCurrent: 0, profiles: [{ options: { scanning: { length: 20 } } }] };
      assert.equal(action, 'parseText');
      assert.ok(typeof params === 'object' && params !== null && 'text' in params);
      assert.equal(params.text, '猫がいる。');
      return [
        {
          source: 'scanning-parser',
          content: [[{ text: '猫', reading: 'ねこ' }], [{ text: 'がいる。', reading: '' }]],
        },
      ];
    }),
  );
  assert.equal(
    await generateSentenceFurigana('猫がいる。', '猫', deps, { error: assert.fail }),
    '<b> 猫[ねこ]</b>がいる。',
  );
  assert.ok(requests.includes('parseText'));
});

test(
  'a stalled parser cannot indefinitely block sentence and media updates',
  { timeout: 15_000 },
  async () => {
    const warnings: string[] = [];
    const deps = createDeps(() => new Promise<never>(() => {}));
    assert.equal(
      await generateSentenceFurigana('猫', undefined, deps, {
        error: () => undefined,
        warn: (message) => warnings.push(message),
      }),
      null,
    );
    assert.equal(warnings.length, 1);
  },
);
