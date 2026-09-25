import assert from 'node:assert/strict';
import test from 'node:test';
import { formatSentenceFurigana } from './sentence-furigana';

const parsed = [
  {
    source: 'scanning-parser',
    content: [
      [{ text: '猫', reading: 'ねこ' }],
      [{ text: 'を' }],
      [{ text: '見', reading: 'み' }, { text: 'た' }],
      [{ text: '。\n' }],
      [{ text: '犬', reading: 'いぬ' }],
      [{ text: 'もいた。' }],
    ],
  },
];

test('formats the complete expanded sentence with readings and the mined word highlighted', () => {
  assert.equal(
    formatSentenceFurigana('猫を見た。\n犬もいた。', parsed, '猫'),
    '<b> 猫[ねこ]</b>を 見[み]た。\n 犬[いぬ]もいた。',
  );
});

test('rejects headword-only and malformed parse results instead of saving partial furigana', () => {
  assert.equal(
    formatSentenceFurigana('猫を見た。', [
      { source: 'scanning-parser', content: [[{ text: '猫', reading: 'ねこ' }]] },
    ]),
    null,
  );
  assert.equal(
    formatSentenceFurigana('猫', [
      { source: 'scanning-parser', content: [[{ text: '猫', reading: 42 }]] },
    ]),
    null,
  );
});

test('escapes literal markup and annotation delimiters and leaves kana unannotated', () => {
  const text = '<猫> [メモ]';
  assert.equal(
    formatSentenceFurigana(text, [
      {
        source: 'scanning-parser',
        content: [
          [{ text: '<' }],
          [{ text: '猫', reading: 'ねこ' }],
          [{ text: '> [' }],
          [{ text: 'メモ', reading: 'めも' }],
          [{ text: ']' }],
        ],
      },
    ]),
    '&lt; 猫[ねこ]&gt; &#91;メモ&#93;',
  );
});

test('highlights the mined word without bolding the rest of a dictionary phrase', () => {
  assert.equal(
    formatSentenceFurigana(
      '行儀を直して',
      [
        {
          content: [
            [
              { text: '行儀', reading: 'ぎょうぎ' },
              { text: 'を' },
              { text: '直', reading: 'なお' },
              { text: 'して' },
            ],
          ],
        },
      ],
      '行儀',
    ),
    '<b> 行儀[ぎょうぎ]</b>を 直[なお]して',
  );
});
