import assert from 'node:assert/strict';
import test from 'node:test';

import { parseInlineMarkdown, parsePluginPanelMarkdown } from './plugin-panel-render';

test('parses the block structure of a typical grammar explanation', () => {
  const blocks = parsePluginPanelMarkdown(
    [
      '## Translation',
      'I already told you,',
      "didn't I?",
      '',
      '- **言った** past of 言う',
      '- `じゃない` seeks agreement',
      '',
      '1. first',
      '2. second',
      '',
      '> 言ったじゃない',
      '---',
      '```',
      '<b>kept literal</b>',
      '```',
    ].join('\n'),
  );

  assert.deepEqual(
    blocks.map((block) => block.kind),
    ['heading', 'paragraph', 'list', 'list', 'quote', 'rule', 'code'],
  );
  assert.deepEqual(blocks[1], {
    kind: 'paragraph',
    lines: [
      [{ kind: 'text', value: 'I already told you,' }],
      [{ kind: 'text', value: "didn't I?" }],
    ],
  });
  assert.deepEqual(blocks[2], {
    kind: 'list',
    ordered: false,
    items: [
      [
        { kind: 'strong', value: '言った' },
        { kind: 'text', value: ' past of 言う' },
      ],
      [
        { kind: 'code', value: 'じゃない' },
        { kind: 'text', value: ' seeks agreement' },
      ],
    ],
  });
  assert.deepEqual(blocks[6], { kind: 'code', value: '<b>kept literal</b>' });
});

test('inline parsing keeps HTML and unmatched markers as text', () => {
  assert.deepEqual(parseInlineMarkdown('<img src=x> and *em* and 2 * 3'), [
    { kind: 'text', value: '<img src=x> and ' },
    { kind: 'em', value: 'em' },
    { kind: 'text', value: ' and 2 * 3' },
  ]);
});

test('underscores inside words stay literal while delimited underscore emphasis still works', () => {
  assert.deepEqual(parseInlineMarkdown('set ctx_settings_value and _note_ or __bold__'), [
    { kind: 'text', value: 'set ctx_settings_value and ' },
    { kind: 'em', value: 'note' },
    { kind: 'text', value: ' or ' },
    { kind: 'strong', value: 'bold' },
  ]);
});
