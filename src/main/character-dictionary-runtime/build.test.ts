import assert from 'node:assert/strict';
import test from 'node:test';
import { applyCollapsibleOpenStatesToTermEntries } from './build';
import type { CharacterDictionaryTermEntry } from './types';

test('applyCollapsibleOpenStatesToTermEntries reapplies configured details open states', () => {
  const termEntries: CharacterDictionaryTermEntry[] = [
    [
      'アルファ',
      'あるふぁ',
      '',
      '',
      0,
      [
        {
          type: 'structured-content',
          content: {
            tag: 'div',
            content: [
              {
                tag: 'details',
                open: false,
                content: [
                  { tag: 'summary', content: 'Description' },
                  { tag: 'div', content: 'body' },
                ],
              },
              {
                tag: 'details',
                open: false,
                content: [
                  { tag: 'summary', content: 'Voiced by' },
                  { tag: 'div', content: 'cv' },
                ],
              },
            ],
          },
        },
      ],
      0,
      'name',
    ],
  ];

  const [entry] = applyCollapsibleOpenStatesToTermEntries(
    termEntries,
    (section) => section === 'description',
  );
  assert.ok(entry);
  const glossaryEntry = entry[5][0] as {
    content: {
      content: Array<{ open?: boolean }>;
    };
  };

  assert.equal(glossaryEntry.content.content[0]?.open, true);
  assert.equal(glossaryEntry.content.content[1]?.open, false);
});
