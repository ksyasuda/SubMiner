import assert from 'node:assert/strict';
import test from 'node:test';
import { applyCollapsibleOpenStatesToTermEntries, buildSnapshotFromCharacters } from './build';
import type { CharacterDictionaryTermEntry, CharacterRecord } from './types';

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

test('buildSnapshotFromCharacters shows Japanese aliases without adding romanized names as lookup entries', () => {
  const character: CharacterRecord = {
    id: 1,
    role: 'main',
    firstNameHint: '',
    fullName: 'Aqua',
    lastNameHint: '',
    nativeName: 'アクア',
    alternativeNames: ['阿久亜'],
    bloodType: '',
    birthday: null,
    description: '',
    imageUrl: null,
    age: '',
    sex: '',
    voiceActors: [],
  };

  const snapshot = buildSnapshotFromCharacters(
    100,
    'KonoSuba',
    [character],
    new Map(),
    new Map(),
    1_700_000_000_000,
    () => false,
  );

  const aquaEntry = snapshot.termEntries.find(([term]) => term === 'アクア');
  assert.ok(aquaEntry);
  const glossaryEntry = aquaEntry[5][0] as {
    content: {
      content: Array<{ content?: unknown }>;
    };
  };
  const wholeGlossary = JSON.stringify(glossaryEntry);

  const knownNames = glossaryEntry.content.content.find((node) => {
    const content = node.content;
    return (
      Array.isArray(content) &&
      content.some(
        (child) =>
          child &&
          typeof child === 'object' &&
          (child as { content?: unknown }).content === 'Known names',
      )
    );
  }) as { content: Array<{ content?: unknown }> } | undefined;
  assert.ok(knownNames, 'expected a Known names block in the character glossary');
  const knownNameItems = JSON.stringify(knownNames.content);
  const terms = snapshot.termEntries.map(([term]) => term);

  assert.match(knownNameItems, /アクア/);
  assert.match(knownNameItems, /阿久亜/);
  assert.doesNotMatch(wholeGlossary, /Aqua/);
  assert.doesNotMatch(knownNameItems, /Aqua/);
  assert.doesNotMatch(knownNameItems, /アクア様/);
  assert.equal(terms.includes('Aqua'), false);
  assert.equal(terms.includes('アクア'), true);
  assert.equal(terms.includes('阿久亜'), true);
});

test('buildSnapshotFromCharacters stores media id in Yomitan structured-content data', () => {
  const character: CharacterRecord = {
    id: 1,
    role: 'main',
    firstNameHint: '',
    fullName: 'Aqua',
    lastNameHint: '',
    nativeName: 'アクア',
    alternativeNames: [],
    bloodType: '',
    birthday: null,
    description: '',
    imageUrl: null,
    age: '',
    sex: '',
    voiceActors: [],
  };

  const snapshot = buildSnapshotFromCharacters(
    21699,
    "KONOSUBA -God's blessing on this wonderful world! 2",
    [character],
    new Map(),
    new Map(),
    1_700_000_000_000,
    () => false,
  );
  const aquaEntry = snapshot.termEntries.find(([term]) => term === 'アクア');
  assert.ok(aquaEntry);
  const glossaryEntry = aquaEntry[5][0] as {
    content: {
      data?: Record<string, string>;
      content: Array<Record<string, unknown>>;
    };
  };

  assert.equal(glossaryEntry.content.data?.subminerMediaId, '21699');
  assert.equal(
    glossaryEntry.content.content.some((node) =>
      Object.prototype.hasOwnProperty.call(node, 'subminerMediaId'),
    ),
    false,
  );
});
