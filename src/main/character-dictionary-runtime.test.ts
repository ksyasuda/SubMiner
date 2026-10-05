import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as path from 'path';
import test from 'node:test';

import { buildSnapshotFromCharacters } from './character-dictionary-runtime/build';
import { getSnapshotPath } from './character-dictionary-runtime/cache';
import {
  EMINENCE_MEDIA_ID,
  EMINENCE_TITLE,
  createRuntime,
  findSection,
  getGlossaryChildren,
  installFakeAniList,
  readStoredZipEntry,
  readTermBank,
  type FakeAniListCharacterEdge,
} from './character-dictionary-runtime/runtime-test-harness';
import type {
  CharacterDictionarySnapshot,
  CharacterRecord,
} from './character-dictionary-runtime/types';

const ALEXIA_IMAGE_URL = 'https://cdn.example.com/character-123.png';
const VA_IMAGE_URL = 'https://cdn.example.com/va-456.jpg';

function eminenceMedia(characters: FakeAniListCharacterEdge[]) {
  return { id: EMINENCE_MEDIA_ID, episodes: 20, title: EMINENCE_TITLE, characters };
}

const alexiaWithVoiceActor: FakeAniListCharacterEdge = {
  role: 'SUPPORTING',
  voiceActors: [
    {
      id: 456,
      name: { full: 'Rina Hidaka', native: '日高里菜' },
      image: { medium: VA_IMAGE_URL },
    },
  ],
  node: {
    id: 123,
    description:
      'Alexia Midgar is the second princess of the Kingdom of Midgar.\n\n__Race:__ Human',
    image: { large: ALEXIA_IMAGE_URL, medium: null },
    name: { full: 'Alexia Midgar', native: 'アレクシア・ミドガル' },
  },
};

test('generateForCurrentMedia writes structured-content glossary entries from AniList characters', async (t) => {
  const anilist = installFakeAniList(t, {
    media: [
      eminenceMedia([
        {
          role: 'SUPPORTING',
          node: {
            id: 123,
            description:
              '__Race:__ Human\nAlexia Midgar is the second princess of the Kingdom of Midgar.',
            image: { large: ALEXIA_IMAGE_URL, medium: null },
            gender: 'Female',
            age: '15',
            dateOfBirth: { month: 9, day: 1 },
            bloodType: 'A',
            name: {
              full: 'Alexia Midgar',
              native: 'アレクシア・ミドガル',
              first: 'Alexia',
              last: 'Midgar',
            },
          },
        },
        {
          role: 'SUPPORTING',
          node: {
            id: 222,
            description: 'Missing native name.',
            image: { large: 'https://example.com/john.png', medium: null },
            name: { full: 'John Smith', native: '', first: 'John', last: 'Smith' },
          },
        },
      ]),
    ],
    images: [ALEXIA_IMAGE_URL],
  });
  const { runtime } = createRuntime();

  const result = await runtime.generateForCurrentMedia();
  const termBank = readTermBank(result.zipPath);

  // Characters without a native name are skipped (no portrait download, no terms) while the rest
  // still generate.
  assert.equal(anilist.requests.includes('https://example.com/john.png'), false);
  assert.equal(
    termBank.some((entry) => JSON.stringify(entry).includes('Missing native name.')),
    false,
  );

  const alexia = termBank.find(([term]) => term === 'アレクシア');
  assert.ok(alexia, 'expected compact native-name variant for character');
  const glossary = alexia[5];
  assert.equal(glossary.length, 1);
  assert.equal((glossary[0] as { type?: string }).type, 'structured-content');
  assert.equal(
    glossary.some((item) => typeof item === 'object' && item.type === 'image'),
    false,
    'image must live inside the structured content, not as a separate glossary item',
  );

  const children = getGlossaryChildren(alexia);
  assert.equal(children[0]?.tag, 'div');
  assert.equal(children[0]?.content, 'アレクシア・ミドガル');
  assert.equal(
    children.some((child) => child.content === 'Alexia Midgar'),
    false,
  );

  const imageWrap = children.find(
    (child) => (child.content as { path?: unknown } | undefined)?.path === 'img/m130298-c123.png',
  );
  assert.ok(imageWrap);
  assert.equal(imageWrap.tag, 'div');
  const image = imageWrap.content as Record<string, unknown>;
  assert.equal(image.tag, 'img');
  assert.equal(image.sizeUnits, 'em');

  const sourceDiv = children.find(
    (child) =>
      typeof child.content === 'string' && child.content.includes('The Eminence in Shadow'),
  );
  assert.equal(sourceDiv?.tag, 'div');

  const roleBadgeDiv = children.find(
    (child) => (child.content as { content?: unknown } | undefined)?.content === 'Main Character',
  );
  assert.ok(roleBadgeDiv);
  assert.equal(roleBadgeDiv.tag, 'div');
  assert.equal((roleBadgeDiv.content as { tag?: string }).tag, 'span');

  const descSection = findSection(children, 'Description');
  assert.ok(descSection, 'expected Description collapsible section');
  assert.equal(descSection.open, false);
  assert.ok(
    String(descSection.content[1]?.content).includes(
      'Alexia Midgar is the second princess of the Kingdom of Midgar.',
    ),
  );

  const infoSection = findSection(children, 'Character Information');
  assert.ok(infoSection, 'expected Character Information section with parsed __Race:__ field');
  assert.equal(infoSection.open, false);
  const info = JSON.stringify(infoSection.content[1]?.content);
  assert.match(info, /Female/);
  assert.match(info, /15 years/);
  assert.match(info, /Blood Type A/);
  assert.match(info, /Birthday: September 1/);
});

test('generateForCurrentMedia applies configured open states to character dictionary sections', async (t) => {
  installFakeAniList(t, {
    media: [eminenceMedia([alexiaWithVoiceActor])],
    images: [ALEXIA_IMAGE_URL, VA_IMAGE_URL],
  });
  const { runtime } = createRuntime({
    getCollapsibleSectionOpenState: (section) =>
      section === 'description' || section === 'voicedBy',
  });

  const result = await runtime.generateForCurrentMedia();
  const alexia = readTermBank(result.zipPath).find(([term]) => term === 'アレクシア');
  assert.ok(alexia);
  const children = getGlossaryChildren(alexia);

  assert.equal(findSection(children, 'Description')?.open, true);
  assert.equal(findSection(children, 'Character Information')?.open, false);
  assert.equal(findSection(children, 'Voiced by')?.open, true);
});

test('generateForCurrentMedia reapplies collapsible open states when using cached snapshot data', async (t) => {
  installFakeAniList(t, {
    media: [eminenceMedia([alexiaWithVoiceActor])],
    images: [ALEXIA_IMAGE_URL, VA_IMAGE_URL],
  });
  const { userDataPath, runtime: runtimeOpen } = createRuntime({
    getCollapsibleSectionOpenState: () => true,
  });
  await runtimeOpen.generateForCurrentMedia();

  const { runtime: runtimeClosed } = createRuntime({
    userDataPath,
    getCollapsibleSectionOpenState: () => false,
    now: () => 1_700_000_000_500,
  });
  const result = await runtimeClosed.generateForCurrentMedia();

  const alexia = readTermBank(result.zipPath).find(([term]) => term === 'アレクシア');
  assert.ok(alexia);
  const sections = getGlossaryChildren(alexia).filter((child) => child.tag === 'details');
  assert.ok(sections.length >= 2);
  assert.ok(sections.every((section) => section.open === false));
});

function makeCharacter(overrides: Partial<CharacterRecord> & { id: number }): CharacterRecord {
  return {
    role: 'main',
    firstNameHint: '',
    fullName: '',
    lastNameHint: '',
    nativeName: '',
    alternativeNames: [],
    bloodType: '',
    birthday: null,
    description: '',
    imageUrl: null,
    age: '',
    sex: '',
    voiceActors: [],
    ...overrides,
  };
}

function buildSnapshot(characters: CharacterRecord[]): CharacterDictionarySnapshot {
  return buildSnapshotFromCharacters(
    1,
    'Test Anime',
    characters,
    new Map(),
    new Map(),
    1_700_000_000_000,
    () => false,
  );
}

const nameTermCases: Array<{
  name: string;
  character: CharacterRecord;
  readings: Array<[term: string, reading: string]>;
}> = [
  {
    name: 'adds kana aliases for romanized names when native name is kanji',
    character: makeCharacter({ id: 1, fullName: 'Satou Kazuma', nativeName: '佐藤和真' }),
    readings: [
      ['カズマ', 'かずま'],
      ['サトウカズマ', 'さとうかずま'],
    ],
  },
  {
    name: 'indexes kanji family and given names using romanized first and last hints',
    character: makeCharacter({
      id: 77,
      firstNameHint: 'Yuuma',
      fullName: 'Yuuma Kunimi',
      lastNameHint: 'Kunimi',
      nativeName: '国見佑真',
    }),
    readings: [
      ['国見', 'くにみ'],
      ['佑真', 'ゆうま'],
    ],
  },
  {
    name: 'indexes alternative character names for alias lookups',
    character: makeCharacter({
      id: 321,
      fullName: 'Cid Kagenou',
      nativeName: 'シド・カゲノー',
      alternativeNames: ['Shadow', 'Minoru Kagenou'],
    }),
    readings: [['シャドウ', 'しゃどう']],
  },
  {
    name: 'uses kanji first and last name hints to build kanji readings',
    character: makeCharacter({
      id: 1,
      firstNameHint: '和真',
      fullName: 'Satou Kazuma',
      lastNameHint: '佐藤',
      nativeName: '佐藤和真',
    }),
    readings: [
      ['佐藤和真', 'さとうかずま'],
      ['佐藤', 'さとう'],
      ['和真', 'かずま'],
    ],
  },
];

for (const { name, character, readings } of nameTermCases) {
  test(`buildSnapshotFromCharacters ${name}`, () => {
    const { termEntries } = buildSnapshot([character]);
    for (const [term, reading] of readings) {
      assert.equal(
        termEntries.find(([entryTerm]) => entryTerm === term)?.[1],
        reading,
        `expected ${term} with reading ${reading}`,
      );
    }
  });
}

test('buildSnapshotFromCharacters preserves duplicate surface forms across different characters', () => {
  const { termEntries } = buildSnapshot([
    makeCharacter({
      id: 111,
      description: 'First Alpha.',
      firstNameHint: 'Alpha',
      fullName: 'Alpha One',
      lastNameHint: 'One',
      nativeName: 'アルファ',
    }),
    makeCharacter({
      id: 222,
      description: 'Second Alpha.',
      firstNameHint: 'Alpha',
      fullName: 'Alpha Two',
      lastNameHint: 'Two',
      nativeName: 'アルファ',
    }),
  ]);

  const glossaries = termEntries
    .filter(([term]) => term === 'アルファ')
    .map((entry) => JSON.stringify(getGlossaryChildren(entry)));
  assert.equal(glossaries.length, 2);
  assert.ok(glossaries.some((value) => value.includes('First Alpha.')));
  assert.ok(glossaries.some((value) => value.includes('Second Alpha.')));
});

const alphaEdge: FakeAniListCharacterEdge = {
  role: 'MAIN',
  node: {
    id: 111,
    description: 'Leader of Shadow Garden.',
    image: { large: 'https://example.com/alpha.png', medium: null },
    name: { full: 'Alpha', native: 'アルファ' },
  },
};

test('getOrCreateCurrentSnapshot reuses cached media resolution without AniList requests', async (t) => {
  const anilist = installFakeAniList(t, {
    media: [eminenceMedia([alphaEdge])],
    images: ['https://example.com/alpha.png'],
  });
  const { runtime, outputDir } = createRuntime();

  const first = await runtime.getOrCreateCurrentSnapshot();
  fs.rmSync(path.join(outputDir, 'anilist-resolution-cache.json'), { force: true });
  const second = await runtime.getOrCreateCurrentSnapshot();

  assert.equal(first.fromCache, false);
  assert.equal(second.fromCache, true);
  assert.equal(anilist.searchCount, 1);
  assert.equal(anilist.characterPageCount, 1);
  assert.equal(fs.existsSync(path.join(outputDir, 'cache.json')), false);

  const snapshot = JSON.parse(
    fs.readFileSync(getSnapshotPath(outputDir, EMINENCE_MEDIA_ID), 'utf8'),
  ) as CharacterDictionarySnapshot;
  assert.equal(snapshot.mediaId, EMINENCE_MEDIA_ID);
  assert.ok(snapshot.termEntries.some(([term]) => term === 'アルファ'));
});

test('generateForCurrentMedia downloads shared voice actor images once per AniList person id', async (t) => {
  const kana = {
    id: 9001,
    name: { full: 'Kana Hanazawa', native: '花澤香菜' },
    image: { large: null, medium: 'https://example.com/kana.png' },
  };
  const anilist = installFakeAniList(t, {
    media: [
      eminenceMedia([
        { ...alphaEdge, voiceActors: [kana] },
        {
          role: 'SUPPORTING',
          voiceActors: [kana],
          node: {
            id: 654,
            description: 'Beta documents Shadow Garden operations.',
            image: { large: 'https://example.com/beta.png', medium: null },
            name: { full: 'Beta', native: 'ベータ' },
          },
        },
      ]),
    ],
    images: [
      'https://example.com/alpha.png',
      'https://example.com/beta.png',
      'https://example.com/kana.png',
    ],
  });
  const { runtime } = createRuntime();

  await runtime.generateForCurrentMedia();

  assert.deepEqual([...anilist.imageRequests].sort(), [
    'https://example.com/alpha.png',
    'https://example.com/beta.png',
    'https://example.com/kana.png',
  ]);
});

test('buildMergedDictionary combines stored snapshots into one stable dictionary', async (t) => {
  const current = { title: 'The Eminence in Shadow', episode: 5 };
  installFakeAniList(t, {
    media: [
      { ...eminenceMedia([alphaEdge]), search: 'The Eminence in Shadow' },
      {
        id: 21,
        episodes: 28,
        title: { english: 'Frieren: Beyond Journey’s End', native: '葬送のフリーレン' },
        search: 'Frieren: Beyond Journey’s End',
        characters: [
          {
            role: 'MAIN',
            node: {
              id: 222,
              description: 'Elven mage.',
              image: { large: 'https://example.com/frieren.png', medium: null },
              name: { full: 'Frieren', native: 'フリーレン' },
            },
          },
        ],
      },
    ],
    images: ['https://example.com/alpha.png', 'https://example.com/frieren.png'],
  });
  const { runtime } = createRuntime({
    getCurrentMediaPath: () => '/tmp/current.mkv',
    getCurrentMediaTitle: () => current.title,
    guessAnilistMediaInfo: async () => ({
      title: current.title,
      season: null,
      episode: current.episode,
      source: 'fallback',
    }),
  });

  await runtime.getOrCreateCurrentSnapshot();
  current.title = 'Frieren: Beyond Journey’s End';
  current.episode = 1;
  await runtime.getOrCreateCurrentSnapshot();

  const merged = await runtime.buildMergedDictionary([21, EMINENCE_MEDIA_ID]);
  const mergedReordered = await runtime.buildMergedDictionary([EMINENCE_MEDIA_ID, 21]);
  const index = JSON.parse(readStoredZipEntry(merged.zipPath, 'index.json').toString('utf8')) as {
    title: string;
  };
  const termBank = readTermBank(merged.zipPath);
  const frieren = termBank.find(([term]) => term === 'フリーレン');
  const alpha = termBank.find(([term]) => term === 'アルファ');

  assert.equal(index.title, 'SubMiner Character Dictionary');
  assert.equal(merged.entryCount >= 2, true);
  assert.equal(merged.revision, mergedReordered.revision);
  assert.equal((frieren?.[5][0] as { type?: string }).type, 'structured-content');
  assert.equal((alpha?.[5][0] as { type?: string }).type, 'structured-content');
});

test('buildMergedDictionary rebuilds snapshots written with an older format version', async (t) => {
  const anilist = installFakeAniList(t, {
    media: [
      eminenceMedia([
        {
          role: 'MAIN',
          node: {
            id: 111,
            description: 'Leader of Shadow Garden.',
            image: null,
            name: { full: 'Cid Kagenou', native: 'シド・カゲノー', alternative: ['Shadow'] },
          },
        },
      ]),
    ],
  });
  const { runtime, outputDir } = createRuntime({
    getCurrentMediaPath: () => null,
    getCurrentMediaTitle: () => null,
    guessAnilistMediaInfo: async () => null,
  });
  const snapshotPath = getSnapshotPath(outputDir, EMINENCE_MEDIA_ID);
  fs.mkdirSync(path.dirname(snapshotPath), { recursive: true });
  fs.writeFileSync(
    snapshotPath,
    JSON.stringify({
      formatVersion: 12,
      mediaId: EMINENCE_MEDIA_ID,
      mediaTitle: 'The Eminence in Shadow',
      entryCount: 1,
      updatedAt: 1_700_000_000_000,
      termEntries: [['stale', '', 'name main', '', 100, ['stale'], 0, '']],
      images: [],
    }),
    'utf8',
  );

  const merged = await runtime.buildMergedDictionary([EMINENCE_MEDIA_ID]);
  const termBank = readTermBank(merged.zipPath);

  assert.equal(anilist.characterPageCount, 1);
  assert.ok(termBank.some(([term]) => term === 'シャドウ'));
  assert.equal(
    termBank.some(([term]) => term === 'stale'),
    false,
  );
});

test('generateForCurrentMedia paces AniList requests and character image downloads', async (t) => {
  const anilist = installFakeAniList(t, {
    media: [
      eminenceMedia([
        {
          role: 'MAIN',
          node: {
            id: 111,
            description: 'First character.',
            image: { large: 'https://example.com/alpha.png', medium: null },
            name: { full: 'Alpha', native: 'アルファ' },
          },
        },
        {
          role: 'SUPPORTING',
          node: {
            id: 222,
            description: 'Second character.',
            image: { large: 'https://example.com/beta.png', medium: null },
            name: { full: 'Beta', native: 'ベータ' },
          },
        },
      ]),
    ],
    images: ['https://example.com/beta.png'],
    missingImages: ['https://example.com/alpha.png'],
  });
  // Number of requests already sent each time the runtime pauses.
  const requestsBeforeSleep: number[] = [];
  const { runtime, outputDir } = createRuntime({
    sleep: async () => {
      requestsBeforeSleep.push(anilist.requests.length);
    },
  });

  await runtime.generateForCurrentMedia();

  // search, [pause], character page, alpha image, [pause], beta image
  assert.deepEqual(requestsBeforeSleep, [1, 3]);
  assert.equal(anilist.requests.length, 4);

  // The 404 portrait is skipped without aborting generation.
  const snapshot = JSON.parse(
    fs.readFileSync(getSnapshotPath(outputDir, EMINENCE_MEDIA_ID), 'utf8'),
  ) as CharacterDictionarySnapshot;
  assert.deepEqual(
    snapshot.images.map((image) => image.path),
    ['img/m130298-c222.png'],
  );
});
