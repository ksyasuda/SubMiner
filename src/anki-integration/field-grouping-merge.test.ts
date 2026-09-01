import assert from 'node:assert/strict';
import test from 'node:test';
import {
  FieldGroupingMergeCollaborator,
  type FieldGroupingMergeNoteInfo,
} from './field-grouping-merge';
import type { AnkiConnectConfig } from '../types/anki';

function resolveFieldName(availableFieldNames: string[], preferredName: string): string | null {
  return (
    availableFieldNames.find(
      (name) => name === preferredName || name.toLowerCase() === preferredName.toLowerCase(),
    ) ?? null
  );
}

function createCollaborator(
  options: {
    config?: Partial<AnkiConnectConfig>;
    currentSubtitleText?: string;
    generatedMedia?: {
      audioField?: string;
      audioValue?: string;
      imageField?: string;
      imageValue?: string;
      miscInfoValue?: string;
    };
    warnings?: Array<{ fieldName: string; reason: string; detail?: string }>;
    fieldGroupingProvider?: 'kiku' | 'senren' | null;
  } = {},
) {
  const warnings = options.warnings ?? [];
  const config = {
    fields: {
      sentence: 'Sentence',
      audio: 'ExpressionAudio',
      image: 'Picture',
      miscInfo: 'MiscInfo',
      ...(options.config?.fields ?? {}),
    },
    ...(options.config ?? {}),
  } as AnkiConnectConfig;

  return {
    collaborator: new FieldGroupingMergeCollaborator({
      getConfig: () => config,
      getEffectiveSentenceCardConfig: () => ({
        sentenceField: 'Sentence',
        audioField: 'SentenceAudio',
        fieldGroupingProvider: options.fieldGroupingProvider ?? 'kiku',
      }),
      getCurrentSubtitleText: () => options.currentSubtitleText,
      resolveFieldName,
      resolveNoteFieldName: (noteInfo, preferredName) => {
        if (!preferredName) return null;
        return resolveFieldName(Object.keys(noteInfo.fields), preferredName);
      },
      extractFields: (fields) =>
        Object.fromEntries(
          Object.entries(fields).map(([key, value]) => [key.toLowerCase(), value.value || '']),
        ),
      processSentence: (mpvSentence) => `${mpvSentence}::processed`,
      generateMediaForMerge: async () => options.generatedMedia ?? {},
      warnFieldParseOnce: (fieldName, reason, detail) => {
        warnings.push({ fieldName, reason, detail });
      },
    }),
    warnings,
  };
}

function makeNote(noteId: number, fields: Record<string, string>): FieldGroupingMergeNoteInfo {
  return {
    noteId,
    fields: Object.fromEntries(Object.entries(fields).map(([key, value]) => [key, { value }])),
  };
}

test('getGroupableFieldNames includes Kiku context fields and omits word audio fields', () => {
  const { collaborator } = createCollaborator({
    config: {
      fields: {
        image: 'Illustration',
        sentence: 'SentenceText',
        audio: 'CustomWordAudio',
        miscInfo: 'ExtraInfo',
      },
    },
  });

  assert.deepEqual(collaborator.getGroupableFieldNames(), [
    'Sentence',
    'SentenceAudio',
    'Picture',
    'Illustration',
    'SentenceText',
    'ExtraInfo',
    'SentenceFurigana',
  ]);
});

test('computeFieldGroupingMergedFields groups both notes and sorts by descending group id when keeping original', async () => {
  const { collaborator } = createCollaborator();

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, {
      Sentence: 'original sentence',
      SentenceAudio: '[sound:original-a.mp3] [sound:original-b.mp3]',
      Picture: '<img src="original.png">',
      MiscInfo: 'original misc',
      ExpressionAudio: '[sound:word.mp3]',
    }),
    makeNote(200, {
      Sentence: 'new sentence',
      SentenceAudio: '[sound:new.mp3]',
      Picture: '<img src="new.png">',
      MiscInfo: 'new misc',
    }),
    false,
  );

  assert.equal(
    merged.Sentence,
    '<span data-group-id="300">original sentence</span><span data-group-id="200">new sentence</span>',
  );
  assert.equal(
    merged.SentenceAudio,
    '<span data-group-id="300">[sound:original-a.mp3] [sound:original-b.mp3]</span><span data-group-id="200">[sound:new.mp3]</span>',
  );
  assert.equal(
    merged.Picture,
    '<img data-group-id="300" src="original.png"><img data-group-id="200" src="new.png">',
  );
  assert.equal(
    merged.MiscInfo,
    '<span data-group-id="300">original misc</span><span data-group-id="200">new misc</span>',
  );
  assert.equal('ExpressionAudio' in merged, false);
});

test('computeFieldGroupingMergedFields sorts original before new when merging original into a newer target', async () => {
  const { collaborator } = createCollaborator();

  const merged = await collaborator.computeFieldGroupingMergedFields(
    200,
    300,
    makeNote(200, {
      Sentence: 'new sentence',
      SentenceAudio: '[sound:new.mp3]',
      Picture: '<img src="new.png">',
      MiscInfo: 'new misc',
    }),
    makeNote(300, {
      Sentence: 'original sentence',
      SentenceAudio: '[sound:original.mp3]',
      Picture: '<img src="original.png">',
      MiscInfo: 'original misc',
    }),
    false,
  );

  assert.equal(
    merged.Sentence,
    '<span data-group-id="300">original sentence</span><span data-group-id="200">new sentence</span>',
  );
  assert.equal(
    merged.SentenceAudio,
    '<span data-group-id="300">[sound:original.mp3]</span><span data-group-id="200">[sound:new.mp3]</span>',
  );
  assert.equal(
    merged.Picture,
    '<img data-group-id="300" src="original.png"><img data-group-id="200" src="new.png">',
  );
  assert.equal(
    merged.MiscInfo,
    '<span data-group-id="300">original misc</span><span data-group-id="200">new misc</span>',
  );
});

test('computeFieldGroupingMergedFields keeps strict fields when source is empty and warns on malformed spans', async () => {
  const { collaborator, warnings } = createCollaborator({
    currentSubtitleText: 'subtitle line',
  });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    3,
    4,
    makeNote(3, {
      Sentence: '<span data-group-id="abc">keep sentence</span>',
      SentenceAudio: '',
    }),
    makeNote(4, {
      Sentence: 'source sentence',
      SentenceAudio: '[sound:source.mp3]',
    }),
    false,
  );

  assert.equal(
    merged.Sentence,
    '<span data-group-id="4">source sentence</span><span data-group-id="3"><span data-group-id="abc">keep sentence</span></span>',
  );
  assert.equal(merged.SentenceAudio, '<span data-group-id="4">[sound:source.mp3]</span>');
  assert.equal(warnings.length, 4);
  assert.deepEqual(
    warnings.map((entry) => entry.reason),
    ['invalid-group-id', 'no-usable-span-entries', 'invalid-group-id', 'no-usable-span-entries'],
  );
});

test('computeFieldGroupingMergedFields uses generated media only when includeGeneratedMedia is true', async () => {
  const generatedMedia = {
    audioField: 'SentenceAudio',
    audioValue: '[sound:generated.mp3]',
    imageField: 'Picture',
    imageValue: '<img src="generated.png">',
    miscInfoValue: 'generated misc',
  };
  const { collaborator: withoutGenerated } = createCollaborator({ generatedMedia });
  const { collaborator: withGenerated } = createCollaborator({ generatedMedia });

  const keep = makeNote(10, {
    SentenceAudio: '',
    Picture: '',
    MiscInfo: '',
  });
  const source = makeNote(11, {
    SentenceAudio: '',
    Picture: '',
    MiscInfo: '',
  });

  const without = await withoutGenerated.computeFieldGroupingMergedFields(
    10,
    11,
    keep,
    source,
    false,
  );
  const withMedia = await withGenerated.computeFieldGroupingMergedFields(
    10,
    11,
    keep,
    source,
    true,
  );

  assert.deepEqual(without, {});
  assert.equal(withMedia.SentenceAudio, '<span data-group-id="11">[sound:generated.mp3]</span>');
  assert.equal(withMedia.Picture, '<img data-group-id="11" src="generated.png">');
  assert.equal(withMedia.MiscInfo, '<span data-group-id="11">generated misc</span>');
});

test('computeFieldGroupingMergedFields merges Senren notes into scene-switching markup', async () => {
  const { collaborator } = createCollaborator({ fieldGroupingProvider: 'senren' });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, {
      word: '語',
      sentence: '<span class="group">前<span class="highlight">語</span>後</span>',
      sentenceAudio: '[sound:original.opus]',
      picture: '<img src="original.webp">',
      miscInfo: '<span class="group">Show EP1 (0:01:00)</span>',
    }),
    makeNote(200, {
      word: '語',
      sentence: '<span class="group">次<span class="highlight">語</span>文</span>',
      sentenceAudio: '[sound:new.opus]',
      picture: '<img src="new.webp">',
      miscInfo: 'Show EP2 (0:02:00)',
    }),
    false,
  );

  assert.equal(
    merged.sentence,
    '<span class="group">前<span class="highlight">語</span>後</span>' +
      '<span class="group">次<span class="highlight">語</span>文</span>',
  );
  assert.equal(merged.sentenceAudio, '[sound:original.opus][sound:new.opus]');
  assert.equal(merged.picture, '<img src="original.webp"><img src="new.webp">');
  assert.equal(
    merged.miscInfo,
    '<span class="group">Show EP1 (0:01:00)</span><span class="group">Show EP2 (0:02:00)</span>',
  );
});

test('Senren merge wraps ungrouped legacy content and preserves numbered groups', async () => {
  const { collaborator } = createCollaborator({ fieldGroupingProvider: 'senren' });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, {
      sentence: 'plain legacy sentence',
      sentenceAudio: '[sound:a.opus][sound:b.opus]',
      miscInfo: '<span class="group2">pinned</span> stray text',
    }),
    makeNote(200, {
      sentence: '<span class="group">new sentence</span>',
      sentenceAudio: '[sound:c.opus]',
      miscInfo: '<span class="group">new misc</span>',
    }),
    false,
  );

  assert.equal(
    merged.sentence,
    '<span class="group">plain legacy sentence</span><span class="group">new sentence</span>',
  );
  assert.equal(merged.sentenceAudio, '[sound:a.opus][sound:b.opus][sound:c.opus]');
  assert.equal(
    merged.miscInfo,
    '<span class="group2">pinned</span><span class="group">stray text</span>' +
      '<span class="group">new misc</span>',
  );
});

test('Senren merge rebases numbered groups from an appended source note', async () => {
  const { collaborator } = createCollaborator({ fieldGroupingProvider: 'senren' });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, {
      sentenceAudio: '[sound:keep-a.opus][sound:keep-b.opus]',
      miscInfo: '<span class="group">keep one</span><span class="group">keep two</span>',
    }),
    makeNote(200, {
      sentenceAudio: '[sound:source-a.opus][sound:source-b.opus]',
      miscInfo: '<span class="group2">source two</span>',
    }),
    false,
  );

  assert.equal(
    merged.sentenceAudio,
    '[sound:keep-a.opus][sound:keep-b.opus][sound:source-a.opus][sound:source-b.opus]',
  );
  assert.equal(
    merged.miscInfo,
    '<span class="group">keep one</span><span class="group">keep two</span>' +
      '<span class="group4">source two</span>',
  );
});

test('Senren merge keeps ungrouped text in place around an existing group span', async () => {
  const { collaborator } = createCollaborator({ fieldGroupingProvider: 'senren' });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, {
      miscInfo: 'leading<span class="group">middle</span>trailing',
      sentenceAudio: '[sound:a.opus][sound:b.opus][sound:c.opus]',
    }),
    makeNote(200, {
      miscInfo: '<span class="group">appended</span>',
      sentenceAudio: '[sound:d.opus]',
    }),
    false,
  );

  // Order must follow the source field, and the two ungrouped runs must stay separate.
  assert.equal(
    merged.miscInfo,
    '<span class="group">leading</span><span class="group">middle</span>' +
      '<span class="group">trailing</span><span class="group">appended</span>',
  );
  assert.equal(merged.sentenceAudio, '[sound:a.opus][sound:b.opus][sound:c.opus][sound:d.opus]');
});

test('Senren merge closes unclosed group spans so later scenes stay siblings', async () => {
  const { collaborator } = createCollaborator({ fieldGroupingProvider: 'senren' });

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, { miscInfo: '<span class="group">a<span class="highlight">b' }),
    makeNote(200, { miscInfo: '<span class="group">next</span>' }),
    false,
  );

  assert.equal(
    merged.miscInfo,
    '<span class="group">a<span class="highlight">b</span></span><span class="group">next</span>',
  );
  const openTags = merged.miscInfo!.match(/<span\b/g)?.length ?? 0;
  const closeTags = merged.miscInfo!.match(/<\/span>/g)?.length ?? 0;
  assert.equal(openTags, closeTags);
});

test('computeFieldGroupingMergedFields clears SentenceFurigana when either note lacks it', async () => {
  const { collaborator } = createCollaborator();

  const merged = await collaborator.computeFieldGroupingMergedFields(
    300,
    200,
    makeNote(300, {
      SentenceFurigana: 'original furigana',
    }),
    makeNote(200, {
      SentenceFurigana: '',
    }),
    false,
  );

  assert.equal(merged.SentenceFurigana, '');
});
