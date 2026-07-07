import assert from 'node:assert/strict';
import test from 'node:test';
import { FieldGroupingService } from './field-grouping';
import type { KikuMergePreviewResponse } from '../types/anki';

type NoteInfo = {
  noteId: number;
  fields: Record<string, { value: string }>;
};

function createHarness(
  options: {
    kikuEnabled?: boolean;
    kikuFieldGrouping?: 'auto' | 'manual' | 'disabled';
    deck?: string;
    noteIds?: number[];
    notesInfo?: NoteInfo[][];
    duplicateNoteId?: number | null;
    trackedDuplicateNoteIds?: number[] | null;
    hasAllConfiguredFields?: boolean;
    manualHandled?: boolean;
    expression?: string | null;
    currentSentenceImageField?: string | undefined;
    onProcessNewCard?: (noteId: number, options?: { skipKikuFieldGrouping?: boolean }) => void;
  } = {},
) {
  const calls: string[] = [];
  const findNotesQueries: Array<{ query: string; maxRetries?: number }> = [];
  const noteInfoRequests: number[][] = [];
  const duplicateRequests: Array<{ expression: string; excludeNoteId: number }> = [];
  const processCalls: Array<{ noteId: number; options?: { skipKikuFieldGrouping?: boolean } }> = [];
  const autoCalls: Array<{ originalNoteId: number; newNoteId: number; expression: string }> = [];
  const manualCalls: Array<{ originalNoteId: number; newNoteId: number; expression: string }> = [];

  const noteInfoQueue = [...(options.notesInfo ?? [])];
  const notes = options.noteIds ?? [2];

  const service = new FieldGroupingService({
    getConfig: () => ({
      fields: {
        word: 'Expression',
      },
    }),
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: options.kikuEnabled ?? true,
      kikuFieldGrouping: options.kikuFieldGrouping ?? 'auto',
      kikuDeleteDuplicateInAuto: true,
    }),
    isUpdateInProgress: () => false,
    getDeck: options.deck ? () => options.deck : undefined,
    withUpdateProgress: async (_message, action) => {
      calls.push('withUpdateProgress');
      return action();
    },
    showOsdNotification: (text) => {
      calls.push(`osd:${text}`);
    },
    findNotes: async (query, findNotesOptions) => {
      findNotesQueries.push({ query, maxRetries: findNotesOptions?.maxRetries });
      return notes;
    },
    notesInfo: async (noteIds) => {
      noteInfoRequests.push([...noteIds]);
      return noteInfoQueue.shift() ?? [];
    },
    extractFields: (fields) =>
      Object.fromEntries(
        Object.entries(fields).map(([key, value]) => [key.toLowerCase(), value.value || '']),
      ),
    findDuplicateNote: async (expression, excludeNoteId) => {
      duplicateRequests.push({ expression, excludeNoteId });
      return options.duplicateNoteId ?? 99;
    },
    getTrackedDuplicateNoteIds: () => options.trackedDuplicateNoteIds ?? null,
    hasAllConfiguredFields: () => options.hasAllConfiguredFields ?? true,
    processNewCard: async (noteId, processOptions) => {
      processCalls.push({ noteId, options: processOptions });
      options.onProcessNewCard?.(noteId, processOptions);
    },
    getSentenceCardImageFieldName: () => options.currentSentenceImageField,
    resolveFieldName: (availableFieldNames, preferredName) =>
      availableFieldNames.find(
        (name) => name === preferredName || name.toLowerCase() === preferredName.toLowerCase(),
      ) ?? null,
    computeFieldGroupingMergedFields: async () => ({}),
    getNoteFieldMap: (noteInfo) =>
      Object.fromEntries(
        Object.entries(noteInfo.fields).map(([key, value]) => [key, value.value || '']),
      ),
    handleFieldGroupingAuto: async (originalNoteId, newNoteId, _newNoteInfo, expression) => {
      autoCalls.push({ originalNoteId, newNoteId, expression });
    },
    handleFieldGroupingManual: async (originalNoteId, newNoteId, _newNoteInfo, expression) => {
      manualCalls.push({ originalNoteId, newNoteId, expression });
      return options.manualHandled ?? true;
    },
  });

  return {
    service,
    calls,
    findNotesQueries,
    noteInfoRequests,
    duplicateRequests,
    processCalls,
    autoCalls,
    manualCalls,
  };
}

type SuccessfulPreview = KikuMergePreviewResponse & {
  ok: true;
  compact: {
    action: {
      keepNoteId: number;
      deleteNoteId: number;
      deleteDuplicate: boolean;
    };
    mergedFields: Record<string, string>;
  };
  full: {
    result: {
      wouldDeleteNoteId: number | null;
    };
  };
};

test('triggerFieldGroupingForLastAddedCard stops when kiku mode is disabled', async () => {
  const harness = createHarness({ kikuEnabled: false });

  await harness.service.triggerFieldGroupingForLastAddedCard();

  assert.deepEqual(harness.calls, ['osd:Kiku mode is not enabled']);
  assert.equal(harness.findNotesQueries.length, 0);
});

test('triggerFieldGroupingForLastAddedCard stops when field grouping is disabled', async () => {
  const harness = createHarness({ kikuFieldGrouping: 'disabled' });

  await harness.service.triggerFieldGroupingForLastAddedCard();

  assert.deepEqual(harness.calls, ['osd:Kiku field grouping is disabled']);
  assert.equal(harness.findNotesQueries.length, 0);
});

test('triggerFieldGroupingForLastAddedCard stops when an update is already in progress', async () => {
  const service = new FieldGroupingService({
    getConfig: () => ({ fields: { word: 'Expression' } }),
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: true,
      kikuFieldGrouping: 'auto',
      kikuDeleteDuplicateInAuto: true,
    }),
    isUpdateInProgress: () => true,
    withUpdateProgress: async () => {
      throw new Error('should not be called');
    },
    showOsdNotification: () => {},
    findNotes: async () => [],
    notesInfo: async () => [],
    extractFields: () => ({}),
    findDuplicateNote: async () => null,
    hasAllConfiguredFields: () => true,
    processNewCard: async () => {},
    getSentenceCardImageFieldName: () => undefined,
    resolveFieldName: () => null,
    computeFieldGroupingMergedFields: async () => ({}),
    getNoteFieldMap: () => ({}),
    handleFieldGroupingAuto: async () => {},
    handleFieldGroupingManual: async () => true,
  });

  await service.triggerFieldGroupingForLastAddedCard();
});

test('triggerFieldGroupingForLastAddedCard finds the newest note and hands off to auto grouping', async () => {
  const harness = createHarness({
    deck: 'Anime Deck',
    noteIds: [3, 7, 5],
    notesInfo: [
      [
        {
          noteId: 7,
          fields: {
            Expression: { value: 'word-7' },
            Sentence: { value: 'line-7' },
          },
        },
      ],
      [
        {
          noteId: 7,
          fields: {
            Expression: { value: 'word-7' },
            Sentence: { value: 'line-7' },
          },
        },
      ],
    ],
    duplicateNoteId: 42,
    hasAllConfiguredFields: true,
  });

  await harness.service.triggerFieldGroupingForLastAddedCard();

  assert.deepEqual(harness.findNotesQueries, [
    { query: '"deck:Anime Deck" added:1', maxRetries: undefined },
  ]);
  assert.deepEqual(harness.noteInfoRequests, [[7], [7]]);
  assert.deepEqual(harness.duplicateRequests, [{ expression: 'word-7', excludeNoteId: 7 }]);
  assert.deepEqual(harness.autoCalls, [
    {
      originalNoteId: 42,
      newNoteId: 7,
      expression: 'word-7',
    },
  ]);
});

test('triggerFieldGroupingForLastAddedCard prefers tracked duplicate note ids before duplicate lookup', async () => {
  const harness = createHarness({
    noteIds: [7],
    notesInfo: [
      [
        {
          noteId: 7,
          fields: {
            Expression: { value: 'word-7' },
            Sentence: { value: 'line-7' },
          },
        },
      ],
      [
        {
          noteId: 7,
          fields: {
            Expression: { value: 'word-7' },
            Sentence: { value: 'line-7' },
          },
        },
      ],
    ],
    trackedDuplicateNoteIds: [12, 40, 25],
    duplicateNoteId: 99,
    hasAllConfiguredFields: true,
  });

  await harness.service.triggerFieldGroupingForLastAddedCard();

  assert.deepEqual(harness.duplicateRequests, []);
  assert.deepEqual(harness.autoCalls, [
    {
      originalNoteId: 40,
      newNoteId: 7,
      expression: 'word-7',
    },
  ]);
});

test('triggerFieldGroupingForLastAddedCard refreshes the card when configured fields are missing', async () => {
  const processCalls: Array<{ noteId: number; options?: { skipKikuFieldGrouping?: boolean } }> = [];
  const harness = createHarness({
    noteIds: [11],
    notesInfo: [
      [
        {
          noteId: 11,
          fields: {
            Expression: { value: 'word-11' },
            Sentence: { value: 'line-11' },
          },
        },
      ],
      [
        {
          noteId: 11,
          fields: {
            Expression: { value: 'word-11' },
            Sentence: { value: 'line-11' },
          },
        },
      ],
    ],
    duplicateNoteId: 13,
    hasAllConfiguredFields: false,
    onProcessNewCard: (noteId, options) => {
      processCalls.push({ noteId, options });
    },
  });

  await harness.service.triggerFieldGroupingForLastAddedCard();

  assert.deepEqual(processCalls, [{ noteId: 11, options: { skipKikuFieldGrouping: true } }]);
  assert.deepEqual(harness.manualCalls, []);
});

test('triggerFieldGroupingForLastAddedCard does not re-notify when manual grouping is declined', async () => {
  const harness = createHarness({
    kikuFieldGrouping: 'manual',
    noteIds: [9],
    notesInfo: [
      [
        {
          noteId: 9,
          fields: {
            Expression: { value: 'word-9' },
            Sentence: { value: 'line-9' },
          },
        },
      ],
      [
        {
          noteId: 9,
          fields: {
            Expression: { value: 'word-9' },
            Sentence: { value: 'line-9' },
          },
        },
      ],
    ],
    duplicateNoteId: 77,
    manualHandled: false,
  });

  await harness.service.triggerFieldGroupingForLastAddedCard();

  assert.deepEqual(harness.manualCalls, [
    {
      originalNoteId: 77,
      newNoteId: 9,
      expression: 'word-9',
    },
  ]);
  // The manual workflow already notifies about its outcome (cancelled/unavailable/failed);
  // the trigger wrapper re-notifying produced two "Field grouping cancelled" toasts.
  assert.equal(harness.calls.filter((call) => call === 'osd:Field grouping cancelled').length, 0);
});

test('buildFieldGroupingPreview returns merged compact and full previews', async () => {
  const service = new FieldGroupingService({
    getConfig: () => ({ fields: { word: 'Expression' } }),
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: true,
      kikuFieldGrouping: 'auto',
      kikuDeleteDuplicateInAuto: true,
    }),
    isUpdateInProgress: () => false,
    withUpdateProgress: async (_message, action) => action(),
    showOsdNotification: () => {},
    findNotes: async () => [],
    notesInfo: async (noteIds) =>
      noteIds.map((noteId) => ({
        noteId,
        fields: {
          Sentence: { value: `sentence-${noteId}` },
          SentenceAudio: { value: `[sound:${noteId}.mp3]` },
          Picture: { value: `<img src="${noteId}.png">` },
          MiscInfo: { value: `misc-${noteId}` },
        },
      })),
    extractFields: () => ({}),
    findDuplicateNote: async () => null,
    hasAllConfiguredFields: () => true,
    processNewCard: async () => {},
    getSentenceCardImageFieldName: () => undefined,
    resolveFieldName: (availableFieldNames, preferredName) =>
      availableFieldNames.find(
        (name) => name === preferredName || name.toLowerCase() === preferredName.toLowerCase(),
      ) ?? null,
    computeFieldGroupingMergedFields: async () => ({
      Sentence: 'merged sentence',
      SentenceAudio: 'merged audio',
      Picture: 'merged picture',
      MiscInfo: 'merged misc',
    }),
    getNoteFieldMap: (noteInfo) =>
      Object.fromEntries(
        Object.entries(noteInfo.fields).map(([key, value]) => [key, value.value || '']),
      ),
    handleFieldGroupingAuto: async () => {},
    handleFieldGroupingManual: async () => true,
  });

  const preview = await service.buildFieldGroupingPreview(1, 2, true);

  assert.equal(preview.ok, true);
  if (!preview.ok) {
    throw new Error(preview.error);
  }
  const successPreview = preview as SuccessfulPreview;
  assert.deepEqual(successPreview.compact.action, {
    keepNoteId: 1,
    deleteNoteId: 2,
    deleteDuplicate: true,
  });
  assert.equal(successPreview.compact.mergedFields.Sentence, 'merged sentence');
  assert.equal(successPreview.full.result.wouldDeleteNoteId, 2);
});

test('buildFieldGroupingPreview reports missing notes cleanly', async () => {
  const service = new FieldGroupingService({
    getConfig: () => ({ fields: { word: 'Expression' } }),
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: true,
      kikuFieldGrouping: 'auto',
      kikuDeleteDuplicateInAuto: true,
    }),
    isUpdateInProgress: () => false,
    withUpdateProgress: async (_message, action) => action(),
    showOsdNotification: () => {},
    findNotes: async () => [],
    notesInfo: async () => [
      {
        noteId: 1,
        fields: {
          Sentence: { value: 'sentence-1' },
        },
      },
    ],
    extractFields: () => ({}),
    findDuplicateNote: async () => null,
    hasAllConfiguredFields: () => true,
    processNewCard: async () => {},
    getSentenceCardImageFieldName: () => undefined,
    resolveFieldName: () => null,
    computeFieldGroupingMergedFields: async () => ({}),
    getNoteFieldMap: () => ({}),
    handleFieldGroupingAuto: async () => {},
    handleFieldGroupingManual: async () => true,
  });

  const preview = await service.buildFieldGroupingPreview(1, 2, false);

  assert.equal(preview.ok, false);
  if (preview.ok) {
    throw new Error('expected preview to fail');
  }
  assert.equal(preview.error, 'Could not load selected notes');
});
