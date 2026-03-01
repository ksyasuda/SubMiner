import test from 'node:test';
import assert from 'node:assert/strict';
import { findDuplicateNote, type NoteInfo } from './duplicate';

function createFieldResolver(noteInfo: NoteInfo, preferredName: string): string | null {
  const names = Object.keys(noteInfo.fields);
  const exact = names.find((name) => name === preferredName);
  if (exact) return exact;
  const lower = preferredName.toLowerCase();
  return names.find((name) => name.toLowerCase() === lower) ?? null;
}

test('findDuplicateNote matches duplicate when candidate uses alternate word/expression field name', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '食べる' },
    },
  };

  const duplicateId = await findDuplicateNote('食べる', 100, currentNote, {
    findNotes: async () => [100, 200],
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Word: { value: '食べる' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
});

test('findDuplicateNote falls back to alias field query when primary field query returns no candidates', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '食べる' },
    },
  };

  const seenQueries: string[] = [];
  const duplicateId = await findDuplicateNote('食べる', 100, currentNote, {
    findNotes: async (query) => {
      seenQueries.push(query);
      if (query.includes('"Expression:')) {
        return [];
      }
      if (query.includes('"word:') || query.includes('"Word:') || query.includes('"expression:')) {
        return [200];
      }
      return [];
    },
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Word: { value: '食べる' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
  assert.equal(seenQueries.length, 2);
});

test('findDuplicateNote checks both source expression/word values when both fields are present', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '昨日は雨だった。' },
      Word: { value: '雨' },
    },
  };

  const seenQueries: string[] = [];
  const duplicateId = await findDuplicateNote('昨日は雨だった。', 100, currentNote, {
    findNotes: async (query) => {
      seenQueries.push(query);
      if (query.includes('昨日は雨だった。')) {
        return [];
      }
      if (
        query.includes('"Word:雨"') ||
        query.includes('"word:雨"') ||
        query.includes('"Expression:雨"')
      ) {
        return [200];
      }
      return [];
    },
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Word: { value: '雨' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
  assert.ok(seenQueries.some((query) => query.includes('昨日は雨だった。')));
  assert.ok(seenQueries.some((query) => query.includes('雨')));
});

test('findDuplicateNote falls back to collection-wide query when deck-scoped query has no matches', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '貴様' },
    },
  };

  const seenQueries: string[] = [];
  const duplicateId = await findDuplicateNote('貴様', 100, currentNote, {
    findNotes: async (query) => {
      seenQueries.push(query);
      if (query.includes('deck:Japanese')) {
        return [];
      }
      if (query.includes('"Expression:貴様"') || query.includes('"Word:貴様"')) {
        return [200];
      }
      return [];
    },
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Expression: { value: '貴様' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
  assert.ok(seenQueries.some((query) => query.includes('deck:Japanese')));
  assert.ok(seenQueries.some((query) => !query.includes('deck:Japanese')));
});

test('findDuplicateNote falls back to plain text query when field queries miss', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '貴様' },
    },
  };

  const seenQueries: string[] = [];
  const duplicateId = await findDuplicateNote('貴様', 100, currentNote, {
    findNotes: async (query) => {
      seenQueries.push(query);
      if (query.includes('Expression:') || query.includes('Word:')) {
        return [];
      }
      if (query.includes('"貴様"')) {
        return [200];
      }
      return [];
    },
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Expression: { value: '貴様' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
  assert.ok(seenQueries.some((query) => query.includes('Expression:')));
  assert.ok(seenQueries.some((query) => query.endsWith('"貴様"')));
});

test('findDuplicateNote exact compare tolerates furigana bracket markup in candidate field', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '貴様' },
    },
  };

  const duplicateId = await findDuplicateNote('貴様', 100, currentNote, {
    findNotes: async () => [200],
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Expression: { value: '貴様[きさま]' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
});

test('findDuplicateNote exact compare tolerates html wrappers in candidate field', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '貴様' },
    },
  };

  const duplicateId = await findDuplicateNote('貴様', 100, currentNote, {
    findNotes: async () => [200],
    notesInfo: async () => [
      {
        noteId: 200,
        fields: {
          Expression: { value: '<span data-x="1">貴様</span>' },
        },
      },
    ],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.equal(duplicateId, 200);
});

test('findDuplicateNote does not disable retries on findNotes calls', async () => {
  const currentNote: NoteInfo = {
    noteId: 100,
    fields: {
      Expression: { value: '貴様' },
    },
  };

  const seenOptions: Array<{ maxRetries?: number } | undefined> = [];
  await findDuplicateNote('貴様', 100, currentNote, {
    findNotes: async (_query, options) => {
      seenOptions.push(options);
      return [];
    },
    notesInfo: async () => [],
    getDeck: () => 'Japanese::Mining',
    resolveFieldName: (noteInfo, preferredName) => createFieldResolver(noteInfo, preferredName),
    logWarn: () => {},
  });

  assert.ok(seenOptions.length > 0);
  assert.ok(seenOptions.every((options) => options?.maxRetries !== 0));
});
