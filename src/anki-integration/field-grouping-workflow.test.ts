import test from 'node:test';
import assert from 'node:assert/strict';
import { FieldGroupingWorkflow } from './field-grouping-workflow';
import type { KikuDuplicateCardInfo, KikuFieldGroupingChoice } from '../types/anki';

type NoteInfo = {
  noteId: number;
  fields: Record<string, { value: string }>;
};

type ManualChoice = {
  keepNoteId: number;
  deleteNoteId: number;
  deleteDuplicate: boolean;
  cancelled: boolean;
};

type FieldGroupingCallback = (data: {
  original: KikuDuplicateCardInfo;
  duplicate: KikuDuplicateCardInfo;
}) => Promise<KikuFieldGroupingChoice>;

function createWorkflowHarness() {
  const updates: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const deleted: number[][] = [];
  const statuses: string[] = [];
  const rememberedMerges: Array<{ deletedNoteId: number; keptNoteId: number }> = [];
  const mergeCalls: Array<{
    keepNoteId: number;
    deleteNoteId: number;
    keepNoteInfoNoteId: number;
    deleteNoteInfoNoteId: number;
  }> = [];
  let manualChoice: ManualChoice | null = null;

  const deps = {
    client: {
      notesInfo: async (noteIds: number[]) =>
        noteIds.map(
          (noteId) =>
            ({
              noteId,
              fields: {
                Expression: { value: `word-${noteId}` },
                Sentence: { value: `line-${noteId}` },
              },
            }) satisfies NoteInfo,
        ),
      updateNoteFields: async (noteId: number, fields: Record<string, string>) => {
        updates.push({ noteId, fields });
      },
      deleteNotes: async (noteIds: number[]) => {
        deleted.push(noteIds);
      },
    },
    getConfig: () => ({
      fields: {
        audio: 'ExpressionAudio',
        image: 'Picture',
      },
      isKiku: {
        deleteDuplicateInAuto: true,
      },
    }),
    getEffectiveSentenceCardConfig: () => ({
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      kikuDeleteDuplicateInAuto: true,
    }),
    getCurrentSubtitleText: () => 'subtitle-text',
    getFieldGroupingCallback: (): FieldGroupingCallback | null => {
      const choice = manualChoice;
      if (choice === null) return null;
      return async () => choice;
    },
    setFieldGroupingCallback: () => undefined,
    computeFieldGroupingMergedFields: async (
      keepNoteId: number,
      deleteNoteId: number,
      keepNoteInfo: NoteInfo,
      deleteNoteInfo: NoteInfo,
    ) => {
      mergeCalls.push({
        keepNoteId,
        deleteNoteId,
        keepNoteInfoNoteId: keepNoteInfo.noteId,
        deleteNoteInfoNoteId: deleteNoteInfo.noteId,
      });
      return {
        Sentence: 'merged sentence',
      };
    },
    extractFields: (fields: Record<string, { value: string }>) => {
      const out: Record<string, string> = {};
      for (const [key, value] of Object.entries(fields)) {
        out[key.toLowerCase()] = value.value;
      }
      return out;
    },
    hasFieldValue: (_noteInfo: NoteInfo, _field?: string) => false,
    addConfiguredTagsToNote: async () => undefined,
    removeTrackedNoteId: () => undefined,
    rememberMergedNoteIds: (deletedNoteId: number, keptNoteId: number) => {
      rememberedMerges.push({ deletedNoteId, keptNoteId });
    },
    showStatusNotification: (message: string) => {
      statuses.push(message);
    },
    showNotification: async () => undefined,
    showOsdNotification: () => undefined,
    logError: () => undefined,
    logInfo: () => undefined,
    truncateSentence: (value: string) => value,
  };

  return {
    workflow: new FieldGroupingWorkflow(deps),
    updates,
    deleted,
    rememberedMerges,
    statuses,
    mergeCalls,
    setManualChoice: (choice: typeof manualChoice) => {
      manualChoice = choice;
    },
    deps,
  };
}

test('FieldGroupingWorkflow auto merge updates keep note and deletes duplicate by default', async () => {
  const harness = createWorkflowHarness();

  await harness.workflow.handleAuto(1, 2, {
    noteId: 2,
    fields: {
      Expression: { value: 'word-2' },
      Sentence: { value: 'line-2' },
    },
  });

  assert.equal(harness.updates.length, 1);
  assert.equal(harness.updates[0]?.noteId, 1);
  assert.deepEqual(harness.deleted, [[2]]);
  assert.deepEqual(harness.rememberedMerges, [{ deletedNoteId: 2, keptNoteId: 1 }]);
  assert.equal(harness.statuses.length, 1);
});

test('FieldGroupingWorkflow manual mode returns false when callback unavailable', async () => {
  const harness = createWorkflowHarness();

  const handled = await harness.workflow.handleManual(1, 2, {
    noteId: 2,
    fields: {
      Expression: { value: 'word-2' },
      Sentence: { value: 'line-2' },
    },
  });

  assert.equal(handled, false);
  assert.equal(harness.updates.length, 0);
});

test('FieldGroupingWorkflow manual keep-new uses new note as merge target and old note as source', async () => {
  const harness = createWorkflowHarness();
  harness.setManualChoice({
    keepNoteId: 2,
    deleteNoteId: 1,
    deleteDuplicate: false,
    cancelled: false,
  });

  const handled = await harness.workflow.handleManual(1, 2, {
    noteId: 2,
    fields: {
      Expression: { value: 'word-2' },
      Sentence: { value: 'line-2' },
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(harness.mergeCalls, [
    {
      keepNoteId: 2,
      deleteNoteId: 1,
      keepNoteInfoNoteId: 2,
      deleteNoteInfoNoteId: 1,
    },
  ]);
});
