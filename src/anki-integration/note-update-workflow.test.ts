import test from 'node:test';
import assert from 'node:assert/strict';
import { NoteUpdateWorkflow } from './note-update-workflow';

type NoteInfo = {
  noteId: number;
  fields: Record<string, { value: string }>;
};

function createWorkflowHarness() {
  const updates: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const notifications: Array<{ noteId: number; label: string | number }> = [];
  const warnings: string[] = [];

  const deps = {
    client: {
      notesInfo: async (_noteIds: number[]) =>
        [
          {
            noteId: 42,
            fields: {
              Expression: { value: 'taberu' },
              Sentence: { value: '' },
            },
          },
        ] satisfies NoteInfo[],
      updateNoteFields: async (noteId: number, fields: Record<string, string>) => {
        updates.push({ noteId, fields });
      },
      storeMediaFile: async () => undefined,
    },
    getConfig: () => ({
      fields: {
        sentence: 'Sentence',
      },
      media: {},
      behavior: {},
    }),
    getCurrentSubtitleText: () => 'subtitle-text',
    getCurrentSubtitleStart: () => 12.3,
    getEffectiveSentenceCardConfig: () => ({
      sentenceField: 'Sentence',
      kikuEnabled: false,
      kikuFieldGrouping: 'disabled' as const,
    }),
    appendKnownWordsFromNoteInfo: (_noteInfo: NoteInfo) => undefined,
    extractFields: (fields: Record<string, { value: string }>) => {
      const out: Record<string, string> = {};
      for (const [key, value] of Object.entries(fields)) {
        out[key.toLowerCase()] = value.value;
      }
      return out;
    },
    findDuplicateNote: async () => null,
    handleFieldGroupingAuto: async () => undefined,
    handleFieldGroupingManual: async () => false,
    processSentence: (text: string) => text,
    resolveConfiguredFieldName: (noteInfo: NoteInfo, preferred?: string) => {
      if (!preferred) return null;
      const names = Object.keys(noteInfo.fields);
      return names.find((name) => name.toLowerCase() === preferred.toLowerCase()) ?? null;
    },
    getResolvedSentenceAudioFieldName: () => null,
    mergeFieldValue: (_existing: string, next: string) => next,
    generateAudioFilename: () => 'audio_1.mp3',
    generateAudio: async () => null,
    generateImageFilename: () => 'image_1.jpg',
    generateImage: async () => null,
    formatMiscInfoPattern: () => '',
    addConfiguredTagsToNote: async () => undefined,
    showNotification: async (noteId: number, label: string | number) => {
      notifications.push({ noteId, label });
    },
    showOsdNotification: (_text: string) => undefined,
    beginUpdateProgress: (_text: string) => undefined,
    endUpdateProgress: () => undefined,
    logWarn: (message: string) => warnings.push(message),
    logInfo: (_message: string) => undefined,
    logError: (_message: string) => undefined,
  };

  return {
    workflow: new NoteUpdateWorkflow(deps),
    updates,
    notifications,
    warnings,
    deps,
  };
}

test('NoteUpdateWorkflow updates sentence field and emits notification', async () => {
  const harness = createWorkflowHarness();

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.equal(harness.updates[0]?.noteId, 42);
  assert.equal(harness.updates[0]?.fields.Sentence, 'subtitle-text');
  assert.equal(harness.notifications.length, 1);
});

test('NoteUpdateWorkflow no-ops when note info is missing', async () => {
  const harness = createWorkflowHarness();
  harness.deps.client.notesInfo = async () => [];

  await harness.workflow.execute(777);

  assert.equal(harness.updates.length, 0);
  assert.equal(harness.notifications.length, 0);
  assert.equal(harness.warnings.length, 1);
});
