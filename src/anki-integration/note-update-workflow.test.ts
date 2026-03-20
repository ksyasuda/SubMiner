import test from 'node:test';
import assert from 'node:assert/strict';
import {
  NoteUpdateWorkflow,
  type NoteUpdateWorkflowDeps,
  type NoteUpdateWorkflowNoteInfo,
} from './note-update-workflow';

function createWorkflowHarness() {
  const updates: Array<{ noteId: number; fields: Record<string, string> }> = [];
  const notifications: Array<{ noteId: number; label: string | number }> = [];
  const warnings: string[] = [];

  const deps: NoteUpdateWorkflowDeps = {
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
        ] satisfies NoteUpdateWorkflowNoteInfo[],
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
    appendKnownWordsFromNoteInfo: (_noteInfo: NoteUpdateWorkflowNoteInfo) => undefined,
    extractFields: (fields: Record<string, { value: string }>) => {
      const out: Record<string, string> = {};
      for (const [key, value] of Object.entries(fields)) {
        out[key.toLowerCase()] = value.value;
      }
      return out;
    },
    findDuplicateNote: async (_expression, _excludeNoteId, _noteInfo) => null,
    handleFieldGroupingAuto: async (_originalNoteId, _newNoteId, _newNoteInfo, _expression) =>
      undefined,
    handleFieldGroupingManual: async (_originalNoteId, _newNoteId, _newNoteInfo, _expression) =>
      false,
    processSentence: (text: string, _noteFields: Record<string, string>) => text,
    resolveConfiguredFieldName: (noteInfo: NoteUpdateWorkflowNoteInfo, preferred?: string) => {
      if (!preferred) return null;
      const names = Object.keys(noteInfo.fields);
      return names.find((name) => name.toLowerCase() === preferred.toLowerCase()) ?? null;
    },
    getResolvedSentenceAudioFieldName: () => null,
    getAnimatedImageLeadInSeconds: async () => 0,
    mergeFieldValue: (_existing: string, next: string, _overwrite: boolean) => next,
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
    logWarn: (message: string, ..._args: unknown[]) => warnings.push(message),
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

test('NoteUpdateWorkflow updates note before auto field grouping merge', async () => {
  const harness = createWorkflowHarness();
  const callOrder: string[] = [];
  let notesInfoCallCount = 0;
  harness.deps.getEffectiveSentenceCardConfig = () => ({
    sentenceField: 'Sentence',
    kikuEnabled: true,
    kikuFieldGrouping: 'auto',
  });
  harness.deps.findDuplicateNote = async () => 99;
  harness.deps.client.notesInfo = async () => {
    notesInfoCallCount += 1;
    if (notesInfoCallCount === 1) {
      return [
        {
          noteId: 42,
          fields: {
            Expression: { value: 'taberu' },
            Sentence: { value: '' },
          },
        },
      ] satisfies NoteUpdateWorkflowNoteInfo[];
    }
    return [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: 'subtitle-text' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];
  };
  harness.deps.client.updateNoteFields = async (noteId, fields) => {
    callOrder.push('update');
    harness.updates.push({ noteId, fields });
  };
  harness.deps.handleFieldGroupingAuto = async (
    _originalNoteId,
    _newNoteId,
    newNoteInfo,
    _expression,
  ) => {
    callOrder.push('auto');
    assert.equal(newNoteInfo.fields.Sentence?.value, 'subtitle-text');
  };

  await harness.workflow.execute(42);

  assert.deepEqual(callOrder, ['update', 'auto']);
  assert.equal(harness.updates.length, 1);
});

test('NoteUpdateWorkflow passes animated image lead-in when syncing avif to word audio', async () => {
  const harness = createWorkflowHarness();
  let receivedLeadInSeconds = 0;

  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          ExpressionAudio: { value: '[sound:word.mp3]' },
          Sentence: { value: '' },
          Picture: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];
  harness.deps.getConfig = () => ({
    fields: {
      sentence: 'Sentence',
      image: 'Picture',
    },
    media: {
      generateImage: true,
      imageType: 'avif',
      syncAnimatedImageToWordAudio: true,
    },
    behavior: {},
  });
  harness.deps.getAnimatedImageLeadInSeconds = async () => 1.25;
  harness.deps.generateImage = async (leadInSeconds?: number) => {
    receivedLeadInSeconds = leadInSeconds ?? 0;
    return Buffer.from('image');
  };

  await harness.workflow.execute(42);

  assert.equal(receivedLeadInSeconds, 1.25);
});
