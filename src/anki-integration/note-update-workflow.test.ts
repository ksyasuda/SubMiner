import test from 'node:test';
import assert from 'node:assert/strict';
import {
  NoteUpdateWorkflow,
  type NoteUpdateWorkflowDeps,
  type NoteUpdateWorkflowNoteInfo,
} from './note-update-workflow';
import type { SubtitleMiningContext } from '../types/subtitle';
import type { CardKind } from '../types/anki';
import { applyCardKindFlagFields } from './card-kinds';

function setCardTypeFields(
  updatedFields: Record<string, string>,
  availableFieldNames: string[],
  cardKind: CardKind,
): void {
  applyCardKindFlagFields(
    updatedFields,
    cardKind,
    (preferredName) =>
      availableFieldNames.find((name) => name.toLowerCase() === preferredName.toLowerCase()) ??
      null,
  );
}

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
      lapisEnabled: false,
      kikuEnabled: false,
      fieldGroupingMode: 'disabled' as const,
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
    setCardTypeFields,
    resolveConfiguredFieldName: (noteInfo: NoteUpdateWorkflowNoteInfo, preferred?: string) => {
      if (!preferred) return null;
      const names = Object.keys(noteInfo.fields);
      return names.find((name) => name.toLowerCase() === preferred.toLowerCase()) ?? null;
    },
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

test('NoteUpdateWorkflow uses configured fields for word-card enrichment with Lapis and Kiku enabled', async () => {
  const harness = createWorkflowHarness();
  harness.deps.getConfig = () => ({
    fields: {
      sentence: 'Context',
      audio: 'ContextAudio',
    },
    media: {
      generateAudio: true,
      generateImage: false,
    },
    behavior: {},
  });
  harness.deps.getEffectiveSentenceCardConfig = () => ({
    sentenceField: 'Sentence',
    lapisEnabled: true,
    kikuEnabled: true,
    fieldGroupingMode: 'disabled',
  });
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          SentenceAudio: { value: '' },
          Context: { value: '' },
          ContextAudio: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];
  harness.deps.generateAudio = async () => Buffer.from('audio');

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Context: 'subtitle-text',
    ContextAudio: '[sound:audio_1.mp3]',
  });
});

test('NoteUpdateWorkflow updates sentence furigana when highlight processor changes it', async () => {
  const harness = createWorkflowHarness();
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'tokugi' },
          Sentence: { value: '' },
          SentenceFurigana: { value: '<span class="term">tokugi</span>' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];
  harness.deps.processSentenceFurigana = (sentenceFurigana) =>
    sentenceFurigana.replace('tokugi', '<b>tokugi</b>');

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Sentence: 'subtitle-text',
    SentenceFurigana: '<span class="term"><b>tokugi</b></span>',
  });
});

test('NoteUpdateWorkflow marks enriched Kiku word cards as word-and-sentence cards', async () => {
  const harness = createWorkflowHarness();
  harness.deps.getEffectiveSentenceCardConfig = () => ({
    sentenceField: 'Sentence',
    lapisEnabled: false,
    kikuEnabled: true,
    fieldGroupingMode: 'manual',
  });
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          IsWordAndSentenceCard: { value: '' },
          IsSentenceCard: { value: '' },
          IsAudioCard: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Sentence: 'subtitle-text',
    IsWordAndSentenceCard: 'x',
    IsSentenceCard: '',
    IsAudioCard: '',
  });
});

test('NoteUpdateWorkflow marks the configured word card kind instead of word-and-sentence', async () => {
  const harness = createWorkflowHarness();
  harness.deps.getEffectiveSentenceCardConfig = () => ({
    sentenceField: 'Sentence',
    lapisEnabled: false,
    kikuEnabled: true,
    fieldGroupingMode: 'manual',
    wordCardKind: 'click',
  });
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          IsWordAndSentenceCard: { value: 'x' },
          IsClickCard: { value: '' },
          IsSentenceCard: { value: '' },
          IsAudioCard: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Sentence: 'subtitle-text',
    IsClickCard: 'x',
    IsWordAndSentenceCard: '',
    IsSentenceCard: '',
    IsAudioCard: '',
  });
});

test('NoteUpdateWorkflow leaves card type flags alone when the word card kind is none', async () => {
  const harness = createWorkflowHarness();
  harness.deps.getEffectiveSentenceCardConfig = () => ({
    sentenceField: 'Sentence',
    lapisEnabled: false,
    kikuEnabled: true,
    fieldGroupingMode: 'manual',
    wordCardKind: 'none',
  });
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          IsWordAndSentenceCard: { value: '' },
          IsSentenceCard: { value: '' },
          IsAudioCard: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Sentence: 'subtitle-text',
  });
});

test('NoteUpdateWorkflow does not set Kiku card flags when Lapis and Kiku are disabled', async () => {
  const harness = createWorkflowHarness();
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          IsWordAndSentenceCard: { value: '' },
          IsSentenceCard: { value: '' },
          IsAudioCard: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Sentence: 'subtitle-text',
  });
});

test('NoteUpdateWorkflow preserves explicit sentence card type during sentence enrichment', async () => {
  const harness = createWorkflowHarness();
  harness.deps.getEffectiveSentenceCardConfig = () => ({
    sentenceField: 'Sentence',
    lapisEnabled: true,
    kikuEnabled: false,
    fieldGroupingMode: 'disabled',
  });
  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'sentence expression' },
          Sentence: { value: '' },
          IsWordAndSentenceCard: { value: '' },
          IsSentenceCard: { value: 'x' },
          IsAudioCard: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.deepEqual(harness.updates[0]?.fields, {
    Sentence: 'subtitle-text',
  });
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
    lapisEnabled: false,
    kikuEnabled: true,
    fieldGroupingMode: 'auto',
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

test('NoteUpdateWorkflow uses subtitle sidebar context for sentence media timing', async () => {
  const harness = createWorkflowHarness();
  const sidebarContext = {
    source: 'subtitle-sidebar' as const,
    text: 'sidebar previous line',
    startTime: 10,
    endTime: 12,
    capturedAtMs: 123,
  };
  let audioContext: unknown = null;
  let imageContext: unknown = null;
  let miscInfoStartTime: number | undefined;

  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: 'sidebar previous line' },
          SentenceAudio: { value: '' },
          Picture: { value: '' },
          MiscInfo: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];
  harness.deps.getConfig = () => ({
    fields: {
      sentence: 'Sentence',
      audio: 'SentenceAudio',
      image: 'Picture',
      miscInfo: 'MiscInfo',
    },
    media: {
      generateAudio: true,
      generateImage: true,
      imageType: 'avif',
    },
    behavior: {},
  });
  harness.deps.getCurrentSubtitleText = () => 'current primary line';
  harness.deps.getCurrentSubtitleStart = () => 20;
  harness.deps.generateAudio = async (context?: SubtitleMiningContext) => {
    audioContext = context ?? null;
    return Buffer.from('audio');
  };
  harness.deps.generateImage = async (_leadInSeconds?: number, context?: SubtitleMiningContext) => {
    imageContext = context ?? null;
    return Buffer.from('image');
  };
  harness.deps.formatMiscInfoPattern = (_fallbackFilename, startTimeSeconds) => {
    miscInfoStartTime = startTimeSeconds;
    return `start:${startTimeSeconds}`;
  };
  (
    harness.deps as NoteUpdateWorkflowDeps & {
      consumeSubtitleMiningContext: () => typeof sidebarContext | null;
    }
  ).consumeSubtitleMiningContext = () => sidebarContext;

  await harness.workflow.execute(42);

  assert.equal(harness.updates.length, 1);
  assert.equal(harness.updates[0]?.fields.Sentence, 'sidebar previous line');
  assert.deepEqual(audioContext, sidebarContext);
  assert.deepEqual(imageContext, sidebarContext);
  assert.equal(miscInfoStartTime, 10);
});

test('NoteUpdateWorkflow snapshots one media range for audio and image without a mining context', async () => {
  const harness = createWorkflowHarness();
  const capturedContext: SubtitleMiningContext = {
    source: 'overlay',
    text: 'subtitle-text',
    startTime: 31.5,
    endTime: 34.25,
  };
  let captureCalls = 0;
  let audioContext: SubtitleMiningContext | null = null;
  let imageContext: SubtitleMiningContext | null = null;
  let miscInfoStartTime: number | undefined;

  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          SentenceAudio: { value: '' },
          Picture: { value: '' },
          MiscInfo: { value: '' },
        },
      },
    ] satisfies NoteUpdateWorkflowNoteInfo[];
  harness.deps.getConfig = () => ({
    fields: {
      sentence: 'Sentence',
      audio: 'SentenceAudio',
      image: 'Picture',
      miscInfo: 'MiscInfo',
    },
    media: {
      generateAudio: true,
      generateImage: true,
      imageType: 'avif',
    },
    behavior: {},
  });
  harness.deps.captureSubtitleMediaContext = () => {
    captureCalls += 1;
    return capturedContext;
  };
  harness.deps.generateAudio = async (context?: SubtitleMiningContext) => {
    audioContext = context ?? null;
    return Buffer.from('audio');
  };
  harness.deps.generateImage = async (_leadInSeconds?: number, context?: SubtitleMiningContext) => {
    imageContext = context ?? null;
    return Buffer.from('image');
  };
  harness.deps.formatMiscInfoPattern = (_fallbackFilename, startTimeSeconds) => {
    miscInfoStartTime = startTimeSeconds;
    return `start:${startTimeSeconds}`;
  };

  await harness.workflow.execute(42);

  assert.equal(captureCalls, 1);
  assert.deepEqual(audioContext, capturedContext);
  assert.deepEqual(imageContext, capturedContext);
  assert.equal(miscInfoStartTime, 31.5);
});

test('NoteUpdateWorkflow queues media updates when YouTube cache is pending', async () => {
  const harness = createWorkflowHarness();
  const queuedUpdates: Array<{
    noteId: number;
    noteInfo: NoteUpdateWorkflowNoteInfo;
    context?: SubtitleMiningContext;
    label: string | number;
  }> = [];
  const mediaCalls: string[] = [];

  harness.deps.client.notesInfo = async () =>
    [
      {
        noteId: 42,
        fields: {
          Expression: { value: 'taberu' },
          Sentence: { value: '' },
          SentenceAudio: { value: '' },
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
      generateAudio: true,
      generateImage: true,
    },
    behavior: {},
  });
  harness.deps.generateAudio = async () => {
    mediaCalls.push('audio');
    return Buffer.from('audio');
  };
  harness.deps.generateImage = async () => {
    mediaCalls.push('image');
    return Buffer.from('image');
  };
  harness.deps.queuePendingYoutubeMediaUpdate = async (job) => {
    queuedUpdates.push(job);
    return true;
  };

  await harness.workflow.execute(42);

  assert.deepEqual(mediaCalls, []);
  assert.equal(queuedUpdates.length, 1);
  assert.equal(queuedUpdates[0]?.noteId, 42);
  assert.equal(queuedUpdates[0]?.label, 'taberu');
  assert.equal(queuedUpdates[0]?.context, undefined);
  assert.deepEqual(harness.updates, [{ noteId: 42, fields: { Sentence: 'subtitle-text' } }]);
});
