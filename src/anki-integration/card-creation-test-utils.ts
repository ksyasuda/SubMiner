import { CardCreationService } from './card-creation';
import type { AnkiConnectConfig } from '../types/anki';

export type CardCreationDeps = ConstructorParameters<typeof CardCreationService>[0];

/** Builds the minimal config the service reads; pass only the sections a test cares about. */
export function createAnkiConfig(
  overrides: {
    fields?: Record<string, string>;
    media?: Record<string, unknown>;
  } = {},
): AnkiConnectConfig {
  return {
    deck: 'Mining',
    fields: { sentence: 'Sentence', audio: 'SentenceAudio', ...overrides.fields },
    media: { generateAudio: false, generateImage: false, ...overrides.media },
    behavior: {},
    ai: false,
  } as AnkiConnectConfig;
}

/**
 * Fake CardCreationService dependencies: a bare mpv session, a client that adds note 42, no media,
 * and field resolvers that match exact field names. Override only what the test exercises.
 */
export function createCardCreationDeps(
  overrides: Partial<CardCreationDeps> = {},
): CardCreationDeps {
  return {
    getConfig: () => createAnkiConfig(),
    getAiConfig: () => ({}),
    getTimingTracker: () => ({}) as never,
    getMpvClient: () =>
      ({
        currentVideoPath: '/video.mp4',
        currentSubText: '字幕',
        currentSubStart: 1,
        currentSubEnd: 2,
        currentTimePos: 1.5,
        currentAudioStreamIndex: 0,
      }) as never,
    client: {
      addNote: async () => 42,
      addTags: async () => undefined,
      notesInfo: async () => [],
      updateNoteFields: async () => undefined,
      storeMediaFile: async () => undefined,
      findNotes: async () => [],
      retrieveMediaFile: async () => '',
      deleteNotes: async () => undefined,
    },
    mediaGenerator: {
      generateAudio: async () => null,
      generateScreenshot: async () => null,
      generateAnimatedImage: async () => null,
    },
    showOsdNotification: () => undefined,
    showUpdateResult: () => undefined,
    showStatusNotification: () => undefined,
    showNotification: async () => undefined,
    beginUpdateProgress: () => undefined,
    endUpdateProgress: () => undefined,
    withUpdateProgress: async (_message, action) => action(),
    resolveConfiguredFieldName: (noteInfo, ...preferredNames) => {
      for (const preferredName of preferredNames) {
        if (preferredName && preferredName in noteInfo.fields) return preferredName;
      }
      return null;
    },
    resolveNoteFieldName: (noteInfo, preferredName) =>
      preferredName && preferredName in noteInfo.fields ? preferredName : null,
    getAnimatedImageLeadInSeconds: async () => 0,
    extractFields: () => ({}),
    processSentence: (sentence) => sentence,
    setCardTypeFields: () => undefined,
    mergeFieldValue: (_existing, newValue) => newValue,
    formatMiscInfoPattern: () => '',
    getEffectiveSentenceCardConfig: () => ({
      model: 'Sentence',
      sentenceField: 'Sentence',
      audioField: 'SentenceAudio',
      lapisEnabled: false,
      kikuEnabled: false,
      fieldGroupingMode: 'disabled',
    }),
    getFallbackDurationSeconds: () => 10,
    appendKnownWordsFromNoteInfo: () => undefined,
    removeKnownWordNote: () => undefined,
    isUpdateInProgress: () => false,
    setUpdateInProgress: () => undefined,
    trackLastAddedNoteId: () => undefined,
    ...overrides,
  };
}

/** Anki client overrides, kept separate so a test can swap one method without redeclaring the rest. */
export function createFakeClient(
  overrides: Partial<CardCreationDeps['client']> = {},
): CardCreationDeps['client'] {
  return { ...createCardCreationDeps().client, ...overrides };
}

export function createFakeMediaGenerator(
  overrides: Partial<CardCreationDeps['mediaGenerator']> = {},
): CardCreationDeps['mediaGenerator'] {
  return { ...createCardCreationDeps().mediaGenerator, ...overrides };
}
