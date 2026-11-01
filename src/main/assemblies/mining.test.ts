import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitleMiningContext } from '../../types';
import type { AppState } from '../state';
import { createMiningAssembly, type MiningAssemblyDeps } from './mining';

type AnkiIntegrationLike = NonNullable<AppState['ankiIntegration']>;
type MpvClientLike = NonNullable<AppState['mpvClient']>;

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(overrides: Partial<MiningAssemblyDeps> = {}): MiningAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      mpvClient: null,
      currentMediaPath: null,
      activeParsedSubtitleCues: [],
      ankiIntegration: null,
      subtitleTimingTracker: null,
      fieldGroupingResolver: null,
      fieldGroupingResolverSequence: 0,
    },
    getConfig: lazy('getConfig'),
    logger: { warn: lazy('warn'), error: lazy('error') },
    overlayManager: {
      getMainWindow: lazy('getMainWindow'),
      getModalWindow: lazy('getModalWindow'),
      getVisibleOverlayVisible: lazy('getVisibleOverlayVisible'),
    },
    overlayModalRuntime: {
      getRestoreVisibleOverlayOnModalClose: lazy('getRestoreVisibleOverlayOnModalClose'),
      waitForModalOpen: lazy('waitForModalOpen'),
      handleOverlayModalClosed: lazy('handleOverlayModalClosed'),
      sendToActiveOverlayWindow: lazy('sendToActiveOverlayWindow'),
    },
    youtube: {
      getCachedMediaPathForCurrentPlayback: lazy('getCachedMediaPathForCurrentPlayback'),
      shouldRequireMediaCacheForCurrentPlayback: lazy('shouldRequireMediaCacheForCurrentPlayback'),
    },
    clipboard: { readText: lazy('readText'), writeText: lazy('writeText') },
    getOverlayNotifications: lazy('getOverlayNotifications'),
    getOverlayModals: lazy('getOverlayModals'),
    setVisibleOverlayVisible: lazy('setVisibleOverlayVisible'),
    ensureOverlayStartupPrereqs: lazy('ensureOverlayStartupPrereqs'),
    ensureOverlayWindowsReadyForVisibilityActions: lazy(
      'ensureOverlayWindowsReadyForVisibilityActions',
    ),
    ankiIntegrationCallbacks: {
      recordCardsMined: lazy('recordCardsMined'),
      getMiningProvenance: lazy('getMiningProvenance'),
      onCardsReady: lazy('onCardsReady'),
      onKnownWordCacheUpdated: lazy('onKnownWordCacheUpdated'),
      generateSentenceFurigana: lazy('generateSentenceFurigana'),
    },
    ...overrides,
  };
}

const lookupContext: SubtitleMiningContext = {
  source: 'overlay',
  text: '今日はいい天気',
  startTime: 3,
  endTime: 5,
  capturedAtMs: 1,
};

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createMiningAssembly does not call its deps while constructing', () => {
  const mining = createMiningAssembly(makeDeps());
  assert.equal(typeof mining.mineSentenceCard, 'function');
  assert.equal(typeof mining.attachToAnkiIntegration, 'function');
});

test('a recorded lookup context reaches exactly one mined card', () => {
  let consumeContext: (() => SubtitleMiningContext | null) | null = null;
  const setters: string[] = [];
  const integration = new Proxy(
    {},
    {
      get: (_target, name: string) => (callback: unknown) => {
        setters.push(name);
        if (name === 'setSubtitleMiningContextConsumer') {
          consumeContext = callback as () => SubtitleMiningContext | null;
        }
      },
    },
  ) as AnkiIntegrationLike;
  const mining = createMiningAssembly(makeDeps());

  mining.attachToAnkiIntegration(integration);
  mining.recordSubtitleMiningContext(lookupContext);

  assert.equal(setters.length, 7);
  assert.ok(consumeContext);
  const consume = consumeContext as () => SubtitleMiningContext | null;
  assert.deepEqual(consume(), lookupContext);
  assert.equal(consume(), null);
});

test('mining without parsed cues falls back to the live mpv line on the video clock', () => {
  const mpvClient = {
    currentSubText: 'こんにちは',
    currentSubStart: 10,
    currentSubEnd: 12,
    currentSubDelay: 0.5,
    currentTimePos: 11,
  } as MpvClientLike;
  const mining = createMiningAssembly(
    makeDeps({ appState: { ...makeDeps().appState, mpvClient } }),
  );

  const context = mining.captureCurrentPrimarySubtitleMiningContext();

  assert.equal(context?.text, 'こんにちは');
  assert.equal(context?.startTime, 10.5);
  assert.equal(context?.endTime, 12.5);
});
