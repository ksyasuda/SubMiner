import { createBindMpvMainEventHandlersHandler } from '../mpv-main-event-bindings';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from '../mpv-main-event-main-deps';
import { createMpvClientRuntimeServiceFactory } from '../mpv-client-runtime-service';
import { createUpdateMpvSubtitleRenderMetricsHandler } from '../mpv-subtitle-render-metrics';
import {
  createMecabTokenizerAndCheckHandler,
  createPrewarmSubtitleDictionariesHandler,
} from '../subtitle-tokenization-warmup';
import type { TokenizerDepsRuntimeOptions } from '../../../core/services/tokenizer';
import {
  createLaunchBackgroundWarmupTaskHandler as createLaunchBackgroundWarmupTaskFromStartup,
  createStartBackgroundWarmupsHandler as createStartBackgroundWarmupsFromStartup,
} from '../startup-warmups';
import type { ComposerInputs, ComposerOutputs } from './contracts';

type BindMpvMainEventHandlersMainDeps = Parameters<
  typeof createBuildBindMpvMainEventHandlersMainDepsHandler
>[0];
type BindMpvMainEventHandlers = ReturnType<typeof createBindMpvMainEventHandlersHandler>;
type BoundMpvClient = Parameters<BindMpvMainEventHandlers>[0];
type RuntimeMpvClient = BoundMpvClient & { connect: () => void };
type MpvClientRuntimeServiceFactoryMainDeps<TMpvClient extends RuntimeMpvClient> = Omit<
  Parameters<typeof createMpvClientRuntimeServiceFactory<TMpvClient>>[0],
  'bindEventHandlers'
>;
type UpdateMpvSubtitleRenderMetricsMainDeps = Parameters<
  typeof createUpdateMpvSubtitleRenderMetricsHandler
>[0];
// Tokenizer deps as main.ts supplies them: the core runtime options plus lookup recording and
// the character-dictionary gate applied in buildTokenizerRuntimeOptions.
type BuildTokenizerDepsMainDeps = TokenizerDepsRuntimeOptions & {
  getJlptEnabled: NonNullable<TokenizerDepsRuntimeOptions['getJlptEnabled']>;
  getCharacterDictionaryEnabled?: () => boolean;
  getFrequencyDictionaryEnabled: NonNullable<
    TokenizerDepsRuntimeOptions['getFrequencyDictionaryEnabled']
  >;
  getFrequencyDictionaryMatchMode: NonNullable<
    TokenizerDepsRuntimeOptions['getFrequencyDictionaryMatchMode']
  >;
  getFrequencyRank: NonNullable<TokenizerDepsRuntimeOptions['getFrequencyRank']>;
  getMinSentenceWordsForNPlusOne: NonNullable<
    TokenizerDepsRuntimeOptions['getMinSentenceWordsForNPlusOne']
  >;
  getYomitanGroupDebugEnabled: NonNullable<
    TokenizerDepsRuntimeOptions['getYomitanGroupDebugEnabled']
  >;
  recordLookup: (hit: boolean) => void;
};
type CreateMecabTokenizerAndCheckMainDeps = Parameters<
  typeof createMecabTokenizerAndCheckHandler
>[0];
type PrewarmSubtitleDictionariesMainDeps = Parameters<
  typeof createPrewarmSubtitleDictionariesHandler
>[0];
type LaunchBackgroundWarmupTaskMainDeps = Parameters<
  typeof createLaunchBackgroundWarmupTaskFromStartup
>[0];
type StartBackgroundWarmupsMainDeps = Omit<
  Parameters<typeof createStartBackgroundWarmupsFromStartup>[0],
  'launchTask' | 'createMecabTokenizerAndCheck' | 'prewarmSubtitleDictionaries'
>;

export type MpvRuntimeComposerOptions<
  TMpvClient extends RuntimeMpvClient,
  TTokenizerRuntimeDeps,
  TTokenizedSubtitle,
> = ComposerInputs<{
  bindMpvMainEventHandlersMainDeps: BindMpvMainEventHandlersMainDeps;
  mpvClientRuntimeServiceFactoryMainDeps: MpvClientRuntimeServiceFactoryMainDeps<TMpvClient>;
  updateMpvSubtitleRenderMetricsMainDeps: UpdateMpvSubtitleRenderMetricsMainDeps;
  tokenizer: {
    buildTokenizerDepsMainDeps: BuildTokenizerDepsMainDeps;
    createTokenizerRuntimeDeps: (deps: TokenizerDepsRuntimeOptions) => TTokenizerRuntimeDeps;
    tokenizeSubtitle: (text: string, deps: TTokenizerRuntimeDeps) => Promise<TTokenizedSubtitle>;
    createMecabTokenizerAndCheckMainDeps: CreateMecabTokenizerAndCheckMainDeps;
    prewarmSubtitleDictionariesMainDeps: PrewarmSubtitleDictionariesMainDeps;
  };
  warmups: {
    launchBackgroundWarmupTaskMainDeps: LaunchBackgroundWarmupTaskMainDeps;
    startBackgroundWarmupsMainDeps: StartBackgroundWarmupsMainDeps;
  };
}>;

export type MpvRuntimeComposerResult<
  TMpvClient extends RuntimeMpvClient,
  TTokenizedSubtitle,
> = ComposerOutputs<{
  bindMpvClientEventHandlers: BindMpvMainEventHandlers;
  createMpvClientRuntimeService: () => TMpvClient;
  updateMpvSubtitleRenderMetrics: ReturnType<typeof createUpdateMpvSubtitleRenderMetricsHandler>;
  tokenizeSubtitle: (text: string) => Promise<TTokenizedSubtitle>;
  createMecabTokenizerAndCheck: () => Promise<void>;
  prewarmSubtitleDictionaries: () => Promise<void>;
  startTokenizationWarmups: () => Promise<void>;
  isTokenizationWarmupReady: () => boolean;
  launchBackgroundWarmupTask: ReturnType<typeof createLaunchBackgroundWarmupTaskFromStartup>;
  startBackgroundWarmups: ReturnType<typeof createStartBackgroundWarmupsFromStartup>;
}>;

// Adapts main's tokenizer deps to the core runtime options: every known-word lookup is
// recorded, and name matching is off while the character dictionary is disabled. Returns a
// fresh object, since the composer wraps onTokenizationReady on it.
function buildTokenizerRuntimeOptions(
  deps: BuildTokenizerDepsMainDeps,
): TokenizerDepsRuntimeOptions {
  const {
    recordLookup,
    getCharacterDictionaryEnabled,
    getNameMatchEnabled,
    getNameMatchImagesEnabled,
    ...runtimeOptions
  } = deps;
  const characterDictionaryEnabled = (): boolean => getCharacterDictionaryEnabled?.() !== false;
  return {
    ...runtimeOptions,
    isKnownWord: (text, reading, options) => {
      const hit = deps.isKnownWord(text, reading, options);
      recordLookup(hit);
      return hit;
    },
    ...(getNameMatchEnabled
      ? { getNameMatchEnabled: () => characterDictionaryEnabled() && getNameMatchEnabled() }
      : {}),
    ...(getNameMatchImagesEnabled
      ? {
          getNameMatchImagesEnabled: () =>
            characterDictionaryEnabled() && getNameMatchImagesEnabled(),
        }
      : {}),
  };
}

export function composeMpvRuntimeHandlers<
  TMpvClient extends RuntimeMpvClient,
  TTokenizerRuntimeDeps,
  TTokenizedSubtitle,
>(
  options: MpvRuntimeComposerOptions<TMpvClient, TTokenizerRuntimeDeps, TTokenizedSubtitle>,
): MpvRuntimeComposerResult<TMpvClient, TTokenizedSubtitle> {
  const bindMpvMainEventHandlersMainDeps = createBuildBindMpvMainEventHandlersMainDepsHandler(
    options.bindMpvMainEventHandlersMainDeps,
  )();
  const bindMpvClientEventHandlers = createBindMpvMainEventHandlersHandler(
    bindMpvMainEventHandlersMainDeps,
  );

  const createMpvClientRuntimeService = createMpvClientRuntimeServiceFactory<TMpvClient>({
    ...options.mpvClientRuntimeServiceFactoryMainDeps,
    bindEventHandlers: (client) => bindMpvClientEventHandlers(client),
  });

  const updateMpvSubtitleRenderMetrics = createUpdateMpvSubtitleRenderMetricsHandler(
    options.updateMpvSubtitleRenderMetricsMainDeps,
  );

  const createMecabTokenizerAndCheck = createMecabTokenizerAndCheckHandler(
    options.tokenizer.createMecabTokenizerAndCheckMainDeps,
  );
  const prewarmSubtitleDictionaries = createPrewarmSubtitleDictionariesHandler(
    options.tokenizer.prewarmSubtitleDictionariesMainDeps,
  );
  const shouldInitializeMecabForAnnotations = (): boolean => {
    const nPlusOneEnabled =
      options.tokenizer.buildTokenizerDepsMainDeps.getNPlusOneEnabled?.() !== false;
    const knownWordsEnabled = options.tokenizer.buildTokenizerDepsMainDeps.getKnownWordsEnabled
      ? options.tokenizer.buildTokenizerDepsMainDeps.getKnownWordsEnabled() !== false
      : nPlusOneEnabled;
    const jlptEnabled = options.tokenizer.buildTokenizerDepsMainDeps.getJlptEnabled() !== false;
    const frequencyEnabled =
      options.tokenizer.buildTokenizerDepsMainDeps.getFrequencyDictionaryEnabled() !== false;
    return knownWordsEnabled || nPlusOneEnabled || jlptEnabled || frequencyEnabled;
  };
  const shouldWarmupAnnotationDictionaries = (): boolean => {
    const jlptEnabled = options.tokenizer.buildTokenizerDepsMainDeps.getJlptEnabled() !== false;
    const frequencyEnabled =
      options.tokenizer.buildTokenizerDepsMainDeps.getFrequencyDictionaryEnabled() !== false;
    return jlptEnabled || frequencyEnabled;
  };
  let tokenizationWarmupInFlight: Promise<void> | null = null;
  let tokenizationPrerequisiteWarmupInFlight: Promise<void> | null = null;
  let tokenizationPrerequisiteWarmupCompleted = false;
  let tokenizationWarmupCompleted = false;
  let tokenizationPlaybackReady = false;
  const markTokenizationPrerequisiteWarmupCompleted = (): void => {
    tokenizationPrerequisiteWarmupCompleted = true;
  };
  const markTokenizationPlaybackReady = (): void => {
    tokenizationPlaybackReady = true;
  };
  const markTokenizationWarmupCompleted = (): void => {
    tokenizationPrerequisiteWarmupCompleted = true;
    tokenizationWarmupCompleted = true;
    tokenizationPlaybackReady = true;
  };
  const backgroundWarmupCoversOnDemandTokenization = (): boolean => {
    if (!options.warmups.startBackgroundWarmupsMainDeps.shouldWarmupYomitanExtension()) {
      return false;
    }
    if (
      shouldInitializeMecabForAnnotations() &&
      !options.warmups.startBackgroundWarmupsMainDeps.shouldWarmupMecab()
    ) {
      return false;
    }
    if (
      shouldWarmupAnnotationDictionaries() &&
      !options.warmups.startBackgroundWarmupsMainDeps.shouldWarmupSubtitleDictionaries()
    ) {
      return false;
    }
    return true;
  };
  const ensureTokenizationPrerequisites = (): Promise<void> => {
    if (tokenizationPrerequisiteWarmupCompleted) {
      return Promise.resolve();
    }
    if (!tokenizationPrerequisiteWarmupInFlight) {
      tokenizationPrerequisiteWarmupInFlight = options.warmups.startBackgroundWarmupsMainDeps
        .ensureYomitanExtensionLoaded()
        .then(() => {
          markTokenizationPrerequisiteWarmupCompleted();
        })
        .finally(() => {
          tokenizationPrerequisiteWarmupInFlight = null;
        });
    }
    return tokenizationPrerequisiteWarmupInFlight;
  };
  const startTokenizationWarmups = (): Promise<void> => {
    if (tokenizationWarmupCompleted) {
      return Promise.resolve();
    }
    if (!tokenizationWarmupInFlight) {
      tokenizationWarmupInFlight = (async () => {
        const warmupTasks: Promise<unknown>[] = [ensureTokenizationPrerequisites()];
        if (
          shouldInitializeMecabForAnnotations() &&
          !options.tokenizer.createMecabTokenizerAndCheckMainDeps.getMecabTokenizer()
        ) {
          warmupTasks.push(createMecabTokenizerAndCheck().catch(() => {}));
        }
        if (shouldWarmupAnnotationDictionaries()) {
          warmupTasks.push(prewarmSubtitleDictionaries().catch(() => {}));
        }
        await Promise.all(warmupTasks);
        markTokenizationWarmupCompleted();
      })().finally(() => {
        tokenizationWarmupInFlight = null;
      });
    }
    return tokenizationWarmupInFlight;
  };
  // Built once and reused for every tokenization: per-call rebuilds create
  // fresh closures, which defeats identity-keyed caches downstream (the JLPT
  // lookup cache keys on the getJlptLevel function, and the mecab availability
  // WeakSet keys on the runtime deps instance).
  let cachedTokenizerRuntimeDeps: TTokenizerRuntimeDeps | null = null;
  const getTokenizerRuntimeDeps = (): TTokenizerRuntimeDeps => {
    if (cachedTokenizerRuntimeDeps) {
      return cachedTokenizerRuntimeDeps;
    }
    const tokenizerMainDeps = buildTokenizerRuntimeOptions(
      options.tokenizer.buildTokenizerDepsMainDeps,
    );
    const baseOnTokenizationReady = tokenizerMainDeps.onTokenizationReady;
    tokenizerMainDeps.onTokenizationReady = (tokenizedText: string): void => {
      if (!shouldWarmupAnnotationDictionaries()) {
        baseOnTokenizationReady?.(tokenizedText);
        return;
      }
      markTokenizationPlaybackReady();
      baseOnTokenizationReady?.(tokenizedText);
      if (!tokenizationWarmupCompleted) {
        void prewarmSubtitleDictionaries({ showLoadingOsd: true }).catch(() => {});
      }
    };
    cachedTokenizerRuntimeDeps = options.tokenizer.createTokenizerRuntimeDeps(tokenizerMainDeps);
    return cachedTokenizerRuntimeDeps;
  };
  const tokenizeSubtitle = async (text: string): Promise<TTokenizedSubtitle> => {
    if (!tokenizationWarmupCompleted) void startTokenizationWarmups();
    await ensureTokenizationPrerequisites();
    return options.tokenizer.tokenizeSubtitle(text, getTokenizerRuntimeDeps());
  };

  const launchBackgroundWarmupTask = createLaunchBackgroundWarmupTaskFromStartup(
    options.warmups.launchBackgroundWarmupTaskMainDeps,
  );
  const startBackgroundWarmups = createStartBackgroundWarmupsFromStartup({
    ...options.warmups.startBackgroundWarmupsMainDeps,
    launchTask: (label, task) => launchBackgroundWarmupTask(label, task),
    createMecabTokenizerAndCheck: () => createMecabTokenizerAndCheck(),
    prewarmSubtitleDictionaries: () => prewarmSubtitleDictionaries(),
    onYomitanExtensionWarmupScheduled: (promise) => {
      if (tokenizationPrerequisiteWarmupCompleted) {
        return;
      }
      const finalizedPromise = promise
        .then(() => {
          markTokenizationPrerequisiteWarmupCompleted();
        })
        .finally(() => {
          if (tokenizationPrerequisiteWarmupInFlight === finalizedPromise) {
            tokenizationPrerequisiteWarmupInFlight = null;
          }
        });
      tokenizationPrerequisiteWarmupInFlight = finalizedPromise;
    },
    onTokenizationWarmupScheduled: (promise) => {
      if (tokenizationWarmupCompleted || !backgroundWarmupCoversOnDemandTokenization()) {
        return;
      }
      const finalizedPromise = promise
        .then(() => {
          markTokenizationWarmupCompleted();
        })
        .finally(() => {
          if (tokenizationWarmupInFlight === finalizedPromise) {
            tokenizationWarmupInFlight = null;
          }
        });
      tokenizationWarmupInFlight = finalizedPromise;
    },
  });

  return {
    bindMpvClientEventHandlers: (client) => bindMpvClientEventHandlers(client),
    createMpvClientRuntimeService,
    updateMpvSubtitleRenderMetrics: (patch) => updateMpvSubtitleRenderMetrics(patch),
    tokenizeSubtitle,
    createMecabTokenizerAndCheck: () => createMecabTokenizerAndCheck(),
    prewarmSubtitleDictionaries: () => prewarmSubtitleDictionaries(),
    startTokenizationWarmups,
    isTokenizationWarmupReady: () => tokenizationPlaybackReady,
    launchBackgroundWarmupTask: (label, task) => launchBackgroundWarmupTask(label, task),
    startBackgroundWarmups: () => startBackgroundWarmups(),
  };
}
