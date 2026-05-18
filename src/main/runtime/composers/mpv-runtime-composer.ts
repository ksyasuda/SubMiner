import { createBindMpvMainEventHandlersHandler } from '../mpv-main-event-bindings';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from '../mpv-main-event-main-deps';
import { createBuildMpvClientRuntimeServiceFactoryDepsHandler } from '../mpv-client-runtime-service-main-deps';
import { createMpvClientRuntimeServiceFactory } from '../mpv-client-runtime-service';
import type { MpvClientRuntimeServiceOptions } from '../mpv-client-runtime-service';
import type { Config } from '../../../types';
import { createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler } from '../mpv-subtitle-render-metrics-main-deps';
import { createUpdateMpvSubtitleRenderMetricsHandler } from '../mpv-subtitle-render-metrics';
import {
  createBuildTokenizerDepsMainHandler,
  createCreateMecabTokenizerAndCheckMainHandler,
  createPrewarmSubtitleDictionariesMainHandler,
} from '../subtitle-tokenization-main-deps';
import {
  createBuildLaunchBackgroundWarmupTaskMainDepsHandler,
  createBuildStartBackgroundWarmupsMainDepsHandler,
} from '../startup-warmups-main-deps';
import {
  createLaunchBackgroundWarmupTaskHandler as createLaunchBackgroundWarmupTaskFromStartup,
  createStartBackgroundWarmupsHandler as createStartBackgroundWarmupsFromStartup,
} from '../startup-warmups';
import type { BuiltMainDeps, ComposerInputs, ComposerOutputs } from './contracts';

type BindMpvMainEventHandlersMainDeps = Parameters<
  typeof createBuildBindMpvMainEventHandlersMainDepsHandler
>[0];
type BindMpvMainEventHandlers = ReturnType<typeof createBindMpvMainEventHandlersHandler>;
type BoundMpvClient = Parameters<BindMpvMainEventHandlers>[0];
type RuntimeMpvClient = BoundMpvClient & { connect: () => void };
type MpvClientRuntimeServiceFactoryMainDeps<TMpvClient extends RuntimeMpvClient> = Omit<
  Parameters<
    typeof createBuildMpvClientRuntimeServiceFactoryDepsHandler<
      TMpvClient,
      Config,
      MpvClientRuntimeServiceOptions
    >
  >[0],
  'bindEventHandlers'
>;
type UpdateMpvSubtitleRenderMetricsMainDeps = Parameters<
  typeof createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler
>[0];
type BuildTokenizerDepsMainDeps = Parameters<typeof createBuildTokenizerDepsMainHandler>[0];
type TokenizerMainDeps = BuiltMainDeps<typeof createBuildTokenizerDepsMainHandler>;
type CreateMecabTokenizerAndCheckMainDeps = Parameters<
  typeof createCreateMecabTokenizerAndCheckMainHandler
>[0];
type PrewarmSubtitleDictionariesMainDeps = Parameters<
  typeof createPrewarmSubtitleDictionariesMainHandler
>[0];
type LaunchBackgroundWarmupTaskMainDeps = Parameters<
  typeof createBuildLaunchBackgroundWarmupTaskMainDepsHandler
>[0];
type StartBackgroundWarmupsMainDeps = Omit<
  Parameters<typeof createBuildStartBackgroundWarmupsMainDepsHandler>[0],
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
    createTokenizerRuntimeDeps: (deps: TokenizerMainDeps) => TTokenizerRuntimeDeps;
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

  const buildMpvClientRuntimeServiceFactoryMainDepsHandler =
    createBuildMpvClientRuntimeServiceFactoryDepsHandler<
      TMpvClient,
      Config,
      MpvClientRuntimeServiceOptions
    >({
      ...options.mpvClientRuntimeServiceFactoryMainDeps,
      bindEventHandlers: (client) => bindMpvClientEventHandlers(client),
    });
  const createMpvClientRuntimeService = (): TMpvClient =>
    createMpvClientRuntimeServiceFactory(buildMpvClientRuntimeServiceFactoryMainDepsHandler())();

  const updateMpvSubtitleRenderMetrics = createUpdateMpvSubtitleRenderMetricsHandler(
    createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler(
      options.updateMpvSubtitleRenderMetricsMainDeps,
    )(),
  );

  const buildTokenizerDepsHandler = createBuildTokenizerDepsMainHandler(
    options.tokenizer.buildTokenizerDepsMainDeps,
  );
  const createMecabTokenizerAndCheck = createCreateMecabTokenizerAndCheckMainHandler(
    options.tokenizer.createMecabTokenizerAndCheckMainDeps,
  );
  const prewarmSubtitleDictionaries = createPrewarmSubtitleDictionariesMainHandler(
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
  const tokenizeSubtitle = async (text: string): Promise<TTokenizedSubtitle> => {
    if (!tokenizationWarmupCompleted) void startTokenizationWarmups();
    await ensureTokenizationPrerequisites();
    const tokenizerMainDeps = buildTokenizerDepsHandler();
    if (shouldWarmupAnnotationDictionaries()) {
      const onTokenizationReady = tokenizerMainDeps.onTokenizationReady;
      tokenizerMainDeps.onTokenizationReady = (tokenizedText: string): void => {
        markTokenizationPlaybackReady();
        onTokenizationReady?.(tokenizedText);
        if (!tokenizationWarmupCompleted) {
          void prewarmSubtitleDictionaries({ showLoadingOsd: true }).catch(() => {});
        }
      };
    }
    return options.tokenizer.tokenizeSubtitle(
      text,
      options.tokenizer.createTokenizerRuntimeDeps(tokenizerMainDeps),
    );
  };

  const launchBackgroundWarmupTask = createLaunchBackgroundWarmupTaskFromStartup(
    createBuildLaunchBackgroundWarmupTaskMainDepsHandler(
      options.warmups.launchBackgroundWarmupTaskMainDeps,
    )(),
  );
  const startBackgroundWarmups = createStartBackgroundWarmupsFromStartup(
    createBuildStartBackgroundWarmupsMainDepsHandler({
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
    })(),
  );

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
