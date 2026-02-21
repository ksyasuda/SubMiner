import { createBindMpvMainEventHandlersHandler } from '../mpv-main-event-bindings';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from '../mpv-main-event-main-deps';
import { createBuildMpvClientRuntimeServiceFactoryDepsHandler } from '../mpv-client-runtime-service-main-deps';
import { createMpvClientRuntimeServiceFactory } from '../mpv-client-runtime-service';
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

type BindMpvMainEventHandlersMainDeps = Parameters<
  typeof createBuildBindMpvMainEventHandlersMainDepsHandler
>[0];
type MpvClientRuntimeServiceFactoryMainDeps = Omit<
  Parameters<typeof createBuildMpvClientRuntimeServiceFactoryDepsHandler>[0],
  'bindEventHandlers'
>;
type UpdateMpvSubtitleRenderMetricsMainDeps = Parameters<
  typeof createBuildUpdateMpvSubtitleRenderMetricsMainDepsHandler
>[0];
type BuildTokenizerDepsMainDeps = Parameters<typeof createBuildTokenizerDepsMainHandler>[0];
type TokenizerMainDeps = ReturnType<ReturnType<typeof createBuildTokenizerDepsMainHandler>>;
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

export type MpvRuntimeComposerOptions<TTokenizerRuntimeDeps, TTokenizedSubtitle> = {
  bindMpvMainEventHandlersMainDeps: BindMpvMainEventHandlersMainDeps;
  mpvClientRuntimeServiceFactoryMainDeps: MpvClientRuntimeServiceFactoryMainDeps;
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
};

export type MpvRuntimeComposerResult<TTokenizedSubtitle> = {
  bindMpvClientEventHandlers: ReturnType<typeof createBindMpvMainEventHandlersHandler>;
  createMpvClientRuntimeService: () => unknown;
  updateMpvSubtitleRenderMetrics: ReturnType<typeof createUpdateMpvSubtitleRenderMetricsHandler>;
  tokenizeSubtitle: (text: string) => Promise<TTokenizedSubtitle>;
  createMecabTokenizerAndCheck: () => Promise<void>;
  prewarmSubtitleDictionaries: () => Promise<void>;
  launchBackgroundWarmupTask: ReturnType<typeof createLaunchBackgroundWarmupTaskFromStartup>;
  startBackgroundWarmups: ReturnType<typeof createStartBackgroundWarmupsFromStartup>;
};

export function composeMpvRuntimeHandlers<TTokenizerRuntimeDeps, TTokenizedSubtitle>(
  options: MpvRuntimeComposerOptions<TTokenizerRuntimeDeps, TTokenizedSubtitle>,
): MpvRuntimeComposerResult<TTokenizedSubtitle> {
  const bindMpvMainEventHandlersMainDeps = createBuildBindMpvMainEventHandlersMainDepsHandler(
    options.bindMpvMainEventHandlersMainDeps,
  )();
  const bindMpvClientEventHandlers = createBindMpvMainEventHandlersHandler(
    bindMpvMainEventHandlersMainDeps,
  );

  const buildMpvClientRuntimeServiceFactoryMainDepsHandler =
    createBuildMpvClientRuntimeServiceFactoryDepsHandler({
      ...options.mpvClientRuntimeServiceFactoryMainDeps,
      bindEventHandlers: (client) => bindMpvClientEventHandlers(client as never),
    });
  const createMpvClientRuntimeService = (): unknown =>
    createMpvClientRuntimeServiceFactory(
      buildMpvClientRuntimeServiceFactoryMainDepsHandler() as never,
    )();

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
  const tokenizeSubtitle = async (text: string): Promise<TTokenizedSubtitle> => {
    await prewarmSubtitleDictionaries();
    return options.tokenizer.tokenizeSubtitle(
      text,
      options.tokenizer.createTokenizerRuntimeDeps(buildTokenizerDepsHandler()),
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
    })(),
  );

  return {
    bindMpvClientEventHandlers: (client) => bindMpvClientEventHandlers(client),
    createMpvClientRuntimeService,
    updateMpvSubtitleRenderMetrics: (patch) => updateMpvSubtitleRenderMetrics(patch),
    tokenizeSubtitle,
    createMecabTokenizerAndCheck: () => createMecabTokenizerAndCheck(),
    prewarmSubtitleDictionaries: () => prewarmSubtitleDictionaries(),
    launchBackgroundWarmupTask: (label, task) => launchBackgroundWarmupTask(label, task),
    startBackgroundWarmups: () => startBackgroundWarmups(),
  };
}
