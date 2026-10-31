import type { ResolvedConfig } from '../../types';
import type { createLogger } from '../../logger';
import {
  createFrequencyDictionaryRuntimeService,
  getFrequencyDictionarySearchPaths,
} from '../frequency-dictionary-runtime';
import { createJlptDictionaryRuntimeService, getJlptDictionarySearchPaths } from '../jlpt-runtime';
import { createBuildDictionaryRootsHandler } from '../runtime/dictionary-roots';
import type { AppState } from '../state';

/**
 * What the bundled annotation dictionaries need. `roots` holds the install and user paths
 * searched for each dictionary's vendor folder; `main.ts` resolves them when it builds this.
 */
export interface DictionariesAssemblyDeps {
  roots: Omit<Parameters<typeof createBuildDictionaryRootsHandler>[0], 'vendorDirName'>;
  appState: Pick<AppState, 'jlptLevelLookup' | 'frequencyRankLookup'>;
  getConfig: () => ResolvedConfig;
  logger: Pick<ReturnType<typeof createLogger>, 'info'>;
}

export type DictionariesAssembly = ReturnType<typeof createDictionariesAssembly>;

/**
 * Owns the JLPT and frequency dictionary lookups that subtitle annotations read. `main.ts` builds
 * it once; tokenization calls `jlpt`/`frequency` to load each lookup on first use.
 */
export function createDictionariesAssembly(deps: DictionariesAssemblyDeps) {
  const { appState, logger } = deps;

  const buildJlptDictionaryRootsHandler = createBuildDictionaryRootsHandler({
    ...deps.roots,
    vendorDirName: 'yomitan-jlpt-vocab',
  });
  const buildFrequencyDictionaryRootsHandler = createBuildDictionaryRootsHandler({
    ...deps.roots,
    vendorDirName: 'frequency-dictionary',
  });

  const jlpt = createJlptDictionaryRuntimeService({
    isJlptEnabled: () => deps.getConfig().subtitleStyle.enableJlpt,
    getSearchPaths: () =>
      getJlptDictionarySearchPaths({ getDictionaryRoots: () => buildJlptDictionaryRootsHandler() }),
    setJlptLevelLookup: (lookup) => {
      appState.jlptLevelLookup = lookup;
    },
    log: (message) => logger.info(`[JLPT] ${message}`),
  });

  const frequency = createFrequencyDictionaryRuntimeService({
    isFrequencyDictionaryEnabled: () => deps.getConfig().subtitleStyle.frequencyDictionary.enabled,
    getSearchPaths: () =>
      getFrequencyDictionarySearchPaths({
        getDictionaryRoots: () =>
          buildFrequencyDictionaryRootsHandler().filter((dictionaryRoot) => dictionaryRoot),
        getSourcePath: () => deps.getConfig().subtitleStyle.frequencyDictionary.sourcePath,
      }),
    setFrequencyRankLookup: (lookup) => {
      appState.frequencyRankLookup = lookup;
    },
    log: (message) => logger.info(`[Frequency] ${message}`),
  });

  return { jlpt, frequency };
}
