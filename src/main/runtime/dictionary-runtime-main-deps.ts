import type { FrequencyDictionaryLookup, JlptLevel } from '../../types';

type JlptLookup = (term: string) => JlptLevel | null;

export function createBuildDictionaryRootsMainHandler(deps: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  userDataPath: string;
  appUserDataPath: string;
  homeDir: string;
  cwd: string;
  joinPath: (...parts: string[]) => string;
}) {
  return () => [
    deps.joinPath(deps.dirname, '..', '..', 'vendor', 'yomitan-jlpt-vocab'),
    deps.joinPath(deps.appPath, 'vendor', 'yomitan-jlpt-vocab'),
    deps.joinPath(deps.resourcesPath, 'yomitan-jlpt-vocab'),
    deps.joinPath(deps.resourcesPath, 'app.asar', 'vendor', 'yomitan-jlpt-vocab'),
    deps.userDataPath,
    deps.appUserDataPath,
    deps.joinPath(deps.homeDir, '.config', 'SubMiner'),
    deps.joinPath(deps.homeDir, '.config', 'subminer'),
    deps.joinPath(deps.homeDir, 'Library', 'Application Support', 'SubMiner'),
    deps.joinPath(deps.homeDir, 'Library', 'Application Support', 'subminer'),
    deps.cwd,
  ];
}

export function createBuildFrequencyDictionaryRootsMainHandler(deps: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  userDataPath: string;
  appUserDataPath: string;
  homeDir: string;
  cwd: string;
  joinPath: (...parts: string[]) => string;
}) {
  return () => [
    deps.joinPath(deps.dirname, '..', '..', 'vendor', 'jiten_freq_global'),
    deps.joinPath(deps.dirname, '..', '..', 'vendor', 'frequency-dictionary'),
    deps.joinPath(deps.appPath, 'vendor', 'jiten_freq_global'),
    deps.joinPath(deps.appPath, 'vendor', 'frequency-dictionary'),
    deps.joinPath(deps.resourcesPath, 'jiten_freq_global'),
    deps.joinPath(deps.resourcesPath, 'frequency-dictionary'),
    deps.joinPath(deps.resourcesPath, 'app.asar', 'vendor', 'jiten_freq_global'),
    deps.joinPath(deps.resourcesPath, 'app.asar', 'vendor', 'frequency-dictionary'),
    deps.userDataPath,
    deps.appUserDataPath,
    deps.joinPath(deps.homeDir, '.config', 'SubMiner'),
    deps.joinPath(deps.homeDir, '.config', 'subminer'),
    deps.joinPath(deps.homeDir, 'Library', 'Application Support', 'SubMiner'),
    deps.joinPath(deps.homeDir, 'Library', 'Application Support', 'subminer'),
    deps.cwd,
  ];
}

export function createBuildJlptDictionaryRuntimeMainDepsHandler(deps: {
  isJlptEnabled: () => boolean;
  getDictionaryRoots: () => string[];
  getJlptDictionarySearchPaths: (deps: { getDictionaryRoots: () => string[] }) => string[];
  setJlptLevelLookup: (lookup: JlptLookup) => void;
  logInfo: (message: string) => void;
}) {
  return () => ({
    isJlptEnabled: () => deps.isJlptEnabled(),
    getSearchPaths: () =>
      deps.getJlptDictionarySearchPaths({
        getDictionaryRoots: () => deps.getDictionaryRoots(),
      }),
    setJlptLevelLookup: (lookup: JlptLookup) => deps.setJlptLevelLookup(lookup),
    log: (message: string) => deps.logInfo(`[JLPT] ${message}`),
  });
}

export function createBuildFrequencyDictionaryRuntimeMainDepsHandler(deps: {
  isFrequencyDictionaryEnabled: () => boolean;
  getDictionaryRoots: () => string[];
  getFrequencyDictionarySearchPaths: (deps: {
    getDictionaryRoots: () => string[];
    getSourcePath: () => string | undefined;
  }) => string[];
  getSourcePath: () => string | undefined;
  setFrequencyRankLookup: (lookup: FrequencyDictionaryLookup) => void;
  logInfo: (message: string) => void;
}) {
  return () => ({
    isFrequencyDictionaryEnabled: () => deps.isFrequencyDictionaryEnabled(),
    getSearchPaths: () =>
      deps.getFrequencyDictionarySearchPaths({
        getDictionaryRoots: () =>
          deps.getDictionaryRoots().filter((dictionaryRoot) => dictionaryRoot),
        getSourcePath: () => deps.getSourcePath(),
      }),
    setFrequencyRankLookup: (lookup: FrequencyDictionaryLookup) =>
      deps.setFrequencyRankLookup(lookup),
    log: (message: string) => deps.logInfo(`[Frequency] ${message}`),
  });
}
