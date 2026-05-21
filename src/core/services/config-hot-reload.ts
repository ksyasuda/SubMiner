import { type ReloadConfigStrictResult } from '../../config';
import type { ConfigValidationWarning } from '../../types';
import type { ResolvedConfig } from '../../types';

export interface ConfigHotReloadDiff {
  hotReloadFields: string[];
  restartRequiredFields: string[];
}

export interface ConfigHotReloadRuntimeDeps {
  getCurrentConfig: () => ResolvedConfig;
  reloadConfigStrict: () => ReloadConfigStrictResult;
  watchConfigPath: (configPath: string, onChange: () => void) => { close: () => void };
  setTimeout: (callback: () => void, delayMs: number) => NodeJS.Timeout;
  clearTimeout: (timeout: NodeJS.Timeout) => void;
  debounceMs?: number;
  onHotReloadApplied: (diff: ConfigHotReloadDiff, config: ResolvedConfig) => void;
  onRestartRequired: (fields: string[]) => void;
  onInvalidConfig: (message: string) => void;
  onValidationWarnings: (configPath: string, warnings: ConfigValidationWarning[]) => void;
}

export interface ConfigHotReloadRuntime {
  start: () => void;
  stop: () => void;
}

function isEqual(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function pathStartsWith(path: string, prefix: string): boolean {
  return path === prefix || path.startsWith(`${prefix}.`);
}

function collectChangedPaths(prev: unknown, next: unknown, prefix = ''): string[] {
  if (isEqual(prev, next)) {
    return [];
  }

  if (!isRecord(prev) || !isRecord(next)) {
    return prefix ? [prefix] : [];
  }

  const keys = new Set([...Object.keys(prev), ...Object.keys(next)]);
  return [...keys].flatMap((key) =>
    collectChangedPaths(prev[key], next[key], prefix ? `${prefix}.${key}` : key),
  );
}

const HOT_RELOAD_ROOTS = ['subtitleStyle', 'keybindings', 'shortcuts', 'subtitleSidebar'] as const;

const HOT_RELOAD_EXACT_OR_PREFIX_PATHS = [
  'secondarySub.defaultMode',
  'mpv.aniskipButtonKey',
  'ankiConnect.ai.enabled',
  'stats.toggleKey',
  'stats.markWatchedKey',
  'logging.level',
  'youtube.primarySubLanguages',
  'jimaku',
  'subsync',
  'ankiConnect.behavior.autoUpdateNewCards',
  'ankiConnect.knownWords.highlightEnabled',
  'ankiConnect.knownWords.refreshMinutes',
  'ankiConnect.knownWords.addMinedWordsImmediately',
  'ankiConnect.knownWords.matchMode',
  'ankiConnect.knownWords.decks',
  'ankiConnect.nPlusOne.enabled',
  'ankiConnect.nPlusOne.minSentenceWords',
  'ankiConnect.fields.word',
  'ankiConnect.fields.audio',
  'ankiConnect.fields.image',
  'ankiConnect.fields.sentence',
  'ankiConnect.fields.miscInfo',
  'ankiConnect.isLapis.sentenceCardModel',
  'ankiConnect.isKiku.fieldGrouping',
] as const;

function hotReloadFieldForChangedPath(path: string): string | null {
  for (const root of HOT_RELOAD_ROOTS) {
    if (pathStartsWith(path, root)) {
      return root;
    }
  }

  for (const hotPath of HOT_RELOAD_EXACT_OR_PREFIX_PATHS) {
    if (pathStartsWith(path, hotPath)) {
      return hotPath === 'jimaku' || hotPath === 'subsync' ? path : hotPath;
    }
  }

  return null;
}

function classifyDiff(prev: ResolvedConfig, next: ResolvedConfig): ConfigHotReloadDiff {
  const hotReloadFields: string[] = [];
  const restartRequiredFields: string[] = [];
  const hotReloadFieldSet = new Set<string>();
  const changedPaths = collectChangedPaths(prev, next);

  for (const path of changedPaths) {
    const hotReloadField = hotReloadFieldForChangedPath(path);
    if (hotReloadField) {
      hotReloadFieldSet.add(hotReloadField);
    }
  }

  const keys = new Set([
    ...(Object.keys(prev) as Array<keyof ResolvedConfig>),
    ...(Object.keys(next) as Array<keyof ResolvedConfig>),
  ]);

  for (const key of keys) {
    if (
      key === 'subtitleStyle' ||
      key === 'keybindings' ||
      key === 'shortcuts' ||
      key === 'subtitleSidebar'
    ) {
      continue;
    }

    const changedPathsForKey = changedPaths.filter((path) => pathStartsWith(path, String(key)));
    const hasRestartRequiredChange = changedPathsForKey.some(
      (path) => !hotReloadFieldForChangedPath(path),
    );
    if (hasRestartRequiredChange) {
      restartRequiredFields.push(String(key));
    }
  }

  hotReloadFields.push(...hotReloadFieldSet);
  return { hotReloadFields, restartRequiredFields };
}

export function createConfigHotReloadRuntime(
  deps: ConfigHotReloadRuntimeDeps,
): ConfigHotReloadRuntime {
  let watcher: { close: () => void } | null = null;
  let timer: NodeJS.Timeout | null = null;
  let watchedPath: string | null = null;
  const debounceMs = deps.debounceMs ?? 250;

  const reloadWithDiff = () => {
    const prev = deps.getCurrentConfig();
    const result = deps.reloadConfigStrict();
    if (!result.ok) {
      deps.onInvalidConfig(`Config reload failed: ${result.error}`);
      return;
    }

    if (watchedPath !== result.path) {
      watchPath(result.path);
    }

    if (result.warnings.length > 0) {
      deps.onValidationWarnings(result.path, result.warnings);
    }

    const diff = classifyDiff(prev, result.config);
    if (diff.hotReloadFields.length > 0) {
      deps.onHotReloadApplied(diff, result.config);
    }
    if (diff.restartRequiredFields.length > 0) {
      deps.onRestartRequired(diff.restartRequiredFields);
    }
  };

  const scheduleReload = () => {
    if (timer) {
      deps.clearTimeout(timer);
    }
    timer = deps.setTimeout(() => {
      timer = null;
      reloadWithDiff();
    }, debounceMs);
  };

  const watchPath = (configPath: string) => {
    watcher?.close();
    watcher = deps.watchConfigPath(configPath, scheduleReload);
    watchedPath = configPath;
  };

  return {
    start: () => {
      if (watcher) {
        return;
      }
      const result = deps.reloadConfigStrict();
      if (!result.ok) {
        deps.onInvalidConfig(`Config watcher startup failed: ${result.error}`);
        return;
      }
      watchPath(result.path);
    },
    stop: () => {
      if (timer) {
        deps.clearTimeout(timer);
        timer = null;
      }
      watcher?.close();
      watcher = null;
      watchedPath = null;
    },
  };
}

export { classifyDiff as classifyConfigHotReloadDiff };
