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

function classifyDiff(prev: ResolvedConfig, next: ResolvedConfig): ConfigHotReloadDiff {
  const hotReloadFields: string[] = [];
  const restartRequiredFields: string[] = [];

  if (!isEqual(prev.subtitleStyle, next.subtitleStyle)) {
    hotReloadFields.push('subtitleStyle');
  }
  if (!isEqual(prev.keybindings, next.keybindings)) {
    hotReloadFields.push('keybindings');
  }
  if (!isEqual(prev.shortcuts, next.shortcuts)) {
    hotReloadFields.push('shortcuts');
  }
  if (!isEqual(prev.subtitleSidebar, next.subtitleSidebar)) {
    hotReloadFields.push('subtitleSidebar');
  }
  if (prev.secondarySub.defaultMode !== next.secondarySub.defaultMode) {
    hotReloadFields.push('secondarySub.defaultMode');
  }
  if (!isEqual(prev.ankiConnect.ai, next.ankiConnect.ai)) {
    hotReloadFields.push('ankiConnect.ai');
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

    if (key === 'secondarySub') {
      const normalizedPrev = {
        ...prev.secondarySub,
        defaultMode: next.secondarySub.defaultMode,
      };
      if (!isEqual(normalizedPrev, next.secondarySub)) {
        restartRequiredFields.push('secondarySub');
      }
      continue;
    }

    if (key === 'ankiConnect') {
      const normalizedPrev = {
        ...prev.ankiConnect,
        ai: {
          enabled: next.ankiConnect.ai.enabled,
          model: prev.ankiConnect.ai.model,
          systemPrompt: prev.ankiConnect.ai.systemPrompt,
        },
      };
      if (!isEqual(normalizedPrev, next.ankiConnect)) {
        restartRequiredFields.push('ankiConnect');
      }
      continue;
    }

    if (!isEqual(prev[key], next[key])) {
      restartRequiredFields.push(String(key));
    }
  }

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
