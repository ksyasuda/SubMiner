import type { ConfigValidationWarning } from '../../types';
import {
  buildConfigWarningNotificationBody,
  buildConfigWarningSummary,
  failStartupFromConfig,
} from '../config-validation';

type ReloadConfigFailure = {
  ok: false;
  path: string;
  error: string;
};

type ReloadConfigSuccess = {
  ok: true;
  path: string;
  warnings: ConfigValidationWarning[];
};

type ReloadConfigStrictResult = ReloadConfigFailure | ReloadConfigSuccess;

export type ReloadConfigRuntimeDeps = {
  reloadConfigStrict: () => ReloadConfigStrictResult;
  logInfo: (message: string) => void;
  logWarning: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  startConfigHotReload: () => void;
  refreshAnilistClientSecretState: (options: { force: boolean }) => Promise<unknown>;
  failHandlers: {
    logError: (details: string) => void;
    showErrorBox: (title: string, details: string) => void;
    quit: () => void;
  };
};

export type CriticalConfigErrorRuntimeDeps = {
  getConfigPath: () => string;
  failHandlers: {
    logError: (details: string) => void;
    showErrorBox: (title: string, details: string) => void;
    quit: () => void;
  };
};

export function createReloadConfigHandler(deps: ReloadConfigRuntimeDeps): () => void {
  return () => {
    const result = deps.reloadConfigStrict();
    if (!result.ok) {
      failStartupFromConfig(
        'SubMiner config parse error',
        `Failed to parse config file at:\n${result.path}\n\nError: ${result.error}\n\nFix the config file and restart SubMiner.`,
        deps.failHandlers,
      );
    }

    deps.logInfo(`Using config file: ${result.path}`);
    if (result.warnings.length > 0) {
      deps.logWarning(buildConfigWarningSummary(result.path, result.warnings));
      deps.showDesktopNotification('SubMiner', {
        body: buildConfigWarningNotificationBody(result.path, result.warnings),
      });
    }

    deps.startConfigHotReload();
    void deps.refreshAnilistClientSecretState({ force: true });
  };
}

export function createCriticalConfigErrorHandler(
  deps: CriticalConfigErrorRuntimeDeps,
): (errors: string[]) => never {
  return (errors: string[]) => {
    const configPath = deps.getConfigPath();
    const details = [
      `Critical config validation failed. File: ${configPath}`,
      '',
      ...errors.map((error, index) => `${index + 1}. ${error}`),
      '',
      'Fix the config file and restart SubMiner.',
    ].join('\n');
    return failStartupFromConfig('SubMiner config validation error', details, deps.failHandlers);
  };
}
