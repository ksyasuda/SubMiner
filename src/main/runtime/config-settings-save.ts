import type { ReloadConfigStrictResult } from '../../config';
import { applyConfigSettingsPatchToContent } from '../../config/settings/jsonc-edit';
import type { ConfigValidationWarning, ResolvedConfig } from '../../types/config';
import type {
  ConfigSettingsPatch,
  ConfigSettingsSaveResult,
  ConfigSettingsSnapshot,
} from '../../types/settings';

export interface ConfigSettingsHotReloadDiff {
  hotReloadFields: string[];
  restartRequiredFields: string[];
}

export interface ConfigSettingsSaveDeps {
  getConfigPath(): string;
  getCurrentConfig(): ResolvedConfig;
  getWarnings(): ConfigValidationWarning[];
  getSnapshot(): ConfigSettingsSnapshot;
  fileExists(path: string): boolean;
  readText(path: string): string;
  writeTextAtomically(path: string, content: string): void;
  deleteFile?(path: string): void;
  reloadConfigStrict(): ReloadConfigStrictResult;
  classifyDiff(prev: ResolvedConfig, next: ResolvedConfig): ConfigSettingsHotReloadDiff;
  applyHotReload(diff: ConfigSettingsHotReloadDiff, config: ResolvedConfig): void;
  getRestartRequiredSections(restartRequiredFields: string[]): string[];
}

export function createSaveConfigSettingsPatchHandler(deps: ConfigSettingsSaveDeps) {
  return (patch: ConfigSettingsPatch): ConfigSettingsSaveResult => {
    if (patch.operations.length === 0) {
      return {
        ok: true,
        snapshot: deps.getSnapshot(),
        hotReloadFields: [],
        restartRequiredFields: [],
        restartRequiredSections: [],
      };
    }

    const configPath = deps.getConfigPath();
    const previousConfig = deps.getCurrentConfig();
    const previousWarnings = deps.getWarnings();
    const hadExistingConfig = deps.fileExists(configPath);
    const content = hadExistingConfig ? deps.readText(configPath) : '{}\n';
    const candidate = applyConfigSettingsPatchToContent({
      content,
      operations: patch.operations,
      previousWarnings,
    });

    if (!candidate.ok) {
      return {
        ok: false,
        warnings: candidate.warnings,
        error: candidate.error,
        hotReloadFields: [],
        restartRequiredFields: [],
        restartRequiredSections: [],
      };
    }

    deps.writeTextAtomically(configPath, candidate.content);
    const reloadResult = deps.reloadConfigStrict();
    if (!reloadResult.ok) {
      if (hadExistingConfig) {
        deps.writeTextAtomically(configPath, content);
      } else if (deps.deleteFile) {
        deps.deleteFile(configPath);
      } else {
        deps.writeTextAtomically(configPath, content);
      }
      return {
        ok: false,
        warnings: [],
        error: reloadResult.error,
        hotReloadFields: [],
        restartRequiredFields: [],
        restartRequiredSections: [],
      };
    }

    const diff = deps.classifyDiff(previousConfig, reloadResult.config);
    if (diff.hotReloadFields.length > 0) {
      deps.applyHotReload(diff, reloadResult.config);
    }

    return {
      ok: true,
      snapshot: deps.getSnapshot(),
      warnings: reloadResult.warnings,
      hotReloadFields: diff.hotReloadFields,
      restartRequiredFields: diff.restartRequiredFields,
      restartRequiredSections: deps.getRestartRequiredSections(diff.restartRequiredFields),
    };
  };
}
