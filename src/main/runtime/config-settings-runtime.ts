import fs from 'node:fs';
import path from 'node:path';
import { buildConfigSettingsSnapshot } from '../../config/settings/jsonc-edit';
import type { ConfigValidationWarning, RawConfig, ResolvedConfig } from '../../types/config';
import type {
  ConfigSettingsField,
  ConfigSettingsSaveResult,
  ConfigSettingsSnapshot,
} from '../../types/settings';
import type { ReloadConfigStrictResult } from '../../config';
import {
  classifyConfigHotReloadDiff,
  type ConfigHotReloadDiff,
} from '../../core/services/config-hot-reload';
import { createSaveConfigSettingsPatchHandler } from './config-settings-save';
import {
  createOpenConfigSettingsWindowHandler,
  type ConfigSettingsWindowLike,
} from './config-settings-window';
import { isConfigSettingsPatch } from './config-settings-ipc';

export interface ConfigSettingsIpcMainLike {
  handle(channel: string, listener: (event: unknown, ...args: unknown[]) => unknown): unknown;
}

export interface ConfigSettingsIpcChannels {
  getConfigSettingsSnapshot: string;
  saveConfigSettingsPatch: string;
  openConfigSettingsFile: string;
  openConfigSettingsWindow: string;
}

export interface ConfigSettingsRuntimeDeps<TWindow extends ConfigSettingsWindowLike> {
  fields: ConfigSettingsField[];
  getConfigPath(): string;
  getRawConfig(): RawConfig;
  getConfig(): ResolvedConfig;
  getWarnings(): ConfigValidationWarning[];
  reloadConfigStrict(): ReloadConfigStrictResult;
  applyHotReload(diff: ConfigHotReloadDiff, config: ResolvedConfig): void;
  getSettingsWindow(): TWindow | null;
  setSettingsWindow(window: TWindow | null): void;
  createSettingsWindow(): TWindow;
  settingsHtmlPath: string;
  openPath(path: string): Promise<string>;
  ipcMain: ConfigSettingsIpcMainLike;
  ipcChannels: ConfigSettingsIpcChannels;
  log?: (message: string) => void;
}

export function writeTextFileAtomically(targetPath: string, content: string): void {
  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  const tempPath = path.join(
    path.dirname(targetPath),
    `.${path.basename(targetPath)}.${process.pid}.${Date.now()}.tmp`,
  );
  try {
    fs.writeFileSync(tempPath, content, 'utf-8');
    fs.renameSync(tempPath, targetPath);
  } catch (error) {
    try {
      fs.rmSync(tempPath, { force: true });
    } catch {
      // Best effort cleanup after a failed atomic write.
    }
    throw error;
  }
}

function getRestartRequiredSettingsSections(
  fields: readonly ConfigSettingsField[],
  restartRequiredFields: string[],
): string[] {
  const sections = new Set<string>();
  for (const field of fields) {
    if (
      restartRequiredFields.some(
        (restartField) =>
          field.configPath === restartField ||
          field.configPath.startsWith(`${restartField}.`) ||
          restartField.startsWith(`${field.configPath}.`),
      )
    ) {
      sections.add(field.section);
    }
  }
  return [...sections].sort();
}

export function createConfigSettingsRuntime<TWindow extends ConfigSettingsWindowLike>(
  deps: ConfigSettingsRuntimeDeps<TWindow>,
) {
  function getSnapshot(): ConfigSettingsSnapshot {
    return buildConfigSettingsSnapshot({
      configPath: deps.getConfigPath(),
      rawConfig: deps.getRawConfig(),
      resolvedConfig: deps.getConfig(),
      warnings: deps.getWarnings(),
      fields: deps.fields,
    });
  }

  const savePatch = createSaveConfigSettingsPatchHandler({
    getConfigPath: () => deps.getConfigPath(),
    getCurrentConfig: () => deps.getConfig(),
    getWarnings: () => deps.getWarnings(),
    getSnapshot,
    fileExists: (targetPath) => fs.existsSync(targetPath),
    readText: (targetPath) => fs.readFileSync(targetPath, 'utf-8'),
    writeTextAtomically: (targetPath, content) => writeTextFileAtomically(targetPath, content),
    deleteFile: (targetPath) => fs.rmSync(targetPath, { force: true }),
    reloadConfigStrict: () => deps.reloadConfigStrict(),
    classifyDiff: (previous, next) => classifyConfigHotReloadDiff(previous, next),
    applyHotReload: (diff, config) => deps.applyHotReload(diff, config),
    getRestartRequiredSections: (fields) => getRestartRequiredSettingsSections(deps.fields, fields),
  });

  function ensureConfigFileExists(): string {
    const configPath = deps.getConfigPath();
    if (!fs.existsSync(configPath)) {
      writeTextFileAtomically(configPath, '{}\n');
    }
    return configPath;
  }

  const openWindow = createOpenConfigSettingsWindowHandler({
    getSettingsWindow: deps.getSettingsWindow,
    setSettingsWindow: deps.setSettingsWindow,
    createSettingsWindow: deps.createSettingsWindow,
    settingsHtmlPath: deps.settingsHtmlPath,
    log: deps.log,
  });

  function invalidPatchResult(): ConfigSettingsSaveResult {
    return {
      ok: false,
      warnings: [],
      error: 'Invalid config settings patch.',
      hotReloadFields: [],
      restartRequiredFields: [],
      restartRequiredSections: [],
    };
  }

  function registerHandlers(): void {
    deps.ipcMain.handle(deps.ipcChannels.getConfigSettingsSnapshot, () => getSnapshot());
    deps.ipcMain.handle(deps.ipcChannels.saveConfigSettingsPatch, (_event, patch: unknown) => {
      if (!isConfigSettingsPatch(patch, deps.fields)) {
        return invalidPatchResult();
      }
      return savePatch(patch);
    });
    deps.ipcMain.handle(deps.ipcChannels.openConfigSettingsFile, async () => {
      const openError = await deps.openPath(ensureConfigFileExists());
      return openError.length === 0;
    });
    deps.ipcMain.handle(deps.ipcChannels.openConfigSettingsWindow, () => openWindow());
  }

  return {
    getSnapshot,
    savePatch,
    openWindow,
    registerHandlers,
  };
}
