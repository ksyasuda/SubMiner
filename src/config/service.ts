import * as fs from 'fs';
import * as path from 'path';
import { ConfigValidationWarning, RawConfig, ResolvedConfig } from '../types';
import { DEFAULT_CONFIG, deepCloneConfig, deepMergeRawConfig } from './definitions';
import { ConfigPaths, loadRawConfig, loadRawConfigStrict } from './load';
import { resolveConfig } from './resolve';

export type ReloadConfigStrictResult =
  | {
      ok: true;
      config: ResolvedConfig;
      warnings: ConfigValidationWarning[];
      path: string;
    }
  | {
      ok: false;
      error: string;
      path: string;
    };

export class ConfigStartupParseError extends Error {
  readonly path: string;
  readonly parseError: string;

  constructor(configPath: string, parseError: string) {
    super(
      `Failed to parse startup config at ${configPath}: ${parseError}. Fix the config file and restart SubMiner.`,
    );
    this.name = 'ConfigStartupParseError';
    this.path = configPath;
    this.parseError = parseError;
  }
}

export class ConfigService {
  private readonly configPaths: ConfigPaths;
  private rawConfig: RawConfig = {};
  private resolvedConfig: ResolvedConfig = deepCloneConfig(DEFAULT_CONFIG);
  private warnings: ConfigValidationWarning[] = [];
  private configPathInUse!: string;

  constructor(configDir: string) {
    this.configPaths = {
      configDir,
      configFileJsonc: path.join(configDir, 'config.jsonc'),
      configFileJson: path.join(configDir, 'config.json'),
    };
    const loadResult = loadRawConfigStrict(this.configPaths);
    if (!loadResult.ok) {
      throw new ConfigStartupParseError(loadResult.path, loadResult.error);
    }
    this.applyResolvedConfig(loadResult.config, loadResult.path);
  }

  getConfigPath(): string {
    return this.configPathInUse;
  }

  getConfig(): ResolvedConfig {
    return deepCloneConfig(this.resolvedConfig);
  }

  getRawConfig(): RawConfig {
    return JSON.parse(JSON.stringify(this.rawConfig)) as RawConfig;
  }

  getWarnings(): ConfigValidationWarning[] {
    return [...this.warnings];
  }

  reloadConfig(): ResolvedConfig {
    const { config, path: configPath } = loadRawConfig(this.configPaths);
    return this.applyResolvedConfig(config, configPath);
  }

  reloadConfigStrict(): ReloadConfigStrictResult {
    const loadResult = loadRawConfigStrict(this.configPaths);
    if (!loadResult.ok) {
      return loadResult;
    }

    const { config, path: configPath } = loadResult;
    const resolvedConfig = this.applyResolvedConfig(config, configPath);
    return {
      ok: true,
      config: resolvedConfig,
      warnings: this.getWarnings(),
      path: configPath,
    };
  }

  saveRawConfig(config: RawConfig): void {
    if (!fs.existsSync(this.configPaths.configDir)) {
      fs.mkdirSync(this.configPaths.configDir, { recursive: true });
    }
    const targetPath = this.configPathInUse.endsWith('.json')
      ? this.configPathInUse
      : this.configPaths.configFileJsonc;
    fs.writeFileSync(targetPath, JSON.stringify(config, null, 2));
    this.applyResolvedConfig(config, targetPath);
  }

  patchRawConfig(patch: RawConfig): void {
    const merged = deepMergeRawConfig(this.getRawConfig(), patch);
    this.saveRawConfig(merged);
  }

  private applyResolvedConfig(config: RawConfig, configPath: string): ResolvedConfig {
    this.rawConfig = config;
    this.configPathInUse = configPath;
    const { resolved, warnings } = resolveConfig(config);
    this.resolvedConfig = resolved;
    this.warnings = warnings;
    return this.getConfig();
  }
}
