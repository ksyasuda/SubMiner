import * as fs from 'fs';
import { RawConfig } from '../types/config';
import { parseConfigContent } from './parse';

export interface ConfigPaths {
  configDir: string;
  configFileJsonc: string;
  configFileJson: string;
}

export interface LoadResult {
  config: RawConfig;
  path: string;
}

export type StrictLoadResult =
  | (LoadResult & { ok: true })
  | {
      ok: false;
      error: string;
      path: string;
    };

function isObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

export function resolveExistingConfigPath(paths: ConfigPaths): string {
  if (fs.existsSync(paths.configFileJsonc)) {
    return paths.configFileJsonc;
  }
  if (fs.existsSync(paths.configFileJson)) {
    return paths.configFileJson;
  }
  return paths.configFileJsonc;
}

export function loadRawConfigStrict(paths: ConfigPaths): StrictLoadResult {
  const configPath = resolveExistingConfigPath(paths);

  if (!fs.existsSync(configPath)) {
    return { ok: true, config: {}, path: configPath };
  }

  try {
    const data = fs.readFileSync(configPath, 'utf-8');
    const parsed = parseConfigContent(configPath, data);
    return {
      ok: true,
      config: isObject(parsed) ? (parsed as RawConfig) : {},
      path: configPath,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown parse error';
    return { ok: false, error: message, path: configPath };
  }
}

export function loadRawConfig(paths: ConfigPaths): LoadResult {
  const strictResult = loadRawConfigStrict(paths);
  if (strictResult.ok) {
    return strictResult;
  }
  return { config: {}, path: strictResult.path };
}
