import path from 'node:path';

import { PLUGIN_API_VERSION, type PluginManifest } from './types';

const PLUGIN_NAME_PATTERN = /^[a-z0-9][a-z0-9._-]*$/i;

/** Plugin ids double as directory names and action prefixes, and action names follow the same rule. */
export function isValidPluginName(id: string): boolean {
  return PLUGIN_NAME_PATTERN.test(id);
}

/** Validates parsed `plugin.json` content. Returns an error message instead of throwing. */
export function parsePluginManifest(
  raw: unknown,
  fallbackName: string,
): { manifest: PluginManifest } | { error: string } {
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    return { error: 'plugin.json must contain a JSON object.' };
  }
  const value = raw as Record<string, unknown>;
  if (typeof value.main !== 'string' || value.main.trim().length === 0) {
    return { error: 'plugin.json needs a "main" entry file.' };
  }
  if (value.name !== undefined && typeof value.name !== 'string') {
    return { error: 'plugin.json "name" must be a string.' };
  }
  if (value.version !== undefined && typeof value.version !== 'string') {
    return { error: 'plugin.json "version" must be a string.' };
  }
  const apiVersion = value.apiVersion === undefined ? 1 : value.apiVersion;
  if (typeof apiVersion !== 'number' || !Number.isInteger(apiVersion) || apiVersion < 1) {
    return { error: 'plugin.json "apiVersion" must be a positive integer.' };
  }
  if (apiVersion !== PLUGIN_API_VERSION) {
    return {
      error: `Plugin requires API ${apiVersion}; this SubMiner supports API ${PLUGIN_API_VERSION}.`,
    };
  }
  return {
    manifest: {
      name: typeof value.name === 'string' && value.name.trim() ? value.name.trim() : fallbackName,
      version: typeof value.version === 'string' ? value.version : null,
      apiVersion,
      main: value.main.trim(),
    },
  };
}

/** Resolves a plugin file (entry, window HTML, preload), rejecting paths that leave the plugin directory. */
export function resolvePluginFilePath(pluginDir: string, file: string): string | null {
  const resolved = path.resolve(pluginDir, file);
  const relative = path.relative(pluginDir, resolved);
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) {
    return null;
  }
  return resolved;
}
