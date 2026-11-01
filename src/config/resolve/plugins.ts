import { ResolveContext } from './context';
import { isObject } from './shared';

/** Keeps well-formed `plugins.<id>` entries; ids are validated again when the plugin loads. */
export function applyPluginsConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;
  if (src.plugins === undefined) return;
  if (!isObject(src.plugins)) {
    warn('plugins', src.plugins, resolved.plugins, 'Expected object keyed by plugin id.');
    return;
  }

  for (const [id, entry] of Object.entries(src.plugins)) {
    const path = `plugins.${id}`;
    if (!isObject(entry)) {
      warn(path, entry, undefined, 'Expected { "enabled": boolean, "settings"?: object }.');
      continue;
    }
    if (entry.enabled !== undefined && typeof entry.enabled !== 'boolean') {
      warn(`${path}.enabled`, entry.enabled, false, 'Expected boolean.');
    }
    if (entry.settings !== undefined && !isObject(entry.settings)) {
      warn(`${path}.settings`, entry.settings, {}, 'Expected object.');
    }
    resolved.plugins[id] = {
      enabled: entry.enabled === true,
      settings: isObject(entry.settings) ? { ...entry.settings } : {},
    };
  }
}
