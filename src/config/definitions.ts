import { RawConfig, ResolvedConfig } from '../types';
import { CORE_DEFAULT_CONFIG } from './definitions/defaults-core';
import { IMMERSION_DEFAULT_CONFIG } from './definitions/defaults-immersion';
import { INTEGRATIONS_DEFAULT_CONFIG } from './definitions/defaults-integrations';
import { SUBTITLE_DEFAULT_CONFIG } from './definitions/defaults-subtitle';
import { buildCoreConfigOptionRegistry } from './definitions/options-core';
import { buildImmersionConfigOptionRegistry } from './definitions/options-immersion';
import { buildIntegrationConfigOptionRegistry } from './definitions/options-integrations';
import { buildSubtitleConfigOptionRegistry } from './definitions/options-subtitle';
import { buildRuntimeOptionRegistry } from './definitions/runtime-options';
import { CONFIG_TEMPLATE_SECTIONS } from './definitions/template-sections';

export { DEFAULT_KEYBINDINGS, SPECIAL_COMMANDS } from './definitions/shared';
export type {
  ConfigOptionRegistryEntry,
  ConfigTemplateSection,
  ConfigValueKind,
  RuntimeOptionRegistryEntry,
} from './definitions/shared';

const {
  subtitlePosition,
  keybindings,
  websocket,
  logging,
  texthooker,
  shortcuts,
  secondarySub,
  subsync,
  auto_start_overlay,
  bind_visible_overlay_to_mpv_sub_visibility,
  invisibleOverlay,
} = CORE_DEFAULT_CONFIG;
const { ankiConnect, jimaku, anilist, jellyfin, youtubeSubgen } = INTEGRATIONS_DEFAULT_CONFIG;
const { subtitleStyle } = SUBTITLE_DEFAULT_CONFIG;
const { immersionTracking } = IMMERSION_DEFAULT_CONFIG;

export const DEFAULT_CONFIG: ResolvedConfig = {
  subtitlePosition,
  keybindings,
  websocket,
  logging,
  texthooker,
  ankiConnect,
  shortcuts,
  secondarySub,
  subsync,
  subtitleStyle,
  auto_start_overlay,
  bind_visible_overlay_to_mpv_sub_visibility,
  jimaku,
  anilist,
  jellyfin,
  youtubeSubgen,
  invisibleOverlay,
  immersionTracking,
};

export const DEFAULT_ANKI_CONNECT_CONFIG = DEFAULT_CONFIG.ankiConnect;

export const RUNTIME_OPTION_REGISTRY = buildRuntimeOptionRegistry(DEFAULT_CONFIG);

export const CONFIG_OPTION_REGISTRY = [
  ...buildCoreConfigOptionRegistry(DEFAULT_CONFIG),
  ...buildSubtitleConfigOptionRegistry(DEFAULT_CONFIG),
  ...buildIntegrationConfigOptionRegistry(DEFAULT_CONFIG, RUNTIME_OPTION_REGISTRY),
  ...buildImmersionConfigOptionRegistry(DEFAULT_CONFIG),
];

export { CONFIG_TEMPLATE_SECTIONS };

export function deepCloneConfig(config: ResolvedConfig): ResolvedConfig {
  return JSON.parse(JSON.stringify(config)) as ResolvedConfig;
}

export function deepMergeRawConfig(base: RawConfig, patch: RawConfig): RawConfig {
  const clone = JSON.parse(JSON.stringify(base)) as Record<string, unknown>;
  const patchObject = patch as Record<string, unknown>;

  const mergeInto = (target: Record<string, unknown>, source: Record<string, unknown>): void => {
    for (const [key, value] of Object.entries(source)) {
      if (
        value !== null &&
        typeof value === 'object' &&
        !Array.isArray(value) &&
        typeof target[key] === 'object' &&
        target[key] !== null &&
        !Array.isArray(target[key])
      ) {
        mergeInto(target[key] as Record<string, unknown>, value as Record<string, unknown>);
      } else {
        target[key] = value;
      }
    }
  };

  mergeInto(clone, patchObject);
  return clone as RawConfig;
}
