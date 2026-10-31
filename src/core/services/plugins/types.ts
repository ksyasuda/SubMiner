import type { PluginMpvEventName } from '../../../plugin-api';
export type * from '../../../plugin-api';

export const PLUGIN_API_VERSION = 1;

/** `plugin.json` at the root of a plugin directory. */
export interface PluginManifest {
  /** Display name; defaults to the directory name. */
  name: string;
  version: string | null;
  /** Omitted by older plugins, which use API v1. */
  apiVersion: number;
  /** Entry module, relative to the plugin directory. */
  main: string;
}

/** One `plugins.<id>` entry from config. */
export interface PluginConfigEntry {
  enabled: boolean;
  settings: Record<string, unknown>;
}

const pluginMpvEvents = {
  'connection-change': true,
  'subtitle-change': true,
  'subtitle-timing': true,
  'time-pos-change': true,
  'duration-change': true,
  'pause-change': true,
  'fullscreen-change': true,
  'secondary-subtitle-change': true,
  'subtitle-track-change': true,
  'secondary-subtitle-track-change': true,
  'secondary-subtitle-delay-change': true,
  'subtitle-track-list-change': true,
  'media-path-change': true,
  'media-title-change': true,
  'secondary-subtitle-visibility': true,
  'client-message': true,
  'file-loaded': true,
  'end-file': true,
} satisfies Record<PluginMpvEventName, true>;

export const PLUGIN_MPV_EVENT_NAMES = Object.keys(pluginMpvEvents) as PluginMpvEventName[];

/** What the host needs to show a plugin window; paths are absolute and already validated. */
export interface PluginWindowRequest {
  /** Same key, same window: reopening focuses the existing one. */
  key: string;
  htmlPath: string;
  preloadPath: string | null;
  title: string;
  width: number;
  height: number;
}
