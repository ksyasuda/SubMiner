import type { Args, LauncherJellyfinConfig, PluginRuntimeConfig } from '../types.js';
import type { ProcessAdapter } from '../process-adapter.js';

export interface LauncherCommandContext {
  args: Args;
  scriptPath: string;
  scriptName: string;
  mpvSocketPath: string;
  pluginRuntimeConfig: PluginRuntimeConfig;
  appPath: string | null;
  launcherJellyfinConfig: LauncherJellyfinConfig;
  processAdapter: ProcessAdapter;
}
