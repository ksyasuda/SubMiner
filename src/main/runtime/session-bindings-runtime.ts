import { sendMpvCommandRuntime, type MpvRuntimeClientLike } from '../../core/services';
import {
  buildPluginSessionBindingsArtifact,
  compileSessionBindings,
} from '../../core/services/session-bindings';
import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import type { CompiledSessionBinding, Keybinding, ResolvedConfig } from '../../types';
import { writeSessionBindingsArtifact } from './session-bindings-artifact';

export interface SessionBindingsRuntimeDeps {
  configDir: string;
  getKeybindings: () => Keybinding[];
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getResolvedConfig: () => ResolvedConfig;
  getMpvClient: () => MpvRuntimeClientLike | null;
  setSessionBindings: (bindings: CompiledSessionBinding[]) => void;
  setSessionBindingsInitialized: (initialized: boolean) => void;
  logWarn: (message: string) => void;
}

export function createSessionBindingsRuntime(deps: SessionBindingsRuntimeDeps): {
  persistSessionBindings: (
    bindings: CompiledSessionBinding[],
    warnings?: ReturnType<typeof compileSessionBindings>['warnings'],
  ) => void;
  refreshCurrentSessionBindings: () => void;
} {
  function resolveSessionBindingPlatform(): 'darwin' | 'win32' | 'linux' {
    if (process.platform === 'darwin') return 'darwin';
    if (process.platform === 'win32') return 'win32';
    return 'linux';
  }

  function compileCurrentSessionBindings(): {
    bindings: CompiledSessionBinding[];
    warnings: ReturnType<typeof compileSessionBindings>['warnings'];
  } {
    return compileSessionBindings({
      keybindings: deps.getKeybindings(),
      shortcuts: deps.getConfiguredShortcuts(),
      statsToggleKey: deps.getResolvedConfig().stats.toggleKey,
      statsMarkWatchedKey: deps.getResolvedConfig().stats.markWatchedKey,
      platform: resolveSessionBindingPlatform(),
      rawConfig: deps.getResolvedConfig(),
    });
  }

  function persistSessionBindings(
    bindings: CompiledSessionBinding[],
    warnings: ReturnType<typeof compileSessionBindings>['warnings'] = [],
  ): void {
    const artifact = buildPluginSessionBindingsArtifact({
      bindings,
      warnings,
      numericSelectionTimeoutMs: deps.getConfiguredShortcuts().multiCopyTimeoutMs,
    });
    writeSessionBindingsArtifact(deps.configDir, artifact);
    deps.setSessionBindings(bindings);
    deps.setSessionBindingsInitialized(true);
    const mpvClient = deps.getMpvClient();
    if (mpvClient?.connected) {
      sendMpvCommandRuntime(mpvClient, ['script-message', 'subminer-reload-session-bindings']);
    }
  }

  function refreshCurrentSessionBindings(): void {
    const compiled = compileCurrentSessionBindings();
    for (const warning of compiled.warnings) {
      deps.logWarn(`[session-bindings] ${warning.message}`);
    }
    persistSessionBindings(compiled.bindings, compiled.warnings);
  }

  return { persistSessionBindings, refreshCurrentSessionBindings };
}
