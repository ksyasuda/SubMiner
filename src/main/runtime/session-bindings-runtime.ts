import { sendMpvCommandRuntime, type MpvRuntimeClientLike } from '../../core/services';
import {
  buildPluginSessionBindingsArtifact,
  compileSessionBindings,
} from '../../core/services/session-bindings';
import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import type { CompiledSessionBinding, Keybinding, ResolvedConfig } from '../../types';
import { writeSessionBindingsArtifact } from './session-bindings-artifact';
import { parseMpvInputBindingKeys } from '../../shared/mpv-input-bindings';
import {
  reserveMpvSequencePrefixes,
  resolveSessionSequenceConflicts,
} from '../../shared/session-key-sequences';
import type { SessionBindingWarning } from '../../types/session-bindings';

export interface SessionBindingsRuntimeDeps {
  configDir: string;
  getKeybindings: () => Keybinding[];
  getConfiguredShortcuts: () => ConfiguredShortcuts;
  getResolvedConfig: () => ResolvedConfig;
  getMpvClient: () =>
    | (MpvRuntimeClientLike & { requestProperty: (name: string) => Promise<unknown> })
    | null;
  setSessionBindings: (bindings: CompiledSessionBinding[]) => void;
  setSessionBindingsInitialized: (initialized: boolean) => void;
  logWarn: (message: string, details?: unknown) => void;
  onBindingsChanged?: (bindings: CompiledSessionBinding[]) => void;
  onWarning?: (warning: SessionBindingWarning) => void;
}

export function createSessionBindingsRuntime(deps: SessionBindingsRuntimeDeps): {
  persistSessionBindings: (
    bindings: CompiledSessionBinding[],
    warnings?: ReturnType<typeof compileSessionBindings>['warnings'],
  ) => void;
  refreshCurrentSessionBindings: () => void;
  refreshMpvSessionBindings: () => Promise<void>;
} {
  let sourceBindings: CompiledSessionBinding[] = [];
  let sourceWarnings: SessionBindingWarning[] = [];
  let nativeSnapshot: {
    client: ReturnType<SessionBindingsRuntimeDeps['getMpvClient']>;
    keys: string[];
  } | null = null;
  let pending: {
    client: ReturnType<SessionBindingsRuntimeDeps['getMpvClient']>;
    promise: Promise<void>;
  } | null = null;
  let publishedSignature: string | null = null;
  let reportedWarnings = new Set<string>();
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
    sourceBindings = bindings;
    sourceWarnings = warnings;
    publishBindings();
  }

  function publishBindings(): void {
    const client = deps.getMpvClient();
    const keys = client?.connected && nativeSnapshot?.client === client ? nativeSnapshot.keys : [];
    const result = resolveSessionSequenceConflicts(
      sourceBindings,
      reserveMpvSequencePrefixes(keys),
    );
    const warnings = [...sourceWarnings, ...result.warnings];
    const signature = JSON.stringify([
      result.bindings,
      warnings,
      deps.getConfiguredShortcuts().multiCopyTimeoutMs,
    ]);
    if (signature === publishedSignature) return;
    const artifact = buildPluginSessionBindingsArtifact({
      bindings: result.bindings,
      warnings,
      numericSelectionTimeoutMs: deps.getConfiguredShortcuts().multiCopyTimeoutMs,
    });
    try {
      writeSessionBindingsArtifact(deps.configDir, artifact);
    } catch (error) {
      deps.logWarn('[session-bindings] Failed to write session bindings artifact');
      throw error;
    }
    publishedSignature = signature;
    deps.setSessionBindings(result.bindings);
    deps.setSessionBindingsInitialized(true);
    const nextWarnings = new Set(warnings.map((warning) => warning.message));
    for (const warning of warnings) {
      if (reportedWarnings.has(warning.message)) continue;
      deps.logWarn(`[session-bindings] ${warning.message}`);
      deps.onWarning?.(warning);
    }
    reportedWarnings = nextWarnings;
    const mpvClient = deps.getMpvClient();
    if (mpvClient?.connected) {
      try {
        sendMpvCommandRuntime(mpvClient, ['script-message', 'subminer-reload-session-bindings']);
      } catch (error) {
        deps.logWarn('[session-bindings] Failed to notify mpv to reload session bindings', error);
      }
    }
    deps.onBindingsChanged?.(result.bindings);
  }

  async function refreshMpvSessionBindings(): Promise<void> {
    const client = deps.getMpvClient();
    if (!client?.connected) {
      nativeSnapshot = null;
      publishBindings();
      return;
    }
    if (pending?.client === client) return pending.promise;
    const promise = (async () => {
      try {
        const raw = await client.requestProperty('input-bindings');
        if (client !== deps.getMpvClient() || !client.connected) return;
        nativeSnapshot = { client, keys: parseMpvInputBindingKeys(raw, { includeIgnored: false }) };
        publishBindings();
      } catch {
        // Keep the last successful snapshot if discovery is temporarily unavailable.
      }
    })();
    const request = { client, promise };
    pending = request;
    try {
      await promise;
    } finally {
      if (pending === request) pending = null;
    }
  }

  function refreshCurrentSessionBindings(): void {
    const compiled = compileCurrentSessionBindings();
    persistSessionBindings(compiled.bindings, compiled.warnings);
    void refreshMpvSessionBindings();
  }

  return { persistSessionBindings, refreshCurrentSessionBindings, refreshMpvSessionBindings };
}
